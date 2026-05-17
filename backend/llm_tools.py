"""LLM が呼べる tool の定義と実行ロジック.

OpenAI function calling のフォーマットで定義し、`build_tool_definitions()` で
LLM に渡す `tools` 配列を取得、`execute_tool_call()` で実行する。

提供するツール:
- get_cells_range: 指定範囲のセル値を 2D 配列で取得
- find_cells: 値の部分一致でセルを検索
- lookup_references: 既存の参照逆引きインデックスを引く (範囲交差)
- list_vba_modules: VBA モジュールとプロシージャのメタ情報一覧
- get_vba_procedure: 指定プロシージャのソースコード本体を取得
- list_sheet_formulas: シート内の数式一覧 (オプションでテキストフィルタ)
"""

from __future__ import annotations

import json
import logging
import os
from typing import Any

from backend.storage import JobNotFoundError, Storage
from core.reference_index import find_overlapping

logger = logging.getLogger(__name__)


# tool 1 件あたり LLM に返す結果の最大文字数. 巨大な find_cells / lookup_references
# 結果でコンテキストを食い潰さないためのセーフティネット.
# 環境変数 TOOL_RESULT_MAX_CHARS で上書き可能 (最低 1000 にクランプ).
_DEFAULT_TOOL_RESULT_MAX_CHARS = 20_000


def _tool_result_max_chars() -> int:
    raw = os.environ.get("TOOL_RESULT_MAX_CHARS")
    if not raw:
        return _DEFAULT_TOOL_RESULT_MAX_CHARS
    try:
        return max(1000, int(raw))
    except ValueError:
        logger.warning(
            "invalid TOOL_RESULT_MAX_CHARS=%r; using default %d",
            raw,
            _DEFAULT_TOOL_RESULT_MAX_CHARS,
        )
        return _DEFAULT_TOOL_RESULT_MAX_CHARS


def _truncate_tool_result(result: str, limit: int) -> str:
    """tool 結果文字列を limit までに切り詰める.

    切り詰め時はマーカーを末尾に付与するが、limit がマーカーより小さい場合は
    マーカーを省略し、必ず limit を守る (本番下限 1000 ではほぼ発生しない).
    """
    if len(result) <= limit:
        return result
    marker = (
        f"\n... [TRUNCATED: tool result was {len(result)} chars, "
        f"capped at {limit}. Refine query or use a smaller range.]"
    )
    if limit <= len(marker):
        return result[:limit]
    return result[: limit - len(marker)] + marker


# OpenAI function calling 仕様の tool 定義 (chat completions の `tools` パラメータに渡す形)
TOOL_DEFINITIONS: list[dict[str, Any]] = [
    {
        "type": "function",
        "function": {
            "name": "get_cells_range",
            "description": (
                "指定したシートの指定範囲のセル値を取得する。"
                "改修したい箇所周辺のセル配置や値を確認するときに使う。"
                "例: シート 'Portfolio' の 6 行目を見たいなら range='A6:Z6'。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet": {"type": "string", "description": "シート名 (例: 'Portfolio')"},
                    "range": {
                        "type": "string",
                        "description": "Excel 範囲表記 (例: 'A6:F6', 'A1:Z20', 'B3')",
                    },
                },
                "required": ["sheet", "range"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "find_cells",
            "description": (
                "セルの value 文字列を部分一致検索する。"
                "ユーザーが述べた項目名や値を Excel 内で探したいときに使う。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "検索文字列"},
                    "sheet": {
                        "type": "string",
                        "description": "シートを絞る場合のシート名 (任意)",
                    },
                    "limit": {
                        "type": "integer",
                        "description": "最大件数 (デフォルト 20, 上限 200)",
                    },
                },
                "required": ["query"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_vba_modules",
            "description": (
                "VBA モジュールとそれに含まれるプロシージャの一覧 (名前・種別・行範囲) を返す。"
                "コード本体は含まず軽量。"
                "ユーザーの要望に関係しそうなプロシージャを特定する起点として使う。"
                "本体を読みたい場合はその後 `get_vba_procedure` を呼ぶ。"
            ),
            "parameters": {
                "type": "object",
                "properties": {},
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_vba_procedure",
            "description": (
                "指定した VBA モジュール内の指定プロシージャ (Sub/Function/Property) の "
                "ソースコード本体を返す。"
                "プロシージャ名は `list_vba_modules` で確認した正確な名前を渡すこと。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "module": {
                        "type": "string",
                        "description": "モジュール名 (例: 'Module1')",
                    },
                    "name": {
                        "type": "string",
                        "description": "プロシージャ名 (例: 'UpdateDaily')",
                    },
                },
                "required": ["module", "name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_sheet_formulas",
            "description": (
                "指定シートの数式一覧を返す。"
                "設計書には参照数が多い上位 10 件しか載らないので、"
                "それ以外の数式や特定関数を使っている数式を探したい時に使う。"
                "`pattern` を渡すと数式テキストの部分一致 (大文字小文字無視) で絞り込める。"
                "例: pattern='SUMIF' で SUMIF を使っている数式だけ返す。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet": {"type": "string", "description": "シート名"},
                    "pattern": {
                        "type": "string",
                        "description": "数式テキストの部分一致フィルタ (任意, 大文字小文字無視)",
                    },
                    "limit": {
                        "type": "integer",
                        "description": "最大件数 (デフォルト 100, 上限 500)",
                    },
                },
                "required": ["sheet"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "lookup_references",
            "description": (
                "あるセルまたは範囲を参照している箇所 (数式 / VBA) を返す。"
                "改修の波及範囲を調べるときに使う。"
                "完全一致だけでなく範囲交差でヒットする (例: target='Input!A5' は "
                "'Input!A:A' を参照している数式も返す)。"
                "シート修飾を付けて呼ぶこと (例: 'Calc!H2')。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "target": {
                        "type": "string",
                        "description": "参照先キー (例: 'Calc!H2', 'Input!A:A')",
                    },
                },
                "required": ["target"],
            },
        },
    },
]


def build_tool_definitions() -> list[dict[str, Any]]:
    """LLM に渡す tools 配列を返す."""
    return TOOL_DEFINITIONS


class ToolExecutionError(Exception):
    """tool 実行中のエラー (引数不正 / 該当データなし等)."""


def execute_tool_call(
    storage: Storage,
    job_id: str,
    name: str,
    arguments: dict[str, Any],
) -> str:
    """tool 1 件を実行し、結果を JSON 文字列で返す.

    Returns:
        LLM に返す tool 結果. JSON 文字列にして返す (OpenAI 仕様)。
        エラー時もメッセージを JSON にして返し、例外は呼び出し側に伝播しない。
    """
    try:
        if name == "get_cells_range":
            result = _exec_get_cells_range(storage, job_id, arguments)
        elif name == "find_cells":
            result = _exec_find_cells(storage, job_id, arguments)
        elif name == "lookup_references":
            result = _exec_lookup_references(storage, job_id, arguments)
        elif name == "list_vba_modules":
            result = _exec_list_vba_modules(storage, job_id, arguments)
        elif name == "get_vba_procedure":
            result = _exec_get_vba_procedure(storage, job_id, arguments)
        elif name == "list_sheet_formulas":
            result = _exec_list_sheet_formulas(storage, job_id, arguments)
        else:
            return json.dumps({"error": f"unknown tool: {name}"}, ensure_ascii=False)
    except Exception as e:  # noqa: BLE001 - tool ループは壊さない
        logger.warning("Tool %s failed: %s", name, e)
        return json.dumps({"error": str(e), "tool": name}, ensure_ascii=False)
    return _truncate_tool_result(result, _tool_result_max_chars())


def _exec_get_cells_range(storage: Storage, job_id: str, args: dict[str, Any]) -> str:
    sheet = args.get("sheet")
    range_str = args.get("range")
    if not sheet or not range_str:
        raise ToolExecutionError("sheet and range are required")
    try:
        result = storage.get_cells_range(job_id, sheet=str(sheet), range_str=str(range_str))
    except JobNotFoundError as e:
        raise ToolExecutionError(f"job not found: {e}") from e
    except FileNotFoundError as e:
        raise ToolExecutionError(f"cells.db not built: {e}") from e
    return json.dumps(result, ensure_ascii=False)


def _exec_find_cells(storage: Storage, job_id: str, args: dict[str, Any]) -> str:
    query = args.get("query")
    if not query:
        raise ToolExecutionError("query is required")
    sheet_arg = args.get("sheet")
    sheet = str(sheet_arg) if sheet_arg else None
    limit_raw = args.get("limit", 20)
    try:
        limit = int(limit_raw)
    except (TypeError, ValueError):
        limit = 20
    try:
        matches = storage.find_cells(job_id, query=str(query), sheet=sheet, limit=limit)
    except JobNotFoundError as e:
        raise ToolExecutionError(f"job not found: {e}") from e
    except FileNotFoundError as e:
        raise ToolExecutionError(f"cells.db not built: {e}") from e
    return json.dumps({"matches": matches, "count": len(matches)}, ensure_ascii=False)


def _exec_list_vba_modules(storage: Storage, job_id: str, args: dict[str, Any]) -> str:
    """VBA モジュールとプロシージャのメタ情報のみ返す (コード本体は含まない)."""
    try:
        wb = storage.load_workbook(job_id)
    except JobNotFoundError as e:
        raise ToolExecutionError(f"job not found: {e}") from e
    except FileNotFoundError as e:
        raise ToolExecutionError(f"workbook not extracted: {e}") from e
    modules = [
        {
            "name": m.name,
            "type": m.type,
            "procedures": [
                {
                    "name": p.name,
                    "kind": p.kind,
                    "start_line": p.start_line,
                    "end_line": p.end_line,
                }
                for p in m.procedures
            ],
        }
        for m in wb.vba_modules
    ]
    return json.dumps({"modules": modules, "count": len(modules)}, ensure_ascii=False)


def _exec_get_vba_procedure(storage: Storage, job_id: str, args: dict[str, Any]) -> str:
    """指定モジュール内の指定プロシージャのソースを返す."""
    module_name = args.get("module")
    proc_name = args.get("name")
    if not module_name or not proc_name:
        raise ToolExecutionError("module and name are required")
    try:
        wb = storage.load_workbook(job_id)
    except JobNotFoundError as e:
        raise ToolExecutionError(f"job not found: {e}") from e
    except FileNotFoundError as e:
        raise ToolExecutionError(f"workbook not extracted: {e}") from e

    module = next((m for m in wb.vba_modules if m.name == str(module_name)), None)
    if module is None:
        raise ToolExecutionError(f"module not found: {module_name}")
    proc = next((p for p in module.procedures if p.name == str(proc_name)), None)
    if proc is None:
        raise ToolExecutionError(f"procedure not found in {module_name}: {proc_name}")

    # 抽出時にプロシージャ単位の code が空でも、モジュール全体コードから行範囲で切り出す
    code = proc.code
    if not code and module.code:
        lines = module.code.splitlines()
        # start_line / end_line は 1-origin
        start = max(1, proc.start_line) - 1
        end = max(proc.end_line, proc.start_line)
        code = "\n".join(lines[start:end])

    return json.dumps(
        {
            "module": module.name,
            "name": proc.name,
            "kind": proc.kind,
            "start_line": proc.start_line,
            "end_line": proc.end_line,
            "code": code,
            "annotation": proc.annotation,
        },
        ensure_ascii=False,
    )


def _exec_list_sheet_formulas(storage: Storage, job_id: str, args: dict[str, Any]) -> str:
    """シートの数式一覧を返す (オプションでテキストフィルタ)."""
    sheet_name = args.get("sheet")
    if not sheet_name:
        raise ToolExecutionError("sheet is required")
    pattern_raw = args.get("pattern")
    pattern = str(pattern_raw).lower() if pattern_raw else None
    limit_raw = args.get("limit", 100)
    try:
        limit = int(limit_raw)
    except (TypeError, ValueError):
        limit = 100
    limit = max(1, min(500, limit))

    try:
        wb = storage.load_workbook(job_id)
    except JobNotFoundError as e:
        raise ToolExecutionError(f"job not found: {e}") from e
    except FileNotFoundError as e:
        raise ToolExecutionError(f"workbook not extracted: {e}") from e

    sheet = next((s for s in wb.sheets if s.name == str(sheet_name)), None)
    if sheet is None:
        raise ToolExecutionError(f"sheet not found: {sheet_name}")

    if pattern:
        matched = [f for f in sheet.formulas if pattern in f.formula.lower()]
    else:
        matched = list(sheet.formulas)

    total = len(matched)
    sliced = matched[:limit]
    return json.dumps(
        {
            "sheet": sheet.name,
            "total": total,
            "returned": len(sliced),
            "truncated": total > limit,
            "formulas": [{"coord": f.coord, "formula": f.formula, "refs": f.refs} for f in sliced],
        },
        ensure_ascii=False,
    )


def _exec_lookup_references(storage: Storage, job_id: str, args: dict[str, Any]) -> str:
    target = args.get("target")
    if not target:
        raise ToolExecutionError("target is required")
    try:
        idx = storage.load_references(job_id)
    except JobNotFoundError as e:
        raise ToolExecutionError(f"job not found: {e}") from e
    except FileNotFoundError as e:
        raise ToolExecutionError(f"references not built: {e}") from e
    # 範囲交差で検索. `Calc!H2` で `Calc!A1:J100` のような上位範囲もヒットさせる.
    refs = find_overlapping(idx, str(target))
    return json.dumps(
        {"refs": [r.model_dump(by_alias=True) for r in refs], "count": len(refs)},
        ensure_ascii=False,
    )
