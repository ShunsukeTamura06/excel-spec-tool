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
- list_workbook_objects: グラフ / ピボット / Power Query・外部接続の棚卸しを取得
- list_analysis_risks: 動的参照など、静的解析では断定できない未解析リスクを取得
- lookup_external_function: Bloomberg BDH/BDP/BDS 等の Add-In 関数定義を引く
- list_external_functions_used: 当該ブックで使われている外部関数とその箇所を列挙
"""

from __future__ import annotations

import json
import logging
import os
from collections import Counter
from typing import Any

from backend.storage import JobNotFoundError, Storage
from core.exceptions import NamedRangeFixError
from core.external_functions import get_function, list_functions
from core.named_range_fix import propose_named_range_fix
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
            "name": "list_workbook_objects",
            "description": (
                "グラフ / ピボットテーブル / Power Query・外部接続の棚卸しを返す。"
                "表や列の変更がチャート、ピボット、外部データ接続に波及するか確認するときに使う。"
                "依存関係は OOXML から明示的に取れた範囲だけで、"
                "Power Query の M コード本文は含まない。"
                "`sheet` を渡すとそのシート上のグラフ・ピボットに絞り込める。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet": {"type": "string", "description": "シート名で絞る場合に指定"},
                    "kind": {
                        "type": "string",
                        "enum": ["chart", "pivot", "power_query", "all"],
                        "description": "取得対象。未指定なら all。",
                    },
                },
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_analysis_risks",
            "description": (
                "静的解析では影響範囲を断定できない未解析リスクを返す。"
                "動的 VBA 参照、ActiveSheet/Selection、イベントマクロ、"
                "INDIRECT/OFFSET、外部接続、元データ不明のグラフ/ピボットなどが対象。"
                "改修手順や波及範囲を答える前に、手動確認が必要なリスクを把握するために使う。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "severity": {
                        "type": "string",
                        "enum": ["high", "medium", "low", "all"],
                        "description": "重大度で絞る。未指定なら all。",
                    },
                    "category": {
                        "type": "string",
                        "description": "カテゴリで絞る場合に指定。",
                    },
                    "limit": {
                        "type": "integer",
                        "description": "最大件数 (デフォルト 100, 上限 500)",
                    },
                },
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "lookup_external_function",
            "description": (
                "Bloomberg / Refinitiv 等の Excel Add-In 関数の定義を引く。"
                "BDH / BDP / BDS のような非標準関数の引数・返り値・使用例・"
                "落とし穴を、当ツールのレジストリから事実情報として取得する。"
                "推測やハルシネーションを避けるため、外部関数の挙動を答える前に必ず呼ぶこと。"
                '未登録ベンダーや未対応関数の場合は `{"error": "..."}` を返す。'
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {
                        "type": "string",
                        "description": "関数名 (例: 'BDH', 'BDP', 'BDS'). 大文字小文字無視.",
                    },
                },
                "required": ["name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "list_external_functions_used",
            "description": (
                "当該ワークブックで実際に使われている外部 Add-In 関数を一覧する。"
                "関数名 / 使用回数 / 主な使用箇所 (シート!セル, 最大 5 件) を返す。"
                "「このブックは Bloomberg 関数をどこで使っているか」を確認するときに使う。"
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
            "name": "lookup_references",
            "description": (
                "あるセルまたは範囲を参照している箇所 (数式 / VBA / グラフ / ピボット) を返す。"
                "改修の波及範囲を調べるときに使う。"
                "完全一致だけでなく範囲交差でヒットする (例: target='Input!A5' は "
                "'Input!A:A' を参照している数式も返す)。"
                "シート修飾を付けて呼ぶこと (例: 'Calc!H2')。"
                "グラフ系列参照とピボット元データは OOXML から明示的に取れた範囲だけが対象。"
                "VBA は静的に確定できる Range/Cells/短縮参照だけが対象で、"
                'Range("A" & row), Range(addr), ActiveSheet, Selection, Offset, Resize '
                "など実行時に決まる参照は検出対象外。"
                "0 件でも動的参照を含めて影響なしとは断定しないこと。"
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
    {
        "type": "function",
        "function": {
            "name": "propose_named_range_fix",
            "description": (
                "名前付き範囲の参照先を書き換えたら何が変わるかを試算する。"
                "この tool 自体はファイルを一切変更しない (読み取り専用の試算)。"
                "ユーザーから名前定義の修正依頼を受けたら、実際に適用する前に必ずこれを呼んで"
                "影響範囲 (波及範囲・既存リスク) を確認し、その内容をユーザーに提示すること。"
                "実際の適用はユーザーが画面上のボタンで明示的に実行するものであり、"
                "この tool を呼んだだけでは何も書き換わらない。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {
                        "type": "string",
                        "description": "書き換える名前付き範囲の名前 (例: 'TaxRate')",
                    },
                    "new_refers_to": {
                        "type": "string",
                        "description": (
                            "新しい参照先 (Excel 形式、シート修飾付き。"
                            '例: "Data!$B$1", "Data!$A$1:$A$200")'
                        ),
                    },
                },
                "required": ["name", "new_refers_to"],
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
        elif name == "list_workbook_objects":
            result = _exec_list_workbook_objects(storage, job_id, arguments)
        elif name == "list_analysis_risks":
            result = _exec_list_analysis_risks(storage, job_id, arguments)
        elif name == "lookup_external_function":
            result = _exec_lookup_external_function(arguments)
        elif name == "list_external_functions_used":
            result = _exec_list_external_functions_used(storage, job_id, arguments)
        elif name == "propose_named_range_fix":
            result = _exec_propose_named_range_fix(storage, job_id, arguments)
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


def _exec_list_workbook_objects(storage: Storage, job_id: str, args: dict[str, Any]) -> str:
    """グラフ / ピボット / Power Query・外部接続の棚卸しを返す."""
    sheet_filter_raw = args.get("sheet")
    sheet_filter = str(sheet_filter_raw) if sheet_filter_raw else None
    kind_raw = args.get("kind") or "all"
    kind = str(kind_raw)
    if kind not in {"chart", "pivot", "power_query", "all"}:
        raise ToolExecutionError("kind must be chart, pivot, power_query, or all")

    try:
        wb = storage.load_workbook(job_id)
    except JobNotFoundError as e:
        raise ToolExecutionError(f"job not found: {e}") from e
    except FileNotFoundError as e:
        raise ToolExecutionError(f"workbook not extracted: {e}") from e

    charts: list[dict[str, Any]] = []
    pivots: list[dict[str, Any]] = []
    for sheet in wb.sheets:
        if sheet_filter and sheet.name != sheet_filter:
            continue
        if kind in {"chart", "all"}:
            charts.extend(
                {
                    "sheet": sheet.name,
                    "name": chart.name,
                    "title": chart.title,
                    "chart_type": chart.chart_type,
                    "anchor": chart.anchor,
                    "series": [series.model_dump() for series in chart.series],
                }
                for chart in sheet.charts
            )
        if kind in {"pivot", "all"}:
            pivots.extend(
                {
                    "sheet": sheet.name,
                    **pivot.model_dump(),
                }
                for pivot in sheet.pivot_tables
            )

    power_queries: list[dict[str, Any]] = []
    if kind in {"power_query", "all"} and not sheet_filter:
        power_queries = [query.model_dump() for query in wb.power_queries]

    return json.dumps(
        {
            "charts": charts,
            "pivot_tables": pivots,
            "power_queries": power_queries,
            "counts": {
                "charts": len(charts),
                "pivot_tables": len(pivots),
                "power_queries": len(power_queries),
            },
            "analysis_scope": (
                "グラフ系列参照とピボット元データは OOXML から明示的に取れた範囲。"
                "Power Query は接続定義と queryTable 出力先の棚卸しで、M コード本文は未解析。"
            ),
        },
        ensure_ascii=False,
    )


def _exec_list_analysis_risks(storage: Storage, job_id: str, args: dict[str, Any]) -> str:
    """静的解析で断定できない未解析リスクを返す."""
    severity = str(args.get("severity") or "all")
    if severity not in {"high", "medium", "low", "all"}:
        raise ToolExecutionError("severity must be high, medium, low, or all")
    category_raw = args.get("category")
    category = str(category_raw) if category_raw else None
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

    risks = list(wb.analysis_risks)
    if severity != "all":
        risks = [risk for risk in risks if risk.severity == severity]
    if category:
        risks = [risk for risk in risks if risk.category == category]

    counts = {
        "high": sum(1 for risk in wb.analysis_risks if risk.severity == "high"),
        "medium": sum(1 for risk in wb.analysis_risks if risk.severity == "medium"),
        "low": sum(1 for risk in wb.analysis_risks if risk.severity == "low"),
    }
    sliced = risks[:limit]
    return json.dumps(
        {
            "risks": [risk.model_dump() for risk in sliced],
            "counts": counts,
            "total": len(risks),
            "returned": len(sliced),
            "truncated": len(risks) > limit,
            "analysis_scope": (
                "ここに出る項目は静的解析では影響範囲を断定できない箇所。"
                "回答時は確認済み事実と未解析リスクを分け、手動確認対象として扱う。"
            ),
        },
        ensure_ascii=False,
    )


def _exec_lookup_external_function(args: dict[str, Any]) -> str:
    """Bloomberg/Refinitiv 等の外部 Add-In 関数の定義をレジストリから引く."""
    name_raw = args.get("name")
    if not name_raw:
        raise ToolExecutionError("name is required")
    fn = get_function(str(name_raw))
    if fn is None:
        known = [f.name for f in list_functions()]
        return json.dumps(
            {
                "error": f"unknown external function: {name_raw}",
                "known": known,
            },
            ensure_ascii=False,
        )
    return json.dumps(fn.model_dump(), ensure_ascii=False)


def _exec_list_external_functions_used(storage: Storage, job_id: str, args: dict[str, Any]) -> str:
    """このジョブの Workbook で使われている外部関数を集計して返す."""
    try:
        wb = storage.load_workbook(job_id)
    except JobNotFoundError as e:
        raise ToolExecutionError(f"job not found: {e}") from e
    except FileNotFoundError as e:
        raise ToolExecutionError(f"workbook not extracted: {e}") from e

    counts: Counter[str] = Counter()
    locations: dict[str, list[str]] = {}
    for sheet in wb.sheets:
        for f in sheet.formulas:
            for fn_name in f.external_functions:
                counts[fn_name] += 1
                if len(locations.get(fn_name, [])) < 5:
                    coord_disp = f.coord if "!" in f.coord else f"{sheet.name}!{f.coord}"
                    locations.setdefault(fn_name, []).append(coord_disp)

    items = []
    for fn_name, cnt in counts.most_common():
        registered = get_function(fn_name)
        items.append(
            {
                "name": fn_name,
                "vendor": registered.vendor if registered else "?",
                "count": cnt,
                "top_locations": locations.get(fn_name, []),
                "short": registered.short if registered else "",
            }
        )
    return json.dumps(
        {"items": items, "total_kinds": len(items), "total_uses": sum(counts.values())},
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
        {
            "refs": [r.model_dump(by_alias=True) for r in refs],
            "count": len(refs),
            "analysis_scope": (
                "静的解析で検出できる数式参照、VBA の静的 Range/Cells/短縮参照、"
                "グラフ系列参照、ピボット元データが対象。"
                "動的に組み立てる VBA 参照や実行時状態依存の参照は含まれない。"
            ),
        },
        ensure_ascii=False,
    )


def _exec_propose_named_range_fix(storage: Storage, job_id: str, args: dict[str, Any]) -> str:
    name = args.get("name")
    new_refers_to = args.get("new_refers_to")
    if not name or not new_refers_to:
        raise ToolExecutionError("name and new_refers_to are required")
    try:
        wb = storage.load_workbook(job_id)
        idx = storage.load_references(job_id)
    except JobNotFoundError as e:
        raise ToolExecutionError(f"job not found: {e}") from e
    except FileNotFoundError as e:
        raise ToolExecutionError(f"workbook or references not built: {e}") from e
    try:
        diff = propose_named_range_fix(wb, idx, str(name), str(new_refers_to))
    except NamedRangeFixError as e:
        raise ToolExecutionError(str(e)) from e
    return json.dumps(
        {
            "proposal": diff.model_dump(by_alias=True),
            "note": (
                "これは試算結果であり、ファイルはまだ変更されていない。"
                "ユーザーが画面のボタンで適用を確定するまで実ファイルは変わらない。"
            ),
        },
        ensure_ascii=False,
    )
