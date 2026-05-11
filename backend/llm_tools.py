"""LLM が呼べる tool の定義と実行ロジック.

OpenAI function calling のフォーマットで定義し、`build_tool_definitions()` で
LLM に渡す `tools` 配列を取得、`execute_tool_call()` で実行する。

提供するツール:
- get_cells_range: 指定範囲のセル値を 2D 配列で取得
- find_cells: 値の部分一致でセルを検索
- lookup_references: 既存の参照逆引きインデックスを引く
"""

from __future__ import annotations

import json
import logging
from typing import Any

from backend.storage import JobNotFoundError, Storage

logger = logging.getLogger(__name__)


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
            "name": "lookup_references",
            "description": (
                "あるセルまたは範囲を参照している箇所 (数式 / VBA) を返す。"
                "改修の波及範囲を調べるときに使う。"
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
            return _exec_get_cells_range(storage, job_id, arguments)
        if name == "find_cells":
            return _exec_find_cells(storage, job_id, arguments)
        if name == "lookup_references":
            return _exec_lookup_references(storage, job_id, arguments)
        return json.dumps({"error": f"unknown tool: {name}"}, ensure_ascii=False)
    except Exception as e:  # noqa: BLE001 - tool ループは壊さない
        logger.warning("Tool %s failed: %s", name, e)
        return json.dumps({"error": str(e), "tool": name}, ensure_ascii=False)


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
    refs = idx.refs.get(str(target), [])
    return json.dumps(
        {"refs": [r.model_dump(by_alias=True) for r in refs], "count": len(refs)},
        ensure_ascii=False,
    )
