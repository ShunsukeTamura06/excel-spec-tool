"""POST /chat/{job_id} と GET /chat/{job_id}/history — 改修対話.

LLM の function calling ループを実装している:
1. system prompt + 履歴 + ユーザー発話 + tools 定義を LLM に渡す
2. 応答に tool_calls があれば、各 tool を実行して結果を tool role の
   メッセージとして追加し、再度 LLM を呼ぶ
3. tool_calls が無くなったら最終応答テキストを返す
4. 暴走防止のため最大反復回数を制限する
"""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from typing import Any

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel

from backend.dependencies import get_llm_client, get_storage
from backend.llm_client import LLMClient, LLMResponse
from backend.llm_tools import build_tool_definitions, execute_tool_call
from backend.storage import JobNotFoundError, Storage
from core.models import ChatMessage

router = APIRouter()
logger = logging.getLogger(__name__)


# tool ループの最大反復回数. これを超えたら強制的に終了する。
MAX_TOOL_ITERATIONS = 8


class ChatRequest(BaseModel):
    """POST /chat/{job_id} のリクエストボディ."""

    message: str


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _build_system_prompt(spec_md: str) -> str:
    """設計書を system prompt に固定する.

    SPEC §5.4 の「応答には改修手順 + 波及範囲を含める」指示に加え、
    cells API / 参照逆引きを tool 経由で参照できることを明示する。
    """
    instructions = (
        "あなたは Excel 改修支援アシスタントです。"
        "ユーザーの改修要望に対し、設計書とツールを参照して具体的な手順と影響範囲を回答してください。\n\n"
        "応答には必ず以下のセクションを含めてください:\n"
        "1. 改修手順 (ユーザーの操作レベル)\n"
        "2. 波及範囲 (影響を受けるセルや VBA)\n\n"
        "設計書には Excel の概観のみ載っています。具体的なセル値や行内容を確認したい時は\n"
        "  - get_cells_range(sheet, range): 指定範囲を読む\n"
        "  - find_cells(query, sheet=?, limit=?): 値を検索する\n"
        "  - lookup_references(target): あるセル/範囲を参照している箇所を引く\n"
        "を呼んでください。情報が足りなければ追加質問してから回答してください。"
    )
    if spec_md:
        return f"{instructions}\n\n---\n# 設計書\n{spec_md}"
    return instructions


def _tool_call_to_assistant_message(resp: LLMResponse) -> dict[str, Any]:
    """LLMResponse の tool_calls を OpenAI 仕様の assistant メッセージに変換."""
    return {
        "role": "assistant",
        "content": resp.content or "",
        "tool_calls": [
            {
                "id": tc.id,
                "type": "function",
                "function": {
                    "name": tc.name,
                    "arguments": json.dumps(tc.arguments, ensure_ascii=False),
                },
            }
            for tc in resp.tool_calls
        ],
    }


def _run_tool_loop(
    llm: LLMClient,
    storage: Storage,
    job_id: str,
    base_messages: list[dict[str, Any]],
) -> tuple[str, list[dict[str, Any]]]:
    """tool 呼び出しを伴うチャット応答ループ.

    Returns:
        (最終アシスタント応答テキスト, ループ中の中間 tool 呼び出し記録).
        中間記録は履歴 jsonl には保存しないがデバッグ・観測用に返す。
    """
    tools = build_tool_definitions()
    messages = list(base_messages)
    tool_trace: list[dict[str, Any]] = []

    for _iteration in range(MAX_TOOL_ITERATIONS):
        resp = llm.chat_completion_with_tools(messages, tools=tools)

        if not resp.tool_calls:
            # 最終応答
            return (resp.content or "", tool_trace)

        # tool 呼び出しを実行し、結果を tool role メッセージで追加
        messages.append(_tool_call_to_assistant_message(resp))
        for tc in resp.tool_calls:
            result_str = execute_tool_call(storage, job_id, tc.name, tc.arguments)
            tool_trace.append(
                {"name": tc.name, "arguments": tc.arguments, "result_preview": result_str[:200]}
            )
            messages.append(
                {
                    "role": "tool",
                    "tool_call_id": tc.id,
                    "name": tc.name,
                    "content": result_str,
                }
            )

    # 上限到達: 最後の応答を返す。content が空の場合は注意書きを付与。
    logger.warning("Tool loop hit MAX_TOOL_ITERATIONS=%d for job %s", MAX_TOOL_ITERATIONS, job_id)
    return (
        (resp.content or "[tool loop max iterations reached; partial result]"),
        tool_trace,
    )


@router.post("/chat/{job_id}")
async def chat(
    job_id: str,
    body: ChatRequest,
    storage: Storage = Depends(get_storage),
    llm: LLMClient = Depends(get_llm_client),
) -> dict[str, Any]:
    """ユーザー発話を受け付け、LLM 応答を返す. 履歴は jsonl に追記する."""
    if not body.message.strip():
        raise HTTPException(status_code=400, detail="message is required")

    try:
        try:
            spec_md = storage.load_spec(job_id)
        except FileNotFoundError:
            spec_md = ""

        history = storage.load_chat_history(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e

    messages: list[dict[str, Any]] = [
        {"role": "system", "content": _build_system_prompt(spec_md)},
    ]
    messages.extend({"role": m.role, "content": m.content} for m in history)
    messages.append({"role": "user", "content": body.message})

    reply, tool_trace = _run_tool_loop(llm, storage, job_id, messages)

    # 履歴に追記 (user → assistant). tool 呼び出しの中間結果は履歴には残さない
    now = _utc_now_iso()
    storage.append_chat_message(
        job_id, ChatMessage(role="user", content=body.message, timestamp=now)
    )
    storage.append_chat_message(job_id, ChatMessage(role="assistant", content=reply, timestamp=now))

    new_history = storage.load_chat_history(job_id)
    return {
        "reply": reply,
        "history": [m.model_dump() for m in new_history],
        "tool_trace": tool_trace,
    }


@router.get("/chat/{job_id}/history")
async def chat_history(
    job_id: str,
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """ジョブのチャット履歴を返す."""
    try:
        history = storage.load_chat_history(job_id)
        # ジョブ存在確認
        storage.get_meta(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e

    return {"history": [m.model_dump() for m in history]}
