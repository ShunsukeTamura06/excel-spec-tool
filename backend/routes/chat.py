"""POST /chat/{job_id} と GET /chat/{job_id}/history — 改修対話."""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel

from backend.dependencies import get_llm_client, get_storage
from backend.llm_client import LLMClient
from backend.storage import JobNotFoundError, Storage
from core.models import ChatMessage

router = APIRouter()


class ChatRequest(BaseModel):
    """POST /chat/{job_id} のリクエストボディ."""

    message: str


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _build_system_prompt(spec_md: str) -> str:
    """設計書を system prompt に固定する.

    SPEC §5.4 の「応答には改修手順 + 波及範囲を含める」指示も付与。
    """
    instructions = (
        "あなたは Excel 改修支援アシスタントです。次の設計書を参照しつつ、"
        "ユーザーの改修要望に答えてください。応答には必ず以下のセクションを含めてください:\n"
        "1. 改修手順 (ユーザーの操作レベル)\n"
        "2. 波及範囲 (影響を受けるセルや VBA)\n"
    )
    if spec_md:
        return f"{instructions}\n---\n# 設計書\n{spec_md}"
    return instructions


@router.post("/chat/{job_id}")
async def chat(
    job_id: str,
    body: ChatRequest,
    storage: Storage = Depends(get_storage),
    llm: LLMClient = Depends(get_llm_client),
) -> dict[str, Any]:
    """ユーザー発話を受け付け、LLM 応答を返す. 履歴も jsonl に追記."""
    if not body.message.strip():
        raise HTTPException(status_code=400, detail="message is required")

    try:
        # 設計書はあれば system prompt に固定、なければ空文字
        try:
            spec_md = storage.load_spec(job_id)
        except FileNotFoundError:
            spec_md = ""

        history = storage.load_chat_history(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e

    # LLM へのメッセージを組み立て
    messages: list[dict[str, str]] = [
        {"role": "system", "content": _build_system_prompt(spec_md)},
    ]
    messages.extend({"role": m.role, "content": m.content} for m in history)
    messages.append({"role": "user", "content": body.message})

    reply = llm.chat_completion(messages)

    # 履歴に追記 (user → assistant の順)
    now = _utc_now_iso()
    storage.append_chat_message(
        job_id, ChatMessage(role="user", content=body.message, timestamp=now)
    )
    storage.append_chat_message(job_id, ChatMessage(role="assistant", content=reply, timestamp=now))

    new_history = storage.load_chat_history(job_id)
    return {
        "reply": reply,
        "history": [m.model_dump() for m in new_history],
    }


@router.get("/chat/{job_id}/history")
async def chat_history(
    job_id: str,
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """ジョブのチャット履歴を返す."""
    try:
        history = storage.load_chat_history(job_id)
        # ジョブ存在確認 (履歴が空でも meta は無いとおかしい)
        storage.get_meta(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e

    return {"history": [m.model_dump() for m in history]}
