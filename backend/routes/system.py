"""GET /system/llm-status — LLM 接続設定の有無を返す.

LLM が未設定 (= MockLLMClient フォールバック) のとき、フロントエンドは
チャット空状態などに onboarding カードを出す. その判定に使う.

セキュリティ: 設定値そのもの (API key / base URL) は返さない. 設定済か否か
と pro/fast モデルラベルのみを返す.
"""

from __future__ import annotations

from typing import Any

from fastapi import APIRouter, Depends

from backend.dependencies import get_llm_client
from backend.llm_client import LLMClient, MockLLMClient, OpenAICompatibleLLMClient

router = APIRouter()


@router.get("/system/llm-status")
async def llm_status(llm: LLMClient = Depends(get_llm_client)) -> dict[str, Any]:
    """LLM クライアントの設定状況を返す.

    Returns:
        configured: True なら実 LLM 接続. False ならモック.
        mode:       "openai_compatible" / "mock" / "unknown".
        pro_model:  pro tier のモデル名 (取れる場合).
        fast_model: fast tier のモデル名 (取れる場合).
    """
    if isinstance(llm, OpenAICompatibleLLMClient):
        return {
            "configured": True,
            "mode": "openai_compatible",
            "pro_model": llm.default_model,
            "fast_model": llm.fast_model,
        }
    if isinstance(llm, MockLLMClient):
        return {
            "configured": False,
            "mode": "mock",
            "pro_model": llm.pro_model,
            "fast_model": llm.fast_model,
        }
    return {
        "configured": False,
        "mode": "unknown",
        "pro_model": "",
        "fast_model": "",
    }
