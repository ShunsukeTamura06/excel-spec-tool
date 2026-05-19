"""POST /feedback — ユーザーフィードバックを受け取って永続化する.

設計方針:
- ユーザーの心理的負担を最小化するため、ほとんどのフィールドが任意.
- 失敗してもユーザーの作業を妨げないよう、エラーは控えめに返す.
- 認証なし (管理画面実装時に必要に応じて追加).
"""

from __future__ import annotations

import logging
import uuid
from datetime import datetime, timezone
from typing import Any

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, Field

from backend.dependencies import get_storage
from backend.storage import Storage
from core.models import Feedback, FeedbackKind

logger = logging.getLogger(__name__)
router = APIRouter()


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


class FeedbackInput(BaseModel):
    """フロントから受け取るフィードバック本体. id / timestamp は backend で付与."""

    kind: FeedbackKind
    comment: str = ""
    page: str = ""
    job_id: str | None = None
    target_id: str | None = None
    target_excerpt: str = Field(default="", max_length=500)
    user_label: str = Field(default="", max_length=80)


_MAX_COMMENT_CHARS = 4000


@router.post("/feedback")
async def submit_feedback(
    body: FeedbackInput,
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """フィードバックを受け付けて永続化する.

    バリデーション:
      - comment は 4000 字までに切る (長すぎは握りつぶさず詰める).
      - kind は Pydantic 側で Literal 制約.
    成功時: {"ok": True, "id": "..."} を返す.
    """
    comment = (body.comment or "").strip()
    if len(comment) > _MAX_COMMENT_CHARS:
        comment = comment[:_MAX_COMMENT_CHARS]

    item = Feedback(
        id=str(uuid.uuid4()),
        timestamp=_utc_now_iso(),
        kind=body.kind,
        comment=comment,
        page=body.page or "",
        job_id=body.job_id,
        target_id=body.target_id,
        target_excerpt=(body.target_excerpt or "")[:500],
        user_label=(body.user_label or "").strip()[:80],
    )

    try:
        storage.append_feedback(item)
    except Exception as e:
        logger.exception("failed to persist feedback")
        # 永続化に失敗してもユーザー体験を壊さないよう 500 は出すが、最小限の情報のみ
        raise HTTPException(status_code=500, detail="failed to save feedback") from e

    logger.info(
        "feedback received: kind=%s job_id=%s comment_chars=%d page=%s",
        item.kind,
        item.job_id or "-",
        len(item.comment),
        item.page,
    )
    return {"ok": True, "id": item.id}
