"""GET /jobs (一覧) と DELETE /jobs/{job_id} (削除).

`GET /jobs` は SPEC.md §5.1 表には無いが、§6.2 で Frontend が
「過去のジョブ一覧」から再選択できる仕様のため、自然な派生として追加する
(B 案で承認済み).
"""

from __future__ import annotations

from typing import Any

from fastapi import APIRouter, Depends, HTTPException

from backend.dependencies import get_storage
from backend.storage import Storage

router = APIRouter()


@router.get("/jobs")
async def list_jobs(
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """登録済みジョブのメタを created_at 降順で返す."""
    metas = storage.list_jobs()
    return {"jobs": [m.model_dump() for m in metas]}


@router.delete("/jobs/{job_id}")
async def delete_job(
    job_id: str,
    storage: Storage = Depends(get_storage),
) -> dict[str, bool]:
    """ジョブを削除. 存在しない場合は 404."""
    try:
        deleted = storage.delete_job(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    if not deleted:
        raise HTTPException(status_code=404, detail="job not found")
    return {"deleted": True}
