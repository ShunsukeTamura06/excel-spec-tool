"""GET /jobs (一覧) と DELETE /jobs/{job_id} (削除).

`GET /jobs` は docs/SPEC.ja.md §5.1 表には無いが、§6.2 で Frontend が
「過去のジョブ一覧」から再選択できる仕様のため、自然な派生として追加する
(B 案で承認済み).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import FileResponse

from backend.dependencies import get_storage
from backend.storage import JobNotFoundError, Storage

router = APIRouter()


@router.get("/jobs")
async def list_jobs(
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """登録済みジョブのメタを created_at 降順で返す."""
    metas = storage.list_jobs()
    return {"jobs": [m.model_dump() for m in metas]}


@router.get("/jobs/{job_id}/download", response_class=FileResponse)
async def download_job(
    job_id: str,
    storage: Storage = Depends(get_storage),
) -> FileResponse:
    """ジョブのExcelファイルを、識別しやすい名前でダウンロードする."""

    try:
        meta = storage.get_meta(job_id)
        path = storage.get_original_path(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {exc}") from exc
    except JobNotFoundError as exc:
        raise HTTPException(status_code=404, detail=f"job not found: {exc}") from exc

    if meta.status == "failed":
        raise HTTPException(
            status_code=409,
            detail=(
                "this job failed processing or verification and cannot be downloaded; "
                "check /jobs/{job_id}/verification for details"
            ),
        )

    original_name = Path(meta.filename)
    download_name = f"{original_name.stem}_xlblueprint{original_name.suffix}"
    media_type = (
        "application/vnd.ms-excel.sheet.macroEnabled.12"
        if original_name.suffix.lower() == ".xlsm"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return FileResponse(
        path,
        filename=download_name,
        media_type=media_type,
    )


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
