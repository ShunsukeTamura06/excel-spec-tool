"""GET /spec/{job_id} — 設計書 Markdown 取得."""

from __future__ import annotations

from typing import Any

from fastapi import APIRouter, Depends, HTTPException

from backend.dependencies import get_storage
from backend.storage import JobNotFoundError, Storage

router = APIRouter()


@router.get("/spec/{job_id}")
async def get_spec(
    job_id: str,
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """生成済み設計書 (Markdown) と meta を返す."""
    try:
        meta = storage.get_meta(job_id)
        spec_md = storage.load_spec(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e
    except FileNotFoundError as e:
        raise HTTPException(
            status_code=409,
            detail="spec not generated yet; call /analyze first",
        ) from e

    return {"spec_md": spec_md, "meta": meta.model_dump()}
