"""GET /references/{job_id}?target=... — 参照逆引き検索."""

from __future__ import annotations

from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Query

from backend.dependencies import get_storage
from backend.storage import JobNotFoundError, Storage

router = APIRouter()


@router.get("/references/{job_id}")
async def get_references(
    job_id: str,
    target: str = Query("", description="参照先キー (例: 'Calc!H2'). 空なら空配列"),
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """指定された参照先を持つ Reference 一覧を返す."""
    try:
        idx = storage.load_references(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e
    except FileNotFoundError as e:
        raise HTTPException(
            status_code=409,
            detail="references not built yet; call /extract first",
        ) from e

    refs = idx.refs.get(target, []) if target else []
    return {"refs": [r.model_dump(by_alias=True) for r in refs]}
