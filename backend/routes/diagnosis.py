"""GET /diagnosis/{job_id} — 根拠付き Excel 診断の取得."""

from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException

from backend.dependencies import get_storage
from backend.storage import JobNotFoundError, Storage
from core.diagnosis import WorkbookDiagnosis

router = APIRouter()


@router.get("/diagnosis/{job_id}")
async def get_diagnosis(
    job_id: str,
    storage: Storage = Depends(get_storage),
) -> WorkbookDiagnosis:
    """分析済みジョブの一般ユーザー向け Excel 診断を返す."""
    try:
        return storage.load_diagnosis(job_id)
    except ValueError as error:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {error}") from error
    except JobNotFoundError as error:
        raise HTTPException(status_code=404, detail="job not found") from error
    except FileNotFoundError as error:
        raise HTTPException(
            status_code=409,
            detail="diagnosis not generated yet; call /analyze first",
        ) from error
