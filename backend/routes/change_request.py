"""POST /change-request/{job_id} — 業務要望を改修依頼書へ整理する."""

from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, Field

from backend.dependencies import get_storage
from backend.storage import JobNotFoundError, Storage
from core.change_request import ChangeBrief, build_change_brief

router = APIRouter()


class ChangeRequestInput(BaseModel):
    """改修依頼書の作成入力."""

    requested_outcome: str = Field(min_length=1, max_length=4000)
    feature_id: str | None = Field(default=None, max_length=40)


@router.post("/change-request/{job_id}")
async def create_change_request(
    job_id: str,
    request: ChangeRequestInput,
    storage: Storage = Depends(get_storage),
) -> ChangeBrief:
    """保存済み診断を根拠に一般ユーザー向け改修依頼書を返す."""
    try:
        diagnosis = storage.load_diagnosis(job_id)
        return build_change_brief(
            diagnosis,
            requested_outcome=request.requested_outcome,
            feature_id=request.feature_id,
        )
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error
    except JobNotFoundError as error:
        raise HTTPException(status_code=404, detail="job not found") from error
    except FileNotFoundError as error:
        raise HTTPException(
            status_code=409,
            detail="diagnosis not generated yet; call /analyze first",
        ) from error
