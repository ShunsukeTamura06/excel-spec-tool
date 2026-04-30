"""POST /analyze/{job_id} — LLM 注釈 + 設計書生成."""

from __future__ import annotations

import logging

from fastapi import APIRouter, Depends, HTTPException

from backend.dependencies import get_llm_client, get_storage
from backend.llm_client import LLMClient
from backend.storage import JobNotFoundError, Storage
from core.spec_generator import generate_spec

logger = logging.getLogger(__name__)
router = APIRouter()


@router.post("/analyze/{job_id}")
async def analyze(
    job_id: str,
    storage: Storage = Depends(get_storage),
    llm: LLMClient = Depends(get_llm_client),  # noqa: ARG001 - 将来使用
) -> dict[str, str]:
    """抽出済みワークブックに LLM 注釈を付け、設計書を生成・保存する.

    Note:
        現状の LLM 注釈ステップは未実装 (TODO)。`generate_spec` のみ実行する。
        実装が入った時点で、SheetInfo.purpose / VbaProcedure.annotation /
        CellFormula.annotation を埋めてから保存する。
    """
    try:
        wb = storage.load_workbook(job_id)
        idx = storage.load_references(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e
    except FileNotFoundError as e:
        raise HTTPException(
            status_code=409,
            detail="workbook not extracted yet; call /extract first",
        ) from e

    # TODO: LLM 注釈をここで付与する (Sheet purpose, VBA procedure annotation, ...)
    spec_md = generate_spec(wb, idx)

    try:
        storage.save_spec(job_id, spec_md)
        storage.update_status(job_id, "analyzed")
    except Exception as e:
        logger.exception("Failed to persist spec for %s", job_id)
        storage.update_status(job_id, "failed")
        raise HTTPException(status_code=500, detail=str(e)) from e

    return {"status": "ok"}
