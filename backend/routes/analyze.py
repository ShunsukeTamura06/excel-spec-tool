"""POST /analyze/{job_id} — LLM 注釈 + 設計書生成."""

from __future__ import annotations

import logging

from fastapi import APIRouter, Depends, HTTPException

from backend.annotators import annotate_workbook
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
    llm: LLMClient = Depends(get_llm_client),
) -> dict[str, str]:
    """抽出済みワークブックに LLM 注釈を付け、設計書を生成・保存する.

    1. extracted.json から Workbook をロード
    2. `annotate_workbook` で SheetInfo.purpose / VbaProcedure.annotation を
       LLM (fast tier) で埋める
    3. 注釈済み Workbook を保存し直す
    4. `generate_spec` で Markdown 設計書を生成・保存
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

    logger.info("analyze started: sheets=%d vba_modules=%d", len(wb.sheets), len(wb.vba_modules))

    # LLM 注釈 (Phase C). 個別失敗は飲み込む実装なので全体は止まらない。
    annotated_wb = annotate_workbook(wb, llm)
    storage.save_workbook(job_id, annotated_wb)

    spec_md = generate_spec(annotated_wb, idx)

    try:
        storage.save_spec(job_id, spec_md)
        storage.update_status(job_id, "analyzed")
        logger.info("analyze completed: spec_chars=%d status=analyzed", len(spec_md))
    except Exception as e:
        logger.exception("analyze failed to persist spec")
        storage.update_status(job_id, "failed")
        raise HTTPException(status_code=500, detail=str(e)) from e

    return {"status": "ok"}
