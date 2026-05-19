"""POST /analyze/{job_id} — LLM 注釈 + 設計書生成.

並行性メモ:
  annotate_workbook は LLM (HTTP 同期) を多数回呼ぶため最も長い処理 (数秒〜
  数分). generate_spec も CPU 寄り. これらを `async def` 本体で直接呼ぶと
  event loop が止まるので、`asyncio.to_thread` で threadpool に逃がす.
"""

from __future__ import annotations

import asyncio
import logging

from fastapi import APIRouter, Depends, HTTPException

from backend.annotators import annotate_workbook
from backend.dependencies import get_llm_client, get_storage
from backend.llm_client import LLMClient
from backend.storage import JobNotFoundError, Storage
from core.models import ReferenceIndex, Workbook
from core.spec_generator import generate_spec

logger = logging.getLogger(__name__)
router = APIRouter()


def _run_analysis(
    storage: Storage,
    llm: LLMClient,
    job_id: str,
    wb: Workbook,
    idx: ReferenceIndex,
) -> int:
    """LLM 注釈 + 設計書生成 + 永続化を行う. threadpool で呼ばれる前提.

    Returns:
        生成された spec.md の文字数 (ログ用).
    """
    annotated_wb = annotate_workbook(wb, llm)
    storage.save_workbook(job_id, annotated_wb)
    spec_md = generate_spec(annotated_wb, idx)
    storage.save_spec(job_id, spec_md)
    storage.update_status(job_id, "analyzed")
    return len(spec_md)


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

    LLM ループは threadpool で実行するので他リクエストはブロックされない.
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

    try:
        spec_chars = await asyncio.to_thread(
            _run_analysis, storage, llm, job_id, wb, idx
        )
    except Exception as e:
        logger.exception("analyze failed")
        # ベストエフォートでステータスを更新 (再度 to_thread で safe に呼ぶ)
        try:
            await asyncio.to_thread(storage.update_status, job_id, "failed")
        except Exception:  # noqa: BLE001
            pass
        raise HTTPException(status_code=500, detail=str(e)) from e

    logger.info("analyze completed: spec_chars=%d status=analyzed", spec_chars)
    return {"status": "ok"}
