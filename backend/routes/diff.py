"""GET /diff — 2ジョブ間の構造差分 + 波及範囲レポート (P1 安全ゲート増分1).

Excel COM での再計算・マクロ実行によるテストはここでは扱わない (別増分)。
セル抽出 (extract_cells_to_sqlite 経由) が同期・重めの I/O のため、
extract/analyze と同じ方針で `asyncio.to_thread` に逃がす。
"""

from __future__ import annotations

import asyncio
import logging
from pathlib import Path

from fastapi import APIRouter, Depends, HTTPException, Query

from backend.dependencies import get_storage
from backend.storage import JobNotFoundError, Storage
from core.exceptions import DiffError
from core.models import Workbook
from core.workbook_diff import diff_workbooks

logger = logging.getLogger(__name__)
router = APIRouter()


def _load_job(storage: Storage, job_id: str, label: str) -> tuple[Workbook, Path]:
    """ジョブの Workbook と原本パスをロードする. 失敗時は label 付きで再送出する."""
    try:
        wb = storage.load_workbook(job_id)
        path = storage.get_original_path(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid {label}_job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail=f"{label} job not found") from e
    except FileNotFoundError as e:
        raise HTTPException(
            status_code=409,
            detail=f"{label} workbook not extracted yet; call /extract first",
        ) from e
    return wb, path


@router.get("/diff")
async def get_diff(
    before_job_id: str = Query(..., description="「前」バージョンの job_id"),
    after_job_id: str = Query(..., description="「後」バージョンの job_id"),
    storage: Storage = Depends(get_storage),
) -> dict[str, object]:
    """2ジョブ間の構造差分 (セル/名前定義/条件付き書式/入力規則/グラフ/ピボット/VBA)

    と波及範囲 (blast radius)・既存リスクをまとめて返す。
    """
    if before_job_id == after_job_id:
        raise HTTPException(
            status_code=400,
            detail="before_job_id and after_job_id must differ",
        )

    before_wb, before_path = _load_job(storage, before_job_id, "before")
    after_wb, after_path = _load_job(storage, after_job_id, "after")

    try:
        before_index = storage.load_references(before_job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid before_job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="before job not found") from e
    except FileNotFoundError as e:
        raise HTTPException(
            status_code=409,
            detail="before references not built yet; call /extract first",
        ) from e

    logger.info("diff started: before=%s after=%s", before_job_id, after_job_id)
    try:
        diff = await asyncio.to_thread(
            diff_workbooks, before_path, after_path, before_wb, after_wb, before_index
        )
    except DiffError as e:
        raise HTTPException(status_code=422, detail=f"diff failed: {e}") from e
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"internal error: {e}") from e

    logger.info("diff completed: is_empty=%s", diff.is_empty())
    return {"diff": diff.model_dump(by_alias=True)}
