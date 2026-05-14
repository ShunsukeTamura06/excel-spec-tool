"""GET /workbook/{job_id} — 抽出済み Workbook 構造を返す.

設計書ページのシート/VBA タブが構造を直接読むために使う. /spec が返す
spec.md は人間向けの読み物だが、こちらは UI が drilldown する用の
機械可読データ.
"""

from __future__ import annotations

import logging
from typing import Any

from fastapi import APIRouter, Depends, HTTPException

from backend.dependencies import get_storage
from backend.storage import JobNotFoundError, Storage

logger = logging.getLogger(__name__)

router = APIRouter()


@router.get("/workbook/{job_id}")
async def get_workbook(
    job_id: str,
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """抽出済み Workbook (sheets / vba_modules / external_links) を返す."""
    try:
        wb = storage.load_workbook(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e
    except FileNotFoundError as e:
        raise HTTPException(
            status_code=409,
            detail="workbook not extracted yet; call /extract first",
        ) from e

    return wb.model_dump()
