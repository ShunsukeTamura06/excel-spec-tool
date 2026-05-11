"""GET /cells/{job_id}/range と /find — セル生データの取得.

設計書には載せていない「全セル」を、必要な時に範囲指定または値検索で
取得するためのエンドポイント。LLM が tool として呼ぶことを想定している。
"""

from __future__ import annotations

from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Query

from backend.dependencies import get_storage
from backend.storage import JobNotFoundError, Storage

router = APIRouter()


@router.get("/cells/{job_id}/range")
async def get_range(
    job_id: str,
    sheet: str = Query(..., description="シート名"),
    range: str = Query(..., description="Excel 範囲 (例: 'A6:F10' or 'A6')"),
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """指定されたシートの指定範囲のセル値を返す."""
    try:
        # job_id 検証は内部で行われる
        return storage.get_cells_range(job_id, sheet=sheet, range_str=range)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e
    except FileNotFoundError as e:
        raise HTTPException(
            status_code=409,
            detail="cells.db not built; call /extract on an xlsx/xlsm first",
        ) from e


@router.get("/cells/{job_id}/find")
async def find(
    job_id: str,
    q: str = Query(..., description="検索文字列 (value への LIKE 検索)"),
    sheet: str | None = Query(None, description="シート名で絞り込み (任意)"),
    limit: int = Query(20, ge=1, le=200),
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """セルの value を部分一致検索する."""
    try:
        results = storage.find_cells(job_id, query=q, sheet=sheet, limit=limit)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e
    except FileNotFoundError as e:
        raise HTTPException(
            status_code=409,
            detail="cells.db not built; call /extract on an xlsx/xlsm first",
        ) from e
    return {"matches": results, "count": len(results)}
