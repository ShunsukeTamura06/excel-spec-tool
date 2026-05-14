"""GET /diagrams/{job_id} — シート依存グラフと VBA コールグラフを返す.

Workbook 抽出済み (extract 完了後) であれば呼べる。LLM 注釈 (analyze) は不要。
"""

from __future__ import annotations

import logging
from typing import Any

from fastapi import APIRouter, Depends, HTTPException

from backend.dependencies import get_storage
from backend.storage import JobNotFoundError, Storage
from core.diagrams import build_diagrams

logger = logging.getLogger(__name__)

router = APIRouter()


@router.get("/diagrams/{job_id}")
async def get_diagrams(
    job_id: str,
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """ジョブのダイアグラム (sheet_deps / vba_calls) を返す.

    Returns:
        {
            "sheet_deps": {"kind": "sheet_deps", "nodes": [...], "edges": [...]},
            "vba_calls":  {"kind": "vba_calls",  "nodes": [...], "edges": [...]},
        }
    """
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

    diagrams = build_diagrams(wb)
    logger.info(
        "diagrams built: sheets=%d sheet_edges=%d procs=%d call_edges=%d",
        len(diagrams.sheet_deps.nodes),
        len(diagrams.sheet_deps.edges),
        len(diagrams.vba_calls.nodes),
        len(diagrams.vba_calls.edges),
    )
    return diagrams.model_dump()
