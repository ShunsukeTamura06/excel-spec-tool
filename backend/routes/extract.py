"""POST /extract — アップロード + Core 抽出."""

from __future__ import annotations

import logging

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile

from backend.dependencies import get_storage
from backend.storage import Storage
from core.exceptions import ExtractionError
from core.extractors.cells import extract_cells_to_sqlite
from core.extractors.vba import extract_vba
from core.extractors.workbook import extract_workbook
from core.reference_index import build_reference_index

logger = logging.getLogger(__name__)
router = APIRouter()


@router.post("/extract")
async def extract(
    file: UploadFile = File(...),
    storage: Storage = Depends(get_storage),
) -> dict[str, str]:
    """ファイルを保存し、Core 抽出を実行して job_id を返す."""
    if not file.filename:
        raise HTTPException(status_code=400, detail="filename is required")

    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="empty file")

    meta = storage.create_job(file.filename, data)
    job_id = meta.job_id

    try:
        path = storage.get_original_path(job_id)
        wb = extract_workbook(path)
        # ストレージ上の保存名 (original.xlsx 等) ではなく、ユーザー提供のファイル名を保持する
        wb.filename = meta.filename
        wb.vba_modules = extract_vba(path)
        idx = build_reference_index(wb)
        storage.save_workbook(job_id, wb)
        storage.save_references(job_id, idx)

        # cells.db (全セル生データ) を生成. .xls は openpyxl で読めないのでスキップ。
        if path.suffix.lower() != ".xls":
            try:
                extract_cells_to_sqlite(path, storage.cells_db_path(job_id))
            except ExtractionError as e:
                # cells.db 失敗は致命的ではない (チャットでの参照のみに使う)
                logger.warning("cells.db build failed for %s: %s", job_id, e)

        storage.update_status(job_id, "extracted")
    except ExtractionError as e:
        logger.exception("Extraction failed for %s", job_id)
        storage.update_status(job_id, "failed")
        raise HTTPException(status_code=422, detail=f"extraction failed: {e}") from e
    except Exception as e:
        logger.exception("Unexpected error during extraction for %s", job_id)
        storage.update_status(job_id, "failed")
        raise HTTPException(status_code=500, detail=f"internal error: {e}") from e

    return {"job_id": job_id}
