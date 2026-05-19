"""POST /extract — アップロード + Core 抽出.

並行性メモ:
  openpyxl / olevba / cells.db 書き出しは数秒〜数分かかる同期処理。`async def`
  本体で直接呼ぶと event loop が止まり、別ユーザーの軽量リクエスト (/jobs,
  /health, /spec 等) も巻き添えで遅延する。そのため `_run_extraction` に
  まとめて切り出し、`asyncio.to_thread` で threadpool に逃がしている。
"""

from __future__ import annotations

import asyncio
import logging
import os

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile

from backend.dependencies import get_storage
from backend.logging_config import set_job_id
from backend.storage import Storage
from core.exceptions import ExtractionError
from core.extractors.cells import extract_cells_to_sqlite
from core.extractors.vba import extract_vba
from core.extractors.workbook import extract_workbook
from core.reference_index import build_reference_index

logger = logging.getLogger(__name__)
router = APIRouter()

# SPEC.md §8: 1ファイル 50MB を上限の目安として想定。環境変数で上書き可能。
MAX_UPLOAD_BYTES = int(os.environ.get("MAX_UPLOAD_BYTES", 50 * 1024 * 1024))
_UPLOAD_CHUNK_SIZE = 1024 * 1024  # 1 MiB


async def _read_capped(file: UploadFile, limit: int) -> bytes:
    """UploadFile を limit までチャンク読みする. 超過したら HTTPException(413).

    Content-Length 由来の `file.size` で早期判定し、ヘッダ詐称や chunked 転送に
    備えて読みながら累積サイズも検証する。
    """
    if file.size is not None and file.size > limit:
        raise HTTPException(
            status_code=413,
            detail=f"file too large (max {limit} bytes, got {file.size})",
        )

    buf = bytearray()
    while True:
        chunk = await file.read(_UPLOAD_CHUNK_SIZE)
        if not chunk:
            break
        buf.extend(chunk)
        if len(buf) > limit:
            raise HTTPException(
                status_code=413,
                detail=f"file too large (max {limit} bytes)",
            )
    return bytes(buf)


def _run_extraction(storage: Storage, job_id: str, filename: str) -> None:
    """ジョブの重い抽出処理本体. threadpool で呼ばれる前提.

    `set_job_id` の ContextVar はこのスレッド内のログにだけ伝播する.
    本体の HTTP リクエスト処理スレッドとは別だが、ログ相関 ID は維持される.
    """
    with set_job_id(job_id):
        try:
            path = storage.get_original_path(job_id)
            wb = extract_workbook(path)
            # ストレージ上の保存名 (original.xlsx 等) ではなく、ユーザー提供のファイル名を保持する
            wb.filename = filename
            wb.vba_modules = extract_vba(path)
            idx = build_reference_index(wb)
            storage.save_workbook(job_id, wb)
            storage.save_references(job_id, idx)
            logger.info(
                "workbook extracted: sheets=%d vba_modules=%d ref_keys=%d",
                len(wb.sheets),
                len(wb.vba_modules),
                len(idx.refs),
            )

            # cells.db (全セル生データ) を生成. .xls は openpyxl で読めないのでスキップ。
            if path.suffix.lower() != ".xls":
                try:
                    cell_count = extract_cells_to_sqlite(path, storage.cells_db_path(job_id))
                    logger.info("cells.db built: cells=%d", cell_count)
                except ExtractionError as e:
                    # cells.db 失敗は致命的ではない (チャットでの参照のみに使う)
                    logger.warning("cells.db build failed: %s", e)

            storage.update_status(job_id, "extracted")
            logger.info("extract completed: status=extracted")
        except ExtractionError:
            logger.exception("extract failed (ExtractionError)")
            storage.update_status(job_id, "failed")
            raise
        except Exception:
            logger.exception("extract failed (unexpected)")
            storage.update_status(job_id, "failed")
            raise


@router.post("/extract")
async def extract(
    file: UploadFile = File(...),
    storage: Storage = Depends(get_storage),
) -> dict[str, str]:
    """ファイルを保存し、Core 抽出を実行して job_id を返す.

    重い同期処理 (openpyxl / olevba / cells.db) は `asyncio.to_thread` に
    逃がして event loop を解放する. これにより別ユーザーのリクエストが
    待たされなくなる.
    """
    if not file.filename:
        raise HTTPException(status_code=400, detail="filename is required")

    data = await _read_capped(file, MAX_UPLOAD_BYTES)
    if not data:
        raise HTTPException(status_code=400, detail="empty file")

    # create_job (ディレクトリ作成 + バイト書き出し) も blocking なので to_thread に逃がす
    meta = await asyncio.to_thread(storage.create_job, file.filename, data)
    job_id = meta.job_id

    logger.info("extract started: filename=%s size=%d bytes", meta.filename, len(data))
    try:
        await asyncio.to_thread(_run_extraction, storage, job_id, meta.filename)
    except ExtractionError as e:
        raise HTTPException(status_code=422, detail=f"extraction failed: {e}") from e
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"internal error: {e}") from e

    return {"job_id": job_id}
