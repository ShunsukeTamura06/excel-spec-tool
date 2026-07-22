"""VBA変更パッケージの生成と、Windows適用後.xlsmの静的検証."""

from __future__ import annotations

import asyncio
import io
import logging
from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from backend.dependencies import get_storage
from backend.routes.extract import MAX_UPLOAD_BYTES, _read_capped, _run_extraction
from backend.routes.refactor import _load_before, _self_verify
from backend.storage import JobNotFoundError, Storage
from core.change_record import ChangeExecutionRecord
from core.exceptions import DiffError, ExtractionError, VbaChangeError
from core.mutation import (
    MutationResult,
    SafeChangePlan,
    VbaProcedureReplaceOperation,
    build_safe_change_plan,
    propose_mutation,
)
from core.vba_package import build_vba_change_package
from core.verification import verify_expected_diff

logger = logging.getLogger(__name__)
router = APIRouter()


class VbaPackageRequest(BaseModel):
    """Windows実行パッケージ生成リクエスト.

    計画本体はチャットの propose_vba_procedure_replace が保存した内容を
    plan_id で引き当てる (クライアントが計画内容を送っても信頼しない)。
    """

    plan_id: str


def _load_pending_vba_plan(
    storage: Storage,
    job_id: str,
    plan_id: str,
) -> tuple[Path, SafeChangePlan]:
    """保存済みのVBA置換計画を plan_id で読み出し、原本パスと安全計画を返す."""

    try:
        stored_plan = storage.load_pending_plan(job_id, plan_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=f"invalid plan_id: {exc}") from exc
    except FileNotFoundError as exc:
        raise HTTPException(
            status_code=404,
            detail=f"pending change plan not found or already used: {exc}",
        ) from exc
    plan = stored_plan.plan
    if plan.source_job_id != job_id:
        raise HTTPException(status_code=400, detail="plan source_job_id does not match URL job_id")
    if plan.requested_provider != "windows_vbide":
        raise HTTPException(
            status_code=422,
            detail="VBA procedure replacement requires the windows_vbide provider",
        )
    if not isinstance(plan.operation, VbaProcedureReplaceOperation):
        raise HTTPException(
            status_code=422,
            detail="VBA change package requires vba_procedure_replace",
        )
    before_wb, before_path, before_index, _ = _load_before(storage, job_id)
    if before_path.suffix.lower() != ".xlsm":
        raise HTTPException(status_code=422, detail="VBA change package requires an .xlsm file")
    try:
        safe_plan = build_safe_change_plan(plan, before_wb, before_index)
    except VbaChangeError as exc:
        raise HTTPException(status_code=422, detail=f"VBA change plan failed: {exc}") from exc
    return before_path, safe_plan


@router.post("/jobs/{job_id}/vba-change/package")
async def download_vba_change_package(
    job_id: str,
    body: VbaPackageRequest,
    storage: Storage = Depends(get_storage),
) -> StreamingResponse:
    """原本を変更せず、Windows Excel/VBIDE用ZIPパッケージを返す.

    このステップは何も適用・確定しない (ZIPを作るだけ)。計画の消費は
    後続の /vba-change/verify で行う。
    """

    try:
        source_path, safe_plan = _load_pending_vba_plan(storage, job_id, body.plan_id)
        package = await asyncio.to_thread(
            build_vba_change_package,
            source_path,
            safe_plan.plan,
        )
    except (JobNotFoundError, FileNotFoundError) as exc:
        raise HTTPException(status_code=404, detail=f"source workbook not found: {exc}") from exc
    except VbaChangeError as exc:
        raise HTTPException(status_code=422, detail=f"VBA package failed: {exc}") from exc
    operation = safe_plan.plan.operation
    assert isinstance(operation, VbaProcedureReplaceOperation)
    logger.info(
        "VBA change package generated: job=%s module=%s procedure=%s bytes=%d",
        job_id,
        operation.module_name,
        operation.procedure_name,
        len(package),
    )
    return StreamingResponse(
        io.BytesIO(package),
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="xlblueprint_vba_change.zip"'},
    )


@router.post("/jobs/{job_id}/vba-change/verify")
async def verify_vba_changed_workbook(
    job_id: str,
    file: UploadFile = File(...),
    plan_id: str = Form(...),
    storage: Storage = Depends(get_storage),
) -> dict[str, object]:
    """Windowsで生成された.xlsmを再抽出し、期待VBA差分と厳密照合する.

    plan_id は propose 時に保存された計画のみを引き当てる。この呼び出しで
    計画を消費 (削除) し、同じ計画を使った再照合 (リプレイ) を防ぐ。
    """

    if not file.filename or Path(file.filename).suffix.lower() != ".xlsm":
        raise HTTPException(status_code=422, detail="revised workbook must be an .xlsm file")
    _, safe_plan = _load_pending_vba_plan(storage, job_id, plan_id)
    storage.consume_pending_plan(job_id, plan_id)
    data = await _read_capped(file, MAX_UPLOAD_BYTES)
    if not data:
        raise HTTPException(status_code=400, detail="empty revised workbook")

    before_wb, before_path, before_index, _ = _load_before(storage, job_id)
    expected_diff = propose_mutation(safe_plan.plan, before_wb, before_index)
    result_meta = await asyncio.to_thread(storage.create_job, file.filename, data)
    try:
        await asyncio.to_thread(
            _run_extraction,
            storage,
            result_meta.job_id,
            file.filename,
        )
        actual_diff = await _self_verify(
            storage,
            before_wb,
            before_path,
            before_index,
            result_meta.job_id,
        )
    except (ExtractionError, DiffError) as exc:
        await asyncio.to_thread(storage.update_status, result_meta.job_id, "failed")
        raise HTTPException(status_code=422, detail=f"VBA result extraction failed: {exc}") from exc
    except HTTPException:
        await asyncio.to_thread(storage.update_status, result_meta.job_id, "failed")
        raise
    except Exception as exc:
        await asyncio.to_thread(storage.update_status, result_meta.job_id, "failed")
        raise HTTPException(status_code=500, detail=f"internal error: {exc}") from exc

    verification = verify_expected_diff(expected_diff, actual_diff)
    source_meta = await asyncio.to_thread(storage.refresh_original_fingerprint, job_id)
    result_meta = storage.get_meta(result_meta.job_id)
    if source_meta.file_sha256 is None or result_meta.file_sha256 is None:
        await asyncio.to_thread(storage.update_status, result_meta.job_id, "failed")
        raise HTTPException(status_code=500, detail="artifact fingerprint is missing")
    provider_result = MutationResult(
        provider="windows_vbide",
        provider_version="Microsoft Excel/VBIDE user execution",
        operation="vba_procedure_replace",
        changed_count=1,
    )
    record = ChangeExecutionRecord(
        source_job_id=job_id,
        result_job_id=result_meta.job_id,
        source_file_sha256=source_meta.file_sha256,
        result_file_sha256=result_meta.file_sha256,
        plan=safe_plan.plan,
        provider_result=provider_result,
        expected_diff=expected_diff,
        actual_diff=actual_diff,
        verification=verification,
    )
    await asyncio.to_thread(storage.save_verification, result_meta.job_id, record)

    if verification.status == "failed":
        await asyncio.to_thread(storage.update_status, result_meta.job_id, "failed")
        raise HTTPException(
            status_code=409,
            detail={
                "message": "verification policy rejected the VBA-modified workbook",
                "new_job_id": result_meta.job_id,
                "verification": verification.model_dump(mode="json"),
            },
        )
    logger.info(
        "VBA changed workbook verified: source=%s result=%s status=%s",
        job_id,
        result_meta.job_id,
        verification.status,
    )
    return {
        "new_job_id": result_meta.job_id,
        "diff": actual_diff.model_dump(by_alias=True),
        "verification": verification.model_dump(mode="json"),
        "plan": safe_plan.plan.model_dump(mode="json"),
        "provider": provider_result.model_dump(mode="json"),
    }
