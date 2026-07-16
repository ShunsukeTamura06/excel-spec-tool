"""安全パターンの自動修正を実ファイルに適用し、新ジョブを作って自己検証する (S2 増分1)。

- POST /jobs/{job_id}/named-range-fix — 名前定義修正
- POST /jobs/{job_id}/formula-fix — 固定参照置換 / 数式範囲拡張

いずれも人間が画面のボタンを押した時だけ呼ばれる。LLM の tool loop からは呼ばれない
(黙って変更しない、docs/VISION.ja.md §4.2)。
"""

from __future__ import annotations

import asyncio
import logging
from pathlib import Path
from typing import Literal

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel

from backend.dependencies import get_storage
from backend.routes.extract import _run_extraction
from backend.storage import JobNotFoundError, Storage
from core.change_record import ChangeExecutionRecord
from core.exceptions import (
    DiffError,
    ExtractionError,
    FormulaFixError,
    MutationProviderError,
    NamedRangeFixError,
    ProviderUnavailableError,
    UnsupportedMutationError,
)
from core.models import ReferenceIndex, Workbook, WorkbookDiff
from core.mutation import (
    FixedRefReplaceOperation,
    MutationPlan,
    MutationProvider,
    MutationResult,
    NamedRangeSetOperation,
    OpenPyxlMutationProvider,
    ProviderName,
    RangeExpansionOperation,
    SafeChangePlan,
    build_safe_change_plan,
    propose_mutation,
)
from core.officecli_provider import OfficeCliMutationProvider
from core.verification import verify_expected_diff
from core.workbook_diff import diff_workbooks

logger = logging.getLogger(__name__)
router = APIRouter()


class NamedRangeFixRequest(BaseModel):
    """POST /jobs/{job_id}/named-range-fix のリクエストボディ."""

    name: str
    new_refers_to: str
    provider: ProviderName = "openpyxl"


class FormulaFixRequest(BaseModel):
    """POST /jobs/{job_id}/formula-fix のリクエストボディ."""

    kind: Literal["fixed_ref_replace", "range_expansion"]
    old_ref: str
    new_ref: str
    provider: ProviderName = "openpyxl"


class ChangePlanRequest(BaseModel):
    """一般ユーザー向け変更計画のリクエストボディ."""

    kind: Literal["range_expansion"]
    old_ref: str
    new_ref: str


class ExecuteChangePlanRequest(BaseModel):
    """確認済みの変更計画をそのまま実行するリクエストボディ."""

    plan: MutationPlan


def _provider_for(name: ProviderName) -> MutationProvider:
    """APIで選択された変更プロバイダーを返す."""

    if name == "officecli":
        return OfficeCliMutationProvider()
    return OpenPyxlMutationProvider()


def _load_before(storage: Storage, job_id: str) -> tuple[Workbook, Path, ReferenceIndex, str]:
    """before 側 (元ジョブ) の一式をロードし、HTTP エラーに変換する."""
    try:
        before_wb = storage.load_workbook(job_id)
        before_path = storage.get_original_path(job_id)
        before_index = storage.load_references(job_id)
        meta = storage.get_meta(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail=f"job not found: {e}") from e
    except FileNotFoundError as e:
        raise HTTPException(
            status_code=409, detail=f"workbook not extracted yet; call /extract first: {e}"
        ) from e
    return before_wb, before_path, before_index, meta.filename


def _apply_and_extract(
    storage: Storage,
    source_job_id: str,
    filename: str,
    provider: MutationProvider,
    plan: MutationPlan,
) -> tuple[str, MutationResult]:
    """実ファイルへの書き込み + 新ジョブ作成 + 抽出を1関数にまとめる (threadpool 前提).

    Returns:
        新しく作られたジョブの job_id とプロバイダー実行結果.
    """
    source_path = storage.get_original_path(source_job_id)
    data = source_path.read_bytes()
    # 元と同じバイト列でまず新ジョブを作る (meta.json 生成・拡張子判定を再利用するため)。
    # その後 apply_fn で新ジョブの original.* を書き換える。
    new_meta = storage.create_job(filename, data)
    new_path = storage.get_original_path(new_meta.job_id)
    try:
        provider_result = provider.apply(plan, source_path, new_path)
        # create_job時点の指紋は元ファイルのコピーを指すため、変更後成果物へ更新する。
        storage.refresh_original_fingerprint(new_meta.job_id)
        _run_extraction(storage, new_meta.job_id, filename)
    except Exception:
        # 適用/抽出に失敗した場合、作りかけの新ジョブを残すとジョブ一覧に
        # 「uploaded のまま進まない」孤児が溜まるため後始末する。
        storage.delete_job(new_meta.job_id)
        raise
    return new_meta.job_id, provider_result


async def _self_verify(
    storage: Storage,
    before_wb: Workbook,
    before_path: Path,
    before_index: ReferenceIndex,
    new_job_id: str,
) -> WorkbookDiff:
    """元ジョブ (before) と新ジョブ (after) を diff_workbooks で比較する."""
    try:
        after_wb = storage.load_workbook(new_job_id)
        after_path = storage.get_original_path(new_job_id)
        diff = await asyncio.to_thread(
            diff_workbooks, before_path, after_path, before_wb, after_wb, before_index
        )
    except DiffError as e:
        raise HTTPException(status_code=422, detail=f"self-verification diff failed: {e}") from e
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"internal error: {e}") from e
    return diff


async def _execute_plan(
    storage: Storage,
    before_wb: Workbook,
    before_path: Path,
    before_index: ReferenceIndex,
    filename: str,
    plan: MutationPlan,
) -> dict[str, object]:
    """変更計画を適用し、期待差分とのpolicy照合と監査保存まで実行する."""

    expected_diff = propose_mutation(plan, before_wb, before_index)
    provider = _provider_for(plan.requested_provider)
    new_job_id, provider_result = await asyncio.to_thread(
        _apply_and_extract,
        storage,
        plan.source_job_id,
        filename,
        provider,
        plan,
    )
    try:
        actual_diff = await _self_verify(storage, before_wb, before_path, before_index, new_job_id)
    except HTTPException:
        await asyncio.to_thread(storage.update_status, new_job_id, "failed")
        raise
    verification = verify_expected_diff(expected_diff, actual_diff)
    try:
        source_meta = await asyncio.to_thread(
            storage.refresh_original_fingerprint, plan.source_job_id
        )
        result_meta = storage.get_meta(new_job_id)
        if source_meta.file_sha256 is None or result_meta.file_sha256 is None:
            raise OSError("artifact fingerprint is missing")
        record = ChangeExecutionRecord(
            source_job_id=plan.source_job_id,
            result_job_id=new_job_id,
            source_file_sha256=source_meta.file_sha256,
            result_file_sha256=result_meta.file_sha256,
            plan=plan,
            provider_result=provider_result,
            expected_diff=expected_diff,
            actual_diff=actual_diff,
            verification=verification,
        )
        await asyncio.to_thread(storage.save_verification, new_job_id, record)
    except OSError as exc:
        await asyncio.to_thread(storage.update_status, new_job_id, "failed")
        raise HTTPException(
            status_code=500,
            detail=f"failed to persist verification evidence: {exc}",
        ) from exc

    if verification.status == "failed":
        await asyncio.to_thread(storage.update_status, new_job_id, "failed")
        raise HTTPException(
            status_code=409,
            detail={
                "message": "verification policy rejected the mutated workbook",
                "new_job_id": new_job_id,
                "verification": verification.model_dump(mode="json"),
            },
        )

    return {
        "new_job_id": new_job_id,
        "diff": actual_diff.model_dump(by_alias=True),
        "verification": verification.model_dump(mode="json"),
        "plan": plan.model_dump(mode="json"),
        "provider": provider_result.model_dump(mode="json"),
    }


@router.get("/mutation-providers")
async def list_mutation_providers() -> dict[str, object]:
    """変更プロバイダーの利用可否と対応範囲を返す."""

    openpyxl_capability = OpenPyxlMutationProvider().capability()
    officecli_capability = await asyncio.to_thread(OfficeCliMutationProvider().capability)
    return {
        "providers": [
            openpyxl_capability.model_dump(mode="json"),
            officecli_capability.model_dump(mode="json"),
        ]
    }


@router.get("/jobs/{job_id}/verification")
async def get_verification_record(
    job_id: str,
    storage: Storage = Depends(get_storage),
) -> dict[str, object]:
    """変更後ジョブに保存された検証監査レコードを返す."""

    try:
        record = await asyncio.to_thread(storage.load_verification, job_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {exc}") from exc
    except JobNotFoundError as exc:
        raise HTTPException(status_code=404, detail=f"job not found: {exc}") from exc
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail="verification record not found") from exc
    return {"verification_record": record.model_dump(mode="json")}


@router.post("/jobs/{job_id}/change-plan", response_model=SafeChangePlan)
async def create_change_plan(
    job_id: str,
    body: ChangePlanRequest,
    storage: Storage = Depends(get_storage),
) -> SafeChangePlan:
    """原本を書き換えず、適用前に確認する範囲拡張計画を作る."""

    before_wb, _, before_index, _ = _load_before(storage, job_id)
    plan = MutationPlan(
        source_job_id=job_id,
        requested_provider="openpyxl",
        operation=RangeExpansionOperation(old_ref=body.old_ref, new_ref=body.new_ref),
    )
    try:
        return build_safe_change_plan(plan, before_wb, before_index)
    except FormulaFixError as exc:
        raise HTTPException(status_code=422, detail=f"change plan failed: {exc}") from exc


@router.post("/jobs/{job_id}/change-plan/execute")
async def execute_change_plan(
    job_id: str,
    body: ExecuteChangePlanRequest,
    storage: Storage = Depends(get_storage),
) -> dict[str, object]:
    """画面で確認した同一計画を適用し、新規ファイルとして検証する."""

    plan = body.plan
    if plan.source_job_id != job_id:
        raise HTTPException(status_code=400, detail="plan source_job_id does not match URL job_id")
    if plan.requested_provider != "openpyxl":
        raise HTTPException(
            status_code=422,
            detail="general-user change plans currently support only the openpyxl provider",
        )
    if not isinstance(plan.operation, RangeExpansionOperation):
        raise HTTPException(
            status_code=422,
            detail="general-user change plans currently support only range expansion",
        )

    before_wb, before_path, before_index, filename = _load_before(storage, job_id)
    try:
        result = await _execute_plan(
            storage,
            before_wb,
            before_path,
            before_index,
            filename,
            plan,
        )
    except FormulaFixError as exc:
        raise HTTPException(status_code=422, detail=f"change execution failed: {exc}") from exc
    except ExtractionError as exc:
        raise HTTPException(
            status_code=422,
            detail=f"extraction of new file failed: {exc}",
        ) from exc
    except ProviderUnavailableError as exc:
        raise HTTPException(
            status_code=503,
            detail=f"mutation provider unavailable: {exc}",
        ) from exc
    except UnsupportedMutationError as exc:
        raise HTTPException(status_code=422, detail=f"unsupported mutation: {exc}") from exc
    except MutationProviderError as exc:
        raise HTTPException(status_code=422, detail=f"mutation provider failed: {exc}") from exc
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"internal error: {exc}") from exc

    logger.info(
        "confirmed change plan executed: source=%s new_job=%s plan=%s verification=%s",
        job_id,
        result["new_job_id"],
        plan.plan_id,
        result["verification"],
    )
    return result


@router.post("/jobs/{job_id}/named-range-fix")
async def apply_named_range_fix_route(
    job_id: str,
    body: NamedRangeFixRequest,
    storage: Storage = Depends(get_storage),
) -> dict[str, object]:
    """名前定義修正を適用し、新ジョブを作って diff で自己検証する.

    手順:
      1. 元ジョブの原本ファイルをコピーし、名前定義を書き換えた新ファイルとして
         新ジョブを作成、フル抽出パイプラインを流す (extract.py の _run_extraction 再利用)。
      2. 元ジョブ (before) と新ジョブ (after) を diff_workbooks で比較し、
         意図した変更 (name の refers_to) だけが起きているかを自己検証する。
    """
    before_wb, before_path, before_index, filename = _load_before(storage, job_id)
    plan = MutationPlan(
        source_job_id=job_id,
        requested_provider=body.provider,
        operation=NamedRangeSetOperation(name=body.name, new_refers_to=body.new_refers_to),
    )

    try:
        result = await _execute_plan(storage, before_wb, before_path, before_index, filename, plan)
    except NamedRangeFixError as e:
        raise HTTPException(status_code=422, detail=f"named range fix failed: {e}") from e
    except ExtractionError as e:
        raise HTTPException(status_code=422, detail=f"extraction of new file failed: {e}") from e
    except ProviderUnavailableError as e:
        raise HTTPException(status_code=503, detail=f"mutation provider unavailable: {e}") from e
    except UnsupportedMutationError as e:
        raise HTTPException(status_code=422, detail=f"unsupported mutation: {e}") from e
    except MutationProviderError as e:
        raise HTTPException(status_code=422, detail=f"mutation provider failed: {e}") from e
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"internal error: {e}") from e

    logger.info(
        "named range fix applied: source=%s new_job=%s name=%s verification=%s",
        job_id,
        result["new_job_id"],
        body.name,
        result["verification"],
    )
    return result


@router.post("/jobs/{job_id}/formula-fix")
async def apply_formula_fix_route(
    job_id: str,
    body: FormulaFixRequest,
    storage: Storage = Depends(get_storage),
) -> dict[str, object]:
    """固定参照置換 / 数式範囲拡張を適用し、新ジョブを作って diff で自己検証する.

    named-range-fix と同じ流れ: 修正済みの新ファイルで新ジョブを作り、
    before/after を diff_workbooks で比較して意図した数式変更だけが
    起きているかを自己検証する。
    """
    before_wb, before_path, before_index, filename = _load_before(storage, job_id)
    operation = (
        RangeExpansionOperation(old_ref=body.old_ref, new_ref=body.new_ref)
        if body.kind == "range_expansion"
        else FixedRefReplaceOperation(old_ref=body.old_ref, new_ref=body.new_ref)
    )
    plan = MutationPlan(
        source_job_id=job_id,
        requested_provider=body.provider,
        operation=operation,
    )

    try:
        result = await _execute_plan(storage, before_wb, before_path, before_index, filename, plan)
    except FormulaFixError as e:
        raise HTTPException(status_code=422, detail=f"formula fix failed: {e}") from e
    except ExtractionError as e:
        raise HTTPException(status_code=422, detail=f"extraction of new file failed: {e}") from e
    except ProviderUnavailableError as e:
        raise HTTPException(status_code=503, detail=f"mutation provider unavailable: {e}") from e
    except UnsupportedMutationError as e:
        raise HTTPException(status_code=422, detail=f"unsupported mutation: {e}") from e
    except MutationProviderError as e:
        raise HTTPException(status_code=422, detail=f"mutation provider failed: {e}") from e
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"internal error: {e}") from e

    logger.info(
        "formula fix applied: source=%s new_job=%s kind=%s old=%s new=%s verification=%s",
        job_id,
        result["new_job_id"],
        body.kind,
        body.old_ref,
        body.new_ref,
        result["verification"],
    )
    return result
