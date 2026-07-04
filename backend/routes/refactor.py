"""安全パターンの自動修正を実ファイルに適用し、新ジョブを作って自己検証する (S2 増分1)。

- POST /jobs/{job_id}/named-range-fix — 名前定義修正
- POST /jobs/{job_id}/formula-fix — 固定参照置換 / 数式範囲拡張

いずれも人間が画面のボタンを押した時だけ呼ばれる。LLM の tool loop からは呼ばれない
(黙って変更しない、docs/VISION.ja.md §4.2)。
"""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Callable
from pathlib import Path
from typing import Literal

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel

from backend.dependencies import get_storage
from backend.routes.extract import _run_extraction
from backend.storage import JobNotFoundError, Storage
from core.exceptions import DiffError, ExtractionError, FormulaFixError, NamedRangeFixError
from core.formula_fix import apply_fixed_ref_replace, apply_range_expansion
from core.models import ReferenceIndex, Workbook, WorkbookDiff
from core.named_range_fix import apply_named_range_fix
from core.workbook_diff import diff_workbooks

logger = logging.getLogger(__name__)
router = APIRouter()

# 実ファイルへの書き込み関数: (元ファイル, 出力先) を受け取り修正を適用する
_ApplyFn = Callable[[Path, Path], None]


class NamedRangeFixRequest(BaseModel):
    """POST /jobs/{job_id}/named-range-fix のリクエストボディ."""

    name: str
    new_refers_to: str


class FormulaFixRequest(BaseModel):
    """POST /jobs/{job_id}/formula-fix のリクエストボディ."""

    kind: Literal["fixed_ref_replace", "range_expansion"]
    old_ref: str
    new_ref: str


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
    apply_fn: _ApplyFn,
) -> str:
    """実ファイルへの書き込み + 新ジョブ作成 + 抽出を1関数にまとめる (threadpool 前提).

    Returns:
        新しく作られたジョブの job_id.
    """
    source_path = storage.get_original_path(source_job_id)
    data = source_path.read_bytes()
    # 元と同じバイト列でまず新ジョブを作る (meta.json 生成・拡張子判定を再利用するため)。
    # その後 apply_fn で新ジョブの original.* を書き換える。
    new_meta = storage.create_job(filename, data)
    new_path = storage.get_original_path(new_meta.job_id)
    try:
        apply_fn(source_path, new_path)
        _run_extraction(storage, new_meta.job_id, filename)
    except Exception:
        # 適用/抽出に失敗した場合、作りかけの新ジョブを残すとジョブ一覧に
        # 「uploaded のまま進まない」孤児が溜まるため後始末する。
        storage.delete_job(new_meta.job_id)
        raise
    return new_meta.job_id


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

    def apply_fn(source: Path, out: Path) -> None:
        apply_named_range_fix(source, body.name, body.new_refers_to, out)

    try:
        new_job_id = await asyncio.to_thread(
            _apply_and_extract, storage, job_id, filename, apply_fn
        )
    except NamedRangeFixError as e:
        raise HTTPException(status_code=422, detail=f"named range fix failed: {e}") from e
    except ExtractionError as e:
        raise HTTPException(status_code=422, detail=f"extraction of new file failed: {e}") from e
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"internal error: {e}") from e

    diff = await _self_verify(storage, before_wb, before_path, before_index, new_job_id)

    logger.info(
        "named range fix applied: source=%s new_job=%s name=%s is_empty=%s",
        job_id,
        new_job_id,
        body.name,
        diff.is_empty(),
    )
    return {"new_job_id": new_job_id, "diff": diff.model_dump(by_alias=True)}


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

    def apply_fn(source: Path, out: Path) -> None:
        if body.kind == "range_expansion":
            apply_range_expansion(source, body.old_ref, body.new_ref, out)
        else:
            apply_fixed_ref_replace(source, body.old_ref, body.new_ref, out)

    try:
        new_job_id = await asyncio.to_thread(
            _apply_and_extract, storage, job_id, filename, apply_fn
        )
    except FormulaFixError as e:
        raise HTTPException(status_code=422, detail=f"formula fix failed: {e}") from e
    except ExtractionError as e:
        raise HTTPException(status_code=422, detail=f"extraction of new file failed: {e}") from e
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"internal error: {e}") from e

    diff = await _self_verify(storage, before_wb, before_path, before_index, new_job_id)

    logger.info(
        "formula fix applied: source=%s new_job=%s kind=%s old=%s new=%s is_empty=%s",
        job_id,
        new_job_id,
        body.kind,
        body.old_ref,
        body.new_ref,
        diff.is_empty(),
    )
    return {"new_job_id": new_job_id, "diff": diff.model_dump(by_alias=True)}
