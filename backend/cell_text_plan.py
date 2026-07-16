"""空セルへの固定テキスト追加計画を組み立てる共通処理."""

from __future__ import annotations

from backend.storage import Storage
from core.mutation import (
    CellTextBatchOperation,
    CellTextEdit,
    MutationPlan,
    SafeChangePlan,
    build_safe_change_plan,
)


def build_cell_text_safe_plan(
    storage: Storage,
    job_id: str,
    edits: list[CellTextEdit],
) -> SafeChangePlan:
    """対象セルが空であることを確認し、OfficeCLI向け変更計画を返す.

    Args:
        storage: 対象ジョブの保存領域。
        job_id: 変更元ジョブID。
        edits: 追加する固定テキストの一覧。

    Returns:
        期待差分を含む、未適用の安全変更計画。

    Raises:
        ValueError: 対応外形式、存在しないシート、重複セル、または空ではないセルの場合。
        FileNotFoundError: 抽出結果やセルDBが存在しない場合。
    """

    source_path = storage.get_original_path(job_id)
    if source_path.suffix.lower() != ".xlsx":
        raise ValueError("cell text changes currently support .xlsx only")
    workbook = storage.load_workbook(job_id)
    sheet_names = {sheet.name for sheet in workbook.sheets}
    seen: set[tuple[str, str]] = set()
    for edit in edits:
        if edit.sheet not in sheet_names:
            raise ValueError(f"sheet not found: {edit.sheet}")
        identity = (edit.sheet, edit.coord)
        if identity in seen:
            raise ValueError(f"duplicate target cell: {edit.sheet}!{edit.coord}")
        seen.add(identity)
        cell_range = storage.get_cells_range(job_id, edit.sheet, edit.coord)
        rows = cell_range.get("rows")
        cell: object = None
        if isinstance(rows, list) and rows:
            first_row = rows[0]
            if isinstance(first_row, list) and first_row:
                cell = first_row[0]
        if isinstance(cell, dict) and (
            cell.get("value") is not None or cell.get("formula") is not None
        ):
            raise ValueError(f"target cell is not empty: {edit.sheet}!{edit.coord}")

    references = storage.load_references(job_id)
    plan = MutationPlan(
        source_job_id=job_id,
        requested_provider="officecli",
        operation=CellTextBatchOperation(edits=edits),
    )
    return build_safe_change_plan(plan, workbook, references)
