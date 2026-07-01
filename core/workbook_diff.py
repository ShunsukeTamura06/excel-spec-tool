"""ワークブック構造差分エンジン (P1 安全ゲートの静的パート).

2バージョンのワークブック (before/after) を比較し、セル・名前付き範囲・条件付き書式・
入力規則・グラフ・ピボットテーブル・VBAモジュールの構造差分と、変更箇所の波及範囲
(blast radius) を計算する。Excel COM による動的検証 (再計算・マクロ実行) はここでは
扱わない (別増分)。

docs/VISION.ja.md §6.2/§6.3 のスパイクで検証済みの手法をそのまま使う: 生XML(zip) の
差分はノイズが多いため、正規化抽出 (`extract_cells_to_sqlite`) を比較基盤にすることで
ノイズを無視しつつ本物の変更だけを検出する。

docs/SPEC.ja.md §4.6 参照。
"""

from __future__ import annotations

import sqlite3
import tempfile
from pathlib import Path

from core.exceptions import DiffError, ExtractionError
from core.extractors.cells import extract_cells_to_sqlite
from core.models import (
    BlastRadiusEntry,
    CellDiff,
    ChangeType,
    ChartDiff,
    ConditionalFormatDiff,
    DataValidationDiff,
    NamedRangeDiff,
    PivotTableDiff,
    ReferenceIndex,
    VbaModuleDiff,
    Workbook,
    WorkbookDiff,
)
from core.reference_index import find_overlapping

# (sheet, coord) -> (value, formula, number_format)
_CellKey = tuple[str, str]
_CellValue = tuple[str | None, str | None, str | None]


def diff_workbooks(
    before_path: Path,
    after_path: Path,
    before_wb: Workbook,
    after_wb: Workbook,
    before_index: ReferenceIndex,
) -> WorkbookDiff:
    """2バージョンのワークブックを比較し、構造差分+波及範囲レポートを作る.

    Args:
        before_path: 「前」バージョンの原本ファイルパス (セル差分の算出に使う).
        after_path: 「後」バージョンの原本ファイルパス.
        before_wb: 「前」の抽出済み Workbook.
        after_wb: 「後」の抽出済み Workbook.
        before_index: 「前」の Workbook から build_reference_index() 済みの ReferenceIndex.
            波及範囲は「変更前の時点で何がその箇所を参照していたか」を見るため before 固定.

    Returns:
        WorkbookDiff。追加/削除/変更の一覧と、変更箇所ごとの波及範囲、
        before_wb.analysis_risks の再掲。

    Raises:
        DiffError: before/after いずれかのセル抽出に失敗した場合.
    """
    cells = _diff_cells(before_path, after_path)
    named_ranges = _diff_named_ranges(before_wb, after_wb)
    conditional_formats = _diff_conditional_formats(before_wb, after_wb)
    data_validations = _diff_data_validations(before_wb, after_wb)
    charts = _diff_charts(before_wb, after_wb)
    pivot_tables = _diff_pivot_tables(before_wb, after_wb)
    vba_modules = _diff_vba_modules(before_wb, after_wb)
    blast_radius = _build_blast_radius(cells, named_ranges, before_index)

    return WorkbookDiff(
        before_filename=before_wb.filename,
        after_filename=after_wb.filename,
        cells=cells,
        named_ranges=named_ranges,
        conditional_formats=conditional_formats,
        data_validations=data_validations,
        charts=charts,
        pivot_tables=pivot_tables,
        vba_modules=vba_modules,
        blast_radius=blast_radius,
        existing_risks=list(before_wb.analysis_risks),
    )


def _dump_cells_by_key(file_path: Path, db_path: Path) -> dict[_CellKey, _CellValue]:
    """セル正規化抽出 (extract_cells_to_sqlite) を実行し、座標をキーにした辞書を返す."""
    try:
        extract_cells_to_sqlite(file_path, db_path)
    except ExtractionError as exc:
        raise DiffError(f"セル抽出に失敗しました ({file_path}): {exc}") from exc

    conn = sqlite3.connect(str(db_path))
    try:
        rows = conn.execute(
            "SELECT sheet, coord, value, formula, number_format FROM cells"
        ).fetchall()
    finally:
        conn.close()
    return {
        (sheet, coord): (value, formula, number_format)
        for sheet, coord, value, formula, number_format in rows
    }


def _diff_cells(before_path: Path, after_path: Path) -> list[CellDiff]:
    """2ファイルをセル単位で比較する (spikes/xlsx_diff_noise で検証済みの手法)."""
    if not before_path.exists():
        raise DiffError(f"before ファイルが見つかりません: {before_path}")
    if not after_path.exists():
        raise DiffError(f"after ファイルが見つかりません: {after_path}")

    try:
        with tempfile.TemporaryDirectory(prefix="workbook_diff_") as tmp:
            tmp_dir = Path(tmp)
            before_cells = _dump_cells_by_key(before_path, tmp_dir / "before.db")
            after_cells = _dump_cells_by_key(after_path, tmp_dir / "after.db")
    except OSError as exc:
        raise DiffError(f"一時ディレクトリの作成に失敗しました: {exc}") from exc

    diffs: list[CellDiff] = []
    for key in before_cells.keys() | after_cells.keys():
        before_v = before_cells.get(key)
        after_v = after_cells.get(key)
        if before_v == after_v:
            continue

        change_type: ChangeType
        if before_v is None:
            change_type = "added"
        elif after_v is None:
            change_type = "removed"
        else:
            change_type = "modified"

        before_value, before_formula, before_number_format = before_v or (None, None, None)
        after_value, after_formula, after_number_format = after_v or (None, None, None)
        sheet, coord = key
        diffs.append(
            CellDiff(
                sheet=sheet,
                coord=coord,
                change_type=change_type,
                before_value=before_value,
                after_value=after_value,
                before_formula=before_formula,
                after_formula=after_formula,
                before_number_format=before_number_format,
                after_number_format=after_number_format,
            )
        )
    return sorted(diffs, key=lambda d: (d.sheet, d.coord))


def _diff_named_ranges(before_wb: Workbook, after_wb: Workbook) -> list[NamedRangeDiff]:
    """名前付き範囲を name をキーに突き合わせる (ワークブック全体で一意な前提)."""
    before_map = {nr.name: nr for sheet in before_wb.sheets for nr in sheet.named_ranges}
    after_map = {nr.name: nr for sheet in after_wb.sheets for nr in sheet.named_ranges}

    diffs: list[NamedRangeDiff] = []
    for name in sorted(before_map.keys() | after_map.keys()):
        before_nr = before_map.get(name)
        after_nr = after_map.get(name)
        if before_nr == after_nr:
            continue
        change_type: ChangeType = (
            "added" if before_nr is None else "removed" if after_nr is None else "modified"
        )
        diffs.append(
            NamedRangeDiff(
                name=name,
                change_type=change_type,
                before_refers_to=before_nr.refers_to if before_nr else None,
                after_refers_to=after_nr.refers_to if after_nr else None,
            )
        )
    return diffs


def _diff_conditional_formats(
    before_wb: Workbook, after_wb: Workbook
) -> list[ConditionalFormatDiff]:
    """条件付き書式を (sheet, range) をキーに突き合わせる."""
    before_map = {
        (sheet.name, cf.range): cf for sheet in before_wb.sheets for cf in sheet.conditional_formats
    }
    after_map = {
        (sheet.name, cf.range): cf for sheet in after_wb.sheets for cf in sheet.conditional_formats
    }

    diffs: list[ConditionalFormatDiff] = []
    for key in sorted(before_map.keys() | after_map.keys()):
        before_cf = before_map.get(key)
        after_cf = after_map.get(key)
        if before_cf == after_cf:
            continue
        sheet, range_ = key
        change_type: ChangeType = (
            "added" if before_cf is None else "removed" if after_cf is None else "modified"
        )
        diffs.append(
            ConditionalFormatDiff(
                sheet=sheet,
                range=range_,
                change_type=change_type,
                before_rule=before_cf.rule if before_cf else None,
                after_rule=after_cf.rule if after_cf else None,
            )
        )
    return diffs


def _diff_data_validations(before_wb: Workbook, after_wb: Workbook) -> list[DataValidationDiff]:
    """入力規則を (sheet, range) をキーに突き合わせる."""
    before_map = {
        (sheet.name, dv.range): dv for sheet in before_wb.sheets for dv in sheet.data_validations
    }
    after_map = {
        (sheet.name, dv.range): dv for sheet in after_wb.sheets for dv in sheet.data_validations
    }

    diffs: list[DataValidationDiff] = []
    for key in sorted(before_map.keys() | after_map.keys()):
        before_dv = before_map.get(key)
        after_dv = after_map.get(key)
        if before_dv == after_dv:
            continue
        sheet, range_ = key
        change_type: ChangeType = (
            "added" if before_dv is None else "removed" if after_dv is None else "modified"
        )
        diffs.append(
            DataValidationDiff(
                sheet=sheet,
                range=range_,
                change_type=change_type,
                before=before_dv,
                after=after_dv,
            )
        )
    return diffs


def _chart_key(sheet_name: str, chart: object) -> tuple[str, str]:
    """グラフのキーを作る. name が空なら anchor をフォールバックにする."""
    name = getattr(chart, "name", "") or getattr(chart, "anchor", "")
    return (sheet_name, name)


def _diff_charts(before_wb: Workbook, after_wb: Workbook) -> list[ChartDiff]:
    """グラフを (sheet, name または anchor) をキーに突き合わせる."""
    before_map = {
        _chart_key(sheet.name, chart): chart for sheet in before_wb.sheets for chart in sheet.charts
    }
    after_map = {
        _chart_key(sheet.name, chart): chart for sheet in after_wb.sheets for chart in sheet.charts
    }

    diffs: list[ChartDiff] = []
    for key in sorted(before_map.keys() | after_map.keys()):
        before_chart = before_map.get(key)
        after_chart = after_map.get(key)
        if before_chart == after_chart:
            continue
        sheet, chart_key = key
        change_type: ChangeType = (
            "added" if before_chart is None else "removed" if after_chart is None else "modified"
        )
        diffs.append(
            ChartDiff(
                sheet=sheet,
                key=chart_key,
                change_type=change_type,
                before=before_chart,
                after=after_chart,
            )
        )
    return diffs


def _diff_pivot_tables(before_wb: Workbook, after_wb: Workbook) -> list[PivotTableDiff]:
    """ピボットテーブルを (sheet, name) をキーに突き合わせる."""
    before_map = {
        (sheet.name, pt.name): pt for sheet in before_wb.sheets for pt in sheet.pivot_tables
    }
    after_map = {
        (sheet.name, pt.name): pt for sheet in after_wb.sheets for pt in sheet.pivot_tables
    }

    diffs: list[PivotTableDiff] = []
    for key in sorted(before_map.keys() | after_map.keys()):
        before_pt = before_map.get(key)
        after_pt = after_map.get(key)
        if before_pt == after_pt:
            continue
        sheet, name = key
        change_type: ChangeType = (
            "added" if before_pt is None else "removed" if after_pt is None else "modified"
        )
        diffs.append(
            PivotTableDiff(
                sheet=sheet,
                name=name,
                change_type=change_type,
                before=before_pt,
                after=after_pt,
            )
        )
    return diffs


def _diff_vba_modules(before_wb: Workbook, after_wb: Workbook) -> list[VbaModuleDiff]:
    """VBAモジュールを name をキーに突き合わせる. コード全文一致で modified 判定."""
    before_map = {m.name: m for m in before_wb.vba_modules}
    after_map = {m.name: m for m in after_wb.vba_modules}

    diffs: list[VbaModuleDiff] = []
    for name in sorted(before_map.keys() | after_map.keys()):
        before_m = before_map.get(name)
        after_m = after_map.get(name)
        if before_m and after_m and before_m.code == after_m.code and before_m.type == after_m.type:
            continue
        if before_m is None and after_m is None:
            continue
        change_type: ChangeType = (
            "added" if before_m is None else "removed" if after_m is None else "modified"
        )
        diffs.append(
            VbaModuleDiff(
                name=name,
                change_type=change_type,
                before_code=before_m.code if before_m else None,
                after_code=after_m.code if after_m else None,
            )
        )
    return diffs


def _build_blast_radius(
    cells: list[CellDiff],
    named_ranges: list[NamedRangeDiff],
    before_index: ReferenceIndex,
) -> list[BlastRadiusEntry]:
    """削除/変更された箇所を、before 時点で参照していた箇所を洗い出す.

    added (before の時点で存在しなかった箇所) は対象外にする — 存在しないものを
    参照できるはずがないため。
    """
    entries: list[BlastRadiusEntry] = []

    for cell in cells:
        if cell.change_type == "added":
            continue
        location = f"{cell.sheet}!{cell.coord}"
        referenced_by = find_overlapping(before_index, location)
        if referenced_by:
            entries.append(
                BlastRadiusEntry(
                    location=location,
                    change_type=cell.change_type,
                    referenced_by=referenced_by,
                )
            )

    for nr in named_ranges:
        if nr.change_type == "added" or nr.before_refers_to is None:
            continue
        referenced_by = find_overlapping(before_index, nr.before_refers_to)
        if referenced_by:
            entries.append(
                BlastRadiusEntry(
                    location=nr.before_refers_to,
                    change_type=nr.change_type,
                    referenced_by=referenced_by,
                )
            )

    return entries
