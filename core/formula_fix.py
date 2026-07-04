"""数式内の参照置換 (S2 増分1: 固定参照置換・数式範囲拡張).

安全パターンのうち「数式が参照している範囲/セルを別の範囲/セルに書き換える」だけを
対象にする。数式の構造 (関数・演算子・引数) は一切変更しない。

docs/VISION.ja.md §4.3 のとおり素朴な文字列置換は使わない。openpyxl の formula
Tokenizer で数式をトークン列に分解し、参照トークン (OPERAND/RANGE) のうち対象と
完全一致するものだけを置換する。これにより文字列リテラル内の "A1" や、部分一致する
別の参照 (例: A1 を置換したいときの A10) を誤って書き換えることがない。

フロー (named_range_fix と同型):
  1. `propose_fixed_ref_replace()` / `propose_range_expansion()` で「適用したら
     どの数式がどう変わるか」を書き込みなしで計算する (tool loop から呼んでよい)。
  2. ユーザーが影響を見て納得したら `apply_fixed_ref_replace()` /
     `apply_range_expansion()` で実ファイルに書き込む (人間の明示操作からのみ呼ぶ。
     tool loop からは呼ばない — docs/VISION.ja.md §4.2)。
"""

from __future__ import annotations

import logging
import re
from pathlib import Path

from openpyxl import load_workbook as load_openpyxl_workbook
from openpyxl.formula.tokenizer import Token, Tokenizer, TokenizerError
from openpyxl.utils.cell import range_boundaries

from core.exceptions import FormulaFixError
from core.models import CellDiff, ReferenceIndex, Workbook, WorkbookDiff
from core.workbook_diff import build_blast_radius

logger = logging.getLogger(__name__)

# 引用符なしで数式に書けるシート名 (それ以外は '...' で囲む)
_SHEET_NO_QUOTE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")


def _split_sheet_ref(ref: str) -> tuple[str | None, str]:
    """'Data!$A$1' → ("Data", "$A$1") に分解する. シート修飾がなければ (None, ref).

    シート名の引用符 ('Other Sheet'!A1) は外し、'' エスケープは ' に戻す。
    """
    ref = ref.strip()
    if "!" not in ref:
        return None, ref
    sheet_part, cell_part = ref.rsplit("!", 1)
    if sheet_part.startswith("'") and sheet_part.endswith("'") and len(sheet_part) >= 2:
        sheet_part = sheet_part[1:-1].replace("''", "'")
    return sheet_part, cell_part


def _normalize(sheet: str, cell_part: str) -> tuple[str, str]:
    """比較用の正規化キーを作る (シート名は大小無視、$ は無視、座標は大文字)."""
    return sheet.casefold(), cell_part.replace("$", "").upper()


def _quote_sheet(sheet: str) -> str:
    """数式に書けるシート名表記にする (必要なら '...' で囲む)."""
    if _SHEET_NO_QUOTE.match(sheet):
        return sheet
    return "'" + sheet.replace("'", "''") + "'"


def _parse_qualified_ref(ref: str, label: str) -> tuple[str, str]:
    """シート修飾付き参照をパースし (シート名, セル部) を返す.

    Args:
        ref: "Data!$B$5" / "'Other Sheet'!A1:A100" 形式の参照.
        label: エラーメッセージ用の引数名 (old_ref / new_ref 等).

    Raises:
        FormulaFixError: シート修飾がない、またはセル部が参照として解釈できない場合.
    """
    sheet, cell_part = _split_sheet_ref(ref)
    if sheet is None or not sheet:
        raise FormulaFixError(f"{label} must be sheet-qualified (e.g. 'Data!$B$5'): {ref}")
    try:
        range_boundaries(cell_part.replace("$", ""))
    except ValueError as exc:
        raise FormulaFixError(f"{label} is not a valid cell/range reference: {ref}") from exc
    return sheet, cell_part


def _replace_refs_in_formula(
    formula: str,
    formula_sheet: str,
    old_key: tuple[str, str],
    new_sheet: str,
    new_cell_part: str,
) -> str | None:
    """数式内の参照トークンのうち old_key と一致するものを置換する.

    Args:
        formula: "=" 始まりの数式文字列.
        formula_sheet: この数式が置かれているシート名 (シート修飾なしトークンの解決に使う).
        old_key: `_normalize()` 済みの置換対象キー.
        new_sheet: 置換後のシート名 (未加工).
        new_cell_part: 置換後のセル部 ($ 表記は呼び出し元の指定を保持).

    Returns:
        置換が1件以上発生した場合は新しい数式文字列、変更なしなら None。
        数式がトークン化できない場合も None (置換対象にしない)。
    """
    try:
        tokenizer = Tokenizer(formula)
    except TokenizerError:
        logger.warning("formula could not be tokenized; skipped: %s", formula)
        return None

    changed = False
    for token in tokenizer.items:
        if token.type != Token.OPERAND or token.subtype != Token.RANGE:
            continue
        token_sheet, token_cell = _split_sheet_ref(token.value)
        resolved_sheet = token_sheet if token_sheet is not None else formula_sheet
        if _normalize(resolved_sheet, token_cell) != old_key:
            continue
        if token_sheet is None and new_sheet.casefold() == formula_sheet.casefold():
            # 元がシート修飾なしで置換後も同じシートなら、修飾なし表記を保つ
            token.value = new_cell_part
        else:
            token.value = f"{_quote_sheet(new_sheet)}!{new_cell_part}"
        changed = True

    if not changed:
        return None
    rendered: str = tokenizer.render()
    return rendered


def _validate_expansion(old_ref: str, new_ref: str) -> None:
    """数式範囲拡張の前提を検証する: 同一シートで、新範囲が旧範囲を包含していること.

    Raises:
        FormulaFixError: シートが異なる、範囲が非有界 (A:A 等)、包含していない、
            または旧範囲と同一の場合.
    """
    old_sheet, old_cell = _parse_qualified_ref(old_ref, "old_range")
    new_sheet, new_cell = _parse_qualified_ref(new_ref, "new_range")
    if old_sheet.casefold() != new_sheet.casefold():
        raise FormulaFixError(
            f"range expansion must stay on the same sheet: {old_ref} -> {new_ref}"
        )
    old_b = range_boundaries(old_cell.replace("$", ""))
    new_b = range_boundaries(new_cell.replace("$", ""))
    if any(v is None for v in old_b) or any(v is None for v in new_b):
        raise FormulaFixError(
            f"range expansion requires bounded ranges (not whole columns/rows): "
            f"{old_ref} -> {new_ref}"
        )
    if old_b == new_b:
        raise FormulaFixError(f"new_range is identical to old_range: {new_ref}")
    o_min_c, o_min_r, o_max_c, o_max_r = old_b
    n_min_c, n_min_r, n_max_c, n_max_r = new_b
    contains = (
        n_min_c <= o_min_c and n_min_r <= o_min_r and n_max_c >= o_max_c and n_max_r >= o_max_r
    )
    if not contains:
        raise FormulaFixError(f"new_range must contain old_range entirely: {old_ref} -> {new_ref}")


def _propose(
    before_wb: Workbook,
    before_index: ReferenceIndex,
    old_ref: str,
    new_ref: str,
) -> WorkbookDiff:
    """old_ref を new_ref に置換したらどの数式が変わるかをメモリ内で試算する."""
    old_sheet, old_cell = _parse_qualified_ref(old_ref, "old_ref")
    new_sheet, new_cell = _parse_qualified_ref(new_ref, "new_ref")
    old_key = _normalize(old_sheet, old_cell)

    changed_cells: list[CellDiff] = []
    for sheet in before_wb.sheets:
        for f in sheet.formulas:
            new_formula = _replace_refs_in_formula(
                f.formula, sheet.name, old_key, new_sheet, new_cell
            )
            if new_formula is None:
                continue
            # CellFormula.coord は "Sheet!A1" 形式、CellDiff.coord はシート修飾なし
            coord = f.coord.rsplit("!", 1)[-1]
            changed_cells.append(
                CellDiff(
                    sheet=sheet.name,
                    coord=coord,
                    change_type="modified",
                    before_formula=f.formula,
                    after_formula=new_formula,
                )
            )
    if not changed_cells:
        raise FormulaFixError(f"no formulas reference {old_ref}")

    blast_radius = build_blast_radius(changed_cells, [], before_index)
    return WorkbookDiff(
        before_filename=before_wb.filename,
        after_filename=before_wb.filename,
        cells=changed_cells,
        blast_radius=blast_radius,
        existing_risks=list(before_wb.analysis_risks),
    )


def propose_fixed_ref_replace(
    before_wb: Workbook,
    before_index: ReferenceIndex,
    old_ref: str,
    new_ref: str,
) -> WorkbookDiff:
    """固定参照置換の試算: old_ref を参照している全数式の置換後を計算する (read-only).

    実ファイルには一切書き込まない。

    Args:
        before_wb: 対象ジョブの抽出済み Workbook.
        before_index: 対象ジョブの ReferenceIndex (波及範囲算出に使う).
        old_ref: 置換対象の参照 (シート修飾必須。例: "Data!$B$5").
        new_ref: 置換後の参照 (シート修飾必須。例: "Data!$B$6").

    Returns:
        WorkbookDiff。cells に変更される数式セルの before/after が入る。

    Raises:
        FormulaFixError: 参照が不正、または old_ref を参照する数式が1つもない場合.
    """
    return _propose(before_wb, before_index, old_ref, new_ref)


def propose_range_expansion(
    before_wb: Workbook,
    before_index: ReferenceIndex,
    old_range: str,
    new_range: str,
) -> WorkbookDiff:
    """数式範囲拡張の試算: old_range を参照している全数式の拡張後を計算する (read-only).

    固定参照置換と同じ置換エンジンを使うが、new_range が old_range を完全に包含する
    「拡張」であることを事前に検証する (縮小や別領域への移動はこのパターンでは扱わない)。

    Args:
        before_wb: 対象ジョブの抽出済み Workbook.
        before_index: 対象ジョブの ReferenceIndex.
        old_range: 拡張対象の範囲 (シート修飾必須。例: "Data!$A$1:$A$100").
        new_range: 拡張後の範囲 (同一シートで old_range を包含すること).

    Returns:
        WorkbookDiff。cells に変更される数式セルの before/after が入る。

    Raises:
        FormulaFixError: 範囲が不正、拡張になっていない、または該当数式がない場合.
    """
    _validate_expansion(old_range, new_range)
    return _propose(before_wb, before_index, old_range, new_range)


def _apply(
    file_path: Path,
    old_ref: str,
    new_ref: str,
    out_path: Path,
) -> int:
    """old_ref → new_ref の置換を実ファイルに適用し、out_path に書き出す.

    Returns:
        置換が発生した数式セルの数.
    """
    old_sheet, old_cell = _parse_qualified_ref(old_ref, "old_ref")
    new_sheet, new_cell = _parse_qualified_ref(new_ref, "new_ref")
    old_key = _normalize(old_sheet, old_cell)

    try:
        wb = load_openpyxl_workbook(file_path, keep_vba=True)
    except Exception as exc:  # noqa: BLE001 - 失敗内容こそ診断材料
        raise FormulaFixError(f"failed to open workbook: {file_path}: {exc}") from exc

    replaced = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                # ArrayFormula 等の非文字列数式は対象外 (propose 側も同様にスキップされる)
                if not isinstance(value, str) or not value.startswith("="):
                    continue
                new_formula = _replace_refs_in_formula(
                    value, ws.title, old_key, new_sheet, new_cell
                )
                if new_formula is None:
                    continue
                cell.value = new_formula
                replaced += 1

    if replaced == 0:
        raise FormulaFixError(f"no formulas reference {old_ref} in {file_path}")

    try:
        wb.save(out_path)
    except Exception as exc:  # noqa: BLE001
        raise FormulaFixError(f"failed to save workbook: {out_path}: {exc}") from exc
    return replaced


def apply_fixed_ref_replace(
    file_path: Path,
    old_ref: str,
    new_ref: str,
    out_path: Path,
) -> int:
    """固定参照置換を適用した新しい xlsx/xlsm を out_path に書き出す.

    file_path 自体は変更しない (別ファイルへの出力)。

    Args:
        file_path: 元ファイル (ジョブの original.*).
        old_ref: 置換対象の参照 (シート修飾必須).
        new_ref: 置換後の参照 (シート修飾必須).
        out_path: 書き出し先パス.

    Returns:
        置換が発生した数式セルの数.

    Raises:
        FormulaFixError: ファイルが開けない、該当数式がない、または保存に失敗した場合.
    """
    return _apply(file_path, old_ref, new_ref, out_path)


def apply_range_expansion(
    file_path: Path,
    old_range: str,
    new_range: str,
    out_path: Path,
) -> int:
    """数式範囲拡張を適用した新しい xlsx/xlsm を out_path に書き出す.

    Args:
        file_path: 元ファイル (ジョブの original.*).
        old_range: 拡張対象の範囲 (シート修飾必須).
        new_range: 拡張後の範囲 (同一シートで old_range を包含すること).
        out_path: 書き出し先パス.

    Returns:
        置換が発生した数式セルの数.

    Raises:
        FormulaFixError: 範囲が不正、拡張になっていない、該当数式がない、
            または開けない/保存できない場合.
    """
    _validate_expansion(old_range, new_range)
    return _apply(file_path, old_range, new_range, out_path)
