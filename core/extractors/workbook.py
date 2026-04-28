"""ワークブック構造抽出モジュール.

openpyxl で .xlsx / .xlsm を開き、数式セル・名前付き範囲・条件付き書式・外部リンクを
Pydantic モデルに詰める。VBA は本モジュールでは扱わない (core.extractors.vba を別途呼ぶ).

SPEC.md §4.2 参照。
"""

from __future__ import annotations

import logging
import re
from pathlib import Path

from core.exceptions import ExtractionError
from core.models import (
    CellFormula,
    ConditionalFormat,
    NamedRange,
    SheetInfo,
    Workbook,
)

logger = logging.getLogger(__name__)


# `Sheet1!$A$1` や `'My Sheet'!A1:B10` の先頭シート名部分を取り出す
_SHEET_PREFIX_RE = re.compile(r"^(?:'([^']+)'|([^!'\s]+))!")


def _extract_sheet_name(refers_to: str) -> str | None:
    """`Sheet1!$A$1` 形式の参照から先頭のシート名を抜き出す.

    ヒットしない (シート名なし) 場合は None.
    """
    m = _SHEET_PREFIX_RE.match(refers_to.strip())
    if not m:
        return None
    return m.group(1) or m.group(2)


def _extract_formula_refs(formula: str) -> list[str]:
    """数式から RANGE トークンを抽出する.

    openpyxl の Tokenizer に頼る。`=SUMIF(Input!A:A, A2, Input!E:E)` から
    `["Input!A:A", "A2", "Input!E:E"]` を得る想定。
    """
    from openpyxl.formula.tokenizer import Tokenizer

    if not formula:
        return []
    if not formula.startswith("="):
        formula = "=" + formula

    try:
        tok = Tokenizer(formula)
    except Exception:  # noqa: BLE001 - 構文エラーは握りつぶしてログ
        logger.debug("Failed to tokenize formula: %r", formula)
        return []

    refs: list[str] = []
    for t in tok.items:
        if getattr(t, "subtype", None) == "RANGE":
            refs.append(t.value)
    return refs


def _extract_formulas(ws: object) -> list[CellFormula]:
    """1シートから数式セルを抽出する."""
    formulas: list[CellFormula] = []
    sheet_title: str = ws.title  # type: ignore[attr-defined]

    for row in ws.iter_rows():  # type: ignore[attr-defined]
        for cell in row:
            if cell.data_type != "f":
                continue
            value = cell.value
            if value is None:
                continue
            formula_str = str(value)
            coord = f"{sheet_title}!{cell.coordinate}"
            refs = _extract_formula_refs(formula_str)
            formulas.append(CellFormula(coord=coord, formula=formula_str, refs=refs))
    return formulas


def _extract_conditional_formats(ws: object) -> list[ConditionalFormat]:
    """1シートから条件付き書式を抽出する."""
    cfs: list[ConditionalFormat] = []
    cf_list = getattr(ws, "conditional_formatting", None)
    if cf_list is None:
        return cfs
    try:
        for rng, rules in cf_list._cf_rules.items():
            range_str = str(rng.sqref) if hasattr(rng, "sqref") else str(rng)
            for rule in rules:
                rule_str = _summarize_cf_rule(rule)
                cfs.append(ConditionalFormat(range=range_str, rule=rule_str))
    except Exception:  # noqa: BLE001
        logger.debug(
            "Failed to extract conditional formats from %s",
            sheet_title := getattr(ws, "title", "?"),
        )
        _ = sheet_title
    return cfs


def _summarize_cf_rule(rule: object) -> str:
    """ConditionalFormatRule を人間が読める短い文字列に."""
    rtype = getattr(rule, "type", None) or "rule"
    operator = getattr(rule, "operator", None)
    formula = getattr(rule, "formula", None)
    parts = [str(rtype)]
    if operator:
        parts.append(str(operator))
    if formula:
        formulas_list = list(formula) if not isinstance(formula, str) else [formula]
        parts.append(",".join(str(f) for f in formulas_list))
    return " ".join(parts)


def _extract_external_links(wb: object) -> list[str]:
    """外部リンクを抽出する.

    `wb._external_links` は private API なので失敗を許容する。
    """
    links: list[str] = []
    raw = getattr(wb, "_external_links", None)
    if not raw:
        return links
    try:
        for link in raw:
            file_link = getattr(link, "file_link", None)
            target = getattr(file_link, "Target", None) if file_link else None
            if target:
                links.append(str(target))
    except Exception:  # noqa: BLE001
        logger.debug("Failed to enumerate external links")
    return links


def _attach_named_ranges(wb: object, sheets_by_name: dict[str, SheetInfo]) -> None:
    """ワークブック定義名を、参照先シートの SheetInfo.named_ranges に振り分ける.

    シートが特定できない (refers_to が解析不能 or 該当シート無し) 場合は
    最初のシートに付ける。
    """
    defined_names = getattr(wb, "defined_names", None)
    if not defined_names:
        return

    try:
        items = list(defined_names.items())
    except Exception:  # noqa: BLE001
        logger.debug("Failed to enumerate defined_names")
        return

    for name, defn in items:
        refers_to = getattr(defn, "value", None) or ""
        sheet_name = _extract_sheet_name(refers_to)
        target_sheet: SheetInfo | None = None
        if sheet_name and sheet_name in sheets_by_name:
            target_sheet = sheets_by_name[sheet_name]
        elif sheets_by_name:
            target_sheet = next(iter(sheets_by_name.values()))
        if target_sheet is not None:
            target_sheet.named_ranges.append(NamedRange(name=name, refers_to=refers_to))


def extract_workbook(file_path: Path) -> Workbook:
    """Excelファイルからシート構造を抽出する.

    Args:
        file_path: 対象ファイルのパス. .xlsx / .xlsm を想定.

    Returns:
        Workbook モデル. VBA モジュールは含めない (vba_modules は空).
        `.xls` (旧バイナリ形式) が渡された場合は sheets=[] で返し警告ログを出す.

    Raises:
        ExtractionError: ファイルが存在しない、または openpyxl が開けない場合.
    """
    if not file_path.exists():
        raise ExtractionError(f"File not found: {file_path}")

    if file_path.suffix.lower() == ".xls":
        logger.warning(
            ".xls (legacy binary) is not supported by openpyxl; "
            "returning empty workbook structure for %s",
            file_path.name,
        )
        return Workbook(filename=file_path.name)

    from openpyxl import load_workbook

    try:
        wb = load_workbook(filename=str(file_path), keep_vba=True, data_only=False)
    except Exception as e:  # noqa: BLE001
        raise ExtractionError(f"openpyxl failed to open {file_path}: {e}") from e

    sheets: list[SheetInfo] = []
    sheets_by_name: dict[str, SheetInfo] = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        info = SheetInfo(
            name=sn,
            rows=ws.max_row or 0,
            cols=ws.max_column or 0,
            formulas=_extract_formulas(ws),
            conditional_formats=_extract_conditional_formats(ws),
        )
        sheets.append(info)
        sheets_by_name[sn] = info

    _attach_named_ranges(wb, sheets_by_name)

    return Workbook(
        filename=file_path.name,
        sheets=sheets,
        external_links=_extract_external_links(wb),
    )
