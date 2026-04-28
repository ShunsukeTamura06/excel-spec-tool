"""core.extractors.workbook のテスト.

xlsx は openpyxl で動的に作って tmp_path に書き出し、それを読み戻して検証する。
バイナリの xlsx fixture をコミットせずに済ませる。
"""

from pathlib import Path

import pytest
from openpyxl import Workbook as OpyWorkbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName

from core.exceptions import ExtractionError
from core.extractors.workbook import (
    _extract_formula_refs,
    _extract_sheet_name,
    extract_workbook,
)

# ---------- 内部ヘルパーの単体テスト ----------


class TestExtractSheetName:
    def test_simple(self) -> None:
        assert _extract_sheet_name("Sheet1!$A$1") == "Sheet1"

    def test_quoted_with_space(self) -> None:
        assert _extract_sheet_name("'My Sheet'!A1:B2") == "My Sheet"

    def test_no_sheet(self) -> None:
        assert _extract_sheet_name("$A$1") is None

    def test_empty(self) -> None:
        assert _extract_sheet_name("") is None


class TestExtractFormulaRefs:
    def test_simple_sum(self) -> None:
        refs = _extract_formula_refs("=SUM(A1:A10)")
        assert "A1:A10" in refs

    def test_cross_sheet(self) -> None:
        refs = _extract_formula_refs("=Sheet2!A1")
        assert "Sheet2!A1" in refs

    def test_sumif_three_refs(self) -> None:
        refs = _extract_formula_refs("=SUMIF(Input!A:A, A2, Input!E:E)")
        assert "Input!A:A" in refs
        assert "A2" in refs
        assert "Input!E:E" in refs

    def test_no_leading_equals(self) -> None:
        # 先頭の "=" が無くても正しく動くこと (補完される)
        refs = _extract_formula_refs("SUM(B1:B5)")
        assert "B1:B5" in refs

    def test_invalid_formula_returns_empty(self) -> None:
        # Tokenizer が壊れる入力でも例外を漏らさないこと
        # 完全に空でもクラッシュしない
        assert _extract_formula_refs("") == []


# ---------- extract_workbook の統合テスト ----------


@pytest.fixture
def simple_xlsx(tmp_path: Path) -> Path:
    """A1=10, A2=20, A3=数式, B1=数式 の xlsx を生成."""
    wb = OpyWorkbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Calc"
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = "=SUM(A1:A2)"
    ws["B1"] = "=A3*2"
    out = tmp_path / "simple.xlsx"
    wb.save(out)
    return out


@pytest.fixture
def cross_sheet_xlsx(tmp_path: Path) -> Path:
    """Sheet1.A1=10, Sheet2.B1=`=Sheet1!A1` の xlsx を生成."""
    wb = OpyWorkbook()
    ws1 = wb.active
    assert ws1 is not None
    ws1.title = "Input"
    ws1["A1"] = 10

    ws2 = wb.create_sheet("Calc")
    ws2["B1"] = "=Input!A1"
    ws2["B2"] = "=SUMIF(Input!A:A, B1, Input!A:A)"

    out = tmp_path / "cross.xlsx"
    wb.save(out)
    return out


@pytest.fixture
def named_range_xlsx(tmp_path: Path) -> Path:
    """名前付き範囲を持つ xlsx."""
    wb = OpyWorkbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Data"
    ws["A1"] = 1
    wb.defined_names["TaxRate"] = DefinedName(name="TaxRate", attr_text="Data!$A$1")

    out = tmp_path / "named.xlsx"
    wb.save(out)
    return out


@pytest.fixture
def conditional_format_xlsx(tmp_path: Path) -> Path:
    """条件付き書式を含む xlsx."""
    wb = OpyWorkbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"
    fill = PatternFill(start_color="FFEE1111", end_color="FFEE1111", fill_type="solid")
    rule = CellIsRule(operator="greaterThan", formula=["100"], fill=fill)
    ws.conditional_formatting.add("A1:A10", rule)

    out = tmp_path / "cf.xlsx"
    wb.save(out)
    return out


class TestExtractFormulas:
    def test_simple(self, simple_xlsx: Path) -> None:
        wb = extract_workbook(simple_xlsx)
        assert wb.filename == "simple.xlsx"
        assert len(wb.sheets) == 1
        sheet = wb.sheets[0]
        assert sheet.name == "Calc"

        coords = {f.coord for f in sheet.formulas}
        assert coords == {"Calc!A3", "Calc!B1"}

        a3 = next(f for f in sheet.formulas if f.coord == "Calc!A3")
        assert a3.formula == "=SUM(A1:A2)"
        assert "A1:A2" in a3.refs

        b1 = next(f for f in sheet.formulas if f.coord == "Calc!B1")
        assert b1.formula == "=A3*2"
        assert "A3" in b1.refs

    def test_cross_sheet_refs(self, cross_sheet_xlsx: Path) -> None:
        wb = extract_workbook(cross_sheet_xlsx)
        assert {s.name for s in wb.sheets} == {"Input", "Calc"}

        calc = next(s for s in wb.sheets if s.name == "Calc")

        b1 = next(f for f in calc.formulas if f.coord == "Calc!B1")
        assert "Input!A1" in b1.refs

        b2 = next(f for f in calc.formulas if f.coord == "Calc!B2")
        # SUMIF の3つの引数すべてが拾えていること
        assert "Input!A:A" in b2.refs
        assert "B1" in b2.refs


class TestNamedRanges:
    def test_named_range_attached_to_correct_sheet(self, named_range_xlsx: Path) -> None:
        wb = extract_workbook(named_range_xlsx)
        data_sheet = next(s for s in wb.sheets if s.name == "Data")
        assert len(data_sheet.named_ranges) == 1
        nr = data_sheet.named_ranges[0]
        assert nr.name == "TaxRate"
        assert "Data" in nr.refers_to


class TestConditionalFormats:
    def test_extracted(self, conditional_format_xlsx: Path) -> None:
        wb = extract_workbook(conditional_format_xlsx)
        sheet = wb.sheets[0]
        assert len(sheet.conditional_formats) >= 1
        cf = sheet.conditional_formats[0]
        assert "A1:A10" in cf.range


class TestXlsHandling:
    """`.xls` (旧バイナリ) が来たら空シートで返すこと."""

    def test_xls_returns_empty_sheets(self, tmp_path: Path) -> None:
        # 中身のないファイルでも .xls 拡張子なら早期return される
        fake_xls = tmp_path / "legacy.xls"
        fake_xls.write_bytes(b"not a real xls")
        wb = extract_workbook(fake_xls)
        assert wb.filename == "legacy.xls"
        assert wb.sheets == []
        assert wb.vba_modules == []


class TestErrors:
    def test_missing_file(self, tmp_path: Path) -> None:
        with pytest.raises(ExtractionError):
            extract_workbook(tmp_path / "no_such_file.xlsx")

    def test_corrupt_xlsx(self, tmp_path: Path) -> None:
        bad = tmp_path / "corrupt.xlsx"
        bad.write_bytes(b"definitely not xlsx")
        with pytest.raises(ExtractionError):
            extract_workbook(bad)


class TestEmptyWorkbook:
    def test_no_formulas_no_named_ranges(self, empty_xlsx: Path) -> None:
        wb = extract_workbook(empty_xlsx)
        assert len(wb.sheets) == 1
        sheet = wb.sheets[0]
        assert sheet.formulas == []
        assert sheet.named_ranges == []
        assert sheet.conditional_formats == []
