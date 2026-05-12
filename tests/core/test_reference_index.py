"""core.reference_index のテスト."""

from core.models import (
    CellFormula,
    Reference,
    ReferenceIndex,
    SheetInfo,
    VbaModule,
    VbaProcedure,
    Workbook,
)
from core.reference_index import (
    ParsedRef,
    _canonical_key,
    _col_num_to_letters,
    _find_enclosing_procedure,
    _parse_ref,
    build_reference_index,
    find_overlapping,
)

# ---------- 内部ヘルパーの単体テスト ----------


class TestColNumToLetters:
    def test_a(self) -> None:
        assert _col_num_to_letters(1) == "A"

    def test_h(self) -> None:
        assert _col_num_to_letters(8) == "H"

    def test_z(self) -> None:
        assert _col_num_to_letters(26) == "Z"

    def test_aa(self) -> None:
        assert _col_num_to_letters(27) == "AA"

    def test_zero_or_negative(self) -> None:
        assert _col_num_to_letters(0) == ""
        assert _col_num_to_letters(-1) == ""


class TestFindEnclosingProcedure:
    def _module(self) -> VbaModule:
        return VbaModule(
            name="Module1",
            type="Module",
            code="",
            procedures=[
                VbaProcedure(name="A", kind="Sub", start_line=10, end_line=20, code=""),
                VbaProcedure(name="B", kind="Function", start_line=30, end_line=40, code=""),
            ],
        )

    def test_inside_procedure(self) -> None:
        m = self._module()
        assert _find_enclosing_procedure(m, 15) == "A"
        assert _find_enclosing_procedure(m, 35) == "B"

    def test_at_boundary(self) -> None:
        m = self._module()
        assert _find_enclosing_procedure(m, 10) == "A"
        assert _find_enclosing_procedure(m, 20) == "A"

    def test_outside_any_procedure(self) -> None:
        m = self._module()
        assert _find_enclosing_procedure(m, 1) is None
        assert _find_enclosing_procedure(m, 25) is None
        assert _find_enclosing_procedure(m, 100) is None


# ---------- 数式側 ----------


class TestFormulaSide:
    def test_sumif_three_refs(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(
                    name="Calc",
                    rows=10,
                    cols=10,
                    formulas=[
                        CellFormula(
                            coord="Calc!H2",
                            formula="=SUMIF(Input!A:A, A2, Input!E:E)",
                            refs=["Input!A:A", "A2", "Input!E:E"],
                        )
                    ],
                )
            ],
        )
        idx = build_reference_index(wb)
        # 3 つのキーが正規化された形で立つこと
        # 同一シート相対参照 `A2` はオーナーシート Calc を補って `Calc!A2` に揃う
        assert "Input!A:A" in idx.refs
        assert "Calc!A2" in idx.refs
        assert "Input!E:E" in idx.refs

        ref = idx.refs["Input!A:A"][0]
        assert ref.kind == "formula"
        assert ref.from_ == "Calc!H2"
        assert ref.to == "Input!A:A"
        assert ref.code == "=SUMIF(Input!A:A, A2, Input!E:E)"

    def test_absolute_refs_normalized(self) -> None:
        """`$H$2` / `$H2` / `H2` は同じキーに正規化される."""
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(
                    name="Calc",
                    rows=10,
                    cols=10,
                    formulas=[
                        CellFormula(coord="Calc!A1", formula="=$H$2", refs=["$H$2"]),
                        CellFormula(coord="Calc!A2", formula="=$H2", refs=["$H2"]),
                        CellFormula(coord="Calc!A3", formula="=H2", refs=["H2"]),
                    ],
                )
            ],
        )
        idx = build_reference_index(wb)
        assert "Calc!H2" in idx.refs
        assert len(idx.refs["Calc!H2"]) == 3
        # `$H$2` 等の raw キーは残らない
        assert "$H$2" not in idx.refs
        assert "$H2" not in idx.refs

    def test_unmatched_query_returns_empty(self) -> None:
        wb = Workbook(filename="t.xlsm")
        idx = build_reference_index(wb)
        assert idx.refs.get("NoSuchSheet!Z99", []) == []

    def test_multiple_sources_to_same_target(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(
                    name="Out",
                    rows=10,
                    cols=10,
                    formulas=[
                        CellFormula(coord="Out!A1", formula="=Calc!H2", refs=["Calc!H2"]),
                        CellFormula(coord="Out!A2", formula="=Calc!H2*2", refs=["Calc!H2"]),
                    ],
                )
            ],
        )
        idx = build_reference_index(wb)
        assert len(idx.refs["Calc!H2"]) == 2
        froms = {r.from_ for r in idx.refs["Calc!H2"]}
        assert froms == {"Out!A1", "Out!A2"}


# ---------- VBA側 ----------


def _module_with_code(code: str, procs: list[VbaProcedure] | None = None) -> VbaModule:
    return VbaModule(
        name="Module1",
        type="Module",
        code=code,
        procedures=procs or [],
    )


def _wb_with_module(module: VbaModule) -> Workbook:
    return Workbook(filename="t.xlsm", vba_modules=[module])


class TestVbaPatterns:
    def test_bare_range(self) -> None:
        code = 'Sub UpdateDaily()\n    Range("A1:J100")\nEnd Sub'
        procs = [VbaProcedure(name="UpdateDaily", kind="Sub", start_line=1, end_line=3, code=code)]
        wb = _wb_with_module(_module_with_code(code, procs))
        idx = build_reference_index(wb)

        assert "A1:J100" in idx.refs
        ref = idx.refs["A1:J100"][0]
        assert ref.kind == "vba"
        assert ref.to == "A1:J100"
        assert ref.code == 'Range("A1:J100")'
        assert ref.from_ == "Module1.UpdateDaily:L2"

    def test_worksheets_range(self) -> None:
        code = 'Sub Foo()\n    Worksheets("Calc").Range("A1")\nEnd Sub'
        procs = [VbaProcedure(name="Foo", kind="Sub", start_line=1, end_line=3, code=code)]
        idx = build_reference_index(_wb_with_module(_module_with_code(code, procs)))

        assert "Calc!A1" in idx.refs
        ref = idx.refs["Calc!A1"][0]
        assert ref.from_ == "Module1.Foo:L2"

    def test_sheets_cells(self) -> None:
        # Cells(2, 8) → 行2列8 = "H2"
        code = 'Sub Foo()\n    Sheets("Calc").Cells(2, 8)\nEnd Sub'
        procs = [VbaProcedure(name="Foo", kind="Sub", start_line=1, end_line=3, code=code)]
        idx = build_reference_index(_wb_with_module(_module_with_code(code, procs)))

        assert "Calc!H2" in idx.refs
        assert idx.refs["Calc!H2"][0].kind == "vba"

    def test_bracket_shorthand(self) -> None:
        code = "Sub Foo()\n    [Calc!A1] = 1\nEnd Sub"
        procs = [VbaProcedure(name="Foo", kind="Sub", start_line=1, end_line=3, code=code)]
        idx = build_reference_index(_wb_with_module(_module_with_code(code, procs)))

        assert "Calc!A1" in idx.refs

    def test_all_four_patterns_in_one_module(self) -> None:
        code = (
            "Sub Foo()\n"
            '    Range("A1:J100")\n'
            '    Worksheets("Calc").Range("B1")\n'
            '    Sheets("Input").Cells(2, 8)\n'
            "    [Out!Z99] = 0\n"
            "End Sub\n"
        )
        procs = [VbaProcedure(name="Foo", kind="Sub", start_line=1, end_line=6, code=code)]
        idx = build_reference_index(_wb_with_module(_module_with_code(code, procs)))

        assert "A1:J100" in idx.refs
        assert "Calc!B1" in idx.refs
        assert "Input!H2" in idx.refs
        assert "Out!Z99" in idx.refs

    def test_from_label_outside_procedure(self) -> None:
        # プロシージャの外で書かれた Range は Module1:L行 形式
        code = 'Range("A1")\nSub Foo()\nEnd Sub\n'
        procs = [VbaProcedure(name="Foo", kind="Sub", start_line=2, end_line=3, code=code)]
        idx = build_reference_index(_wb_with_module(_module_with_code(code, procs)))

        assert "A1" in idx.refs
        ref = idx.refs["A1"][0]
        assert ref.from_ == "Module1:L1"

    def test_no_double_match_when_patterns_overlap(self) -> None:
        # Worksheets("Calc").Range("A1") は SHEETS_RANGE で1件、
        # 末尾の Range(...) を BARE が拾わないこと
        code = 'Sub Foo()\n    Worksheets("Calc").Range("A1")\nEnd Sub'
        procs = [VbaProcedure(name="Foo", kind="Sub", start_line=1, end_line=3, code=code)]
        idx = build_reference_index(_wb_with_module(_module_with_code(code, procs)))

        assert "Calc!A1" in idx.refs
        assert len(idx.refs["Calc!A1"]) == 1
        # 単独の "A1" として誤検出されていない
        assert "A1" not in idx.refs

    def test_comment_is_ignored(self) -> None:
        # コメント中の Range は無視される
        code = 'Sub Foo()\n    \' Range("X1")\nEnd Sub'
        procs = [VbaProcedure(name="Foo", kind="Sub", start_line=1, end_line=3, code=code)]
        idx = build_reference_index(_wb_with_module(_module_with_code(code, procs)))

        assert "X1" not in idx.refs


# ---------- 空入力 ----------


class TestEmpty:
    def test_empty_workbook(self) -> None:
        idx = build_reference_index(Workbook(filename="empty.xlsm"))
        assert idx.refs == {}

    def test_empty_module_code(self) -> None:
        wb = _wb_with_module(_module_with_code(""))
        idx = build_reference_index(wb)
        assert idx.refs == {}


# ---------- 参照の正規化 / パース ----------


class TestParseRef:
    def test_single_cell(self) -> None:
        p = _parse_ref("A1")
        assert p == ParsedRef(None, 1, 1, 1, 1)

    def test_owner_sheet_supplied(self) -> None:
        p = _parse_ref("A1", owner_sheet="Calc")
        assert p == ParsedRef("Calc", 1, 1, 1, 1)

    def test_qualified_overrides_owner(self) -> None:
        # シート修飾付きなら owner_sheet は使われない
        p = _parse_ref("Input!H2", owner_sheet="Calc")
        assert p == ParsedRef("Input", 2, 8, 2, 8)

    def test_absolute_markers_stripped(self) -> None:
        assert _parse_ref("$A$1") == ParsedRef(None, 1, 1, 1, 1)
        assert _parse_ref("Calc!$H$2") == ParsedRef("Calc", 2, 8, 2, 8)

    def test_lowercase_normalized(self) -> None:
        assert _parse_ref("a1:b2") == ParsedRef(None, 1, 1, 2, 2)

    def test_range(self) -> None:
        assert _parse_ref("A1:B10") == ParsedRef(None, 1, 1, 10, 2)

    def test_full_column(self) -> None:
        p = _parse_ref("A:A")
        assert p is not None
        assert p.sheet is None
        assert (p.min_col, p.max_col) == (1, 1)
        assert p.min_row == 1
        assert p.max_row > 1_000_000  # Excel hard limit

    def test_full_column_with_sheet(self) -> None:
        p = _parse_ref("Input!A:C")
        assert p is not None
        assert p.sheet == "Input"
        assert (p.min_col, p.max_col) == (1, 3)

    def test_full_row(self) -> None:
        p = _parse_ref("1:5")
        assert p is not None
        assert (p.min_row, p.max_row) == (1, 5)
        assert p.min_col == 1
        assert p.max_col > 10_000

    def test_quoted_sheet_name(self) -> None:
        p = _parse_ref("'My Sheet'!A1")
        assert p == ParsedRef("My Sheet", 1, 1, 1, 1)

    def test_quoted_sheet_with_apostrophe(self) -> None:
        # シート名内の `''` は `'` のエスケープ
        p = _parse_ref("'It''s'!A1")
        assert p == ParsedRef("It's", 1, 1, 1, 1)

    def test_unparseable(self) -> None:
        assert _parse_ref("") is None
        assert _parse_ref("   ") is None
        assert _parse_ref("MyTable[Col1]") is None  # テーブル参照
        assert _parse_ref("SomeFunction(x)") is None


class TestCanonicalKey:
    def test_single_cell(self) -> None:
        assert _canonical_key(ParsedRef("Calc", 2, 8, 2, 8)) == "Calc!H2"

    def test_range(self) -> None:
        assert _canonical_key(ParsedRef("Calc", 1, 1, 10, 2)) == "Calc!A1:B10"

    def test_full_column(self) -> None:
        from core.reference_index import _EXCEL_MAX_ROW

        assert _canonical_key(ParsedRef("Input", 1, 1, _EXCEL_MAX_ROW, 1)) == "Input!A:A"

    def test_full_column_multi(self) -> None:
        from core.reference_index import _EXCEL_MAX_ROW

        assert _canonical_key(ParsedRef("Input", 1, 1, _EXCEL_MAX_ROW, 3)) == "Input!A:C"

    def test_full_row(self) -> None:
        from core.reference_index import _EXCEL_MAX_COL

        assert _canonical_key(ParsedRef("Calc", 1, 1, 1, _EXCEL_MAX_COL)) == "Calc!1:1"

    def test_no_sheet(self) -> None:
        assert _canonical_key(ParsedRef(None, 1, 1, 1, 1)) == "A1"


class TestOverlap:
    def test_same_cell(self) -> None:
        a = ParsedRef("Calc", 1, 1, 1, 1)
        assert a.overlaps(a)

    def test_cell_in_range(self) -> None:
        cell = ParsedRef("Calc", 5, 8, 5, 8)
        rng = ParsedRef("Calc", 1, 1, 100, 10)
        assert rng.overlaps(cell)
        assert cell.overlaps(rng)

    def test_cell_in_full_column(self) -> None:
        from core.reference_index import _EXCEL_MAX_ROW

        col = ParsedRef("Input", 1, 1, _EXCEL_MAX_ROW, 1)
        cell = ParsedRef("Input", 5, 1, 5, 1)
        assert col.overlaps(cell)
        assert cell.overlaps(col)

    def test_different_sheet(self) -> None:
        a = ParsedRef("Calc", 1, 1, 1, 1)
        b = ParsedRef("Input", 1, 1, 1, 1)
        assert not a.overlaps(b)

    def test_disjoint_same_sheet(self) -> None:
        a = ParsedRef("Calc", 1, 1, 1, 1)
        b = ParsedRef("Calc", 10, 10, 10, 10)
        assert not a.overlaps(b)

    def test_none_sheet_is_wildcard(self) -> None:
        # VBA bare 参照 (sheet=None) はシート不問で重なる扱い
        wild = ParsedRef(None, 1, 1, 100, 10)
        in_calc = ParsedRef("Calc", 5, 5, 5, 5)
        in_input = ParsedRef("Input", 5, 5, 5, 5)
        assert wild.overlaps(in_calc)
        assert wild.overlaps(in_input)
        assert in_calc.overlaps(wild)


# ---------- find_overlapping ----------


class TestFindOverlapping:
    def test_exact_match(self) -> None:
        idx = ReferenceIndex(
            refs={"Calc!H2": [Reference(kind="formula", from_="Out!A1", to="Calc!H2")]}
        )
        results = find_overlapping(idx, "Calc!H2")
        assert len(results) == 1
        assert results[0].from_ == "Out!A1"

    def test_cell_hits_full_column(self) -> None:
        idx = ReferenceIndex(
            refs={
                "Input!A:A": [
                    Reference(
                        kind="formula",
                        from_="Calc!H2",
                        to="Input!A:A",
                        code="=SUMIF(Input!A:A,...)",
                    )
                ]
            }
        )
        results = find_overlapping(idx, "Input!A5")
        assert len(results) == 1
        assert results[0].from_ == "Calc!H2"

    def test_cell_hits_enclosing_range(self) -> None:
        idx = ReferenceIndex(
            refs={
                "Calc!A1:J100": [Reference(kind="vba", from_="Module1.Foo:L2", to="Calc!A1:J100")]
            }
        )
        results = find_overlapping(idx, "Calc!H2")
        assert len(results) == 1

    def test_no_hit_different_sheet(self) -> None:
        idx = ReferenceIndex(
            refs={"Input!A:A": [Reference(kind="formula", from_="X!A1", to="Input!A:A")]}
        )
        results = find_overlapping(idx, "Other!A5")
        assert results == []

    def test_vba_bare_is_wildcard(self) -> None:
        # シート不明 (VBA bare) の参照は target のシートに関わらず重なれば返す
        idx = ReferenceIndex(
            refs={
                "A1:J100": [
                    Reference(
                        kind="vba",
                        from_="Module1.Foo:L2",
                        to="A1:J100",
                        code='Range("A1:J100")',
                    )
                ]
            }
        )
        results = find_overlapping(idx, "Calc!H2")
        assert len(results) == 1

    def test_unparseable_target_falls_back_to_exact(self) -> None:
        idx = ReferenceIndex(
            refs={"MyTable[Col1]": [Reference(kind="formula", from_="X!A1", to="MyTable[Col1]")]}
        )
        results = find_overlapping(idx, "MyTable[Col1]")
        assert len(results) == 1
        # 別のテーブル名はヒットしない
        assert find_overlapping(idx, "Other[Col1]") == []

    def test_owner_sheet_for_unqualified_target(self) -> None:
        idx = ReferenceIndex(
            refs={"Calc!H2": [Reference(kind="formula", from_="X!A1", to="Calc!H2")]}
        )
        # 修飾なし `H2` を owner_sheet=Calc で問い合わせるとヒット
        assert len(find_overlapping(idx, "H2", owner_sheet="Calc")) == 1
        # owner_sheet 省略時は sheet=None として全シートにマッチ
        assert len(find_overlapping(idx, "H2")) == 1
