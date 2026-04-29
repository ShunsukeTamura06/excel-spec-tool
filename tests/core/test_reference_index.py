"""core.reference_index のテスト."""

from core.models import (
    CellFormula,
    SheetInfo,
    VbaModule,
    VbaProcedure,
    Workbook,
)
from core.reference_index import (
    _col_num_to_letters,
    _find_enclosing_procedure,
    build_reference_index,
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
        # 3つのキーがそれぞれ立つこと
        assert "Input!A:A" in idx.refs
        assert "A2" in idx.refs
        assert "Input!E:E" in idx.refs

        ref = idx.refs["Input!A:A"][0]
        assert ref.kind == "formula"
        assert ref.from_ == "Calc!H2"
        assert ref.to == "Input!A:A"
        assert ref.code == "=SUMIF(Input!A:A, A2, Input!E:E)"

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
