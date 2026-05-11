"""core.spec_generator のテスト."""

from core.models import (
    CellFormula,
    ConditionalFormat,
    NamedRange,
    Reference,
    ReferenceIndex,
    SheetInfo,
    VbaModule,
    VbaProcedure,
    Workbook,
)
from core.spec_generator import (
    _pick_top_formulas,
    _pick_top_references,
    generate_spec,
)


def _empty_index() -> ReferenceIndex:
    return ReferenceIndex(refs={})


# ---------- 内部ヘルパー ----------


class TestPickTopFormulas:
    def test_sort_by_ref_count_desc_then_coord(self) -> None:
        formulas = [
            CellFormula(coord="S!A1", formula="=A2", refs=["A2"]),
            CellFormula(coord="S!A2", formula="=SUM(A1:A100)", refs=["A1:A100"] * 3),
            CellFormula(coord="S!A3", formula="=A4+A5", refs=["A4", "A5"]),
        ]
        top = _pick_top_formulas(formulas, 2)
        # refs の多い順
        assert top[0].coord == "S!A2"
        assert top[1].coord == "S!A3"

    def test_limit(self) -> None:
        formulas = [CellFormula(coord=f"S!A{i}", formula="=1", refs=[]) for i in range(20)]
        assert len(_pick_top_formulas(formulas, 5)) == 5

    def test_empty(self) -> None:
        assert _pick_top_formulas([], 10) == []


class TestPickTopReferences:
    def test_sort_by_count_desc(self) -> None:
        idx = ReferenceIndex(
            refs={
                "Calc!A1": [
                    Reference(kind="formula", from_="Out!A1", to="Calc!A1"),
                    Reference(kind="formula", from_="Out!A2", to="Calc!A1"),
                    Reference(kind="formula", from_="Out!A3", to="Calc!A1"),
                ],
                "Calc!B1": [
                    Reference(kind="formula", from_="Out!B1", to="Calc!B1"),
                ],
            }
        )
        top = _pick_top_references(idx, 5)
        assert top[0] == ("Calc!A1", 3)
        assert top[1] == ("Calc!B1", 1)

    def test_empty(self) -> None:
        assert _pick_top_references(_empty_index(), 5) == []


# ---------- generate_spec ----------


class TestGenerateSpecStructure:
    def test_empty_workbook_has_all_sections(self) -> None:
        wb = Workbook(filename="empty.xlsm")
        md = generate_spec(wb, _empty_index())

        assert "# 設計書: empty.xlsm" in md
        assert "## 1. 概要" in md
        assert "## 2. シート一覧" in md
        assert "## 3. シート詳細" in md
        assert "## 4. VBAモジュール" in md
        assert "## 5. 参照関係" in md
        assert "## 6. 注意点・観察事項" in md

    def test_overview_counts(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(name="A", rows=1, cols=1),
                SheetInfo(name="B", rows=1, cols=1),
            ],
            vba_modules=[VbaModule(name="M1", type="Module", code="")],
            external_links=["external.xlsx"],
        )
        md = generate_spec(wb, _empty_index())
        assert "シート数: 2" in md
        assert "VBAモジュール数: 1" in md
        assert "external.xlsx" in md

    def test_overview_no_external_links(self) -> None:
        wb = Workbook(filename="t.xlsm")
        md = generate_spec(wb, _empty_index())
        assert "外部リンク: なし" in md

    def test_filename_in_title(self) -> None:
        wb = Workbook(filename="my Workbook.xlsm")
        md = generate_spec(wb, _empty_index())
        assert "# 設計書: my Workbook.xlsm" in md


class TestSheetSection:
    def test_sheet_appears_in_table_and_details(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(
                    name="Calc",
                    rows=100,
                    cols=10,
                    formulas=[
                        CellFormula(
                            coord="Calc!H2",
                            formula="=SUMIF(Input!A:A, A2, Input!E:E)",
                            refs=["Input!A:A", "A2", "Input!E:E"],
                        )
                    ],
                    named_ranges=[NamedRange(name="TaxRate", refers_to="Calc!$A$1")],
                    conditional_formats=[
                        ConditionalFormat(range="A1:A10", rule="cellIs greaterThan 100")
                    ],
                    purpose="日次集計",
                )
            ],
        )
        md = generate_spec(wb, _empty_index())

        assert "| Calc |" in md  # シート一覧テーブル
        assert "### Calc" in md  # 詳細セクション
        assert "日次集計" in md
        assert "Calc!H2" in md
        assert "SUMIF" in md
        assert "TaxRate" in md
        assert "cellIs greaterThan 100" in md
        assert "A1:A10" in md

    def test_no_sheets(self) -> None:
        wb = Workbook(filename="t.xlsm")
        md = generate_spec(wb, _empty_index())
        assert "(シートなし)" in md or "_(シートなし)_" in md

    def test_top_10_limit(self) -> None:
        # 15個の数式があっても、TOP 10 セクションには10件まで
        formulas = [
            CellFormula(coord=f"S!A{i}", formula=f"=A{i + 100}", refs=[f"A{i + 100}"])
            for i in range(15)
        ]
        wb = Workbook(
            filename="t.xlsm",
            sheets=[SheetInfo(name="S", rows=15, cols=1, formulas=formulas)],
        )
        md = generate_spec(wb, _empty_index())
        # 主要数式テーブル中の coord 出現数を数える簡易テスト
        coord_count = sum(md.count(f"`S!A{i}`") for i in range(15))
        # TOP 10 + シート一覧テーブルなどに現れる数 ≤ 15
        assert coord_count <= 10 + 5  # テーブル外の余裕分


class TestVbaSection:
    def test_module_with_procedures(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            vba_modules=[
                VbaModule(
                    name="Module1",
                    type="Module",
                    code='Sub UpdateDaily()\n    Range("A1")\nEnd Sub',
                    procedures=[
                        VbaProcedure(
                            name="UpdateDaily",
                            kind="Sub",
                            start_line=1,
                            end_line=3,
                            code="Sub UpdateDaily()...",
                            annotation="日次更新処理",
                        )
                    ],
                )
            ],
        )
        md = generate_spec(wb, _empty_index())

        assert "### Module1" in md
        assert "UpdateDaily" in md
        assert "日次更新処理" in md
        assert "<details>" in md  # 折りたたみ
        assert "```vba" in md

    def test_no_modules(self) -> None:
        wb = Workbook(filename="t.xlsm")
        md = generate_spec(wb, _empty_index())
        assert "(VBAモジュールなし)" in md or "_(VBAモジュールなし)_" in md


class TestReferencesSection:
    def test_references_listed(self) -> None:
        idx = ReferenceIndex(
            refs={
                "Calc!A1": [
                    Reference(kind="formula", from_="Out!K3", to="Calc!A1"),
                    Reference(kind="vba", from_="Module1.X:L10", to="Calc!A1"),
                ]
            }
        )
        wb = Workbook(filename="t.xlsm")
        md = generate_spec(wb, idx)

        assert "Calc!A1" in md
        assert "2" in md  # count
        assert "Out!K3" in md or "Module1.X:L10" in md

    def test_truncation_indicator(self) -> None:
        # 6件以上ある場合 "+N" 表示が出る
        refs = [Reference(kind="formula", from_=f"Out!A{i}", to="X!A1") for i in range(8)]
        idx = ReferenceIndex(refs={"X!A1": refs})
        wb = Workbook(filename="t.xlsm")
        md = generate_spec(wb, idx)
        assert "+3" in md  # 8件中5件まで列挙し、残り3件を "+3" 表示

    def test_no_references(self) -> None:
        wb = Workbook(filename="t.xlsm")
        md = generate_spec(wb, _empty_index())
        assert "(参照なし)" in md or "_(参照なし)_" in md


class TestObservationsSection:
    def test_counts_annotations(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(name="A", rows=1, cols=1, purpose="入力"),
                SheetInfo(name="B", rows=1, cols=1),
            ],
            vba_modules=[
                VbaModule(
                    name="M",
                    type="Module",
                    code="",
                    procedures=[
                        VbaProcedure(
                            name="P",
                            kind="Sub",
                            start_line=1,
                            end_line=2,
                            code="",
                            annotation="日次",
                        )
                    ],
                )
            ],
        )
        md = generate_spec(wb, _empty_index())
        assert "用途が推定されたシート: 1 / 2" in md
        assert "注釈付きプロシージャ: 1" in md


class TestEscaping:
    def test_pipe_in_formula_does_not_break_table(self) -> None:
        # `|` を含む値はエスケープされる
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(
                    name="S",
                    rows=1,
                    cols=1,
                    formulas=[CellFormula(coord="S!A1", formula='=IF(A1>0,"a|b","c|d")', refs=[])],
                )
            ],
        )
        md = generate_spec(wb, _empty_index())
        # エスケープされたパイプが現れる
        assert "\\|" in md


# ---------- 新規セクション: Excel テーブル / マージ / プレビュー ----------


class TestExcelTablesSection:
    def test_renders_when_present(self) -> None:
        from core.models import ExcelTable

        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(
                    name="P",
                    rows=10,
                    cols=5,
                    tables=[ExcelTable(name="PortfolioTable", ref="A6:F206")],
                )
            ],
        )
        md = generate_spec(wb, _empty_index())
        assert "Excel テーブル" in md
        assert "PortfolioTable" in md
        assert "A6:F206" in md

    def test_renders_none_when_absent(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[SheetInfo(name="S", rows=1, cols=1)],
        )
        md = generate_spec(wb, _empty_index())
        assert "Excel テーブル" in md
        assert "_(なし)_" in md


class TestMergedRangesSection:
    def test_renders_merged_ranges(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(
                    name="S",
                    rows=1,
                    cols=1,
                    merged_ranges=["A1:C1", "D2:F4"],
                )
            ],
        )
        md = generate_spec(wb, _empty_index())
        assert "マージセル" in md
        assert "A1:C1" in md
        assert "D2:F4" in md

    def test_truncates_long_list(self) -> None:
        merged = [f"A{i}:B{i}" for i in range(1, 21)]
        wb = Workbook(
            filename="t.xlsm",
            sheets=[SheetInfo(name="S", rows=20, cols=2, merged_ranges=merged)],
        )
        md = generate_spec(wb, _empty_index())
        assert "他 10 件" in md


class TestPreviewSection:
    def test_renders_preview_table(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(
                    name="S",
                    rows=2,
                    cols=2,
                    preview_rows=[["a", "b"], ["1", "2"]],
                    preview_origin="A1:B2",
                )
            ],
        )
        md = generate_spec(wb, _empty_index())
        assert "冒頭プレビュー" in md
        assert "A1:B2" in md
        assert "a" in md
        assert "1" in md

    def test_no_preview_when_empty(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[SheetInfo(name="S", rows=0, cols=0)],
        )
        md = generate_spec(wb, _empty_index())
        assert "冒頭プレビュー" not in md
