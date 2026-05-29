"""core.risk_analyzer のテスト."""

from core.models import (
    CellFormula,
    ChartObject,
    PivotTableInfo,
    PowerQueryInfo,
    SheetInfo,
    VbaModule,
    Workbook,
)
from core.risk_analyzer import detect_analysis_risks


class TestDetectAnalysisRisks:
    def test_detects_dynamic_formula(self) -> None:
        wb = Workbook(
            filename="t.xlsx",
            sheets=[
                SheetInfo(
                    name="Calc",
                    rows=1,
                    cols=1,
                    formulas=[
                        CellFormula(
                            coord="Calc!A1",
                            formula='=INDIRECT("Input!"&A2)',
                            refs=[],
                        )
                    ],
                )
            ],
        )

        risks = detect_analysis_risks(wb)

        assert len(risks) == 1
        assert risks[0].category == "dynamic_formula"
        assert risks[0].severity == "high"
        assert risks[0].location == "Calc!A1"

    def test_detects_dynamic_vba_and_runtime_state(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            vba_modules=[
                VbaModule(
                    name="Module1",
                    type="Module",
                    code=(
                        "Sub X()\n"
                        '  addr = "A" & row\n'
                        "  Range(addr).Value = 1\n"
                        "  Selection.Offset(1, 0).Value = 2\n"
                        "End Sub\n"
                    ),
                )
            ],
        )

        risks = detect_analysis_risks(wb)
        categories = {risk.category for risk in risks}

        assert "dynamic_vba" in categories
        assert "runtime_state" in categories
        assert any("Range" in risk.evidence for risk in risks)
        assert any("Selection" in risk.evidence for risk in risks)

    def test_ignores_literal_range_and_comments(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            vba_modules=[
                VbaModule(
                    name="Module1",
                    type="Module",
                    code=(
                        "Sub X()\n"
                        '  Range("A1").Value = 1\n'
                        "  ' Range(addr).Value = 2\n"
                        '  s = "ActiveSheet"\n'
                        "End Sub\n"
                    ),
                )
            ],
        )

        risks = detect_analysis_risks(wb)

        assert risks == []

    def test_detects_event_macro(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            vba_modules=[
                VbaModule(
                    name="Sheet1",
                    type="Document",
                    code="Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n",
                )
            ],
        )

        risks = detect_analysis_risks(wb)

        assert len(risks) == 1
        assert risks[0].category == "event_macro"
        assert risks[0].severity == "high"

    def test_detects_external_dependencies_and_unknown_objects(self) -> None:
        wb = Workbook(
            filename="t.xlsx",
            sheets=[
                SheetInfo(
                    name="Report",
                    rows=10,
                    cols=10,
                    charts=[ChartObject(name="Chart 1", chart_type="barChart")],
                    pivot_tables=[PivotTableInfo(name="PivotTable1")],
                )
            ],
            external_links=["../source.xlsx"],
            power_queries=[PowerQueryInfo(name="Query - Sales", kind="power_query")],
        )

        risks = detect_analysis_risks(wb)
        categories = {risk.category for risk in risks}

        assert "external_dependency" in categories
        assert "unknown_object_dependency" in categories
        assert any(risk.location == "Connection:Query - Sales" for risk in risks)
