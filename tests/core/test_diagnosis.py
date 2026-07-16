"""根拠付き Excel 診断のテスト."""

from core.diagnosis import build_workbook_diagnosis
from core.models import (
    AnalysisRisk,
    CellFormula,
    FormControl,
    Reference,
    ReferenceIndex,
    SheetInfo,
    VbaModule,
    VbaProcedure,
    Workbook,
)


def test_build_diagnosis_connects_button_to_annotated_macro() -> None:
    """ボタンと対応マクロから利用者向け機能を根拠付きで作る."""
    workbook = Workbook(
        filename="月次集計.xlsm",
        sheets=[
            SheetInfo(
                name="操作",
                rows=10,
                cols=5,
                purpose="月次集計を実行する画面",
                form_controls=[FormControl(text="集計実行", macro="RunSummary", anchor="B2")],
            )
        ],
        vba_modules=[
            VbaModule(
                name="Main",
                type="Module",
                code="",
                procedures=[
                    VbaProcedure(
                        name="RunSummary",
                        kind="Sub",
                        start_line=1,
                        end_line=5,
                        code="Sub RunSummary(): End Sub",
                        annotation="入力データを月次集計します。",
                        side_effects=["結果シート"],
                    )
                ],
            )
        ],
    )

    diagnosis = build_workbook_diagnosis(workbook, ReferenceIndex())

    assert diagnosis.features[0].name == "集計実行"
    assert diagnosis.features[0].summary == "入力データを月次集計します。"
    assert diagnosis.features[0].confidence == "explicit"
    assert diagnosis.features[0].outputs == ["結果シート"]
    assert len(diagnosis.features[0].evidence_ids) == 2
    assert all(
        evidence_id in {item.id for item in diagnosis.evidence}
        for evidence_id in diagnosis.features[0].evidence_ids
    )


def test_build_diagnosis_does_not_invent_business_purpose() -> None:
    """用途の根拠がなければ不明と明示する."""
    workbook = Workbook(
        filename="unknown.xlsx",
        sheets=[
            SheetInfo(
                name="Sheet1",
                rows=2,
                cols=2,
                formulas=[CellFormula(coord="B2", formula="=A2", refs=["A2"])],
            )
        ],
    )

    diagnosis = build_workbook_diagnosis(workbook, ReferenceIndex())

    assert diagnosis.overview.confidence == "unknown"
    assert "特定できません" in diagnosis.overview.text
    assert diagnosis.features == []
    assert diagnosis.coverage.formulas == 1


def test_build_diagnosis_marks_reference_endpoints_as_inferred() -> None:
    """参照グラフ由来の入力・出力候補は推定として扱う."""
    workbook = Workbook(
        filename="flow.xlsx",
        sheets=[
            SheetInfo(name="入力", rows=10, cols=2),
            SheetInfo(name="結果", rows=10, cols=2),
        ],
    )
    references = ReferenceIndex(
        refs={
            "入力!A1": [
                Reference(kind="formula", **{"from": "結果!B1"}, to="入力!A1", code="=入力!A1")
            ]
        }
    )

    diagnosis = build_workbook_diagnosis(workbook, references)

    assert diagnosis.inputs[0].confidence == "inferred"
    assert "入力" in diagnosis.inputs[0].text
    assert diagnosis.outputs[0].confidence == "inferred"
    assert "結果" in diagnosis.outputs[0].text


def test_build_diagnosis_adds_direct_consumers_to_feature_impact() -> None:
    """入力シート機能に、そのシートを参照する直接利用先を紐付ける."""
    workbook = Workbook(
        filename="impact.xlsx",
        sheets=[
            SheetInfo(
                name="商品マスタ",
                rows=10,
                cols=2,
                purpose="商品情報を管理します。",
            ),
            SheetInfo(name="月次レポート", rows=10, cols=2),
        ],
    )
    references = ReferenceIndex(
        refs={
            "商品マスタ!A1": [
                Reference(
                    kind="formula",
                    **{"from": "月次レポート!B1"},
                    to="商品マスタ!A1",
                    code="=商品マスタ!A1",
                )
            ]
        }
    )

    diagnosis = build_workbook_diagnosis(workbook, references)

    assert diagnosis.features[0].related_sheets == ["商品マスタ", "月次レポート"]
    assert diagnosis.features[0].outputs == ["月次レポート"]


def test_build_diagnosis_surfaces_external_dependencies_and_risks() -> None:
    """外部依存と未解析リスクを根拠付きで表示する."""
    workbook = Workbook(
        filename="linked.xlsx",
        sheets=[SheetInfo(name="結果", rows=1, cols=1)],
        external_links=["master.xlsx"],
        analysis_risks=[
            AnalysisRisk(
                category="external_dependency",
                severity="high",
                location="結果!A1",
                evidence="[master.xlsx]Data!A1",
                description="外部ファイルがないと結果が変わる可能性があります。",
                recommendation="外部ファイルを確認してください。",
            )
        ],
    )

    diagnosis = build_workbook_diagnosis(workbook, ReferenceIndex())

    assert diagnosis.external_dependencies[0].confidence == "explicit"
    assert diagnosis.warnings[0].confidence == "explicit"
    assert diagnosis.coverage.external_dependencies == 1


def test_build_diagnosis_groups_repeated_risk_descriptions() -> None:
    """同種リスクを一覧で重複表示せず、根拠だけを束ねる."""
    risks = [
        AnalysisRisk(
            category="dynamic_formula",
            severity="medium",
            location=f"集計!A{row}",
            evidence="=OFFSET(A1, 1, 0)",
            description="OFFSETにより参照先が実行時に変わります。",
            recommendation="実行環境で確認してください。",
        )
        for row in (1, 2)
    ]
    workbook = Workbook(
        filename="risk.xlsx",
        sheets=[SheetInfo(name="集計", rows=2, cols=1)],
        analysis_risks=risks,
    )

    diagnosis = build_workbook_diagnosis(workbook, ReferenceIndex())

    assert len(diagnosis.warnings) == 1
    assert len(diagnosis.warnings[0].evidence_ids) == 2
