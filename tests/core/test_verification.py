"""期待差分と実差分を照合するverification policyのテスト."""

from __future__ import annotations

from core.models import (
    AnalysisRisk,
    BlastRadiusEntry,
    CellDiff,
    NamedRangeDiff,
    Reference,
    WorkbookDiff,
)
from core.verification import verify_expected_diff


def _diff(*, cells: list[CellDiff] | None = None) -> WorkbookDiff:
    """検証テスト用の最小WorkbookDiffを作る."""

    return WorkbookDiff(
        before_filename="before.xlsx",
        after_filename="after.xlsx",
        cells=cells or [],
    )


def test_exact_change_passes() -> None:
    """期待差分と実差分が一致すればpassedになる."""

    change = CellDiff(
        sheet="Data",
        coord="C1",
        change_type="modified",
        before_formula="=$B$5*2",
        after_formula="=$B$6*2",
    )

    report = verify_expected_diff(_diff(cells=[change]), _diff(cells=[change]))

    assert report.status == "passed"
    assert report.expected_change_count == 1
    assert report.actual_change_count == 1
    assert report.violations == []


def test_unexpected_change_fails() -> None:
    """計画にないセル変更が加われば不合格になる."""

    expected = CellDiff(sheet="Data", coord="C1", change_type="modified", after_formula="=1")
    unexpected = CellDiff(sheet="Data", coord="D1", change_type="modified", after_value="x")

    report = verify_expected_diff(
        _diff(cells=[expected]),
        _diff(cells=[expected, unexpected]),
    )

    assert report.status == "failed"
    assert [violation.code for violation in report.violations] == ["unexpected_change"]
    assert report.violations[0].key == "Data!D1"


def test_missing_expected_change_fails() -> None:
    """プロバイダーが予定した変更を適用しなければ不合格になる."""

    expected = CellDiff(sheet="Data", coord="C1", change_type="modified", after_formula="=1")

    report = verify_expected_diff(_diff(cells=[expected]), _diff())

    assert report.status == "failed"
    assert report.violations[0].code == "missing_expected_change"


def test_mismatched_change_fails() -> None:
    """同じセルでも変更後の内容が計画と違えば不合格になる."""

    expected = CellDiff(sheet="Data", coord="C1", change_type="modified", after_formula="=1")
    actual = CellDiff(sheet="Data", coord="C1", change_type="modified", after_formula="=2")

    report = verify_expected_diff(_diff(cells=[expected]), _diff(cells=[actual]))

    assert report.status == "failed"
    assert report.violations[0].code == "mismatched_change"


def test_unchanged_number_format_is_context_not_a_mismatch() -> None:
    """実差分だけが同一表示形式を補足しても、変更内容は一致と判定する."""

    expected = CellDiff(
        sheet="Data",
        coord="C1",
        change_type="modified",
        before_formula="=A1",
        after_formula="=A2",
    )
    actual = expected.model_copy(
        update={"before_number_format": "0.0", "after_number_format": "0.0"}
    )

    report = verify_expected_diff(_diff(cells=[expected]), _diff(cells=[actual]))

    assert report.status == "passed"
    assert report.violations == []


def test_changed_number_format_still_fails() -> None:
    """数式変更に表示形式変更が混入した場合は予定外差分として拒否する."""

    expected = CellDiff(
        sheet="Data",
        coord="C1",
        change_type="modified",
        before_formula="=A1",
        after_formula="=A2",
    )
    actual = expected.model_copy(
        update={"before_number_format": "0.0", "after_number_format": "0.00"}
    )

    report = verify_expected_diff(_diff(cells=[expected]), _diff(cells=[actual]))

    assert report.status == "failed"
    assert report.violations[0].code == "mismatched_change"


def test_blast_radius_requires_review() -> None:
    """差分が一致しても既存参照先があれば人間確認へ回す."""

    expected = WorkbookDiff(
        before_filename="before.xlsx",
        after_filename="after.xlsx",
        named_ranges=[
            NamedRangeDiff(
                name="TaxRate",
                change_type="modified",
                before_refers_to="Data!$A$1",
                after_refers_to="Data!$A$2",
            )
        ],
    )
    actual = expected.model_copy(
        update={
            "blast_radius": [
                BlastRadiusEntry(
                    location="Data!A1",
                    change_type="modified",
                    referenced_by=[Reference(kind="formula", from_="Calc!B1", to="Data!A1")],
                )
            ]
        }
    )

    report = verify_expected_diff(expected, actual)

    assert report.status == "needs_review"
    assert "1件" in report.warnings[0]


def test_high_analysis_risk_requires_review() -> None:
    """高リスクな未解析項目が残る成果物は自動合格させない."""

    change = CellDiff(sheet="Data", coord="C1", change_type="modified", after_formula="=1")
    actual = _diff(cells=[change])
    actual.existing_risks = [
        AnalysisRisk(
            category="dynamic_formula",
            severity="high",
            location="Data!D1",
            evidence="=INDIRECT(A1)",
            description="参照先が動的です",
            recommendation="Excelで確認してください",
        )
    ]

    report = verify_expected_diff(_diff(cells=[change]), actual)

    assert report.status == "needs_review"
