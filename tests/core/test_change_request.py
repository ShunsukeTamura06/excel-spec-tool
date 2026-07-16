"""改修依頼書ビルダーのテスト."""

import pytest

from core.change_request import build_change_brief
from core.diagnosis import GroundedClaim, WorkbookDiagnosis, WorkbookFeature


def _diagnosis() -> WorkbookDiagnosis:
    """テスト用の最小診断を返す."""
    return WorkbookDiagnosis(
        filename="tool.xlsm",
        headline=GroundedClaim(text="Excelツール", confidence="explicit"),
        overview=GroundedClaim(text="月次集計を行います。", confidence="inferred"),
        features=[
            WorkbookFeature(
                id="F001",
                name="集計実行",
                summary="入力データを月次集計します。",
                confidence="explicit",
                entry_points=["操作!B2"],
                related_sheets=["操作", "結果"],
                outputs=["結果シート"],
                evidence_ids=["E001", "E002"],
            )
        ],
    )


def test_build_change_brief_uses_selected_feature_evidence() -> None:
    """対象機能の現状・影響候補・根拠を引き継ぐ."""
    brief = build_change_brief(_diagnosis(), "集計結果に担当部署を追加したい", "F001")

    assert brief.title == "「集計実行」の改修"
    assert brief.current_behavior == "入力データを月次集計します。"
    assert brief.requested_outcome == "集計結果に担当部署を追加したい"
    assert "結果" in brief.affected_areas
    assert brief.evidence_ids == ["E001", "E002"]
    assert brief.automation == "needs_review"
    assert len(brief.acceptance_criteria) >= 3


def test_build_change_brief_can_target_whole_workbook() -> None:
    """機能を選べない場合もファイル全体の相談として整理する."""
    brief = build_change_brief(_diagnosis(), "入力ミスを減らしたい")

    assert brief.feature_id is None
    assert brief.current_behavior == "月次集計を行います。"


@pytest.mark.parametrize("value", ["", "   ", "\n"])
def test_build_change_brief_rejects_empty_request(value: str) -> None:
    """空の業務要望は受け付けない."""
    with pytest.raises(ValueError, match="must not be empty"):
        build_change_brief(_diagnosis(), value)


def test_build_change_brief_rejects_unknown_feature() -> None:
    """診断にない機能IDを対象にできない."""
    with pytest.raises(ValueError, match="unknown feature_id"):
        build_change_brief(_diagnosis(), "変更したい", "F999")
