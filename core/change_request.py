"""業務要望と根拠付き診断から改修依頼書を組み立てる."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

from core.diagnosis import WorkbookDiagnosis, WorkbookFeature


class ChangeBrief(BaseModel):
    """ユーザーの自然文要望を安全な改修相談へ渡すための依頼書."""

    title: str
    feature_id: str | None = None
    current_behavior: str
    requested_outcome: str
    affected_areas: list[str] = Field(default_factory=list)
    evidence_ids: list[str] = Field(default_factory=list)
    clarification_questions: list[str] = Field(default_factory=list)
    acceptance_criteria: list[str] = Field(default_factory=list)
    automation: Literal["supported", "needs_review", "unsupported"] = "needs_review"
    automation_reason: str
    next_step: str


def _find_feature(diagnosis: WorkbookDiagnosis, feature_id: str | None) -> WorkbookFeature | None:
    """指定IDの機能を返し、存在しないIDはエラーにする."""
    if feature_id is None:
        return None
    for feature in diagnosis.features:
        if feature.id == feature_id:
            return feature
    raise ValueError(f"unknown feature_id: {feature_id}")


def _is_range_expansion_request(outcome: str) -> bool:
    """数式の参照範囲拡張を明示した要望かを保守的に判定する."""
    normalized = outcome.lower().replace(" ", "").replace("　", "")
    range_targets = ("参照範囲", "数式範囲", "集計範囲")
    expansion_actions = ("広げ", "拡張", "伸ば", "増や")
    direct_phrases = (
        "最終行を増や",
        "最終行を伸ば",
        "行数を増や",
        "データ件数を増や",
        "expandtherange",
        "extendtherange",
    )
    explicit_range_expansion = any(target in normalized for target in range_targets) and any(
        action in normalized for action in expansion_actions
    )
    return explicit_range_expansion or any(phrase in normalized for phrase in direct_phrases)


def build_change_brief(
    diagnosis: WorkbookDiagnosis,
    requested_outcome: str,
    feature_id: str | None = None,
) -> ChangeBrief:
    """診断と業務要望から決定的な改修依頼書を生成する.

    この関数は任意変更を自動適用できるとは判定しない。変更計画への変換と
    MutationProvider の能力照合が完了するまでは必ず ``needs_review`` とする。

    Args:
        diagnosis: 対象ファイルの根拠付き診断。
        requested_outcome: 利用者が実現したい業務上の結果。
        feature_id: 診断に含まれる対象機能ID。ファイル全体なら ``None``。

    Returns:
        改修相談と受入条件確認に使う構造化依頼書。

    Raises:
        ValueError: 要望が空、または機能IDが存在しない場合。
    """
    outcome = " ".join(requested_outcome.strip().split())
    if not outcome:
        raise ValueError("requested_outcome must not be empty")

    feature = _find_feature(diagnosis, feature_id)
    if feature is None:
        title = f"{diagnosis.filename} 全体の改修"
        current_behavior = diagnosis.overview.text
        affected_areas = [item.location for item in diagnosis.evidence[:8]]
        evidence_ids = [item.id for item in diagnosis.evidence[:8]]
    else:
        title = f"「{feature.name}」の改修"
        current_behavior = feature.summary
        affected_areas = [*feature.entry_points, *feature.related_sheets]
        affected_areas.extend(f"入力: {item}" for item in feature.inputs)
        affected_areas.extend(f"出力: {item}" for item in feature.outputs)
        evidence_ids = feature.evidence_ids

    affected_areas = list(dict.fromkeys(item for item in affected_areas if item))
    questions = [
        "変更後、どの操作をしたときに、どの結果になれば完了ですか？",
        "現在の動作で残さなければならない条件や例外はありますか？",
        "確認に使える代表的な入力データと期待結果はありますか？",
    ]
    if diagnosis.external_dependencies:
        questions.append("外部ファイルや外部データがない場合は、どう動くべきですか？")
    if diagnosis.warnings:
        questions.append("未解析の動的処理を、Excel実行環境で確認できますか？")

    criteria = [
        f"変更後に「{outcome}」を代表データで確認できる",
        "対象外のシート、数式、名前定義、VBAに意図しない構造差分がない",
        "原本が保持され、修正版と変更記録を別ファイルとして確認できる",
    ]
    if diagnosis.warnings:
        criteria.append("診断で示された未解析リスクについて、確認結果または未確認の記録がある")

    range_expansion_requested = _is_range_expansion_request(outcome)

    return ChangeBrief(
        title=title,
        feature_id=feature_id,
        current_behavior=current_behavior,
        requested_outcome=outcome,
        affected_areas=affected_areas,
        evidence_ids=evidence_ids,
        clarification_questions=questions,
        acceptance_criteria=criteria,
        automation="supported" if range_expansion_requested else "needs_review",
        automation_reason=(
            "数式が参照するデータ範囲の拡張として、適用前に変更計画を確認できます。"
            if range_expansion_requested
            else "要望に合う改修案をExcelの根拠から作成します。"
        ),
        next_step=(
            "対象範囲と新しい最終行を確認してください。"
            if range_expansion_requested
            else "Excelを調査し、具体的な改修案を提示します。"
        ),
    )
