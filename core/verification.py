"""期待差分と実差分を照合する検証policy gate.

変更プロバイダーの成功終了を安全性の根拠にせず、propose段階で作った期待差分と、
変更後ファイルを再抽出して得た実差分が一致した場合だけ構造検証を通過させる。
"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from core.models import WorkbookDiff

VerificationStatus = Literal["passed", "needs_review", "failed"]
DiffSection = Literal[
    "cells",
    "named_ranges",
    "conditional_formats",
    "data_validations",
    "charts",
    "pivot_tables",
    "vba_modules",
]


class DiffEvidence(BaseModel):
    """policy判定で比較した1件の構造差分."""

    section: DiffSection
    key: str
    change: dict[str, Any]


class PolicyViolation(BaseModel):
    """期待差分と実差分の不一致."""

    code: Literal["missing_expected_change", "unexpected_change", "mismatched_change"]
    section: DiffSection
    key: str
    message: str
    expected: DiffEvidence | None = None
    actual: DiffEvidence | None = None


class VerificationReport(BaseModel):
    """成果物を通過・要確認・不合格に分類する検証結果."""

    status: VerificationStatus
    policy_id: str = "exact-structural-diff-v1"
    expected_change_count: int
    actual_change_count: int
    violations: list[PolicyViolation] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


def _evidence_key(section: DiffSection, change: dict[str, Any]) -> str:
    """差分カテゴリごとの安定した照合キーを返す."""

    if section == "cells":
        return f"{change['sheet']}!{change['coord']}"
    if section == "named_ranges":
        return str(change["name"])
    if section in {"conditional_formats", "data_validations"}:
        return f"{change['sheet']}!{change['range']}"
    if section == "charts":
        return f"{change['sheet']}!{change['key']}"
    if section == "pivot_tables":
        return f"{change['sheet']}!{change['name']}"
    return str(change["name"])


def _collect_evidence(diff: WorkbookDiff) -> dict[tuple[DiffSection, str], DiffEvidence]:
    """WorkbookDiffの構造差分をカテゴリ・キーで逆引き可能にする."""

    sections: tuple[DiffSection, ...] = (
        "cells",
        "named_ranges",
        "conditional_formats",
        "data_validations",
        "charts",
        "pivot_tables",
        "vba_modules",
    )
    evidence: dict[tuple[DiffSection, str], DiffEvidence] = {}
    for section in sections:
        for item in getattr(diff, section):
            change = item.model_dump(mode="json", exclude_none=True)
            key = _evidence_key(section, change)
            evidence[(section, key)] = DiffEvidence(section=section, key=key, change=change)
    return evidence


def verify_expected_diff(expected: WorkbookDiff, actual: WorkbookDiff) -> VerificationReport:
    """期待した構造差分と変更後に観測した実差分を厳密照合する.

    Args:
        expected: propose段階で生成した許可差分。
        actual: 変更後ファイルをフル抽出して得た差分。

    Returns:
        不一致はfailed、動的リスクや波及先があればneeds_review、それ以外はpassed。
    """

    expected_items = _collect_evidence(expected)
    actual_items = _collect_evidence(actual)
    violations: list[PolicyViolation] = []

    for identity in sorted(expected_items.keys() | actual_items.keys()):
        expected_item = expected_items.get(identity)
        actual_item = actual_items.get(identity)
        section, key = identity
        if expected_item is None and actual_item is not None:
            violations.append(
                PolicyViolation(
                    code="unexpected_change",
                    section=section,
                    key=key,
                    message=f"許可されていない変更を検出しました: {section}:{key}",
                    actual=actual_item,
                )
            )
        elif expected_item is not None and actual_item is None:
            violations.append(
                PolicyViolation(
                    code="missing_expected_change",
                    section=section,
                    key=key,
                    message=f"予定した変更が適用されていません: {section}:{key}",
                    expected=expected_item,
                )
            )
        elif (
            expected_item is not None
            and actual_item is not None
            and expected_item.change != actual_item.change
        ):
            violations.append(
                PolicyViolation(
                    code="mismatched_change",
                    section=section,
                    key=key,
                    message=f"変更内容が計画と一致しません: {section}:{key}",
                    expected=expected_item,
                    actual=actual_item,
                )
            )

    warnings: list[str] = []
    if actual.blast_radius:
        warnings.append(f"変更箇所を参照する既存箇所が{len(actual.blast_radius)}件あります。")
    high_risks = [risk for risk in actual.existing_risks if risk.severity == "high"]
    if high_risks:
        warnings.append(f"静的解析で断定できない高リスク項目が{len(high_risks)}件あります。")

    if violations:
        status: VerificationStatus = "failed"
    elif warnings:
        status = "needs_review"
    else:
        status = "passed"

    return VerificationReport(
        status=status,
        expected_change_count=len(expected_items),
        actual_change_count=len(actual_items),
        violations=violations,
        warnings=warnings,
    )
