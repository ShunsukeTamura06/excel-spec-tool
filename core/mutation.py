"""変更計画と変更プロバイダーの共通境界.

変更を作る実装と、変更後を検証する xlblueprint の責務を分離する。変更計画は
プロバイダー非依存で表現し、既存の openpyxl 実装や外部 CLI を同じ契約で扱う。
"""

from __future__ import annotations

import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Annotated, Literal, Protocol

from pydantic import BaseModel, Field

from core.exceptions import UnsupportedMutationError
from core.formula_fix import (
    apply_fixed_ref_replace,
    apply_range_expansion,
    propose_fixed_ref_replace,
    propose_range_expansion,
)
from core.models import CellDiff, ReferenceIndex, Workbook, WorkbookDiff
from core.named_range_fix import apply_named_range_fix, propose_named_range_fix
from core.vba_change import propose_vba_procedure_replace

MutationKind = Literal[
    "named_range_set",
    "fixed_ref_replace",
    "range_expansion",
    "cell_text_batch",
    "vba_procedure_replace",
]
ProviderName = Literal["openpyxl", "officecli", "windows_vbide"]


def _utc_now_iso() -> str:
    """現在時刻をUTCのISO 8601文字列で返す."""

    return datetime.now(timezone.utc).isoformat()


class NamedRangeSetOperation(BaseModel):
    """ワークブックスコープの名前定義を変更する操作."""

    kind: Literal["named_range_set"] = "named_range_set"
    name: str
    new_refers_to: str


class FixedRefReplaceOperation(BaseModel):
    """数式内の固定参照を別の参照へ置換する操作."""

    kind: Literal["fixed_ref_replace"] = "fixed_ref_replace"
    old_ref: str
    new_ref: str


class RangeExpansionOperation(BaseModel):
    """数式内の参照範囲を包含関係を保って拡張する操作."""

    kind: Literal["range_expansion"] = "range_expansion"
    old_ref: str
    new_ref: str


class CellTextEdit(BaseModel):
    """空セルへ追加する固定テキスト1件."""

    sheet: str = Field(min_length=1, max_length=31)
    coord: str = Field(pattern=r"^[A-Z]{1,3}[1-9][0-9]*$")
    value: str = Field(min_length=1, max_length=2000)


class CellTextBatchOperation(BaseModel):
    """既存セルを上書きせず、複数の空セルへ固定テキストを追加する操作."""

    kind: Literal["cell_text_batch"] = "cell_text_batch"
    edits: list[CellTextEdit] = Field(min_length=1, max_length=50)


class VbaProcedureReplaceOperation(BaseModel):
    """既存VBAプロシージャを同名・同種の完全なコードへ置換する操作."""

    kind: Literal["vba_procedure_replace"] = "vba_procedure_replace"
    module_name: str = Field(min_length=1, max_length=128)
    procedure_name: str = Field(min_length=1, max_length=128)
    new_code: str = Field(min_length=1, max_length=12_000)


MutationOperation = Annotated[
    NamedRangeSetOperation
    | FixedRefReplaceOperation
    | RangeExpansionOperation
    | CellTextBatchOperation
    | VbaProcedureReplaceOperation,
    Field(discriminator="kind"),
]


class MutationPlan(BaseModel):
    """監査可能な、プロバイダー非依存の変更計画."""

    plan_id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    created_at: str = Field(default_factory=_utc_now_iso)
    source_job_id: str
    requested_provider: ProviderName = "openpyxl"
    operation: MutationOperation


class SafeChangePlan(BaseModel):
    """一般ユーザーが適用前に確認する、期待差分付きの安全変更計画."""

    plan: MutationPlan
    automation: Literal["supported", "needs_review"]
    can_apply: bool = True
    title: str
    summary: str
    expected_diff: WorkbookDiff
    expected_change_count: int
    affected_locations: list[str] = Field(default_factory=list)
    preconditions: list[str] = Field(default_factory=list)
    acceptance_criteria: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    verification_scope: str


class ProviderCapability(BaseModel):
    """変更プロバイダーが現在提供できる機能."""

    name: ProviderName
    available: bool
    version: str | None = None
    supported_extensions: list[str] = Field(default_factory=list)
    supported_operations: list[MutationKind] = Field(default_factory=list)
    unavailable_reason: str | None = None


class MutationResult(BaseModel):
    """変更プロバイダーによる適用結果の監査用メタデータ."""

    provider: ProviderName
    provider_version: str | None = None
    operation: MutationKind
    changed_count: int


class MutationProvider(Protocol):
    """元ファイルを変更せず、別ファイルへ変更結果を書き出す契約."""

    def capability(self) -> ProviderCapability:
        """現在の利用可否と対応範囲を返す."""

    def apply(self, plan: MutationPlan, source_path: Path, out_path: Path) -> MutationResult:
        """planをsourceへ適用し、結果をout_pathへ書き出す."""


def propose_mutation(
    plan: MutationPlan,
    before_wb: Workbook,
    before_index: ReferenceIndex,
) -> WorkbookDiff:
    """変更計画をファイルへ書き込まず、期待される構造差分を計算する.

    Args:
        plan: 実行予定の変更計画。
        before_wb: 変更前の抽出済みワークブック。
        before_index: 変更前の逆参照インデックス。

    Returns:
        policy gateで実差分と照合する期待差分。
    """

    operation = plan.operation
    if isinstance(operation, VbaProcedureReplaceOperation):
        return propose_vba_procedure_replace(
            before_wb,
            operation.module_name,
            operation.procedure_name,
            operation.new_code,
        )
    if isinstance(operation, CellTextBatchOperation):
        seen: set[tuple[str, str]] = set()
        cells: list[CellDiff] = []
        for edit in operation.edits:
            identity = (edit.sheet, edit.coord)
            if identity in seen:
                raise UnsupportedMutationError(
                    f"duplicate cell text edit: {edit.sheet}!{edit.coord}"
                )
            seen.add(identity)
            cells.append(
                CellDiff(
                    sheet=edit.sheet,
                    coord=edit.coord,
                    change_type="added",
                    after_value=edit.value,
                )
            )
        return WorkbookDiff(
            before_filename=before_wb.filename,
            after_filename=before_wb.filename,
            cells=cells,
            existing_risks=list(before_wb.analysis_risks),
        )
    if isinstance(operation, NamedRangeSetOperation):
        return propose_named_range_fix(
            before_wb,
            before_index,
            operation.name,
            operation.new_refers_to,
        )
    if isinstance(operation, FixedRefReplaceOperation):
        return propose_fixed_ref_replace(
            before_wb,
            before_index,
            operation.old_ref,
            operation.new_ref,
        )
    return propose_range_expansion(
        before_wb,
        before_index,
        operation.old_ref,
        operation.new_ref,
    )


def build_safe_change_plan(
    plan: MutationPlan,
    before_wb: Workbook,
    before_index: ReferenceIndex,
) -> SafeChangePlan:
    """変更を書き込まず、一般ユーザー向けの変更計画と期待差分を作る.

    Args:
        plan: 実行時にも同じIDで使用する変更計画。
        before_wb: 変更前の抽出済みワークブック。
        before_index: 変更前の参照索引。

    Returns:
        期待差分、確認事項、静的検証境界を含む変更計画。
    """
    expected_diff = propose_mutation(plan, before_wb, before_index)
    operation = plan.operation
    changed_locations = [f"{item.sheet}!{item.coord}" for item in expected_diff.cells]
    changed_locations.extend(item.name for item in expected_diff.named_ranges)
    changed_locations.extend(f"VBA:{item.name}" for item in expected_diff.vba_modules)
    expected_change_count = (
        len(expected_diff.cells)
        + len(expected_diff.named_ranges)
        + len(expected_diff.conditional_formats)
        + len(expected_diff.data_validations)
        + len(expected_diff.charts)
        + len(expected_diff.pivot_tables)
        + len(expected_diff.vba_modules)
    )

    if isinstance(operation, VbaProcedureReplaceOperation):
        title = f"{operation.module_name}.{operation.procedure_name} を置き換える"
        summary = (
            f"既存VBAプロシージャ {operation.module_name}.{operation.procedure_name} を"
            "Windows Excelで置き換えるパッケージを作ります。"
        )
        preconditions = [
            "Windows版Microsoft Excelを利用できる",
            "対象VBAプロジェクトがロックされていない",
            "VBAプロジェクト オブジェクト モデルへのアクセスを一時的に許可できる",
            "置換後コードが同名・同種の完全なSubまたはFunctionである",
        ]
    elif isinstance(operation, CellTextBatchOperation):
        title = "説明テキストを追加する"
        summary = f"{len(operation.edits)}個の空セルへ説明テキストを追加します。"
        preconditions = [
            "変更対象として表示されたセルが現在空欄である",
            "追加する説明文が業務上の意味と一致している",
        ]
    elif isinstance(operation, RangeExpansionOperation):
        title = "数式が参照するデータ範囲を広げる"
        summary = (
            f"{operation.old_ref} を {operation.new_ref} へ広げ、"
            f"{len(expected_diff.cells)}件の数式参照を更新します。"
        )
        preconditions = [
            "新しい範囲が現在の範囲を完全に含んでいる",
            "追加する行の列構成が現在のデータと同じである",
            "変更対象として表示された数式が業務上の想定と一致している",
        ]
    elif isinstance(operation, NamedRangeSetOperation):
        title = "名前付き範囲の参照先を変更する"
        summary = f"{operation.name} の参照先を {operation.new_refers_to} へ変更します。"
        preconditions = ["新しい参照先が存在する", "名前の用途が変更後の範囲と一致する"]
    else:
        title = "数式の固定参照を置き換える"
        summary = (
            f"{operation.old_ref} を {operation.new_ref} へ置き換え、"
            f"{len(expected_diff.cells)}件の数式を更新します。"
        )
        preconditions = [
            "新しい参照先が存在する",
            "変更対象として表示された数式が業務上の想定と一致している",
        ]

    warnings = list(dict.fromkeys(item.description for item in expected_diff.existing_risks))
    if isinstance(operation, VbaProcedureReplaceOperation):
        warnings.insert(
            0,
            "Windowsで適用後、revised.xlsmを戻して静的差分検証するまで完了扱いにしません。",
        )
        warnings.insert(
            1,
            "VBA署名がある場合、保存により署名が無効になります。",
        )
    if expected_diff.blast_radius:
        warnings.insert(
            0,
            f"変更対象を参照する箇所が{len(expected_diff.blast_radius)}件あります。",
        )
    automation: Literal["supported", "needs_review"] = "needs_review" if warnings else "supported"

    return SafeChangePlan(
        plan=plan,
        automation=automation,
        title=title,
        summary=summary,
        expected_diff=expected_diff,
        expected_change_count=expected_change_count,
        affected_locations=changed_locations,
        preconditions=preconditions,
        acceptance_criteria=[
            "表示された予定変更と、適用後に再抽出した実差分が一致する",
            "予定外のセル・名前定義・書式・入力規則・グラフ・ピボット・VBA変更がない",
            "原本が保持され、修正版と監査記録を取得できる",
        ],
        warnings=warnings,
        verification_scope=(
            (
                "Windowsで生成された.xlsmを再抽出し、期待したVBAモジュール差分だけを"
                "静的検証します。VBAのコンパイルとマクロ実行結果はまだ保証しません。"
            )
            if isinstance(operation, VbaProcedureReplaceOperation)
            else (
                "Macではファイル構造の一致を検証します。Excelでの再計算値とマクロ実行は"
                "Windows + Microsoft Excelでの追加確認が必要です。"
            )
        ),
    )


class OpenPyxlMutationProvider:
    """既存の限定安全修正を共通契約へ接続するプロバイダー."""

    _VERSION = "openpyxl-adapter-v1"

    def capability(self) -> ProviderCapability:
        """openpyxlプロバイダーの対応範囲を返す."""

        return ProviderCapability(
            name="openpyxl",
            available=True,
            version=self._VERSION,
            supported_extensions=[".xlsx", ".xlsm"],
            supported_operations=["named_range_set", "fixed_ref_replace", "range_expansion"],
        )

    def apply(self, plan: MutationPlan, source_path: Path, out_path: Path) -> MutationResult:
        """既存の決定的な修正関数で変更計画を適用する.

        Args:
            plan: 適用する変更計画。
            source_path: 変更しない元ファイル。
            out_path: 変更後ファイルの出力先。

        Returns:
            適用件数を含む結果。

        Raises:
            UnsupportedMutationError: 対応外の拡張子が指定された場合。
        """

        suffix = source_path.suffix.lower()
        if suffix not in {".xlsx", ".xlsm"}:
            raise UnsupportedMutationError(f"openpyxl provider does not support {suffix}")

        operation = plan.operation
        changed_count: int
        if isinstance(operation, NamedRangeSetOperation):
            apply_named_range_fix(
                source_path,
                operation.name,
                operation.new_refers_to,
                out_path,
            )
            changed_count = 1
        elif isinstance(operation, FixedRefReplaceOperation):
            changed_count = apply_fixed_ref_replace(
                source_path,
                operation.old_ref,
                operation.new_ref,
                out_path,
            )
        elif isinstance(operation, RangeExpansionOperation):
            changed_count = apply_range_expansion(
                source_path,
                operation.old_ref,
                operation.new_ref,
                out_path,
            )
        else:  # pragma: no cover - Pydanticのdiscriminated unionが防ぐ防御分岐
            raise UnsupportedMutationError(f"unsupported operation: {operation!r}")

        return MutationResult(
            provider="openpyxl",
            provider_version=self._VERSION,
            operation=operation.kind,
            changed_count=changed_count,
        )
