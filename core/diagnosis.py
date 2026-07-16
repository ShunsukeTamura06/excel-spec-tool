"""抽出結果から根拠付きの Excel 診断を組み立てる.

LLM の有無にかかわらず同じ入力から同じ診断を生成し、推定と抽出事実を混同しない。
"""

from __future__ import annotations

from collections import defaultdict
from typing import Literal

from pydantic import BaseModel, Field

from core.models import FormControl, ReferenceIndex, SheetInfo, VbaProcedure, Workbook

Confidence = Literal["explicit", "inferred", "unknown"]


class DiagnosisEvidence(BaseModel):
    """診断中の主張を裏付ける、抽出済みの事実."""

    id: str
    kind: Literal["sheet", "control", "vba", "reference", "external", "risk"]
    location: str
    detail: str


class GroundedClaim(BaseModel):
    """確度と根拠を持つ短い説明."""

    text: str
    confidence: Confidence
    evidence_ids: list[str] = Field(default_factory=list)


class WorkbookFeature(BaseModel):
    """利用者が認識できる業務機能の候補."""

    id: str
    name: str
    summary: str
    confidence: Confidence
    entry_points: list[str] = Field(default_factory=list)
    related_sheets: list[str] = Field(default_factory=list)
    inputs: list[str] = Field(default_factory=list)
    outputs: list[str] = Field(default_factory=list)
    evidence_ids: list[str] = Field(default_factory=list)


class DiagnosisCoverage(BaseModel):
    """診断で確認できた構造の件数."""

    sheets: int = 0
    formulas: int = 0
    vba_procedures: int = 0
    controls: int = 0
    external_dependencies: int = 0


class WorkbookDiagnosis(BaseModel):
    """一般ユーザー向けの根拠付き Excel 診断."""

    filename: str
    headline: GroundedClaim
    overview: GroundedClaim
    features: list[WorkbookFeature] = Field(default_factory=list)
    inputs: list[GroundedClaim] = Field(default_factory=list)
    outputs: list[GroundedClaim] = Field(default_factory=list)
    external_dependencies: list[GroundedClaim] = Field(default_factory=list)
    warnings: list[GroundedClaim] = Field(default_factory=list)
    evidence: list[DiagnosisEvidence] = Field(default_factory=list)
    coverage: DiagnosisCoverage = Field(default_factory=DiagnosisCoverage)
    limitations: list[str] = Field(default_factory=list)


class _EvidenceBuilder:
    """連番の証拠IDを付けながら証拠一覧を作る内部ヘルパー."""

    def __init__(self) -> None:
        self.items: list[DiagnosisEvidence] = []

    def add(
        self,
        kind: Literal["sheet", "control", "vba", "reference", "external", "risk"],
        location: str,
        detail: str,
    ) -> str:
        """証拠を追加してIDを返す."""
        evidence_id = f"E{len(self.items) + 1:03d}"
        self.items.append(
            DiagnosisEvidence(id=evidence_id, kind=kind, location=location, detail=detail)
        )
        return evidence_id


def _sheet_summary(sheet: SheetInfo) -> str:
    """シートから明示的に数えられる構造を要約する."""
    parts = [f"使用範囲 {sheet.rows}行×{sheet.cols}列"]
    if sheet.formulas:
        parts.append(f"数式 {len(sheet.formulas)}件")
    if sheet.tables:
        parts.append(f"テーブル {len(sheet.tables)}件")
    if sheet.charts:
        parts.append(f"グラフ {len(sheet.charts)}件")
    if sheet.pivot_tables:
        parts.append(f"ピボット {len(sheet.pivot_tables)}件")
    if sheet.form_controls:
        parts.append(f"操作部品 {len(sheet.form_controls)}件")
    return "、".join(parts)


def _procedure_by_name(wb: Workbook) -> dict[str, tuple[str, VbaProcedure]]:
    """大文字小文字を無視した VBA プロシージャ索引を返す."""
    result: dict[str, tuple[str, VbaProcedure]] = {}
    for module in wb.vba_modules:
        for procedure in module.procedures:
            result[procedure.name.casefold()] = (module.name, procedure)
    return result


def _is_generic_sheet_name(name: str) -> bool:
    """Excelの初期名に近い、業務上の意味を持たないシート名か判定する."""
    normalized = name.casefold().replace(" ", "")
    return normalized.startswith("sheet") or normalized.startswith("シート")


def _control_feature(
    feature_no: int,
    sheet: SheetInfo,
    control: FormControl,
    evidence: _EvidenceBuilder,
    procedures: dict[str, tuple[str, VbaProcedure]],
) -> WorkbookFeature:
    """フォームコントロールを利用者向け機能へ変換する."""
    label = control.text or control.name or control.macro or f"{control.kind}操作"
    location = f"{sheet.name}!{control.anchor}" if control.anchor else sheet.name
    detail = f"{control.kind}「{label}」"
    if control.macro:
        detail += f" からマクロ {control.macro} を起動"
    evidence_ids = [evidence.add("control", location, detail)]
    summary = f"「{label}」を操作して処理を実行します。"
    inputs = list(sheet.inputs)
    outputs = list(sheet.outputs)

    procedure_entry = procedures.get(control.macro.split("!")[-1].casefold())
    if procedure_entry is not None:
        module_name, procedure = procedure_entry
        evidence_ids.append(
            evidence.add(
                "vba",
                f"{module_name}.{procedure.name}",
                procedure.annotation
                or f"{procedure.kind}、{procedure.start_line}〜{procedure.end_line}行",
            )
        )
        if procedure.annotation:
            summary = procedure.annotation
        outputs.extend(procedure.side_effects)

    return WorkbookFeature(
        id=f"F{feature_no:03d}",
        name=label,
        summary=summary,
        confidence="explicit",
        entry_points=[location],
        related_sheets=[sheet.name],
        inputs=list(dict.fromkeys(inputs)),
        outputs=list(dict.fromkeys(outputs)),
        evidence_ids=evidence_ids,
    )


def _sheet_graph(idx: ReferenceIndex, sheet_names: set[str]) -> tuple[set[str], set[str]]:
    """参照グラフから入力側・出力側のシート候補を返す."""
    outgoing: dict[str, set[str]] = defaultdict(set)
    incoming: dict[str, set[str]] = defaultdict(set)
    for target, refs in idx.refs.items():
        target_sheet = target.split("!", 1)[0].strip("'") if "!" in target else ""
        if target_sheet not in sheet_names:
            continue
        for ref in refs:
            source_sheet = ref.from_.split("!", 1)[0].strip("'") if "!" in ref.from_ else ""
            if source_sheet in sheet_names and source_sheet != target_sheet:
                outgoing[source_sheet].add(target_sheet)
                incoming[target_sheet].add(source_sheet)
    inputs = {name for name in sheet_names if incoming[name] and not outgoing[name]}
    outputs = {name for name in sheet_names if outgoing[name] and not incoming[name]}
    return inputs, outputs


def _sheet_relations(
    idx: ReferenceIndex,
    sheet_names: set[str],
) -> tuple[dict[str, set[str]], dict[str, set[str]]]:
    """各シートの依存元と、そのシートを参照する利用先を返す."""
    dependencies: dict[str, set[str]] = defaultdict(set)
    consumers: dict[str, set[str]] = defaultdict(set)
    for target, refs in idx.refs.items():
        target_sheet = target.split("!", 1)[0].strip("'") if "!" in target else ""
        if target_sheet not in sheet_names:
            continue
        for ref in refs:
            source_sheet = ref.from_.split("!", 1)[0].strip("'") if "!" in ref.from_ else ""
            if source_sheet in sheet_names and source_sheet != target_sheet:
                dependencies[source_sheet].add(target_sheet)
                consumers[target_sheet].add(source_sheet)
    return dependencies, consumers


def build_workbook_diagnosis(wb: Workbook, idx: ReferenceIndex) -> WorkbookDiagnosis:
    """抽出済みモデルから決定的な根拠付き診断を生成する.

    Args:
        wb: 抽出・注釈済みのワークブック。
        idx: セル、VBA、オブジェクト間の参照索引。

    Returns:
        一般ユーザー向けの診断。推定項目は ``inferred`` として返す。
    """
    evidence = _EvidenceBuilder()
    sheet_evidence: dict[str, str] = {}
    for sheet in wb.sheets:
        detail = _sheet_summary(sheet)
        if sheet.purpose:
            detail += f"。解析上の用途: {sheet.purpose}"
        sheet_evidence[sheet.name] = evidence.add("sheet", sheet.name, detail)

    formula_count = sum(len(sheet.formulas) for sheet in wb.sheets)
    controls = [(sheet, control) for sheet in wb.sheets for control in sheet.form_controls]
    procedure_count = sum(len(module.procedures) for module in wb.vba_modules)
    facts = [f"{len(wb.sheets)}シート", f"{formula_count}数式"]
    if procedure_count:
        facts.append(f"{procedure_count}個の自動処理")
    if controls:
        facts.append(f"{len(controls)}個の操作部品")
    headline_ids = list(sheet_evidence.values())
    headline = GroundedClaim(
        text=f"{wb.filename} は、{'・'.join(facts)}を含むExcelツールです。",
        confidence="explicit",
        evidence_ids=headline_ids,
    )

    purpose_sheets = [sheet for sheet in wb.sheets if sheet.purpose]
    if purpose_sheets:
        overview = GroundedClaim(
            text=" ".join(f"{sheet.name}: {sheet.purpose}" for sheet in purpose_sheets[:4]),
            confidence="inferred",
            evidence_ids=[sheet_evidence[sheet.name] for sheet in purpose_sheets[:4]],
        )
    else:
        named_sheets = [sheet for sheet in wb.sheets if not _is_generic_sheet_name(sheet.name)]
        if named_sheets:
            shown_names = "、".join(sheet.name for sheet in named_sheets[:6])
            suffix = "など" if len(named_sheets) > 6 else ""
            overview = GroundedClaim(
                text=(
                    f"シート名と構造から、{shown_names}{suffix}を扱う構成と推定できます。"
                    "実際の業務手順は利用者への確認が必要です。"
                ),
                confidence="inferred",
                evidence_ids=[sheet_evidence[sheet.name] for sheet in named_sheets[:6]],
            )
        else:
            overview = GroundedClaim(
                text="ファイル構造は確認できましたが、業務上の目的は構造だけでは特定できません。",
                confidence="unknown",
                evidence_ids=headline_ids,
            )

    sheet_names = {sheet.name for sheet in wb.sheets}
    graph_inputs, graph_outputs = _sheet_graph(idx, sheet_names)
    sheet_dependencies, sheet_consumers = _sheet_relations(idx, sheet_names)

    features: list[WorkbookFeature] = []
    procedures = _procedure_by_name(wb)
    for sheet, control in controls:
        features.append(_control_feature(len(features) + 1, sheet, control, evidence, procedures))

    for module in wb.vba_modules:
        for procedure in module.procedures:
            if not procedure.triggers or any(
                procedure.name.casefold() == control.macro.split("!")[-1].casefold()
                for _, control in controls
            ):
                continue
            evidence_id = evidence.add(
                "vba",
                f"{module.name}.{procedure.name}",
                procedure.annotation or f"起動契機: {', '.join(procedure.triggers)}",
            )
            features.append(
                WorkbookFeature(
                    id=f"F{len(features) + 1:03d}",
                    name=procedure.annotation or procedure.name,
                    summary=procedure.annotation or "Excelの操作や状態を契機に実行される処理です。",
                    confidence="inferred",
                    entry_points=procedure.triggers,
                    outputs=procedure.side_effects,
                    evidence_ids=[evidence_id],
                )
            )

    featured_sheets = {name for feature in features for name in feature.related_sheets}
    for sheet in wb.sheets:
        if sheet.name in featured_sheets:
            continue
        if not (
            sheet.purpose
            or sheet.tables
            or sheet.charts
            or sheet.pivot_tables
            or sheet.name in graph_outputs
        ):
            continue
        features.append(
            WorkbookFeature(
                id=f"F{len(features) + 1:03d}",
                name=sheet.purpose or f"{sheet.name}シートの処理",
                summary=sheet.purpose or _sheet_summary(sheet),
                confidence="inferred",
                related_sheets=[sheet.name],
                inputs=sheet.inputs,
                outputs=sheet.outputs,
                evidence_ids=[sheet_evidence[sheet.name]],
            )
        )
        if len(features) >= 8:
            break

    for feature in features:
        if len(feature.related_sheets) != 1:
            continue
        primary_sheet = feature.related_sheets[0]
        dependency_sheets = sorted(sheet_dependencies[primary_sheet])
        consumers = sorted(sheet_consumers[primary_sheet])
        feature.related_sheets = list(
            dict.fromkeys([primary_sheet, *dependency_sheets, *consumers])
        )
        feature.inputs = list(dict.fromkeys([*feature.inputs, *dependency_sheets]))
        feature.outputs = list(dict.fromkeys([*feature.outputs, *consumers]))

    input_claims: list[GroundedClaim] = []
    output_claims: list[GroundedClaim] = []
    for sheet in wb.sheets:
        for item in sheet.inputs:
            input_claims.append(
                GroundedClaim(
                    text=f"{sheet.name} は {item} を入力として使う可能性があります。",
                    confidence="inferred",
                    evidence_ids=[sheet_evidence[sheet.name]],
                )
            )
        for item in sheet.outputs:
            output_claims.append(
                GroundedClaim(
                    text=f"{sheet.name} から {item} へ結果が渡る可能性があります。",
                    confidence="inferred",
                    evidence_ids=[sheet_evidence[sheet.name]],
                )
            )

    if not input_claims:
        input_claims.extend(
            GroundedClaim(
                text=f"{name} は他シートから参照される入力側の候補です。",
                confidence="inferred",
                evidence_ids=[sheet_evidence[name]],
            )
            for name in sorted(graph_inputs)
        )
    if not output_claims:
        output_claims.extend(
            GroundedClaim(
                text=f"{name} は他シートを参照する出力側の候補です。",
                confidence="inferred",
                evidence_ids=[sheet_evidence[name]],
            )
            for name in sorted(graph_outputs)
        )

    dependencies: list[GroundedClaim] = []
    for link in wb.external_links:
        evidence_id = evidence.add("external", "外部リンク", link)
        dependencies.append(
            GroundedClaim(
                text=f"外部ファイル参照: {link}",
                confidence="explicit",
                evidence_ids=[evidence_id],
            )
        )
    for query in wb.power_queries:
        detail = query.name
        if query.source:
            detail += f" ({query.source})"
        evidence_id = evidence.add("external", query.target_sheet or "外部接続", detail)
        dependencies.append(
            GroundedClaim(
                text=f"外部データ接続: {query.name}",
                confidence=query.confidence,
                evidence_ids=[evidence_id],
            )
        )

    warnings_by_text: dict[str, GroundedClaim] = {}
    for risk in wb.analysis_risks:
        evidence_id = evidence.add("risk", risk.location, risk.evidence)
        existing = warnings_by_text.get(risk.description)
        if existing is not None:
            existing.evidence_ids.append(evidence_id)
            continue
        warnings_by_text[risk.description] = GroundedClaim(
            text=risk.description,
            confidence=risk.confidence,
            evidence_ids=[evidence_id],
        )

    return WorkbookDiagnosis(
        filename=wb.filename,
        headline=headline,
        overview=overview,
        features=features,
        inputs=input_claims,
        outputs=output_claims,
        external_dependencies=dependencies,
        warnings=list(warnings_by_text.values()),
        evidence=evidence.items,
        coverage=DiagnosisCoverage(
            sheets=len(wb.sheets),
            formulas=formula_count,
            vba_procedures=procedure_count,
            controls=len(controls),
            external_dependencies=len(dependencies),
        ),
        limitations=[
            "静的解析のため、実行時にだけ決まるVBA・数式・外部システムの挙動は断定できません。",
            "業務ルールや利用手順は、ファイル内に根拠がない場合は特定できません。",
            "Excel実行環境での再計算・マクロ実行が完了するまでは、動作保証ではありません。",
        ],
    )
