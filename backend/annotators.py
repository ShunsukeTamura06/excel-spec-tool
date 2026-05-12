"""LLM による Workbook 注釈付け (Phase C).

`annotate_workbook(wb, llm)` を呼ぶと、SheetInfo.purpose と VbaProcedure.annotation
を LLM 推論で埋めて返す。生成された Workbook を `/analyze` から保存し、設計書生成に渡す。

設計方針:
- 注釈ステップは大量呼び出し (シート N 件 + プロシージャ M 件) になるので、すべて
  `tier="fast"` で呼ぶ。チャットと違い caching の恩恵は薄い (毎回内容が違う) ため、
  fast モデルでコストとレイテンシを抑える。
- 1 シート / 1 プロシージャごとに独立した LLM 呼び出し。LLM が失敗しても全体は止めず、
  対象だけ注釈空のままにしてログに warning を残す。
- VBA コードが極端に長い場合は先頭 N 字に切り詰める (実 LLM のコンテキスト窓を超えない
  ための安全弁)。
- core 層は LLM を知らないため (CLAUDE.md §0)、本モジュールは backend 層に置く。
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

from core.models import (
    SheetInfo,
    VbaModule,
    VbaProcedure,
    Workbook,
)

if TYPE_CHECKING:
    from backend.llm_client import LLMClient

logger = logging.getLogger(__name__)


# LLM に渡す VBA コード本体の最大文字数 (これを超えたら冒頭のみ + truncated 注記).
# 社内 LLM のコンテキスト窓 (8k〜32k tokens 想定) を超えないための安全弁。
_VBA_CODE_MAX_CHARS = 8000

# シート用 system prompt
_SHEET_PURPOSE_PROMPT = (
    "あなたは Excel 設計書の作成者です。"
    "渡されたシートの構造情報 (名前 / 行数 / 列数 / 主要数式の一部 / 名前付き範囲 / "
    "冒頭プレビュー) から、このシートの『用途』を 1〜2 文 (日本語) で簡潔に述べてください。"
    "推測の場合は『〜と思われる』のように曖昧さを残してください。事実が乏しい場合は"
    "『用途不明』で構いません。"
)

# VBA プロシージャ用 system prompt
_VBA_PROC_PROMPT = (
    "あなたは Excel/VBA 設計書の作成者です。"
    "渡された VBA プロシージャのソースから、このプロシージャが『何をするか』を"
    "1 文 (日本語、80 字以内) で説明してください。"
    "事実に基づき、推測は避けてください。"
)


def _format_sheet_brief(sheet: SheetInfo) -> str:
    """シート用 LLM プロンプトの user content を組み立てる."""
    lines = [
        f"シート名: {sheet.name}",
        f"行数: {sheet.rows} / 列数: {sheet.cols}",
        f"数式セル数: {len(sheet.formulas)}",
    ]
    if sheet.named_ranges:
        nr_str = ", ".join(f"{nr.name}={nr.refers_to}" for nr in sheet.named_ranges[:5])
        lines.append(f"名前付き範囲 (一部): {nr_str}")
    if sheet.formulas:
        sample = sheet.formulas[:5]
        lines.append("主要数式 (一部):")
        for f in sample:
            lines.append(f"  {f.coord}: {f.formula}")
    if sheet.preview_rows:
        lines.append(f"冒頭プレビュー ({sheet.preview_origin}):")
        for row in sheet.preview_rows[:5]:
            cells = [str(c) if c is not None else "" for c in row[:10]]
            lines.append("  | " + " | ".join(cells) + " |")
    return "\n".join(lines)


def _format_procedure_brief(module: VbaModule, proc: VbaProcedure) -> str:
    """プロシージャ用 LLM プロンプトの user content."""
    code = proc.code
    if not code and module.code:
        # 行範囲で切り出し
        lines_all = module.code.splitlines()
        start = max(1, proc.start_line) - 1
        end = max(proc.end_line, proc.start_line)
        code = "\n".join(lines_all[start:end])

    truncated = False
    if len(code) > _VBA_CODE_MAX_CHARS:
        code = code[:_VBA_CODE_MAX_CHARS]
        truncated = True

    parts = [
        f"モジュール: {module.name} ({module.type})",
        f"プロシージャ: {proc.kind} {proc.name} (行 {proc.start_line}-{proc.end_line})",
        "ソースコード:",
        "```vba",
        code,
        "```",
    ]
    if truncated:
        parts.append("(注: コードが長いため冒頭のみを抜粋しています)")
    return "\n".join(parts)


def _safe_annotate(
    llm: LLMClient,
    prompt: str,
    content: str,
    description: str,
) -> str:
    """LLM 注釈呼び出し. 失敗は warning ログを残して空文字を返す."""
    try:
        result = llm.annotate_text(prompt, content)
    except Exception:  # noqa: BLE001 - 個別失敗は全体を止めない
        logger.exception("annotate failed for %s", description)
        return ""
    return (result or "").strip()


def annotate_workbook(wb: Workbook, llm: LLMClient) -> Workbook:
    """Workbook の各シート用途と VBA プロシージャに LLM 注釈を付与する.

    既存の `purpose` / `annotation` が空のものだけ埋める (べき等)。
    入力 wb は変更せず、注釈済みコピーを返す。

    Args:
        wb: 抽出済み Workbook
        llm: LLMClient. annotate_text(prompt, content, tier="fast") を呼ぶ。

    Returns:
        注釈付き Workbook (新しいインスタンス)
    """
    annotated_sheets: list[SheetInfo] = []
    for sheet in wb.sheets:
        if sheet.purpose:
            annotated_sheets.append(sheet)
            continue
        brief = _format_sheet_brief(sheet)
        purpose = _safe_annotate(llm, _SHEET_PURPOSE_PROMPT, brief, f"sheet '{sheet.name}'")
        if purpose:
            logger.info("annotated sheet '%s': %s", sheet.name, purpose[:60])
        annotated_sheets.append(sheet.model_copy(update={"purpose": purpose}))

    annotated_modules: list[VbaModule] = []
    for module in wb.vba_modules:
        annotated_procs: list[VbaProcedure] = []
        for proc in module.procedures:
            if proc.annotation:
                annotated_procs.append(proc)
                continue
            brief = _format_procedure_brief(module, proc)
            annot = _safe_annotate(
                llm,
                _VBA_PROC_PROMPT,
                brief,
                f"procedure '{module.name}.{proc.name}'",
            )
            if annot:
                logger.info(
                    "annotated procedure '%s.%s': %s",
                    module.name,
                    proc.name,
                    annot[:60],
                )
            annotated_procs.append(proc.model_copy(update={"annotation": annot}))
        annotated_modules.append(module.model_copy(update={"procedures": annotated_procs}))

    return wb.model_copy(update={"sheets": annotated_sheets, "vba_modules": annotated_modules})
