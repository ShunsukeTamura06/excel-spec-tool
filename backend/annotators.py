"""LLM による Workbook 注釈付け (Phase C + P1-2 構造化).

`annotate_workbook(wb, llm)` を呼ぶと、SheetInfo / VbaProcedure の以下フィールドを
LLM 推論で埋めて返す:

  SheetInfo:
    - purpose            (1〜2 文)
    - inputs             (依存元: 他シート / 外部)
    - outputs            (出力先)
    - main_calculations  (主要計算の自然言語説明)
    - usage_scenario     (想定利用シーン)

  VbaProcedure:
    - annotation    (1 文サマリ)
    - side_effects  (書き込み先)
    - triggers      (想定起動契機)
    - calls         (内部で呼ぶ他プロシージャ名)

設計方針:
- 大量呼び出しを想定し全て `tier="fast"`. キャッシュ恩恵が薄いタスク.
- 1 シート / 1 プロシージャあたり LLM 呼び出しは 1 回. JSON で複数フィールドを
  まとめて返してもらうことで呼び出し数を最小化.
- LLM 応答は不安定なので、JSON パースを段階的に試行 (生 JSON → ``` フェンス除去 →
  失敗時は空注釈). 失敗しても全体は止めず、対象だけ空のままにする.
- VBA コードが極端に長い場合は先頭 N 字に切り詰める (コンテキスト窓の保護).
- core 層は LLM を知らないため (CLAUDE.md §0)、本モジュールは backend 層に置く.
"""

from __future__ import annotations

import json
import logging
import re
from typing import TYPE_CHECKING, Any

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


# ====================================================================
# プロンプト
# ====================================================================

_SHEET_ANNOTATION_PROMPT = (
    "あなたは Excel 設計書の作成者です。"
    "渡されたシートの構造情報を読み、以下の JSON 形式で回答してください。"
    "推測の場合は値を曖昧に (例: '〜と思われる') してください。"
    "事実が乏しい項目は空文字列 / 空配列で構いません。"
    "JSON 以外の前後文・コードフェンスは出力しないでください。\n"
    "\n"
    "{\n"
    '  "purpose": "シート用途を 1〜2 文",\n'
    '  "inputs": ["依存元 (他シート名 or 外部) を 0〜5 件"],\n'
    '  "outputs": ["出力先 (他シート / 帳票) を 0〜5 件"],\n'
    '  "main_calculations": ["主要計算の自然言語説明を 0〜5 件"],\n'
    '  "usage_scenario": "想定利用シーン (誰がいつ何のために使うか)"\n'
    "}\n"
)

_VBA_PROC_ANNOTATION_PROMPT = (
    "あなたは Excel / VBA 設計書の作成者です。"
    "渡された VBA プロシージャを読み、以下の JSON 形式で回答してください。"
    "事実に基づき、推測は避けてください。"
    "JSON 以外の前後文・コードフェンスは出力しないでください。\n"
    "\n"
    "{\n"
    '  "annotation": "プロシージャの目的を 1 文 (80 字以内)",\n'
    '  "side_effects": ["書き込み先セル/シート/外部 を 0〜5 件"],\n'
    '  "triggers": ["想定起動契機 (ボタン / イベント / 手動 など) を 0〜3 件"],\n'
    '  "calls": ["コード中で呼んでいる他プロシージャ名を 0〜10 件"]\n'
    "}\n"
)


# ====================================================================
# プロンプトの user content 構築
# ====================================================================

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


# ====================================================================
# JSON パース (寛容に)
# ====================================================================

# ```json ... ``` フェンスや前後の余計なテキストを除去するための regex.
# 非貪欲で最初の { ... } ブロックを掴む.
_JSON_OBJECT_RE = re.compile(r"\{[\s\S]*\}")


def _parse_llm_json(text: str) -> dict[str, Any] | None:
    """LLM 応答から JSON オブジェクトを取り出す. 失敗時 None.

    試行順:
      1. テキストそのまま json.loads
      2. ```json ... ``` フェンスを取り除いて json.loads
      3. テキスト全体から最初の `{...}` ブロックを正規表現で抽出して json.loads

    どれも失敗すれば None.
    """
    if not text:
        return None

    candidates: list[str] = []
    stripped = text.strip()
    candidates.append(stripped)

    # ``` で始まるフェンスを除去
    fenced = re.sub(r"^```[a-zA-Z]*\s*", "", stripped)
    fenced = re.sub(r"\s*```\s*$", "", fenced)
    if fenced != stripped:
        candidates.append(fenced)

    # 最初の { ... } を抽出
    m = _JSON_OBJECT_RE.search(stripped)
    if m:
        candidates.append(m.group(0))

    for cand in candidates:
        try:
            parsed = json.loads(cand)
        except (ValueError, TypeError):
            continue
        if isinstance(parsed, dict):
            return parsed
    return None


def _str_field(payload: dict[str, Any], key: str) -> str:
    """payload[key] を str として安全に取り出す. 欠落 / 型不一致なら空文字."""
    v = payload.get(key)
    return v.strip() if isinstance(v, str) else ""


def _list_str_field(payload: dict[str, Any], key: str, limit: int = 20) -> list[str]:
    """payload[key] を list[str] として取り出す. 型不一致は無視, 上限件数で切る."""
    v = payload.get(key)
    if not isinstance(v, list):
        return []
    out: list[str] = []
    for item in v[:limit]:
        if isinstance(item, str):
            s = item.strip()
            if s:
                out.append(s)
    return out


# ====================================================================
# LLM 呼び出し
# ====================================================================

def _safe_annotate_json(
    llm: LLMClient,
    prompt: str,
    content: str,
    description: str,
) -> dict[str, Any]:
    """JSON 注釈呼び出し. 失敗 / パース不能は warning ログを残して {} を返す.

    呼び出し側は得られた dict から各フィールドを取り出す。dict が空の場合は
    そのフィールドは未注釈のままになる (= 既存値を維持)。
    """
    try:
        raw = llm.annotate_text(prompt, content)
    except Exception:  # noqa: BLE001 - 個別失敗は全体を止めない
        logger.exception("annotate failed for %s", description)
        return {}

    parsed = _parse_llm_json(raw or "")
    if parsed is None:
        logger.warning(
            "annotate JSON parse failed for %s; raw=%r",
            description,
            (raw or "")[:200],
        )
        return {}
    return parsed


# ====================================================================
# 公開関数
# ====================================================================

def annotate_workbook(wb: Workbook, llm: LLMClient) -> Workbook:
    """Workbook の各シート / プロシージャに構造化 LLM 注釈を付与する.

    既存の `purpose` / `annotation` が埋まっているものはスキップ (べき等)。
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
            # 既に注釈済み → そのまま (べき等性)
            annotated_sheets.append(sheet)
            continue
        brief = _format_sheet_brief(sheet)
        payload = _safe_annotate_json(
            llm, _SHEET_ANNOTATION_PROMPT, brief, f"sheet '{sheet.name}'"
        )
        purpose = _str_field(payload, "purpose")
        inputs = _list_str_field(payload, "inputs")
        outputs = _list_str_field(payload, "outputs")
        main_calcs = _list_str_field(payload, "main_calculations")
        scenario = _str_field(payload, "usage_scenario")
        if purpose or inputs or outputs or main_calcs or scenario:
            logger.info(
                "annotated sheet '%s': purpose=%r inputs=%d outputs=%d calcs=%d",
                sheet.name,
                purpose[:60],
                len(inputs),
                len(outputs),
                len(main_calcs),
            )
        annotated_sheets.append(
            sheet.model_copy(
                update={
                    "purpose": purpose,
                    "inputs": inputs,
                    "outputs": outputs,
                    "main_calculations": main_calcs,
                    "usage_scenario": scenario,
                },
            )
        )

    annotated_modules: list[VbaModule] = []
    for module in wb.vba_modules:
        annotated_procs: list[VbaProcedure] = []
        for proc in module.procedures:
            if proc.annotation:
                annotated_procs.append(proc)
                continue
            brief = _format_procedure_brief(module, proc)
            payload = _safe_annotate_json(
                llm,
                _VBA_PROC_ANNOTATION_PROMPT,
                brief,
                f"procedure '{module.name}.{proc.name}'",
            )
            annot = _str_field(payload, "annotation")
            side_effects = _list_str_field(payload, "side_effects")
            triggers = _list_str_field(payload, "triggers")
            calls = _list_str_field(payload, "calls")
            if annot or side_effects or triggers or calls:
                logger.info(
                    "annotated proc '%s.%s': %r side=%d trig=%d calls=%d",
                    module.name,
                    proc.name,
                    annot[:60],
                    len(side_effects),
                    len(triggers),
                    len(calls),
                )
            annotated_procs.append(
                proc.model_copy(
                    update={
                        "annotation": annot,
                        "side_effects": side_effects,
                        "triggers": triggers,
                        "calls": calls,
                    },
                )
            )
        annotated_modules.append(module.model_copy(update={"procedures": annotated_procs}))

    return wb.model_copy(update={"sheets": annotated_sheets, "vba_modules": annotated_modules})
