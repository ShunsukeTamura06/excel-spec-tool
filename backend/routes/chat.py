"""POST /chat/{job_id} と GET /chat/{job_id}/history — 改修対話.

LLM の function calling ループを実装している:
1. system prompt + 履歴 + ユーザー発話 + tools 定義を LLM に渡す
2. 応答に tool_calls があれば、各 tool を実行して結果を tool role の
   メッセージとして追加し、再度 LLM を呼ぶ
3. tool_calls が無くなったら最終応答テキストを返す
4. 暴走防止のため最大反復回数を制限する

並行性メモ:
  LLM クライアントは同期 HTTP. tool ループ全体で数秒〜数分かかるため、
  `async def` 本体で直接走らせると event loop がブロックされ他リクエストも
  巻き添えで詰まる. `_run_tool_loop` 全体を `asyncio.to_thread` で threadpool
  に逃がして event loop を解放する.
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
from collections.abc import AsyncIterator, Callable
from datetime import datetime, timezone
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from backend.dependencies import get_llm_client, get_storage
from backend.llm_client import LLMClient, LLMResponse, Usage
from backend.llm_tools import build_tool_definitions, execute_tool_call
from backend.storage import JobNotFoundError, Storage
from core.models import ChatMessage

router = APIRouter()
logger = logging.getLogger(__name__)


# tool ループの最大反復回数. 正確性を優先し、複数観点の確認を許容する。
MAX_TOOL_ITERATIONS = 16

# LLM に送る履歴の上限ペア数 (Phase B-2). user/assistant の 1 往復 = 1 ペア.
# これを超えた古いペアは「過去のやりとり概要」として 1 つの system message に集約する。
# 履歴ファイル (chat_history.jsonl) 自体は完全保存される。LLM 文脈窓だけ絞る。
_DEFAULT_HISTORY_PAIRS = 10
# 古い履歴を要約する際の、1 メッセージあたりの最大文字数 (それ以上は切り詰める)
_SUMMARY_MESSAGE_MAX_CHARS = 120
_TOOL_LOOP_FALLBACK_TEXTS = (
    "[tool loop max iterations reached; partial result]",
    "ツール確認が上限に達し、最終回答を生成できませんでした。",
    "確認できた情報だけでは、この質問に正確に回答できませんでした。",
)
ProgressEmitter = Callable[[str, dict[str, Any]], None]

_TOOL_PROGRESS_MESSAGES = {
    "get_cells_range": "セル範囲を確認中",
    "find_cells": "セルを検索中",
    "lookup_references": "参照関係を確認中",
    "list_vba_modules": "VBAモジュール一覧を確認中",
    "get_vba_procedure": "VBAコードを確認中",
    "list_sheet_formulas": "シートの数式を確認中",
    "list_workbook_objects": "Excelオブジェクトを確認中",
    "list_analysis_risks": "未解析リスクを確認中",
    "lookup_external_function": "外部関数の定義を確認中",
    "list_external_functions_used": "外部関数の使用箇所を確認中",
}


def _get_history_pairs_limit() -> int:
    """環境変数 CHAT_HISTORY_LIMIT_PAIRS で上書き可能. 0 以下なら制限なし扱い."""
    raw = os.environ.get("CHAT_HISTORY_LIMIT_PAIRS")
    if not raw:
        return _DEFAULT_HISTORY_PAIRS
    try:
        return int(raw)
    except ValueError:
        logger.warning(
            "invalid CHAT_HISTORY_LIMIT_PAIRS=%r; using default %d",
            raw,
            _DEFAULT_HISTORY_PAIRS,
        )
        return _DEFAULT_HISTORY_PAIRS


def _summarize_old_turns(old: list[ChatMessage]) -> str:
    """古いメッセージ列を 1 つの system テキストに圧縮する.

    LLM を呼ばずに機械的に圧縮する: 各メッセージを冒頭 N 文字に切り詰めて箇条書きにする。
    情報ロスはあるが、コスト 0 / 遅延 0 で履歴の大筋を残せる。
    """
    if not old:
        return ""
    lines = [f"# 過去のやりとり概要 ({len(old)} メッセージ, 古い順)"]
    for m in old:
        content = (m.content or "").strip().replace("\n", " ")
        if len(content) > _SUMMARY_MESSAGE_MAX_CHARS:
            content = content[: _SUMMARY_MESSAGE_MAX_CHARS - 1] + "…"
        lines.append(f"- {m.role}: {content}")
    return "\n".join(lines)


def _trim_history(
    history: list[ChatMessage],
    max_pairs: int,
) -> tuple[list[ChatMessage], str | None]:
    """履歴を直近 `max_pairs` ペア分に絞り込む.

    Args:
        history: 全履歴 (jsonl から読んだもの).
        max_pairs: 保持する直近ペア数. 0 以下なら全件保持 (トリムしない).

    Returns:
        (recent_messages, summary_text_or_none)
        トリムが発生しなかった場合 summary は None.
    """
    if max_pairs <= 0:
        return list(history), None
    max_msgs = max_pairs * 2
    if len(history) <= max_msgs:
        return list(history), None
    old = history[:-max_msgs]
    recent = history[-max_msgs:]
    return recent, _summarize_old_turns(old)


def _is_tool_loop_fallback_message(message: ChatMessage) -> bool:
    """ツールループ上限時の内部フォールバック応答かどうかを返す."""
    if message.role != "assistant":
        return False
    return any(text in message.content for text in _TOOL_LOOP_FALLBACK_TEXTS)


def _history_for_llm(history: list[ChatMessage]) -> list[ChatMessage]:
    """LLM に渡す履歴から、以後の会話を汚す内部エラー応答を除外する."""
    return [m for m in history if not _is_tool_loop_fallback_message(m)]


def _emit_progress(
    emit_progress: ProgressEmitter | None,
    event: str,
    data: dict[str, Any],
) -> None:
    """進捗イベントを送信する.

    Args:
        emit_progress: 進捗イベントの送信先. None の場合は何もしない。
        event: SSE のイベント名。
        data: JSON 化するイベント本文。
    """
    if emit_progress is not None:
        emit_progress(event, data)


def _tool_progress_message(tool_name: str) -> str:
    """ツール名をユーザーに見せる短い進捗文言へ変換する.

    Args:
        tool_name: LLM が要求したツール名。

    Returns:
        ユーザー向けの進捗メッセージ。
    """
    return _TOOL_PROGRESS_MESSAGES.get(tool_name, "設計書情報を確認中")


def _tool_signature(name: str, arguments: dict[str, Any]) -> str:
    """同一ツール呼び出しを判定するための安定したキーを返す.

    Args:
        name: ツール名。
        arguments: ツール引数。

    Returns:
        ツール名と正規化済み引数を含む文字列キー。
    """
    return f"{name}:{json.dumps(arguments, ensure_ascii=False, sort_keys=True)}"


def _sse_event(event: str, data: dict[str, Any]) -> str:
    """Server-Sent Events 形式の文字列へ変換する.

    Args:
        event: イベント名。
        data: JSON 化するイベント本文。

    Returns:
        SSE 1 イベント分の文字列。
    """
    payload = json.dumps(data, ensure_ascii=False)
    return f"event: {event}\ndata: {payload}\n\n"


def _build_unresolved_reply(tool_trace: list[dict[str, Any]]) -> str:
    """正確な回答を生成できなかった場合の deterministic な応答を作る.

    Args:
        tool_trace: 実行済みツールの記録。

    Returns:
        ユーザーに返す明確な限界説明と代替案。
    """
    lines = [
        "確認できた情報だけでは、この質問に正確に回答できませんでした。",
        "",
        "確認済みの内容:",
    ]
    if tool_trace:
        for item in tool_trace[-5:]:
            name = str(item.get("name", "unknown"))
            args = json.dumps(item.get("arguments", {}), ensure_ascii=False, sort_keys=True)
            lines.append(f"- {name} {args}")
    else:
        lines.append("- 参照できる追加情報を取得できませんでした。")
    lines.extend(
        [
            "",
            "代替案:",
            "- 対象のシート名、セル範囲、VBAモジュール名、または知りたい観点を"
            "もう少し具体化してください。",
            "- 変更前後の期待動作や、影響を確認したい項目を指定してください。",
            "- 必要であれば設計書タブや参照検索で対象箇所を先に絞り込んでから質問してください。",
        ]
    )
    return "\n".join(lines)


class ChatRequest(BaseModel):
    """POST /chat/{job_id} のリクエストボディ."""

    message: str


class ChatSessionCreateRequest(BaseModel):
    """POST /chat/{job_id}/sessions のリクエストボディ."""

    title: str = "新しい相談"


class ChatSessionUpdateRequest(BaseModel):
    """PATCH /chat/{job_id}/sessions/{session_id} のリクエストボディ."""

    title: str | None = None
    archived: bool | None = None


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


_SYSTEM_INSTRUCTIONS = "\n".join(
    [
        "あなたは Excelツール改修支援AI です。",
        "ユーザーの改修要望に対し、設計書とツールを参照して具体的な手順と影響範囲を回答してください。",
        "",
        "# 基本姿勢 (必ず守ること)",
        "",
        "## 1. 事実に基づいて回答する",
        "- 設計書とツールから実際に取得した情報「のみ」を根拠にする",
        "- 推測・憶測・一般論で回答しない",
        "- 回答中には「設計書の◯◯にこう書いてある」"
        "「ツール get_cells_range で F1 を確認した結果『実現損益』」のように"
        "必ず根拠を明示する",
        "- セル番地・行・列・シート名・項目名は必ずツールで実物を確認してから書く",
        "",
        "## 2. 確証がなければ確認してから答える",
        "- 「F 列に実現損益があるはず」と思っても、"
        "まず get_cells_range か find_cells で実値を確認する",
        "- 設計書の preview に映っていない領域 (例: 50 行目以降、20 列目以降) の値は"
        "決して推測しない。必ずツールで取得する",
        "- 波及範囲を述べる前に lookup_references を呼んで実際の参照元を確認する",
        "- 改修可否や安全性を述べる前に list_analysis_risks を呼び、未解析リスクを確認する",
        "- lookup_references の結果が 0 件でも、動的 VBA 参照まで含めて影響がないとは断定しない",
        "",
        "## 3. 分からないことは素直に認める",
        "- 設計書とツールでも判定できない場合は"
        "「該当情報が見つかりません」「設計書からは判断できません」と素直に伝える",
        "- 「たぶん」「おそらく」「だと思います」で曖昧に断定しない",
        "- VBA の動的挙動・実行時条件・外部システム連携など、"
        "静的解析では分からないことは「実行時の動作は確認できません」と明言する",
        "- 推測で穴埋めするくらいなら、ユーザーに聞き返す or 「分かりません」と答える",
        "",
        "## 4. 不明確な要望は遠慮なく聞き返す",
        "- 「あの表」「この項目」「いつもの集計」など曖昧な参照は、"
        "どのシート/行/列/名前のことか必ず聞き返す",
        "- 「変更したい」「直したい」だけでは何をどう変えるか (追加/修正/削除/移動) を聞く",
        "- 複数解釈が可能な質問は、回答する前に意図を確認する",
        "- 質問を遠慮しない。1往復で解決しなくてよい。曖昧なまま回答するより聞き返すのが正しい",
        "",
        "## 5. 回答漏れを防ぐ",
        "- ユーザーの質問に複数の論点が含まれていたら、すべてに答える",
        "- 波及範囲を聞かれたら参照元を網羅する (件数 + 主要箇所)",
        "- 参照解析で検出できない動的参照の可能性が残る場合は、"
        "「静的解析で確認できた範囲」と「未確認の可能性」を分けて書く",
        "- 未解析リスクがある場合は、回答の最後ではなく波及範囲の直後に明示する",
        "- 不明な部分があってもそれを「不明」として明示すれば回答漏れにはならない",
        "",
        "# 参照解析の前提",
        "",
        "- 数式参照は openpyxl の Tokenizer で抽出した RANGE トークンを使う",
        '- VBA 参照は静的に確定できる `Range("A1")`, '
        '`Worksheets("Sheet").Range("A1")`, `Sheets("Sheet").Cells(2, 8)`, '
        '`[Sheet!A1]`, 単純なシート変数 (`Set ws = Worksheets("Sheet")`)、'
        '`With Sheets("Sheet")` 内の `.Range` / `.Cells` を対象にする',
        "- コメントや文字列リテラル内の `Range(...)` は参照として扱わない",
        '- `Range("A" & row)`, `Range(addr)`, `Worksheets(sheetName)`, '
        "`Cells(row, col)`, `ActiveSheet`, `Selection`, `CurrentRegion`, "
        "`UsedRange`, `Offset`, `Resize`, `Evaluate`, `Application.Run` など、"
        "実行時に決まる参照は検出対象外",
        "- 参照検索で見つからない場合は「静的解析では検出されない」と表現し、"
        "影響が完全に無いとは断定しない",
        "- VBA の波及範囲が重要な場合は、lookup_references だけで終えず、"
        "list_vba_modules と get_vba_procedure で該当コードを確認する",
        "- グラフ、ピボット、Power Query / 外部接続への波及があり得る場合は、"
        "list_workbook_objects で棚卸しを確認する",
        "- 未解析リスクは list_analysis_risks で確認し、手動確認対象として扱う",
        "",
        "# 応答フォーマット",
        "",
        "通常の回答には以下のセクションを必ず含めてください:",
        "",
        "1. **確認できた事実**: ツールや設計書から直接確認できた情報と根拠",
        "2. **波及範囲**: 影響を受けるセル / VBA / グラフ / ピボット / 接続を一覧 "
        "(件数 + 主要箇所、調査根拠としてどのツールを呼んだか)",
        "3. **未解析リスク**: list_analysis_risks の結果や、静的解析で断定できない点",
        "4. **改修手順**: ユーザーの操作レベル (どのセル/タブを開き、何を入力するか) を具体的に",
        "5. **手動確認チェックリスト**: 改修後に Excel 上で確認すべき"
        "操作・再計算・更新・ボタン押下",
        "",
        "禁止表現:",
        "- 「影響ありません」「使われていません」「安全です」と断定しない",
        "- 代わりに「静的解析で確認できた範囲では」「未解析リスクとして」を使う",
        "",
        "ただし、まずユーザーに質問する必要がある場合は、"
        "フォーマットに従わず質問だけを返してください。",
        "情報不足で回答できないときは、何が分からないかを明示して、"
        "ユーザーに追加情報を求めてください。",
        "",
        "# 使えるツール",
        "",
        "設計書には Excel の概観 "
        "(シート一覧 / 先頭 20 行 × 20 列のプレビュー / 表構造 / "
        "VBA モジュール・プロシージャの目録 / 名前付き範囲 / 主要数式 TOP10) のみ載っています。",
        "VBA のソースコード本体や TOP10 から漏れた数式は設計書に含まれません。"
        "詳細な値・行内容・参照関係・VBA コードを確認したい時は、以下のツールを必ず呼んでください:",
        "",
        "- `get_cells_range(sheet, range)`:",
        '    例: get_cells_range("Portfolio", "A6:Z6")',
        "    指定範囲のセル値・数式・データ型を 2D 配列で返す。プレビュー外の領域を見たい時はこれ",
        "- `find_cells(query, sheet?, limit?)`:",
        '    例: find_cells("実現損益", sheet="Portfolio")',
        "    値の部分一致でセルを検索する。ユーザーが言った項目名がどこにあるか分からない時はこれ",
        "- `lookup_references(target)`:",
        '    例: lookup_references("Calc!H2")',
        "    あるセル/範囲を参照している箇所 (数式・VBA) を返す。波及範囲調査ではこれを必ず使う。",
        "    完全一致だけでなく範囲交差でヒットする (`Input!A5` で `Input!A:A` 参照もヒット)。",
        "    target はシート修飾付きで指定する (例: 'Calc!H2', 'Input!A:A')。",
        "    ただし VBA は静的に確定できる参照のみ。0 件でも動的参照の可能性は残る。",
        "- `list_vba_modules()`:",
        "    VBA モジュールとプロシージャの一覧 (名前・種別・行範囲) を軽量に返す。",
        "    どのプロシージャが該当しそうかを最初に絞り込むときに使う。",
        "- `get_vba_procedure(module, name)`:",
        '    例: get_vba_procedure("Module1", "UpdateDaily")',
        "    指定したプロシージャのソースコード本体を返す。",
        "    設計書には VBA 全コードを載せていないので、ロジックを確認したい時はこれを呼ぶ。",
        "- `list_sheet_formulas(sheet, pattern?, limit?)`:",
        '    例: list_sheet_formulas("Calc", pattern="SUMIF")',
        "    シートの数式一覧 (設計書 TOP10 から漏れた数式や特定関数の検索) を返す。",
        "- `list_workbook_objects(sheet?, kind?)`:",
        '    例: list_workbook_objects(kind="pivot")',
        "    グラフ / ピボット / Power Query・外部接続の棚卸しを返す。",
        "    表や列の変更が可視化・集計・外部接続に波及するか確認するときに使う。",
        "- `list_analysis_risks(severity?, category?, limit?)`:",
        '    例: list_analysis_risks(severity="high")',
        "    動的 VBA 参照、イベント処理、INDIRECT/OFFSET、外部接続など、"
        "静的解析では断定できない未解析リスクを返す。",
        "- `lookup_external_function(name)`:",
        '    例: lookup_external_function("BDH")',
        "    Bloomberg / Refinitiv 等の Excel Add-In 関数の定義 "
        "(引数 / 返り値 / 例 / 落とし穴) を返す。",
        "    BDH / BDP / BDS のような非標準関数の挙動を答える前に必ず呼ぶこと "
        "(推測やハルシネーションを避けるため)。",
        "- `list_external_functions_used()`:",
        "    このブックで使われている外部 Add-In 関数の一覧 (回数 / 主な使用箇所) を返す。",
        "    「このブックは Bloomberg をどこで使っている？」のような問いの起点に使う。",
        "",
        "確実性を優先して必要な範囲でツールを呼んでください。",
        "同じ観点の確認を繰り返さず、十分な根拠が集まったら回答をまとめてください。",
    ]
)


def _build_system_prompt(spec_md: str) -> str:
    """設計書を system prompt に固定する.

    SPEC §5.4 の「応答には改修手順 + 波及範囲を含める」指示に加え、
    LLM の振る舞いガイドライン (事実主義 / 分からない時は認める / 質問してよい /
    回答漏れを防ぐ) を明文化する。
    """
    if spec_md:
        return f"{_SYSTEM_INSTRUCTIONS}\n\n---\n# 設計書\n{spec_md}"
    return _SYSTEM_INSTRUCTIONS


def _tool_call_to_assistant_message(resp: LLMResponse) -> dict[str, Any]:
    """LLMResponse の tool_calls を OpenAI 仕様の assistant メッセージに変換."""
    return {
        "role": "assistant",
        "content": resp.content or "",
        "tool_calls": [
            {
                "id": tc.id,
                "type": "function",
                "function": {
                    "name": tc.name,
                    "arguments": json.dumps(tc.arguments, ensure_ascii=False),
                },
            }
            for tc in resp.tool_calls
        ],
    }


def _run_tool_loop(
    llm: LLMClient,
    storage: Storage,
    job_id: str,
    base_messages: list[dict[str, Any]],
    emit_progress: ProgressEmitter | None = None,
) -> tuple[str, list[dict[str, Any]]]:
    """tool 呼び出しを伴うチャット応答ループ.

    Args:
        llm: LLM クライアント。
        storage: ジョブ保存先。
        job_id: 対象ジョブ ID。
        base_messages: system prompt・履歴・ユーザー発話。
        emit_progress: 進捗イベント送信コールバック。

    Returns:
        (最終アシスタント応答テキスト, ループ中の中間 tool 呼び出し記録).
        中間記録は履歴 jsonl には保存しないがデバッグ・観測用に返す。
    """
    tools = build_tool_definitions()
    messages = list(base_messages)
    tool_trace: list[dict[str, Any]] = []
    total_usage = Usage()
    tool_result_cache: dict[str, str] = {}
    repeat_counts: dict[str, int] = {}

    for iteration in range(MAX_TOOL_ITERATIONS):
        _emit_progress(
            emit_progress,
            "status",
            {
                "message": (
                    "設計書と質問を照合中" if iteration == 0 else "追加で必要な根拠を確認中"
                ),
                "iteration": iteration + 1,
            },
        )
        # チャット中は tier 固定 (= pro). 理由:
        # ツールループ内で pro/fast を切り替えると、prompt caching が
        # モデル単位で分離されているため、切替時点でキャッシュが失効してしまう。
        # 入力 ≫ 出力 のチャットでは、caching を活かす方がトータル安い。
        # fast tier は LLM 注釈バッチ (Phase C) など、キャッシュ恩恵が薄い大量呼び出しで使う。
        #
        # cache_prefix=1: 先頭 1 メッセージ (system: 行動指針 + 設計書) を
        # キャッシュ対象として LLM に伝える (Phase B-3)。
        # 履歴トリムの summary は 2 つ目以降に置いてあるので、トリムが起きてもプレフィックスは
        # 安定し、キャッシュは効き続ける。
        logger.info("llm iteration=%d tier=pro messages=%d", iteration + 1, len(messages))
        resp = llm.chat_completion_with_tools(messages, tools=tools, tier="pro", cache_prefix=1)

        if resp.usage is not None:
            total_usage = total_usage + resp.usage
            logger.info(
                "llm usage iter=%d model=%s prompt=%d completion=%d total=%d cached=%d",
                iteration + 1,
                resp.model or "?",
                resp.usage.prompt_tokens,
                resp.usage.completion_tokens,
                resp.usage.total_tokens,
                resp.usage.cached_tokens,
            )

        if not resp.tool_calls:
            logger.info(
                "llm final response: iterations=%d tool_calls_total=%d reply_chars=%d "
                "cumulative_prompt=%d completion=%d total=%d cached=%d",
                iteration + 1,
                len(tool_trace),
                len(resp.content or ""),
                total_usage.prompt_tokens,
                total_usage.completion_tokens,
                total_usage.total_tokens,
                total_usage.cached_tokens,
            )
            return (resp.content or "", tool_trace)

        # tool 呼び出しを実行し、結果を tool role メッセージで追加
        messages.append(_tool_call_to_assistant_message(resp))
        for tc in resp.tool_calls:
            logger.info("tool call: name=%s args=%s", tc.name, tc.arguments)
            signature = _tool_signature(tc.name, tc.arguments)
            repeat_counts[signature] = repeat_counts.get(signature, 0) + 1
            progress_message = _tool_progress_message(tc.name)
            _emit_progress(
                emit_progress,
                "tool_start",
                {
                    "message": progress_message,
                    "tool_name": tc.name,
                    "iteration": iteration + 1,
                },
            )
            if signature in tool_result_cache:
                result_str = tool_result_cache[signature]
                if repeat_counts[signature] >= 3:
                    result_str = (
                        f"{result_str}\n\n"
                        "[system note] 同じ条件で確認済みです。新しい根拠が必要な場合は"
                        "別の観点でツールを使ってください。正確に判断できる根拠が不足する場合は、"
                        "不足している情報と代替案を明示してください。"
                    )
            else:
                result_str = execute_tool_call(storage, job_id, tc.name, tc.arguments)
                tool_result_cache[signature] = result_str
            logger.info(
                "tool result: name=%s result_chars=%d preview=%s",
                tc.name,
                len(result_str),
                result_str[:120].replace("\n", " "),
            )
            tool_trace.append(
                {"name": tc.name, "arguments": tc.arguments, "result_preview": result_str[:200]}
            )
            _emit_progress(
                emit_progress,
                "tool_result",
                {
                    "message": f"{progress_message}が完了しました",
                    "tool_name": tc.name,
                    "iteration": iteration + 1,
                },
            )
            messages.append(
                {
                    "role": "tool",
                    "tool_call_id": tc.id,
                    "name": tc.name,
                    "content": result_str,
                }
            )

    # 確認継続後も最終回答に至らない場合:
    # 曖昧に推測せず、正確に判断できるかどうかを LLM に判定させる。
    logger.warning(
        "tool loop hit MAX_TOOL_ITERATIONS=%d tool_calls_total=%d "
        "cumulative_prompt=%d completion=%d total=%d cached=%d",
        MAX_TOOL_ITERATIONS,
        len(tool_trace),
        total_usage.prompt_tokens,
        total_usage.completion_tokens,
        total_usage.total_tokens,
        total_usage.cached_tokens,
    )
    _emit_progress(
        emit_progress,
        "status",
        {"message": "確認済み情報で回答可否を判断中", "iteration": MAX_TOOL_ITERATIONS},
    )
    messages.append(
        {
            "role": "system",
            "content": (
                "追加のツール確認は行わず、ここまでに得たツール結果と設計書だけを根拠に"
                "最終回答を日本語で作成してください。正確に回答できる根拠がある場合だけ"
                "断定してください。根拠が不足して正確に判断できない場合は、できない理由、"
                "確認済みの内容、ユーザーが次に取れる代替案を明確に述べてください。"
            ),
        }
    )
    reply = llm.chat_completion(
        messages,
        tier="pro",
        cache_prefix=1,
    )
    logger.info("llm forced final response: reply_chars=%d", len(reply or ""))
    reply = (reply or resp.content or "").strip()
    if not reply:
        reply = _build_unresolved_reply(tool_trace)
    return (reply, tool_trace)


def _build_chat_messages(
    storage: Storage, job_id: str, session_id: str, user_message: str
) -> list[dict[str, Any]]:
    """LLM に渡すチャットメッセージ列を構築する.

    Args:
        storage: ジョブ保存先。
        job_id: 対象ジョブ ID。
        session_id: チャットセッション ID。
        user_message: 今回のユーザー発話。

    Returns:
        system prompt、履歴、今回発話を含む LLM 入力。
    """
    try:
        try:
            spec_md = storage.load_spec(job_id)
        except FileNotFoundError:
            spec_md = ""

        history = storage.load_chat_history(job_id, session_id=session_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job or chat session not found") from e

    # 履歴を直近 N ペアに絞り、古い分は summary に集約 (Phase B-2)
    # ツールループ上限時のフォールバック文は、次ターン以降の LLM 文脈から除外する。
    prompt_history = _history_for_llm(history)
    max_pairs = _get_history_pairs_limit()
    recent, summary = _trim_history(prompt_history, max_pairs=max_pairs)
    if summary:
        logger.info(
            "history trimmed: total=%d kept=%d summarized=%d",
            len(prompt_history),
            len(recent),
            len(prompt_history) - len(recent),
        )

    messages: list[dict[str, Any]] = [
        {"role": "system", "content": _build_system_prompt(spec_md)},
    ]
    if summary:
        # system 直後に置く. 設計書 (大プレフィックス) のキャッシュ性を壊さないよう、
        # main system prompt とは別メッセージにする。
        messages.append({"role": "system", "content": summary})
    messages.extend({"role": m.role, "content": m.content} for m in recent)
    messages.append({"role": "user", "content": user_message})
    return messages


def _persist_chat_turn(
    storage: Storage,
    job_id: str,
    session_id: str,
    user_message: str,
    reply: str,
) -> list[ChatMessage]:
    """ユーザー発話とアシスタント応答を履歴に保存して再読込する.

    Args:
        storage: ジョブ保存先。
        job_id: 対象ジョブ ID。
        session_id: チャットセッション ID。
        user_message: 今回のユーザー発話。
        reply: アシスタント応答。

    Returns:
        保存後のチャット履歴。
    """
    now = _utc_now_iso()
    storage.append_chat_message(
        job_id,
        ChatMessage(role="user", content=user_message, timestamp=now),
        session_id=session_id,
    )
    storage.append_chat_message(
        job_id,
        ChatMessage(role="assistant", content=reply, timestamp=now),
        session_id=session_id,
    )
    return storage.load_chat_history(job_id, session_id=session_id)


@router.post("/chat/{job_id}")
async def chat(
    job_id: str,
    body: ChatRequest,
    session_id: str = Query(default="default"),
    storage: Storage = Depends(get_storage),
    llm: LLMClient = Depends(get_llm_client),
) -> dict[str, Any]:
    """ユーザー発話を受け付け、LLM 応答を返す. 履歴は jsonl に追記する."""
    if not body.message.strip():
        raise HTTPException(status_code=400, detail="message is required")
    logger.info("chat request: message_chars=%d", len(body.message))

    messages = _build_chat_messages(storage, job_id, session_id, body.message)

    # LLM ループ全体を threadpool に逃がす (event loop を解放).
    reply, tool_trace = await asyncio.to_thread(_run_tool_loop, llm, storage, job_id, messages)

    # 履歴に追記 (user → assistant). tool 呼び出しの中間結果は履歴には残さない.
    # ファイル書き込みも一応 blocking なので to_thread に逃がす.
    new_history = await asyncio.to_thread(
        _persist_chat_turn,
        storage,
        job_id,
        session_id,
        body.message,
        reply,
    )
    return {
        "reply": reply,
        "history": [m.model_dump() for m in new_history],
        "tool_trace": tool_trace,
    }


@router.post("/chat/{job_id}/stream")
async def chat_stream(
    job_id: str,
    body: ChatRequest,
    session_id: str = Query(default="default"),
    storage: Storage = Depends(get_storage),
    llm: LLMClient = Depends(get_llm_client),
) -> StreamingResponse:
    """ユーザー発話を受け付け、進捗イベントと最終応答を SSE で返す."""
    if not body.message.strip():
        raise HTTPException(status_code=400, detail="message is required")
    logger.info("chat stream request: message_chars=%d", len(body.message))
    messages = _build_chat_messages(storage, job_id, session_id, body.message)

    async def _events() -> AsyncIterator[str]:
        queue: asyncio.Queue[tuple[str, dict[str, Any]] | None] = asyncio.Queue()
        loop = asyncio.get_running_loop()

        def _emit(event: str, data: dict[str, Any]) -> None:
            loop.call_soon_threadsafe(queue.put_nowait, (event, data))

        def _work() -> None:
            try:
                reply, tool_trace = _run_tool_loop(
                    llm,
                    storage,
                    job_id,
                    messages,
                    emit_progress=_emit,
                )
                new_history = _persist_chat_turn(storage, job_id, session_id, body.message, reply)
                _emit(
                    "final",
                    {
                        "reply": reply,
                        "history": [m.model_dump() for m in new_history],
                        "tool_trace": tool_trace,
                    },
                )
            except Exception:
                logger.exception("chat stream failed")
                _emit(
                    "error",
                    {
                        "message": (
                            "応答生成中にエラーが発生しました。"
                            "時間を置いて再度試すか、質問範囲を具体化してください。"
                        )
                    },
                )
            finally:
                loop.call_soon_threadsafe(queue.put_nowait, None)

        task = asyncio.create_task(asyncio.to_thread(_work))
        while True:
            item = await queue.get()
            if item is None:
                break
            event, data = item
            yield _sse_event(event, data)
        await task

    return StreamingResponse(
        _events(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@router.get("/chat/{job_id}/history")
async def chat_history(
    job_id: str,
    session_id: str = Query(default="default"),
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """ジョブのチャット履歴を返す."""
    try:
        history = storage.load_chat_history(job_id, session_id=session_id)
        # ジョブ存在確認
        storage.get_meta(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job or chat session not found") from e

    return {"history": [m.model_dump() for m in history]}


@router.get("/chat/{job_id}/sessions")
async def list_chat_sessions(
    job_id: str,
    include_archived: bool = Query(default=False),
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """ジョブに紐づくチャットセッション一覧を返す."""
    try:
        sessions = storage.list_chat_sessions(job_id, include_archived=include_archived)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e
    return {"sessions": [s.model_dump() for s in sessions]}


@router.post("/chat/{job_id}/sessions")
async def create_chat_session(
    job_id: str,
    body: ChatSessionCreateRequest,
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """ジョブ配下に新しいチャットセッションを作成する."""
    try:
        session = storage.create_chat_session(job_id, title=body.title)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e
    return {"session": session.model_dump()}


@router.patch("/chat/{job_id}/sessions/{session_id}")
async def update_chat_session(
    job_id: str,
    session_id: str,
    body: ChatSessionUpdateRequest,
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """チャットセッションのタイトルまたはアーカイブ状態を更新する."""
    try:
        session = storage.update_chat_session(
            job_id,
            session_id,
            title=body.title,
            archived=body.archived,
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job or chat session not found") from e
    return {"session": session.model_dump()}
