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
    "propose_named_range_fix": "名前定義修正の影響範囲を試算中",
    "propose_fixed_ref_replace": "固定参照置換の影響範囲を試算中",
    "propose_range_expansion": "数式範囲拡張の影響範囲を試算中",
    "propose_cell_text_edits": "説明テキスト追加の変更内容を準備中",
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


def _tool_trace_result(result_str: str) -> Any:
    """tool 結果を UI 用の構造化値に変換する.

    Args:
        result_str: tool が返した JSON 文字列。

    Returns:
        JSON として読めれば dict/list 等を返す。切り詰め済みなどで読めない場合は None。
    """
    try:
        return json.loads(result_str)
    except json.JSONDecodeError:
        return None


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
        "あなたは xlblueprint のチャットアシスタントです。",
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
        "- 数式、VBA、セル構造を変更する提案で波及範囲を述べる前に、"
        "lookup_references を呼んで実際の参照元を確認する。表示形式だけの提案では"
        "セルごとの参照検索を繰り返さない",
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
        "## 4. ユーザーの負荷を最小化する",
        "- 質問する前に、設計書とツールでアプリ側が判断できることを先に調べる",
        "- 質問は、回答によって改修内容が大きく変わる場合だけに限定する",
        "- 質問が必要でも一度に1問だけにする。最大3つの短い選択肢を示し、"
        "推奨案を先頭に置く。合理的に任せられる場合は「おすすめで任せる」を含める",
        "- 代表データ、受入条件、影響範囲の列挙を定型的にユーザーへ要求しない。"
        "必要な調査と草案作成はアプリ側で行う",
        "- 曖昧さが軽微なら、妥当な仮定を一文で示して改修案を先に提示する",
        "",
        "## 5. 判断に必要な要点へ絞る",
        "- ユーザーの目的達成に影響する論点を優先し、調査過程や重複する根拠は省く",
        "- 波及範囲を聞かれたら件数と主要箇所を要約する。全セルや全数式は、"
        "ユーザーが詳細を求めた場合だけ列挙する",
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
        "# 応答の見せ方",
        "",
        "- 結論または変更後の姿を最初に書く。調査過程から書き始めない",
        "- 通常は「提案」「変更する場所」「確認が必要なこと」の最大3ブロックに収める",
        "- 初回の改修案は原則400文字以内、箇条書きは4項目以内にする。"
        "収まらない詳細は省略し、必要ならユーザーが後から開ける形にする",
        "- ユーザーが明示的に求めていない数式一覧、セル一覧、ツール名、調査手順は本文に載せない",
        "- 空のブロックは表示しない。見出し、箇条書き、注意書きを必要以上に増やさない",
        "- 根拠と未解析リスクは各1〜2行の要点に圧縮する。完全な波及範囲、"
        "受入条件、手動確認チェックリストは、適用直前またはユーザーが詳細を求めた時だけ示す",
        "- 改修手順が必要な場合も、ユーザーが次に行う操作を先頭に短く示す。"
        "VBA・数式の書き換えを伴う場合は下記「修正提案の出し方」に従うこと",
        "",
        "禁止表現:",
        "- 「影響ありません」「使われていません」「安全です」と断定しない",
        "- 代わりに「静的解析で確認できた範囲では」「未解析リスクとして」を使う",
        "",
        "# 修正提案の出し方 (VBA・数式を書き換える場合)",
        "",
        "ユーザーは VBA に不慣れであることが前提です (SPEC §1.2)。"
        "「この行を~に直してください」「~の処理をこう変えてください」のような"
        "部分的な差分の説明だけでは、書き換え箇所を取り違えて別のバグを"
        "埋め込むおそれがあります。改修手順は必ずコピー&ペーストだけで"
        "完結する形で提示してください。",
        "",
        "## VBA プロシージャを書き換える場合",
        "- 提案前に必ず `get_vba_procedure` で現在のプロシージャ全体を取得する。"
        "全体を確認しないまま書き換え案を出さない",
        "- 提案するコードは変更箇所の抜粋ではなく、"
        "`Sub ... End Sub` / `Function ... End Function` を含む"
        "書き換え後のプロシージャ全体をコードブロックで示す",
        "- 「私が関数全体のコードを表示するので、それをそのままコピーして"
        "[モジュール名] モジュールの [プロシージャ名] プロシージャを"
        "(既存のコードを削除してから) 書き換えてください」のように、"
        "コピー先とやる操作をそのまま実行できる言葉で明記する",
        "",
        "## セルの数式を書き換える場合",
        "- 対象セル (シート名+セル番地) を明記し、書き換え後の完全な式"
        "(`=` から始まる全体) をコードブロックで示す。式の一部だけを"
        "示して「ここを削ってください」とは言わない",
        "- 数式そのものを別の式に置き換える場合は「[シート名]![セル番地] を選択し、"
        "この式をそのまま貼り付けて Enter を押してください」と明記する",
        "- 数式を削除して結果を固定値にする場合は「この式をコピーし、"
        "[シート名]![セル番地] に『値のみ貼り付け』"
        "(右クリック→貼り付けのオプション→値) をしてください」のように、"
        "数式ではなく値として貼り付けることを明記する",
        "- 複数セルにまたがる場合もセルごとに完全な式を示す。"
        "「◯列全体に同様の処理をしてください」のような曖昧な指示はしない",
        "",
        "禁止:",
        "- 行番号や部分抜粋だけを示して「ここを~に直してください」と指示すること",
        "- 「このロジックを~のように変えてください」だけで、"
        "書き換え後の完全なコード・数式を示さないこと",
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
        "- `propose_named_range_fix(name, new_refers_to)`:",
        '    例: propose_named_range_fix("TaxRate", "Data!$B$1")',
        "    名前付き範囲の参照先を書き換えたら何が変わるかを試算する読み取り専用ツール。"
        "このツール自体はファイルを一切変更しない。",
        "- `propose_fixed_ref_replace(old_ref, new_ref)`:",
        '    例: propose_fixed_ref_replace("Data!$B$5", "Data!$B$6")',
        "    数式内の固定参照を別の参照に置き換えたらどの数式がどう変わるかを試算する"
        "読み取り専用ツール。参照はシート修飾付きで指定する。",
        "- `propose_range_expansion(old_range, new_range)`:",
        '    例: propose_range_expansion("Data!$A$1:$A$100", "Data!$A$1:$A$200")',
        "    数式が参照している範囲を広げたらどの数式がどう変わるかを試算する"
        "読み取り専用ツール。データ行が増えて集計範囲が足りない、という依頼で使う。"
        "new_range は old_range と同一シートで old_range を包含すること。",
        "- `propose_cell_text_edits(edits)`:",
        '    例: propose_cell_text_edits([{"sheet":"Output","coord":"C3","value":"説明"}])',
        "    現在空欄のセルへ説明、注記、見出し等の固定テキストを追加する計画を作る"
        "読み取り専用ツール。既存値や数式の上書きは拒否される。",
        "",
        "# 自動適用できる修正依頼を受けたとき",
        "",
        "- ユーザーから該当パターンの修正依頼を受けたら、まず対応する propose 系ツール"
        "(`propose_named_range_fix` / `propose_fixed_ref_replace` / `propose_range_expansion` / "
        "`propose_cell_text_edits`)"
        "を呼び、変更される箇所・波及範囲・既存リスクを試算してから提示すること",
        "- 実際の適用は必ずユーザーが画面上のボタンで明示的に行う。"
        "あなた自身がファイルを書き換えることはできないし、してはいけない",
        "- propose 系ツールで対応できる依頼に対して、Excelを手作業で編集する手順を"
        "長々と案内しない。変更カードの「修正版を作る」ボタンを案内すること",
        "- 提案した内容と実際の適用結果が食い違わないよう、propose 系ツールに渡す引数は"
        "ユーザーの意図を正確に反映した値にすること",
        "- 上記パターンに当てはまらない修正 (VBA 変更・行列の挿入削除・複雑な数式の"
        "書き換え等) は自動適用できない。従来どおりコピペ完結の改修手順を提示すること",
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


def _requests_cell_text_change(messages: list[dict[str, Any]]) -> bool:
    """最新の依頼が空セルへの説明追加として扱える可能性が高いか判定する."""

    user_message = next(
        (
            str(message.get("content", ""))
            for message in reversed(messages)
            if message.get("role") == "user"
        ),
        "",
    ).lower()
    text_terms = ("説明", "注記", "見出し", "description", "note", "header")
    action_terms = ("追加", "入れ", "加え", "書き", "作り", "設け", "add", "insert")
    return any(term in user_message for term in text_terms) and any(
        term in user_message for term in action_terms
    )


def _needs_cell_text_tool_retry(
    messages: list[dict[str, Any]],
    tool_trace: list[dict[str, Any]],
    retry_count: int,
) -> bool:
    """対応可能な説明追加依頼で、変更カード未作成なら最大2回再誘導する."""

    return (
        retry_count < 2
        and _requests_cell_text_change(messages)
        and not any(item.get("name") == "propose_cell_text_edits" for item in tool_trace)
    )


def _final_reply_for_tool_trace(
    content: str,
    tool_trace: list[dict[str, Any]],
) -> str:
    """操作カードが主役の応答は、LLMの長文を短い操作案内へ置き換える."""

    cell_text_item = next(
        (item for item in tool_trace if item.get("name") == "propose_cell_text_edits"),
        None,
    )
    if cell_text_item is None:
        return content
    result = cell_text_item.get("result")
    safe_plan = result.get("safe_plan") if isinstance(result, dict) else None
    summary = safe_plan.get("summary") if isinstance(safe_plan, dict) else None
    lead = str(summary) if summary else "説明テキストの追加案を作りました。"
    return (
        f"{lead}\n\n"
        "下の「変更内容を確認」で内容を見て、「修正版を作る」を押してください。"
        "原本は変更されません。"
    )


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
    cell_text_retry_count = 0

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
            if _needs_cell_text_tool_retry(messages, tool_trace, cell_text_retry_count):
                cell_text_retry_count += 1
                logger.info("retrying supported cell text request with explicit tool guidance")
                if resp.content:
                    messages.append({"role": "assistant", "content": resp.content})
                messages.append(
                    {
                        "role": "system",
                        "content": (
                            "この依頼は空セルへの説明テキスト追加として自動対応できます。"
                            "手作業のExcel操作手順を最終回答にしないでください。"
                            "対象セルをまだ確認していなければ get_cells_range を呼び、"
                            "確認済みなら既存値や数式を上書きしない edits を組み立てて、"
                            "必ず propose_cell_text_edits を呼んで変更カードを作成してください。"
                            "列の挿入や書式変更は提案せず、現在空欄のセルへの値追加だけに限定します。"
                        ),
                    }
                )
                continue
            reply = _final_reply_for_tool_trace(resp.content or "", tool_trace)
            logger.info(
                "llm final response: iterations=%d tool_calls_total=%d reply_chars=%d "
                "cumulative_prompt=%d completion=%d total=%d cached=%d",
                iteration + 1,
                len(tool_trace),
                len(reply),
                total_usage.prompt_tokens,
                total_usage.completion_tokens,
                total_usage.total_tokens,
                total_usage.cached_tokens,
            )
            return (reply, tool_trace)

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
                {
                    "name": tc.name,
                    "arguments": tc.arguments,
                    "result_preview": result_str[:200],
                    "result": _tool_trace_result(result_str),
                }
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
    tool_trace: list[dict[str, Any]],
) -> list[ChatMessage]:
    """ユーザー発話とアシスタント応答を履歴に保存して再読込する.

    Args:
        storage: ジョブ保存先。
        job_id: 対象ジョブ ID。
        session_id: チャットセッション ID。
        user_message: 今回のユーザー発話。
        reply: アシスタント応答。
        tool_trace: アシスタント応答の根拠として表示するツール実行結果。

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
        ChatMessage(role="assistant", content=reply, timestamp=now, tool_trace=tool_trace),
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
        tool_trace,
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
                new_history = _persist_chat_turn(
                    storage,
                    job_id,
                    session_id,
                    body.message,
                    reply,
                    tool_trace,
                )
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
