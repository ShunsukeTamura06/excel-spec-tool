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
from datetime import datetime, timezone
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, Query
from pydantic import BaseModel

from backend.dependencies import get_llm_client, get_storage
from backend.llm_client import LLMClient, LLMResponse, Usage
from backend.llm_tools import build_tool_definitions, execute_tool_call
from backend.storage import JobNotFoundError, Storage
from core.models import ChatMessage

router = APIRouter()
logger = logging.getLogger(__name__)


# tool ループの最大反復回数. これを超えたら強制的に終了する。
MAX_TOOL_ITERATIONS = 8

# LLM に送る履歴の上限ペア数 (Phase B-2). user/assistant の 1 往復 = 1 ペア.
# これを超えた古いペアは「過去のやりとり概要」として 1 つの system message に集約する。
# 履歴ファイル (chat_history.jsonl) 自体は完全保存される。LLM 文脈窓だけ絞る。
_DEFAULT_HISTORY_PAIRS = 10
# 古い履歴を要約する際の、1 メッセージあたりの最大文字数 (それ以上は切り詰める)
_SUMMARY_MESSAGE_MAX_CHARS = 120


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
        "",
        "# 応答フォーマット",
        "",
        "通常の回答には以下のセクションを必ず含めてください:",
        "",
        "1. **改修手順**: ユーザーの操作レベル (どのセル/タブを開き、何を入力するか) を具体的に",
        "2. **波及範囲**: 影響を受けるセルや VBA を一覧 "
        "(件数 + 主要箇所、調査根拠としてどのツールを呼んだか)",
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
        "ツールの呼び出しに上限はありません。確実性を優先して、必要な回数呼んでください。",
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
) -> tuple[str, list[dict[str, Any]]]:
    """tool 呼び出しを伴うチャット応答ループ.

    Returns:
        (最終アシスタント応答テキスト, ループ中の中間 tool 呼び出し記録).
        中間記録は履歴 jsonl には保存しないがデバッグ・観測用に返す。
    """
    tools = build_tool_definitions()
    messages = list(base_messages)
    tool_trace: list[dict[str, Any]] = []
    total_usage = Usage()

    for iteration in range(MAX_TOOL_ITERATIONS):
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
            result_str = execute_tool_call(storage, job_id, tc.name, tc.arguments)
            logger.info(
                "tool result: name=%s result_chars=%d preview=%s",
                tc.name,
                len(result_str),
                result_str[:120].replace("\n", " "),
            )
            tool_trace.append(
                {"name": tc.name, "arguments": tc.arguments, "result_preview": result_str[:200]}
            )
            messages.append(
                {
                    "role": "tool",
                    "tool_call_id": tc.id,
                    "name": tc.name,
                    "content": result_str,
                }
            )

    # 上限到達: 最後の応答を返す。content が空の場合は注意書きを付与。
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
    return (
        (resp.content or "[tool loop max iterations reached; partial result]"),
        tool_trace,
    )


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
    max_pairs = _get_history_pairs_limit()
    recent, summary = _trim_history(history, max_pairs=max_pairs)
    if summary:
        logger.info(
            "history trimmed: total=%d kept=%d summarized=%d",
            len(history),
            len(recent),
            len(history) - len(recent),
        )

    messages: list[dict[str, Any]] = [
        {"role": "system", "content": _build_system_prompt(spec_md)},
    ]
    if summary:
        # system 直後に置く. 設計書 (大プレフィックス) のキャッシュ性を壊さないよう、
        # main system prompt とは別メッセージにする。
        messages.append({"role": "system", "content": summary})
    messages.extend({"role": m.role, "content": m.content} for m in recent)
    messages.append({"role": "user", "content": body.message})

    # LLM ループ全体を threadpool に逃がす (event loop を解放).
    reply, tool_trace = await asyncio.to_thread(_run_tool_loop, llm, storage, job_id, messages)

    # 履歴に追記 (user → assistant). tool 呼び出しの中間結果は履歴には残さない.
    # ファイル書き込みも一応 blocking なので to_thread に逃がす.
    now = _utc_now_iso()

    def _persist_and_reload() -> list[ChatMessage]:
        storage.append_chat_message(
            job_id,
            ChatMessage(role="user", content=body.message, timestamp=now),
            session_id=session_id,
        )
        storage.append_chat_message(
            job_id,
            ChatMessage(role="assistant", content=reply, timestamp=now),
            session_id=session_id,
        )
        return storage.load_chat_history(job_id, session_id=session_id)

    new_history = await asyncio.to_thread(_persist_and_reload)
    return {
        "reply": reply,
        "history": [m.model_dump() for m in new_history],
        "tool_trace": tool_trace,
    }


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
