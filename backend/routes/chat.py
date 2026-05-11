"""POST /chat/{job_id} と GET /chat/{job_id}/history — 改修対話.

LLM の function calling ループを実装している:
1. system prompt + 履歴 + ユーザー発話 + tools 定義を LLM に渡す
2. 応答に tool_calls があれば、各 tool を実行して結果を tool role の
   メッセージとして追加し、再度 LLM を呼ぶ
3. tool_calls が無くなったら最終応答テキストを返す
4. 暴走防止のため最大反復回数を制限する
"""

from __future__ import annotations

import json
import logging
from datetime import datetime, timezone
from typing import Any

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel

from backend.dependencies import get_llm_client, get_storage
from backend.llm_client import LLMClient, LLMResponse
from backend.llm_tools import build_tool_definitions, execute_tool_call
from backend.storage import JobNotFoundError, Storage
from core.models import ChatMessage

router = APIRouter()
logger = logging.getLogger(__name__)


# tool ループの最大反復回数. これを超えたら強制的に終了する。
MAX_TOOL_ITERATIONS = 8


class ChatRequest(BaseModel):
    """POST /chat/{job_id} のリクエストボディ."""

    message: str


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


_SYSTEM_INSTRUCTIONS = "\n".join(
    [
        "あなたは Excel 改修支援アシスタントです。",
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
        "- 不明な部分があってもそれを「不明」として明示すれば回答漏れにはならない",
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
        "(シート一覧 / 先頭 20 行 × 20 列のプレビュー / 表構造 / VBA 抽出結果 / "
        "名前付き範囲 / 主要数式 TOP10) のみ載っています。",
        "詳細な値・行内容・参照関係を確認したい時は、以下のツールを必ず呼んでください:",
        "",
        "- `get_cells_range(sheet, range)`:",
        '    例: get_cells_range("Portfolio", "A6:Z6")',
        "    指定範囲のセル値・数式・データ型を 2D 配列で返す。プレビュー外の領域を見たい時はこれ",
        "- `find_cells(query, sheet?, limit?)`:",
        '    例: find_cells("実現損益", sheet="Portfolio")',
        "    値の部分一致でセルを検索する。ユーザーが言った項目名がどこにあるか分からない時はこれ",
        "- `lookup_references(target)`:",
        '    例: lookup_references("Calc!H2")',
        "    あるセル/範囲を参照している箇所 (数式・VBA) を返す。波及範囲調査ではこれを必ず使う",
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

    for _iteration in range(MAX_TOOL_ITERATIONS):
        resp = llm.chat_completion_with_tools(messages, tools=tools)

        if not resp.tool_calls:
            # 最終応答
            return (resp.content or "", tool_trace)

        # tool 呼び出しを実行し、結果を tool role メッセージで追加
        messages.append(_tool_call_to_assistant_message(resp))
        for tc in resp.tool_calls:
            result_str = execute_tool_call(storage, job_id, tc.name, tc.arguments)
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
    logger.warning("Tool loop hit MAX_TOOL_ITERATIONS=%d for job %s", MAX_TOOL_ITERATIONS, job_id)
    return (
        (resp.content or "[tool loop max iterations reached; partial result]"),
        tool_trace,
    )


@router.post("/chat/{job_id}")
async def chat(
    job_id: str,
    body: ChatRequest,
    storage: Storage = Depends(get_storage),
    llm: LLMClient = Depends(get_llm_client),
) -> dict[str, Any]:
    """ユーザー発話を受け付け、LLM 応答を返す. 履歴は jsonl に追記する."""
    if not body.message.strip():
        raise HTTPException(status_code=400, detail="message is required")

    try:
        try:
            spec_md = storage.load_spec(job_id)
        except FileNotFoundError:
            spec_md = ""

        history = storage.load_chat_history(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e

    messages: list[dict[str, Any]] = [
        {"role": "system", "content": _build_system_prompt(spec_md)},
    ]
    messages.extend({"role": m.role, "content": m.content} for m in history)
    messages.append({"role": "user", "content": body.message})

    reply, tool_trace = _run_tool_loop(llm, storage, job_id, messages)

    # 履歴に追記 (user → assistant). tool 呼び出しの中間結果は履歴には残さない
    now = _utc_now_iso()
    storage.append_chat_message(
        job_id, ChatMessage(role="user", content=body.message, timestamp=now)
    )
    storage.append_chat_message(job_id, ChatMessage(role="assistant", content=reply, timestamp=now))

    new_history = storage.load_chat_history(job_id)
    return {
        "reply": reply,
        "history": [m.model_dump() for m in new_history],
        "tool_trace": tool_trace,
    }


@router.get("/chat/{job_id}/history")
async def chat_history(
    job_id: str,
    storage: Storage = Depends(get_storage),
) -> dict[str, Any]:
    """ジョブのチャット履歴を返す."""
    try:
        history = storage.load_chat_history(job_id)
        # ジョブ存在確認
        storage.get_meta(job_id)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"invalid job_id: {e}") from e
    except JobNotFoundError as e:
        raise HTTPException(status_code=404, detail="job not found") from e

    return {"history": [m.model_dump() for m in history]}
