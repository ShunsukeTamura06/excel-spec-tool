"""社内LLM クライアント.

SPEC.md §5.3 に基づき、OpenAI 互換 API を想定したインターフェースを提供する。

設計:
- `LLMClient` Protocol で実装を切り替え可能にする
- `MockLLMClient`: 決定的な応答を返す。テスト・開発時のデフォルト
- `OpenAICompatibleLLMClient`: 実 API 接続のプレースホルダ。
  社内 LLM の仕様確定後 Shun が実装を埋める想定。現状はメソッド呼び出しで
  NotImplementedError を投げる
- 環境変数 (LLM_BASE_URL / LLM_API_KEY / LLM_MODEL) が揃っていれば実クライアントを、
  そうでなければモックを返す `get_default_client()` ファクトリ
- モジュールレベル `chat_completion()` / `annotate_text()` は SPEC §5.3 の公開
  シグネチャ。内部でデフォルトクライアントに委譲

Function calling:
- `chat_completion_with_tools(messages, tools, model)` で tool 呼び出し可能。
  応答は `LLMResponse(content, tool_calls)`. content だけなら完了、
  tool_calls があれば呼び出し側でツールを実行し結果を message に追記して再度
  呼ぶループを構成する。

外部クラウドへのデータ送信は SPEC.md §1.3 で禁止されているため、実クライアントは
社内 LLM (プロキシ経由 / ホワイトリスト URL) のみを叩く前提。
"""

from __future__ import annotations

import logging
import os
from collections import deque
from dataclasses import dataclass, field
from typing import Any, Protocol, runtime_checkable

logger = logging.getLogger(__name__)


# ChatGPT 互換 API のメッセージ型. 単純な dict.
# {"role": "user"|"assistant"|"system"|"tool", "content": "..."} + 必要なら
# "tool_calls" や "tool_call_id" 等を含む
ChatMessageDict = dict[str, Any]


@dataclass
class LLMToolCall:
    """LLM が要求した tool 呼び出し 1 件."""

    id: str
    name: str
    arguments: dict[str, Any]


@dataclass
class LLMResponse:
    """tool calling 対応の応答.

    content と tool_calls は同時に空でないこともあり得る (OpenAI 仕様)。
    content のみ → 通常応答
    tool_calls のみ → tool 実行を要求している
    """

    content: str | None = None
    tool_calls: list[LLMToolCall] = field(default_factory=list)


@runtime_checkable
class LLMClient(Protocol):
    """LLM クライアントのインターフェース."""

    def chat_completion(
        self,
        messages: list[ChatMessageDict],
        model: str | None = None,
    ) -> str:
        """チャット補完 (SPEC §5.3). 最終アシスタント応答テキストを返す."""
        ...

    def annotate_text(self, prompt: str, content: str) -> str:
        """注釈用ユーティリティ."""
        ...

    def chat_completion_with_tools(
        self,
        messages: list[ChatMessageDict],
        tools: list[dict[str, Any]],
        model: str | None = None,
    ) -> LLMResponse:
        """tool 呼び出し可能なチャット補完."""
        ...


class MockLLMClient:
    """テスト/開発用の決定的なクライアント.

    通常呼び出しはエコー風の応答を返す。
    tool calling 用には `queue_response()` で次に返す応答を予約できる。
    """

    def __init__(self) -> None:
        self._queue: deque[LLMResponse] = deque()

    def queue_response(self, response: LLMResponse) -> None:
        """次回 chat_completion_with_tools 呼び出しに返す応答を予約する."""
        self._queue.append(response)

    def chat_completion(
        self,
        messages: list[ChatMessageDict],
        model: str | None = None,
    ) -> str:
        last_user = ""
        for m in reversed(messages):
            if m.get("role") == "user":
                last_user = str(m.get("content", ""))
                break
        model_label = model or "mock"
        return f"[mock:{model_label}] received: {last_user}"

    def annotate_text(self, prompt: str, content: str) -> str:
        head = content.strip().splitlines()[0] if content.strip() else ""
        head = head[:80]
        return f"[mock annotation] prompt={prompt!r} head={head!r}"

    def chat_completion_with_tools(
        self,
        messages: list[ChatMessageDict],
        tools: list[dict[str, Any]],
        model: str | None = None,
    ) -> LLMResponse:
        # 予約済み応答があればそれを返す
        if self._queue:
            return self._queue.popleft()
        # デフォルト: tool は呼ばずに通常応答
        return LLMResponse(content=self.chat_completion(messages, model=model))


class OpenAICompatibleLLMClient:
    """OpenAI 互換 API への HTTP クライアント (プレースホルダ).

    `LLM_BASE_URL` / `LLM_API_KEY` / `LLM_MODEL` を読んで `httpx.post` する想定。
    社内仕様確定後に実装する。現状は呼び出し時に NotImplementedError を投げる。
    """

    def __init__(self, base_url: str, api_key: str, default_model: str) -> None:
        self.base_url = base_url
        self.api_key = api_key
        self.default_model = default_model

    def chat_completion(
        self,
        messages: list[ChatMessageDict],
        model: str | None = None,
    ) -> str:
        # TODO(Shun): 社内 LLM の仕様確定後、httpx で /chat/completions を叩く実装に差し替える。
        raise NotImplementedError(
            "OpenAICompatibleLLMClient.chat_completion is not implemented yet."
        )

    def annotate_text(self, prompt: str, content: str) -> str:
        raise NotImplementedError("OpenAICompatibleLLMClient.annotate_text is not implemented yet.")

    def chat_completion_with_tools(
        self,
        messages: list[ChatMessageDict],
        tools: list[dict[str, Any]],
        model: str | None = None,
    ) -> LLMResponse:
        # TODO(Shun): tools 配列を payload に含めて社内 LLM を呼び、
        # 応答の choices[0].message.tool_calls を LLMToolCall に詰めて返す。
        raise NotImplementedError(
            "OpenAICompatibleLLMClient.chat_completion_with_tools is not implemented yet."
        )


def _read_env() -> tuple[str | None, str | None, str | None]:
    return (
        os.environ.get("LLM_BASE_URL") or None,
        os.environ.get("LLM_API_KEY") or None,
        os.environ.get("LLM_MODEL") or None,
    )


def get_default_client() -> LLMClient:
    """環境変数を見てデフォルトクライアントを返す.

    `LLM_BASE_URL` と `LLM_API_KEY` の両方が設定されていれば
    `OpenAICompatibleLLMClient` を、そうでなければ `MockLLMClient` を返す。
    モデル名は `LLM_MODEL` から (なければ `"gpt-5.2"`).
    """
    base_url, api_key, model = _read_env()
    if base_url and api_key:
        return OpenAICompatibleLLMClient(
            base_url=base_url,
            api_key=api_key,
            default_model=model or "gpt-5.2",
        )
    logger.info("LLM env not set; falling back to MockLLMClient")
    return MockLLMClient()


# ---------- SPEC.md §5.3 が定める公開関数 ----------


def chat_completion(
    messages: list[ChatMessageDict],
    model: str | None = None,
) -> str:
    """デフォルトクライアントで chat_completion を実行する."""
    return get_default_client().chat_completion(messages, model=model)


def annotate_text(prompt: str, content: str) -> str:
    """デフォルトクライアントで annotate_text を実行する."""
    return get_default_client().annotate_text(prompt, content)
