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

外部クラウドへのデータ送信は SPEC.md §1.3 で禁止されているため、実クライアントは
社内 LLM (プロキシ経由 / ホワイトリスト URL) のみを叩く前提。
"""

from __future__ import annotations

import logging
import os
from typing import Protocol, runtime_checkable

logger = logging.getLogger(__name__)


# ChatGPT 互換 API のメッセージ型. dict のままだが、最低限のキーだけドキュメント化:
#   {"role": "user" | "assistant" | "system", "content": "..."}
ChatMessageDict = dict[str, str]


@runtime_checkable
class LLMClient(Protocol):
    """LLM クライアントのインターフェース."""

    def chat_completion(
        self,
        messages: list[ChatMessageDict],
        model: str | None = None,
    ) -> str:
        """チャット補完. 最終アシスタント応答 (テキスト) を返す."""
        ...

    def annotate_text(self, prompt: str, content: str) -> str:
        """注釈用ユーティリティ. 与えた content を prompt に従って解釈した結果を返す."""
        ...


class MockLLMClient:
    """テスト/開発用の決定的なクライアント.

    実 API は叩かず、入力に基づいた合成応答を返す。
    """

    def chat_completion(
        self,
        messages: list[ChatMessageDict],
        model: str | None = None,
    ) -> str:
        last_user = ""
        for m in reversed(messages):
            if m.get("role") == "user":
                last_user = m.get("content", "")
                break
        model_label = model or "mock"
        return f"[mock:{model_label}] received: {last_user}"

    def annotate_text(self, prompt: str, content: str) -> str:
        # content が長すぎる場合は冒頭だけ要約に使う
        head = content.strip().splitlines()[0] if content.strip() else ""
        head = head[:80]
        return f"[mock annotation] prompt={prompt!r} head={head!r}"


class OpenAICompatibleLLMClient:
    """OpenAI 互換 API への HTTP クライアント (プレースホルダ).

    `LLM_BASE_URL` / `LLM_API_KEY` / `LLM_MODEL` を読んで `httpx.post` する想定。
    社内仕様確定後に実装する。現状はモック以外を選んだことを検出するため、
    呼び出し時に NotImplementedError を投げる。
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
            "OpenAICompatibleLLMClient.chat_completion is not implemented yet. "
            "Replace with internal LLM HTTP call."
        )

    def annotate_text(self, prompt: str, content: str) -> str:
        # TODO(Shun): 同上。chat_completion を組み立てて呼ぶ薄いラッパで足りる想定。
        raise NotImplementedError("OpenAICompatibleLLMClient.annotate_text is not implemented yet.")


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
