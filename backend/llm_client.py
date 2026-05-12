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

import json
import logging
import os
from collections import deque
from dataclasses import dataclass, field
from typing import Any, Literal, Protocol, runtime_checkable

logger = logging.getLogger(__name__)


# ChatGPT 互換 API のメッセージ型. 単純な dict.
# {"role": "user"|"assistant"|"system"|"tool", "content": "..."} + 必要なら
# "tool_calls" や "tool_call_id" 等を含む
ChatMessageDict = dict[str, Any]


# モデルの tier. "pro" = 高精度モデル, "fast" = 高速・低コストモデル.
# 呼び出し側がタスクの難易度に応じて指定する。実モデル名は環境変数 LLM_MODEL_PRO /
# LLM_MODEL_FAST または個別クライアントの設定で解決される。
ModelTier = Literal["pro", "fast"]
_DEFAULT_PRO_MODEL = "gpt-5.2"


@dataclass
class LLMToolCall:
    """LLM が要求した tool 呼び出し 1 件."""

    id: str
    name: str
    arguments: dict[str, Any]


@dataclass
class Usage:
    """LLM 呼び出し 1 回分のトークン消費量.

    OpenAI 互換 API の `response.usage` をそのまま正規化したもの。
    `cached_tokens` は prompt のうちキャッシュ命中した分 (Phase B で活用予定)。
    """

    prompt_tokens: int = 0
    completion_tokens: int = 0
    total_tokens: int = 0
    cached_tokens: int = 0

    def __add__(self, other: Usage) -> Usage:
        return Usage(
            prompt_tokens=self.prompt_tokens + other.prompt_tokens,
            completion_tokens=self.completion_tokens + other.completion_tokens,
            total_tokens=self.total_tokens + other.total_tokens,
            cached_tokens=self.cached_tokens + other.cached_tokens,
        )


@dataclass
class LLMResponse:
    """tool calling 対応の応答.

    content と tool_calls は同時に空でないこともあり得る (OpenAI 仕様)。
    content のみ → 通常応答
    tool_calls のみ → tool 実行を要求している

    `usage` と `model` は観測用 (どのモデルが何トークン使ったかをログ集計するため)。
    実装が usage を取れない場合は None でも構わない。
    """

    content: str | None = None
    tool_calls: list[LLMToolCall] = field(default_factory=list)
    usage: Usage | None = None
    model: str | None = None


@runtime_checkable
class LLMClient(Protocol):
    """LLM クライアントのインターフェース.

    `tier` は高精度 ("pro") / 高速 ("fast") の選択ヒント。実モデル名は
    クライアントの設定 (LLM_MODEL_PRO / LLM_MODEL_FAST) から解決される。
    `model` を明示指定すると tier より優先される。

    `cache_prefix` は prompt caching (Phase B-3) のヒント:
      - 値は「messages の先頭から数えて、キャッシュ対象とみなすメッセージ数」
      - 0 (デフォルト) なら明示的なキャッシュマーカは付けない
      - OpenAI 互換 API は通常自動プレフィックスキャッシュなので 0 でも効くが、
        Anthropic 系のように明示マーカ (`cache_control: ephemeral`) が必要な
        プロバイダで使う想定。実装側がプロバイダ仕様に合わせて翻訳する。
    """

    def chat_completion(
        self,
        messages: list[ChatMessageDict],
        model: str | None = None,
        tier: ModelTier = "pro",
        cache_prefix: int = 0,
    ) -> str:
        """チャット補完 (SPEC §5.3). 最終アシスタント応答テキストを返す."""
        ...

    def annotate_text(self, prompt: str, content: str, tier: ModelTier = "fast") -> str:
        """注釈用ユーティリティ. 大量呼び出し前提なので fast がデフォルト."""
        ...

    def chat_completion_with_tools(
        self,
        messages: list[ChatMessageDict],
        tools: list[dict[str, Any]],
        model: str | None = None,
        tier: ModelTier = "pro",
        cache_prefix: int = 0,
    ) -> LLMResponse:
        """tool 呼び出し可能なチャット補完."""
        ...


def _estimate_usage(prompt_chars: int, completion_chars: int) -> Usage:
    """文字数からトークンを概算する (英文 ~4chars/token を仮定).

    モック・テスト用。実 API はサーバー応答の usage を使うのでこの推定は通らない。
    """
    p = max(1, prompt_chars // 4)
    c = max(0, completion_chars // 4)
    return Usage(prompt_tokens=p, completion_tokens=c, total_tokens=p + c)


def _sum_message_chars(messages: list[ChatMessageDict]) -> int:
    """messages の content 文字数を合計する (mock の usage 推定用)."""
    total = 0
    for m in messages:
        content = m.get("content")
        if isinstance(content, str):
            total += len(content)
    return total


class MockLLMClient:
    """テスト/開発用の決定的なクライアント.

    通常呼び出しはエコー風の応答を返す。
    tool calling 用には `queue_response()` で次に返す応答を予約できる。
    予約された応答に usage / model が無ければ、メッセージ長と tier から補完する。

    `pro_model` / `fast_model` は tier 解決のための擬似モデル名で、テストから
    上書きできる (Task C 検証用)。
    """

    def __init__(
        self,
        pro_model: str = "mock-pro",
        fast_model: str = "mock-fast",
    ) -> None:
        self._queue: deque[LLMResponse] = deque()
        self.pro_model = pro_model
        self.fast_model = fast_model
        # 直近の tier / model 履歴 (テストでアサート用)
        self.calls: list[dict[str, Any]] = []
        # model 別の直前 cacheable プレフィックス内容 (caching シミュレーション用).
        # 同じ model に対して同一プレフィックスが再度送られたら cached_tokens に
        # その分を計上する。モデルが切り替わったらキャッシュは別物として扱う。
        self._prefix_cache: dict[str, str] = {}

    def queue_response(self, response: LLMResponse) -> None:
        """次回 chat_completion_with_tools 呼び出しに返す応答を予約する."""
        self._queue.append(response)

    def _resolve(self, model: str | None, tier: ModelTier) -> str:
        """`model` 明示があれば優先、無ければ tier から解決する."""
        if model:
            return model
        return self.pro_model if tier == "pro" else self.fast_model

    def _simulate_cache_hit(
        self,
        model_label: str,
        messages: list[ChatMessageDict],
        cache_prefix: int,
    ) -> int:
        """同一 (model, prefix 内容) が直前に来ていたら cached_tokens を返す.

        副作用として、今回のプレフィックス内容を `_prefix_cache` に記録する。
        """
        if cache_prefix <= 0:
            return 0
        prefix_content = "".join(str(m.get("content", "")) for m in messages[:cache_prefix])
        previous = self._prefix_cache.get(model_label)
        self._prefix_cache[model_label] = prefix_content
        if previous is not None and previous == prefix_content:
            return max(1, len(prefix_content) // 4)
        return 0

    def chat_completion(
        self,
        messages: list[ChatMessageDict],
        model: str | None = None,
        tier: ModelTier = "pro",
        cache_prefix: int = 0,
    ) -> str:
        last_user = ""
        for m in reversed(messages):
            if m.get("role") == "user":
                last_user = str(m.get("content", ""))
                break
        model_label = self._resolve(model, tier)
        content = f"[mock:{model_label}] received: {last_user}"
        usage = _estimate_usage(_sum_message_chars(messages), len(content))
        usage.cached_tokens = self._simulate_cache_hit(model_label, messages, cache_prefix)
        logger.info(
            "llm chat_completion: model=%s tier=%s prompt_tokens=%d completion_tokens=%d "
            "total=%d cached=%d",
            model_label,
            tier,
            usage.prompt_tokens,
            usage.completion_tokens,
            usage.total_tokens,
            usage.cached_tokens,
        )
        return content

    def annotate_text(self, prompt: str, content: str, tier: ModelTier = "fast") -> str:
        head = content.strip().splitlines()[0] if content.strip() else ""
        head = head[:80]
        result = f"[mock annotation] prompt={prompt!r} head={head!r}"
        usage = _estimate_usage(len(prompt) + len(content), len(result))
        model_label = self._resolve(None, tier)
        logger.info(
            "llm annotate_text: model=%s tier=%s prompt_tokens=%d completion_tokens=%d total=%d",
            model_label,
            tier,
            usage.prompt_tokens,
            usage.completion_tokens,
            usage.total_tokens,
        )
        return result

    def chat_completion_with_tools(
        self,
        messages: list[ChatMessageDict],
        tools: list[dict[str, Any]],
        model: str | None = None,
        tier: ModelTier = "pro",
        cache_prefix: int = 0,
    ) -> LLMResponse:
        model_label = self._resolve(model, tier)
        self.calls.append(
            {
                "tier": tier,
                "model": model_label,
                "messages": list(messages),
                "cache_prefix": cache_prefix,
            }
        )
        cached_tokens = self._simulate_cache_hit(model_label, messages, cache_prefix)
        # 予約済み応答があれば優先. usage / model が欠けていれば補完する.
        if self._queue:
            resp = self._queue.popleft()
            if resp.usage is None:
                content_len = len(resp.content or "")
                resp.usage = _estimate_usage(_sum_message_chars(messages), content_len)
                resp.usage.cached_tokens = cached_tokens
            if resp.model is None:
                resp.model = model_label
            return resp
        # デフォルト: tool は呼ばずに通常応答
        # Note: 内部で chat_completion を呼ぶと _simulate_cache_hit が二重に走るので
        # ここでは展開せずに content だけ生成する
        last_user = ""
        for m in reversed(messages):
            if m.get("role") == "user":
                last_user = str(m.get("content", ""))
                break
        content = f"[mock:{model_label}] received: {last_user}"
        usage = _estimate_usage(_sum_message_chars(messages), len(content))
        usage.cached_tokens = cached_tokens
        return LLMResponse(content=content, usage=usage, model=model_label)


class OpenAICompatibleLLMClient:
    """OpenAI 互換 API への HTTP クライアント.

    `openai` SDK の `OpenAI(base_url=..., api_key=...)` で社内 LLM を叩く。
    pro / fast の 2 モデルを保持し、`tier` で切り替える。

    cache_prefix について:
      OpenAI 純正 API は自動プレフィックスキャッシュなので、本実装では cache_prefix
      パラメタは受け取るだけで明示的なマーカは付けない。
      最低 1024 tokens のプレフィックス + 同一内容なら自動でキャッシュ命中する。
      Anthropic 互換 (cache_control マーカ必須) のプロバイダに切り替える場合は、
      `_send_chat_completion` 内で messages[cache_prefix - 1] に
      {"cache_control": {"type": "ephemeral"}} を付加する処理を追加すること。

    Args:
        base_url: API ベース URL (例: `http://internal-llm.example.com/v1`)
        api_key:  API キー
        pro_model:  高精度モデル名 (例: `gpt-5.2`)
        fast_model: 高速モデル名 (例: `gpt-5.2-mini`). 省略時は pro_model と同じ。
    """

    def __init__(
        self,
        base_url: str,
        api_key: str,
        pro_model: str,
        fast_model: str | None = None,
    ) -> None:
        self.base_url = base_url
        self.api_key = api_key
        self.pro_model = pro_model
        self.fast_model = fast_model or pro_model
        # SDK クライアントを 1 個保持. import を遅延させてテストでスタブ可能に。
        from openai import OpenAI

        self._client = OpenAI(base_url=base_url, api_key=api_key)

    # 後方互換: 旧コードが `client.default_model` を参照するケースを残す.
    @property
    def default_model(self) -> str:
        return self.pro_model

    def _resolve(self, model: str | None, tier: ModelTier) -> str:
        if model:
            return model
        return self.pro_model if tier == "pro" else self.fast_model

    def _extract_usage(self, raw_usage: Any) -> Usage:
        """OpenAI 応答の usage オブジェクトを Usage に正規化する."""
        if raw_usage is None:
            return Usage()
        prompt = int(getattr(raw_usage, "prompt_tokens", 0) or 0)
        completion = int(getattr(raw_usage, "completion_tokens", 0) or 0)
        total = int(getattr(raw_usage, "total_tokens", 0) or 0) or prompt + completion
        cached = 0
        details = getattr(raw_usage, "prompt_tokens_details", None)
        if details is not None:
            cached = int(getattr(details, "cached_tokens", 0) or 0)
        return Usage(
            prompt_tokens=prompt,
            completion_tokens=completion,
            total_tokens=total,
            cached_tokens=cached,
        )

    def chat_completion(
        self,
        messages: list[ChatMessageDict],
        model: str | None = None,
        tier: ModelTier = "pro",
        cache_prefix: int = 0,
    ) -> str:
        model_label = self._resolve(model, tier)
        response = self._client.chat.completions.create(
            model=model_label,
            messages=messages,  # type: ignore[arg-type]
        )
        usage = self._extract_usage(getattr(response, "usage", None))
        logger.info(
            "llm chat_completion: model=%s tier=%s prompt_tokens=%d completion_tokens=%d "
            "total=%d cached=%d",
            model_label,
            tier,
            usage.prompt_tokens,
            usage.completion_tokens,
            usage.total_tokens,
            usage.cached_tokens,
        )
        choice = response.choices[0] if response.choices else None
        return (choice.message.content if choice and choice.message else "") or ""

    def annotate_text(self, prompt: str, content: str, tier: ModelTier = "fast") -> str:
        """注釈生成. デフォルト tier=fast (大量呼び出し前提)."""
        messages: list[ChatMessageDict] = [
            {"role": "system", "content": prompt},
            {"role": "user", "content": content},
        ]
        return self.chat_completion(messages, tier=tier)

    def chat_completion_with_tools(
        self,
        messages: list[ChatMessageDict],
        tools: list[dict[str, Any]],
        model: str | None = None,
        tier: ModelTier = "pro",
        cache_prefix: int = 0,
    ) -> LLMResponse:
        model_label = self._resolve(model, tier)
        kwargs: dict[str, Any] = {
            "model": model_label,
            "messages": messages,
        }
        if tools:
            kwargs["tools"] = tools
        response = self._client.chat.completions.create(**kwargs)  # type: ignore[arg-type]
        usage = self._extract_usage(getattr(response, "usage", None))

        content: str | None = None
        tool_calls: list[LLMToolCall] = []
        if response.choices:
            msg = response.choices[0].message
            content = msg.content
            raw_tool_calls = getattr(msg, "tool_calls", None) or []
            for tc in raw_tool_calls:
                # OpenAI SDK の ChatCompletionMessageToolCall を LLMToolCall に変換
                fn = getattr(tc, "function", None)
                if fn is None:
                    continue
                args_raw = getattr(fn, "arguments", "") or ""
                try:
                    args = json.loads(args_raw) if args_raw else {}
                except json.JSONDecodeError:
                    logger.warning(
                        "tool_call %s has non-JSON arguments: %r",
                        getattr(tc, "id", "?"),
                        args_raw[:120],
                    )
                    args = {}
                tool_calls.append(
                    LLMToolCall(
                        id=getattr(tc, "id", "") or "",
                        name=getattr(fn, "name", "") or "",
                        arguments=args,
                    )
                )

        return LLMResponse(
            content=content,
            tool_calls=tool_calls,
            usage=usage,
            model=model_label,
        )


def _read_env() -> tuple[str | None, str | None, str | None, str | None, str | None]:
    """環境変数を読む.

    Returns:
        (base_url, api_key, model, model_pro, model_fast)
        - model は旧 LLM_MODEL (後方互換用 fallback)
        - model_pro / model_fast が新規 (Task C)
    """
    return (
        os.environ.get("LLM_BASE_URL") or None,
        os.environ.get("LLM_API_KEY") or None,
        os.environ.get("LLM_MODEL") or None,
        os.environ.get("LLM_MODEL_PRO") or None,
        os.environ.get("LLM_MODEL_FAST") or None,
    )


def get_default_client() -> LLMClient:
    """環境変数を見てデフォルトクライアントを返す.

    `LLM_BASE_URL` と `LLM_API_KEY` の両方が設定されていれば
    `OpenAICompatibleLLMClient` を、そうでなければ `MockLLMClient` を返す。

    モデル解決の優先順位:
        pro:   LLM_MODEL_PRO  → LLM_MODEL → デフォルト ("gpt-5.2")
        fast:  LLM_MODEL_FAST → LLM_MODEL_PRO → LLM_MODEL → デフォルト
    LLM_MODEL_FAST が未設定なら pro と同じモデルが使われる (= モデル切替なし)。
    """
    base_url, api_key, model, model_pro, model_fast = _read_env()
    if base_url and api_key:
        pro = model_pro or model or _DEFAULT_PRO_MODEL
        fast = model_fast or pro
        return OpenAICompatibleLLMClient(
            base_url=base_url,
            api_key=api_key,
            pro_model=pro,
            fast_model=fast,
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
