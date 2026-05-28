"""backend.llm_client のテスト."""

from typing import Any
from unittest.mock import MagicMock, patch

import pytest

from backend import llm_client
from backend.llm_client import (
    LLMClient,
    MockLLMClient,
    OpenAICompatibleLLMClient,
    Usage,
    annotate_text,
    chat_completion,
    get_default_client,
)


def _fake_openai_response(
    content: str | None = None,
    tool_calls: list[Any] | None = None,
    prompt_tokens: int = 10,
    completion_tokens: int = 5,
    cached_tokens: int = 0,
) -> MagicMock:
    """openai SDK の ChatCompletion 応答を模す MagicMock を返す."""
    msg = MagicMock()
    msg.content = content
    msg.tool_calls = tool_calls or None

    choice = MagicMock()
    choice.message = msg

    resp = MagicMock()
    resp.choices = [choice]

    usage = MagicMock()
    usage.prompt_tokens = prompt_tokens
    usage.completion_tokens = completion_tokens
    usage.total_tokens = prompt_tokens + completion_tokens
    details = MagicMock()
    details.cached_tokens = cached_tokens
    usage.prompt_tokens_details = details
    resp.usage = usage
    return resp


def _fake_tool_call(call_id: str, name: str, arguments_json: str) -> MagicMock:
    """ChatCompletionMessageToolCall を模す."""
    tc = MagicMock()
    tc.id = call_id
    fn = MagicMock()
    fn.name = name
    fn.arguments = arguments_json
    tc.function = fn
    return tc


def _make_openai_client(
    pro_model: str = "pro",
    fast_model: str = "fast",
) -> tuple[OpenAICompatibleLLMClient, MagicMock]:
    """OpenAI SDK の OpenAI クラスを差し替えた client を返す.

    Returns:
        (client, fake_sdk) — fake_sdk.chat.completions.create が MagicMock.
    """
    with patch("openai.OpenAI") as openai_cls:
        fake = MagicMock()
        openai_cls.return_value = fake
        client = OpenAICompatibleLLMClient(
            base_url="http://x", api_key="k", pro_model=pro_model, fast_model=fast_model
        )
    return client, fake


# ---------- Protocol 適合 ----------


class TestProtocolConformance:
    def test_mock_implements_protocol(self) -> None:
        assert isinstance(MockLLMClient(), LLMClient)

    def test_openai_implements_protocol(self) -> None:
        client, _ = _make_openai_client(pro_model="m")
        assert isinstance(client, LLMClient)


# ---------- MockLLMClient ----------


class TestMockClient:
    def test_chat_completion_echoes_last_user_message(self) -> None:
        client = MockLLMClient()
        result = client.chat_completion(
            [
                {"role": "system", "content": "You are an assistant."},
                {"role": "user", "content": "Hello"},
                {"role": "assistant", "content": "Hi"},
                {"role": "user", "content": "How are you?"},
            ]
        )
        assert "How are you?" in result
        assert "[mock" in result

    def test_chat_completion_with_no_user_message(self) -> None:
        result = MockLLMClient().chat_completion([{"role": "system", "content": "system only"}])
        assert isinstance(result, str)
        # 最後のユーザー発話が空でもクラッシュしない
        assert "[mock" in result

    def test_chat_completion_includes_model_label(self) -> None:
        result = MockLLMClient().chat_completion(
            [{"role": "user", "content": "hi"}], model="gpt-test"
        )
        assert "gpt-test" in result

    def test_annotate_text_returns_empty_json(self) -> None:
        # LLM 未設定時のデフォルトとして使われるため、annotators._safe_annotate_json が
        # JSON パースして空辞書 → 注釈スキップになるよう "{}" を返す。
        # 以前は prompt の raw 文字列を返していたが、それが設計書 / ダイアグラム
        # ノードに leak していたため変更した。
        result = MockLLMClient().annotate_text(
            "summarize this VBA module", "Sub Hello()\nMsgBox 'hi'\nEnd Sub"
        )
        assert result == "{}"

    def test_annotate_empty_content(self) -> None:
        # 空の content でも例外を出さない
        result = MockLLMClient().annotate_text("any prompt", "")
        assert result == "{}"


# ---------- OpenAICompatibleLLMClient ----------


class TestOpenAICompatibleClient:
    def test_construction_records_attributes(self) -> None:
        client, _ = _make_openai_client(pro_model="gpt-5.2", fast_model="gpt-5.2-mini")
        assert client.base_url == "http://x"
        assert client.api_key == "k"
        assert client.default_model == "gpt-5.2"
        assert client.fast_model == "gpt-5.2-mini"

    def test_chat_completion_returns_content(self) -> None:
        client, fake = _make_openai_client()
        fake.chat.completions.create.return_value = _fake_openai_response(content="hello")
        result = client.chat_completion([{"role": "user", "content": "hi"}])
        assert result == "hello"
        # 正しく pro_model でリクエストされた
        call = fake.chat.completions.create.call_args
        assert call.kwargs["model"] == "pro"

    def test_chat_completion_tier_fast_uses_fast_model(self) -> None:
        client, fake = _make_openai_client()
        fake.chat.completions.create.return_value = _fake_openai_response(content="x")
        client.chat_completion([{"role": "user", "content": "hi"}], tier="fast")
        assert fake.chat.completions.create.call_args.kwargs["model"] == "fast"

    def test_chat_completion_explicit_model_overrides(self) -> None:
        client, fake = _make_openai_client()
        fake.chat.completions.create.return_value = _fake_openai_response(content="x")
        client.chat_completion([{"role": "user", "content": "hi"}], model="explicit", tier="fast")
        assert fake.chat.completions.create.call_args.kwargs["model"] == "explicit"

    def test_chat_completion_empty_choices(self) -> None:
        client, fake = _make_openai_client()
        empty = MagicMock()
        empty.choices = []
        empty.usage = None
        fake.chat.completions.create.return_value = empty
        assert client.chat_completion([{"role": "user", "content": "hi"}]) == ""

    def test_annotate_text_uses_fast_by_default(self) -> None:
        client, fake = _make_openai_client()
        fake.chat.completions.create.return_value = _fake_openai_response(content="note")
        result = client.annotate_text("describe this", "code body")
        assert result == "note"
        msgs = fake.chat.completions.create.call_args.kwargs["messages"]
        assert msgs[0]["role"] == "system"
        assert "describe this" in msgs[0]["content"]
        assert msgs[1]["role"] == "user"
        assert "code body" in msgs[1]["content"]
        # fast モデルが使われた
        assert fake.chat.completions.create.call_args.kwargs["model"] == "fast"

    def test_chat_with_tools_parses_tool_calls(self) -> None:
        client, fake = _make_openai_client()
        tc = _fake_tool_call("call_1", "find_cells", '{"query": "x"}')
        fake.chat.completions.create.return_value = _fake_openai_response(
            content=None, tool_calls=[tc]
        )
        resp = client.chat_completion_with_tools(
            [{"role": "user", "content": "q"}],
            tools=[{"type": "function", "function": {"name": "find_cells"}}],
        )
        assert resp.content is None
        assert len(resp.tool_calls) == 1
        assert resp.tool_calls[0].id == "call_1"
        assert resp.tool_calls[0].name == "find_cells"
        assert resp.tool_calls[0].arguments == {"query": "x"}
        assert resp.model == "pro"
        # tools パラメタが SDK に渡されている
        assert "tools" in fake.chat.completions.create.call_args.kwargs

    def test_chat_with_tools_invalid_json_args(self) -> None:
        client, fake = _make_openai_client()
        tc = _fake_tool_call("call_1", "find_cells", "not valid json")
        fake.chat.completions.create.return_value = _fake_openai_response(
            content=None, tool_calls=[tc]
        )
        resp = client.chat_completion_with_tools([{"role": "user", "content": "q"}], tools=[])
        # 不正 JSON は空 dict にフォールバック、クラッシュしない
        assert resp.tool_calls[0].arguments == {}

    def test_chat_with_tools_empty_args_string(self) -> None:
        client, fake = _make_openai_client()
        tc = _fake_tool_call("call_1", "noop", "")
        fake.chat.completions.create.return_value = _fake_openai_response(
            content=None, tool_calls=[tc]
        )
        resp = client.chat_completion_with_tools([{"role": "user", "content": "q"}], tools=[])
        assert resp.tool_calls[0].arguments == {}

    def test_usage_extracted_with_cache(self) -> None:
        client, fake = _make_openai_client()
        fake.chat.completions.create.return_value = _fake_openai_response(
            content="ok", prompt_tokens=100, completion_tokens=20, cached_tokens=80
        )
        resp = client.chat_completion_with_tools([{"role": "user", "content": "q"}], tools=[])
        assert resp.usage is not None
        assert resp.usage.prompt_tokens == 100
        assert resp.usage.completion_tokens == 20
        assert resp.usage.total_tokens == 120
        assert resp.usage.cached_tokens == 80

    def test_usage_missing_handled_gracefully(self) -> None:
        client, fake = _make_openai_client()
        resp_mock = MagicMock()
        resp_mock.choices = [MagicMock()]
        resp_mock.choices[0].message.content = "ok"
        resp_mock.choices[0].message.tool_calls = None
        resp_mock.usage = None
        fake.chat.completions.create.return_value = resp_mock
        resp = client.chat_completion_with_tools([{"role": "user", "content": "q"}], tools=[])
        assert resp.usage is not None
        assert resp.usage.total_tokens == 0


class TestExtractUsage:
    """OpenAICompatibleLLMClient._extract_usage の細部."""

    def test_handles_none(self) -> None:
        client, _ = _make_openai_client()
        u = client._extract_usage(None)
        assert u == Usage()

    def test_recomputes_total_when_missing(self) -> None:
        client, _ = _make_openai_client()
        raw = MagicMock()
        raw.prompt_tokens = 50
        raw.completion_tokens = 10
        raw.total_tokens = 0  # 未提供
        raw.prompt_tokens_details = None
        u = client._extract_usage(raw)
        assert u.total_tokens == 60


# ---------- get_default_client (env-based factory) ----------


class TestGetDefaultClient:
    def test_no_env_returns_mock(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.delenv("LLM_BASE_URL", raising=False)
        monkeypatch.delenv("LLM_API_KEY", raising=False)
        monkeypatch.delenv("LLM_MODEL", raising=False)
        client = get_default_client()
        assert isinstance(client, MockLLMClient)

    def test_only_base_url_returns_mock(self, monkeypatch: pytest.MonkeyPatch) -> None:
        # api_key が無ければモック (実 API には接続できない)
        monkeypatch.setenv("LLM_BASE_URL", "http://internal")
        monkeypatch.delenv("LLM_API_KEY", raising=False)
        client = get_default_client()
        assert isinstance(client, MockLLMClient)

    def test_only_api_key_returns_mock(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.delenv("LLM_BASE_URL", raising=False)
        monkeypatch.setenv("LLM_API_KEY", "k")
        client = get_default_client()
        assert isinstance(client, MockLLMClient)

    def test_both_set_returns_openai_client(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("LLM_BASE_URL", "http://internal-llm/v1")
        monkeypatch.setenv("LLM_API_KEY", "secret")
        monkeypatch.setenv("LLM_MODEL", "gpt-5.2")
        client = get_default_client()
        assert isinstance(client, OpenAICompatibleLLMClient)
        assert client.default_model == "gpt-5.2"

    def test_default_model_when_not_set(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("LLM_BASE_URL", "http://x")
        monkeypatch.setenv("LLM_API_KEY", "k")
        monkeypatch.delenv("LLM_MODEL", raising=False)
        client = get_default_client()
        assert isinstance(client, OpenAICompatibleLLMClient)
        assert client.default_model == "gpt-5.2"

    def test_empty_env_treated_as_unset(self, monkeypatch: pytest.MonkeyPatch) -> None:
        # 空文字列は未設定と同じ扱い
        monkeypatch.setenv("LLM_BASE_URL", "")
        monkeypatch.setenv("LLM_API_KEY", "")
        client = get_default_client()
        assert isinstance(client, MockLLMClient)


# ---------- module-level convenience functions ----------


class TestModuleLevelFunctions:
    def test_chat_completion_uses_default_client(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.delenv("LLM_BASE_URL", raising=False)
        monkeypatch.delenv("LLM_API_KEY", raising=False)
        result = chat_completion([{"role": "user", "content": "hi"}])
        assert "[mock" in result

    def test_annotate_text_uses_default_client(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.delenv("LLM_BASE_URL", raising=False)
        monkeypatch.delenv("LLM_API_KEY", raising=False)
        result = annotate_text("describe", "hello")
        # mock は空 JSON を返す (raw prompt leak 回避のため)
        assert result == "{}"

    def test_real_client_invokes_openai_sdk(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("LLM_BASE_URL", "http://x")
        monkeypatch.setenv("LLM_API_KEY", "k")
        with patch("openai.OpenAI") as openai_cls:
            fake = MagicMock()
            fake.chat.completions.create.return_value = _fake_openai_response(content="real call")
            openai_cls.return_value = fake
            result = chat_completion([{"role": "user", "content": "hi"}])
        assert result == "real call"


# ---------- module exports ----------


class TestModuleSurface:
    def test_public_symbols_exported(self) -> None:
        # SPEC §5.3 が要求する公開関数
        assert callable(llm_client.chat_completion)
        assert callable(llm_client.annotate_text)


# ---------- function calling 拡張 ----------


from backend.llm_client import LLMResponse, LLMToolCall  # noqa: E402


class TestMockToolCalling:
    def test_default_returns_content_only(self) -> None:
        c = MockLLMClient()
        resp = c.chat_completion_with_tools([{"role": "user", "content": "hi"}], tools=[])
        assert isinstance(resp, LLMResponse)
        assert resp.content is not None
        assert resp.tool_calls == []

    def test_queued_response_with_tool_call(self) -> None:
        c = MockLLMClient()
        c.queue_response(
            LLMResponse(
                content=None,
                tool_calls=[LLMToolCall(id="t1", name="get_cells_range", arguments={"sheet": "S"})],
            )
        )
        c.queue_response(LLMResponse(content="final answer"))

        first = c.chat_completion_with_tools([{"role": "user", "content": "x"}], tools=[])
        assert first.tool_calls
        assert first.tool_calls[0].name == "get_cells_range"

        second = c.chat_completion_with_tools([{"role": "user", "content": "x"}], tools=[])
        assert second.content == "final answer"
        assert second.tool_calls == []


class TestOpenAICompatibleToolCalling:
    def test_uses_pro_by_default(self) -> None:
        c, fake = _make_openai_client()
        fake.chat.completions.create.return_value = _fake_openai_response(content="ok")
        resp = c.chat_completion_with_tools([{"role": "user", "content": "x"}], tools=[])
        assert resp.model == "pro"
        assert fake.chat.completions.create.call_args.kwargs["model"] == "pro"

    def test_tier_fast_uses_fast_model(self) -> None:
        c, fake = _make_openai_client()
        fake.chat.completions.create.return_value = _fake_openai_response(content="ok")
        resp = c.chat_completion_with_tools(
            [{"role": "user", "content": "x"}], tools=[], tier="fast"
        )
        assert resp.model == "fast"


# ---------- Usage / model 観測 ----------


# Usage は冒頭で import 済み


class TestUsage:
    def test_zero_addition(self) -> None:
        a = Usage(prompt_tokens=10, completion_tokens=5, total_tokens=15, cached_tokens=2)
        b = Usage(prompt_tokens=4, completion_tokens=1, total_tokens=5, cached_tokens=0)
        c = a + b
        assert c.prompt_tokens == 14
        assert c.completion_tokens == 6
        assert c.total_tokens == 20
        assert c.cached_tokens == 2


class TestMockUsagePropagation:
    def test_default_response_has_usage(self) -> None:
        resp = MockLLMClient().chat_completion_with_tools(
            [{"role": "user", "content": "hello"}], tools=[]
        )
        assert resp.usage is not None
        assert resp.usage.total_tokens >= 1
        assert resp.usage.prompt_tokens >= 1
        # tier=pro デフォルトなので pro_model に解決される
        assert resp.model == "mock-pro"

    def test_default_response_passes_model_label(self) -> None:
        resp = MockLLMClient().chat_completion_with_tools(
            [{"role": "user", "content": "hi"}], tools=[], model="gpt-test"
        )
        assert resp.model == "gpt-test"

    def test_queued_response_without_usage_is_filled(self) -> None:
        c = MockLLMClient()
        # usage 未指定で queue
        c.queue_response(LLMResponse(content="answer"))
        resp = c.chat_completion_with_tools(
            [{"role": "user", "content": "a longer prompt that should generate tokens"}],
            tools=[],
        )
        assert resp.usage is not None
        assert resp.usage.total_tokens >= 1
        assert resp.model == "mock-pro"

    def test_queued_response_preserves_explicit_usage(self) -> None:
        c = MockLLMClient()
        c.queue_response(
            LLMResponse(
                content="x",
                usage=Usage(prompt_tokens=100, completion_tokens=50, total_tokens=150),
                model="pro-model",
            )
        )
        resp = c.chat_completion_with_tools([{"role": "user", "content": "q"}], tools=[])
        # 明示された usage / model は上書きされない
        assert resp.usage is not None
        assert resp.usage.prompt_tokens == 100
        assert resp.usage.completion_tokens == 50
        assert resp.model == "pro-model"


# ---------- Task C: モデル tier (pro / fast) ----------


class TestMockTierResolution:
    def test_default_tier_is_pro(self) -> None:
        c = MockLLMClient()
        resp = c.chat_completion_with_tools([{"role": "user", "content": "x"}], tools=[])
        assert resp.model == "mock-pro"
        assert c.calls[-1]["tier"] == "pro"

    def test_fast_tier_resolves_to_fast_model(self) -> None:
        c = MockLLMClient()
        resp = c.chat_completion_with_tools(
            [{"role": "user", "content": "x"}], tools=[], tier="fast"
        )
        assert resp.model == "mock-fast"
        assert c.calls[-1]["tier"] == "fast"

    def test_explicit_model_overrides_tier(self) -> None:
        c = MockLLMClient()
        resp = c.chat_completion_with_tools(
            [{"role": "user", "content": "x"}], tools=[], model="custom", tier="fast"
        )
        assert resp.model == "custom"

    def test_custom_tier_models(self) -> None:
        c = MockLLMClient(pro_model="big", fast_model="small")
        r1 = c.chat_completion_with_tools([{"role": "user", "content": "x"}], tools=[], tier="pro")
        r2 = c.chat_completion_with_tools([{"role": "user", "content": "x"}], tools=[], tier="fast")
        assert r1.model == "big"
        assert r2.model == "small"


class TestOpenAITierResolution:
    def test_constructor_stores_pro_and_fast(self) -> None:
        c = OpenAICompatibleLLMClient(
            base_url="http://x", api_key="k", pro_model="big", fast_model="small"
        )
        assert c.pro_model == "big"
        assert c.fast_model == "small"

    def test_fast_defaults_to_pro_when_not_set(self) -> None:
        c = OpenAICompatibleLLMClient(base_url="http://x", api_key="k", pro_model="only")
        assert c.fast_model == "only"

    def test_resolve_uses_tier(self) -> None:
        c = OpenAICompatibleLLMClient(
            base_url="http://x", api_key="k", pro_model="big", fast_model="small"
        )
        assert c._resolve(None, "pro") == "big"
        assert c._resolve(None, "fast") == "small"
        assert c._resolve("custom", "fast") == "custom"


class TestMockCacheSimulation:
    """Phase B-3: MockLLMClient のキャッシュ命中シミュレーション."""

    def test_no_cache_when_prefix_zero(self) -> None:
        c = MockLLMClient()
        msgs = [{"role": "system", "content": "stable system"}, {"role": "user", "content": "q"}]
        r1 = c.chat_completion_with_tools(msgs, tools=[], cache_prefix=0)
        r2 = c.chat_completion_with_tools(msgs, tools=[], cache_prefix=0)
        assert r1.usage is not None and r1.usage.cached_tokens == 0
        assert r2.usage is not None and r2.usage.cached_tokens == 0

    def test_first_call_miss_second_call_hit(self) -> None:
        c = MockLLMClient()
        system = "a" * 400  # 同じ system を 2 回送る
        msgs1 = [{"role": "system", "content": system}, {"role": "user", "content": "q1"}]
        msgs2 = [{"role": "system", "content": system}, {"role": "user", "content": "q2"}]
        r1 = c.chat_completion_with_tools(msgs1, tools=[], cache_prefix=1)
        r2 = c.chat_completion_with_tools(msgs2, tools=[], cache_prefix=1)
        assert r1.usage is not None and r1.usage.cached_tokens == 0  # 初回は miss
        assert r2.usage is not None and r2.usage.cached_tokens > 0  # 2回目は hit

    def test_cache_miss_when_prefix_changes(self) -> None:
        c = MockLLMClient()
        msgs1 = [{"role": "system", "content": "version 1"}, {"role": "user", "content": "q"}]
        msgs2 = [{"role": "system", "content": "version 2"}, {"role": "user", "content": "q"}]
        c.chat_completion_with_tools(msgs1, tools=[], cache_prefix=1)
        r2 = c.chat_completion_with_tools(msgs2, tools=[], cache_prefix=1)
        assert r2.usage is not None and r2.usage.cached_tokens == 0

    def test_cache_isolated_per_model(self) -> None:
        """tier 切替 (= model 切替) はキャッシュを共有しない."""
        c = MockLLMClient()
        msgs = [{"role": "system", "content": "stable"}, {"role": "user", "content": "q"}]
        # pro でキャッシュをシード
        c.chat_completion_with_tools(msgs, tools=[], cache_prefix=1, tier="pro")
        # 同じ内容で fast を呼ぶ → 別 model なのでキャッシュ無し
        r_fast = c.chat_completion_with_tools(msgs, tools=[], cache_prefix=1, tier="fast")
        assert r_fast.usage is not None and r_fast.usage.cached_tokens == 0
        # pro を再度呼ぶ → ヒット
        r_pro2 = c.chat_completion_with_tools(msgs, tools=[], cache_prefix=1, tier="pro")
        assert r_pro2.usage is not None and r_pro2.usage.cached_tokens > 0

    def test_only_first_n_messages_count(self) -> None:
        """cache_prefix=1 なら 2 つ目以降のメッセージが変わってもキャッシュ命中する."""
        c = MockLLMClient()
        sys_msg = {"role": "system", "content": "stable"}
        msgs1 = [sys_msg, {"role": "user", "content": "first"}]
        msgs2 = [sys_msg, {"role": "user", "content": "completely different"}]
        c.chat_completion_with_tools(msgs1, tools=[], cache_prefix=1)
        r2 = c.chat_completion_with_tools(msgs2, tools=[], cache_prefix=1)
        # user 部分が変わっても system が同じなのでヒット
        assert r2.usage is not None and r2.usage.cached_tokens > 0


class TestGetDefaultClientTier:
    def test_pro_fast_explicit_envs(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("LLM_BASE_URL", "http://x")
        monkeypatch.setenv("LLM_API_KEY", "k")
        monkeypatch.setenv("LLM_MODEL_PRO", "big-model")
        monkeypatch.setenv("LLM_MODEL_FAST", "small-model")
        monkeypatch.delenv("LLM_MODEL", raising=False)
        c = get_default_client()
        assert isinstance(c, OpenAICompatibleLLMClient)
        assert c.pro_model == "big-model"
        assert c.fast_model == "small-model"

    def test_fast_falls_back_to_pro(self, monkeypatch: pytest.MonkeyPatch) -> None:
        # LLM_MODEL_FAST 未設定なら fast = pro
        monkeypatch.setenv("LLM_BASE_URL", "http://x")
        monkeypatch.setenv("LLM_API_KEY", "k")
        monkeypatch.setenv("LLM_MODEL_PRO", "big")
        monkeypatch.delenv("LLM_MODEL_FAST", raising=False)
        monkeypatch.delenv("LLM_MODEL", raising=False)
        c = get_default_client()
        assert isinstance(c, OpenAICompatibleLLMClient)
        assert c.pro_model == "big"
        assert c.fast_model == "big"

    def test_legacy_llm_model_only(self, monkeypatch: pytest.MonkeyPatch) -> None:
        # LLM_MODEL のみ設定 (旧スタイル): pro/fast 両方に展開
        monkeypatch.setenv("LLM_BASE_URL", "http://x")
        monkeypatch.setenv("LLM_API_KEY", "k")
        monkeypatch.setenv("LLM_MODEL", "legacy")
        monkeypatch.delenv("LLM_MODEL_PRO", raising=False)
        monkeypatch.delenv("LLM_MODEL_FAST", raising=False)
        c = get_default_client()
        assert isinstance(c, OpenAICompatibleLLMClient)
        assert c.pro_model == "legacy"
        assert c.fast_model == "legacy"

    def test_pro_overrides_legacy_llm_model(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("LLM_BASE_URL", "http://x")
        monkeypatch.setenv("LLM_API_KEY", "k")
        monkeypatch.setenv("LLM_MODEL", "legacy")
        monkeypatch.setenv("LLM_MODEL_PRO", "new-pro")
        monkeypatch.delenv("LLM_MODEL_FAST", raising=False)
        c = get_default_client()
        assert isinstance(c, OpenAICompatibleLLMClient)
        assert c.pro_model == "new-pro"
        # fast は LLM_MODEL_PRO の値 (legacy ではなく) にフォールバック
        assert c.fast_model == "new-pro"
