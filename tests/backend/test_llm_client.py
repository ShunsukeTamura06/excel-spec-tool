"""backend.llm_client のテスト."""

import pytest

from backend import llm_client
from backend.llm_client import (
    LLMClient,
    MockLLMClient,
    OpenAICompatibleLLMClient,
    annotate_text,
    chat_completion,
    get_default_client,
)

# ---------- Protocol 適合 ----------


class TestProtocolConformance:
    def test_mock_implements_protocol(self) -> None:
        assert isinstance(MockLLMClient(), LLMClient)

    def test_openai_implements_protocol(self) -> None:
        client = OpenAICompatibleLLMClient(base_url="http://x", api_key="k", default_model="m")
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

    def test_annotate_text(self) -> None:
        result = MockLLMClient().annotate_text(
            "summarize this VBA module", "Sub Hello()\nMsgBox 'hi'\nEnd Sub"
        )
        assert "[mock annotation]" in result
        assert "summarize" in result

    def test_annotate_empty_content(self) -> None:
        # 空の content でも例外を出さない
        result = MockLLMClient().annotate_text("any prompt", "")
        assert isinstance(result, str)


# ---------- OpenAICompatibleLLMClient ----------


class TestOpenAICompatibleClientPlaceholder:
    def test_chat_completion_raises_not_implemented(self) -> None:
        client = OpenAICompatibleLLMClient(base_url="http://x", api_key="k", default_model="m")
        with pytest.raises(NotImplementedError):
            client.chat_completion([{"role": "user", "content": "hi"}])

    def test_annotate_text_raises_not_implemented(self) -> None:
        client = OpenAICompatibleLLMClient(base_url="http://x", api_key="k", default_model="m")
        with pytest.raises(NotImplementedError):
            client.annotate_text("p", "c")

    def test_construction_records_attributes(self) -> None:
        client = OpenAICompatibleLLMClient(
            base_url="http://internal-llm/v1",
            api_key="secret",
            default_model="gpt-5.2",
        )
        assert client.base_url == "http://internal-llm/v1"
        assert client.api_key == "secret"
        assert client.default_model == "gpt-5.2"


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
        assert "[mock annotation]" in result

    def test_real_client_propagates_not_implemented(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("LLM_BASE_URL", "http://x")
        monkeypatch.setenv("LLM_API_KEY", "k")
        with pytest.raises(NotImplementedError):
            chat_completion([{"role": "user", "content": "hi"}])


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
    def test_raises_not_implemented(self) -> None:
        c = OpenAICompatibleLLMClient(base_url="http://x", api_key="k", default_model="m")
        with pytest.raises(NotImplementedError):
            c.chat_completion_with_tools([{"role": "user", "content": "x"}], tools=[])
