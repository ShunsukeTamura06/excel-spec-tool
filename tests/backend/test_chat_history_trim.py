"""backend.routes.chat 内の履歴トリム/サマリ生成のテスト (Phase B-2)."""

from __future__ import annotations

import pytest

from backend.routes.chat import (
    _get_history_pairs_limit,
    _summarize_old_turns,
    _trim_history,
)
from core.models import ChatMessage


def _msgs(*pairs: tuple[str, str]) -> list[ChatMessage]:
    """簡易ヘルパ: (user_text, assistant_text) のペアから ChatMessage 列を作る."""
    out: list[ChatMessage] = []
    for i, (u, a) in enumerate(pairs):
        out.append(ChatMessage(role="user", content=u, timestamp=f"t{i}"))
        out.append(ChatMessage(role="assistant", content=a, timestamp=f"t{i}"))
    return out


class TestTrimHistory:
    def test_under_limit_returns_all(self) -> None:
        history = _msgs(("q1", "a1"), ("q2", "a2"))
        recent, summary = _trim_history(history, max_pairs=5)
        assert recent == history
        assert summary is None

    def test_exact_limit_returns_all(self) -> None:
        history = _msgs(("q1", "a1"), ("q2", "a2"))
        recent, summary = _trim_history(history, max_pairs=2)
        assert len(recent) == 4
        assert summary is None

    def test_over_limit_keeps_recent_and_summarizes_rest(self) -> None:
        history = _msgs(
            ("old1-q", "old1-a"),
            ("old2-q", "old2-a"),
            ("old3-q", "old3-a"),
            ("recent1-q", "recent1-a"),
            ("recent2-q", "recent2-a"),
        )
        recent, summary = _trim_history(history, max_pairs=2)
        assert len(recent) == 4
        # 残ったのは直近 2 ペアのみ
        assert recent[0].content == "recent1-q"
        assert recent[-1].content == "recent2-a"
        # サマリには古い 3 ペア (6 メッセージ) の情報が入っている
        assert summary is not None
        assert "old1-q" in summary
        assert "old3-a" in summary
        # 直近メッセージはサマリに入らない
        assert "recent1-q" not in summary

    def test_max_pairs_zero_disables_trim(self) -> None:
        history = _msgs(*[(f"q{i}", f"a{i}") for i in range(20)])
        recent, summary = _trim_history(history, max_pairs=0)
        assert recent == history
        assert summary is None

    def test_max_pairs_negative_disables_trim(self) -> None:
        history = _msgs(("q1", "a1"))
        recent, summary = _trim_history(history, max_pairs=-1)
        assert recent == history
        assert summary is None

    def test_empty_history(self) -> None:
        recent, summary = _trim_history([], max_pairs=5)
        assert recent == []
        assert summary is None


class TestSummarizeOldTurns:
    def test_empty_returns_empty_string(self) -> None:
        assert _summarize_old_turns([]) == ""

    def test_includes_role_and_content(self) -> None:
        msgs = [
            ChatMessage(role="user", content="what is X?", timestamp="t1"),
            ChatMessage(role="assistant", content="X is foo", timestamp="t1"),
        ]
        s = _summarize_old_turns(msgs)
        assert "user:" in s
        assert "assistant:" in s
        assert "what is X?" in s
        assert "X is foo" in s
        assert "2 メッセージ" in s

    def test_truncates_long_content(self) -> None:
        long_text = "x" * 500
        msgs = [ChatMessage(role="user", content=long_text, timestamp="t1")]
        s = _summarize_old_turns(msgs)
        # 切り詰め記号 "…" が入る
        assert "…" in s
        # 全文は入らない (500 文字は超えない)
        assert long_text not in s

    def test_newlines_in_content_collapsed(self) -> None:
        msgs = [
            ChatMessage(role="user", content="line1\nline2\nline3", timestamp="t1"),
        ]
        s = _summarize_old_turns(msgs)
        # 各メッセージは 1 行になる
        assert "line1 line2 line3" in s


class TestHistoryLimitEnv:
    def test_default_when_unset(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.delenv("CHAT_HISTORY_LIMIT_PAIRS", raising=False)
        assert _get_history_pairs_limit() == 10

    def test_env_override(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("CHAT_HISTORY_LIMIT_PAIRS", "3")
        assert _get_history_pairs_limit() == 3

    def test_invalid_value_falls_back_to_default(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("CHAT_HISTORY_LIMIT_PAIRS", "not-a-number")
        assert _get_history_pairs_limit() == 10

    def test_zero_value_passes_through(self, monkeypatch: pytest.MonkeyPatch) -> None:
        # 0 はそのまま返る (_trim_history 側でトリム無効化として扱う)
        monkeypatch.setenv("CHAT_HISTORY_LIMIT_PAIRS", "0")
        assert _get_history_pairs_limit() == 0
