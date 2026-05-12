"""backend.logging_config のテスト."""

from __future__ import annotations

import asyncio
import logging
from pathlib import Path

import pytest

from backend.logging_config import (
    JobIdFilter,
    JobIdLoggingMiddleware,
    configure_logging,
    current_job_id,
    set_job_id,
)


class TestJobIdFilter:
    def test_no_job_id_uses_dash(self) -> None:
        record = logging.LogRecord(
            name="t",
            level=logging.INFO,
            pathname=__file__,
            lineno=1,
            msg="hi",
            args=(),
            exc_info=None,
        )
        JobIdFilter().filter(record)
        assert record.job_id == "-"

    def test_within_set_job_id(self) -> None:
        record = logging.LogRecord(
            name="t",
            level=logging.INFO,
            pathname=__file__,
            lineno=1,
            msg="hi",
            args=(),
            exc_info=None,
        )
        with set_job_id("job-123"):
            JobIdFilter().filter(record)
        assert record.job_id == "job-123"

    def test_set_job_id_cleared_after_block(self) -> None:
        with set_job_id("inside"):
            assert current_job_id() == "inside"
        assert current_job_id() is None

    def test_set_job_id_nested(self) -> None:
        with set_job_id("outer"):
            assert current_job_id() == "outer"
            with set_job_id("inner"):
                assert current_job_id() == "inner"
            assert current_job_id() == "outer"


class TestConfigureLogging:
    def test_respects_log_level(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("LOG_LEVEL", "WARNING")
        monkeypatch.delenv("LOG_FILE", raising=False)
        configure_logging()
        assert logging.getLogger().level == logging.WARNING

    def test_unknown_level_falls_back_to_info(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("LOG_LEVEL", "BOGUS")
        monkeypatch.delenv("LOG_FILE", raising=False)
        configure_logging()
        assert logging.getLogger().level == logging.INFO

    def test_log_file_writes_to_file(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        log_path = tmp_path / "out.log"
        monkeypatch.setenv("LOG_LEVEL", "INFO")
        monkeypatch.setenv("LOG_FILE", str(log_path))
        configure_logging()
        with set_job_id("abc-999"):
            logging.getLogger("test").info("hello world")
        # FileHandler はバッファリングするので明示的に flush
        for h in logging.getLogger().handlers:
            h.flush()
        contents = log_path.read_text(encoding="utf-8")
        assert "hello world" in contents
        assert "abc-999" in contents
        assert "INFO" in contents

    def test_idempotent_no_duplicate_handlers(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.delenv("LOG_FILE", raising=False)
        configure_logging()
        first = len(logging.getLogger().handlers)
        configure_logging()
        configure_logging()
        assert len(logging.getLogger().handlers) == first


class TestJobIdLoggingMiddleware:
    def test_extracts_job_id_from_path(self) -> None:
        seen: dict[str, str | None] = {}

        async def app(scope: dict, receive, send) -> None:  # type: ignore[no-untyped-def]
            seen["job_id"] = current_job_id()

        mw = JobIdLoggingMiddleware(app)  # type: ignore[arg-type]
        scope = {
            "type": "http",
            "path": "/chat/12345678-1234-4234-8234-1234567890ab",
        }

        async def receive() -> dict:  # type: ignore[type-arg]
            return {"type": "http.request"}

        async def send(msg: dict) -> None:  # type: ignore[type-arg]
            return None

        asyncio.run(mw(scope, receive, send))
        assert seen["job_id"] == "12345678-1234-4234-8234-1234567890ab"
        # ミドルウェア外では戻っている
        assert current_job_id() is None

    def test_no_match_leaves_context_clear(self) -> None:
        seen: dict[str, str | None] = {}

        async def app(scope: dict, receive, send) -> None:  # type: ignore[no-untyped-def]
            seen["job_id"] = current_job_id()

        mw = JobIdLoggingMiddleware(app)  # type: ignore[arg-type]
        scope = {"type": "http", "path": "/health"}

        async def receive() -> dict:  # type: ignore[type-arg]
            return {"type": "http.request"}

        async def send(msg: dict) -> None:  # type: ignore[type-arg]
            return None

        asyncio.run(mw(scope, receive, send))
        assert seen["job_id"] is None

    def test_non_http_passthrough(self) -> None:
        called = {"hit": False}

        async def app(scope: dict, receive, send) -> None:  # type: ignore[no-untyped-def]
            called["hit"] = True

        mw = JobIdLoggingMiddleware(app)  # type: ignore[arg-type]

        async def receive() -> dict:  # type: ignore[type-arg]
            return {}

        async def send(msg: dict) -> None:  # type: ignore[type-arg]
            return None

        asyncio.run(mw({"type": "lifespan"}, receive, send))
        assert called["hit"] is True
