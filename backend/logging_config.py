"""バックエンドのロギング設定モジュール.

`configure_logging()` で root logger をセットアップする。
ジョブを跨いだリクエスト相関を取るために `job_id` を contextvar として持ち、
`set_job_id()` ブロックや `JobIdLoggingMiddleware` で設定された値が
全ログレコードに自動的に注入される。

環境変数:
    LOG_LEVEL: ログレベル (DEBUG/INFO/WARNING/ERROR, デフォルト INFO)
    LOG_FILE:  出力先ファイル. 指定しなければ stderr.

利用例:
    >>> with set_job_id("abc-..."):
    ...     logger.info("extracted workbook")  # ログに job_id が付与される
"""

from __future__ import annotations

import logging
import os
import re
import sys
from collections.abc import Iterator
from contextlib import contextmanager
from contextvars import ContextVar

from starlette.types import ASGIApp, Receive, Scope, Send

# 現在のリクエストに紐づく job_id. 未設定なら None.
_job_id_var: ContextVar[str | None] = ContextVar("job_id", default=None)


class JobIdFilter(logging.Filter):
    """LogRecord に現在の job_id を `record.job_id` として注入するフィルタ.

    contextvar が未設定の場合は `-` を入れる (フォーマット時に空欄にならないため)。
    """

    def filter(self, record: logging.LogRecord) -> bool:
        record.job_id = _job_id_var.get() or "-"
        return True


_LOG_FORMAT = "%(asctime)s %(levelname)-7s [%(job_id)s] %(name)s: %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"


def configure_logging() -> None:
    """環境変数を見て root logger を構成する.

    idempotent: 既存のハンドラは除去してから付け直すので、複数回呼び出しても
    ハンドラが重複しない (テストで TestClient 経由に複数回 create_app される
    ケースを想定)。
    """
    level_name = os.environ.get("LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)
    log_file = os.environ.get("LOG_FILE") or None

    root = logging.getLogger()
    root.setLevel(level)

    for h in list(root.handlers):
        root.removeHandler(h)

    handler: logging.Handler
    if log_file:
        handler = logging.FileHandler(log_file, encoding="utf-8")
    else:
        handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter(_LOG_FORMAT, datefmt=_DATE_FORMAT))
    handler.addFilter(JobIdFilter())
    root.addHandler(handler)

    # uvicorn / fastapi は独自 logger を持つので root のフォーマットを効かせる
    for name in ("uvicorn", "uvicorn.error", "uvicorn.access", "fastapi"):
        lg = logging.getLogger(name)
        lg.handlers = []
        lg.propagate = True
        lg.setLevel(level)


@contextmanager
def set_job_id(job_id: str | None) -> Iterator[None]:
    """ブロック内のログに job_id を相関 ID として付与する.

    `extract` のようにジョブ作成「後」に相関を確立したい箇所で使う。
    通常の `/{job_id}` パスを持つ route は `JobIdLoggingMiddleware` が
    自動で設定するので明示呼び出しは不要。
    """
    token = _job_id_var.set(job_id)
    try:
        yield
    finally:
        _job_id_var.reset(token)


def current_job_id() -> str | None:
    """現在の contextvar 上の job_id を返す (主にテスト用)."""
    return _job_id_var.get()


# `/{prefix}/{uuid}` の uuid 部分を job_id として拾うパターン.
# UUIDv4 厳密形ではなく緩い 8-4-4-4-12 ヘックスで拾う (storage 側で再検証される)。
_JOB_ID_PATH_RE = re.compile(
    r"/(?:analyze|spec|references|chat|cells|jobs)/"
    r"([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})"
)


class JobIdLoggingMiddleware:
    """URL path から job_id を抽出して contextvar に積む ASGI ミドルウェア.

    BaseHTTPMiddleware ではなく素の ASGI 実装を使う理由: BaseHTTPMiddleware は
    内部で別タスクを生成するため、そこで設定した contextvar が下流ハンドラから
    見えないバグが Starlette に既知 (issue #1715)。素の ASGI なら同一タスクで
    `__call__` が走るので contextvar が確実に伝播する。
    """

    def __init__(self, app: ASGIApp) -> None:
        self.app = app

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        path = scope.get("path", "")
        m = _JOB_ID_PATH_RE.search(path) if isinstance(path, str) else None
        if m is None:
            await self.app(scope, receive, send)
            return
        token = _job_id_var.set(m.group(1))
        try:
            await self.app(scope, receive, send)
        finally:
            _job_id_var.reset(token)
