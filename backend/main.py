"""FastAPI アプリケーションのエントリポイント.

起動例:
    uvicorn backend.main:app --reload --port 8000
"""

from __future__ import annotations

import os

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from backend.logging_config import JobIdLoggingMiddleware, configure_logging
from backend.routes import (
    analyze,
    cells,
    chat,
    diagrams,
    extract,
    jobs,
    references,
    spec,
    workbook,
)

# Nuxt dev server (default port 3000). 環境変数で上書き可能 (カンマ区切り).
_DEFAULT_CORS_ORIGINS = "http://localhost:3000,http://127.0.0.1:3000"


def _cors_origins() -> list[str]:
    raw = os.environ.get("CORS_ALLOW_ORIGINS", _DEFAULT_CORS_ORIGINS)
    return [o.strip() for o in raw.split(",") if o.strip()]


def create_app() -> FastAPI:
    """FastAPI アプリを構築して返す."""
    configure_logging()

    app = FastAPI(
        title="Excel改修支援ツール Backend",
        version="0.1.0",
        description="VBA/数式/参照を含むExcelの統合設計書生成 + 改修対話を提供する.",
    )

    # CORS — Nuxt フロントエンド (別ポート) からのアクセスを許可する.
    app.add_middleware(
        CORSMiddleware,
        allow_origins=_cors_origins(),
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # path から job_id を抽出して相関 ID をログに乗せる
    app.add_middleware(JobIdLoggingMiddleware)

    app.include_router(extract.router)
    app.include_router(analyze.router)
    app.include_router(spec.router)
    app.include_router(references.router)
    app.include_router(chat.router)
    app.include_router(cells.router)
    app.include_router(diagrams.router)
    app.include_router(workbook.router)
    app.include_router(jobs.router)

    @app.get("/health")
    async def health() -> dict[str, str]:
        return {"status": "ok"}

    return app


app = create_app()
