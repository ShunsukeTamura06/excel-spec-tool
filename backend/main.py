"""FastAPI アプリケーションのエントリポイント.

起動例:
    uvicorn backend.main:app --reload --port 8000
"""

from __future__ import annotations

import logging

from fastapi import FastAPI

from backend.routes import analyze, chat, extract, jobs, references, spec

logging.basicConfig(level=logging.INFO)


def create_app() -> FastAPI:
    """FastAPI アプリを構築して返す."""
    app = FastAPI(
        title="Excel改修支援ツール Backend",
        version="0.1.0",
        description="VBA/数式/参照を含むExcelの統合設計書生成 + 改修対話を提供する.",
    )

    app.include_router(extract.router)
    app.include_router(analyze.router)
    app.include_router(spec.router)
    app.include_router(references.router)
    app.include_router(chat.router)
    app.include_router(jobs.router)

    @app.get("/health")
    async def health() -> dict[str, str]:
        return {"status": "ok"}

    return app


app = create_app()
