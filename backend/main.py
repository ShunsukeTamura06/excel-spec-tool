"""FastAPI アプリケーションのエントリポイント.

起動例:
    uvicorn backend.main:app --reload --port 8000
"""

from __future__ import annotations

from fastapi import FastAPI

from backend.logging_config import JobIdLoggingMiddleware, configure_logging
from backend.routes import analyze, cells, chat, extract, jobs, references, spec


def create_app() -> FastAPI:
    """FastAPI アプリを構築して返す."""
    configure_logging()

    app = FastAPI(
        title="Excel改修支援ツール Backend",
        version="0.1.0",
        description="VBA/数式/参照を含むExcelの統合設計書生成 + 改修対話を提供する.",
    )

    # path から job_id を抽出して相関 ID をログに乗せる
    app.add_middleware(JobIdLoggingMiddleware)

    app.include_router(extract.router)
    app.include_router(analyze.router)
    app.include_router(spec.router)
    app.include_router(references.router)
    app.include_router(chat.router)
    app.include_router(cells.router)
    app.include_router(jobs.router)

    @app.get("/health")
    async def health() -> dict[str, str]:
        return {"status": "ok"}

    return app


app = create_app()
