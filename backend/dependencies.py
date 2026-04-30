"""FastAPI 依存性注入用のヘルパー.

ルートで `Depends(get_storage)` / `Depends(get_llm_client)` を使えるよう、
モジュールレベルで定義。テストでは `app.dependency_overrides` で差し替える。
"""

from __future__ import annotations

from backend.llm_client import LLMClient, get_default_client
from backend.storage import Storage

_storage_singleton: Storage | None = None


def get_storage() -> Storage:
    """環境変数 JOBS_DIR を見て Storage を返す (プロセス内シングルトン)."""
    global _storage_singleton
    if _storage_singleton is None:
        _storage_singleton = Storage.from_env()
    return _storage_singleton


def get_llm_client() -> LLMClient:
    """環境変数を見てデフォルト LLM クライアントを返す."""
    return get_default_client()


def reset_storage_singleton() -> None:
    """テスト用: Storage シングルトンをクリアする."""
    global _storage_singleton
    _storage_singleton = None
