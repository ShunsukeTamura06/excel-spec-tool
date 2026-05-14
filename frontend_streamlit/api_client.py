"""Backend (FastAPI) を叩く httpx ラッパー.

SPEC.md §6.3 に基づき、Frontend は本クラスを通してのみ Backend を呼ぶ。
Backend URL は環境変数 `BACKEND_URL` から (デフォルト `http://localhost:8000`)。
"""

from __future__ import annotations

import os
from types import TracebackType
from typing import Any

import httpx

DEFAULT_BACKEND_URL = "http://localhost:8000"
DEFAULT_TIMEOUT = 60.0


class BackendError(Exception):
    """Backend からエラー応答が返ってきた場合."""

    def __init__(self, status_code: int, detail: str) -> None:
        super().__init__(f"backend error {status_code}: {detail}")
        self.status_code = status_code
        self.detail = detail


class BackendClient:
    """Backend (FastAPI) クライアント.

    `httpx.Client` を内部に保持する. テストでは `http_client` 引数で
    `httpx.Client(transport=httpx.ASGITransport(app=app))` を渡し、
    実 HTTP を経由せずに Backend を呼べる。
    """

    def __init__(
        self,
        base_url: str | None = None,
        http_client: httpx.Client | None = None,
        timeout: float = DEFAULT_TIMEOUT,
    ) -> None:
        self.base_url = (base_url or os.environ.get("BACKEND_URL") or DEFAULT_BACKEND_URL).rstrip(
            "/"
        )
        if http_client is not None:
            self._client = http_client
            self._owns_client = False
        else:
            self._client = httpx.Client(timeout=timeout)
            self._owns_client = True

    # ---------- lifecycle ----------

    def close(self) -> None:
        if self._owns_client:
            self._client.close()

    def __enter__(self) -> BackendClient:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        self.close()

    # ---------- internals ----------

    def _url(self, path: str) -> str:
        if not path.startswith("/"):
            path = "/" + path
        return f"{self.base_url}{path}"

    def _check(self, r: httpx.Response) -> None:
        if r.is_success:
            return
        try:
            payload = r.json()
            detail = (
                payload.get("detail", str(payload)) if isinstance(payload, dict) else str(payload)
            )
        except Exception:  # noqa: BLE001
            detail = r.text or r.reason_phrase
        raise BackendError(status_code=r.status_code, detail=str(detail))

    # ---------- endpoints ----------

    def extract(
        self,
        filename: str,
        data: bytes,
        mime: str = "application/octet-stream",
    ) -> str:
        """POST /extract -> job_id."""
        r = self._client.post(
            self._url("/extract"),
            files={"file": (filename, data, mime)},
        )
        self._check(r)
        return str(r.json()["job_id"])

    def analyze(self, job_id: str) -> dict[str, Any]:
        """POST /analyze/{job_id}."""
        r = self._client.post(self._url(f"/analyze/{job_id}"))
        self._check(r)
        return dict(r.json())

    def get_spec(self, job_id: str) -> dict[str, Any]:
        """GET /spec/{job_id} -> {spec_md, meta}."""
        r = self._client.get(self._url(f"/spec/{job_id}"))
        self._check(r)
        return dict(r.json())

    def get_references(self, job_id: str, target: str) -> list[dict[str, Any]]:
        """GET /references/{job_id}?target=... -> [Reference dict, ...]."""
        r = self._client.get(
            self._url(f"/references/{job_id}"),
            params={"target": target},
        )
        self._check(r)
        return list(r.json().get("refs", []))

    def chat(self, job_id: str, message: str) -> dict[str, Any]:
        """POST /chat/{job_id} -> {reply, history}."""
        r = self._client.post(
            self._url(f"/chat/{job_id}"),
            json={"message": message},
        )
        self._check(r)
        return dict(r.json())

    def get_chat_history(self, job_id: str) -> list[dict[str, Any]]:
        """GET /chat/{job_id}/history -> [ChatMessage dict, ...]."""
        r = self._client.get(self._url(f"/chat/{job_id}/history"))
        self._check(r)
        return list(r.json().get("history", []))

    def list_jobs(self) -> list[dict[str, Any]]:
        """GET /jobs -> [JobMeta dict, ...]."""
        r = self._client.get(self._url("/jobs"))
        self._check(r)
        return list(r.json().get("jobs", []))

    def delete_job(self, job_id: str) -> bool:
        """DELETE /jobs/{job_id} -> bool."""
        r = self._client.delete(self._url(f"/jobs/{job_id}"))
        self._check(r)
        return bool(r.json().get("deleted", False))

    def health(self) -> bool:
        """GET /health -> True."""
        try:
            r = self._client.get(self._url("/health"))
            return r.is_success
        except httpx.HTTPError:
            return False
