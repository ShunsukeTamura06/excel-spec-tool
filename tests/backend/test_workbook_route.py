"""GET /workbook/{job_id} のテスト."""

from __future__ import annotations

import uuid

from fastapi.testclient import TestClient

from backend.storage import Storage
from core.models import CellFormula, SheetInfo, Workbook


def _seed(storage: Storage, wb: Workbook) -> str:
    meta = storage.create_job("x.xlsx", b"d")
    storage.save_workbook(meta.job_id, wb)
    storage.update_status(meta.job_id, "extracted")
    return meta.job_id


class TestWorkbookRoute:
    def test_returns_workbook_json(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        wb = Workbook(
            filename="demo.xlsx",
            sheets=[
                SheetInfo(
                    name="A",
                    rows=2,
                    cols=1,
                    formulas=[
                        CellFormula(coord="A1", formula="=1+1", refs=[]),
                    ],
                ),
            ],
        )
        job_id = _seed(backend_storage, wb)

        r = client.get(f"/workbook/{job_id}")
        assert r.status_code == 200
        body = r.json()
        assert body["filename"] == "demo.xlsx"
        assert len(body["sheets"]) == 1
        assert body["sheets"][0]["name"] == "A"
        assert body["sheets"][0]["formulas"][0]["formula"] == "=1+1"
        assert body["vba_modules"] == []

    def test_404_when_missing(self, client: TestClient) -> None:
        r = client.get(f"/workbook/{uuid.uuid4()}")
        assert r.status_code == 404

    def test_409_when_not_extracted(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        meta = backend_storage.create_job("foo.xlsm", b"d")
        r = client.get(f"/workbook/{meta.job_id}")
        assert r.status_code == 409

    def test_400_on_invalid_id(self, client: TestClient) -> None:
        r = client.get("/workbook/not-a-uuid")
        assert r.status_code == 400
