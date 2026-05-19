"""GET /external-functions と GET /external-functions/used/{job_id} のテスト."""

from __future__ import annotations

import uuid

from fastapi.testclient import TestClient

from backend.storage import Storage
from core.models import CellFormula, SheetInfo, Workbook


class TestRegistryEndpoint:
    def test_returns_bloomberg_functions(self, client: TestClient) -> None:
        r = client.get("/external-functions")
        assert r.status_code == 200
        body = r.json()
        names = {f["name"] for f in body["functions"]}
        assert {"BDH", "BDP", "BDS"} <= names
        # vendor 一覧
        assert "Bloomberg" in body["vendors"]
        # 各定義に必要なフィールドが揃っている
        bdh = next(f for f in body["functions"] if f["name"] == "BDH")
        assert bdh["vendor"] == "Bloomberg"
        assert bdh["signature"].startswith("=BDH(")
        assert bdh["params"]
        assert bdh["examples"]


class TestUsageEndpoint:
    def test_aggregates_usage_per_function(
        self,
        client: TestClient,
        backend_storage: Storage,
    ) -> None:
        meta = backend_storage.create_job("demo.xlsm", b"d")
        wb = Workbook(
            filename="demo.xlsm",
            sheets=[
                SheetInfo(
                    name="Port",
                    rows=10,
                    cols=5,
                    formulas=[
                        CellFormula(
                            coord="A2",
                            formula='=BDP("AAPL US Equity", "PX_LAST")',
                            external_functions=["BDP"],
                        ),
                        CellFormula(
                            coord="A3",
                            formula='=BDP("MSFT US Equity", "PX_LAST")',
                            external_functions=["BDP"],
                        ),
                        CellFormula(
                            coord="B2",
                            formula='=BDH("AAPL US Equity", "PX_LAST", "-1Y")',
                            external_functions=["BDH"],
                        ),
                    ],
                )
            ],
        )
        backend_storage.save_workbook(meta.job_id, wb)
        backend_storage.update_status(meta.job_id, "extracted")

        r = client.get(f"/external-functions/used/{meta.job_id}")
        assert r.status_code == 200
        body = r.json()
        assert body["total_kinds"] == 2
        assert body["total_uses"] == 3
        names = [it["name"] for it in body["items"]]
        assert names == ["BDP", "BDH"]
        bdp = next(it for it in body["items"] if it["name"] == "BDP")
        assert bdp["count"] == 2
        assert bdp["vendor"] == "Bloomberg"
        assert bdp["registered"] is True
        # 全使用箇所が入る (5 件キャップではない)
        assert len(bdp["locations"]) == 2
        assert bdp["locations"][0]["coord"].startswith("Port!")

    def test_empty_when_no_external_functions(
        self,
        client: TestClient,
        backend_storage: Storage,
    ) -> None:
        meta = backend_storage.create_job("demo.xlsx", b"d")
        wb = Workbook(filename="demo.xlsx", sheets=[SheetInfo(name="S", rows=1, cols=1)])
        backend_storage.save_workbook(meta.job_id, wb)
        backend_storage.update_status(meta.job_id, "extracted")

        r = client.get(f"/external-functions/used/{meta.job_id}")
        assert r.status_code == 200
        body = r.json()
        assert body["items"] == []
        assert body["total_kinds"] == 0
        assert body["total_uses"] == 0

    def test_404_when_job_missing(self, client: TestClient) -> None:
        r = client.get(f"/external-functions/used/{uuid.uuid4()}")
        assert r.status_code == 404

    def test_409_when_not_extracted(
        self,
        client: TestClient,
        backend_storage: Storage,
    ) -> None:
        meta = backend_storage.create_job("demo.xlsx", b"d")
        r = client.get(f"/external-functions/used/{meta.job_id}")
        assert r.status_code == 409

    def test_400_on_invalid_job_id(self, client: TestClient) -> None:
        r = client.get("/external-functions/used/not-a-uuid")
        assert r.status_code == 400
