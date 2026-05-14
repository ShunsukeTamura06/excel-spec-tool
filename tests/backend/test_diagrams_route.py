"""GET /diagrams/{job_id} のテスト."""

from __future__ import annotations

import uuid

import pytest
from fastapi.testclient import TestClient

from backend.storage import Storage
from core.models import (
    CellFormula,
    SheetInfo,
    VbaModule,
    VbaProcedure,
    Workbook,
)


def _seed_job(storage: Storage, wb: Workbook) -> str:
    """meta + workbook だけ書き込んだジョブを作る (extract 完了相当)."""
    meta = storage.create_job("demo.xlsx", b"dummy")
    storage.save_workbook(meta.job_id, wb)
    storage.update_status(meta.job_id, "extracted")
    return meta.job_id


@pytest.fixture
def three_sheet_wb() -> Workbook:
    return Workbook(
        filename="demo.xlsx",
        sheets=[
            SheetInfo(name="Input", rows=10, cols=5),
            SheetInfo(
                name="Calc",
                rows=10,
                cols=5,
                formulas=[
                    CellFormula(
                        coord="A2",
                        formula="=Input!A2",
                        refs=["Input!A2"],
                    ),
                    CellFormula(
                        coord="B2",
                        formula="=Input!B2*2",
                        refs=["Input!B2"],
                    ),
                ],
            ),
            SheetInfo(
                name="Output",
                rows=2,
                cols=1,
                formulas=[
                    CellFormula(
                        coord="A1",
                        formula="=SUM(Calc!A:A)",
                        refs=["Calc!A:A"],
                    ),
                ],
            ),
        ],
        vba_modules=[
            VbaModule(
                name="M1",
                type="Module",
                code="",
                procedures=[
                    VbaProcedure(
                        name="Entry",
                        kind="Sub",
                        start_line=1,
                        end_line=3,
                        code="Sub Entry()\n    Call Helper\nEnd Sub",
                    ),
                    VbaProcedure(
                        name="Helper",
                        kind="Sub",
                        start_line=4,
                        end_line=5,
                        code="Sub Helper()\nEnd Sub",
                    ),
                ],
            ),
        ],
    )


class TestDiagramsRoute:
    def test_returns_both_diagrams(
        self,
        client: TestClient,
        backend_storage: Storage,
        three_sheet_wb: Workbook,
    ) -> None:
        job_id = _seed_job(backend_storage, three_sheet_wb)

        r = client.get(f"/diagrams/{job_id}")
        assert r.status_code == 200
        body = r.json()
        assert set(body.keys()) == {"sheet_deps", "vba_calls"}

        sd = body["sheet_deps"]
        assert sd["kind"] == "sheet_deps"
        node_ids = {n["id"] for n in sd["nodes"]}
        assert node_ids == {"Input", "Calc", "Output"}
        edges = {(e["src"], e["dst"]): e for e in sd["edges"]}
        assert ("Calc", "Input") in edges
        assert ("Output", "Calc") in edges
        assert edges[("Calc", "Input")]["weight"] == 2

        vc = body["vba_calls"]
        assert vc["kind"] == "vba_calls"
        node_ids = {n["id"] for n in vc["nodes"]}
        assert node_ids == {"M1.Entry", "M1.Helper"}
        vc_edges = {(e["src"], e["dst"]) for e in vc["edges"]}
        assert ("M1.Entry", "M1.Helper") in vc_edges

    def test_404_when_job_missing(self, client: TestClient) -> None:
        # 存在しないが形式は正しい UUIDv4
        nonexistent = str(uuid.uuid4())
        r = client.get(f"/diagrams/{nonexistent}")
        assert r.status_code == 404
        assert "not found" in r.json()["detail"].lower()

    def test_409_when_workbook_not_extracted(
        self,
        client: TestClient,
        backend_storage: Storage,
    ) -> None:
        """create_job 直後 (extracted.json 未生成) は 409."""
        meta = backend_storage.create_job("foo.xlsm", b"data")
        r = client.get(f"/diagrams/{meta.job_id}")
        assert r.status_code == 409
        assert "extract" in r.json()["detail"].lower()

    def test_400_on_invalid_job_id(self, client: TestClient) -> None:
        r = client.get("/diagrams/not-a-uuid")
        assert r.status_code == 400

    def test_empty_workbook(
        self,
        client: TestClient,
        backend_storage: Storage,
    ) -> None:
        """シートも VBA も無い workbook でも 200 で空配列を返す."""
        wb = Workbook(filename="empty.xlsx")
        job_id = _seed_job(backend_storage, wb)
        r = client.get(f"/diagrams/{job_id}")
        assert r.status_code == 200
        body = r.json()
        assert body["sheet_deps"]["nodes"] == []
        assert body["sheet_deps"]["edges"] == []
        assert body["vba_calls"]["nodes"] == []
        assert body["vba_calls"]["edges"] == []
