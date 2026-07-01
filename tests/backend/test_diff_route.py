"""GET /diff のテスト."""

from __future__ import annotations

import io
import uuid

import openpyxl
from fastapi.testclient import TestClient

from backend.storage import Storage
from core.extractors.workbook import extract_workbook
from core.models import AnalysisRisk
from core.reference_index import build_reference_index


def _xlsx_bytes(cells: dict[str, object], sheet_name: str = "Sheet") -> bytes:
    """簡易xlsxをバイト列で生成する. cells は {"A1": 値または数式文字列, ...}."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = sheet_name
    for coord, value in cells.items():
        ws[coord] = value
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _seed_extracted_job(
    storage: Storage,
    data: bytes,
    filename: str = "x.xlsx",
    extra_risks: list[AnalysisRisk] | None = None,
) -> str:
    """extract相当の処理をテスト内で組み立て、抽出済みジョブを1件作る."""
    meta = storage.create_job(filename, data)
    path = storage.get_original_path(meta.job_id)
    wb = extract_workbook(path)
    wb.filename = filename
    if extra_risks:
        wb.analysis_risks = extra_risks
    idx = build_reference_index(wb)
    storage.save_workbook(meta.job_id, wb)
    storage.save_references(meta.job_id, idx)
    storage.update_status(meta.job_id, "extracted")
    return meta.job_id


class TestDiffRoute:
    def test_returns_diff_for_two_jobs(self, client: TestClient, backend_storage: Storage) -> None:
        before_id = _seed_extracted_job(backend_storage, _xlsx_bytes({"A1": 1, "B1": "=A1+1"}))
        after_id = _seed_extracted_job(backend_storage, _xlsx_bytes({"A1": 1, "B1": "=A1+2"}))

        r = client.get(f"/diff?before_job_id={before_id}&after_job_id={after_id}")
        assert r.status_code == 200
        cells = r.json()["diff"]["cells"]
        assert len(cells) == 1
        assert cells[0]["coord"] == "B1"
        assert cells[0]["change_type"] == "modified"

    def test_blast_radius_included_for_removed_referenced_range(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        before_id = _seed_extracted_job(
            backend_storage,
            _xlsx_bytes({"A1": 10, "A2": 20, "B1": "=SUM(A1:A2)"}),
            filename="before.xlsx",
        )
        after_id = _seed_extracted_job(
            backend_storage,
            _xlsx_bytes({"B1": "=SUM(A1:A2)"}),
            filename="after.xlsx",
        )

        r = client.get(f"/diff?before_job_id={before_id}&after_job_id={after_id}")
        assert r.status_code == 200
        blast_radius = r.json()["diff"]["blast_radius"]
        locations = {entry["location"] for entry in blast_radius}
        assert "Sheet!A1" in locations or "Sheet!A2" in locations

    def test_no_changes_between_identical_jobs(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        data = _xlsx_bytes({"A1": 1, "B1": "hello"})
        before_id = _seed_extracted_job(backend_storage, data, filename="a.xlsx")
        after_id = _seed_extracted_job(backend_storage, data, filename="b.xlsx")

        r = client.get(f"/diff?before_job_id={before_id}&after_job_id={after_id}")
        assert r.status_code == 200
        diff = r.json()["diff"]
        assert diff["cells"] == []
        assert diff["blast_radius"] == []

    def test_existing_risks_passed_through(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        risk = AnalysisRisk(
            category="dynamic_formula",
            severity="medium",
            location="Sheet!A1",
            evidence="INDIRECT(...)",
            description="動的参照",
            recommendation="確認してください",
        )
        before_id = _seed_extracted_job(
            backend_storage, _xlsx_bytes({"A1": 1}), filename="a.xlsx", extra_risks=[risk]
        )
        after_id = _seed_extracted_job(backend_storage, _xlsx_bytes({"A1": 1}), filename="b.xlsx")

        r = client.get(f"/diff?before_job_id={before_id}&after_job_id={after_id}")
        assert r.status_code == 200
        risks = r.json()["diff"]["existing_risks"]
        assert len(risks) == 1
        assert risks[0]["category"] == "dynamic_formula"


class TestDiffRouteErrors:
    def test_400_invalid_before_job_id(self, client: TestClient, backend_storage: Storage) -> None:
        after_id = _seed_extracted_job(backend_storage, _xlsx_bytes({}))
        r = client.get(f"/diff?before_job_id=not-a-uuid&after_job_id={after_id}")
        assert r.status_code == 400

    def test_400_invalid_after_job_id(self, client: TestClient, backend_storage: Storage) -> None:
        before_id = _seed_extracted_job(backend_storage, _xlsx_bytes({}))
        r = client.get(f"/diff?before_job_id={before_id}&after_job_id=not-a-uuid")
        assert r.status_code == 400

    def test_400_same_job_id_for_both(self, client: TestClient, backend_storage: Storage) -> None:
        job_id = _seed_extracted_job(backend_storage, _xlsx_bytes({}))
        r = client.get(f"/diff?before_job_id={job_id}&after_job_id={job_id}")
        assert r.status_code == 400

    def test_404_before_job_not_found(self, client: TestClient, backend_storage: Storage) -> None:
        after_id = _seed_extracted_job(backend_storage, _xlsx_bytes({}))
        r = client.get(f"/diff?before_job_id={uuid.uuid4()}&after_job_id={after_id}")
        assert r.status_code == 404

    def test_404_after_job_not_found(self, client: TestClient, backend_storage: Storage) -> None:
        before_id = _seed_extracted_job(backend_storage, _xlsx_bytes({}))
        r = client.get(f"/diff?before_job_id={before_id}&after_job_id={uuid.uuid4()}")
        assert r.status_code == 404

    def test_409_before_not_extracted(self, client: TestClient, backend_storage: Storage) -> None:
        meta = backend_storage.create_job("x.xlsx", _xlsx_bytes({}))
        after_id = _seed_extracted_job(backend_storage, _xlsx_bytes({}))
        r = client.get(f"/diff?before_job_id={meta.job_id}&after_job_id={after_id}")
        assert r.status_code == 409

    def test_409_after_not_extracted(self, client: TestClient, backend_storage: Storage) -> None:
        before_id = _seed_extracted_job(backend_storage, _xlsx_bytes({}))
        meta = backend_storage.create_job("y.xlsx", _xlsx_bytes({}))
        r = client.get(f"/diff?before_job_id={before_id}&after_job_id={meta.job_id}")
        assert r.status_code == 409
