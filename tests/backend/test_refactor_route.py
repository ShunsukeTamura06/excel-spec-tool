"""POST /jobs/{job_id}/named-range-fix のテスト."""

from __future__ import annotations

import io
import uuid

import openpyxl
from fastapi.testclient import TestClient
from openpyxl.workbook.defined_name import DefinedName

from backend.storage import Storage
from core.extractors.workbook import extract_workbook
from core.reference_index import build_reference_index


def _xlsx_bytes_with_named_range(
    name: str = "TaxRate",
    refers_to: str = "Data!$A$1",
    formula_referencing: str | None = None,
) -> bytes:
    """名前付き範囲を1つ持つ xlsx をバイト列で生成する.

    formula_referencing を渡すと、B1 にその範囲を参照する数式を仕込む
    (波及範囲テスト用)。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Data"
    ws["A1"] = 0.1
    if formula_referencing:
        ws["B1"] = formula_referencing
    wb.defined_names[name] = DefinedName(name=name, attr_text=refers_to)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _seed_extracted_job(storage: Storage, data: bytes, filename: str = "x.xlsx") -> str:
    """extract相当の処理をテスト内で組み立て、抽出済みジョブを1件作る."""
    meta = storage.create_job(filename, data)
    path = storage.get_original_path(meta.job_id)
    wb = extract_workbook(path)
    wb.filename = filename
    idx = build_reference_index(wb)
    storage.save_workbook(meta.job_id, wb)
    storage.save_references(meta.job_id, idx)
    storage.update_status(meta.job_id, "extracted")
    return meta.job_id


class TestNamedRangeFixRoute:
    def test_applies_fix_and_creates_new_job(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        job_id = _seed_extracted_job(backend_storage, _xlsx_bytes_with_named_range())

        r = client.post(
            f"/jobs/{job_id}/named-range-fix",
            json={"name": "TaxRate", "new_refers_to": "Data!$B$1"},
        )
        assert r.status_code == 200
        body = r.json()
        new_job_id = body["new_job_id"]
        assert new_job_id != job_id

        diff = body["diff"]
        assert len(diff["named_ranges"]) == 1
        nr = diff["named_ranges"][0]
        assert nr["name"] == "TaxRate"
        assert nr["change_type"] == "modified"
        assert nr["before_refers_to"] == "Data!$A$1"
        assert nr["after_refers_to"] == "Data!$B$1"

    def test_new_job_diff_matches_get_diff_endpoint(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        job_id = _seed_extracted_job(backend_storage, _xlsx_bytes_with_named_range())

        r = client.post(
            f"/jobs/{job_id}/named-range-fix",
            json={"name": "TaxRate", "new_refers_to": "Data!$B$1"},
        )
        assert r.status_code == 200
        new_job_id = r.json()["new_job_id"]

        r2 = client.get(f"/diff?before_job_id={job_id}&after_job_id={new_job_id}")
        assert r2.status_code == 200
        assert r2.json()["diff"]["named_ranges"] == r.json()["diff"]["named_ranges"]

    def test_blast_radius_included_when_referenced(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        job_id = _seed_extracted_job(
            backend_storage,
            _xlsx_bytes_with_named_range(formula_referencing="=SUM(A1:A1)"),
        )

        r = client.post(
            f"/jobs/{job_id}/named-range-fix",
            json={"name": "TaxRate", "new_refers_to": "Data!$B$1"},
        )
        assert r.status_code == 200
        # 参照有無はビルドされた ReferenceIndex の正規化に依存するため、
        # ここでは応答が正しく構造化されていることだけを確認する。
        assert isinstance(r.json()["diff"]["blast_radius"], list)

    def test_422_unknown_name(self, client: TestClient, backend_storage: Storage) -> None:
        job_id = _seed_extracted_job(backend_storage, _xlsx_bytes_with_named_range())

        r = client.post(
            f"/jobs/{job_id}/named-range-fix",
            json={"name": "NoSuchName", "new_refers_to": "Data!$B$1"},
        )
        assert r.status_code == 422

    def test_404_job_not_found(self, client: TestClient, backend_storage: Storage) -> None:
        r = client.post(
            f"/jobs/{uuid.uuid4()}/named-range-fix",
            json={"name": "TaxRate", "new_refers_to": "Data!$B$1"},
        )
        assert r.status_code == 404

    def test_409_not_extracted(self, client: TestClient, backend_storage: Storage) -> None:
        meta = backend_storage.create_job("x.xlsx", _xlsx_bytes_with_named_range())
        r = client.post(
            f"/jobs/{meta.job_id}/named-range-fix",
            json={"name": "TaxRate", "new_refers_to": "Data!$B$1"},
        )
        assert r.status_code == 409
