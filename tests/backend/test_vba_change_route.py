"""VBA変更パッケージ生成・戻りファイル検証APIのテスト."""

from __future__ import annotations

import io
import json
import zipfile

from fastapi.testclient import TestClient

from backend.storage import Storage
from core.models import ReferenceIndex, VbaModule, VbaProcedure, Workbook, WorkbookDiff
from core.mutation import MutationPlan, VbaProcedureReplaceOperation, propose_mutation


def _seed_vba_job(storage: Storage) -> tuple[str, Workbook]:
    """VBA抽出済みのダミー.xlsmジョブを作る."""

    meta = storage.create_job("tool.xlsm", b"dummy-xlsm")
    code = 'Option Explicit\n\nPublic Sub UpdateReport()\n    Range("A1").Value = 1\nEnd Sub\n'
    workbook = Workbook(
        filename="tool.xlsm",
        vba_modules=[
            VbaModule(
                name="Module1",
                type="Module",
                code=code,
                procedures=[
                    VbaProcedure(
                        name="UpdateReport",
                        kind="Sub",
                        start_line=3,
                        end_line=5,
                        code=('Public Sub UpdateReport()\n    Range("A1").Value = 1\nEnd Sub'),
                    )
                ],
            )
        ],
    )
    storage.save_workbook(meta.job_id, workbook)
    storage.save_references(meta.job_id, ReferenceIndex())
    storage.update_status(meta.job_id, "extracted")
    return meta.job_id, workbook


def _plan(job_id: str) -> MutationPlan:
    return MutationPlan(
        source_job_id=job_id,
        requested_provider="windows_vbide",
        operation=VbaProcedureReplaceOperation(
            module_name="Module1",
            procedure_name="UpdateReport",
            new_code=('Public Sub UpdateReport()\n    Range("A1").Value = 2\nEnd Sub'),
        ),
    )


def test_downloads_windows_vba_package(
    client: TestClient,
    backend_storage: Storage,
) -> None:
    """原本を含むWindows用ZIPをダウンロードできる."""

    job_id, _ = _seed_vba_job(backend_storage)
    response = client.post(
        f"/jobs/{job_id}/vba-change/package",
        json={"plan": _plan(job_id).model_dump(mode="json")},
    )

    assert response.status_code == 200
    assert response.headers["content-type"] == "application/zip"
    with zipfile.ZipFile(io.BytesIO(response.content)) as archive:
        assert archive.read("original.xlsm") == b"dummy-xlsm"
        assert "ProcStartLine" in archive.read("apply_vba_change.ps1").decode("utf-8-sig")


def test_verifies_returned_xlsm_against_expected_vba_diff(
    client: TestClient,
    backend_storage: Storage,
    monkeypatch,
) -> None:
    """Windows適用後ファイルの実差分が計画通りなら合格させる."""

    job_id, workbook = _seed_vba_job(backend_storage)
    plan = _plan(job_id)
    expected = propose_mutation(plan, workbook, ReferenceIndex())

    def fake_extraction(
        storage: Storage,
        result_job_id: str,
        filename: str,
    ) -> None:
        storage.save_workbook(
            result_job_id,
            Workbook(filename=filename),
        )
        storage.save_references(result_job_id, ReferenceIndex())
        storage.update_status(result_job_id, "extracted")

    async def fake_self_verify(*args, **kwargs) -> WorkbookDiff:
        return expected

    monkeypatch.setattr("backend.routes.vba_change._run_extraction", fake_extraction)
    monkeypatch.setattr("backend.routes.vba_change._self_verify", fake_self_verify)
    response = client.post(
        f"/jobs/{job_id}/vba-change/verify",
        files={"file": ("revised.xlsm", b"revised-xlsm", "application/octet-stream")},
        data={"plan_json": plan.model_dump_json()},
    )

    assert response.status_code == 200
    body = response.json()
    assert body["verification"]["status"] == "passed"
    assert body["provider"]["provider"] == "windows_vbide"
    assert body["diff"]["vba_modules"][0]["name"] == "Module1"


def test_rejects_unexpected_vba_result(
    client: TestClient,
    backend_storage: Storage,
    monkeypatch,
) -> None:
    """期待したVBA差分がない戻りファイルを不合格にする."""

    job_id, _ = _seed_vba_job(backend_storage)
    plan = _plan(job_id)

    def fake_extraction(
        storage: Storage,
        result_job_id: str,
        filename: str,
    ) -> None:
        storage.update_status(result_job_id, "extracted")

    async def fake_self_verify(*args, **kwargs) -> WorkbookDiff:
        return WorkbookDiff(
            before_filename="tool.xlsm",
            after_filename="revised.xlsm",
        )

    monkeypatch.setattr("backend.routes.vba_change._run_extraction", fake_extraction)
    monkeypatch.setattr("backend.routes.vba_change._self_verify", fake_self_verify)
    response = client.post(
        f"/jobs/{job_id}/vba-change/verify",
        files={"file": ("revised.xlsm", b"unchanged-xlsm", "application/octet-stream")},
        data={"plan_json": json.dumps(plan.model_dump(mode="json"))},
    )

    assert response.status_code == 409
