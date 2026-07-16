"""Windows VBA変更パッケージ生成のテスト."""

from __future__ import annotations

import io
import json
import zipfile
from pathlib import Path

from core.mutation import MutationPlan, VbaProcedureReplaceOperation
from core.vba_package import build_vba_change_package


def test_package_contains_original_plan_script_and_replacement(tmp_path: Path) -> None:
    """Windowsで適用するための必要ファイルをZIPへ含める."""

    source = tmp_path / "tool.xlsm"
    source.write_bytes(b"dummy-xlsm")
    plan = MutationPlan(
        source_job_id="00000000-0000-4000-8000-000000000000",
        requested_provider="windows_vbide",
        operation=VbaProcedureReplaceOperation(
            module_name="Module1",
            procedure_name="UpdateReport",
            new_code="Public Sub UpdateReport()\nEnd Sub",
        ),
    )

    package = build_vba_change_package(source, plan)

    with zipfile.ZipFile(io.BytesIO(package)) as archive:
        assert set(archive.namelist()) == {
            "original.xlsm",
            "replacement.bas",
            "plan.json",
            "apply_vba_change.ps1",
            "README.txt",
        }
        assert archive.read("original.xlsm") == b"dummy-xlsm"
        assert archive.read("replacement.bas").decode() == ("Public Sub UpdateReport()\nEnd Sub")
        manifest = json.loads(archive.read("plan.json"))
        assert manifest["operation"]["kind"] == "vba_procedure_replace"
        script_bytes = archive.read("apply_vba_change.ps1")
        assert script_bytes.startswith(b"\xef\xbb\xbf")
        script = script_bytes.decode("utf-8-sig")
        assert "AutomationSecurity = 3" in script
        assert "EnableEvents = $false" in script
        assert "ProcStartLine" in script
        assert "ProcBodyLine" in script
        assert "DeleteLines" in script
        assert "InsertLines" in script


def test_package_warns_when_vba_signature_exists(tmp_path: Path) -> None:
    """VBA署名パーツを検出してREADMEとmanifestへ警告を残す."""

    source = tmp_path / "signed.xlsm"
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("xl/vbaProject.bin", b"vba")
        archive.writestr("xl/vbaProjectSignature.bin", b"signature")
    source.write_bytes(buffer.getvalue())
    plan = MutationPlan(
        source_job_id="00000000-0000-4000-8000-000000000000",
        requested_provider="windows_vbide",
        operation=VbaProcedureReplaceOperation(
            module_name="Module1",
            procedure_name="UpdateReport",
            new_code="Public Sub UpdateReport()\nEnd Sub",
        ),
    )

    with zipfile.ZipFile(io.BytesIO(build_vba_change_package(source, plan))) as archive:
        manifest = json.loads(archive.read("plan.json"))
        assert manifest["vba_signature_present"] is True
        assert "digital signature" in archive.read("README.txt").decode()
