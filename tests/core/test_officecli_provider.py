"""OfficeCLI変更プロバイダーのプロセス境界テスト."""

from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

from core.exceptions import MutationProviderError, UnsupportedMutationError
from core.mutation import FixedRefReplaceOperation, MutationPlan, NamedRangeSetOperation
from core.officecli_provider import OfficeCliMutationProvider


def _plan() -> MutationPlan:
    """OfficeCLIが対応する名前定義変更計画を返す."""

    return MutationPlan(
        source_job_id="00000000-0000-4000-8000-000000000000",
        requested_provider="officecli",
        operation=NamedRangeSetOperation(name="TaxRate", new_refers_to="Data!$A$2"),
    )


def test_capability_reports_missing_binary(monkeypatch: pytest.MonkeyPatch) -> None:
    """未導入環境では例外でなく利用不可capabilityを返す."""

    monkeypatch.setattr("core.officecli_provider.shutil.which", lambda _: None)

    capability = OfficeCliMutationProvider().capability()

    assert not capability.available
    assert capability.supported_operations == ["named_range_set"]
    assert capability.unavailable_reason


def test_apply_uses_argument_array_and_json_contract(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """shellを使わず隔離コピーだけをOfficeCLIへ渡す."""

    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    source.write_bytes(b"xlsx-placeholder")
    calls: list[list[str]] = []

    def fake_run(args: list[str], **_: object) -> subprocess.CompletedProcess[str]:
        calls.append(args)
        if args[-1] == "--version":
            return subprocess.CompletedProcess(args, 0, stdout="officecli 1.0.136\n", stderr="")
        return subprocess.CompletedProcess(args, 0, stdout='{"success":true}', stderr="")

    monkeypatch.setattr("core.officecli_provider.shutil.which", lambda _: "/opt/officecli")
    monkeypatch.setattr("core.officecli_provider.subprocess.run", fake_run)

    result = OfficeCliMutationProvider().apply(_plan(), source, output)

    assert source.read_bytes() == b"xlsx-placeholder"
    assert output.read_bytes() == b"xlsx-placeholder"
    assert calls[1] == [
        "/opt/officecli",
        "set",
        str(output),
        "/namedrange[TaxRate]",
        "--prop",
        "ref=Data!$A$2",
        "--json",
    ]
    assert result.provider_version == "officecli 1.0.136"


def test_apply_rejects_unsupported_operation(tmp_path: Path) -> None:
    """OfficeCLIで未検証の操作をcapability以上に実行しない."""

    source = tmp_path / "source.xlsx"
    source.write_bytes(b"xlsx-placeholder")
    plan = MutationPlan(
        source_job_id="00000000-0000-4000-8000-000000000000",
        requested_provider="officecli",
        operation=FixedRefReplaceOperation(old_ref="Data!A1", new_ref="Data!A2"),
    )

    with pytest.raises(UnsupportedMutationError, match="does not support operation"):
        OfficeCliMutationProvider(executable="/opt/officecli").apply(
            plan, source, tmp_path / "output.xlsx"
        )


def test_apply_rejects_invalid_json(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """終了コード0でも非JSON応答なら成功扱いしない."""

    source = tmp_path / "source.xlsx"
    source.write_bytes(b"xlsx-placeholder")

    def fake_run(args: list[str], **_: object) -> subprocess.CompletedProcess[str]:
        stdout = "officecli 1.0.136\n" if args[-1] == "--version" else "not-json"
        return subprocess.CompletedProcess(args, 0, stdout=stdout, stderr="")

    monkeypatch.setattr("core.officecli_provider.shutil.which", lambda _: "/opt/officecli")
    monkeypatch.setattr("core.officecli_provider.subprocess.run", fake_run)

    with pytest.raises(MutationProviderError, match="invalid JSON"):
        OfficeCliMutationProvider().apply(_plan(), source, tmp_path / "output.xlsx")


def test_apply_rejects_explicit_failure_payload(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """終了コード0でもOfficeCLI自身が失敗を返した場合は成功扱いしない."""

    source = tmp_path / "source.xlsx"
    source.write_bytes(b"xlsx-placeholder")

    def fake_run(args: list[str], **_: object) -> subprocess.CompletedProcess[str]:
        stdout = "officecli 1.0.136\n" if args[-1] == "--version" else '{"success":false}'
        return subprocess.CompletedProcess(args, 0, stdout=stdout, stderr="")

    monkeypatch.setattr("core.officecli_provider.shutil.which", lambda _: "/opt/officecli")
    monkeypatch.setattr("core.officecli_provider.subprocess.run", fake_run)

    with pytest.raises(MutationProviderError, match="reported mutation failure"):
        OfficeCliMutationProvider().apply(_plan(), source, tmp_path / "output.xlsx")
