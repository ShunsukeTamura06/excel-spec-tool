"""OfficeCLIを任意の変更エンジンとして使うアダプター.

OfficeCLIは外部プロセスとして隔離し、shellを介さない引数配列・タイムアウト・JSON応答の
検証を必須にする。xlsxの名前定義変更と、空セルへの固定テキスト一括追加を公開する。
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
from pathlib import Path

from core.exceptions import (
    MutationProviderError,
    ProviderUnavailableError,
    UnsupportedMutationError,
)
from core.mutation import (
    CellTextBatchOperation,
    MutationPlan,
    MutationResult,
    NamedRangeSetOperation,
    ProviderCapability,
)


class OfficeCliMutationProvider:
    """OfficeCLIのCLI/JSON契約をMutationProviderへ変換する."""

    def __init__(self, executable: str | None = None, timeout_seconds: float = 60.0) -> None:
        """実行ファイルとタイムアウトを設定する.

        Args:
            executable: OfficeCLI実行ファイル。未指定時はOFFICECLI_BINまたはPATHから解決。
            timeout_seconds: 1操作の最大実行秒数。
        """

        configured = executable or os.environ.get("OFFICECLI_BIN") or "officecli"
        self._configured_executable = configured
        self._timeout_seconds = timeout_seconds

    def _resolve_executable(self) -> str | None:
        """設定値を実行可能ファイルのパスへ解決する."""

        return shutil.which(self._configured_executable)

    def _subprocess_env(self) -> dict[str, str]:
        """OfficeCLIへLLM認証情報を渡さない子プロセス環境を作る."""

        env = dict(os.environ)
        for key in ("LLM_API_KEY", "OPENAI_API_KEY", "ANTHROPIC_API_KEY"):
            env.pop(key, None)
        env["OFFICECLI_SKIP_UPDATE"] = "1"
        return env

    def _version(self, executable: str) -> str:
        """OfficeCLIのバージョン文字列を取得する."""

        try:
            completed = subprocess.run(
                [executable, "--version"],
                check=False,
                capture_output=True,
                text=True,
                env=self._subprocess_env(),
                timeout=min(self._timeout_seconds, 5.0),
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise ProviderUnavailableError(f"failed to execute OfficeCLI: {exc}") from exc
        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout).strip()
            raise ProviderUnavailableError(f"OfficeCLI --version failed: {detail}")
        return completed.stdout.strip() or completed.stderr.strip() or "unknown"

    def capability(self) -> ProviderCapability:
        """OfficeCLIの利用可否と、意図的に限定した対応範囲を返す."""

        executable = self._resolve_executable()
        if executable is None:
            return ProviderCapability(
                name="officecli",
                available=False,
                supported_extensions=[".xlsx"],
                supported_operations=["named_range_set", "cell_text_batch"],
                unavailable_reason="officecli executable was not found",
            )
        try:
            version = self._version(executable)
        except ProviderUnavailableError as exc:
            return ProviderCapability(
                name="officecli",
                available=False,
                supported_extensions=[".xlsx"],
                supported_operations=["named_range_set", "cell_text_batch"],
                unavailable_reason=str(exc),
            )
        return ProviderCapability(
            name="officecli",
            available=True,
            version=version,
            supported_extensions=[".xlsx"],
            supported_operations=["named_range_set", "cell_text_batch"],
        )

    def apply(self, plan: MutationPlan, source_path: Path, out_path: Path) -> MutationResult:
        """OfficeCLIで対応済みのxlsx変更を別ファイルへ適用する.

        Args:
            plan: OfficeCLI対応操作を含む計画。
            source_path: 変更しない元xlsx。
            out_path: OfficeCLIが変更する隔離コピー。

        Returns:
            OfficeCLIのバージョンを含む適用結果。

        Raises:
            ProviderUnavailableError: OfficeCLIが利用できない場合。
            UnsupportedMutationError: xlsx/対応済み操作以外が指定された場合。
            MutationProviderError: OfficeCLIが失敗または不正JSONを返した場合。
        """

        operation = plan.operation
        if source_path.suffix.lower() != ".xlsx":
            raise UnsupportedMutationError("OfficeCLI provider currently supports .xlsx only")
        if not isinstance(operation, (NamedRangeSetOperation, CellTextBatchOperation)):
            raise UnsupportedMutationError(
                f"OfficeCLI provider does not support operation: {operation.kind}"
            )
        if isinstance(operation, NamedRangeSetOperation):
            if "[" in operation.name or "]" in operation.name:
                raise MutationProviderError("named range contains unsupported path characters")
            if operation.new_refers_to.startswith("="):
                raise MutationProviderError("OfficeCLI named-range ref must not start with '='")

        executable = self._resolve_executable()
        if executable is None:
            raise ProviderUnavailableError("officecli executable was not found")
        version = self._version(executable)

        try:
            if source_path.resolve() == out_path.resolve():
                raise MutationProviderError("source_path and out_path must be different")
            shutil.copy2(source_path, out_path)
            if isinstance(operation, NamedRangeSetOperation):
                args = [
                    executable,
                    "set",
                    str(out_path),
                    f"/namedrange[{operation.name}]",
                    "--prop",
                    f"ref={operation.new_refers_to}",
                    "--json",
                ]
                changed_count = 1
            else:
                commands = [
                    {
                        "command": "set",
                        "path": f"/{edit.sheet}/{edit.coord}",
                        "props": {"value": edit.value},
                    }
                    for edit in operation.edits
                ]
                args = [
                    executable,
                    "batch",
                    str(out_path),
                    "--commands",
                    json.dumps(commands, ensure_ascii=False, separators=(",", ":")),
                    "--stop-on-error",
                    "--json",
                ]
                changed_count = len(operation.edits)
            completed = subprocess.run(
                args,
                check=False,
                capture_output=True,
                text=True,
                env=self._subprocess_env(),
                timeout=self._timeout_seconds,
            )
        except subprocess.TimeoutExpired as exc:
            raise MutationProviderError(
                f"OfficeCLI timed out after {self._timeout_seconds:g} seconds"
            ) from exc
        except OSError as exc:
            raise MutationProviderError(f"failed to execute OfficeCLI: {exc}") from exc

        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout).strip()
            raise MutationProviderError(f"OfficeCLI mutation failed: {detail}")
        try:
            response = json.loads(completed.stdout)
        except json.JSONDecodeError as exc:
            raise MutationProviderError("OfficeCLI returned invalid JSON") from exc
        if not isinstance(response, (dict, list)):
            raise MutationProviderError("OfficeCLI returned an unexpected JSON payload")
        if isinstance(response, dict) and (
            response.get("success") is False or response.get("error")
        ):
            raise MutationProviderError(f"OfficeCLI reported mutation failure: {response}")
        if isinstance(operation, CellTextBatchOperation):
            data = response.get("data") if isinstance(response, dict) else None
            summary = data.get("summary") if isinstance(data, dict) else None
            if not isinstance(summary, dict):
                raise MutationProviderError("OfficeCLI batch response has no summary")
            try:
                succeeded = int(summary.get("succeeded", 0) or 0)
                failed = int(summary.get("failed", 0) or 0)
            except (TypeError, ValueError) as exc:
                raise MutationProviderError(
                    "OfficeCLI batch response has invalid summary counts"
                ) from exc
            if failed or succeeded != changed_count:
                raise MutationProviderError(
                    "OfficeCLI batch did not apply every planned change: "
                    f"succeeded={succeeded} failed={failed} expected={changed_count}"
                )

        return MutationResult(
            provider="officecli",
            provider_version=version,
            operation=operation.kind,
            changed_count=changed_count,
        )
