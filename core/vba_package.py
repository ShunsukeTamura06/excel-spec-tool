"""Windows Excel/VBIDEでVBA置換を適用するZIPパッケージを生成する."""

from __future__ import annotations

import io
import json
import zipfile
from pathlib import Path

from core.exceptions import VbaChangeError
from core.mutation import MutationPlan, VbaProcedureReplaceOperation

_POWERSHELL_SCRIPT = r"""$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$source = Join-Path $root "original.xlsm"
$output = Join-Path $root "revised.xlsm"
$replacementPath = Join-Path $root "replacement.bas"
$planPath = Join-Path $root "plan.json"
$resultPath = Join-Path $root "result.json"

$excel = $null
$workbook = $null

try {
    if (-not (Test-Path -LiteralPath $source)) {
        throw "original.xlsm が見つかりません。ZIPを展開したフォルダ内で実行してください。"
    }
    Copy-Item -LiteralPath $source -Destination $output -Force
    $plan = Get-Content -LiteralPath $planPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $replacement = Get-Content -LiteralPath $replacementPath -Raw -Encoding UTF8

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open($output, 0, $false)

    try {
        $component = $workbook.VBProject.VBComponents.Item(
            [string]$plan.operation.module_name
        )
    }
    catch {
        throw (
            "VBAプロジェクトへアクセスできません。対象モジュール、プロジェクトロック、" +
            "Excelの「VBAプロジェクト オブジェクト モデルへのアクセスを" +
            "信頼する」設定を確認してください。"
        )
    }

    $codeModule = $component.CodeModule
    $procedureName = [string]$plan.operation.procedure_name
    $procedureKind = 0
    $procedureStart = $codeModule.ProcStartLine($procedureName, $procedureKind)
    $bodyLine = $codeModule.ProcBodyLine($procedureName, $procedureKind)
    $procedureCount = $codeModule.ProcCountLines($procedureName, $procedureKind)
    if ($procedureStart -le 0 -or $bodyLine -le 0 -or $procedureCount -le 0) {
        throw "対象プロシージャが見つかりません: $procedureName"
    }

    $bodyText = $codeModule.Lines($bodyLine, 1)
    $endKeyword = if ($bodyText -match "(?i)\bFunction\b") { "Function" } else { "Sub" }
    $lastCandidate = $procedureStart + $procedureCount - 1
    $endLine = 0
    for ($line = $bodyLine; $line -le $lastCandidate; $line++) {
        if ($codeModule.Lines($line, 1) -match "^\s*End\s+$endKeyword\b") {
            $endLine = $line
            break
        }
    }
    if ($endLine -le 0) {
        throw "対象プロシージャのEnd $endKeyword が見つかりません: $procedureName"
    }

    $lineCount = $endLine - $bodyLine + 1
    $codeModule.DeleteLines($bodyLine, $lineCount)
    $codeModule.InsertLines($bodyLine, $replacement)
    $workbook.Save()

    @{
        status = "applied"
        output = $output
        module = [string]$plan.operation.module_name
        procedure = $procedureName
        note = "xlblueprintへrevised.xlsmを戻して静的差分検証を実行してください。"
    } | ConvertTo-Json | Set-Content -LiteralPath $resultPath -Encoding UTF8
    Write-Host "VBA変更を適用しました: $output"
}
catch {
    @{
        status = "failed"
        error = $_.Exception.Message
    } | ConvertTo-Json | Set-Content -LiteralPath $resultPath -Encoding UTF8
    Write-Error $_.Exception.Message
    exit 1
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    if ($null -ne $workbook) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($null -ne $excel) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""


def _has_vba_signature(source_bytes: bytes) -> bool:
    """OOXMLパッケージにVBA署名パーツが含まれるか返す."""

    try:
        with zipfile.ZipFile(io.BytesIO(source_bytes)) as archive:
            return any(
                name.casefold().endswith("vbaprojectsignature.bin") for name in archive.namelist()
            )
    except zipfile.BadZipFile:
        return False


def build_vba_change_package(
    source_path: Path,
    plan: MutationPlan,
) -> bytes:
    """原本とWindows適用スクリプトを含むZIPを生成する.

    Args:
        source_path: 変更しない原本 `.xlsm`。
        plan: `vba_procedure_replace` 操作を含む変更計画。

    Returns:
        ZIPファイルのバイト列。

    Raises:
        VbaChangeError: 形式または操作がパッケージ対象外の場合。
    """

    if source_path.suffix.lower() != ".xlsm":
        raise VbaChangeError("VBA change packages require an .xlsm workbook")
    operation = plan.operation
    if not isinstance(operation, VbaProcedureReplaceOperation):
        raise VbaChangeError("VBA package requires a vba_procedure_replace operation")

    source_bytes = source_path.read_bytes()
    signature_present = _has_vba_signature(source_bytes)
    manifest = {
        **plan.model_dump(mode="json"),
        "vba_signature_present": signature_present,
        "static_verification_only": True,
    }
    readme_lines = [
        "xlblueprint VBA change package",
        "",
        "Requirements:",
        "- Windows with Microsoft Excel installed",
        "- Excel Trust Center: Trust access to the VBA project object model",
        "- An unlocked VBA project",
        "",
        "Steps:",
        "1. Extract this ZIP to a local folder.",
        "2. Review replacement.bas and plan.json.",
        "3. Run: powershell.exe -NoProfile -File .\\apply_vba_change.ps1",
        "4. Confirm revised.xlsm and result.json were created.",
        "5. Upload revised.xlsm to xlblueprint for static diff verification.",
        "",
        "Safety:",
        "- The script edits a copy named revised.xlsm; original.xlsm remains unchanged.",
        "- Excel events and macro execution are disabled while the workbook is opened.",
        "- The script does not compile or execute the changed macro.",
        "- Disable Trust access to the VBA project object model again after use.",
    ]
    if signature_present:
        readme_lines.extend(
            [
                "",
                "WARNING: A VBA digital signature was detected.",
                (
                    "Saving the changed workbook invalidates that signature. "
                    "Re-sign it before distribution."
                ),
            ]
        )

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("original.xlsm", source_bytes)
        archive.writestr(
            "replacement.bas",
            operation.new_code.replace("\r\n", "\n").replace("\r", "\n"),
        )
        archive.writestr(
            "plan.json",
            json.dumps(manifest, ensure_ascii=False, indent=2),
        )
        archive.writestr(
            "apply_vba_change.ps1",
            b"\xef\xbb\xbf" + _POWERSHELL_SCRIPT.encode("utf-8"),
        )
        archive.writestr("README.txt", "\n".join(readme_lines) + "\n")
    return output.getvalue()
