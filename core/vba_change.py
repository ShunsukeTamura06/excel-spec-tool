"""既存VBAプロシージャの安全な全置換計画を作る."""

from __future__ import annotations

from core.exceptions import VbaChangeError
from core.extractors.vba import _PROC_END_RE, _parse_procedures
from core.models import VbaModuleDiff, Workbook, WorkbookDiff


def normalize_vba_code(code: str) -> str:
    """VBAコードの改行と末尾空白を正規化する.

    Args:
        code: 正規化対象のVBAコード。

    Returns:
        LF改行で、各行末の空白と末尾の空行を除いたコード。
    """

    lines = code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    return "\n".join(line.rstrip() for line in lines).rstrip("\n")


def propose_vba_procedure_replace(
    workbook: Workbook,
    module_name: str,
    procedure_name: str,
    new_code: str,
) -> WorkbookDiff:
    """既存Sub/Functionを完全な新コードへ置換した期待差分を作る.

    Args:
        workbook: 変更前の抽出済みワークブック。
        module_name: 対象VBAモジュール名。
        procedure_name: 対象SubまたはFunction名。
        new_code: 宣言行からEnd行までを含む置換後コード。

    Returns:
        対象モジュール1件だけをmodifiedとする期待構造差分。

    Raises:
        VbaChangeError: 対象不在、重複、Property、宣言不一致、不正コードの場合。
    """

    if not new_code.strip():
        raise VbaChangeError("replacement VBA code must not be empty")
    if len(new_code) > 12_000:
        raise VbaChangeError("replacement VBA code is too large")
    if "Attribute VB_" in new_code:
        raise VbaChangeError("replacement code must not contain hidden VBA attributes")

    modules = [
        item for item in workbook.vba_modules if item.name.casefold() == module_name.casefold()
    ]
    if len(modules) != 1:
        raise VbaChangeError(f"VBA module must exist exactly once: {module_name}")
    module = modules[0]
    procedures = [
        item for item in module.procedures if item.name.casefold() == procedure_name.casefold()
    ]
    if len(procedures) != 1:
        raise VbaChangeError(
            f"VBA procedure must exist exactly once: {module.name}.{procedure_name}"
        )
    before_procedure = procedures[0]
    if before_procedure.kind == "Property":
        raise VbaChangeError("Property procedures are not supported in the first VBA package")

    parsed = _parse_procedures(normalize_vba_code(new_code))
    if len(parsed) != 1:
        raise VbaChangeError("replacement code must contain exactly one complete procedure")
    replacement = parsed[0]
    if replacement.name.casefold() != before_procedure.name.casefold():
        raise VbaChangeError(f"replacement procedure name must remain {before_procedure.name}")
    if replacement.kind != before_procedure.kind:
        raise VbaChangeError(f"replacement procedure kind must remain {before_procedure.kind}")
    # _parse_procedures は抽出用途では「End が無ければ次の宣言 or ファイル末尾で
    # 打ち切る」寛容な挙動をする (End 抜けの実ファイルを読めるようにするため)。
    # ここでは適用対象そのものを検証しているため、その寛容さは事故のもとになる
    # (End Sub/Function が無いコードでもここを通過し、壊れたモジュールを
    # 「差分一致」として提示してしまう)。最終行が対応する End 文であることを
    # 明示的に確認する。
    replacement_lines = [line for line in replacement.code.splitlines() if line.strip()]
    end_match = replacement_lines and _PROC_END_RE.match(replacement_lines[-1])
    if not end_match or end_match.group("kind").capitalize() != replacement.kind:
        raise VbaChangeError(
            f"replacement code must end with a matching End {replacement.kind} statement"
        )

    before_module_code = normalize_vba_code(module.code)
    lines = before_module_code.split("\n")
    start_index = before_procedure.start_line - 1
    end_index = before_procedure.end_line
    if start_index < 0 or end_index > len(lines) or start_index >= end_index:
        raise VbaChangeError("extracted VBA procedure line range is invalid")
    current_slice = normalize_vba_code("\n".join(lines[start_index:end_index]))
    if before_procedure.code.strip() and current_slice != normalize_vba_code(before_procedure.code):
        raise VbaChangeError("extracted VBA procedure no longer matches its module")

    replacement_lines = normalize_vba_code(new_code).split("\n")
    after_module_code = "\n".join([*lines[:start_index], *replacement_lines, *lines[end_index:]])
    return WorkbookDiff(
        before_filename=workbook.filename,
        after_filename=workbook.filename,
        vba_modules=[
            VbaModuleDiff(
                name=module.name,
                change_type="modified",
                before_code=before_module_code,
                after_code=normalize_vba_code(after_module_code),
            )
        ],
        existing_risks=list(workbook.analysis_risks),
    )
