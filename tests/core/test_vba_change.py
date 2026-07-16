"""VBAプロシージャ置換計画のテスト."""

from __future__ import annotations

import pytest

from core.exceptions import VbaChangeError
from core.models import VbaModule, VbaProcedure, Workbook
from core.vba_change import normalize_vba_code, propose_vba_procedure_replace


def _workbook() -> Workbook:
    code = (
        'Attribute VB_Name = "Module1"\r\n'
        "Option Explicit\r\n"
        "\r\n"
        "Public Sub UpdateReport()\r\n"
        '    Range("A1").Value = 1\r\n'
        "End Sub\r\n"
        "\r\n"
        "Public Function DoubleValue(ByVal value As Long) As Long\r\n"
        "    DoubleValue = value * 2\r\n"
        "End Function\r\n"
    )
    return Workbook(
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
                        start_line=4,
                        end_line=6,
                        code=('Public Sub UpdateReport()\n    Range("A1").Value = 1\nEnd Sub'),
                    ),
                    VbaProcedure(
                        name="DoubleValue",
                        kind="Function",
                        start_line=8,
                        end_line=10,
                        code=(
                            "Public Function DoubleValue(ByVal value As Long) As Long\n"
                            "    DoubleValue = value * 2\n"
                            "End Function"
                        ),
                    ),
                ],
            )
        ],
    )


def test_proposes_exact_module_change() -> None:
    """対象プロシージャ以外を維持したモジュール差分を作る."""

    diff = propose_vba_procedure_replace(
        _workbook(),
        "Module1",
        "UpdateReport",
        (
            "Public Sub UpdateReport()\n"
            '    Range("A1").Value = 2\n'
            '    Range("A2").Value = "done"\n'
            "End Sub"
        ),
    )

    assert len(diff.vba_modules) == 1
    change = diff.vba_modules[0]
    assert change.name == "Module1"
    assert 'Range("A1").Value = 1' in (change.before_code or "")
    assert 'Range("A1").Value = 2' in (change.after_code or "")
    assert "DoubleValue = value * 2" in (change.after_code or "")
    assert "\r" not in (change.after_code or "")


def test_rejects_different_procedure_name() -> None:
    """別名のプロシージャへの置換を拒否する."""

    with pytest.raises(VbaChangeError, match="name must remain"):
        propose_vba_procedure_replace(
            _workbook(),
            "Module1",
            "UpdateReport",
            "Public Sub OtherName()\nEnd Sub",
        )


def test_rejects_partial_code() -> None:
    """宣言・Endを含まない断片コードを拒否する."""

    with pytest.raises(VbaChangeError, match="exactly one complete procedure"):
        propose_vba_procedure_replace(
            _workbook(),
            "Module1",
            "UpdateReport",
            'Range("A1").Value = 2',
        )


def test_normalize_vba_code_ignores_line_ending_noise() -> None:
    """CRLF/LFと行末空白を同一のVBAコードとして正規化する."""

    assert normalize_vba_code("Sub A()\r\nEnd Sub  \r\n") == "Sub A()\nEnd Sub"
