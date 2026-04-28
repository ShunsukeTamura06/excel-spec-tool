"""core.extractors.vba のテスト."""

from pathlib import Path

import pytest

from core.exceptions import ExtractionError
from core.extractors.vba import (
    _module_name_from_filename,
    _module_type_from_filename,
    _parse_procedures,
    extract_vba,
)


class TestParseProcedures:
    """`_parse_procedures` (内部ヘルパー) のテスト."""

    def test_simple_sub(self) -> None:
        code = 'Sub Hello()\n    MsgBox "hi"\nEnd Sub'
        procs = _parse_procedures(code)
        assert len(procs) == 1
        assert procs[0].name == "Hello"
        assert procs[0].kind == "Sub"
        assert procs[0].start_line == 1
        assert procs[0].end_line == 3

    def test_function(self) -> None:
        code = "Function Square(x As Long) As Long\n    Square = x * x\nEnd Function"
        procs = _parse_procedures(code)
        assert len(procs) == 1
        assert procs[0].kind == "Function"
        assert procs[0].name == "Square"

    def test_property_get(self) -> None:
        code = "Property Get TheAnswer() As Integer\n    TheAnswer = 42\nEnd Property"
        procs = _parse_procedures(code)
        assert len(procs) == 1
        assert procs[0].kind == "Property"
        assert procs[0].name == "TheAnswer"

    def test_all_three_kinds(self) -> None:
        code = (
            "Sub A()\nEnd Sub\n"
            "\n"
            "Function B() As Long\nEnd Function\n"
            "\n"
            "Property Let C(v As Long)\nEnd Property\n"
        )
        procs = _parse_procedures(code)
        assert [p.kind for p in procs] == ["Sub", "Function", "Property"]
        assert [p.name for p in procs] == ["A", "B", "C"]

    def test_public_private_static_modifiers(self) -> None:
        code = (
            "Public Sub A()\nEnd Sub\n"
            "Private Function B() As Long\nEnd Function\n"
            "Friend Static Sub C()\nEnd Sub\n"
        )
        procs = _parse_procedures(code)
        assert [p.name for p in procs] == ["A", "B", "C"]

    def test_no_procedures(self) -> None:
        code = "Option Explicit\n' just a comment\nDim x As Long\n"
        assert _parse_procedures(code) == []

    def test_line_numbers_are_1_based_and_inclusive(self) -> None:
        code = (
            "Option Explicit\n"  # line 1
            "\n"  # line 2
            "Sub Hello()\n"  # line 3
            '    MsgBox "hi"\n'  # line 4
            "End Sub\n"  # line 5
        )
        procs = _parse_procedures(code)
        assert len(procs) == 1
        assert procs[0].start_line == 3
        assert procs[0].end_line == 5
        # コードは start..end の行を結合したもの
        assert "Sub Hello()" in procs[0].code
        assert "End Sub" in procs[0].code
        assert "Option Explicit" not in procs[0].code

    def test_multiple_procedures_in_one_module(self) -> None:
        code = "Sub A()\n    Debug.Print 1\nEnd Sub\n\nSub B()\n    Debug.Print 2\nEnd Sub\n"
        procs = _parse_procedures(code)
        assert len(procs) == 2
        assert procs[0].name == "A"
        assert procs[1].name == "B"


class TestModuleNameAndType:
    def test_bas_is_module(self) -> None:
        assert _module_type_from_filename("Module1.bas") == "Module"

    def test_cls_is_class(self) -> None:
        assert _module_type_from_filename("Class1.cls") == "Class"

    def test_frm_is_form(self) -> None:
        assert _module_type_from_filename("UserForm1.frm") == "Form"

    def test_unknown_defaults_to_module(self) -> None:
        assert _module_type_from_filename("Foo") == "Module"

    def test_name_strips_extension(self) -> None:
        assert _module_name_from_filename("Module1.bas") == "Module1"
        assert _module_name_from_filename("ThisWorkbook.cls") == "ThisWorkbook"


class TestExtractVbaFromBas:
    """olevba の TEXT モードを使った統合テスト. .bas プレーンテキストfixtureに対して動作."""

    def test_simple_macro_bas(self, fixtures_dir: Path) -> None:
        path = fixtures_dir / "simple_macro.bas"
        modules = extract_vba(path)

        assert len(modules) == 1
        m = modules[0]
        assert m.type == "Module"

        # Sub / Function / Property がそれぞれ検出されている
        kinds = sorted(p.kind for p in m.procedures)
        assert kinds == ["Function", "Property", "Sub"]
        names = {p.name for p in m.procedures}
        assert names == {"Hello", "Square", "TheAnswer"}

    def test_another_module_bas(self, fixtures_dir: Path) -> None:
        path = fixtures_dir / "another_module.bas"
        modules = extract_vba(path)

        assert len(modules) == 1
        m = modules[0]
        proc_names = {p.name for p in m.procedures}
        assert proc_names == {"UpdateDaily", "IsValid"}


class TestExtractVbaNoMacros:
    """VBAを含まないxlsxは空リストを返す."""

    def test_xlsx_without_vba(self, empty_xlsx: Path) -> None:
        result = extract_vba(empty_xlsx)
        assert result == []


class TestExtractVbaErrors:
    def test_missing_file_raises(self, tmp_path: Path) -> None:
        with pytest.raises(ExtractionError):
            extract_vba(tmp_path / "does_not_exist.xlsm")


# TODO: 実 .xlsm / .xls バイナリでの統合テスト. パスワード保護 VBA のテスト.
#       現状は LibreOffice headless での VBA 注入が安定しないため見送り。
#       Shun が手動で fixture を追加した時点で有効化する。
