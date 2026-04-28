"""core.models のテスト."""

import pytest
from pydantic import ValidationError

from core.models import (
    CellFormula,
    ConditionalFormat,
    JobMeta,
    NamedRange,
    Reference,
    ReferenceIndex,
    SheetInfo,
    VbaModule,
    VbaProcedure,
    Workbook,
)


class TestVbaProcedure:
    def test_valid(self) -> None:
        p = VbaProcedure(
            name="UpdateDaily",
            kind="Sub",
            start_line=10,
            end_line=50,
            code="Sub UpdateDaily()\nEnd Sub",
        )
        assert p.name == "UpdateDaily"
        assert p.annotation == ""

    def test_kind_literal_rejects_invalid(self) -> None:
        with pytest.raises(ValidationError):
            VbaProcedure(
                name="X",
                kind="Macro",
                start_line=1,
                end_line=2,
                code="",  # type: ignore[arg-type]
            )


class TestVbaModule:
    def test_default_procedures_is_empty_list(self) -> None:
        m = VbaModule(name="Module1", type="Module", code="")
        assert m.procedures == []

    def test_each_instance_has_independent_default_list(self) -> None:
        m1 = VbaModule(name="A", type="Module", code="")
        m2 = VbaModule(name="B", type="Module", code="")
        m1.procedures.append(VbaProcedure(name="P", kind="Sub", start_line=1, end_line=2, code=""))
        assert m2.procedures == []


class TestCellFormula:
    def test_valid(self) -> None:
        f = CellFormula(
            coord="Calc!H2",
            formula="=SUMIF(Input!A:A, A2, Input!E:E)",
            refs=["Input!A:A", "Calc!A2", "Input!E:E"],
        )
        assert len(f.refs) == 3


class TestNamedRange:
    def test_valid(self) -> None:
        n = NamedRange(name="顧客マスタ", refers_to="Calc!$A$2:$D$5000")
        assert n.name == "顧客マスタ"


class TestConditionalFormat:
    def test_valid(self) -> None:
        c = ConditionalFormat(range="A1:A10", rule="cellIs greaterThan 100")
        assert c.range == "A1:A10"


class TestSheetInfo:
    def test_minimal(self) -> None:
        s = SheetInfo(name="Calc", rows=100, cols=10)
        assert s.formulas == []
        assert s.named_ranges == []
        assert s.conditional_formats == []
        assert s.purpose == ""


class TestWorkbook:
    def test_minimal(self) -> None:
        w = Workbook(filename="test.xlsm")
        assert w.sheets == []
        assert w.vba_modules == []
        assert w.external_links == []

    def test_full_tree_serialization_roundtrip(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(
                    name="Calc",
                    rows=10,
                    cols=5,
                    formulas=[CellFormula(coord="Calc!A1", formula="=1+1", refs=[])],
                ),
            ],
            vba_modules=[
                VbaModule(
                    name="M1",
                    type="Module",
                    code="Sub X()\nEnd Sub",
                    procedures=[
                        VbaProcedure(name="X", kind="Sub", start_line=1, end_line=2, code="Sub X()")
                    ],
                )
            ],
            external_links=["other.xlsx"],
        )
        dumped = wb.model_dump_json()
        restored = Workbook.model_validate_json(dumped)
        assert restored == wb


class TestReference:
    def test_construct_with_python_name(self) -> None:
        r = Reference(kind="formula", from_="Output!K3", to="Calc!H2", code="=Calc!H2")
        assert r.from_ == "Output!K3"

    def test_construct_with_json_alias(self) -> None:
        r = Reference.model_validate(
            {"kind": "vba", "from": "Module1.UpdateDaily:L47", "to": "Calc!A1", "code": ""}
        )
        assert r.from_ == "Module1.UpdateDaily:L47"

    def test_dump_uses_alias(self) -> None:
        r = Reference(kind="formula", from_="A!1", to="B!1")
        dumped = r.model_dump(by_alias=True)
        assert "from" in dumped
        assert "from_" not in dumped


class TestReferenceIndex:
    def test_default_refs_is_empty(self) -> None:
        idx = ReferenceIndex()
        assert idx.refs == {}

    def test_lookup(self) -> None:
        idx = ReferenceIndex(
            refs={
                "Calc!H2": [
                    Reference(kind="formula", from_="Output!K3", to="Calc!H2"),
                    Reference(kind="vba", from_="Module1.X:L10", to="Calc!H2"),
                ]
            }
        )
        assert len(idx.refs["Calc!H2"]) == 2


class TestJobMeta:
    def test_valid(self) -> None:
        m = JobMeta(
            job_id="abc-123",
            filename="t.xlsm",
            created_at="2026-04-28T15:00:00",
            status="uploaded",
        )
        assert m.status == "uploaded"

    def test_status_literal_rejects_invalid(self) -> None:
        with pytest.raises(ValidationError):
            JobMeta(
                job_id="x",
                filename="x",
                created_at="x",
                status="processing",  # type: ignore[arg-type]
            )
