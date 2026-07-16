"""変更計画とopenpyxlプロバイダー境界のテスト."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest
from openpyxl.workbook.defined_name import DefinedName

from core.exceptions import UnsupportedMutationError
from core.extractors.workbook import extract_workbook
from core.mutation import (
    MutationPlan,
    NamedRangeSetOperation,
    OpenPyxlMutationProvider,
    propose_mutation,
)
from core.reference_index import build_reference_index


def _named_range_workbook(path: Path) -> None:
    """名前定義を持つxlsxを作成する."""

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet["A1"] = 1
    workbook.defined_names["TaxRate"] = DefinedName(name="TaxRate", attr_text="Data!$A$1")
    workbook.save(path)
    workbook.close()


def test_propose_and_apply_share_one_plan(tmp_path: Path) -> None:
    """同じ変更計画を期待差分生成と実適用の両方に使える."""

    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    _named_range_workbook(source)
    workbook = extract_workbook(source)
    index = build_reference_index(workbook)
    plan = MutationPlan(
        source_job_id="00000000-0000-4000-8000-000000000000",
        operation=NamedRangeSetOperation(
            name="TaxRate",
            new_refers_to="Data!$A$2",
        ),
    )

    expected = propose_mutation(plan, workbook, index)
    result = OpenPyxlMutationProvider().apply(plan, source, output)

    assert expected.named_ranges[0].after_refers_to == "Data!$A$2"
    assert result.provider == "openpyxl"
    assert result.changed_count == 1
    assert source.read_bytes() != output.read_bytes()
    changed = openpyxl.load_workbook(output)
    assert changed.defined_names["TaxRate"].attr_text == "Data!$A$2"
    changed.close()


def test_openpyxl_capability_is_explicit() -> None:
    """対応形式と操作をcapabilityとして機械取得できる."""

    capability = OpenPyxlMutationProvider().capability()

    assert capability.available
    assert capability.supported_extensions == [".xlsx", ".xlsm"]
    assert set(capability.supported_operations) == {
        "named_range_set",
        "fixed_ref_replace",
        "range_expansion",
    }


def test_openpyxl_rejects_unsupported_extension(tmp_path: Path) -> None:
    """対応外形式を曖昧に処理せず明示的に拒否する."""

    source = tmp_path / "source.xls"
    source.write_bytes(b"not-an-xls")
    plan = MutationPlan(
        source_job_id="00000000-0000-4000-8000-000000000000",
        operation=NamedRangeSetOperation(name="TaxRate", new_refers_to="Data!$A$2"),
    )

    with pytest.raises(UnsupportedMutationError, match="does not support"):
        OpenPyxlMutationProvider().apply(plan, source, tmp_path / "output.xls")
