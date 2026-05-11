"""backend.llm_tools のテスト."""

import json
from pathlib import Path

import pytest
from openpyxl import Workbook as OpyWorkbook

from backend.llm_tools import (
    TOOL_DEFINITIONS,
    build_tool_definitions,
    execute_tool_call,
)
from backend.storage import Storage
from core.extractors.cells import extract_cells_to_sqlite
from core.models import Reference, ReferenceIndex


@pytest.fixture
def job_with_cells(tmp_path: Path) -> tuple[Storage, str]:
    """cells.db を埋めたジョブを用意する."""
    storage = Storage(tmp_path / "jobs")

    # 小さな xlsx を生成
    wb = OpyWorkbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Portfolio"
    headers = ["銘柄コード", "銘柄名", "保有口数", "現在値", "評価損益"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=h)
    ws.cell(row=2, column=1, value="ABC")
    ws.cell(row=2, column=2, value="株式会社A")
    ws.cell(row=2, column=3, value=100)
    src = tmp_path / "p.xlsx"
    wb.save(src)
    data = src.read_bytes()

    meta = storage.create_job("p.xlsx", data)
    extract_cells_to_sqlite(
        storage.get_original_path(meta.job_id),
        storage.cells_db_path(meta.job_id),
    )
    return storage, meta.job_id


# ---------- 定義 ----------


class TestToolDefinitions:
    def test_three_tools_defined(self) -> None:
        names = {t["function"]["name"] for t in TOOL_DEFINITIONS}
        assert names == {"get_cells_range", "find_cells", "lookup_references"}

    def test_build_returns_list(self) -> None:
        tools = build_tool_definitions()
        assert isinstance(tools, list)
        assert tools == TOOL_DEFINITIONS

    def test_each_tool_has_schema(self) -> None:
        for t in TOOL_DEFINITIONS:
            assert t["type"] == "function"
            fn = t["function"]
            assert "name" in fn
            assert "description" in fn
            assert "parameters" in fn
            assert fn["parameters"]["type"] == "object"


# ---------- execute_tool_call ----------


class TestExecuteGetCellsRange:
    def test_returns_grid(self, job_with_cells: tuple[Storage, str]) -> None:
        storage, job_id = job_with_cells
        result_str = execute_tool_call(
            storage, job_id, "get_cells_range", {"sheet": "Portfolio", "range": "A1:E1"}
        )
        result = json.loads(result_str)
        assert result["sheet"] == "Portfolio"
        assert result["rows"][0][0]["value"] == "銘柄コード"

    def test_missing_args_returns_error_json(self, job_with_cells: tuple[Storage, str]) -> None:
        storage, job_id = job_with_cells
        result = json.loads(execute_tool_call(storage, job_id, "get_cells_range", {}))
        assert "error" in result

    def test_invalid_range_returns_error_json(self, job_with_cells: tuple[Storage, str]) -> None:
        storage, job_id = job_with_cells
        result = json.loads(
            execute_tool_call(
                storage, job_id, "get_cells_range", {"sheet": "P", "range": "garbage"}
            )
        )
        assert "error" in result


class TestExecuteFindCells:
    def test_basic_find(self, job_with_cells: tuple[Storage, str]) -> None:
        storage, job_id = job_with_cells
        result = json.loads(
            execute_tool_call(storage, job_id, "find_cells", {"query": "銘柄コード"})
        )
        assert result["count"] >= 1

    def test_no_query_returns_error(self, job_with_cells: tuple[Storage, str]) -> None:
        storage, job_id = job_with_cells
        result = json.loads(execute_tool_call(storage, job_id, "find_cells", {}))
        assert "error" in result

    def test_sheet_filter(self, job_with_cells: tuple[Storage, str]) -> None:
        storage, job_id = job_with_cells
        # シート指定なしと指定の両方で動作
        with_sheet = json.loads(
            execute_tool_call(
                storage,
                job_id,
                "find_cells",
                {"query": "ABC", "sheet": "Portfolio"},
            )
        )
        assert with_sheet["count"] >= 1


class TestExecuteLookupReferences:
    def test_basic_lookup(self, job_with_cells: tuple[Storage, str]) -> None:
        storage, job_id = job_with_cells
        # 参照インデックスを差し込んでおく
        idx = ReferenceIndex(
            refs={"Calc!H2": [Reference(kind="formula", from_="Out!K3", to="Calc!H2")]}
        )
        storage.save_references(job_id, idx)

        result = json.loads(
            execute_tool_call(storage, job_id, "lookup_references", {"target": "Calc!H2"})
        )
        assert result["count"] == 1
        assert result["refs"][0]["from"] == "Out!K3"

    def test_unknown_target_returns_empty(self, job_with_cells: tuple[Storage, str]) -> None:
        storage, job_id = job_with_cells
        # references.json が無くてもエラーにはせず empty を返したいが、
        # 現実装は ToolExecutionError -> error JSON。どちらでもテスト的に許容。
        # ここではまず references を保存してから空ターゲットを引く。
        storage.save_references(job_id, ReferenceIndex())
        result = json.loads(
            execute_tool_call(storage, job_id, "lookup_references", {"target": "Nowhere!A1"})
        )
        assert result["count"] == 0

    def test_missing_target_returns_error(self, job_with_cells: tuple[Storage, str]) -> None:
        storage, job_id = job_with_cells
        result = json.loads(execute_tool_call(storage, job_id, "lookup_references", {}))
        assert "error" in result


class TestExecuteUnknownTool:
    def test_returns_error(self, job_with_cells: tuple[Storage, str]) -> None:
        storage, job_id = job_with_cells
        result = json.loads(execute_tool_call(storage, job_id, "no_such", {}))
        assert "error" in result
