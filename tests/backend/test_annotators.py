"""backend.annotators のテスト (Phase C + P1-2 構造化)."""

from __future__ import annotations

import json
from typing import Any

import pytest

from backend.annotators import (
    _VBA_CODE_MAX_CHARS,
    _format_procedure_brief,
    _format_sheet_brief,
    _list_str_field,
    _parse_llm_json,
    _str_field,
    annotate_workbook,
)
from core.models import (
    CellFormula,
    NamedRange,
    SheetInfo,
    VbaModule,
    VbaProcedure,
    Workbook,
)


class _RecordingLLM:
    """annotate_text の呼び出し履歴を記録し、決定的な JSON 応答を返すスタブ.

    `response` は str (そのまま返す) か dict (JSON にして返す) を受け取る.
    LLMClient プロトコル全体は実装しない.
    """

    def __init__(
        self,
        response: str | dict[str, Any] | None = None,
        raise_on_call: bool = False,
    ) -> None:
        self.calls: list[dict[str, Any]] = []
        self.response = response if response is not None else {"purpose": "STUB"}
        self.raise_on_call = raise_on_call

    def annotate_text(self, prompt: str, content: str, tier: str = "fast") -> str:
        self.calls.append({"prompt": prompt, "content": content, "tier": tier})
        if self.raise_on_call:
            raise RuntimeError("simulated LLM failure")
        if isinstance(self.response, dict):
            return json.dumps(self.response, ensure_ascii=False)
        return self.response


# ---------- formatting helpers ----------


class TestFormatSheetBrief:
    def test_includes_basic_fields(self) -> None:
        s = SheetInfo(
            name="Portfolio",
            rows=100,
            cols=10,
            formulas=[CellFormula(coord="Portfolio!A1", formula="=SUM(B1:B10)", refs=[])],
            named_ranges=[NamedRange(name="TaxRate", refers_to="Portfolio!$A$1")],
            preview_rows=[["銘柄", "金額"], ["ABC", "100"]],
            preview_origin="A1:B2",
        )
        text = _format_sheet_brief(s)
        assert "Portfolio" in text
        assert "100" in text  # rows
        assert "SUM" in text
        assert "TaxRate" in text
        assert "銘柄" in text

    def test_empty_sheet_does_not_crash(self) -> None:
        text = _format_sheet_brief(SheetInfo(name="E", rows=0, cols=0))
        assert "E" in text


class TestFormatProcedureBrief:
    def test_uses_proc_code_when_set(self) -> None:
        m = VbaModule(name="M", type="Module", code="ignored")
        p = VbaProcedure(
            name="P",
            kind="Sub",
            start_line=1,
            end_line=2,
            code="Sub P()\nEnd Sub",
        )
        text = _format_procedure_brief(m, p)
        assert "Sub P()" in text
        assert "ignored" not in text  # proc.code が優先される

    def test_falls_back_to_module_lines(self) -> None:
        m = VbaModule(
            name="M",
            type="Module",
            code="Sub A()\nEnd Sub\nSub B()\nEnd Sub\n",
        )
        p = VbaProcedure(name="B", kind="Sub", start_line=3, end_line=4, code="")
        text = _format_procedure_brief(m, p)
        assert "Sub B()" in text
        assert "Sub A()" not in text

    def test_truncates_huge_code(self) -> None:
        big_code = "Sub Big()\n" + ('    Range("A1")\n' * 5000) + "End Sub\n"
        m = VbaModule(name="M", type="Module", code="")
        p = VbaProcedure(
            name="Big",
            kind="Sub",
            start_line=1,
            end_line=10,
            code=big_code,
        )
        text = _format_procedure_brief(m, p)
        # コード本体は _VBA_CODE_MAX_CHARS 以下に切られる
        # (text 全体には他のメタも含むが、コード部分は確実に切れている)
        assert len(text) < len(big_code) + 500
        assert "冒頭のみを抜粋" in text
        assert _VBA_CODE_MAX_CHARS >= 1000  # sanity


# ---------- annotate_workbook ----------


class TestAnnotateWorkbook:
    def test_annotates_empty_workbook(self) -> None:
        llm = _RecordingLLM()
        annotated = annotate_workbook(Workbook(filename="t.xlsm"), llm)  # type: ignore[arg-type]
        assert annotated.sheets == []
        assert annotated.vba_modules == []
        assert llm.calls == []

    def test_annotates_each_sheet(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(name="S1", rows=1, cols=1),
                SheetInfo(name="S2", rows=1, cols=1),
            ],
        )
        llm = _RecordingLLM(
            response={
                "purpose": "日次集計",
                "inputs": ["Input"],
                "outputs": ["Output"],
                "main_calculations": ["SUMIF で集計"],
                "usage_scenario": "毎朝担当者が確認",
            },
        )
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        s1, s2 = result.sheets
        assert s1.purpose == "日次集計"
        assert s1.inputs == ["Input"]
        assert s1.outputs == ["Output"]
        assert s1.main_calculations == ["SUMIF で集計"]
        assert s1.usage_scenario == "毎朝担当者が確認"
        # 2 枚目も同じ JSON を返すので埋まる
        assert s2.purpose == "日次集計"
        assert len(llm.calls) == 2

    def test_skips_already_annotated_sheets(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(name="S1", rows=1, cols=1, purpose="既存"),
                SheetInfo(name="S2", rows=1, cols=1),
            ],
        )
        llm = _RecordingLLM(response={"purpose": "new"})
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        assert result.sheets[0].purpose == "既存"  # 既存は維持
        assert result.sheets[1].purpose == "new"
        # 1 件だけ呼ばれた (既存は skip)
        assert len(llm.calls) == 1

    def test_annotates_procedures(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            vba_modules=[
                VbaModule(
                    name="M",
                    type="Module",
                    code="",
                    procedures=[
                        VbaProcedure(
                            name="P1",
                            kind="Sub",
                            start_line=1,
                            end_line=2,
                            code="Sub P1()\nEnd Sub",
                        ),
                        VbaProcedure(
                            name="P2",
                            kind="Sub",
                            start_line=3,
                            end_line=4,
                            code="Sub P2()\nEnd Sub",
                        ),
                    ],
                )
            ],
        )
        llm = _RecordingLLM(
            response={
                "annotation": "日次処理",
                "side_effects": ["Calc!A:Z"],
                "triggers": ["ボタン"],
                "calls": ["Helper"],
            },
        )
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        procs = result.vba_modules[0].procedures
        assert procs[0].annotation == "日次処理"
        assert procs[0].side_effects == ["Calc!A:Z"]
        assert procs[0].triggers == ["ボタン"]
        assert procs[0].calls == ["Helper"]
        assert procs[1].annotation == "日次処理"
        assert len(llm.calls) == 2

    def test_skips_already_annotated_procedures(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            vba_modules=[
                VbaModule(
                    name="M",
                    type="Module",
                    code="",
                    procedures=[
                        VbaProcedure(
                            name="P1",
                            kind="Sub",
                            start_line=1,
                            end_line=2,
                            code="x",
                            annotation="既存",
                        ),
                        VbaProcedure(name="P2", kind="Sub", start_line=3, end_line=4, code="x"),
                    ],
                )
            ],
        )
        llm = _RecordingLLM(response={"annotation": "new"})
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        procs = result.vba_modules[0].procedures
        assert procs[0].annotation == "既存"
        assert procs[1].annotation == "new"
        assert len(llm.calls) == 1

    def test_llm_failure_leaves_empty_annotation(self, caplog: pytest.LogCaptureFixture) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[SheetInfo(name="S", rows=1, cols=1)],
        )
        llm = _RecordingLLM(raise_on_call=True)
        with caplog.at_level("WARNING", logger="backend.annotators"):
            result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        # 失敗時は空のまま、例外は出ない
        assert result.sheets[0].purpose == ""
        # warning ログが残る
        assert any("annotate failed" in r.message for r in caplog.records)

    def test_returns_new_workbook_does_not_mutate(self) -> None:
        sheet = SheetInfo(name="S", rows=1, cols=1)
        wb = Workbook(filename="t.xlsm", sheets=[sheet])
        llm = _RecordingLLM(response={"purpose": "new purpose"})
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        # 元の SheetInfo は変更されていない
        assert wb.sheets[0].purpose == ""
        assert result.sheets[0].purpose == "new purpose"

    def test_strips_whitespace_within_json_values(self) -> None:
        wb = Workbook(filename="t.xlsm", sheets=[SheetInfo(name="S", rows=1, cols=1)])
        llm = _RecordingLLM(
            response={"purpose": "  \n trimmed \n  ", "inputs": ["  In  ", ""]},
        )
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        # JSON 値内の前後空白は除去される. list 内の空文字は filter される.
        assert result.sheets[0].purpose == "trimmed"
        assert result.sheets[0].inputs == ["In"]

    def test_malformed_json_leaves_fields_empty(
        self, caplog: pytest.LogCaptureFixture
    ) -> None:
        wb = Workbook(filename="t.xlsm", sheets=[SheetInfo(name="S", rows=1, cols=1)])
        llm = _RecordingLLM(response="this is not JSON at all")
        with caplog.at_level("WARNING", logger="backend.annotators"):
            result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        assert result.sheets[0].purpose == ""
        assert result.sheets[0].inputs == []
        assert any("JSON parse failed" in r.message for r in caplog.records)

    def test_json_in_code_fence_is_parsed(self) -> None:
        """LLM が ```json ... ``` で包んでも拾えること."""
        wb = Workbook(filename="t.xlsm", sheets=[SheetInfo(name="S", rows=1, cols=1)])
        llm = _RecordingLLM(
            response='```json\n{"purpose": "fenced ok"}\n```',
        )
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        assert result.sheets[0].purpose == "fenced ok"

    def test_extracts_json_with_leading_chatter(self) -> None:
        """LLM が前置きを付けても JSON 部分だけ拾える."""
        wb = Workbook(filename="t.xlsm", sheets=[SheetInfo(name="S", rows=1, cols=1)])
        llm = _RecordingLLM(
            response='了解しました。以下が結果です:\n{"purpose": "extracted"}',
        )
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        assert result.sheets[0].purpose == "extracted"


# ---------- JSON パーサ単体 ----------


class TestParseLlmJson:
    def test_plain_json(self) -> None:
        assert _parse_llm_json('{"a": 1}') == {"a": 1}

    def test_code_fence(self) -> None:
        assert _parse_llm_json('```json\n{"a": 1}\n```') == {"a": 1}

    def test_with_leading_text(self) -> None:
        assert _parse_llm_json('blah blah {"a": 1} trailing') == {"a": 1}

    def test_invalid_returns_none(self) -> None:
        assert _parse_llm_json("not json") is None
        assert _parse_llm_json("") is None

    def test_array_at_top_level_returns_none(self) -> None:
        # トップレベル配列は dict ではないので None.
        assert _parse_llm_json("[1, 2, 3]") is None


class TestFieldExtractors:
    def test_str_field_handles_missing_and_wrong_type(self) -> None:
        assert _str_field({"x": "hi"}, "x") == "hi"
        assert _str_field({"x": 1}, "x") == ""
        assert _str_field({}, "x") == ""
        assert _str_field({"x": "  spaced  "}, "x") == "spaced"

    def test_list_str_field_filters_non_strings_and_empty(self) -> None:
        payload: dict[str, Any] = {"xs": ["a", "", "  ", "b", 1, None, "c"]}
        assert _list_str_field(payload, "xs") == ["a", "b", "c"]

    def test_list_str_field_respects_limit(self) -> None:
        payload: dict[str, Any] = {"xs": [str(i) for i in range(50)]}
        assert len(_list_str_field(payload, "xs", limit=5)) == 5
