"""backend.annotators のテスト (Phase C)."""

from __future__ import annotations

from typing import Any

import pytest

from backend.annotators import (
    _VBA_CODE_MAX_CHARS,
    _format_procedure_brief,
    _format_sheet_brief,
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
    """annotate_text の呼び出し履歴を記録し、決定的な応答を返すスタブ.

    LLMClient プロトコル全体は実装しない (annotate_workbook が使うのは annotate_text のみ)。
    """

    def __init__(self, response: str = "STUB ANNOTATION", raise_on_call: bool = False) -> None:
        self.calls: list[dict[str, Any]] = []
        self.response = response
        self.raise_on_call = raise_on_call

    def annotate_text(self, prompt: str, content: str, tier: str = "fast") -> str:
        self.calls.append({"prompt": prompt, "content": content, "tier": tier})
        if self.raise_on_call:
            raise RuntimeError("simulated LLM failure")
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
        llm = _RecordingLLM(response="日次集計")
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        assert result.sheets[0].purpose == "日次集計"
        assert result.sheets[1].purpose == "日次集計"
        assert len(llm.calls) == 2

    def test_skips_already_annotated_sheets(self) -> None:
        wb = Workbook(
            filename="t.xlsm",
            sheets=[
                SheetInfo(name="S1", rows=1, cols=1, purpose="既存"),
                SheetInfo(name="S2", rows=1, cols=1),
            ],
        )
        llm = _RecordingLLM(response="new")
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
        llm = _RecordingLLM(response="日次処理")
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        procs = result.vba_modules[0].procedures
        assert procs[0].annotation == "日次処理"
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
        llm = _RecordingLLM(response="new")
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
        llm = _RecordingLLM(response="new purpose")
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        # 元の SheetInfo は変更されていない
        assert wb.sheets[0].purpose == ""
        assert result.sheets[0].purpose == "new purpose"

    def test_strips_whitespace_from_llm_response(self) -> None:
        wb = Workbook(filename="t.xlsm", sheets=[SheetInfo(name="S", rows=1, cols=1)])
        llm = _RecordingLLM(response="  \n trimmed \n  ")
        result = annotate_workbook(wb, llm)  # type: ignore[arg-type]
        assert result.sheets[0].purpose == "trimmed"
