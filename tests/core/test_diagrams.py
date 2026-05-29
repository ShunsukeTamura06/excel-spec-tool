"""core.diagrams のテスト.

シート依存グラフと VBA コールグラフのビルダーを検証する.
"""

from __future__ import annotations

from core.diagrams import (
    build_diagrams,
    build_sheet_dependency_graph,
    build_vba_call_graph,
)
from core.models import (
    CellFormula,
    SheetInfo,
    VbaModule,
    VbaProcedure,
    Workbook,
)


def _make_wb_3sheets() -> Workbook:
    """Input -> Calc -> Output の三段依存. 自己ループ用に Calc 内参照も含む."""
    return Workbook(
        filename="demo.xlsx",
        sheets=[
            SheetInfo(name="Input", rows=10, cols=5),
            SheetInfo(
                name="Calc",
                rows=10,
                cols=5,
                formulas=[
                    CellFormula(
                        coord="A2",
                        formula="=Input!A2*Input!B2",
                        refs=["Input!A2", "Input!B2"],
                    ),
                    CellFormula(
                        coord="B2",
                        formula="=Input!C2",
                        refs=["Input!C2"],
                    ),
                    CellFormula(
                        coord="C2",
                        formula="=A2+B2",
                        refs=["A2", "B2"],  # 自己ループは除外されるはず
                    ),
                ],
            ),
            SheetInfo(
                name="Output",
                rows=2,
                cols=2,
                formulas=[
                    CellFormula(
                        coord="A1",
                        formula="=SUM(Calc!A2:A10)",
                        refs=["Calc!A2:A10"],
                    ),
                    CellFormula(
                        coord="A2",
                        formula="=SUM(Calc!B2:B10)",
                        refs=["Calc!B2:B10"],
                    ),
                ],
            ),
        ],
    )


class TestSheetDependencyGraph:
    def test_all_sheets_become_nodes(self) -> None:
        wb = _make_wb_3sheets()
        d = build_sheet_dependency_graph(wb)
        node_ids = {n.id for n in d.nodes}
        assert node_ids == {"Input", "Calc", "Output"}
        # 各ノードは kind="sheet" でラベルがシート名
        for n in d.nodes:
            assert n.kind == "sheet"
            assert n.label == n.id

    def test_cross_sheet_edges_with_weights(self) -> None:
        wb = _make_wb_3sheets()
        d = build_sheet_dependency_graph(wb)
        edge_map = {(e.src, e.dst): e for e in d.edges}

        # Calc -> Input: 3 参照 (A2, B2, C2)
        assert ("Calc", "Input") in edge_map
        assert edge_map[("Calc", "Input")].weight == 3
        assert edge_map[("Calc", "Input")].kind == "formula"

        # Output -> Calc: 2 参照
        assert ("Output", "Calc") in edge_map
        assert edge_map[("Output", "Calc")].weight == 2

    def test_self_loops_are_excluded(self) -> None:
        wb = _make_wb_3sheets()
        d = build_sheet_dependency_graph(wb)
        for e in d.edges:
            assert e.src != e.dst, f"self-loop should be excluded: {e}"

    def test_external_sheet_reference_is_skipped(self) -> None:
        """ブック内に存在しないシートへの参照はエッジにしない."""
        wb = Workbook(
            filename="x.xlsx",
            sheets=[
                SheetInfo(
                    name="Main",
                    rows=1,
                    cols=1,
                    formulas=[
                        CellFormula(
                            coord="A1",
                            formula="=Ghost!A1",
                            refs=["Ghost!A1"],
                        ),
                    ],
                ),
            ],
        )
        d = build_sheet_dependency_graph(wb)
        assert d.edges == []
        assert {n.id for n in d.nodes} == {"Main"}

    def test_no_formulas_no_edges(self) -> None:
        wb = Workbook(
            filename="empty.xlsx",
            sheets=[
                SheetInfo(name="A", rows=0, cols=0),
                SheetInfo(name="B", rows=0, cols=0),
            ],
        )
        d = build_sheet_dependency_graph(wb)
        assert d.edges == []
        assert len(d.nodes) == 2

    def test_meta_contains_formula_count(self) -> None:
        wb = _make_wb_3sheets()
        d = build_sheet_dependency_graph(wb)
        meta = {n.id: n.meta for n in d.nodes}
        assert meta["Calc"]["formulas"] == 3
        assert meta["Output"]["formulas"] == 2
        assert meta["Input"]["formulas"] == 0


# ---------- VBA コールグラフ ----------


def _proc(name: str, code: str, start: int = 1, end: int | None = None) -> VbaProcedure:
    if end is None:
        end = start + max(0, len(code.splitlines()) - 1)
    return VbaProcedure(
        name=name,
        kind="Sub",
        start_line=start,
        end_line=end,
        code=code,
    )


class TestVbaCallGraph:
    def test_simple_call_statement(self) -> None:
        """`Call Foo` 文を検出する."""
        modules = [
            VbaModule(
                name="M1",
                type="Module",
                code="Sub Entry()\n    Call Foo\nEnd Sub\nSub Foo()\nEnd Sub",
                procedures=[
                    _proc("Entry", "Sub Entry()\n    Call Foo\nEnd Sub", 1, 3),
                    _proc("Foo", "Sub Foo()\nEnd Sub", 4, 5),
                ],
            ),
        ]
        d = build_vba_call_graph(modules)
        edges = {(e.src, e.dst) for e in d.edges}
        assert ("M1.Entry", "M1.Foo") in edges

    def test_function_call_expression(self) -> None:
        """`x = Helper(1)` のような式形式の呼び出しを検出する."""
        helper_code = "Function Helper(n As Long) As Long\n    Helper = n * 2\nEnd Function"
        modules = [
            VbaModule(
                name="M1",
                type="Module",
                code="",
                procedures=[
                    _proc(
                        "Entry",
                        "Sub Entry()\n    Dim x\n    x = Helper(1)\nEnd Sub",
                        1,
                        4,
                    ),
                    VbaProcedure(
                        name="Helper",
                        kind="Function",
                        start_line=5,
                        end_line=7,
                        code=helper_code,
                    ),
                ],
            ),
        ]
        d = build_vba_call_graph(modules)
        edges = {(e.src, e.dst) for e in d.edges}
        assert ("M1.Entry", "M1.Helper") in edges

    def test_multiple_calls_aggregate_weight(self) -> None:
        modules = [
            VbaModule(
                name="M1",
                type="Module",
                code="",  # 個別 proc.code で十分なのでモジュール code は空
                procedures=[
                    _proc(
                        "A",
                        "Sub A()\n    Call B\n    Call B\n    Call B\nEnd Sub",
                        1,
                        5,
                    ),
                    _proc("B", "Sub B()\nEnd Sub", 6, 7),
                ],
            ),
        ]
        d = build_vba_call_graph(modules)
        edge = next(e for e in d.edges if e.src == "M1.A" and e.dst == "M1.B")
        assert edge.weight == 3

    def test_cross_module_call(self) -> None:
        modules = [
            VbaModule(
                name="Main",
                type="Module",
                code="",
                procedures=[
                    _proc(
                        "Run",
                        "Sub Run()\n    Utils.Calc 1, 2\n    Call Calc\nEnd Sub",
                        1,
                        4,
                    ),
                ],
            ),
            VbaModule(
                name="Utils",
                type="Module",
                code="",
                procedures=[_proc("Calc", "Sub Calc(a, b)\nEnd Sub", 1, 2)],
            ),
        ]
        d = build_vba_call_graph(modules)
        edges = {(e.src, e.dst) for e in d.edges}
        # 別モジュールの Calc を見つけられる
        assert ("Main.Run", "Utils.Calc") in edges

    def test_strings_and_comments_are_ignored(self) -> None:
        """コメントや文字列リテラル中の同名は呼び出しと見なさない."""
        modules = [
            VbaModule(
                name="M1",
                type="Module",
                code="",
                procedures=[
                    _proc(
                        "Entry",
                        'Sub Entry()\n    \' Call Foo\n    MsgBox "Foo"\nEnd Sub',
                        1,
                        4,
                    ),
                    _proc("Foo", "Sub Foo()\nEnd Sub", 5, 6),
                ],
            ),
        ]
        d = build_vba_call_graph(modules)
        assert d.edges == []

    def test_recursion_excluded(self) -> None:
        """自己再帰呼び出しはエッジにしない (描画上ノイズになるため)."""
        modules = [
            VbaModule(
                name="M1",
                type="Module",
                code="",
                procedures=[
                    _proc(
                        "Fact",
                        (
                            "Function Fact(n As Long) As Long\n"
                            "    If n <= 1 Then Fact = 1 "
                            "Else Fact = n * Fact(n - 1)\n"
                            "End Function"
                        ),
                        1,
                        3,
                    ),
                ],
            ),
        ]
        d = build_vba_call_graph(modules)
        # ノードは登録される
        assert {n.id for n in d.nodes} == {"M1.Fact"}
        # 自己再帰は edges に出ない
        assert d.edges == []

    def test_keywords_are_not_treated_as_calls(self) -> None:
        """`If` `End` `Dim` 等の VBA キーワードと同名プロシージャを混同しない."""
        modules = [
            VbaModule(
                name="M1",
                type="Module",
                code="",
                procedures=[
                    _proc(
                        "Entry",
                        "Sub Entry()\n    If True Then Exit Sub\nEnd Sub",
                        1,
                        3,
                    ),
                    # キーワードと同名のプロシージャ — 通常はないが、誤検出回避の確認
                    _proc("End", "Sub End()\nEnd Sub", 4, 5),
                ],
            ),
        ]
        d = build_vba_call_graph(modules)
        edges = {(e.src, e.dst) for e in d.edges}
        # キーワードはマッチしないので、 'End' プロシージャへのエッジは無い
        assert ("M1.Entry", "M1.End") not in edges

    def test_empty_modules(self) -> None:
        d = build_vba_call_graph([])
        assert d.nodes == []
        assert d.edges == []

    def test_build_diagrams_combines_both(self) -> None:
        wb = _make_wb_3sheets()
        wb.vba_modules.append(
            VbaModule(
                name="M1",
                type="Module",
                code="",
                procedures=[
                    _proc("A", "Sub A()\n    Call B\nEnd Sub", 1, 3),
                    _proc("B", "Sub B()\nEnd Sub", 4, 5),
                ],
            )
        )
        out = build_diagrams(wb)
        assert out.sheet_deps.kind == "sheet_deps"
        assert out.vba_calls.kind == "vba_calls"
        assert any(e.src == "Calc" and e.dst == "Input" for e in out.sheet_deps.edges)
        assert any(e.src == "M1.A" and e.dst == "M1.B" for e in out.vba_calls.edges)
