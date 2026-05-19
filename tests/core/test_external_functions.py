"""core.external_functions のテスト.

レジストリと数式検出ロジックを検証する.
"""

from __future__ import annotations

from core.external_functions import (
    ExternalFunction,
    detect_in_formula,
    get_function,
    list_functions,
)


class TestRegistry:
    def test_lookup_known_function(self) -> None:
        bdh = get_function("BDH")
        assert bdh is not None
        assert bdh.name == "BDH"
        assert bdh.vendor == "Bloomberg"
        assert "ヒストリカル" in bdh.short

    def test_lookup_case_insensitive(self) -> None:
        assert get_function("bdh") is not None
        assert get_function("Bdh") is not None
        assert get_function("BDH") is not None

    def test_lookup_unknown_returns_none(self) -> None:
        assert get_function("SUM") is None
        assert get_function("UNKNOWN_FN") is None
        assert get_function("") is None

    def test_list_functions_contains_bdh_bdp_bds(self) -> None:
        funcs = list_functions()
        names = {f.name for f in funcs}
        assert {"BDH", "BDP", "BDS"} <= names

    def test_list_by_vendor(self) -> None:
        bloomberg = list_functions(vendor="Bloomberg")
        assert len(bloomberg) >= 3
        assert all(f.vendor == "Bloomberg" for f in bloomberg)

        empty = list_functions(vendor="Refinitiv")
        assert empty == []

    def test_definitions_have_examples_and_params(self) -> None:
        for fn in list_functions():
            assert fn.signature
            assert fn.examples, f"{fn.name}: examples missing"
            assert fn.params, f"{fn.name}: params missing"
            assert fn.long, f"{fn.name}: long description missing"


class TestDetectInFormula:
    def test_detects_bdh(self) -> None:
        assert detect_in_formula('=BDH("AAPL US Equity", "PX_LAST", TODAY()-30)') == ["BDH"]

    def test_detects_bdp(self) -> None:
        assert detect_in_formula('=BDP("USDJPY Curncy", "PX_LAST")') == ["BDP"]

    def test_detects_bds(self) -> None:
        assert detect_in_formula('=BDS("SPX Index", "INDX_MEMBERS")') == ["BDS"]

    def test_multiple_distinct_in_one_formula(self) -> None:
        formula = '=BDP("AAPL US Equity", "PX_LAST") + SUM(BDH("MSFT US Equity", "PX_LAST", "-1Y"))'
        # 順序が保たれてユニーク化される
        assert detect_in_formula(formula) == ["BDP", "BDH"]

    def test_ignores_standard_excel_functions(self) -> None:
        assert detect_in_formula("=SUM(A1:A10) + AVERAGE(B1:B10)") == []
        assert detect_in_formula("=IF(A1>0, VLOOKUP(B1, C:D, 2, FALSE), 0)") == []

    def test_ignores_identifier_without_paren(self) -> None:
        """関数呼び出しでない単なる識別子は誤検出しない."""
        # "BDH" という文字列が引数内にあっても、後ろが ( でなければ無視
        assert detect_in_formula('="this BDH text is just a string"') == []

    def test_case_insensitive_match(self) -> None:
        """小文字 / 大小混在でも検出されて大文字で返る."""
        assert detect_in_formula('=bdh("X", "Y", "Z")') == ["BDH"]
        assert detect_in_formula('=Bdp("X", "Y")') == ["BDP"]

    def test_lower_case_excel_function_not_matched(self) -> None:
        """sum() は登録されていないので拾わない."""
        assert detect_in_formula("=sum(A1:A10)") == []

    def test_empty_or_blank_returns_empty(self) -> None:
        assert detect_in_formula("") == []
        assert detect_in_formula("   ") == []

    def test_substring_of_other_identifier_not_matched(self) -> None:
        """`ABDH(` のような他関数を BDH と誤検出しない (単語境界)."""
        assert detect_in_formula('=ABDH("X")') == []
        assert detect_in_formula('=X_BDH("Y")') == []

    def test_isinstance_model(self) -> None:
        fn = get_function("BDH")
        assert isinstance(fn, ExternalFunction)
