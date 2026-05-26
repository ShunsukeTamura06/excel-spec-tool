"""system prompt の内容に対するスナップショット的検証.

「LLM の振る舞いガイドラインがプロンプトから抜け落ちない」ことを保証する.
ガイドライン自体は人間の議論で決めるものなので、ここでは「キーワードが
含まれていること」だけを軽くテストする。
"""

from backend.routes.chat import _build_system_prompt


class TestSystemPromptContainsCoreGuidelines:
    def test_includes_basic_role(self) -> None:
        prompt = _build_system_prompt("")
        assert "Excelツール改修支援AI" in prompt

    def test_includes_fact_based_principle(self) -> None:
        # 事実ベース / 推測禁止
        prompt = _build_system_prompt("")
        assert "事実" in prompt
        assert "推測" in prompt or "推測しない" in prompt or "推測・憶測" in prompt

    def test_includes_dont_know_is_ok(self) -> None:
        # 分からないと言ってよい
        prompt = _build_system_prompt("")
        assert "分からない" in prompt or "判断できません" in prompt or "見つかりません" in prompt

    def test_includes_ask_user_freely(self) -> None:
        # 質問してよい
        prompt = _build_system_prompt("")
        assert "聞き返" in prompt or "質問" in prompt

    def test_includes_no_omission(self) -> None:
        # 回答漏れ防止
        prompt = _build_system_prompt("")
        assert "回答漏れ" in prompt or "すべてに答える" in prompt or "網羅" in prompt

    def test_includes_response_format_sections(self) -> None:
        prompt = _build_system_prompt("")
        assert "確認できた事実" in prompt
        assert "改修手順" in prompt
        assert "波及範囲" in prompt
        assert "未解析リスク" in prompt
        assert "手動確認チェックリスト" in prompt

    def test_includes_tool_names(self) -> None:
        prompt = _build_system_prompt("")
        assert "get_cells_range" in prompt
        assert "find_cells" in prompt
        assert "lookup_references" in prompt
        assert "list_workbook_objects" in prompt
        assert "list_analysis_risks" in prompt

    def test_includes_reference_analysis_limitations(self) -> None:
        prompt = _build_system_prompt("")
        assert "参照解析の前提" in prompt
        assert "静的に確定できる" in prompt
        assert "Range(addr)" in prompt
        assert "影響が完全に無いとは断定しない" in prompt
        assert "影響ありません" in prompt

    def test_includes_spec_when_provided(self) -> None:
        prompt = _build_system_prompt("# 設計書: x.xlsm\n本文")
        assert "# 設計書: x.xlsm" in prompt
        assert "本文" in prompt

    def test_no_spec_works(self) -> None:
        # spec が空でもプロンプトは生成される
        prompt = _build_system_prompt("")
        assert len(prompt) > 100
        # 設計書セクションマーカーは出ない
        assert "# 設計書\n" not in prompt
