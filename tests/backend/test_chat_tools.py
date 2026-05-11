"""backend.routes.chat の function calling ループ統合テスト.

MockLLMClient.queue_response() でツール呼び出しをスクリプト化し、
チャットルートが想定通りに tool ループを回すことを検証する。
"""

from collections.abc import Iterator
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook as OpyWorkbook

from backend.dependencies import get_llm_client, get_storage
from backend.llm_client import LLMResponse, LLMToolCall, MockLLMClient
from backend.main import create_app
from backend.storage import Storage


@pytest.fixture
def storage(tmp_path: Path) -> Storage:
    return Storage(tmp_path / "jobs")


@pytest.fixture
def scripted_llm() -> MockLLMClient:
    """テストごとに queue_response でスクリプトを組む."""
    return MockLLMClient()


@pytest.fixture
def app_client(storage: Storage, scripted_llm: MockLLMClient) -> Iterator[TestClient]:
    app = create_app()
    app.dependency_overrides[get_storage] = lambda: storage
    app.dependency_overrides[get_llm_client] = lambda: scripted_llm
    with TestClient(app) as c:
        yield c
    app.dependency_overrides.clear()


@pytest.fixture
def xlsx_bytes(tmp_path: Path) -> bytes:
    wb = OpyWorkbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Portfolio"
    headers = ["銘柄コード", "銘柄名", "保有口数", "現在値", "評価損益", "実現損益"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=1, column=i, value=h)
    ws.cell(row=2, column=1, value="ABC")
    ws.cell(row=2, column=2, value="株式会社A")
    out = tmp_path / "p.xlsx"
    wb.save(out)
    return out.read_bytes()


def _setup_job(client: TestClient, body: bytes) -> str:
    r = client.post(
        "/extract",
        files={"file": ("p.xlsx", body, "application/octet-stream")},
    )
    assert r.status_code == 200, r.text
    return r.json()["job_id"]


class TestToolLoop:
    def test_no_tool_calls_returns_simple_reply(
        self, app_client: TestClient, xlsx_bytes: bytes, scripted_llm: MockLLMClient
    ) -> None:
        # 1ターンだけで終わるシナリオ
        scripted_llm.queue_response(LLMResponse(content="ok"))

        job_id = _setup_job(app_client, xlsx_bytes)
        r = app_client.post(f"/chat/{job_id}", json={"message": "hi"})
        assert r.status_code == 200, r.text
        body = r.json()
        assert body["reply"] == "ok"
        assert body["tool_trace"] == []

    def test_single_tool_call_then_final_reply(
        self, app_client: TestClient, xlsx_bytes: bytes, scripted_llm: MockLLMClient
    ) -> None:
        # 1) LLM が find_cells を呼ぶ
        # 2) tool 結果を受け取って最終応答
        scripted_llm.queue_response(
            LLMResponse(
                content=None,
                tool_calls=[
                    LLMToolCall(
                        id="t1",
                        name="find_cells",
                        arguments={"query": "実現損益"},
                    )
                ],
            )
        )
        scripted_llm.queue_response(LLMResponse(content="実現損益は F1 です"))

        job_id = _setup_job(app_client, xlsx_bytes)
        r = app_client.post(f"/chat/{job_id}", json={"message": "実現損益はどこ？"})
        assert r.status_code == 200, r.text
        body = r.json()
        assert body["reply"] == "実現損益は F1 です"
        assert len(body["tool_trace"]) == 1
        assert body["tool_trace"][0]["name"] == "find_cells"
        assert body["tool_trace"][0]["arguments"] == {"query": "実現損益"}

    def test_multiple_tool_calls(
        self, app_client: TestClient, xlsx_bytes: bytes, scripted_llm: MockLLMClient
    ) -> None:
        # find -> get_range -> 最終応答
        scripted_llm.queue_response(
            LLMResponse(
                content=None,
                tool_calls=[
                    LLMToolCall(id="t1", name="find_cells", arguments={"query": "実現損益"})
                ],
            )
        )
        scripted_llm.queue_response(
            LLMResponse(
                content=None,
                tool_calls=[
                    LLMToolCall(
                        id="t2",
                        name="get_cells_range",
                        arguments={"sheet": "Portfolio", "range": "A1:F1"},
                    )
                ],
            )
        )
        scripted_llm.queue_response(LLMResponse(content="G 列に追加してください"))

        job_id = _setup_job(app_client, xlsx_bytes)
        r = app_client.post(f"/chat/{job_id}", json={"message": "MTM 列追加"})
        assert r.status_code == 200, r.text
        body = r.json()
        assert body["reply"] == "G 列に追加してください"
        assert len(body["tool_trace"]) == 2
        assert [t["name"] for t in body["tool_trace"]] == [
            "find_cells",
            "get_cells_range",
        ]

    def test_tool_loop_caps_iterations(
        self, app_client: TestClient, xlsx_bytes: bytes, scripted_llm: MockLLMClient
    ) -> None:
        # 無限ループ風: 常に tool_call を返す応答を 10 回キュー
        for i in range(10):
            scripted_llm.queue_response(
                LLMResponse(
                    content=None,
                    tool_calls=[
                        LLMToolCall(
                            id=f"t{i}",
                            name="find_cells",
                            arguments={"query": "x"},
                        )
                    ],
                )
            )
        job_id = _setup_job(app_client, xlsx_bytes)
        r = app_client.post(f"/chat/{job_id}", json={"message": "loop"})
        assert r.status_code == 200
        # MAX_TOOL_ITERATIONS で打ち切られていること
        body = r.json()
        from backend.routes.chat import MAX_TOOL_ITERATIONS

        assert len(body["tool_trace"]) == MAX_TOOL_ITERATIONS

    def test_history_contains_only_user_and_assistant(
        self, app_client: TestClient, xlsx_bytes: bytes, scripted_llm: MockLLMClient
    ) -> None:
        # tool 呼び出しを挟んでも、保存される履歴は user/assistant ペアのみ
        scripted_llm.queue_response(
            LLMResponse(
                content=None,
                tool_calls=[LLMToolCall(id="t1", name="find_cells", arguments={"query": "x"})],
            )
        )
        scripted_llm.queue_response(LLMResponse(content="final"))

        job_id = _setup_job(app_client, xlsx_bytes)
        app_client.post(f"/chat/{job_id}", json={"message": "q"})

        r = app_client.get(f"/chat/{job_id}/history")
        history = r.json()["history"]
        assert len(history) == 2
        assert {h["role"] for h in history} == {"user", "assistant"}

    def test_unknown_tool_does_not_crash(
        self, app_client: TestClient, xlsx_bytes: bytes, scripted_llm: MockLLMClient
    ) -> None:
        # 存在しない tool 名を LLM が要求してもクラッシュしない
        scripted_llm.queue_response(
            LLMResponse(
                content=None,
                tool_calls=[LLMToolCall(id="t1", name="no_such_tool", arguments={})],
            )
        )
        scripted_llm.queue_response(LLMResponse(content="recovered"))

        job_id = _setup_job(app_client, xlsx_bytes)
        r = app_client.post(f"/chat/{job_id}", json={"message": "x"})
        assert r.status_code == 200
        assert r.json()["reply"] == "recovered"
