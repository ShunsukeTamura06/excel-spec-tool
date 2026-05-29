"""POST /feedback と Storage.append_feedback / list_feedback のテスト."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from fastapi.testclient import TestClient

from backend.storage import Storage
from core.models import Feedback

# ---------- Storage 単体 ----------


class TestStorageFeedback:
    def test_append_and_list(self, tmp_path: Path) -> None:
        storage = Storage(tmp_path / "jobs")
        item = Feedback(
            id="a1",
            timestamp="2026-05-19T10:00:00+00:00",
            kind="improvement",
            comment="シート選択をキーボードで切替えたい",
            page="/spec/abc",
        )
        storage.append_feedback(item)
        out = storage.list_feedback()
        assert len(out) == 1
        assert out[0].id == "a1"
        assert out[0].comment.startswith("シート選択")

    def test_jsonl_file_layout(self, tmp_path: Path) -> None:
        storage = Storage(tmp_path / "jobs")
        item = Feedback(
            id="b1",
            timestamp="2026-05-19T12:34:56+00:00",
            kind="thumbs_up",
        )
        storage.append_feedback(item)
        expected = tmp_path / "jobs" / "_feedback" / "2026-05-19.jsonl"
        assert expected.is_file()
        # 1 行 1 件で書かれている
        lines = expected.read_text(encoding="utf-8").splitlines()
        assert len(lines) == 1
        parsed = json.loads(lines[0])
        assert parsed["id"] == "b1"
        assert parsed["kind"] == "thumbs_up"

    def test_list_sorted_desc_across_days(self, tmp_path: Path) -> None:
        storage = Storage(tmp_path / "jobs")
        # 別日に 2 件
        storage.append_feedback(
            Feedback(id="old", timestamp="2026-05-01T08:00:00+00:00", kind="bug")
        )
        storage.append_feedback(
            Feedback(id="new", timestamp="2026-05-19T08:00:00+00:00", kind="bug")
        )
        out = storage.list_feedback()
        assert [it.id for it in out] == ["new", "old"]

    def test_skips_malformed_lines(self, tmp_path: Path) -> None:
        storage = Storage(tmp_path / "jobs")
        fb_dir = tmp_path / "jobs" / "_feedback"
        fb_dir.mkdir(parents=True)
        # 正しい 1 行 + 壊れた 1 行 + 空行
        line_ok = json.dumps(
            {
                "id": "x",
                "timestamp": "2026-05-19T10:00:00+00:00",
                "kind": "other",
                "comment": "",
            }
        )
        (fb_dir / "2026-05-19.jsonl").write_text(
            line_ok + "\n" + "{not json}\n" + "\n",
            encoding="utf-8",
        )
        out = storage.list_feedback()
        assert len(out) == 1
        assert out[0].id == "x"

    def test_list_empty_when_no_feedback(self, tmp_path: Path) -> None:
        storage = Storage(tmp_path / "jobs")
        assert storage.list_feedback() == []


# ---------- POST /feedback ----------


class TestFeedbackRoute:
    def test_submit_thumbs_up_minimal(self, client: TestClient) -> None:
        r = client.post("/feedback", json={"kind": "thumbs_up"})
        assert r.status_code == 200
        body = r.json()
        assert body["ok"] is True
        assert "id" in body

    def test_submit_with_full_context(self, client: TestClient, backend_storage: Storage) -> None:
        payload = {
            "kind": "improvement",
            "comment": "VBA のシンタックスハイライトが欲しい",
            "page": "/spec/abc",
            "job_id": "00000000-0000-4000-8000-000000000001",
            "target_id": "msg-123",
            "target_excerpt": "改修手順: ...",
            "user_label": "田中",
        }
        r = client.post("/feedback", json=payload)
        assert r.status_code == 200

        stored = backend_storage.list_feedback()
        assert len(stored) == 1
        item = stored[0]
        assert item.kind == "improvement"
        assert item.comment == payload["comment"]
        assert item.page == payload["page"]
        assert item.job_id == payload["job_id"]
        assert item.target_id == payload["target_id"]
        assert item.user_label == "田中"

    def test_invalid_kind_rejected(self, client: TestClient) -> None:
        r = client.post("/feedback", json={"kind": "bogus"})
        assert r.status_code == 422  # Pydantic Literal violation

    def test_comment_truncated_at_4000_chars(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        long_comment = "あ" * 5000
        r = client.post("/feedback", json={"kind": "other", "comment": long_comment})
        assert r.status_code == 200
        stored = backend_storage.list_feedback()
        assert len(stored[0].comment) == 4000

    def test_id_and_timestamp_assigned_by_server(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        # client は id / timestamp を送らないが、保存時には埋まる
        r = client.post("/feedback", json={"kind": "bug", "comment": "壊れた"})
        assert r.status_code == 200
        body = r.json()
        item = backend_storage.list_feedback()[0]
        # サーバ採番された UUID と一致
        assert item.id == body["id"]
        # ISO8601 タイムゾーン付きであること (粗チェック)
        assert "T" in item.timestamp
        # 直近 (過去 60 秒以内)
        now = datetime.now(timezone.utc)
        ts = datetime.fromisoformat(item.timestamp)
        assert abs((now - ts).total_seconds()) < 60

    def test_thumbs_down_with_target(self, client: TestClient, backend_storage: Storage) -> None:
        r = client.post(
            "/feedback",
            json={
                "kind": "thumbs_down",
                "target_id": "2026-05-19T10:11:12+00:00",
                "target_excerpt": "改修手順: A1 セルを編集する...",
                "job_id": "00000000-0000-4000-8000-000000000002",
            },
        )
        assert r.status_code == 200
        item = backend_storage.list_feedback()[0]
        assert item.kind == "thumbs_down"
        assert item.target_id == "2026-05-19T10:11:12+00:00"
        assert "改修手順" in item.target_excerpt
