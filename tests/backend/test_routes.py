"""backend FastAPI ルートの統合テスト."""

import uuid

from fastapi.testclient import TestClient

from backend.storage import Storage

# ---------- /health ----------


class TestHealth:
    def test_health(self, client: TestClient) -> None:
        r = client.get("/health")
        assert r.status_code == 200
        assert r.json() == {"status": "ok"}


# ---------- /extract ----------


class TestExtract:
    def test_extract_xlsx(self, client: TestClient, sample_xlsx_bytes: bytes) -> None:
        r = client.post(
            "/extract",
            files={
                "file": (
                    "sample.xlsx",
                    sample_xlsx_bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
        )
        assert r.status_code == 200, r.text
        body = r.json()
        assert "job_id" in body
        # UUIDv4 形式
        uuid.UUID(body["job_id"], version=4)

    def test_extract_xls_returns_empty_sheets(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        # 中身が無くても .xls なら早期 return パスでジョブが完成する
        r = client.post(
            "/extract",
            files={"file": ("legacy.xls", b"not real xls", "application/octet-stream")},
        )
        assert r.status_code == 200, r.text
        job_id = r.json()["job_id"]
        wb = backend_storage.load_workbook(job_id)
        assert wb.sheets == []
        assert backend_storage.get_meta(job_id).status == "extracted"

    def test_extract_corrupt_xlsx_returns_422(self, client: TestClient) -> None:
        r = client.post(
            "/extract",
            files={"file": ("bad.xlsx", b"not a real xlsx", "application/octet-stream")},
        )
        assert r.status_code == 422

    def test_extract_empty_file_returns_400(self, client: TestClient) -> None:
        r = client.post(
            "/extract",
            files={"file": ("a.xlsx", b"", "application/octet-stream")},
        )
        assert r.status_code == 400

    def test_extract_persists_workbook_and_references(
        self,
        client: TestClient,
        sample_xlsx_bytes: bytes,
        backend_storage: Storage,
    ) -> None:
        r = client.post(
            "/extract",
            files={"file": ("sample.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        )
        job_id = r.json()["job_id"]
        wb = backend_storage.load_workbook(job_id)
        assert wb.filename == "sample.xlsx"
        assert any(f.coord == "Calc!A3" for f in wb.sheets[0].formulas)
        idx = backend_storage.load_references(job_id)
        # =SUM(A1:A2) が逆引きインデックスに登場する
        assert "A1:A2" in idx.refs


# ---------- /analyze ----------


class TestAnalyze:
    def test_analyze_creates_spec(
        self, client: TestClient, sample_xlsx_bytes: bytes, backend_storage: Storage
    ) -> None:
        # まず extract
        job_id = client.post(
            "/extract",
            files={"file": ("a.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        ).json()["job_id"]

        r = client.post(f"/analyze/{job_id}")
        assert r.status_code == 200, r.text
        assert r.json() == {"status": "ok"}

        spec_md = backend_storage.load_spec(job_id)
        assert "# 設計書: a.xlsx" in spec_md
        assert backend_storage.get_meta(job_id).status == "analyzed"

    def test_analyze_invalid_job_id(self, client: TestClient) -> None:
        r = client.post("/analyze/not-a-uuid")
        assert r.status_code == 400

    def test_analyze_missing_job(self, client: TestClient) -> None:
        r = client.post(f"/analyze/{uuid.uuid4()}")
        assert r.status_code == 404

    def test_analyze_before_extract_returns_409(
        self, client: TestClient, backend_storage: Storage
    ) -> None:
        # ジョブ作成だけして extract 完了前 (extracted.json が無い状態)
        meta = backend_storage.create_job("a.xlsm", b"x")
        r = client.post(f"/analyze/{meta.job_id}")
        assert r.status_code == 409


# ---------- /spec ----------


class TestSpec:
    def test_get_spec(self, client: TestClient, sample_xlsx_bytes: bytes) -> None:
        job_id = client.post(
            "/extract",
            files={"file": ("a.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        ).json()["job_id"]
        client.post(f"/analyze/{job_id}")

        r = client.get(f"/spec/{job_id}")
        assert r.status_code == 200
        body = r.json()
        assert "spec_md" in body
        assert "# 設計書" in body["spec_md"]
        assert body["meta"]["job_id"] == job_id

    def test_get_spec_before_analyze_returns_409(
        self, client: TestClient, sample_xlsx_bytes: bytes
    ) -> None:
        job_id = client.post(
            "/extract",
            files={"file": ("a.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        ).json()["job_id"]
        r = client.get(f"/spec/{job_id}")
        assert r.status_code == 409

    def test_get_spec_invalid_id(self, client: TestClient) -> None:
        assert client.get("/spec/not-uuid").status_code == 400

    def test_get_spec_missing(self, client: TestClient) -> None:
        assert client.get(f"/spec/{uuid.uuid4()}").status_code == 404


# ---------- /references ----------


class TestReferences:
    def test_lookup_existing_target(self, client: TestClient, sample_xlsx_bytes: bytes) -> None:
        job_id = client.post(
            "/extract",
            files={"file": ("a.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        ).json()["job_id"]

        r = client.get(f"/references/{job_id}", params={"target": "A1:A2"})
        assert r.status_code == 200
        refs = r.json()["refs"]
        assert len(refs) >= 1
        # JSON 上は from キー (alias)
        assert "from" in refs[0]
        assert "from_" not in refs[0]

    def test_lookup_unknown_target_returns_empty(
        self, client: TestClient, sample_xlsx_bytes: bytes
    ) -> None:
        job_id = client.post(
            "/extract",
            files={"file": ("a.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        ).json()["job_id"]
        r = client.get(f"/references/{job_id}", params={"target": "NoSuch!Z9"})
        assert r.status_code == 200
        assert r.json() == {"refs": []}

    def test_lookup_invalid_id(self, client: TestClient) -> None:
        assert client.get("/references/not-uuid").status_code == 400

    def test_lookup_missing(self, client: TestClient) -> None:
        assert client.get(f"/references/{uuid.uuid4()}").status_code == 404


# ---------- /chat ----------


class TestChat:
    def _new_analyzed_job(self, client: TestClient, sample_xlsx_bytes: bytes) -> str:
        job_id = client.post(
            "/extract",
            files={"file": ("a.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        ).json()["job_id"]
        client.post(f"/analyze/{job_id}")
        return job_id

    def test_chat_returns_reply_and_appends_history(
        self, client: TestClient, sample_xlsx_bytes: bytes
    ) -> None:
        job_id = self._new_analyzed_job(client, sample_xlsx_bytes)
        r = client.post(f"/chat/{job_id}", json={"message": "How do I update A3?"})
        assert r.status_code == 200, r.text
        body = r.json()
        assert "reply" in body
        assert "[mock" in body["reply"]
        assert len(body["history"]) == 2  # user + assistant
        assert body["history"][0]["role"] == "user"
        assert body["history"][1]["role"] == "assistant"

    def test_chat_history_grows_across_calls(
        self, client: TestClient, sample_xlsx_bytes: bytes
    ) -> None:
        job_id = self._new_analyzed_job(client, sample_xlsx_bytes)
        client.post(f"/chat/{job_id}", json={"message": "first"})
        client.post(f"/chat/{job_id}", json={"message": "second"})
        r = client.get(f"/chat/{job_id}/history")
        assert r.status_code == 200
        history = r.json()["history"]
        assert len(history) == 4
        assert history[0]["content"] == "first"
        assert history[2]["content"] == "second"

    def test_chat_works_without_spec(self, client: TestClient, sample_xlsx_bytes: bytes) -> None:
        # 設計書未生成 (analyze 前) でもチャットは動作する (spec は空 system prompt)
        job_id = client.post(
            "/extract",
            files={"file": ("a.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        ).json()["job_id"]
        r = client.post(f"/chat/{job_id}", json={"message": "hello"})
        assert r.status_code == 200

    def test_chat_empty_message_returns_400(
        self, client: TestClient, sample_xlsx_bytes: bytes
    ) -> None:
        job_id = self._new_analyzed_job(client, sample_xlsx_bytes)
        r = client.post(f"/chat/{job_id}", json={"message": "  "})
        assert r.status_code == 400

    def test_chat_invalid_id(self, client: TestClient) -> None:
        r = client.post("/chat/not-uuid", json={"message": "hi"})
        assert r.status_code == 400

    def test_chat_missing_job(self, client: TestClient) -> None:
        r = client.post(f"/chat/{uuid.uuid4()}", json={"message": "hi"})
        assert r.status_code == 404

    def test_history_empty_for_fresh_job(
        self, client: TestClient, sample_xlsx_bytes: bytes
    ) -> None:
        job_id = client.post(
            "/extract",
            files={"file": ("a.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        ).json()["job_id"]
        r = client.get(f"/chat/{job_id}/history")
        assert r.status_code == 200
        assert r.json() == {"history": []}


# ---------- /jobs ----------


class TestJobs:
    def test_list_empty(self, client: TestClient) -> None:
        assert client.get("/jobs").json() == {"jobs": []}

    def test_list_returns_created_jobs(self, client: TestClient, sample_xlsx_bytes: bytes) -> None:
        client.post(
            "/extract",
            files={"file": ("a.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        )
        client.post(
            "/extract",
            files={"file": ("b.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        )
        body = client.get("/jobs").json()
        assert len(body["jobs"]) == 2
        names = {j["filename"] for j in body["jobs"]}
        assert names == {"a.xlsx", "b.xlsx"}

    def test_delete_existing(
        self, client: TestClient, sample_xlsx_bytes: bytes, backend_storage: Storage
    ) -> None:
        job_id = client.post(
            "/extract",
            files={"file": ("a.xlsx", sample_xlsx_bytes, "application/octet-stream")},
        ).json()["job_id"]
        r = client.delete(f"/jobs/{job_id}")
        assert r.status_code == 200
        assert r.json() == {"deleted": True}
        # 物理削除されている
        assert not (backend_storage.jobs_dir / job_id).exists()

    def test_delete_missing(self, client: TestClient) -> None:
        assert client.delete(f"/jobs/{uuid.uuid4()}").status_code == 404

    def test_delete_invalid_id(self, client: TestClient) -> None:
        assert client.delete("/jobs/not-uuid").status_code == 400
