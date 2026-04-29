"""backend.storage のテスト."""

import json
import re
import uuid
from pathlib import Path

import pytest

from backend.storage import JobNotFoundError, Storage, _safe_suffix, _validate_job_id
from core.models import (
    CellFormula,
    ChatMessage,
    Reference,
    ReferenceIndex,
    SheetInfo,
    Workbook,
)

# ---------- 内部ヘルパー ----------


class TestValidateJobId:
    def test_valid_uuid_v4(self) -> None:
        _validate_job_id(str(uuid.uuid4()))

    def test_invalid_string_raises(self) -> None:
        with pytest.raises(ValueError):
            _validate_job_id("not-a-uuid")

    def test_uuid_v1_rejected(self) -> None:
        # v1 はバージョン桁が `1` なので v4 正規表現に当たらない
        with pytest.raises(ValueError):
            _validate_job_id(str(uuid.uuid1()))

    def test_path_traversal_rejected(self) -> None:
        with pytest.raises(ValueError):
            _validate_job_id("../etc/passwd")
        with pytest.raises(ValueError):
            _validate_job_id("a/b")


class TestSafeSuffix:
    def test_xlsm(self) -> None:
        assert _safe_suffix("foo.xlsm") == ".xlsm"

    def test_xls_uppercase(self) -> None:
        assert _safe_suffix("foo.XLS") == ".xls"

    def test_no_suffix(self) -> None:
        assert _safe_suffix("foo") == ".bin"

    def test_path_traversal(self) -> None:
        # 拡張子部分にスラッシュなどが入った異常系は .bin
        assert _safe_suffix("foo.x/m") == ".bin"

    def test_too_long_suffix(self) -> None:
        assert _safe_suffix("foo." + "a" * 50) == ".bin"


# ---------- Storage 基本 ----------


@pytest.fixture
def storage(tmp_path: Path) -> Storage:
    return Storage(tmp_path / "jobs")


class TestCreateJob:
    def test_returns_meta_and_writes_files(self, storage: Storage) -> None:
        meta = storage.create_job("input.xlsm", b"hello")
        assert re.match(
            r"^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$",
            meta.job_id,
        )
        assert meta.filename == "input.xlsm"
        assert meta.status == "uploaded"

        d = storage.jobs_dir / meta.job_id
        assert d.is_dir()
        assert (d / "original.xlsm").read_bytes() == b"hello"
        assert (d / "meta.json").is_file()

    def test_directory_permissions_700(self, storage: Storage) -> None:
        meta = storage.create_job("a.xlsm", b"x")
        d = storage.jobs_dir / meta.job_id
        mode = d.stat().st_mode & 0o777
        assert mode == 0o700

    def test_preserves_original_extension(self, storage: Storage) -> None:
        meta = storage.create_job("legacy.xls", b"x")
        d = storage.jobs_dir / meta.job_id
        assert (d / "original.xls").is_file()

    def test_falls_back_to_bin_for_unsafe_suffix(self, storage: Storage) -> None:
        meta = storage.create_job("weird.../passwd", b"x")
        d = storage.jobs_dir / meta.job_id
        assert (d / "original.bin").is_file()


class TestListJobs:
    def test_lists_all_in_descending_time(self, storage: Storage) -> None:
        m1 = storage.create_job("a.xlsm", b"1")
        m2 = storage.create_job("b.xlsm", b"2")
        listed = storage.list_jobs()
        assert {m.job_id for m in listed} == {m1.job_id, m2.job_id}
        # created_at 降順 (新しいものが先頭)
        assert listed[0].created_at >= listed[-1].created_at

    def test_ignores_non_uuid_subdirs(self, storage: Storage) -> None:
        # 関係ないディレクトリは無視される (パスインジェクション対策)
        (storage.jobs_dir / "junk").mkdir()
        meta = storage.create_job("a.xlsm", b"1")
        listed = storage.list_jobs()
        assert [m.job_id for m in listed] == [meta.job_id]

    def test_empty(self, storage: Storage) -> None:
        assert storage.list_jobs() == []


class TestDeleteJob:
    def test_removes_directory(self, storage: Storage) -> None:
        meta = storage.create_job("a.xlsm", b"1")
        d = storage.jobs_dir / meta.job_id
        assert d.is_dir()
        assert storage.delete_job(meta.job_id) is True
        assert not d.exists()

    def test_returns_false_when_missing(self, storage: Storage) -> None:
        assert storage.delete_job(str(uuid.uuid4())) is False

    def test_invalid_job_id_raises(self, storage: Storage) -> None:
        with pytest.raises(ValueError):
            storage.delete_job("../etc")


# ---------- meta / status ----------


class TestMeta:
    def test_get_meta(self, storage: Storage) -> None:
        meta = storage.create_job("a.xlsm", b"1")
        loaded = storage.get_meta(meta.job_id)
        assert loaded == meta

    def test_get_meta_not_found(self, storage: Storage) -> None:
        with pytest.raises(JobNotFoundError):
            storage.get_meta(str(uuid.uuid4()))

    def test_update_status(self, storage: Storage) -> None:
        meta = storage.create_job("a.xlsm", b"1")
        updated = storage.update_status(meta.job_id, "extracted")
        assert updated.status == "extracted"
        # 永続化されていること
        assert storage.get_meta(meta.job_id).status == "extracted"


# ---------- workbook / spec / references roundtrip ----------


class TestWorkbookRoundtrip:
    def test_save_and_load(self, storage: Storage) -> None:
        meta = storage.create_job("a.xlsm", b"1")
        wb = Workbook(
            filename="a.xlsm",
            sheets=[
                SheetInfo(
                    name="S",
                    rows=1,
                    cols=1,
                    formulas=[CellFormula(coord="S!A1", formula="=1+1", refs=[])],
                )
            ],
        )
        storage.save_workbook(meta.job_id, wb)
        loaded = storage.load_workbook(meta.job_id)
        assert loaded == wb


class TestSpecRoundtrip:
    def test_save_and_load(self, storage: Storage) -> None:
        meta = storage.create_job("a.xlsm", b"1")
        text = "# 設計書\n\nhello"
        storage.save_spec(meta.job_id, text)
        assert storage.load_spec(meta.job_id) == text


class TestReferencesRoundtrip:
    def test_save_and_load(self, storage: Storage) -> None:
        meta = storage.create_job("a.xlsm", b"1")
        idx = ReferenceIndex(
            refs={
                "Calc!A1": [
                    Reference(kind="formula", from_="Out!K3", to="Calc!A1", code="=Calc!A1"),
                ]
            }
        )
        storage.save_references(meta.job_id, idx)
        loaded = storage.load_references(meta.job_id)
        assert loaded == idx

    def test_serialized_uses_from_alias(self, storage: Storage) -> None:
        meta = storage.create_job("a.xlsm", b"1")
        idx = ReferenceIndex(refs={"X": [Reference(kind="formula", from_="A!1", to="X")]})
        storage.save_references(meta.job_id, idx)
        path = storage.jobs_dir / meta.job_id / "references.json"
        raw = json.loads(path.read_text(encoding="utf-8"))
        # JSON 上は "from" キー
        assert "from" in raw["refs"]["X"][0]
        assert "from_" not in raw["refs"]["X"][0]


# ---------- chat ----------


class TestChat:
    def test_append_and_load(self, storage: Storage) -> None:
        meta = storage.create_job("a.xlsm", b"1")
        m1 = ChatMessage(role="user", content="hi", timestamp="2026-04-30T00:00:00Z")
        m2 = ChatMessage(role="assistant", content="hello", timestamp="2026-04-30T00:00:01Z")
        storage.append_chat_message(meta.job_id, m1)
        storage.append_chat_message(meta.job_id, m2)
        loaded = storage.load_chat_history(meta.job_id)
        assert loaded == [m1, m2]

    def test_empty_history(self, storage: Storage) -> None:
        meta = storage.create_job("a.xlsm", b"1")
        assert storage.load_chat_history(meta.job_id) == []

    def test_skips_malformed_lines(self, storage: Storage, tmp_path: Path) -> None:
        meta = storage.create_job("a.xlsm", b"1")
        path = storage.jobs_dir / meta.job_id / "chat_history.jsonl"
        path.write_text(
            '{"role":"user","content":"ok","timestamp":"t"}\nGARBAGE\n', encoding="utf-8"
        )
        loaded = storage.load_chat_history(meta.job_id)
        assert len(loaded) == 1
        assert loaded[0].content == "ok"

    def test_jsonl_uses_append_mode(self, storage: Storage) -> None:
        # 各 append が独立に1行追加し、既存内容を破壊しないこと
        meta = storage.create_job("a.xlsm", b"1")
        for i in range(5):
            storage.append_chat_message(
                meta.job_id,
                ChatMessage(role="user", content=str(i), timestamp=f"t{i}"),
            )
        path = storage.jobs_dir / meta.job_id / "chat_history.jsonl"
        lines = path.read_text(encoding="utf-8").strip().split("\n")
        assert len(lines) == 5


# ---------- env ----------


class TestFromEnv:
    def test_uses_jobs_dir_env(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        target = tmp_path / "custom_jobs"
        monkeypatch.setenv("JOBS_DIR", str(target))
        s = Storage.from_env()
        assert s.jobs_dir == target
        assert target.is_dir()


# ---------- security: path traversal across all entrypoints ----------


class TestPathTraversal:
    @pytest.mark.parametrize(
        "method",
        [
            "get_meta",
            "get_original_path",
            "load_workbook",
            "load_spec",
            "load_references",
            "load_chat_history",
            "update_status",
            "save_spec",
            "save_workbook",
            "save_references",
            "append_chat_message",
        ],
    )
    def test_invalid_job_id_rejected(self, storage: Storage, method: str) -> None:
        # update_status / save_spec などは追加引数が必要
        bad_id = "../etc/passwd"
        if method == "update_status":
            with pytest.raises(ValueError):
                storage.update_status(bad_id, "uploaded")
        elif method == "save_spec":
            with pytest.raises(ValueError):
                storage.save_spec(bad_id, "x")
        elif method == "save_workbook":
            with pytest.raises(ValueError):
                storage.save_workbook(bad_id, Workbook(filename="x"))
        elif method == "save_references":
            with pytest.raises(ValueError):
                storage.save_references(bad_id, ReferenceIndex())
        elif method == "append_chat_message":
            with pytest.raises(ValueError):
                storage.append_chat_message(
                    bad_id, ChatMessage(role="user", content="x", timestamp="t")
                )
        else:
            with pytest.raises(ValueError):
                getattr(storage, method)(bad_id)
