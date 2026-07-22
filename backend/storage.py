"""ローカルファイル永続化モジュール.

docs/SPEC.ja.md §5.2 に基づき、ジョブごとに以下の構成でファイルを管理する:

    {jobs_dir}/{job_id}/
    ├── original.{ext}       # アップロード原本
    ├── extracted.json       # Workbook モデル (Core 抽出結果)
    ├── diagnosis.json       # 根拠付きの一般ユーザー向け Excel 診断
    ├── spec.md              # 生成済み設計書
    ├── references.json      # ReferenceIndex モデル
    ├── verification.json    # 変更計画・実差分・policy gate の監査レコード
    ├── chat_history.jsonl   # 1行1メッセージ追記
    └── meta.json            # JobMeta モデル

セキュリティ上の留意点:
- job_id は UUIDv4 形式に限定 (パスインジェクション防止)
- ジョブディレクトリのパーミッションは 0o700
- chat_history.jsonl は append モードで書く
"""

from __future__ import annotations

import json
import logging
import os
import re
import shutil
import uuid
from datetime import datetime, timezone
from hashlib import sha256
from pathlib import Path

from core.change_record import ChangeExecutionRecord
from core.diagnosis import WorkbookDiagnosis
from core.models import (
    ChatMessage,
    ChatSessionMeta,
    Feedback,
    JobMeta,
    ReferenceIndex,
    Workbook,
)
from core.mutation import SafeChangePlan

logger = logging.getLogger(__name__)


_DEFAULT_JOBS_DIR = "./jobs"
_DIR_MODE = 0o700

# UUID v4: 8-4-4-4-12 hex, with version 4 nibble at the right position
_UUID_V4_RE = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$",
    re.IGNORECASE,
)
_DEFAULT_CHAT_SESSION_ID = "default"
_DEFAULT_CHAT_SESSION_TITLE = "既存の相談"


class StorageError(Exception):
    """Storage 層の例外."""


class JobNotFoundError(StorageError):
    """指定 job_id が存在しない."""


def _validate_job_id(job_id: str) -> None:
    """job_id が UUIDv4 形式か検証する. 不正なら ValueError を投げる."""
    if not isinstance(job_id, str) or not _UUID_V4_RE.match(job_id):
        raise ValueError(f"invalid job_id (must be UUIDv4): {job_id!r}")


def _validate_plan_id(plan_id: str) -> None:
    """plan_id が UUIDv4 形式か検証する. 不正なら ValueError を投げる."""
    if not isinstance(plan_id, str) or not _UUID_V4_RE.match(plan_id):
        raise ValueError(f"invalid plan_id (must be UUIDv4): {plan_id!r}")


def _validate_chat_session_id(session_id: str) -> None:
    """チャット session_id を検証する. default または UUIDv4 のみ許可."""
    if session_id == _DEFAULT_CHAT_SESSION_ID:
        return
    if not isinstance(session_id, str) or not _UUID_V4_RE.match(session_id):
        raise ValueError(f"invalid chat session_id: {session_id!r}")


def _utc_now_iso() -> str:
    """現在時刻を ISO8601 文字列で返す (タイムゾーン UTC)."""
    return datetime.now(timezone.utc).isoformat()


def _safe_suffix(filename: str) -> str:
    """アップロードファイル名から拡張子だけ取り出して安全化する.

    パストラバーサルやパスセパレータを除外し、`.xlsm` `.xls` 等を返す。
    不審な値の場合は `.bin` にフォールバック。
    """
    suffix = Path(filename).suffix
    if not suffix:
        return ".bin"
    # `.` の後ろに英数字 (1〜10 文字) のみ許容
    if not re.match(r"^\.[A-Za-z0-9]{1,10}$", suffix):
        return ".bin"
    return suffix.lower()


def _file_sha256(data: bytes) -> str:
    """ファイル内容の SHA-256 ダイジェストを 16 進文字列で返す."""
    return sha256(data).hexdigest()


def _chat_title_from_message(message: str) -> str:
    """最初のユーザー発話からセッションタイトルを作る."""
    text = " ".join((message or "").strip().split())
    if not text:
        return "新しい相談"
    return text[:40]


def _message_preview(message: str) -> str:
    """セッション一覧用にメッセージ本文を短く整形する."""
    text = " ".join((message or "").strip().split())
    return text[:80]


class Storage:
    """ジョブ永続化のファサード."""

    def __init__(self, jobs_dir: Path | str) -> None:
        self.jobs_dir = Path(jobs_dir)
        self.jobs_dir.mkdir(parents=True, exist_ok=True, mode=_DIR_MODE)

    # ------------------------------------------------------------------ env

    @classmethod
    def from_env(cls) -> Storage:
        """環境変数 JOBS_DIR を見て Storage を作る. 未設定なら ./jobs."""
        return cls(os.environ.get("JOBS_DIR", _DEFAULT_JOBS_DIR))

    # ------------------------------------------------------------ paths

    def _job_dir(self, job_id: str) -> Path:
        _validate_job_id(job_id)
        return self.jobs_dir / job_id

    def _require_job_dir(self, job_id: str) -> Path:
        p = self._job_dir(job_id)
        if not p.is_dir():
            raise JobNotFoundError(f"job not found: {job_id}")
        return p

    def get_original_path(self, job_id: str) -> Path:
        d = self._require_job_dir(job_id)
        candidates = sorted(d.glob("original.*"))
        if not candidates:
            raise JobNotFoundError(f"original file not found in {d}")
        return candidates[0]

    def get_verification_path(self, job_id: str) -> Path:
        """変更後ジョブの検証監査レコードのパスを返す."""

        return self._require_job_dir(job_id) / "verification.json"

    def get_pending_plan_path(self, job_id: str, plan_id: str) -> Path:
        """未実行の変更計画 (SafeChangePlan) のパスを返す."""

        _validate_plan_id(plan_id)
        return self._require_job_dir(job_id) / "pending_plans" / f"{plan_id}.json"

    # --------------------------------------------------------- create / list

    def create_job(self, filename: str, data: bytes) -> JobMeta:
        """新規ジョブを作成し原本を保存. UUIDv4 を採番して JobMeta を返す."""
        job_id = str(uuid.uuid4())
        d = self.jobs_dir / job_id
        d.mkdir(mode=_DIR_MODE)

        suffix = _safe_suffix(filename)
        original = d / f"original{suffix}"
        original.write_bytes(data)

        meta = JobMeta(
            job_id=job_id,
            filename=filename,
            created_at=_utc_now_iso(),
            status="uploaded",
            file_sha256=_file_sha256(data),
            file_size=len(data),
        )
        self._write_json(d / "meta.json", meta.model_dump())
        return meta

    def find_duplicate_job(self, data: bytes) -> JobMeta | None:
        """同一内容の利用可能な既存ジョブを返す.

        Args:
            data: アップロードされたファイルのバイト列。

        Returns:
            SHA-256 が一致し、分析済みの既存ジョブ。見つからない場合は None。
        """
        digest = _file_sha256(data)
        size = len(data)
        for meta in self.list_jobs():
            if meta.status != "analyzed":
                continue
            meta = self._ensure_file_fingerprint(meta)
            if meta.file_sha256 == digest and meta.file_size == size:
                return meta
        return None

    def list_jobs(self) -> list[JobMeta]:
        """登録済みジョブのメタを ISO 時刻降順で返す."""
        metas: list[JobMeta] = []
        for entry in self.jobs_dir.iterdir():
            if not entry.is_dir():
                continue
            if not _UUID_V4_RE.match(entry.name):
                continue
            meta_path = entry / "meta.json"
            if not meta_path.is_file():
                continue
            try:
                metas.append(self.get_meta(entry.name))
            except StorageError:
                logger.warning("Skipping unreadable job: %s", entry.name)
        metas.sort(key=lambda m: m.created_at, reverse=True)
        return metas

    def delete_job(self, job_id: str) -> bool:
        """ジョブディレクトリを丸ごと削除. 存在しなかった場合は False."""
        try:
            d = self._require_job_dir(job_id)
        except JobNotFoundError:
            return False
        shutil.rmtree(d)
        return True

    # ------------------------------------------------------ verification

    def save_verification(self, job_id: str, record: ChangeExecutionRecord) -> None:
        """変更計画・期待差分・実差分・policy判定を保存する."""

        path = self.get_verification_path(job_id)
        self._write_json(path, record.model_dump(mode="json"))

    def load_verification(self, job_id: str) -> ChangeExecutionRecord:
        """保存済みの変更検証監査レコードを読み込む."""

        path = self.get_verification_path(job_id)
        if not path.is_file():
            raise FileNotFoundError(f"verification record not found: {path}")
        return ChangeExecutionRecord.model_validate_json(path.read_text(encoding="utf-8"))

    # ------------------------------------------------------- pending plans

    def save_pending_plan(self, job_id: str, plan: SafeChangePlan) -> None:
        """propose段階で画面に表示した計画を、実行時の照合用に保存する.

        plan_id をキーにファイルへ保存する。実行 (execute/verify) 時はこの
        保存内容だけを信頼し、クライアントが送るリクエストボディの計画は
        参照しない (改ざん・すり替え防止)。
        """

        path = self.get_pending_plan_path(job_id, plan.plan.plan_id)
        path.parent.mkdir(parents=True, exist_ok=True, mode=_DIR_MODE)
        self._write_json(path, plan.model_dump(mode="json"))

    def load_pending_plan(self, job_id: str, plan_id: str) -> SafeChangePlan:
        """保存済みの未実行計画を読み込む.

        Raises:
            FileNotFoundError: 該当 plan_id が存在しない、または既に消費済みの場合。
        """

        path = self.get_pending_plan_path(job_id, plan_id)
        if not path.is_file():
            raise FileNotFoundError(f"pending plan not found: {path}")
        return SafeChangePlan.model_validate_json(path.read_text(encoding="utf-8"))

    def consume_pending_plan(self, job_id: str, plan_id: str) -> None:
        """計画を使い捨てにする (再実行によるすり替え・リプレイを防ぐ)."""

        path = self.get_pending_plan_path(job_id, plan_id)
        path.unlink(missing_ok=True)

    # ------------------------------------------------------------ meta

    def get_meta(self, job_id: str) -> JobMeta:
        d = self._require_job_dir(job_id)
        return JobMeta.model_validate_json((d / "meta.json").read_text(encoding="utf-8"))

    def _ensure_file_fingerprint(self, meta: JobMeta) -> JobMeta:
        """旧形式の meta.json にファイル指紋を補完して返す."""
        if meta.file_sha256 and meta.file_size is not None:
            return meta
        try:
            return self.refresh_original_fingerprint(meta.job_id)
        except (OSError, JobNotFoundError):
            return meta

    def refresh_original_fingerprint(self, job_id: str) -> JobMeta:
        """現在の成果物からハッシュとサイズを再計算してmetaへ保存する.

        Args:
            job_id: 指紋を更新するジョブID。

        Returns:
            更新後のジョブメタ情報。

        Raises:
            JobNotFoundError: ジョブまたは原本ファイルが存在しない場合。
            OSError: ファイルを読み書きできない場合。
        """

        meta = self.get_meta(job_id)
        try:
            path = self.get_original_path(job_id)
            data = path.read_bytes()
        except (OSError, JobNotFoundError):
            logger.warning("Failed to fingerprint original file for job: %s", job_id)
            raise

        updated = meta.model_copy(
            update={
                "file_sha256": _file_sha256(data),
                "file_size": len(data),
            }
        )
        self._write_json(self._job_dir(job_id) / "meta.json", updated.model_dump())
        return updated

    def update_status(
        self,
        job_id: str,
        status: str,
    ) -> JobMeta:
        """status を更新して meta.json を書き直す."""
        meta = self.get_meta(job_id)
        new_meta = meta.model_copy(update={"status": status})
        self._write_json(self._job_dir(job_id) / "meta.json", new_meta.model_dump())
        return new_meta

    # ------------------------------------------ workbook / diagnosis / spec / refs

    def save_workbook(self, job_id: str, wb: Workbook) -> None:
        d = self._require_job_dir(job_id)
        self._write_json(d / "extracted.json", wb.model_dump())

    def load_workbook(self, job_id: str) -> Workbook:
        d = self._require_job_dir(job_id)
        return Workbook.model_validate_json((d / "extracted.json").read_text(encoding="utf-8"))

    def save_diagnosis(self, job_id: str, diagnosis: WorkbookDiagnosis) -> None:
        """根拠付き Excel 診断を保存する."""
        d = self._require_job_dir(job_id)
        self._write_json(d / "diagnosis.json", diagnosis.model_dump())

    def load_diagnosis(self, job_id: str) -> WorkbookDiagnosis:
        """保存済みの根拠付き Excel 診断を返す."""
        d = self._require_job_dir(job_id)
        return WorkbookDiagnosis.model_validate_json(
            (d / "diagnosis.json").read_text(encoding="utf-8")
        )

    def save_spec(self, job_id: str, spec_md: str) -> None:
        d = self._require_job_dir(job_id)
        (d / "spec.md").write_text(spec_md, encoding="utf-8")

    def load_spec(self, job_id: str) -> str:
        d = self._require_job_dir(job_id)
        return (d / "spec.md").read_text(encoding="utf-8")

    def save_references(self, job_id: str, idx: ReferenceIndex) -> None:
        d = self._require_job_dir(job_id)
        # by_alias=True で from_ -> "from" としてシリアライズ
        self._write_json(d / "references.json", idx.model_dump(by_alias=True))

    def load_references(self, job_id: str) -> ReferenceIndex:
        d = self._require_job_dir(job_id)
        return ReferenceIndex.model_validate_json(
            (d / "references.json").read_text(encoding="utf-8")
        )

    # ------------------------------------------------------------ chat

    def _chat_sessions_dir(self, job_id: str) -> Path:
        d = self._require_job_dir(job_id) / "chat_sessions"
        d.mkdir(mode=_DIR_MODE, exist_ok=True)
        return d

    def _chat_sessions_meta_path(self, job_id: str) -> Path:
        return self._chat_sessions_dir(job_id) / "sessions.json"

    def _chat_history_path(self, job_id: str, session_id: str) -> Path:
        _validate_chat_session_id(session_id)
        d = self._require_job_dir(job_id)
        if session_id == _DEFAULT_CHAT_SESSION_ID:
            return d / "chat_history.jsonl"
        return self._chat_sessions_dir(job_id) / f"{session_id}.jsonl"

    def _load_chat_sessions_meta(self, job_id: str) -> list[ChatSessionMeta]:
        path = self._chat_sessions_meta_path(job_id)
        if not path.is_file():
            return []
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            logger.warning("Malformed chat sessions meta: %s", path)
            return []
        if not isinstance(raw, list):
            logger.warning("Unexpected chat sessions meta format: %s", path)
            return []
        metas: list[ChatSessionMeta] = []
        for item in raw:
            try:
                metas.append(ChatSessionMeta.model_validate(item))
            except Exception:  # noqa: BLE001
                logger.warning("Skipping malformed chat session meta in %s", path)
        return metas

    def _save_chat_sessions_meta(self, job_id: str, metas: list[ChatSessionMeta]) -> None:
        path = self._chat_sessions_meta_path(job_id)
        data = [m.model_dump() for m in metas]
        self._write_json(path, data)

    def _default_chat_session_meta(self, job_id: str) -> ChatSessionMeta:
        history = self._read_chat_history_file(
            self._chat_history_path(job_id, _DEFAULT_CHAT_SESSION_ID)
        )
        now = _utc_now_iso()
        updated_at = history[-1].timestamp if history else now
        preview = _message_preview(history[-1].content) if history else ""
        title = _DEFAULT_CHAT_SESSION_TITLE if history else "新しい相談"
        if history:
            first_user = next((m.content for m in history if m.role == "user" and m.content), "")
            if first_user:
                title = _chat_title_from_message(first_user)
        return ChatSessionMeta(
            session_id=_DEFAULT_CHAT_SESSION_ID,
            title=title,
            created_at=history[0].timestamp if history else now,
            updated_at=updated_at,
            archived=False,
            last_message_preview=preview,
            message_count=len(history),
        )

    def list_chat_sessions(
        self,
        job_id: str,
        *,
        include_archived: bool = False,
    ) -> list[ChatSessionMeta]:
        """ジョブに紐づくチャットセッションを更新日時降順で返す."""
        self._require_job_dir(job_id)
        metas = self._load_chat_sessions_meta(job_id)
        if not any(m.session_id == _DEFAULT_CHAT_SESSION_ID for m in metas):
            metas.append(self._default_chat_session_meta(job_id))
            self._save_chat_sessions_meta(job_id, metas)
        filtered = metas if include_archived else [m for m in metas if not m.archived]
        filtered.sort(key=lambda m: m.updated_at, reverse=True)
        return filtered

    def create_chat_session(self, job_id: str, title: str = "新しい相談") -> ChatSessionMeta:
        """ジョブ配下に新しいチャットセッションを作成する."""
        self._require_job_dir(job_id)
        metas = self._load_chat_sessions_meta(job_id)
        session_id = str(uuid.uuid4())
        now = _utc_now_iso()
        meta = ChatSessionMeta(
            session_id=session_id,
            title=title.strip()[:40] or "新しい相談",
            created_at=now,
            updated_at=now,
        )
        metas.append(meta)
        self._save_chat_sessions_meta(job_id, metas)
        return meta

    def get_or_create_chat_session(
        self,
        job_id: str,
        session_id: str = _DEFAULT_CHAT_SESSION_ID,
    ) -> ChatSessionMeta:
        """指定セッションを返す。default は必要に応じて作成する."""
        _validate_chat_session_id(session_id)
        metas = self._load_chat_sessions_meta(job_id)
        for meta in metas:
            if meta.session_id == session_id:
                return meta
        if session_id == _DEFAULT_CHAT_SESSION_ID:
            meta = self._default_chat_session_meta(job_id)
            metas.append(meta)
            self._save_chat_sessions_meta(job_id, metas)
            return meta
        raise JobNotFoundError(f"chat session not found: {session_id}")

    def update_chat_session(
        self,
        job_id: str,
        session_id: str,
        *,
        title: str | None = None,
        archived: bool | None = None,
    ) -> ChatSessionMeta:
        """チャットセッションのタイトルまたはアーカイブ状態を更新する."""
        _validate_chat_session_id(session_id)
        metas = self._load_chat_sessions_meta(job_id)
        for i, meta in enumerate(metas):
            if meta.session_id != session_id:
                continue
            updates: dict[str, object] = {"updated_at": _utc_now_iso()}
            if title is not None:
                updates["title"] = title.strip()[:40] or "新しい相談"
            if archived is not None:
                updates["archived"] = archived
            updated = meta.model_copy(update=updates)
            metas[i] = updated
            self._save_chat_sessions_meta(job_id, metas)
            return updated
        raise JobNotFoundError(f"chat session not found: {session_id}")

    def append_chat_message(
        self,
        job_id: str,
        message: ChatMessage,
        session_id: str = _DEFAULT_CHAT_SESSION_ID,
    ) -> None:
        meta = self.get_or_create_chat_session(job_id, session_id)
        path = self._chat_history_path(job_id, session_id)
        line = json.dumps(message.model_dump(), ensure_ascii=False) + "\n"
        with path.open("a", encoding="utf-8") as f:
            f.write(line)

        updates: dict[str, object] = {
            "updated_at": message.timestamp,
            "last_message_preview": _message_preview(message.content),
            "message_count": meta.message_count + 1,
        }
        if (
            meta.message_count == 0
            and meta.title in {"新しい相談", _DEFAULT_CHAT_SESSION_TITLE}
            and message.role == "user"
        ):
            updates["title"] = _chat_title_from_message(message.content)
        updated = meta.model_copy(update=updates)
        metas = self._load_chat_sessions_meta(job_id)
        for i, item in enumerate(metas):
            if item.session_id == session_id:
                metas[i] = updated
                break
        else:
            metas.append(updated)
        self._save_chat_sessions_meta(job_id, metas)

    def _read_chat_history_file(self, path: Path) -> list[ChatMessage]:
        if not path.is_file():
            return []
        msgs: list[ChatMessage] = []
        with path.open("r", encoding="utf-8") as f:
            for raw in f:
                raw = raw.strip()
                if not raw:
                    continue
                try:
                    msgs.append(ChatMessage.model_validate_json(raw))
                except Exception:  # noqa: BLE001
                    logger.warning("Skipping malformed chat line in %s", path)
        return msgs

    def load_chat_history(
        self,
        job_id: str,
        session_id: str = _DEFAULT_CHAT_SESSION_ID,
    ) -> list[ChatMessage]:
        self.get_or_create_chat_session(job_id, session_id)
        return self._read_chat_history_file(self._chat_history_path(job_id, session_id))

    # ------------------------------------------------------------ feedback
    # ユーザーからのフィードバックは jobs と独立に永続化する (ジョブ削除でも残す).
    # 構造: <jobs_dir>/_feedback/<YYYY-MM-DD>.jsonl  (1 日 1 ファイル, 1 行 1 件)
    # `_` プレフィックスで UUID ベースの job ディレクトリ列挙にひっかからない.

    def _feedback_dir(self) -> Path:
        d = self.jobs_dir / "_feedback"
        d.mkdir(parents=True, exist_ok=True, mode=_DIR_MODE)
        return d

    def append_feedback(self, item: Feedback) -> None:
        """フィードバック 1 件を JSONL に追記する.

        ファイル名は item.timestamp の日付部分 (YYYY-MM-DD) から決まる.
        破損行が混ざってもテキスト追記なので影響範囲は当該行のみ.
        """
        d = self._feedback_dir()
        # timestamp は ISO8601, 先頭 10 字が "YYYY-MM-DD"
        date_part = (item.timestamp or "")[:10] or "unknown"
        path = d / f"{date_part}.jsonl"
        line = json.dumps(item.model_dump(), ensure_ascii=False) + "\n"
        with path.open("a", encoding="utf-8") as f:
            f.write(line)

    def list_feedback(self, limit: int = 100) -> list[Feedback]:
        """全 jsonl を読み、新しい順に最大 limit 件返す.

        現状は管理画面 (将来) 用のシンプル実装. 件数が増えたら集約処理を入れる.
        """
        d = self.jobs_dir / "_feedback"
        if not d.is_dir():
            return []
        items: list[Feedback] = []
        # 日付降順 (ファイル名が YYYY-MM-DD.jsonl なので文字列ソートで OK)
        for path in sorted(d.glob("*.jsonl"), reverse=True):
            try:
                with path.open("r", encoding="utf-8") as f:
                    for raw in f:
                        raw = raw.strip()
                        if not raw:
                            continue
                        try:
                            items.append(Feedback.model_validate_json(raw))
                        except Exception:  # noqa: BLE001
                            logger.warning("Skipping malformed feedback line in %s", path)
            except OSError:
                logger.warning("Failed to read feedback file %s", path)
            if len(items) >= limit * 2:  # 早期終了 (時刻ソート前なのでバッファ多めに)
                break
        items.sort(key=lambda x: x.timestamp, reverse=True)
        return items[:limit]

    # ------------------------------------------------------------ cells

    def cells_db_path(self, job_id: str) -> Path:
        """cells.db のパス. 物理的に存在するかは別途確認."""
        return self._job_dir(job_id) / "cells.db"

    def has_cells_db(self, job_id: str) -> bool:
        """cells.db が生成済みかどうか."""
        try:
            return self.cells_db_path(job_id).is_file()
        except ValueError:
            return False

    def get_cells_range(
        self,
        job_id: str,
        sheet: str,
        range_str: str,
    ) -> dict[str, object]:
        """指定範囲のセルを 2D 配列で返す.

        Args:
            job_id: ジョブ ID
            sheet: シート名
            range_str: "A6:F10" 形式. 単一セル "A6" でも可

        Returns:
            {"sheet": ..., "range": ..., "origin_row": int, "origin_col": int,
             "rows": [[...], ...]} の dict.
            空セルは null. value (計算結果) と formula (数式テキスト) を別フィールドで返す.
        """
        import sqlite3

        from openpyxl.utils.cell import range_boundaries

        path = self.cells_db_path(job_id)
        if not path.is_file():
            raise FileNotFoundError(f"cells.db not built for job {job_id}")

        try:
            min_col, min_row, max_col, max_row = range_boundaries(range_str)
        except Exception as e:  # noqa: BLE001
            raise ValueError(f"invalid range: {range_str!r}: {e}") from e
        if None in (min_col, min_row, max_col, max_row):
            raise ValueError(f"invalid range (open-ended not supported here): {range_str!r}")

        conn = sqlite3.connect(str(path))
        try:
            cur = conn.execute(
                "SELECT row, col, value, formula, data_type FROM cells "
                "WHERE sheet=? AND row BETWEEN ? AND ? AND col BETWEEN ? AND ?",
                (sheet, min_row, max_row, min_col, max_col),
            )
            n_rows = max_row - min_row + 1
            n_cols = max_col - min_col + 1
            grid: list[list[dict[str, object] | None]] = [
                [None for _ in range(n_cols)] for _ in range(n_rows)
            ]
            for row, col, value, formula, data_type in cur.fetchall():
                ri = row - min_row
                ci = col - min_col
                if 0 <= ri < n_rows and 0 <= ci < n_cols:
                    grid[ri][ci] = {
                        "value": value,
                        "formula": formula,
                        "data_type": data_type,
                    }
        finally:
            conn.close()

        return {
            "sheet": sheet,
            "range": range_str,
            "origin_row": min_row,
            "origin_col": min_col,
            "rows": grid,
        }

    def find_cells(
        self,
        job_id: str,
        query: str,
        sheet: str | None = None,
        limit: int = 20,
    ) -> list[dict[str, object]]:
        """セルの value を部分一致検索する.

        Args:
            job_id: ジョブ ID
            query: 検索文字列 (LIKE '%query%' 相当). 空文字列なら空配列を返す.
            sheet: シート名で絞る (None なら全シート)
            limit: 最大件数 (1〜200 にクランプ)

        Returns:
            [{"sheet", "row", "col", "coord", "value", "formula"}, ...]
        """
        import sqlite3

        if not query:
            return []
        path = self.cells_db_path(job_id)
        if not path.is_file():
            raise FileNotFoundError(f"cells.db not built for job {job_id}")

        limit = max(1, min(200, int(limit)))
        conn = sqlite3.connect(str(path))
        try:
            sql = "SELECT sheet, row, col, coord, value, formula FROM cells WHERE value LIKE ?"
            params: list[object] = [f"%{query}%"]
            if sheet:
                sql += " AND sheet = ?"
                params.append(sheet)
            sql += " LIMIT ?"
            params.append(limit)
            cur = conn.execute(sql, params)
            results: list[dict[str, object]] = []
            for s, r, c, coord, value, formula in cur.fetchall():
                results.append(
                    {
                        "sheet": s,
                        "row": r,
                        "col": c,
                        "coord": coord,
                        "value": value,
                        "formula": formula,
                    }
                )
            return results
        finally:
            conn.close()

    # ----------------------------------------------------------- internals

    @staticmethod
    def _write_json(path: Path, payload: object) -> None:
        path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
