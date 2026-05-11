"""ローカルファイル永続化モジュール.

SPEC.md §5.2 に基づき、ジョブごとに以下の構成でファイルを管理する:

    {jobs_dir}/{job_id}/
    ├── original.{ext}       # アップロード原本
    ├── extracted.json       # Workbook モデル (Core 抽出結果)
    ├── spec.md              # 生成済み設計書
    ├── references.json      # ReferenceIndex モデル
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
from pathlib import Path

from core.models import (
    ChatMessage,
    JobMeta,
    ReferenceIndex,
    Workbook,
)

logger = logging.getLogger(__name__)


_DEFAULT_JOBS_DIR = "./jobs"
_DIR_MODE = 0o700

# UUID v4: 8-4-4-4-12 hex, with version 4 nibble at the right position
_UUID_V4_RE = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$",
    re.IGNORECASE,
)


class StorageError(Exception):
    """Storage 層の例外."""


class JobNotFoundError(StorageError):
    """指定 job_id が存在しない."""


def _validate_job_id(job_id: str) -> None:
    """job_id が UUIDv4 形式か検証する. 不正なら ValueError を投げる."""
    if not isinstance(job_id, str) or not _UUID_V4_RE.match(job_id):
        raise ValueError(f"invalid job_id (must be UUIDv4): {job_id!r}")


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
        )
        self._write_json(d / "meta.json", meta.model_dump())
        return meta

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

    # ------------------------------------------------------------ meta

    def get_meta(self, job_id: str) -> JobMeta:
        d = self._require_job_dir(job_id)
        return JobMeta.model_validate_json((d / "meta.json").read_text(encoding="utf-8"))

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

    # ------------------------------------------------------ workbook / spec / refs

    def save_workbook(self, job_id: str, wb: Workbook) -> None:
        d = self._require_job_dir(job_id)
        self._write_json(d / "extracted.json", wb.model_dump())

    def load_workbook(self, job_id: str) -> Workbook:
        d = self._require_job_dir(job_id)
        return Workbook.model_validate_json((d / "extracted.json").read_text(encoding="utf-8"))

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

    def append_chat_message(self, job_id: str, message: ChatMessage) -> None:
        d = self._require_job_dir(job_id)
        path = d / "chat_history.jsonl"
        line = json.dumps(message.model_dump(), ensure_ascii=False) + "\n"
        with path.open("a", encoding="utf-8") as f:
            f.write(line)

    def load_chat_history(self, job_id: str) -> list[ChatMessage]:
        d = self._require_job_dir(job_id)
        path = d / "chat_history.jsonl"
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
