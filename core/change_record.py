"""変更計画から検証結果までを保持する監査レコード."""

from __future__ import annotations

from datetime import datetime, timezone

from pydantic import BaseModel, Field

from core.models import WorkbookDiff
from core.mutation import MutationPlan, MutationResult
from core.verification import VerificationReport


def _utc_now_iso() -> str:
    """現在時刻をUTCのISO 8601文字列で返す."""

    return datetime.now(timezone.utc).isoformat()


class ChangeExecutionRecord(BaseModel):
    """誰が何で変更し、何を期待し、実際に何が起きたかを保存する."""

    created_at: str = Field(default_factory=_utc_now_iso)
    source_job_id: str
    result_job_id: str
    source_file_sha256: str
    result_file_sha256: str
    plan: MutationPlan
    provider_result: MutationResult
    expected_diff: WorkbookDiff
    actual_diff: WorkbookDiff
    verification: VerificationReport
