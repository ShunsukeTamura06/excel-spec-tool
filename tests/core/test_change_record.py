"""変更検証監査レコードのシリアライズ契約テスト."""

from core.change_record import ChangeExecutionRecord
from core.models import WorkbookDiff
from core.mutation import MutationPlan, MutationResult, NamedRangeSetOperation
from core.verification import VerificationReport


def test_change_record_round_trips_all_evidence() -> None:
    """成果物ハッシュを含む検証証拠がJSON往復で欠落しない."""

    plan = MutationPlan(
        source_job_id="00000000-0000-4000-8000-000000000000",
        operation=NamedRangeSetOperation(name="TaxRate", new_refers_to="Data!$A$2"),
    )
    diff = WorkbookDiff(before_filename="before.xlsx", after_filename="after.xlsx")
    record = ChangeExecutionRecord(
        source_job_id=plan.source_job_id,
        result_job_id="00000000-0000-4000-8000-000000000001",
        source_file_sha256="a" * 64,
        result_file_sha256="b" * 64,
        plan=plan,
        provider_result=MutationResult(
            provider="openpyxl",
            provider_version="test",
            operation="named_range_set",
            changed_count=1,
        ),
        expected_diff=diff,
        actual_diff=diff,
        verification=VerificationReport(
            status="passed",
            expected_change_count=0,
            actual_change_count=0,
        ),
    )

    restored = ChangeExecutionRecord.model_validate_json(record.model_dump_json())

    assert restored == record
    assert restored.source_file_sha256 != restored.result_file_sha256
