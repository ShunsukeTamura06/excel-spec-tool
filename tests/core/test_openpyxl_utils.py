"""openpyxlリソース解放補助のテスト."""

from core.openpyxl_utils import close_workbook


class _Closable:
    """close呼び出し回数を記録するテスト用オブジェクト."""

    def __init__(self) -> None:
        self.close_count = 0

    def close(self) -> None:
        """呼び出し回数を1増やす."""

        self.close_count += 1


class _Workbook:
    """vba_archiveを持つ最小のWorkbook代替."""

    def __init__(self) -> None:
        self.vba_archive = _Closable()
        self.close_count = 0

    def close(self) -> None:
        """Workbook側のclose呼び出し回数を1増やす."""

        self.close_count += 1


def test_close_workbook_closes_vba_archive_and_workbook() -> None:
    """keep_vba用ZipFileとWorkbook本体の両方を閉じる."""

    workbook = _Workbook()

    close_workbook(workbook)

    assert workbook.vba_archive.close_count == 1
    assert workbook.close_count == 1


def test_close_workbook_accepts_object_without_close() -> None:
    """closeを持たない代替オブジェクトでも例外にしない."""

    close_workbook(object())
