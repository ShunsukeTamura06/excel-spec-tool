"""openpyxlオブジェクトのライフサイクル補助."""

from __future__ import annotations


def close_workbook(workbook: object) -> None:
    """Workbook本体とkeep_vba用の複製ZipFileを確実に閉じる.

    openpyxlの通常の ``Workbook.close()`` はnormal modeで生成された
    ``vba_archive`` を閉じないため、明示的に両方を解放する。

    Args:
        workbook: openpyxlが返したWorkbookオブジェクト。
    """

    vba_archive = getattr(workbook, "vba_archive", None)
    close_vba_archive = getattr(vba_archive, "close", None)
    if callable(close_vba_archive):
        close_vba_archive()

    close = getattr(workbook, "close", None)
    if callable(close):
        close()
