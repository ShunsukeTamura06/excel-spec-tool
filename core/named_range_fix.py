"""名前付き範囲の定義修正 (S2 増分1: 限定自動リファクタの最初のパターン).

安全パターンの中で最も影響範囲が狭い「名前定義の参照先を書き換える」だけを対象にする。
数式の書き換えや削除は扱わない (別増分)。ワークブックスコープの名前定義
(`wb.defined_names`) のみを対象とし、シートスコープの名前定義は対象外。

フロー:
  1. `propose_named_range_fix()` で「適用したら何が変わるか」を書き込みなしで計算する
     (チャットの tool loop から自動で呼んでよい、read-only)。
  2. ユーザーが影響を見て納得したら `apply_named_range_fix()` で実ファイルに書き込む
     (人間の明示操作からのみ呼ぶ。tool loop からは呼ばない — docs/VISION.ja.md §4.2)。
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook as load_openpyxl_workbook
from openpyxl.workbook.defined_name import DefinedName

from core.exceptions import NamedRangeFixError
from core.models import ReferenceIndex, Workbook, WorkbookDiff
from core.openpyxl_utils import close_workbook
from core.workbook_diff import build_blast_radius, diff_named_ranges


def propose_named_range_fix(
    before_wb: Workbook,
    before_index: ReferenceIndex,
    name: str,
    new_refers_to: str,
) -> WorkbookDiff:
    """name の refers_to を new_refers_to に変えたら何が変わるかをメモリ内で試算する.

    実ファイルには一切書き込まない。before_wb を複製し、対象の NamedRange だけを
    差し替えた仮想の after を作って diff を計算する (このパターンではセル/条件付き
    書式等は変化しないため、それらは必ず空になる)。

    Args:
        before_wb: 対象ジョブの抽出済み Workbook.
        before_index: 対象ジョブの ReferenceIndex (波及範囲算出に使う).
        name: 書き換える名前付き範囲の名前.
        new_refers_to: 新しい参照先 (例: "Data!$A$1:$A$100").

    Returns:
        WorkbookDiff。named_ranges に1件、該当があれば blast_radius に掲載される。

    Raises:
        NamedRangeFixError: name が before_wb 内のどのシートにも見つからない場合.
    """
    after_wb = before_wb.model_copy(deep=True)
    found = False
    for sheet in after_wb.sheets:
        for nr in sheet.named_ranges:
            if nr.name == name:
                nr.refers_to = new_refers_to
                found = True
    if not found:
        raise NamedRangeFixError(f"named range not found: {name}")

    named_ranges = diff_named_ranges(before_wb, after_wb)
    blast_radius = build_blast_radius([], named_ranges, before_index)

    return WorkbookDiff(
        before_filename=before_wb.filename,
        after_filename=before_wb.filename,
        named_ranges=named_ranges,
        blast_radius=blast_radius,
        existing_risks=list(before_wb.analysis_risks),
    )


def apply_named_range_fix(
    file_path: Path,
    name: str,
    new_refers_to: str,
    out_path: Path,
) -> None:
    """name の定義を new_refers_to に書き換えた新しい xlsx/xlsm を out_path に書き出す.

    file_path 自体は変更しない (別ファイルへの出力)。

    Args:
        file_path: 元ファイル (ジョブの original.*).
        name: 書き換える名前付き範囲の名前.
        new_refers_to: 新しい参照先.
        out_path: 書き出し先パス.

    Raises:
        NamedRangeFixError: file_path が開けない、name が定義されていない、
            または保存に失敗した場合.
    """
    try:
        wb = load_openpyxl_workbook(file_path, keep_vba=True)
    except Exception as exc:  # noqa: BLE001 - 失敗内容こそ診断材料
        raise NamedRangeFixError(f"failed to open workbook: {file_path}: {exc}") from exc

    try:
        if name not in wb.defined_names:
            raise NamedRangeFixError(f"named range not found in {file_path}: {name}")

        wb.defined_names[name] = DefinedName(name=name, attr_text=new_refers_to)
        wb.save(out_path)
    except NamedRangeFixError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise NamedRangeFixError(f"failed to save workbook: {out_path}: {exc}") from exc
    finally:
        close_workbook(wb)
