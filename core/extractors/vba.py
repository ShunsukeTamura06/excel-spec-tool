"""VBA抽出モジュール.

olevba (oletools) で .xlsm / .xls / .bas からVBAソースを取り出し、
モジュールとプロシージャを Pydantic モデルに詰める。

SPEC.md §4.1 参照。
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Literal

from core.exceptions import ExtractionError
from core.models import VbaModule, VbaProcedure

logger = logging.getLogger(__name__)


_PROCEDURE_KIND = Literal["Sub", "Function", "Property"]

# 行頭で `[Public|Private|Friend] [Static] (Sub|Function|Property [Get|Let|Set]) Name(...)`
# にマッチする正規表現. `End Sub` などとは区別したいので `^` 起点とする。
_PROC_START_RE = re.compile(
    r"""
    ^[ \t]*                                  # 行頭の空白
    (?:Public[ \t]+|Private[ \t]+|Friend[ \t]+)?
    (?:Static[ \t]+)?
    (?P<kind>Sub|Function|Property)
    (?:[ \t]+(?:Get|Let|Set))?               # Property の場合の修飾子
    [ \t]+
    (?P<name>[A-Za-z_][A-Za-z0-9_]*)
    """,
    re.IGNORECASE | re.VERBOSE,
)

# 対応する End ステートメント
_PROC_END_RE = re.compile(
    r"^[ \t]*End[ \t]+(?P<kind>Sub|Function|Property)\b",
    re.IGNORECASE,
)


def _parse_procedures(code: str) -> list[VbaProcedure]:
    """1モジュールのソースから手続きを切り出す.

    `Sub` / `Function` / `Property` 宣言行から対応する `End ...` 行までを1プロシージャとする。
    入れ子は VBA の言語仕様上発生しないので素朴に走査する。

    Args:
        code: モジュールのソース全体.

    Returns:
        検出された VbaProcedure のリスト. 行番号は1始まり.
    """
    lines = code.splitlines()
    procs: list[VbaProcedure] = []

    i = 0
    n = len(lines)
    while i < n:
        m = _PROC_START_RE.match(lines[i])
        if not m:
            i += 1
            continue

        start = i
        kind_str = m.group("kind").capitalize()
        # `Property Get/Let/Set` も kind は "Property" として扱う (SPEC.md の Literal 型に合わせる)
        kind: _PROCEDURE_KIND = kind_str  # type: ignore[assignment]
        name = m.group("name")

        # 対応する End を探す
        j = i + 1
        end = i  # 見つからなければ宣言行のみ
        while j < n:
            em = _PROC_END_RE.match(lines[j])
            if em and em.group("kind").capitalize() == kind_str:
                end = j
                break
            # 別のプロシージャ宣言が現れたらそこで打ち切る (End 抜け対策)
            if _PROC_START_RE.match(lines[j]):
                end = j - 1
                j -= 1
                break
            j += 1
        else:
            end = n - 1

        procs.append(
            VbaProcedure(
                name=name,
                kind=kind,
                start_line=start + 1,
                end_line=end + 1,
                code="\n".join(lines[start : end + 1]),
            )
        )
        i = end + 1

    return procs


def _module_type_from_filename(vba_filename: str) -> Literal["Module", "Class", "Form", "Document"]:
    """olevba が返すモジュールファイル名から種別を推定する.

    olevba は通常 `.bas`(標準モジュール) / `.cls`(クラス/Document) / `.frm`(フォーム) を返す。
    ThisWorkbook や Sheet などのドキュメントモジュールも `.cls` で返るが、
    名前から区別するのは難しいので Class として扱う。
    """
    name_lower = vba_filename.lower()
    if name_lower.endswith(".bas"):
        return "Module"
    if name_lower.endswith(".cls"):
        return "Class"
    if name_lower.endswith(".frm"):
        return "Form"
    return "Module"


def _module_name_from_filename(vba_filename: str) -> str:
    """olevba の vba_filename ("Module1.bas" 等) からモジュール名を抜き出す."""
    stem = Path(vba_filename).stem
    return stem or vba_filename


def extract_vba(file_path: Path) -> list[VbaModule]:
    """Excelファイル (.xlsm / .xls) または .bas からVBAモジュールを抽出する.

    Args:
        file_path: 対象ファイルのパス.

    Returns:
        抽出された VbaModule のリスト. VBAが含まれない場合は空リスト.

    Raises:
        ExtractionError: olevba がパースに失敗した場合.

    Notes:
        - パスワード保護されたVBAは現状サポートしない (TODO).
        - olevba の型情報がないため、内部で `type: ignore` を使う。
    """
    # olevba は遅延 import (起動コスト低減のため)
    from oletools.olevba import VBA_Parser

    if not file_path.exists():
        raise ExtractionError(f"File not found: {file_path}")

    try:
        parser = VBA_Parser(str(file_path))
    except Exception as e:  # noqa: BLE001 - olevba は様々な例外を投げる
        raise ExtractionError(f"Failed to open VBA parser for {file_path}: {e}") from e

    try:
        if not parser.detect_vba_macros():
            return []

        # TODO: パスワード保護 VBA への対応 (現状は olevba がエラーを返したらログ出力して空にする)
        modules: list[VbaModule] = []
        for _filename, _stream_path, vba_filename, vba_code in parser.extract_macros():
            if not vba_code:
                continue
            code_str = (
                vba_code.decode("utf-8", errors="replace")
                if isinstance(vba_code, bytes)
                else vba_code
            )
            module = VbaModule(
                name=_module_name_from_filename(vba_filename),
                type=_module_type_from_filename(vba_filename),
                code=code_str,
                procedures=_parse_procedures(code_str),
            )
            modules.append(module)
        return modules
    except ExtractionError:
        raise
    except Exception as e:  # noqa: BLE001
        logger.exception("Unexpected failure while extracting VBA from %s", file_path)
        raise ExtractionError(f"VBA extraction failed for {file_path}: {e}") from e
    finally:
        parser.close()
