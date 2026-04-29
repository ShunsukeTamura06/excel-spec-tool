"""参照インデックス構築モジュール.

Workbook (数式 + VBA) を走査し、「あるセル/範囲を参照しているもの」を
逆引きできる ReferenceIndex を作る。

SPEC.md §4.3 参照。
"""

from __future__ import annotations

import logging
import re
from collections import defaultdict

from core.models import Reference, ReferenceIndex, VbaModule, Workbook

logger = logging.getLogger(__name__)


# --- VBA 参照パターン ---------------------------------------------------------
# 完璧なパースは目指さない。CLAUDE.md §4 の最低限テスト要件を満たすことが目標。

# `Worksheets("Calc").Range("A1:B2")` または `Sheets("Calc").Range("A1")`
_RE_SHEETS_RANGE = re.compile(
    r"""
    (?:Worksheets|Sheets)\s*\(\s*"(?P<sheet>[^"]+)"\s*\)
    \s*\.\s*Range\s*\(\s*"(?P<range>[^"]+)"\s*\)
    """,
    re.IGNORECASE | re.VERBOSE,
)

# `Sheets("Calc").Cells(2, 8)` → 行2列8 = H2
_RE_SHEETS_CELLS = re.compile(
    r"""
    (?:Worksheets|Sheets)\s*\(\s*"(?P<sheet>[^"]+)"\s*\)
    \s*\.\s*Cells\s*\(\s*(?P<row>\d+)\s*,\s*(?P<col>\d+)\s*\)
    """,
    re.IGNORECASE | re.VERBOSE,
)

# `[Calc!A1]` 短縮表記
_RE_BRACKET = re.compile(r"\[(?P<sheet>[^\[\]!]+)!(?P<range>[^\[\]]+)\]")

# `Range("A1:J100")` シート修飾なし (現在のシート)
# 上の SHEETS_RANGE が消費した後の残りに対して適用するため、最後に走らせる。
_RE_RANGE_BARE = re.compile(
    r"""
    (?<![A-Za-z0-9_\.])    # 直前が識別子でない (Worksheets(...).Range は除外)
    Range\s*\(\s*"(?P<range>[^"]+)"\s*\)
    """,
    re.IGNORECASE | re.VERBOSE,
)


def _col_num_to_letters(col: int) -> str:
    """1始まりの列番号をExcel列文字に変換 (1 -> A, 27 -> AA)."""
    if col < 1:
        return ""
    s = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        s = chr(ord("A") + rem) + s
    return s


def _qualify(sheet: str | None, ref: str) -> str:
    """シート名と範囲を結合して `Sheet!Range` 形式に. シートが None なら範囲のみ."""
    if sheet:
        return f"{sheet}!{ref}"
    return ref


def _find_enclosing_procedure(module: VbaModule, line_num: int) -> str | None:
    """指定行を含むプロシージャ名を返す. どのプロシージャにも属さなければ None."""
    for p in module.procedures:
        if p.start_line <= line_num <= p.end_line:
            return p.name
    return None


def _from_label(module: VbaModule, line_num: int) -> str:
    """VBA参照の from_ ラベルを `Module1.UpdateDaily:L47` の形に整える."""
    proc = _find_enclosing_procedure(module, line_num)
    if proc:
        return f"{module.name}.{proc}:L{line_num}"
    return f"{module.name}:L{line_num}"


def _mask_span(buf: list[str], span: tuple[int, int]) -> None:
    """`buf` の `span` 範囲 (開始, 終了) を空白で塗りつぶす (in-place)."""
    start, end = span
    for i in range(start, end):
        if 0 <= i < len(buf):
            buf[i] = " "


def _scan_vba_module(module: VbaModule) -> list[Reference]:
    """1モジュールの code を走査して Reference を抽出する.

    重複検出を避けるため、長く特異なパターンから順にマッチを記録し、
    マッチ済み区間は後段のパターンが拾わないように line ごとにマスクする。
    """
    refs: list[Reference] = []
    if not module.code:
        return refs

    lines = module.code.splitlines()
    for idx, raw_line in enumerate(lines, start=1):
        # コメント (`'`) より右側は除外する.
        # 文字列リテラル内の `'` は誤検出するが、簡素化のためスキップ。
        line = raw_line.split("'", 1)[0]

        # マスク用 (マッチした範囲を空白で塗りつぶし、後段パターンの誤検出を防ぐ)
        masked = list(line)

        # 1. Worksheets/Sheets("...").Range("...")
        for m in _RE_SHEETS_RANGE.finditer(line):
            refs.append(
                Reference(
                    kind="vba",
                    from_=_from_label(module, idx),
                    to=_qualify(m.group("sheet"), m.group("range")),
                    code=m.group(0),
                )
            )
            _mask_span(masked, m.span())

        # 2. Worksheets/Sheets("...").Cells(r, c)
        cur = "".join(masked)
        for m in _RE_SHEETS_CELLS.finditer(cur):
            row = int(m.group("row"))
            col = int(m.group("col"))
            cell_ref = f"{_col_num_to_letters(col)}{row}"
            refs.append(
                Reference(
                    kind="vba",
                    from_=_from_label(module, idx),
                    to=_qualify(m.group("sheet"), cell_ref),
                    code=m.group(0),
                )
            )
            _mask_span(masked, m.span())

        # 3. [Calc!A1] 短縮表記
        cur = "".join(masked)
        for m in _RE_BRACKET.finditer(cur):
            refs.append(
                Reference(
                    kind="vba",
                    from_=_from_label(module, idx),
                    to=_qualify(m.group("sheet").strip(), m.group("range").strip()),
                    code=m.group(0),
                )
            )
            _mask_span(masked, m.span())

        # 4. Range("...") シート修飾なし
        cur = "".join(masked)
        for m in _RE_RANGE_BARE.finditer(cur):
            refs.append(
                Reference(
                    kind="vba",
                    from_=_from_label(module, idx),
                    to=m.group("range"),
                    code=m.group(0),
                )
            )

    return refs


def build_reference_index(wb: Workbook) -> ReferenceIndex:
    """Workbook 全体から参照の逆引きインデックスを構築する.

    Args:
        wb: Workbook モデル. sheets/formulas と vba_modules/code が埋まっている前提.

    Returns:
        ReferenceIndex: refs[参照先] -> [Reference, ...]
    """
    bucket: dict[str, list[Reference]] = defaultdict(list)

    # 数式側: 各 CellFormula.refs を逆引きする
    for sheet in wb.sheets:
        for f in sheet.formulas:
            for target in f.refs:
                bucket[target].append(
                    Reference(
                        kind="formula",
                        from_=f.coord,
                        to=target,
                        code=f.formula,
                    )
                )

    # VBA側: code を正規表現で走査
    for module in wb.vba_modules:
        for ref in _scan_vba_module(module):
            bucket[ref.to].append(ref)

    return ReferenceIndex(refs=dict(bucket))
