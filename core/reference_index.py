"""参照インデックス構築モジュール.

Workbook (数式 + VBA) を走査し、「あるセル/範囲を参照しているもの」を
逆引きできる ReferenceIndex を作る。

参照キーは正規化済み (シート修飾を補い、`$` を除去し、列文字を大文字化) で
格納されるため、`Calc!H2` / `$H$2` / `h2` などの表記揺れは同一キーになる。
範囲交差での検索は `find_overlapping()` を使う。

SPEC.md §4.3 参照。
"""

from __future__ import annotations

import logging
import re
from collections import defaultdict
from dataclasses import dataclass

from core.models import Reference, ReferenceIndex, VbaModule, Workbook

logger = logging.getLogger(__name__)


# Excel の論理上の最大行・列. 全列指定 (A:A) を「行=1..MAX」として表現する目的で使う.
_EXCEL_MAX_ROW = 1_048_576
_EXCEL_MAX_COL = 16_384


@dataclass(frozen=True)
class ParsedRef:
    """参照を構造化した形で保持する.

    シートが None の場合、`overlaps()` は wildcard 扱いになる
    (VBA の bare `Range("A1")` のようにシート不明の参照を表す)。
    """

    sheet: str | None
    min_row: int
    min_col: int
    max_row: int
    max_col: int

    def overlaps(self, other: ParsedRef) -> bool:
        """別の ParsedRef と範囲が重なるか.

        シートが両方とも判明していて異なる場合は不一致。どちらかが
        None ならシート不問 (wildcard)。
        """
        if self.sheet is not None and other.sheet is not None and self.sheet != other.sheet:
            return False
        return not (
            self.max_row < other.min_row
            or self.min_row > other.max_row
            or self.max_col < other.min_col
            or self.min_col > other.max_col
        )


# シート修飾を切り出す: `Sheet1!A1` / `'My Sheet'!A1` / `''Quoted''!A1` の3形態
_SHEET_QUALIFIED_RE = re.compile(r"^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$")
# 全列指定 (例 `A:A`, `AA:AB`)
_FULL_COL_RE = re.compile(r"^([A-Z]+):([A-Z]+)$")
# 全行指定 (例 `1:5`)
_FULL_ROW_RE = re.compile(r"^(\d+):(\d+)$")


def _parse_ref(ref: str, owner_sheet: str | None = None) -> ParsedRef | None:
    """参照文字列を `ParsedRef` に正規化する.

    Args:
        ref: 参照文字列. 例: `A1`, `Calc!H2:H100`, `'My Sheet'!A1`,
            `$A$1:$B$2`, `A:A`, `1:1`.
        owner_sheet: シート修飾が ref に含まれていない場合に補うシート名。
            数式から取り出した ref ではその数式が属するシートを渡すことで、
            同一シート内の相対参照 (`A2`) も `Calc!A2` として正規化される。

    Returns:
        パースできれば `ParsedRef`. ファンクション名・テーブル名など範囲として
        解釈できない場合は None.
    """
    s = ref.strip()
    if not s:
        return None

    sheet = owner_sheet
    range_part = s
    m = _SHEET_QUALIFIED_RE.match(s)
    if m:
        raw_sheet = m.group(1) if m.group(1) is not None else m.group(2)
        # Excel は引用符内で `''` を `'` のエスケープにする
        sheet = raw_sheet.replace("''", "'").strip()
        range_part = m.group(3).strip()

    range_part = range_part.replace("$", "").upper()
    if not range_part:
        return None

    m_col = _FULL_COL_RE.match(range_part)
    if m_col:
        try:
            from openpyxl.utils.cell import column_index_from_string

            a = column_index_from_string(m_col.group(1))
            b = column_index_from_string(m_col.group(2))
        except Exception:  # noqa: BLE001
            return None
        return ParsedRef(sheet, 1, min(a, b), _EXCEL_MAX_ROW, max(a, b))

    m_row = _FULL_ROW_RE.match(range_part)
    if m_row:
        a_row = int(m_row.group(1))
        b_row = int(m_row.group(2))
        return ParsedRef(sheet, min(a_row, b_row), 1, max(a_row, b_row), _EXCEL_MAX_COL)

    try:
        from openpyxl.utils.cell import range_boundaries

        min_col, min_row, max_col, max_row = range_boundaries(range_part)
    except Exception:  # noqa: BLE001
        return None
    if min_col is None or min_row is None or max_col is None or max_row is None:
        return None
    return ParsedRef(
        sheet,
        int(min_row),
        int(min_col),
        int(max_row),
        int(max_col),
    )


def _canonical_key(parsed: ParsedRef) -> str:
    """`ParsedRef` から逆引きインデックスのキー文字列を生成する.

    フォーマット:
      - シートあり: `Sheet!RANGE`
      - シートなし (VBA bare): `RANGE`

    RANGE は次のルールで再構築する:
      - 全列: `A:A` / `A:Z`
      - 全行: `1:1` / `1:5`
      - 単一セル: `A1`
      - 通常範囲: `A1:B2`
    """
    from openpyxl.utils import get_column_letter

    full_col = parsed.min_row == 1 and parsed.max_row == _EXCEL_MAX_ROW
    full_row = parsed.min_col == 1 and parsed.max_col == _EXCEL_MAX_COL

    if full_col and not full_row:
        a = get_column_letter(parsed.min_col)
        b = get_column_letter(parsed.max_col)
        range_str = f"{a}:{a}" if a == b else f"{a}:{b}"
    elif full_row and not full_col:
        a_row = parsed.min_row
        b_row = parsed.max_row
        range_str = f"{a_row}:{a_row}" if a_row == b_row else f"{a_row}:{b_row}"
    else:
        a_col = get_column_letter(parsed.min_col)
        b_col = get_column_letter(parsed.max_col)
        if parsed.min_row == parsed.max_row and parsed.min_col == parsed.max_col:
            range_str = f"{a_col}{parsed.min_row}"
        else:
            range_str = f"{a_col}{parsed.min_row}:{b_col}{parsed.max_row}"

    if parsed.sheet:
        return f"{parsed.sheet}!{range_str}"
    return range_str


def _normalize_target(raw: str, owner_sheet: str | None = None) -> tuple[str, ParsedRef | None]:
    """raw 文字列を (canonical_key, parsed) に正規化する.

    パース不能なら (raw, None) をそのまま返す (テーブル名・関数名等)。
    """
    parsed = _parse_ref(raw, owner_sheet=owner_sheet)
    if parsed is None:
        return raw, None
    return _canonical_key(parsed), parsed


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

    キーは canonical 形式 (シート修飾付き / `$` 除去 / 列大文字) に正規化される。
    パース不能な参照 (関数名・テーブル名等) は raw 文字列のままキーになる。

    Args:
        wb: Workbook モデル. sheets/formulas と vba_modules/code が埋まっている前提.

    Returns:
        ReferenceIndex: refs[正規化キー] -> [Reference, ...]
    """
    bucket: dict[str, list[Reference]] = defaultdict(list)

    # 数式側: 各 CellFormula.refs を逆引きする. 同一シート内の相対参照 (例 `A2`)
    # にはオーナーシート名を補って `Calc!A2` に揃える.
    for sheet in wb.sheets:
        for f in sheet.formulas:
            for raw_target in f.refs:
                key, _ = _normalize_target(raw_target, owner_sheet=sheet.name)
                bucket[key].append(
                    Reference(
                        kind="formula",
                        from_=f.coord,
                        to=key,
                        code=f.formula,
                    )
                )

    # VBA側: code を正規表現で走査. _scan_vba_module が返した raw `to` を
    # canonical キーに置換して登録する.
    for module in wb.vba_modules:
        for ref in _scan_vba_module(module):
            key, _ = _normalize_target(ref.to, owner_sheet=None)
            bucket[key].append(
                Reference(
                    kind=ref.kind,
                    from_=ref.from_,
                    to=key,
                    code=ref.code,
                )
            )

    return ReferenceIndex(refs=dict(bucket))


def find_overlapping(
    index: ReferenceIndex,
    target: str,
    owner_sheet: str | None = None,
) -> list[Reference]:
    """target と範囲が重なる参照を全件返す.

    完全一致だけでなく、範囲交差で検出する。例:
      - index に `Input!A:A` があり target=`Input!A5` → ヒット
      - index に `Calc!A1:J100` があり target=`Calc!H2` → ヒット
      - index に sheet 不明の VBA bare 参照 (`A1:J100`) があれば、
        target のシートに関係なく範囲が重なればヒット (wildcard 扱い)

    パース不能なキーは厳密文字列一致でのみマッチする。

    Args:
        index: `build_reference_index` の結果.
        target: 検索したい参照先 (例: `Calc!H2`, `Input!A5`).
        owner_sheet: target がシート修飾なしの場合に補うシート名 (任意).

    Returns:
        重なる Reference を index 登録順で返す. 重複は除去しない
        (1 つの数式が同じセルを 2 回参照していれば 2 件返る)。
    """
    target_key, target_parsed = _normalize_target(target, owner_sheet=owner_sheet)

    # target がパース不能: 完全一致のみ
    if target_parsed is None:
        return list(index.refs.get(target_key, []))

    results: list[Reference] = []
    seen_keys: set[str] = set()

    # 1) 完全一致 (高速パス)
    exact = index.refs.get(target_key)
    if exact:
        results.extend(exact)
        seen_keys.add(target_key)

    # 2) 範囲交差
    for key, refs in index.refs.items():
        if key in seen_keys:
            continue
        parsed = _parse_ref(key, owner_sheet=None)
        if parsed is None:
            continue
        if parsed.overlaps(target_parsed):
            results.extend(refs)

    return results
