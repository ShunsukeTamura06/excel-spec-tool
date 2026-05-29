"""未解析リスク検出モジュール.

静的解析で依存関係を断定できない箇所を、LLM とユーザーに明示するための
best-effort 検出を行う。目的は完全な解釈ではなく、危険な断定を避けること。
"""

from __future__ import annotations

import re
from collections.abc import Iterable

from core.models import AnalysisRisk, Workbook

_DYNAMIC_FORMULA_FUNCTIONS = {
    "INDIRECT": "文字列から参照先を組み立てるため、静的な参照解析では追跡できません。",
    "OFFSET": "基準セルから実行時に範囲をずらすため、影響範囲が不確定です。",
    "CELL": "環境や参照セルのメタ情報に依存する可能性があります。",
    "INFO": "Excel 実行環境に依存する情報を返すため、再現性確認が必要です。",
}

_RUNTIME_STATE_PATTERNS = {
    "ActiveSheet": "実行時にアクティブなシートへ依存します。",
    "Selection": "実行時の選択範囲へ依存します。",
    "CurrentRegion": "周辺セルの空白状態で対象範囲が変わります。",
    "UsedRange": "Excel が認識する使用範囲に依存します。",
    "Offset": "実行時に基準範囲からずれた範囲を参照します。",
    "Resize": "実行時に範囲サイズを変更します。",
    "Evaluate": "文字列式を実行時に評価します。",
    "Application.Run": "実行時に呼び出し先マクロが決まる可能性があります。",
}

_EVENT_PROC_RE = re.compile(
    r"^\s*(?:Private\s+)?Sub\s+"
    r"(?P<name>(?:Worksheet|Workbook)_[A-Za-z0-9_]+)\s*\(",
    re.IGNORECASE,
)
_RANGE_CALL_RE = re.compile(r"\bRange\s*\((?P<args>[^)]*)\)", re.IGNORECASE)
_CELLS_CALL_RE = re.compile(r"\bCells\s*\((?P<args>[^)]*)\)", re.IGNORECASE)
_WORKSHEET_DYNAMIC_RE = re.compile(r"\b(?:Worksheets|Sheets)\s*\((?P<args>[^)]*)\)", re.IGNORECASE)


def detect_analysis_risks(wb: Workbook) -> list[AnalysisRisk]:
    """Workbook から未解析リスクを検出する.

    Args:
        wb: 抽出済み Workbook。

    Returns:
        未解析・不確定としてユーザーに明示すべきリスク一覧。
    """
    risks: list[AnalysisRisk] = []
    risks.extend(_detect_formula_risks(wb))
    risks.extend(_detect_vba_risks(wb))
    risks.extend(_detect_external_dependency_risks(wb))
    risks.extend(_detect_object_dependency_risks(wb))
    return _dedupe_risks(risks)


def _detect_formula_risks(wb: Workbook) -> Iterable[AnalysisRisk]:
    """動的参照や環境依存の数式関数を検出する."""
    for sheet in wb.sheets:
        for formula in sheet.formulas:
            upper = formula.formula.upper()
            for fn_name, description in _DYNAMIC_FORMULA_FUNCTIONS.items():
                if not re.search(rf"\b{re.escape(fn_name)}\s*\(", upper):
                    continue
                yield AnalysisRisk(
                    category="dynamic_formula",
                    severity="high" if fn_name in {"INDIRECT", "OFFSET"} else "medium",
                    location=formula.coord,
                    evidence=formula.formula[:500],
                    description=f"{fn_name} 関数を含むため、{description}",
                    recommendation=(
                        "参照先や返り値を Excel 上で確認し、"
                        "改修後に再計算結果を手動確認してください。"
                    ),
                )
            for external_fn in formula.external_functions:
                yield AnalysisRisk(
                    category="external_dependency",
                    severity="medium",
                    location=formula.coord,
                    evidence=formula.formula[:500],
                    description=f"外部 Add-In 関数 `{external_fn}` を使用しています。",
                    recommendation=(
                        "外部データ取得条件、再計算タイミング、"
                        "Add-In 利用可否を実環境で確認してください。"
                    ),
                )


def _detect_vba_risks(wb: Workbook) -> Iterable[AnalysisRisk]:
    """VBA 内の動的参照・実行時状態依存・イベント処理を検出する."""
    for module in wb.vba_modules:
        for line_no, raw_line in enumerate(module.code.splitlines(), start=1):
            line = _strip_vba_comment(raw_line)
            if not line.strip():
                continue
            masked = _mask_vba_strings(line)
            location = f"{module.name}:L{line_no}"

            event_match = _EVENT_PROC_RE.match(masked)
            if event_match:
                event_name = event_match.group("name")
                yield AnalysisRisk(
                    category="event_macro",
                    severity="high",
                    location=location,
                    evidence=raw_line.strip()[:500],
                    description=(
                        f"`{event_name}` はユーザー操作やブック起動時に"
                        "暗黙実行される可能性があります。"
                    ),
                    recommendation=(
                        "改修後に該当イベントの発火条件と副作用を実Excelで確認してください。"
                    ),
                )

            for token, description in _RUNTIME_STATE_PATTERNS.items():
                if not _contains_token(masked, token):
                    continue
                yield AnalysisRisk(
                    category="runtime_state",
                    severity="high" if token in {"ActiveSheet", "Selection"} else "medium",
                    location=location,
                    evidence=raw_line.strip()[:500],
                    description=description,
                    recommendation=(
                        "静的参照だけでは影響範囲を断定できません。"
                        "実行前後の対象範囲を確認してください。"
                    ),
                )

            for match in _RANGE_CALL_RE.finditer(masked):
                args = match.group("args").strip()
                if args and not _is_single_string_literal(args):
                    yield AnalysisRisk(
                        category="dynamic_vba",
                        severity="high",
                        location=location,
                        evidence=raw_line.strip()[:500],
                        description=(
                            "`Range(...)` の参照先が文字列リテラルではなく、実行時に決まります。"
                        ),
                        recommendation=(
                            "変数に入るアドレスを追跡し、対象セル/範囲を手動確認してください。"
                        ),
                    )

            for match in _CELLS_CALL_RE.finditer(masked):
                args = [p.strip() for p in match.group("args").split(",")]
                if len(args) >= 2 and (not args[0].isdigit() or not args[1].isdigit()):
                    yield AnalysisRisk(
                        category="dynamic_vba",
                        severity="high",
                        location=location,
                        evidence=raw_line.strip()[:500],
                        description="`Cells(row, col)` の行または列が実行時に決まります。",
                        recommendation=(
                            "ループ変数や条件分岐を確認し、対象範囲を手動確認してください。"
                        ),
                    )

            for match in _WORKSHEET_DYNAMIC_RE.finditer(masked):
                args = match.group("args").strip()
                if args and not _is_single_string_literal(args) and not args.isdigit():
                    yield AnalysisRisk(
                        category="dynamic_vba",
                        severity="high",
                        location=location,
                        evidence=raw_line.strip()[:500],
                        description=(
                            "`Worksheets(...)` / `Sheets(...)` の対象シートが実行時に決まります。"
                        ),
                        recommendation=(
                            "シート名変数の値を確認し、参照先シートを手動確認してください。"
                        ),
                    )


def _detect_external_dependency_risks(wb: Workbook) -> Iterable[AnalysisRisk]:
    """外部リンク・Power Query / 接続定義のリスクを検出する."""
    for link in wb.external_links:
        yield AnalysisRisk(
            category="external_dependency",
            severity="medium",
            location="Workbook.external_links",
            evidence=link[:500],
            description="外部ブックリンクがあります。",
            recommendation=(
                "リンク先ファイルの存在、更新タイミング、改修対象との関係を確認してください。"
            ),
        )

    for query in wb.power_queries:
        target = query.target_sheet or query.target_name
        yield AnalysisRisk(
            category="external_dependency",
            severity="high" if query.kind == "power_query" else "medium",
            location=f"Connection:{query.name}",
            evidence=(query.source or query.command or query.description or query.name)[:500],
            description=(
                "Power Query / 外部接続があります。初期対応では "
                "M コード本文や接続先内部依存は解析しません。"
            ),
            recommendation=(
                f"出力先 `{target}` と接続の更新結果を Excel 上で確認してください。"
                if target
                else "接続の出力先と更新結果を Excel 上で確認してください。"
            ),
        )


def _detect_object_dependency_risks(wb: Workbook) -> Iterable[AnalysisRisk]:
    """グラフ・ピボットの依存先が取れないケースを検出する."""
    for sheet in wb.sheets:
        for chart in sheet.charts:
            if chart.series and all(s.values_ref or s.categories_ref for s in chart.series):
                continue
            yield AnalysisRisk(
                category="unknown_object_dependency",
                severity="medium",
                location=f"{sheet.name}!Chart:{chart.title or chart.name or chart.chart_type}",
                evidence=chart.model_dump_json()[:500],
                description="グラフは存在しますが、系列参照を一部または全部抽出できませんでした。",
                recommendation="Excel 上でグラフのデータソース範囲を確認してください。",
                confidence="unknown",
            )
        for pivot in sheet.pivot_tables:
            if pivot.source_ref or pivot.source_name:
                continue
            yield AnalysisRisk(
                category="unknown_object_dependency",
                severity="medium",
                location=f"{sheet.name}!Pivot:{pivot.name}",
                evidence=pivot.model_dump_json()[:500],
                description="ピボットテーブルは存在しますが、元データ範囲を抽出できませんでした。",
                recommendation="Excel 上でピボットテーブルの元データ範囲を確認してください。",
                confidence="unknown",
            )


def _strip_vba_comment(line: str) -> str:
    """VBA のコメント部分を除去する.

    Args:
        line: VBA ソース 1 行。

    Returns:
        コメント除去後の行。
    """
    in_string = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            if in_string and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            in_string = not in_string
        elif ch == "'" and not in_string:
            return line[:i]
        i += 1
    return line


def _mask_vba_strings(line: str) -> str:
    """文字列リテラルの中身を空文字相当に置き換える.

    Args:
        line: コメント除去済みの VBA ソース 1 行。

    Returns:
        文字列中のキーワードを誤検出しないためのマスク済み文字列。
    """
    out: list[str] = []
    in_string = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            out.append(ch)
            if in_string and i + 1 < len(line) and line[i + 1] == '"':
                out.append('"')
                i += 2
                continue
            in_string = not in_string
            i += 1
            continue
        out.append(" " if in_string else ch)
        i += 1
    return "".join(out)


def _is_single_string_literal(value: str) -> bool:
    """単一の VBA 文字列リテラルかどうかを返す.

    Args:
        value: 引数文字列。

    Returns:
        `"A1"` のような単一文字列なら True。`"A" & row` は False。
    """
    s = value.strip()
    if not (s.startswith('"') and s.endswith('"')):
        return False
    escaped = False
    for i, ch in enumerate(s[1:-1], start=1):
        if ch != '"':
            escaped = False
            continue
        if escaped:
            escaped = False
            continue
        if i + 1 < len(s) - 1 and s[i + 1] == '"':
            escaped = True
            continue
        return False
    return True


def _contains_token(text: str, token: str) -> bool:
    """VBA キーワード/メンバー名を大文字小文字無視で検出する."""
    pattern = rf"(?<![A-Za-z0-9_]){re.escape(token)}(?![A-Za-z0-9_])"
    return re.search(pattern, text, re.I) is not None


def _dedupe_risks(risks: Iterable[AnalysisRisk]) -> list[AnalysisRisk]:
    """同一リスクを重複除去する."""
    seen: set[tuple[str, str, str, str]] = set()
    out: list[AnalysisRisk] = []
    for risk in risks:
        key = (risk.category, risk.severity, risk.location, risk.evidence)
        if key in seen:
            continue
        seen.add(key)
        out.append(risk)
    return out
