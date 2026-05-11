"""統合設計書 Markdown 生成モジュール.

Workbook と ReferenceIndex から、人が読める統合設計書 (Markdown) を生成する。
LLM 注釈は別関数 (annotate_with_llm) で付与し、本モジュールは LLM 非依存とする。

SPEC.md §4.4 参照。
"""

from __future__ import annotations

from collections.abc import Iterable

from core.models import (
    CellFormula,
    ReferenceIndex,
    Workbook,
)

_TOP_FORMULAS_PER_SHEET = 10
_TOP_REFERENCES_GLOBAL = 20


def _md_escape(s: str) -> str:
    """Markdown テーブル中で扱いづらい文字をエスケープする (パイプと改行)."""
    return s.replace("|", "\\|").replace("\n", " ")


def _section_overview(wb: Workbook) -> str:
    """## 1. 概要 を生成."""
    ext_links_str = ", ".join(wb.external_links) if wb.external_links else "なし"
    lines = [
        "## 1. 概要",
        "",
        f"- ファイル名: `{wb.filename}`",
        f"- シート数: {len(wb.sheets)}",
        f"- VBAモジュール数: {len(wb.vba_modules)}",
        f"- 外部リンク: {ext_links_str}",
    ]
    return "\n".join(lines)


def _section_sheet_list(wb: Workbook) -> str:
    """## 2. シート一覧 をテーブルで生成."""
    lines = [
        "## 2. シート一覧",
        "",
        "| シート名 | 行数 | 列数 | 数式数 | 名前付き範囲 | 条件付き書式 | 用途 |",
        "|---|---:|---:|---:|---:|---:|---|",
    ]
    if not wb.sheets:
        lines.append("| (なし) | - | - | - | - | - | - |")
    for s in wb.sheets:
        lines.append(
            "| {name} | {rows} | {cols} | {f} | {nr} | {cf} | {purpose} |".format(
                name=_md_escape(s.name),
                rows=s.rows,
                cols=s.cols,
                f=len(s.formulas),
                nr=len(s.named_ranges),
                cf=len(s.conditional_formats),
                purpose=_md_escape(s.purpose) if s.purpose else "-",
            )
        )
    return "\n".join(lines)


def _pick_top_formulas(formulas: Iterable[CellFormula], n: int) -> list[CellFormula]:
    """重要そうな数式を上位 n 件選ぶ.

    現状ヒューリスティック: 参照数が多い順 → coord 昇順.
    """
    return sorted(formulas, key=lambda f: (-len(f.refs), f.coord))[:n]


def _section_sheet_details(wb: Workbook) -> str:
    """## 3. シート詳細 (シートごとのサブセクション)."""
    lines = ["## 3. シート詳細"]
    if not wb.sheets:
        lines += ["", "_(シートなし)_"]
        return "\n".join(lines)

    for s in wb.sheets:
        lines += ["", f"### {s.name}"]
        if s.purpose:
            lines += ["", f"- 用途（推定）: {s.purpose}"]
        else:
            lines += ["", "- 用途（推定）: _未設定_"]

        # 主要数式
        top = _pick_top_formulas(s.formulas, _TOP_FORMULAS_PER_SHEET)
        lines += ["", f"#### 主要数式 (TOP {min(_TOP_FORMULAS_PER_SHEET, len(s.formulas))})"]
        if not top:
            lines += ["", "_(数式なし)_"]
        else:
            lines += [
                "",
                "| セル | 数式 | 参照先 | 注釈 |",
                "|---|---|---|---|",
            ]
            for f in top:
                refs_str = ", ".join(f.refs) if f.refs else "-"
                annot = f.annotation if f.annotation else ""
                lines.append(
                    "| `{c}` | `{fm}` | {r} | {a} |".format(
                        c=_md_escape(f.coord),
                        fm=_md_escape(f.formula),
                        r=_md_escape(refs_str),
                        a=_md_escape(annot) if annot else "-",
                    )
                )

        # 名前付き範囲
        lines += ["", "#### 名前付き範囲"]
        if not s.named_ranges:
            lines += ["", "_(なし)_"]
        else:
            lines += ["", "| 名前 | 参照先 |", "|---|---|"]
            for nr in s.named_ranges:
                lines.append(f"| {_md_escape(nr.name)} | `{_md_escape(nr.refers_to)}` |")

        # 条件付き書式
        lines += ["", "#### 条件付き書式"]
        if not s.conditional_formats:
            lines += ["", "_(なし)_"]
        else:
            lines += ["", "| 範囲 | ルール |", "|---|---|"]
            for cf in s.conditional_formats:
                lines.append(f"| `{_md_escape(cf.range)}` | {_md_escape(cf.rule)} |")

        # Excel テーブル (ListObject) — ユーザーが明示的に定義したテーブルのみ
        lines += ["", "#### Excel テーブル（明示的に定義されているもの）"]
        if not s.tables:
            lines += ["", "_(なし)_"]
        else:
            lines += ["", "| 名前 | 範囲 | ヘッダ行数 |", "|---|---|---:|"]
            for t in s.tables:
                lines.append(
                    f"| `{_md_escape(t.name)}` | `{_md_escape(t.ref)}` | {t.header_row_count} |"
                )

        # マージセル — 件数が多い場合は先頭のみ
        if s.merged_ranges:
            shown = s.merged_ranges[:10]
            extra = len(s.merged_ranges) - len(shown)
            lines += ["", "#### マージセル"]
            lines += ["", ", ".join(f"`{_md_escape(m)}`" for m in shown)]
            if extra > 0:
                lines.append(f"(他 {extra} 件)")

        # 冒頭プレビュー (literal, 解釈なし)
        lines += _render_preview_block(s)

    return "\n".join(lines)


def _render_preview_block(sheet: object) -> list[str]:
    """シート冒頭プレビューを Markdown テーブルとして描画する.

    解釈は一切加えない。生のセル値をそのまま並べるだけ。
    """
    rows = getattr(sheet, "preview_rows", []) or []
    origin = getattr(sheet, "preview_origin", "") or ""
    if not rows:
        return []

    from openpyxl.utils import get_column_letter

    n_cols = max(len(r) for r in rows)
    header = ["行"] + [get_column_letter(c + 1) for c in range(n_cols)]
    sep = ["---"] + ["---"] * n_cols

    out: list[str] = [
        "",
        f"#### 冒頭プレビュー (`{origin}`)",
        "",
        "_解釈は加えていません。シートの先頭領域に literal に並んでいる値です。_",
        "",
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(sep) + " |",
    ]
    for row_idx, row in enumerate(rows, start=1):
        cells = []
        for c in range(n_cols):
            v = row[c] if c < len(row) else None
            cells.append(_md_escape(v) if v is not None else "")
        out.append(f"| {row_idx} | " + " | ".join(cells) + " |")
    return out


def _section_vba_modules(wb: Workbook) -> str:
    """## 4. VBAモジュール."""
    lines = ["## 4. VBAモジュール"]
    if not wb.vba_modules:
        lines += ["", "_(VBAモジュールなし)_"]
        return "\n".join(lines)

    for m in wb.vba_modules:
        lines += ["", f"### {m.name} ({m.type})"]

        # プロシージャ一覧
        lines += ["", "#### プロシージャ一覧"]
        if not m.procedures:
            lines += ["", "_(なし)_"]
        else:
            lines += [
                "",
                "| 種別 | 名前 | 行 | 注釈 |",
                "|---|---|---|---|",
            ]
            for p in m.procedures:
                annot = _md_escape(p.annotation) if p.annotation else "-"
                lines.append(
                    f"| {p.kind} | `{_md_escape(p.name)}` | {p.start_line}-{p.end_line} | {annot} |"
                )

        # ソースコード (折りたたみ)
        lines += [
            "",
            "<details><summary>ソースコード</summary>",
            "",
            "```vba",
            m.code if m.code else "",
            "```",
            "",
            "</details>",
        ]
    return "\n".join(lines)


def _pick_top_references(ref_index: ReferenceIndex, n: int) -> list[tuple[str, int]]:
    """参照件数が多いキー上位 n 件を (key, count) で返す."""
    counts = [(k, len(v)) for k, v in ref_index.refs.items()]
    counts.sort(key=lambda kv: (-kv[1], kv[0]))
    return counts[:n]


def _section_references(ref_index: ReferenceIndex) -> str:
    """## 5. 参照関係（抜粋）."""
    lines = ["## 5. 参照関係（抜粋）"]
    top = _pick_top_references(ref_index, _TOP_REFERENCES_GLOBAL)
    if not top:
        lines += ["", "_(参照なし)_"]
        return "\n".join(lines)

    lines += [
        "",
        f"参照件数の多い上位 {len(top)} 件 (参照先 → 参照元の数):",
        "",
        "| 参照先 | 件数 | 参照元 (一部) |",
        "|---|---:|---|",
    ]
    for key, count in top:
        sources = ref_index.refs.get(key, [])
        # 先頭5件だけ表示
        from_list = ", ".join(_md_escape(r.from_) for r in sources[:5])
        if len(sources) > 5:
            from_list += f", … (+{len(sources) - 5})"
        lines.append(f"| `{_md_escape(key)}` | {count} | {from_list} |")
    return "\n".join(lines)


def _section_observations(wb: Workbook) -> str:
    """## 6. 注意点・観察事項.

    現状はプレースホルダ. LLM 注釈ステップで埋める想定。
    SheetInfo.purpose / VbaProcedure.annotation など既に埋まっている注釈の
    総数だけ簡易サマリで出す。
    """
    sheet_with_purpose = sum(1 for s in wb.sheets if s.purpose)
    proc_with_annot = sum(1 for m in wb.vba_modules for p in m.procedures if p.annotation)
    formulas_with_annot = sum(1 for s in wb.sheets for f in s.formulas if f.annotation)

    lines = [
        "## 6. 注意点・観察事項",
        "",
        "_LLM注釈ステップで追記される予定。_",
        "",
        f"- 用途が推定されたシート: {sheet_with_purpose} / {len(wb.sheets)}",
        f"- 注釈付きプロシージャ: {proc_with_annot}",
        f"- 注釈付き数式: {formulas_with_annot}",
    ]
    return "\n".join(lines)


def generate_spec(wb: Workbook, ref_index: ReferenceIndex) -> str:
    """Workbook と ReferenceIndex から統合設計書 (Markdown) を生成する.

    LLM 非依存. 注釈フィールド (SheetInfo.purpose, VbaProcedure.annotation,
    CellFormula.annotation) が埋まっていれば出力に反映される。

    Args:
        wb: 抽出済みワークブック.
        ref_index: 参照インデックス.

    Returns:
        Markdown 形式の設計書文字列 (末尾改行付き).
    """
    sections = [
        f"# 設計書: {wb.filename}",
        "",
        _section_overview(wb),
        "",
        _section_sheet_list(wb),
        "",
        _section_sheet_details(wb),
        "",
        _section_vba_modules(wb),
        "",
        _section_references(ref_index),
        "",
        _section_observations(wb),
        "",
    ]
    return "\n".join(sections)
