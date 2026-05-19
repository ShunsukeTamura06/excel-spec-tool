"""統合設計書 Markdown 生成モジュール.

Workbook と ReferenceIndex から、人が読める統合設計書 (Markdown) を生成する。
LLM 注釈は別関数 (annotate_with_llm) で付与し、本モジュールは LLM 非依存とする。

SPEC.md §4.4 参照。
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Iterable

from core.diagrams import Diagram, build_diagrams
from core.external_functions import get_function, list_functions
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

        # 構造化注釈 (LLM が JSON で返した詳細). 各フィールドは任意.
        if s.usage_scenario:
            lines += [f"- 想定利用シーン: {s.usage_scenario}"]
        if s.inputs:
            lines += [f"- 依存元 (IN): {', '.join(s.inputs)}"]
        if s.outputs:
            lines += [f"- 出力先 (OUT): {', '.join(s.outputs)}"]
        if s.main_calculations:
            lines += ["", "#### 主要計算 (LLM 説明)"]
            for c in s.main_calculations:
                lines.append(f"- {c}")

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

        # データ検証 (入力規則)
        if s.data_validations:
            lines += ["", "#### 入力規則"]
            lines += [
                "",
                "| 範囲 | 種別 | 値 / 数式 | プロンプト |",
                "|---|---|---|---|",
            ]
            for dv in s.data_validations:
                op = f" {dv.operator}" if dv.operator else ""
                value = f"{_md_escape(dv.formula)}{_md_escape(op)}".strip() if dv.formula else "-"
                lines.append(
                    f"| `{_md_escape(dv.range)}` | {_md_escape(dv.type)} | "
                    f"{value} | {_md_escape(dv.prompt) if dv.prompt else '-'} |"
                )

        # フォームコントロール (ボタン → マクロ)
        if s.form_controls:
            lines += ["", "#### フォームコントロール（ボタン等）"]
            lines += [
                "",
                "| 種別 | 表示テキスト | 配置 | 紐づけマクロ |",
                "|---|---|---|---|",
            ]
            for fc in s.form_controls:
                text = _md_escape(fc.text) if fc.text else "-"
                anchor = f"`{_md_escape(fc.anchor)}`" if fc.anchor else "-"
                macro = f"`{_md_escape(fc.macro)}`" if fc.macro else "-"
                lines.append(f"| {_md_escape(fc.kind)} | {text} | {anchor} | {macro} |")

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
    """## 4. VBAモジュール.

    プロシージャの目録のみを載せ、ソースコード本体は載せない。
    本体が必要な場合は LLM 側の `get_vba_procedure(module, name)` ツールを使う
    (設計書のサイズ削減 = Phase B-1)。
    """
    lines = ["## 4. VBAモジュール"]
    if not wb.vba_modules:
        lines += ["", "_(VBAモジュールなし)_"]
        return "\n".join(lines)

    total_lines = 0
    for m in wb.vba_modules:
        module_line_count = len(m.code.splitlines()) if m.code else 0
        total_lines += module_line_count
        lines += [
            "",
            f"### {m.name} ({m.type})",
            "",
            f"- 行数: {module_line_count}",
            f"- プロシージャ数: {len(m.procedures)}",
        ]

        lines += ["", "#### プロシージャ一覧"]
        if not m.procedures:
            lines += ["", "_(なし)_"]
        else:
            lines += [
                "",
                "| 種別 | 名前 | 行 | 注釈 | 副作用 | 起動契機 | 呼出先 |",
                "|---|---|---|---|---|---|---|",
            ]
            for p in m.procedures:
                annot = _md_escape(p.annotation) if p.annotation else "-"
                side = _md_escape(", ".join(p.side_effects)) if p.side_effects else "-"
                trig = _md_escape(", ".join(p.triggers)) if p.triggers else "-"
                calls = _md_escape(", ".join(p.calls)) if p.calls else "-"
                lines.append(
                    f"| {p.kind} | `{_md_escape(p.name)}` | "
                    f"{p.start_line}-{p.end_line} | {annot} | "
                    f"{side} | {trig} | {calls} |"
                )

    lines += [
        "",
        "_ソースコード本体は設計書には含まれません。"
        "LLM が `get_vba_procedure(module, name)` ツールで取得します。_",
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


def _mermaid_label(s: str) -> str:
    """Mermaid のノードラベル ["..."] 内で安全な文字列にする.

    ダブルクォート・バックスラッシュ・改行・パイプを除去/置換する。
    日本語などはそのまま通る。
    """
    return (
        s.replace("\\", "\\\\")
        .replace('"', "'")
        .replace("\n", " ")
        .replace("|", "/")
    )


def _mermaid_id(prefix: str, index: int) -> str:
    """順序由来の安定 ID. ノード ID にシート名等を直接使うと記号で壊れる."""
    return f"{prefix}{index}"


_MERMAID_FENCE_OPEN = "```mermaid"
_MERMAID_FENCE_CLOSE = "```"


def _render_mermaid_graph(diagram: Diagram, direction: str = "LR") -> list[str]:
    """Diagram を mermaid `graph` 構文の行リストにして返す.

    ノード ID は順序由来 (n0, n1, ...) で安定化させ、ラベルにオリジナル名を載せる。
    エッジに weight があれば `|×N|` で併記。
    """
    lines: list[str] = [f"graph {direction}"]
    # node id 解決
    id_map: dict[str, str] = {}
    for i, n in enumerate(diagram.nodes):
        nid = _mermaid_id("n", i)
        id_map[n.id] = nid
        label = _mermaid_label(n.label or n.id)
        lines.append(f'    {nid}["{label}"]')

    for e in diagram.edges:
        src = id_map.get(e.src)
        dst = id_map.get(e.dst)
        if not src or not dst:
            continue
        if e.weight >= 2:
            lines.append(f"    {src} -->|×{e.weight}| {dst}")
        else:
            lines.append(f"    {src} --> {dst}")
    return lines


def _count_external_functions(wb: Workbook) -> Counter[str]:
    """Workbook 全体で外部関数の使用回数を集計する."""
    c: Counter[str] = Counter()
    for sheet in wb.sheets:
        for f in sheet.formulas:
            for name in f.external_functions:
                c[name] += 1
    return c


def _external_function_top_locations(
    wb: Workbook, name: str, limit: int = 5
) -> list[str]:
    """指定外部関数の使用箇所セル座標を先頭 limit 件返す."""
    out: list[str] = []
    for sheet in wb.sheets:
        for f in sheet.formulas:
            if name in f.external_functions:
                out.append(f"{sheet.name}!{f.coord.split('!', 1)[-1]}"
                           if "!" not in f.coord else f.coord)
                if len(out) >= limit:
                    return out
    return out


def _section_external_functions(wb: Workbook) -> str:
    """## N. 外部関数 (Bloomberg 等) — 検出された外部 Add-In 関数の一覧と定義."""
    lines: list[str] = ["## 6. 外部関数 (Bloomberg / Refinitiv 等)"]
    counts = _count_external_functions(wb)
    if not counts:
        lines += [
            "",
            "_検出された外部 Add-In 関数はありません。_",
            "",
            "(対応ベンダー: " + ", ".join(
                sorted({f.vendor for f in list_functions()})
            ) + ")",
        ]
        return "\n".join(lines)

    # サマリ
    lines += [
        "",
        f"検出された外部関数: **{len(counts)} 種類 / 使用箇所 {sum(counts.values())} 件**",
        "",
        "| 関数 | ベンダー | 使用回数 | 主要箇所 (TOP 5) | 概要 |",
        "|---|---|---:|---|---|",
    ]
    for name, cnt in counts.most_common():
        fn = get_function(name)
        vendor = fn.vendor if fn else "?"
        short = _md_escape(fn.short) if fn else "_(レジストリ未登録)_"
        locs = _external_function_top_locations(wb, name, limit=5)
        locs_str = ", ".join(f"`{_md_escape(c)}`" for c in locs) if locs else "-"
        lines.append(f"| `{name}` | {vendor} | {cnt} | {locs_str} | {short} |")

    # 各関数の詳細定義
    lines += ["", "### 関数定義"]
    for name, _ in counts.most_common():
        fn = get_function(name)
        if fn is None:
            continue
        lines += ["", f"#### `{fn.name}` ({fn.vendor})"]
        lines += ["", fn.long]
        lines += ["", f"**シグネチャ**: `{fn.signature}`"]
        if fn.params:
            lines += ["", "**引数**:", ""]
            for p in fn.params:
                req = "必須" if p.required else "任意"
                type_str = f" ({p.type})" if p.type else ""
                lines.append(f"- `{p.name}`{type_str} — {req}: {p.description}")
        if fn.returns:
            lines += ["", f"**返り値**: {fn.returns}"]
        if fn.examples:
            lines += ["", "**使用例**:", "", "```excel"]
            lines += fn.examples
            lines += ["```"]
        if fn.notes:
            lines += ["", "**注意点**:", ""]
            for note in fn.notes:
                lines.append(f"- {note}")
        if fn.doc_url:
            lines += ["", f"**公式参考**: {fn.doc_url}"]

    return "\n".join(lines)


def _section_diagrams(wb: Workbook) -> str:
    """## 7. 依存グラフ — シート依存 + VBA コールを mermaid で埋め込む.

    GitHub / VSCode / 多くの Markdown ビューアが mermaid を直接描画するため、
    ダウンロード後でも依存関係を視覚的に追える。Web UI には別途インタラクティブな
    Vue Flow タブがある。
    """
    diagrams = build_diagrams(wb)
    lines: list[str] = ["## 7. 依存グラフ"]
    lines += [
        "",
        "_GitHub / VSCode 等の mermaid 対応ビューアで描画されます._"
        " インタラクティブ表示は Web UI の「ダイアグラム」タブを参照してください。",
    ]

    # シート依存
    lines += ["", "### 7.1 シート依存"]
    sd = diagrams.sheet_deps
    if not sd.nodes:
        lines += ["", "_(シートなし)_"]
    elif not sd.edges:
        lines += ["", "_(シート間の参照なし)_"]
    else:
        lines += [
            "",
            f"ノード {len(sd.nodes)} 件 / エッジ {len(sd.edges)} 本",
            "",
            _MERMAID_FENCE_OPEN,
        ]
        lines += _render_mermaid_graph(sd, direction="LR")
        lines += [_MERMAID_FENCE_CLOSE]

    # VBA コール
    lines += ["", "### 7.2 VBA コール"]
    vc = diagrams.vba_calls
    if not vc.nodes:
        lines += ["", "_(VBA プロシージャなし)_"]
    elif not vc.edges:
        lines += ["", "_(プロシージャ間の呼び出しなし)_"]
    else:
        lines += [
            "",
            f"ノード {len(vc.nodes)} 件 / エッジ {len(vc.edges)} 本",
            "",
            _MERMAID_FENCE_OPEN,
        ]
        lines += _render_mermaid_graph(vc, direction="LR")
        lines += [_MERMAID_FENCE_CLOSE]

    return "\n".join(lines)


def _section_observations(wb: Workbook) -> str:
    """## 8. 注意点・観察事項.

    現状はプレースホルダ. LLM 注釈ステップで埋める想定。
    SheetInfo.purpose / VbaProcedure.annotation など既に埋まっている注釈の
    総数だけ簡易サマリで出す。
    """
    sheet_with_purpose = sum(1 for s in wb.sheets if s.purpose)
    proc_with_annot = sum(1 for m in wb.vba_modules for p in m.procedures if p.annotation)
    formulas_with_annot = sum(1 for s in wb.sheets for f in s.formulas if f.annotation)

    lines = [
        "## 8. 注意点・観察事項",
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
        _section_external_functions(wb),
        "",
        _section_diagrams(wb),
        "",
        _section_observations(wb),
        "",
    ]
    return "\n".join(sections)
