"""図解 (シート依存グラフ・VBA コールグラフ) のビルダー.

Workbook (数式 + VBA) から、フロントエンドが Vue Flow 等で描画するための
ノード/エッジ構造を構築する。完璧なパースは目指さず、参照インデックスと
同様に「最低限の主要パターンを捕捉する」方針 (docs/SPEC.ja.md §4.3 参照)。
"""

from __future__ import annotations

import logging
import re
from collections import Counter
from typing import Literal

from pydantic import BaseModel, Field

from core.models import VbaModule, Workbook
from core.reference_index import _parse_ref

logger = logging.getLogger(__name__)


NodeKind = Literal["sheet", "module", "procedure"]
EdgeKind = Literal["formula", "call"]


class DiagramNode(BaseModel):
    """グラフのノード 1 件.

    `id` はグラフ内でユニーク。`label` は描画ラベル。`meta` には件数等の
    補足情報を載せる (例: シートなら formulas/rows、プロシージャなら kind)。
    """

    id: str
    label: str
    kind: NodeKind
    meta: dict[str, str | int] = Field(default_factory=dict)


class DiagramEdge(BaseModel):
    """有向エッジ 1 件. `weight` は集約後の出現回数."""

    src: str
    dst: str
    weight: int = 1
    kind: EdgeKind


class Diagram(BaseModel):
    """単一のグラフ構造."""

    kind: Literal["sheet_deps", "vba_calls"]
    nodes: list[DiagramNode] = Field(default_factory=list)
    edges: list[DiagramEdge] = Field(default_factory=list)


class DiagramSet(BaseModel):
    """1 つのジョブから生成される全グラフ.

    現状は 2 種類. 将来 'vba_to_sheet' 等を追加しても backward compatible。
    """

    sheet_deps: Diagram
    vba_calls: Diagram


# ----- シート依存グラフ ------------------------------------------------------


def build_sheet_dependency_graph(wb: Workbook) -> Diagram:
    """シート間の参照依存をグラフ化する.

    数式の `refs` を走査し、`source_sheet -> target_sheet` のエッジを作る。
    同一シート内の参照 (自己ループ) は除外する。エッジは weight (参照件数)
    で集約される。すべてのシートはノードとして必ず登録する (参照ゼロでも)。

    Args:
        wb: Workbook モデル.

    Returns:
        Diagram(kind="sheet_deps").
    """
    pair_counts: Counter[tuple[str, str]] = Counter()
    sheet_names = {s.name for s in wb.sheets}

    for sheet in wb.sheets:
        src = sheet.name
        for f in sheet.formulas:
            for raw in f.refs:
                parsed = _parse_ref(raw, owner_sheet=src)
                if parsed is None or parsed.sheet is None:
                    continue
                dst = parsed.sheet
                if dst == src:
                    continue  # 自己ループ除外
                if dst not in sheet_names:
                    # ブック内に存在しないシート (外部参照や打ち間違い) はスキップ
                    continue
                pair_counts[(src, dst)] += 1

    # 各シートが「最も多く参照しているシート」を求める. ノード本体の表示に使う.
    top_target_by_src: dict[str, tuple[str, int]] = {}
    for (src, dst), count in pair_counts.most_common():
        if src not in top_target_by_src:
            top_target_by_src[src] = (dst, count)

    # 全シートをノードとして登録 (孤立シートも描画したい)
    nodes: list[DiagramNode] = []
    for s in wb.sheets:
        meta: dict[str, str | int] = {
            "formulas": len(s.formulas),
            "rows": s.rows,
            "cols": s.cols,
            "purpose": s.purpose,
        }
        # 代表数式: 最初の数式を 1 件だけ載せる. ノード本体で「このシートは何を
        # しているのか」をひと目で推測できるようにする目的なので、種類ではなく
        # 「具体的な例」を 1 つ出す方が分かりやすい.
        if s.formulas:
            sample = s.formulas[0]
            meta["sample_formula"] = sample.formula[:60]
            meta["sample_coord"] = sample.coord
        # 最頻参照先シート (このシートが最も多く読んでいる相手)
        if s.name in top_target_by_src:
            dst, count = top_target_by_src[s.name]
            meta["top_target"] = dst
            meta["top_target_count"] = count
        nodes.append(DiagramNode(id=s.name, label=s.name, kind="sheet", meta=meta))

    edges = [
        DiagramEdge(src=s, dst=d, weight=w, kind="formula")
        for (s, d), w in pair_counts.most_common()
    ]
    return Diagram(kind="sheet_deps", nodes=nodes, edges=edges)


# ----- VBA コールグラフ ------------------------------------------------------

# プロシージャ名として有効な VBA 識別子. ASCII 英数 + アンダースコア.
# (日本語識別子は VBA でも使えるが珍しいので対象外)
_VBA_IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")


def _strip_strings_and_comments(line: str) -> str:
    """文字列リテラルとコメントを除いた行を返す.

    `Call Foo "literal"` → `Call Foo ""` のような単純置換。VBA の '`' (バッククォート)
    コメントは存在しないので `'` 以降をカットすれば十分。
    """
    # コメント除去
    line = line.split("'", 1)[0]
    # ダブルクォート文字列を空に. VBA はエスケープが `""` なので雑に置換でも害は少ない
    line = re.sub(r'"[^"\n]*"', '""', line)
    return line


def _proc_decl_pattern(name: str) -> re.Pattern[str]:
    """プロシージャ宣言行 (`Sub Foo`, `Function Foo`, `Property Get Foo` 等)
    を識別するための正規表現. VBA はキーワード大文字小文字を区別しない."""
    return re.compile(
        rf"^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+{re.escape(name)}\b",
        re.IGNORECASE,
    )


def build_vba_call_graph(modules: list[VbaModule]) -> Diagram:
    """VBA プロシージャ間のコールグラフを構築する.

    アルゴリズム:
      1. 全プロシージャを `Module.Proc` 形式のノードとして登録.
      2. 各プロシージャのコード片を走査し、識別子トークンを抽出.
         自分自身でなく、かつ既知のプロシージャ名と一致すれば呼び出しとみなす。
      3. 同名プロシージャが複数モジュールに存在する場合、同一モジュール内を
         優先しつつ、見つからなければ他モジュールの同名を採用する。
      4. 重複はカウントして weight に集約.

    完璧な検出は目指さない (動的バインディング・evaluate・実行時生成は除外)。

    Args:
        modules: VBA モジュール一覧.

    Returns:
        Diagram(kind="vba_calls").
    """
    # ノード構築: name → list[(module_name, kind)]
    # 1 プロシージャ = 1 ノード. id は "Module.Proc" でユニーク化.
    nodes: list[DiagramNode] = []
    # name (lower) → list[node_id]
    name_to_ids: dict[str, list[str]] = {}
    # node_id → module_name
    id_to_module: dict[str, str] = {}

    for module in modules:
        for proc in module.procedures:
            node_id = f"{module.name}.{proc.name}"
            nodes.append(
                DiagramNode(
                    id=node_id,
                    label=proc.name,
                    kind="procedure",
                    meta={
                        "module": module.name,
                        "kind": proc.kind,
                        "start_line": proc.start_line,
                        "end_line": proc.end_line,
                    },
                )
            )
            name_to_ids.setdefault(proc.name.lower(), []).append(node_id)
            id_to_module[node_id] = module.name

    # VBA キーワード等の誤検出回避リスト
    keywords = {
        "if",
        "then",
        "else",
        "elseif",
        "end",
        "for",
        "next",
        "do",
        "while",
        "wend",
        "loop",
        "until",
        "select",
        "case",
        "with",
        "exit",
        "function",
        "sub",
        "property",
        "get",
        "let",
        "set",
        "dim",
        "redim",
        "const",
        "as",
        "byval",
        "byref",
        "optional",
        "paramarray",
        "is",
        "not",
        "and",
        "or",
        "xor",
        "eqv",
        "imp",
        "mod",
        "true",
        "false",
        "nothing",
        "null",
        "empty",
        "me",
        "new",
        "call",
        "type",
        "enum",
        "in",
        "on",
        "error",
        "resume",
        "goto",
        "return",
        "stop",
        "private",
        "public",
        "friend",
        "static",
        "rem",
        "to",
        "step",
        "each",
        "single",
        "double",
        "integer",
        "long",
        "string",
        "boolean",
        "variant",
        "object",
        "date",
        "currency",
        "byte",
        "decimal",
        "lbound",
        "ubound",
    }

    pair_counts: Counter[tuple[str, str]] = Counter()

    for module in modules:
        for proc in module.procedures:
            src_id = f"{module.name}.{proc.name}"
            for raw_line in proc.code.splitlines():
                # プロシージャ宣言行は呼び出しと見なさない
                if _proc_decl_pattern(proc.name).match(raw_line):
                    continue
                line = _strip_strings_and_comments(raw_line)
                for tok_match in _VBA_IDENT_RE.finditer(line):
                    tok = tok_match.group(0)
                    tok_lower = tok.lower()
                    if tok_lower in keywords:
                        continue
                    if tok_lower not in name_to_ids:
                        continue
                    if tok_lower == proc.name.lower():
                        # 直前が `End ` や `Sub ` 等の宣言/終了でないことは
                        # _proc_decl_pattern で除外済。それ以外の自己呼び出しは許容。
                        # ただし宣言行で来た場合に限り除外したいので continue にはしない。
                        pass
                    # 候補解決: 同一モジュール内優先, なければ name_to_ids の先頭
                    candidates = name_to_ids[tok_lower]
                    same_mod = [c for c in candidates if id_to_module[c] == module.name]
                    dst_id = same_mod[0] if same_mod else candidates[0]
                    if dst_id == src_id:
                        # 再帰呼び出しは描画しない
                        continue
                    pair_counts[(src_id, dst_id)] += 1

    edges = [
        DiagramEdge(src=s, dst=d, weight=w, kind="call") for (s, d), w in pair_counts.most_common()
    ]
    return Diagram(kind="vba_calls", nodes=nodes, edges=edges)


# ----- まとめ ----------------------------------------------------------------


def build_diagrams(wb: Workbook) -> DiagramSet:
    """Workbook から両グラフをまとめて構築する."""
    return DiagramSet(
        sheet_deps=build_sheet_dependency_graph(wb),
        vba_calls=build_vba_call_graph(wb.vba_modules),
    )
