"""外部関数レジストリ.

`ExternalFunction` モデルと、関数名検索 / 数式中の検出ユーティリティを提供する。
ベンダーごとのコンテンツは `bloomberg.py` 等の別ファイルに置く。
"""

from __future__ import annotations

import re

from pydantic import BaseModel, Field


class ExternalFunctionParam(BaseModel):
    """外部関数の引数 1 件."""

    name: str
    description: str
    required: bool = True
    type: str = ""  # "security" / "field" / "date" / "string" / "number" / "array"


class ExternalFunction(BaseModel):
    """1 ベンダーの 1 関数の定義."""

    name: str  # "BDH" (大文字で統一)
    vendor: str  # "Bloomberg" / "Refinitiv" / ...
    short: str  # 1 文サマリ (一覧表示用)
    long: str  # 複数段落の詳しい説明
    signature: str  # "=BDH(security, fields, start_date, [end_date], [options])"
    params: list[ExternalFunctionParam] = Field(default_factory=list)
    returns: str = ""  # 何が返るか (型・形状)
    examples: list[str] = Field(default_factory=list)  # "=BDH(\"AAPL US Equity\", ...)"
    notes: list[str] = Field(default_factory=list)  # よくある落とし穴
    doc_url: str = ""  # 公式ドキュメントへのリンク (terminal 内など)


def _normalize_name(name: str) -> str:
    return name.strip().upper()


# ベンダーモジュールから関数を集める. 循環 import を避けるため遅延.
def _load_all() -> dict[str, ExternalFunction]:
    out: dict[str, ExternalFunction] = {}
    from core.external_functions import bloomberg

    for fn in bloomberg.FUNCTIONS:
        out[_normalize_name(fn.name)] = fn
    # 将来のベンダー追加はここに append
    return out


# 起動時に 1 度だけ構築. テストや変更時はモジュールを再 import すれば再構築される.
_REGISTRY: dict[str, ExternalFunction] | None = None


def _registry() -> dict[str, ExternalFunction]:
    global _REGISTRY
    if _REGISTRY is None:
        _REGISTRY = _load_all()
    return _REGISTRY


# ------------------------------ public API ------------------------------


def get_function(name: str) -> ExternalFunction | None:
    """関数名 (大小無視) から定義を取得. 未登録なら None."""
    if not name:
        return None
    return _registry().get(_normalize_name(name))


def list_functions(vendor: str | None = None) -> list[ExternalFunction]:
    """登録されている関数を返す. vendor 指定で絞り込み可."""
    funcs = list(_registry().values())
    if vendor:
        v = vendor.strip().lower()
        funcs = [f for f in funcs if f.vendor.lower() == v]
    funcs.sort(key=lambda f: (f.vendor, f.name))
    return funcs


# 「単語境界 + 大文字英字 1 文字以上 + ( + ...」を関数呼び出しとみなす.
# 厳密な構文解析より、レジストリにある関数名と突き合わせる方が安全.
_FUNC_CALL_RE = re.compile(r"(?<![A-Za-z0-9_])([A-Za-z][A-Za-z0-9_]*)\s*\(")


def detect_in_formula(formula: str) -> list[str]:
    """数式中に現れた外部関数名 (大文字正規化) を順序保ちでユニーク返却.

    `=BDH("AAPL US Equity", "PX_LAST") + BDP("MSFT US Equity", "NAME")` から
    `["BDH", "BDP"]` を返す。Excel 組込関数 (SUM 等) や VBA 識別子は無視。
    """
    if not formula:
        return []
    seen: list[str] = []
    known = _registry()
    for m in _FUNC_CALL_RE.finditer(formula):
        name = _normalize_name(m.group(1))
        if name in known and name not in seen:
            seen.append(name)
    return seen
