"""外部 Add-In 関数 (Bloomberg / Refinitiv 等) の知識ベース.

業務 Excel では Bloomberg Excel Add-In の `=BDH(...)` `=BDP(...)` のような
非標準関数が頻出する。これらは Excel の組込関数ではないため、LLM が知らない
場合 / 知っていても引数仕様を hallucinate する場合がある.

本パッケージは:
  - 関数定義を `ExternalFunction` レジストリとして保持
  - 数式中の外部関数呼び出しを検出 (`detect_in_formula`)
  - LLM ツールから引ける (`backend/llm_tools.py`)
  - 設計書とフロントエンドの「外部関数」タブで表示

ベンダーを追加するときは `core/external_functions/<vendor>.py` を作って
`registry._VENDORS` に登録する。レジストリ自体はベンダー非依存に作る。

License note:
  関数シグネチャ・引数名は事実 (公開情報) であり著作物性はない. 説明文は
  すべて自作の日本語要約で、ベンダー公式ドキュメントの翻訳ではない.
"""

from core.external_functions.registry import (
    ExternalFunction,
    ExternalFunctionParam,
    detect_in_formula,
    get_function,
    list_functions,
)

__all__ = [
    "ExternalFunction",
    "ExternalFunctionParam",
    "detect_in_formula",
    "get_function",
    "list_functions",
]
