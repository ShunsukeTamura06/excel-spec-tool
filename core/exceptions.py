"""Core層の例外クラス."""


class CoreError(Exception):
    """Core層の基底例外."""


class ExtractionError(CoreError):
    """Excel/VBA抽出時の失敗."""


class UnsupportedFormatError(CoreError):
    """対応していないファイル形式が渡された."""


class DiffError(CoreError):
    """workbook_diff の実行時失敗 (before/after ファイルの読み込み・cells.db 構築失敗)."""


class NamedRangeFixError(CoreError):
    """名前付き範囲の修正 (提案/適用) に関する失敗 (対象名が見つからない等)."""


class FormulaFixError(CoreError):
    """数式内参照の修正 (提案/適用) に関する失敗 (参照が不正・該当数式なし等)."""


class MutationProviderError(CoreError):
    """外部/内部の変更プロバイダーによる適用処理の失敗."""


class ProviderUnavailableError(MutationProviderError):
    """要求された変更プロバイダーが実行環境で利用できない."""


class UnsupportedMutationError(MutationProviderError):
    """変更プロバイダーが要求された形式または操作をサポートしない."""
