"""Core層の例外クラス."""


class CoreError(Exception):
    """Core層の基底例外."""


class ExtractionError(CoreError):
    """Excel/VBA抽出時の失敗."""


class UnsupportedFormatError(CoreError):
    """対応していないファイル形式が渡された."""
