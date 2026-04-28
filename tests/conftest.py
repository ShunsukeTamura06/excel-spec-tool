"""共通テストフィクスチャ."""

from pathlib import Path

import pytest

FIXTURES_DIR = Path(__file__).parent / "fixtures"


@pytest.fixture
def fixtures_dir() -> Path:
    """tests/fixtures/ への絶対パス."""
    return FIXTURES_DIR


@pytest.fixture
def empty_xlsx(tmp_path: Path) -> Path:
    """VBAを含まないxlsxを動的生成する.

    openpyxl で xlsx を作るだけ。olevba から見ると vba_macros は無し。
    """
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws["A1"] = "hello"
    out = tmp_path / "no_vba.xlsx"
    wb.save(out)
    return out
