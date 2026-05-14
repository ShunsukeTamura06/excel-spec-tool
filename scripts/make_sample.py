"""デモ用サンプル `.xlsx` を生成する.

3 シート構成 (Input / Calc / Output) で、シート間参照・名前付き範囲・
条件付き書式を含む小さな在庫管理ブックを作る。Excel 改修支援ツールの
シート依存グラフ・参照逆引きの動作確認に使う。

実行:
    uv run python scripts/make_sample.py

出力:
    frontend/public/samples/inventory_sample.xlsx
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

# fmt: off
PRODUCTS: list[tuple[str, str, int, int, int]] = [
    # (code,    name,           unit_price, stock, monthly_sales)
    ("P001",  "白米 5kg",          1980,  120,  85),
    ("P002",  "玄米 5kg",          2480,   60,  22),
    ("P003",  "もち米 2kg",        1280,   40,  15),
    ("P004",  "雑穀ミックス 1kg",  1580,   90,  45),
    ("P005",  "押し麦 800g",        780,  150,  60),
    ("P006",  "オートミール 500g",  680,  200, 120),
    ("P007",  "黒米 500g",         1180,   25,   8),
    ("P008",  "赤米 500g",         1280,   30,  12),
    ("P009",  "ハト麦 300g",        980,   45,  18),
    ("P010",  "古代米セット",      2880,   15,   5),
]
# fmt: on


HEADER_FILL = PatternFill("solid", fgColor="1E3A8A")  # indigo-900
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(size=14, bold=True, color="312E81")  # indigo-950


def _set_header(ws, row: int, values: list[str]) -> None:
    for col_idx, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=col_idx, value=v)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")


def _auto_width(ws, columns: list[int], width: float = 14) -> None:
    for col in columns:
        ws.column_dimensions[get_column_letter(col)].width = width


def build_input_sheet(wb: Workbook) -> None:
    ws = wb.active
    ws.title = "Input"

    ws["A1"] = "在庫マスタ"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:E1")

    _set_header(ws, 3, ["商品コード", "商品名", "単価", "在庫数", "月間販売数"])
    for i, (code, name, price, stock, sales) in enumerate(PRODUCTS, start=4):
        ws.cell(row=i, column=1, value=code)
        ws.cell(row=i, column=2, value=name)
        ws.cell(row=i, column=3, value=price).number_format = "#,##0"
        ws.cell(row=i, column=4, value=stock).number_format = "#,##0"
        ws.cell(row=i, column=5, value=sales).number_format = "#,##0"

    _auto_width(ws, [1, 2, 3, 4, 5], width=14)
    ws.column_dimensions["B"].width = 22


def build_calc_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("Calc")

    ws["A1"] = "派生計算"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:E1")

    _set_header(ws, 3, ["商品コード", "商品名", "在庫金額", "月間売上見込", "在庫回転月数"])

    # 行ごとに Input シートを参照する数式を入れる
    for i, _ in enumerate(PRODUCTS, start=4):
        input_row = i  # Input シートの該当行と Calc シートの該当行は同じ番号
        ws.cell(row=i, column=1, value=f"=Input!A{input_row}")
        ws.cell(row=i, column=2, value=f"=VLOOKUP(A{i},商品マスタ,2,FALSE)")
        # 在庫金額 = Input!C × Input!D
        ws.cell(row=i, column=3, value=f"=Input!C{input_row}*Input!D{input_row}")
        ws.cell(row=i, column=3).number_format = "#,##0"
        # 月間売上見込 = Input!C × Input!E
        ws.cell(row=i, column=4, value=f"=Input!C{input_row}*Input!E{input_row}")
        ws.cell(row=i, column=4).number_format = "#,##0"
        # 在庫回転月数 = Input!D ÷ Input!E (ゼロ除算ガード)
        ws.cell(
            row=i,
            column=5,
            value=f"=IF(Input!E{input_row}=0,0,Input!D{input_row}/Input!E{input_row})",
        )
        ws.cell(row=i, column=5).number_format = "0.0"

    # 条件付き書式: 在庫回転月数が 6 以上のセルを赤背景にする (在庫過剰)
    ws.conditional_formatting.add(
        f"E4:E{3 + len(PRODUCTS)}",
        CellIsRule(
            operator="greaterThanOrEqual",
            formula=["6"],
            fill=PatternFill("solid", fgColor="FCA5A5"),  # red-300
        ),
    )
    # 在庫金額のヒートマップ
    ws.conditional_formatting.add(
        f"C4:C{3 + len(PRODUCTS)}",
        ColorScaleRule(
            start_type="min",
            start_color="ECFCCB",  # lime-100
            mid_type="percentile",
            mid_value=50,
            mid_color="FDE68A",  # amber-200
            end_type="max",
            end_color="FCA5A5",  # red-300
        ),
    )

    _auto_width(ws, [1, 2, 3, 4, 5], width=16)
    ws.column_dimensions["B"].width = 22


def build_output_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("Output")

    ws["A1"] = "経営サマリ"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:B1")

    _set_header(ws, 3, ["指標", "値"])

    rows = [
        ("商品数",             "=COUNTA(Input!A4:A100)"),
        ("総在庫金額",         "=SUM(Calc!C4:C100)"),
        ("月間売上見込合計",   "=SUM(Calc!D4:D100)"),
        ("平均単価",           "=AVERAGE(Input!C4:C100)"),
        ("在庫過剰商品数",     "=COUNTIF(Calc!E4:E100,\">=6\")"),
        ("売れ筋商品数 (月販 50 以上)", "=COUNTIF(Input!E4:E100,\">=50\")"),
        ("最大在庫金額",       "=MAX(Calc!C4:C100)"),
        ("最小回転月数",       "=MIN(Calc!E4:E100)"),
    ]
    for i, (label, formula) in enumerate(rows, start=4):
        ws.cell(row=i, column=1, value=label)
        c = ws.cell(row=i, column=2, value=formula)
        if "AVERAGE" in formula or "回転" in label:
            c.number_format = "0.0"
        elif "金額" in label or "売上" in label or "単価" in label:
            c.number_format = "#,##0"
        else:
            c.number_format = "#,##0"

    _auto_width(ws, [1, 2], width=22)
    ws.column_dimensions["A"].width = 28


def add_named_range(wb: Workbook) -> None:
    """商品マスタ = Input!$A$4:$E$13 を名前付き範囲として登録."""
    last_row = 3 + len(PRODUCTS)
    wb.defined_names["商品マスタ"] = DefinedName(
        name="商品マスタ",
        attr_text=f"Input!$A$4:$E${last_row}",
    )


def main() -> Path:
    wb = Workbook()
    build_input_sheet(wb)
    build_calc_sheet(wb)
    build_output_sheet(wb)
    add_named_range(wb)

    out = Path(__file__).resolve().parents[1] / "frontend" / "public" / "samples" / "inventory_sample.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


if __name__ == "__main__":
    p = main()
    print(f"wrote: {p} ({p.stat().st_size:,} bytes)")
