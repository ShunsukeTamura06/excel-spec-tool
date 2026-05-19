"""Excelツール改修支援AIのデータモデル定義.

SPEC.md §3 に基づく Pydantic モデル群。Core / Backend / Frontend で共有する。
"""

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


class VbaProcedure(BaseModel):
    """VBAプロシージャ (Sub / Function / Property) 1件."""

    name: str
    kind: Literal["Sub", "Function", "Property"]
    start_line: int
    end_line: int
    code: str
    annotation: str = ""


class VbaModule(BaseModel):
    """VBAモジュール 1件."""

    name: str
    type: Literal["Module", "Class", "Form", "Document"]
    code: str
    procedures: list[VbaProcedure] = Field(default_factory=list)


class CellFormula(BaseModel):
    """数式セル 1件."""

    coord: str
    formula: str
    refs: list[str] = Field(default_factory=list)
    annotation: str = ""


class NamedRange(BaseModel):
    """名前付き範囲."""

    name: str
    refers_to: str


class ConditionalFormat(BaseModel):
    """条件付き書式."""

    range: str
    rule: str


class ExcelTable(BaseModel):
    """Excel テーブル機能 (ListObject) で明示的に定義された表.

    `openpyxl.worksheet.table.Table` から取得する確定情報。
    ヒューリスティック検出ではない。
    """

    name: str
    ref: str
    header_row_count: int = 1


class SheetInfo(BaseModel):
    """シート 1枚の情報."""

    name: str
    rows: int
    cols: int
    formulas: list[CellFormula] = Field(default_factory=list)
    named_ranges: list[NamedRange] = Field(default_factory=list)
    conditional_formats: list[ConditionalFormat] = Field(default_factory=list)
    tables: list[ExcelTable] = Field(default_factory=list)
    merged_ranges: list[str] = Field(default_factory=list)
    # 先頭 N 行 × M 列の literal プレビュー (解釈なしの生値).
    # 各行は等長で、空セルは None.
    preview_rows: list[list[str | None]] = Field(default_factory=list)
    preview_origin: str = ""  # "A1" 等. 何処を起点に取ったかの記録
    purpose: str = ""


class Workbook(BaseModel):
    """ワークブック全体の抽出結果."""

    filename: str
    sheets: list[SheetInfo] = Field(default_factory=list)
    vba_modules: list[VbaModule] = Field(default_factory=list)
    external_links: list[str] = Field(default_factory=list)


class Reference(BaseModel):
    """参照1件 (数式 or VBA から、あるセル/範囲への参照)."""

    model_config = ConfigDict(populate_by_name=True)

    kind: Literal["formula", "vba"]
    from_: str = Field(alias="from")
    to: str
    code: str = ""


class ReferenceIndex(BaseModel):
    """逆引きインデックス: あるセル/範囲を参照しているもの一覧.

    キーは参照先 (例: "Calc!H2:H5000") で、値はそこを参照している箇所のリスト。
    """

    refs: dict[str, list[Reference]] = Field(default_factory=dict)


class JobMeta(BaseModel):
    """ジョブのメタ情報. ストレージの meta.json に対応."""

    job_id: str
    filename: str
    created_at: str
    status: Literal["uploaded", "extracted", "analyzed", "failed"]


class ChatMessage(BaseModel):
    """チャット履歴の1メッセージ. chat_history.jsonl の1行に対応.

    SPEC.md §3 にはこのモデルは明記されていないが、§5.4 のチャット仕様で
    必要なため追加する (実装の細部).
    """

    role: Literal["user", "assistant", "system"]
    content: str
    timestamp: str
