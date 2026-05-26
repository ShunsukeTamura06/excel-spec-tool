"""Excelツール改修支援AIのデータモデル定義.

SPEC.md §3 に基づく Pydantic モデル群。Core / Backend / Frontend で共有する。
"""

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field


class VbaProcedure(BaseModel):
    """VBAプロシージャ (Sub / Function / Property) 1件."""

    name: str
    kind: Literal["Sub", "Function", "Property"]
    start_line: int
    end_line: int
    code: str

    # ----- LLM 注釈フィールド (P1-2 で構造化) -----
    annotation: str = ""  # 1 文サマリ
    side_effects: list[str] = Field(default_factory=list)  # 書き込み先 (セル / シート / 外部)
    triggers: list[str] = Field(default_factory=list)  # 想定起動契機 (ボタン / イベント / 手動)
    calls: list[str] = Field(default_factory=list)  # 内部で呼ぶ他プロシージャ名 (LLM が判定)


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
    # 外部 Add-In 関数 (Bloomberg BDH/BDP/BDS, Refinitiv 等) の利用検出.
    # `core.external_functions.detect_in_formula` の結果. 標準関数 (SUM 等) は含まない.
    external_functions: list[str] = Field(default_factory=list)


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


class ChartSeries(BaseModel):
    """グラフ系列 1件.

    `values_ref` / `categories_ref` は OOXML の `c:f` から直接取れたセル参照。
    取得できない場合は空文字にする。
    """

    name: str = ""
    values_ref: str = ""
    categories_ref: str = ""


class ChartObject(BaseModel):
    """シート上の Excel グラフ 1件.

    OOXML の DrawingML / chart パーツから best-effort で抽出する。
    """

    name: str = ""
    chart_type: str = ""
    title: str = ""
    anchor: str = ""
    series: list[ChartSeries] = Field(default_factory=list)


class PivotTableInfo(BaseModel):
    """ピボットテーブル 1件.

    `source_*` は pivotCacheDefinition から直接取れた元データ情報。
    取れない場合は空文字のままにする。
    """

    name: str
    anchor: str = ""
    cache_id: str = ""
    source_type: str = ""
    source_sheet: str = ""
    source_ref: str = ""
    source_name: str = ""
    row_fields: list[str] = Field(default_factory=list)
    column_fields: list[str] = Field(default_factory=list)
    value_fields: list[str] = Field(default_factory=list)
    filter_fields: list[str] = Field(default_factory=list)


class PowerQueryInfo(BaseModel):
    """Power Query / 外部データ接続の棚卸し情報.

    Power Query は保存形式の揺れが大きいため、初期対応では接続定義と
    出力先 queryTable から明示的に取れた情報だけを持つ。
    """

    name: str
    kind: Literal["power_query", "connection"] = "connection"
    connection_id: str = ""
    connection_type: str = ""
    description: str = ""
    refresh_on_load: bool = False
    target_sheet: str = ""
    target_name: str = ""
    source: str = ""
    command: str = ""
    m_code: str = ""
    confidence: Literal["explicit", "inferred", "unknown"] = "explicit"


class AnalysisRisk(BaseModel):
    """静的解析では断定できない未解析リスク 1件.

    LLM が「影響なし」と誤断定しないためのガードレールとして使う。
    """

    category: Literal[
        "dynamic_vba",
        "runtime_state",
        "dynamic_formula",
        "external_dependency",
        "event_macro",
        "unknown_object_dependency",
    ]
    severity: Literal["high", "medium", "low"]
    location: str
    evidence: str
    description: str
    recommendation: str
    confidence: Literal["explicit", "inferred", "unknown"] = "explicit"


class DataValidation(BaseModel):
    """セルに設定された入力規則 (リスト / 数値範囲 / 日付 等).

    主用途は「このセルにはどんな値が入るか」をチャット LLM が即答できるようにすること.
    """

    range: str  # 適用範囲 (sqref). 例: "A2:A100", "C5"
    type: str  # "list" / "whole" / "decimal" / "date" / "time" / "textLength" / "custom"
    formula: str = ""  # formula1. リストなら値そのもの / セル参照
    operator: str = ""  # "between" / "greaterThan" 等. リストでは空
    prompt: str = ""  # 入力時のヒントメッセージ
    error: str = ""  # 入力エラー時のメッセージ
    allow_blank: bool = True


class FormControl(BaseModel):
    """シート上のフォームコントロール (ボタン / チェックボックス 等) と VBA の紐付け.

    .xlsm の VML ドローイング (`xl/drawings/vmlDrawing*.vml`) を best-effort で
    パースして抽出する。OnAction で指定された VBA マクロ名と、可能なら表示テキストや
    アンカーセルも記録する。
    """

    kind: str = "button"  # "button" / "checkbox" / "dropdown" / "scrollbar" 等
    name: str = ""  # コントロール名 (取れる場合)
    text: str = ""  # ボタン表面の表示テキスト (取れる場合)
    macro: str = ""  # 紐づけマクロ名 (FmlaMacro)
    anchor: str = ""  # 配置セル (FromRow,FromCol) を "A1" 形式に変換したもの


class SheetInfo(BaseModel):
    """シート 1枚の情報."""

    name: str
    rows: int
    cols: int
    formulas: list[CellFormula] = Field(default_factory=list)
    named_ranges: list[NamedRange] = Field(default_factory=list)
    conditional_formats: list[ConditionalFormat] = Field(default_factory=list)
    tables: list[ExcelTable] = Field(default_factory=list)
    charts: list[ChartObject] = Field(default_factory=list)
    pivot_tables: list[PivotTableInfo] = Field(default_factory=list)
    merged_ranges: list[str] = Field(default_factory=list)
    data_validations: list[DataValidation] = Field(default_factory=list)
    form_controls: list[FormControl] = Field(default_factory=list)
    # 先頭 N 行 × M 列の literal プレビュー (解釈なしの生値).
    # 各行は等長で、空セルは None.
    preview_rows: list[list[str | None]] = Field(default_factory=list)
    preview_origin: str = ""  # "A1" 等. 何処を起点に取ったかの記録

    # ----- LLM 注釈フィールド (P1-2 で構造化) -----
    # 後方互換: いずれも default は空. 既存ストレージから読み込めるよう default を維持する.
    purpose: str = ""  # 用途を 1〜2 文で
    inputs: list[str] = Field(default_factory=list)  # 依存元 (他シート名 / 外部)
    outputs: list[str] = Field(default_factory=list)  # 出力先
    main_calculations: list[str] = Field(default_factory=list)  # 主要計算の自然言語説明
    usage_scenario: str = ""  # 想定利用シーン


class Workbook(BaseModel):
    """ワークブック全体の抽出結果."""

    filename: str
    sheets: list[SheetInfo] = Field(default_factory=list)
    vba_modules: list[VbaModule] = Field(default_factory=list)
    external_links: list[str] = Field(default_factory=list)
    power_queries: list[PowerQueryInfo] = Field(default_factory=list)
    analysis_risks: list[AnalysisRisk] = Field(default_factory=list)


class Reference(BaseModel):
    """参照1件 (数式 or VBA から、あるセル/範囲への参照)."""

    model_config = ConfigDict(populate_by_name=True)

    kind: Literal["formula", "vba", "chart", "pivot", "power_query"]
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
    file_sha256: str | None = None
    file_size: int | None = None


class ChatMessage(BaseModel):
    """チャット履歴の1メッセージ. chat_history.jsonl の1行に対応.

    SPEC.md §3 にはこのモデルは明記されていないが、§5.4 のチャット仕様で
    必要なため追加する (実装の細部).
    """

    role: Literal["user", "assistant", "system"]
    content: str
    timestamp: str
    tool_trace: list[dict[str, Any]] = Field(default_factory=list)


class ChatSessionMeta(BaseModel):
    """チャットセッションのメタ情報.

    1つの Excel ジョブに複数の相談スレッドをぶら下げるための軽量メタ。
    """

    session_id: str
    title: str
    created_at: str
    updated_at: str
    archived: bool = False
    last_message_preview: str = ""
    message_count: int = 0


FeedbackKind = Literal[
    "thumbs_up",  # チャット応答に良い
    "thumbs_down",  # チャット応答にダメ
    "improvement",  # 改善要望
    "bug",  # 不具合報告
    "other",  # その他
]


class Feedback(BaseModel):
    """ユーザーからのフィードバック 1 件.

    永続化先: <jobs_dir>/_feedback/<YYYY-MM-DD>.jsonl
    UUID + ISO 時刻で 1 件ずつ append-only に書き出す.

    心理的負担を下げるため、`comment` 以外はすべて optional または自動収集.
    """

    id: str  # UUIDv4
    timestamp: str  # ISO 8601 (UTC)
    kind: FeedbackKind
    comment: str = ""  # 自由記述 (任意)

    # 自動収集コンテキスト (フィードバックを意味付けるための情報)
    page: str = ""  # ブラウザ URL パス (例: "/spec/abc-...")
    job_id: str | None = None  # 関連ジョブ
    target_id: str | None = None  # 対象オブジェクト (例: チャット応答のタイムスタンプ)
    target_excerpt: str = ""  # 対象の短い抜粋 (例: チャット応答の先頭 200 字)
    user_label: str = ""  # 任意の自己申告ラベル (匿名 OK)
