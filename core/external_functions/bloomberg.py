"""Bloomberg Excel Add-In の関数定義.

対象: BDH / BDP / BDS の 3 大関数. 説明文はすべて自作日本語要約.

参考: Bloomberg Terminal の HELP HELP コマンドおよび WAPI <GO> 画面の
ドキュメント (顧客は Bloomberg Anywhere ライセンスを保有しているため
社内利用に問題なし). 関数シグネチャ・引数名は公開事実情報.
"""

from __future__ import annotations

from core.external_functions.registry import ExternalFunction, ExternalFunctionParam

# ---------------------------------------------------------------- BDH

BDH = ExternalFunction(
    name="BDH",
    vendor="Bloomberg",
    short="Bloomberg Data History — ヒストリカル時系列データを取得する",
    long=(
        "BDH は指定された証券 (security) と項目 (field) について、指定期間の"
        "ヒストリカル値を 2 次元配列で返す Array 関数。Bloomberg Excel Add-In の"
        "中で最も多用される関数で、株価・為替・金利・指標などの過去データを"
        "Excel に取り込む用途に使う。\n"
        "\n"
        "返り値は配列なのでセルを跨いで広がる。先頭列は日付 (または期間ラベル) で、"
        "残りの列が field 毎の値となる。新しい配列数式 (動的配列対応 Excel) では"
        "通常のセル入力で動くが、レガシー Excel では Ctrl+Shift+Enter での"
        "配列確定が必要なケースがある。"
    ),
    signature='=BDH(security, fields, start_date, [end_date], [optional_args...])',
    params=[
        ExternalFunctionParam(
            name="security",
            description=(
                "証券識別子。Bloomberg ticker + yellow key の形式 "
                '(例: "AAPL US Equity", "USDJPY Curncy", "GT10 Govt", "SPX Index")。'
                "セル参照でも可。"
            ),
            type="security",
        ),
        ExternalFunctionParam(
            name="fields",
            description=(
                'データ項目コード。1 つの文字列 (例: "PX_LAST") または範囲参照で'
                "複数指定可。代表的: PX_LAST, PX_OPEN, PX_HIGH, PX_LOW, "
                "VOLUME, CUR_MKT_CAP, BEST_EPS, TOT_RETURN_INDEX_GROSS_DVDS。"
            ),
            type="field",
        ),
        ExternalFunctionParam(
            name="start_date",
            description=(
                "開始日。Excel 日付値、セル参照、または特殊文字列 "
                '("-1Y", "-3M", "TODAY()") も可。'
            ),
            type="date",
        ),
        ExternalFunctionParam(
            name="end_date",
            description=(
                "終了日 (省略時は最新日)。start_date と同様の形式。"
            ),
            required=False,
            type="date",
        ),
        ExternalFunctionParam(
            name="optional_args",
            description=(
                "Per_argument: 期間粒度 (DAILY/WEEKLY/MONTHLY/QUARTERLY/YEARLY)。"
                "Days: 営業日/暦日 (W/A/T 等)。Fill: 欠損埋め (P=previous, B=blank)。"
                "CDR: カレンダー (5D, 7D, ローカル休日 等)。Currency: 通貨換算。"
                "他多数。`name=value` の形で並べる。"
            ),
            required=False,
            type="string",
        ),
    ],
    returns=(
        "2 次元配列。先頭列が日付、以降の列が field 値。期間内に値がない場合は"
        "#N/A や空セルで返る (Fill オプションで挙動を制御可能)。"
    ),
    examples=[
        '=BDH("AAPL US Equity", "PX_LAST", "2024-01-01", "2024-12-31")',
        '=BDH("USDJPY Curncy", "PX_LAST", TODAY()-365, TODAY(), "Per", "DAILY")',
        '=BDH(A1, B1:B3, "-3M", TODAY(), "Days", "W", "Fill", "P")',
    ],
    notes=[
        "security は厳密に Bloomberg ticker + yellow key. ティッカーだけだと #N/A になる。",
        (
            "fields コードは Bloomberg 仕様 (FLDS <GO>) で確認する。"
            "SUM や AVERAGE のような Excel 関数名とは別物。"
        ),
        "BDH は配列を返すので、上書きされる範囲に既存データがあると #SPILL! / 配列衝突になる。",
        "オプション引数は name/value のペアで並ぶ。引数の順序ではなく名前で解釈される。",
        "営業日カレンダーや為替換算の指定は実務で頻出するので、optional_args を見落とさないこと。",
    ],
    doc_url="https://www.bloomberg.com/professional/support/api-library/  (terminal 内 WAPI <GO>)",
)


# ---------------------------------------------------------------- BDP

BDP = ExternalFunction(
    name="BDP",
    vendor="Bloomberg",
    short="Bloomberg Data Point — 最新時点の単一データ値を取得する",
    long=(
        "BDP は指定証券 / 項目の「最新 (現時点) の 1 値」を返す Scalar 関数。"
        "リアルタイム価格・名称・通貨・セクター分類など、時系列ではないスナップショット情報の取得に使う。\n"
        "\n"
        "リアルタイム購読設定が有効なフィールド (例: PX_LAST) の場合、ティック更新で"
        "自動再計算される。BDH と違い 1 セルに 1 値が返るだけなので spilling は起こらない。"
    ),
    signature='=BDP(security, field, [optional_args...])',
    params=[
        ExternalFunctionParam(
            name="security",
            description=(
                'Bloomberg ticker + yellow key 形式の証券識別子。'
                'BDH と同じ規則 (例: "AAPL US Equity")。'
            ),
            type="security",
        ),
        ExternalFunctionParam(
            name="field",
            description=(
                'データ項目コード (1 つ)。例: PX_LAST, NAME, INDUSTRY_SECTOR, '
                "CRNCY, CUR_MKT_CAP, EQY_DVD_YLD_IND。FLDS <GO> で検索可能。"
            ),
            type="field",
        ),
        ExternalFunctionParam(
            name="optional_args",
            description=(
                "オーバーライド (e.g. EQY_FUND_CRNCY=JPY) や換算指定。"
                "name/value のペアで並べる。"
            ),
            required=False,
            type="string",
        ),
    ],
    returns="単一値 (数値 / 文字列 / 日付)。取得不能時は #N/A。",
    examples=[
        '=BDP("AAPL US Equity", "PX_LAST")',
        '=BDP("7203 JP Equity", "NAME")',
        '=BDP("USDJPY Curncy", "PX_LAST")',
        '=BDP(A2, "INDUSTRY_SECTOR")',
    ],
    notes=[
        (
            "BDP はリアルタイム購読対象フィールドだとセルが自動更新される。"
            "再計算が頻発するシートでは負荷源になることがある。"
        ),
        "ヒストリカルが欲しい時は BDH を使う。BDP に日付を渡しても無視される。",
        "結果が #N/A になる場合、ticker の yellow key 抜けや field 名のスペルミスがよくある原因。",
    ],
    doc_url="https://www.bloomberg.com/professional/support/api-library/  (terminal 内 WAPI <GO>)",
)


# ---------------------------------------------------------------- BDS

BDS = ExternalFunction(
    name="BDS",
    vendor="Bloomberg",
    short="Bloomberg Data Set — 構造化された複数値 (テーブル) を取得する",
    long=(
        "BDS は指定証券について「構造化された複数行・複数列のテーブル」を返す Array 関数。"
        "BDP が 1 値、BDH が 1 項目の時系列なのに対し、BDS は項目自体が "
        "コレクション (例: 構成銘柄一覧、配当履歴、株主名簿、関連企業) になっている場合に使う。\n"
        "\n"
        "返り値は 2 次元配列で、行数・列数は項目によって可変。Bloomberg 公式の field "
        "リストで「Bulk Data」と分類されているフィールドが BDS の対象。"
    ),
    signature='=BDS(security, field, [optional_args...])',
    params=[
        ExternalFunctionParam(
            name="security",
            description=(
                "証券識別子。インデックスの構成銘柄を取りたい場合は "
                '"SPX Index" のように Index で指定。'
            ),
            type="security",
        ),
        ExternalFunctionParam(
            name="field",
            description=(
                'Bulk data フィールド。例: INDX_MEMBERS (構成銘柄), DVD_HIST_ALL '
                "(配当履歴), TOP_20_HOLDERS_PUBLIC_FILINGS, RELATED_SECURITIES。"
            ),
            type="field",
        ),
        ExternalFunctionParam(
            name="optional_args",
            description=(
                "Header=Y で列ヘッダ付き、Cols でカラム数指定、period や cdr "
                "等の絞り込み引数を name/value で並べる。"
            ),
            required=False,
            type="string",
        ),
    ],
    returns=(
        "2 次元配列 (行数・列数は field 依存)。Header=Y を渡せば 1 行目に "
        "列名が入る。データが存在しない場合は #N/A もしくは空配列。"
    ),
    examples=[
        '=BDS("SPX Index", "INDX_MEMBERS")',
        '=BDS("AAPL US Equity", "DVD_HIST_ALL", "Header", "Y")',
        '=BDS("7203 JP Equity", "TOP_20_HOLDERS_PUBLIC_FILINGS")',
    ],
    notes=[
        (
            "BDS は出力サイズが動的に決まるので、"
            "貼り付けた場所の下方/右方にあるデータを書き換える。レイアウト崩れに注意。"
        ),
        (
            "Bulk data フィールドは Bloomberg 内で別管理。"
            "FLDS <GO> でフィールド分類を「Bulk」に絞ると見つけやすい。"
        ),
        (
            "巨大なテーブル (構成銘柄数千件など) を頻繁に取ると "
            "Bloomberg 側の制限 (DAPI 制限) にかかることがある。"
        ),
    ],
    doc_url="https://www.bloomberg.com/professional/support/api-library/  (terminal 内 WAPI <GO>)",
)


# レジストリが読み込む順序付きリスト. 追加時はここに append.
FUNCTIONS: list[ExternalFunction] = [BDH, BDP, BDS]
