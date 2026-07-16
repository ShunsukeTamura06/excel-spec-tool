# xlblueprint 仕様書

## 1. 概要

長年使われてきた業務用Excelツール（VBA/数式/シート構造を含む `.xlsm` `.xls`）を、
**確実に解析し、安全に保守・改修できるようにする Excel 専用の道具**。

中核は Excel 固有の複雑な仕組み（VBA・セル・シート・オブジェクト・名前定義・数式・参照）
を確実に扱う `core`（純Python ライブラリ）であり、Web アプリケーション（backend +
frontend）はそれを使う入口の一つ（VBA が書けない社内担当者向け）。もう一つの入口として、
Claude Code 等のコーディングエージェント向けの MCP サーバー / CLI をロードマップ上で
用意する。背景・上位の目的は [docs/VISION.ja.md](./VISION.ja.md) §0 を参照。

### 1.1 ユーザーフロー

Web UI の主語は Excel の内部構造ではなく、一般ユーザーが達成したい次の3つの仕事とする。

#### A. 「このExcelを調べる」

1. ユーザーがExcelツールを選ぶ
2. アプリが、用途、使い方、入力、出力、主要機能、外部依存、注意点をまとめた
   **Excel診断**を生成する
3. 診断の各説明には、根拠となるシート、ボタン、数式、VBA、外部接続等を紐付ける
4. 確定できない説明は「推定」または「不明」と表示し、事実のように断定しない
5. ユーザーは必要なときだけ、シート、VBA、参照関係等の技術詳細を開く

#### B. 「このExcelについて質問する」

1. ユーザーがExcel診断または共通ナビゲーションから質問画面を開く
2. 用途、使い方、入力と出力、数字の根拠、機能、注意点等の質問テンプレートを選ぶか、
   自由文で質問する
   - 入力欄が空ならテンプレートを挿入する
   - 入力済みなら内容を自動で上書きせず、「後ろに追加」「置き換える」「キャンセル」を
     ユーザーが明示的に選ぶ
3. アプリが診断の根拠と技術詳細を参照し、確定事項・推定・不明を区別して回答する
4. 質問するために「直したい」導線や改修依頼書を経由させない

#### C. 「このExcelを直したい」

1. ユーザーが対象機能を選び、業務上の要望を自然文で入力する
2. アプリが、現状機能、変更したい結果、影響候補、確認事項、受入条件を
   **改修依頼書**として整理する
3. 根拠付きの解析結果を文脈にして改修相談を行い、変更可能なパターンでは変更計画を作る
4. アプリが原本を保持したまま修正版コピーを生成する
5. 構造差分、意図しない変更、未解析リスク、検証可能範囲を提示し、ユーザーが検収する

通常の画面では、Claude Code、OfficeCLI、openpyxl、MutationProvider 等の実装技術を
選択させない。これらは変更・検証能力を提供する内部実装として必要に応じて利用する。
利用可能な実装が限定される場合も、ユーザーには「現在自動で変更できる内容」と
「人または別環境での確認が必要な内容」として説明する。

> C の構造検証部分と、3種類の限定変更は実装済み。COM 再計算・マクロ実行による
> 動的検証と、VBAを含む任意変更は未実装。段階分け・実現可能性の検証結果・根拠は
> [docs/VISION.ja.md](./VISION.ja.md) を参照。

### 1.2 想定ユーザー

Web UI（本書の主対象）の想定ユーザー:

- 主担当: 業務とExcelの基本操作は理解しているが、引き継いだツールの内部構造を
  安全に読み解き、改修する自信がない保守担当者
- 副担当: VBAをある程度読めるが、未知のツールを短時間で把握し、変更漏れを避けたい担当者

Web UI 上でのこのユーザーの仕事は、診断結果を確認し、業務要望と受入条件を伝え、
「差分・波及範囲・リスクを見て承認する」ことに限定する（VISION.ja.md §0.2）。
コーディングエージェント（Claude Code 等）向けには別途 MCP サーバー / CLI を
用意する（ロードマップ、VISION.ja.md §4.4）。想定ユーザーを1つの画面に合わせて
丁寧化しすぎない — 入口を分けることでペルソナのブレを解消する。

### 1.3 設計原則

- **主張には根拠を付ける**: 用途や機能の説明は抽出した事実へ紐付ける。根拠のない
  推測を確定事項として表示しない
- **不明を残せる**: 静的解析だけでは分からない実行時挙動、外部システム、利用者の
  業務ルールは未保証として明示する
- **業務語を先、技術語を後にする**: 通常UIは「入力」「出力」「機能」「直したい内容」
  を中心にし、シート名・数式・VBA等は根拠や詳細として表示する
- **原本を変更しない**: 変更は常にコピーへ適用し、計画、実差分、検証結果、ハッシュを
  監査記録として残す

- **データの持ち出しを前提にしない**: LLM 呼び出しは OpenAI 互換 API
  であれば何でも差し替え可能 (ローカル LLM / セルフホスト LLM /
  クラウド LLM)。エアギャップ環境や閉鎖網にもデプロイできる構成を維持する
- **インターネットアクセスなしで動く**: フロントエンドはフォント・アイコン
  含めて同梱し、ランタイムで外部 CDN を叩かない
- **モデルは差し替え可能**: モデル名・エンドポイントは全て環境変数。
  例として OpenAI / Anthropic / Ollama / vLLM / 任意の OpenAI 互換
- **Windows + Excel クライアントへのオンプレ実行が前提 (P1 以降)**: 安全ゲート
  (§1.1 step5-6) は実際の Excel COM/VBIDE で再計算・マクロ実行して検証するため、
  Windows + Excel が入ったクライアント環境へのオンプレデプロイを前提とする。
  P0 (構造抽出・設計書・チャット) は OS 非依存のまま。LLM 呼び出しに関する
  上記の原則 (OpenAI 互換・エアギャップ対応) はこの前提でも変わらず維持する

## 2. アーキテクチャ

3層構成で、依存方向は `frontend → backend → core` の単方向のみ。Excelを書き換える
実装は `MutationProvider` 境界の外側に置き、変更後成果物は必ず core の再抽出・差分・
policy gate に戻す。

```
┌─────────────────────────┐         ┌─────────────────────────┐
│  Frontend (Nuxt 3 SPA)  │  HTTP   │  Backend (FastAPI)      │
│  - アップロードUI       │ ──────> │  - /extract             │
│  - 設計書表示 + 図解    │ <────── │  - /analyze             │
│  - 診断・改修依頼画面   │         │  - /diagnosis           │
│  - 技術詳細・相談       │         │  - /change-request      │
│                         │         │  - /chat, /diagrams     │
└─────────────────────────┘         └─────────────────────────┘
                                              │
                                              ↓
                                    ┌─────────────────────┐
                                    │  Core (純Python)    │
                                    │  - olevba抽出       │
                                    │  - openpyxl抽出     │
                                    │  - 参照インデックス │
                                    │  - 設計書生成       │
                                    └─────────────────────┘
```

### 2.1 各層の責務

| 層 | 責務 | やってはいけないこと |
|---|---|---|
| Core | Excel分解・参照解析・設計書生成。純Pythonライブラリとして単独で使える | HTTP/UI/LLMを知る |
| Backend | CoreをHTTPで包む。永続化、LLM呼び出し、チャット履歴管理 | Excelの解釈ロジックを持つ |
| Frontend | UI、Backendを叩く、表示 | Excel処理ロジックを持つ |

`core` は Excel 専用の道具の本体であり、Web UI（frontend + backend）はそれを使う
入口の一つに過ぎない（VISION.ja.md §0）。ロードマップ上、`core` を薄くラップする
MCP サーバー / CLI をもう一つの入口として追加する予定（現時点では未実装）。

### 2.2 ディレクトリ構成

```
xlblueprint/
├── core/
│   ├── __init__.py
│   ├── extractors/
│   │   ├── __init__.py
│   │   ├── vba.py              # olevbaラッパー
│   │   └── workbook.py         # openpyxlラッパー
│   ├── reference_index.py      # 参照インデックス構築
│   ├── spec_generator.py       # 設計書Markdown生成
│   └── models.py               # Pydanticモデル
│
├── backend/
│   ├── __init__.py
│   ├── main.py                 # FastAPIアプリ定義
│   ├── routes/
│   │   ├── __init__.py
│   │   ├── extract.py
│   │   ├── analyze.py
│   │   ├── spec.py
│   │   ├── references.py
│   │   └── chat.py
│   ├── storage.py              # ローカルファイル永続化
│   └── llm_client.py           # OpenAI 互換 LLM API クライアント
│
├── frontend/                   # Nuxt 3 SPA (TypeScript)
│   ├── nuxt.config.ts
│   ├── package.json
│   ├── app.vue
│   ├── layouts/                # 共通レイアウト (sidebar + content)
│   ├── pages/                  # ルーティング: /, /spec/[jobId], /chat/[jobId]
│   ├── components/             # JobCard, DiagramView 等
│   ├── composables/            # useBackend (typed API client)
│   ├── stores/                 # Pinia: 現在ジョブ等
│   └── types/                  # Pydantic と対応する TypeScript 型
│
├── tests/
│   ├── core/
│   ├── backend/
│   └── fixtures/               # サンプルExcel
│
├── docs/SPEC.ja.md
├── CLAUDE.md
├── README.md
├── pyproject.toml
└── requirements.txt
```

## 3. データモデル (core/models.py)

すべてPydanticで定義。CoreとBackendとFrontendで共有する。

```python
from pydantic import BaseModel
from typing import Literal

class VbaProcedure(BaseModel):
    name: str                          # "UpdateDaily"
    kind: Literal["Sub", "Function", "Property"]
    start_line: int
    end_line: int
    code: str                          # 該当プロシージャのコード片
    annotation: str = ""               # LLMによる意図注釈

class VbaModule(BaseModel):
    name: str                          # "Module1"
    type: Literal["Module", "Class", "Form", "Document"]
    code: str                          # モジュール全体のコード
    procedures: list[VbaProcedure] = []

class CellFormula(BaseModel):
    coord: str                         # "Calc!H2"
    formula: str                       # "=SUMIF(Input!A:A, A2, Input!E:E)"
    refs: list[str] = []               # 参照先 ["Input!A:A", "Calc!A2", "Input!E:E"]
    annotation: str = ""               # LLMによる意図注釈

class NamedRange(BaseModel):
    name: str                          # "顧客マスタ"
    refers_to: str                     # "Calc!$A$2:$D$5000"

class ConditionalFormat(BaseModel):
    range: str
    rule: str

class SheetInfo(BaseModel):
    name: str
    rows: int                          # 使用行数
    cols: int                          # 使用列数
    formulas: list[CellFormula] = []
    named_ranges: list[NamedRange] = []
    conditional_formats: list[ConditionalFormat] = []
    purpose: str = ""                  # LLM推定の用途

class Workbook(BaseModel):
    filename: str
    sheets: list[SheetInfo] = []
    vba_modules: list[VbaModule] = []
    external_links: list[str] = []

class Reference(BaseModel):
    kind: Literal["formula", "vba"]
    from_: str                         # "Output!K3" or "Module1.UpdateDaily:L47"
    to: str                            # "Calc!H2"
    code: str = ""                     # 該当箇所の生コード

class ReferenceIndex(BaseModel):
    """逆引きインデックス: あるセル/範囲を参照しているもの一覧"""
    refs: dict[str, list[Reference]]   # key: 参照先 ("Calc!H2:H5000")

class JobMeta(BaseModel):
    job_id: str
    filename: str
    created_at: str                    # ISO形式
    status: Literal["uploaded", "extracted", "analyzed", "failed"]
```

## 4. Core層の仕様

### 4.1 `core/extractors/vba.py`

```python
def extract_vba(file_path: Path) -> list[VbaModule]
```

- olevba (`oletools`) でVBAソースを抽出
- 各モジュールをパースして `VbaModule` に詰める
- `procedures` は `Sub`/`Function`/`Property` の宣言行を正規表現で検出して切り出す
- VBAプロジェクトロックされた `.xls` も `olevba` でデコード可能。失敗時は空リストを返してログに残す
- `.xls`（古いバイナリ形式）と `.xlsm` 両対応

### 4.2 `core/extractors/workbook.py`

```python
def extract_workbook(file_path: Path) -> Workbook
```

- `openpyxl.load_workbook(keep_vba=True)` でロード
- 各シートを走査:
  - `cell.data_type == 'f'` のセルから `CellFormula` を作る
  - 数式は `openpyxl.formula.Tokenizer` で分解し、`subtype == "RANGE"` のトークンを `refs` に詰める
  - 名前付き範囲は `wb.defined_names` から取得
  - 条件付き書式は `ws.conditional_formatting` から
  - Excel テーブル、入力規則、フォームコントロール、グラフ、ピボットテーブルを抽出
- グラフ / ピボット / Power Query・外部接続は OOXML パーツを直接読む:
  - グラフは `xl/drawings/*.xml` と `xl/charts/chart*.xml` から種類・配置・系列参照を取得
  - ピボットは `xl/pivotTables/` と `xl/pivotCache/` から配置・元データ・主要フィールドを取得
  - Power Query は `xl/connections.xml` と `queryTable` から接続・出力先を棚卸しする
  - 接続文字列内のパスワード / トークン等は必ずマスクする
- `.xls` は `openpyxl` で読めない。`.xls` が来た場合は `extract_vba` のみ実行し、`Workbook.sheets` は空で返す。READMEでも明記
- 外部リンクは `wb._external_links` から（プライベートAPI使用、バージョン依存に注意）

### 4.3 `core/reference_index.py`

```python
def build_reference_index(wb: Workbook) -> ReferenceIndex
```

- 数式側: `wb.sheets[*].formulas[*].refs` を逆引きする
- グラフ側: chart XML から明示的に取れた系列値・カテゴリ範囲を逆引きする
- ピボット側: pivot cache から明示的に取れた元データ範囲を逆引きする
- VBA側: `wb.vba_modules[*].code` を正規表現で走査
  - 対象パターン:
    - `Range("A1:J100")` → 現在のシート
    - `Worksheets("Calc").Range("A1")` → 指定シート
    - `Sheets("Calc").Cells(2, 8)` → 指定シート + 行列
    - `[Calc!A1]` → 短縮表記
  - 検出した参照は `Reference(kind="vba", from_="Module1.UpdateDaily:L47", to="Calc!A1:J100", code='Range("A1:J100")')` で登録
- **完璧なパースは目指さない**。捕捉漏れの可能性は未解析リスクとして明示し、
  LLM は推測で補完しない
- 必ずテストケースで「最低限これだけは捕捉する」を担保する

### 4.5 `core/risk_analyzer.py`

```python
def detect_analysis_risks(wb: Workbook) -> list[AnalysisRisk]
```

- 静的解析では影響範囲を断定できない箇所を検出する
- 対象例:
  - 動的 VBA 参照: `Range(addr)`, `Cells(row, col)`, `Worksheets(sheetName)`
  - 実行時状態依存: `ActiveSheet`, `Selection`, `CurrentRegion`, `UsedRange`, `Offset`, `Resize`
  - 暗黙実行: `Worksheet_Change`, `Workbook_Open` 等のイベントプロシージャ
  - 動的数式: `INDIRECT`, `OFFSET`, `CELL`, `INFO`
  - 外部依存: 外部リンク、Power Query / 外部接続、外部 Add-In 関数
  - 依存先不明のグラフ / ピボット
- 目的は完全解析ではなく、チャット回答で「影響なし」と誤断定しないための
  手動確認リストを作ること

### 4.4 `core/spec_generator.py`

```python
def generate_spec(wb: Workbook, ref_index: ReferenceIndex) -> str
```

- Markdown形式で統合設計書を出力
- セクション構成:

```markdown
# 設計書: {filename}

## 1. 概要
- ファイル名 / シート数 / VBAモジュール数 / 外部リンク

## 2. シート一覧
（テーブルで一覧）

## 3. シート詳細
### {シート名}
- 用途（推定）
- 主要数式 (TOP 10程度)
- 名前付き範囲
- 条件付き書式

## 4. VBAモジュール
### {モジュール名}
- プロシージャ一覧（注釈付き）
- ソースコード（折りたたみ想定）

## 5. 参照関係（抜粋）
- 主要なシート間・VBA-シート間の依存

## 6. 注意点・観察事項
- LLMが検出したデッドコード/アンチパターン（後段で追加される）
```

- LLM注釈は別関数 `annotate_with_llm(wb, llm_client) -> Workbook` で行い、`generate_spec` 自体はLLM非依存

### 4.6 `core/workbook_diff.py` (P1 安全ゲート)

```python
def diff_workbooks(before_path, after_path, before_wb, after_wb, before_index) -> WorkbookDiff
```

- 2バージョンのワークブックの構造差分 (セル・名前付き範囲・条件付き書式・
  入力規則・グラフ・ピボット・VBAモジュール) と波及範囲 (blast radius) を計算する
- セル差分は生XMLではなく正規化抽出 (`extract_cells_to_sqlite`) を比較基盤にして
  保存ノイズを無視する (VISION §6.2/§6.3 のスパイクで実証済みの手法)

### 4.7 `core/named_range_fix.py` / `core/formula_fix.py` (S2 安全パターン自動修正)

§1.1 の 6 (自動適用) のうち、実装済みの安全パターン:

| パターン | propose (試算, read-only) | apply (実ファイル書き込み) |
|---|---|---|
| 名前定義修正 | `propose_named_range_fix(wb, idx, name, new_refers_to)` | `apply_named_range_fix(path, name, new_refers_to, out)` |
| 固定参照置換 | `propose_fixed_ref_replace(wb, idx, old_ref, new_ref)` | `apply_fixed_ref_replace(path, old_ref, new_ref, out)` |
| 数式範囲拡張 | `propose_range_expansion(wb, idx, old_range, new_range)` | `apply_range_expansion(path, old_range, new_range, out)` |

- propose はチャットの tool loop から LLM が呼んでよい (ファイル未変更の試算)。
  apply は人間が画面のボタンを押した時だけ backend 経由で呼ばれる
  (「黙って変更しない」、VISION §4.2)
- 数式の書き換えは文字列置換ではなく openpyxl の formula Tokenizer による
  参照トークン単位の置換 (VISION §4.3)。文字列リテラルや部分一致する別参照は
  影響を受けない
- 範囲拡張は「同一シートで旧範囲を包含する新範囲」だけを受け付ける
  (縮小・移動は対象外)
- apply 後は新ジョブを作り、`diff_workbooks` で before/after を比較して
  意図した変更だけが起きたかを自己検証する

### 4.8 `core/mutation.py` / `core/verification.py` (検証コントロールプレーン)

- 変更意図はプロバイダー非依存の `MutationPlan` として記録する
- `MutationProvider` は元ファイルを変更せず、隔離された出力先へ成果物を作る。
  組み込み `openpyxl` と任意の OfficeCLI アダプターを同じ契約で扱う
- OfficeCLI 対応は現時点で `.xlsx` の名前定義変更のみ。`.xlsm`、数式変更、VBA変更は
  未対応として明示的に拒否し、対応範囲を能力APIで返す
- propose段階の期待構造差分と、成果物をフル再抽出して得た実構造差分を
  `exact-structural-diff-v1` で照合する。予定変更の欠落・予定外変更・内容不一致は
  `failed`、完全一致でも波及先または高リスク項目があれば `needs_review`、それ以外を
  `passed` とする
- `ChangeExecutionRecord` に変更計画、プロバイダーと版、変更前後のSHA-256、期待差分、
  実差分、判定を保存する。
  プロバイダーの終了成功だけを安全性の証拠にはしない
- この判定が保証するのは観測対象の構造一致であり、Excel再計算値・マクロ副作用などの
  動的挙動は保証しない。COM検証が追加されるまでUIにもこの境界を残す

### 4.9 一般ユーザー向け安全変更フロー

- 改修依頼書の作成後、現在自動対応できる限定変更を一般ユーザー向けに提示する
- 最初の完成経路は「数式が参照するデータ範囲の拡張」とする
  - 抽出済み数式から、シート修飾された有界範囲と利用数を候補として表示する
  - ユーザーは対象範囲と新しい最終行を選ぶ。数式文字列を直接編集させない
  - 新範囲が旧範囲を包含しない場合は変更計画を作らない
- 適用前に `MutationPlan`、期待差分、変更される数式数、影響候補、既存リスク、
  静的検証の限界を表示し、ユーザーの明示承認を要求する
- 承認後は、表示したものと同じ `MutationPlan.plan_id` を実行・監査記録へ保存する
- 原本は変更せず、修正版を新ジョブとして作成する
- 適用後は期待差分と実差分のpolicy判定を一般ユーザー向けに表示し、修正版を
  ダウンロードできるようにする
- 対応外の改修依頼は自動適用可能と見せず、根拠付き相談へ案内する

## 5. Backend層の仕様 (FastAPI)

### 5.1 エンドポイント

| Method | Path | リクエスト | レスポンス | 説明 |
|---|---|---|---|---|
| POST | `/extract` | multipart/form-data: `file` | `{"job_id": "..."}` | アップロード+構造抽出 |
| POST | `/analyze/{job_id}` | - | `{"status": "ok"}` | LLM注釈を付与し設計書生成 |
| GET | `/diagnosis/{job_id}` | - | `WorkbookDiagnosis` | 根拠付きExcel診断を取得 |
| POST | `/change-request/{job_id}` | `{"requested_outcome": "...", "feature_id": "F001"}` | `ChangeBrief` | 業務要望を改修依頼書へ整理 |
| POST | `/jobs/{job_id}/change-plan` | `{"kind": "range_expansion", "old_ref": "...", "new_ref": "..."}` | `SafeChangePlan` | 適用前の期待差分と検証条件を作成 |
| POST | `/jobs/{job_id}/change-plan/execute` | `{"plan": {...}}` | 修正ジョブ・実差分・検証結果 | 表示済み変更計画を明示承認後に適用 |
| GET | `/jobs/{job_id}/download` | - | Excelファイル | 原本または検証済み修正版を取得 |
| GET | `/spec/{job_id}` | - | `{"spec_md": "...", "meta": {...}}` | 設計書取得 |
| GET | `/references/{job_id}` | query: `target` | `{"refs": [...]}` | 特定セルへの参照検索 |
| POST | `/chat/{job_id}` | `{"message": "..."}` | `{"reply": "...", "history": [...]}` | 改修対話 |
| GET | `/chat/{job_id}/history` | - | `{"history": [...]}` | チャット履歴取得 |
| DELETE | `/jobs/{job_id}` | - | `{"deleted": true}` | ジョブ削除 |
| GET | `/diff` | query: `before_job_id`, `after_job_id` | `{"diff": {...}}` | 2ジョブ間の構造差分 (§4.6) |
| POST | `/jobs/{job_id}/named-range-fix` | `{"name": "...", "new_refers_to": "..."}` | `{"new_job_id": "...", "diff": {...}}` | 名前定義修正の適用+自己検証 (§4.7) |
| POST | `/jobs/{job_id}/formula-fix` | `{"kind": "fixed_ref_replace" \| "range_expansion", "old_ref": "...", "new_ref": "..."}` | `{"new_job_id": "...", "diff": {...}}` | 固定参照置換/範囲拡張の適用+自己検証 (§4.7) |
| GET | `/mutation-providers` | - | `{"providers": [...]}` | 変更プロバイダーの利用可否・版・対応範囲 |
| GET | `/jobs/{job_id}/verification` | - | `{"verification_record": {...}}` | 変更計画と検証証拠の監査レコード |

`named-range-fix` / `formula-fix` は任意の `provider` (`openpyxl` / `officecli`) を受け取り、
レスポンスに従来の `new_job_id` / `diff` に加えて `plan` / `provider` / `verification` を返す。
policy判定が `failed` の場合は HTTP 409 とし、失敗した成果物は正常完了として扱わない。

### 5.2 ストレージ (backend/storage.py)

ローカルファイル方式。

```
/var/xlblueprint/jobs/
└── {job_id}/                       # job_id = UUIDv4
    ├── original.xlsm               # アップロード原本
    ├── extracted.json              # Workbookモデル (Coreの抽出結果)
    ├── diagnosis.json              # 根拠・確度付きExcel診断
    ├── spec.md                     # 生成済み設計書
    ├── references.json             # ReferenceIndexモデル
    ├── chat_history.jsonl          # 1行1メッセージで追記
    ├── verification.json           # 変更計画・期待/実差分・policy判定
    └── meta.json                   # JobMetaモデル
```

実装上のルール:
- `job_id` は受信時にUUID形式バリデーション（パスインジェクション防止）
- ディレクトリ作成時のパーミッションは700
- チャット履歴は `jsonl` で追記。`open(..., "a")` を使う
- 7日経過したジョブはcron+別スクリプトで削除（実装は別タスク）

### 5.3 LLMクライアント (backend/llm_client.py)

OpenAI互換APIを想定。設定は環境変数から：

```
LLM_BASE_URL=http://internal-llm.example.com/v1
LLM_API_KEY=...
LLM_MODEL=gpt-5.2
```

公開関数:
```python
def chat_completion(messages: list[dict], model: str = None) -> str
def annotate_text(prompt: str, content: str) -> str
```

実 API への接続は `openai` SDK の OpenAI 互換クライアント (`base_url` / `api_key` / `model`) で実装済み。環境変数が揃っていなければ `MockLLMClient` にフォールバックする。

### 5.4 `/chat` の挙動

1. `chat_history.jsonl` から既存履歴を読む
2. 設計書 (`spec.md`) を system prompt に固定
3. 参照インデックス検索ツール（function calling）をLLMに渡す
4. 応答後、user/assistantメッセージをjsonlに追記
5. 応答には必ず以下のセクションを含めるようsystem promptで指示:
   - 改修手順（ユーザーの操作レベル）
   - 波及範囲（参照インデックスから引いた、影響を受けるセル/VBA）
6. 改修手順は「差分」ではなく、コピー&ペーストで完結する形で提示するようsystem promptで指示:
   - VBA: 書き換え後のプロシージャ全体を提示し、既存プロシージャを丸ごと
     置き換える操作（どのモジュールのどのプロシージャを、どう置き換えるか）を明記する
   - 数式: 書き換え後の完全な式を提示し、対象セル（シート名+セル番地）への
     貼り付け方法（数式として貼り付けるか、値のみ貼り付けるか）を明記する
   - §1.2 の想定ユーザー（VBA に不慣れ）が誤った箇所を書き換えてバグを
     埋め込むリスクを下げるための措置。§1.1 の6（自動適用、S2増分）が
     対応するパターンを広げるまでの暫定策

## 6. Frontend層の仕様 (Nuxt 3 SPA)

### 6.1 技術スタック

- **Nuxt 3** + TypeScript + SPA モード (`ssr: false`)
- **Nuxt UI v3** (Tailwind v4 ベース) — コンポーネント / ダークモード / アクセシビリティ
- **Vue Flow** (`@vue-flow/core`) — シート依存・VBA コールグラフのインタラクティブ描画
- **Pinia** — 状態管理 (現在ジョブ等)
- **`@nuxtjs/mdc`** + Shiki — 設計書 Markdown 描画
- HTTP: Nuxt 標準 `$fetch` / `ofetch`
- パッケージ管理: pnpm (corepack)

### 6.2 画面構成

- `pages/index.vue`: ホーム — Excel一覧 + 「このExcelを調べる」入口
- `pages/spec/[jobId].vue`: Excel診断ページ。タブ構成:
  - 診断 (用途・機能・入力・出力・外部依存・注意点・根拠)
  - シート (シート選択 + 数式/名前付き範囲/プレビュー)
  - VBA (モジュール/プロシージャ詳細)
  - ダイアグラム (シート依存図 / VBA コールグラフ)
  - 参照検索 (逆引き)
- `pages/chat/[jobId].vue`: 改修対話チャット
- `pages/change/[jobId].vue`: 対象機能と業務要望から改修依頼書を作成
  - 対応可能な範囲拡張では、候補選択、変更計画、明示承認、適用、検証結果、
    修正版ダウンロードまでを同一画面で行う
- 共通レイアウトに選択中Excel + 一般ユーザー向けナビゲーション

### 6.3 状態管理

- Pinia store `useJobStore` で現在のジョブ ID と job リストを管理
- リロード耐性: 選択中 jobId を `localStorage` に永続化
- URL に jobId を含めることでブックマーク可能 (`/spec/<jobId>`)

### 6.4 API 呼び出し (composables/useBackend.ts)

Nuxt の `$fetch` を typed client にラップ。Backend URL は `runtimeConfig.public.backendUrl`
（環境変数 `NUXT_PUBLIC_BACKEND_URL` 由来、デフォルト `http://localhost:8001`）。

開発時の CORS: Backend は `http://localhost:3001` (Nuxt dev server) を許可する
`CORSMiddleware` を設定する。本番ではフロントを backend と同一オリジンにデプロイ
する想定 (静的書き出しを backend が配信、または同一ホストの reverse proxy 配下)。

## 7. 依存パッケージ

### Python (uv)

```
# core
oletools>=0.60
openpyxl>=3.1
pydantic>=2.0

# backend
fastapi>=0.110
uvicorn[standard]>=0.27
python-multipart>=0.0.9
httpx>=0.27
openai>=2.36   # OpenAI 互換 API クライアント

# dev
pytest>=8.0
pytest-asyncio>=0.23
ruff>=0.4
mypy>=1.9
```

### Frontend (pnpm)

```
nuxt              ^3
@nuxt/ui          ^3
@vue-flow/core    ^1
@vue-flow/controls
@vue-flow/background
@nuxtjs/mdc       ^0
pinia             ^2
@pinia/nuxt       ^0
```

## 8. 起動方法 (開発時)

```bash
# Backend
uv run uvicorn backend.main:app --reload --port 8001

# Frontend (別ターミナル)
cd frontend && pnpm install && pnpm dev   # http://localhost:3001
```

環境変数:
```
# Backend
JOBS_DIR=/var/xlblueprint/jobs    # 開発時は ./jobs
LLM_BASE_URL=http://...
LLM_API_KEY=...
LLM_MODEL=gpt-5.2
CORS_ALLOW_ORIGINS=http://localhost:3001  # カンマ区切り複数可

# Frontend
NUXT_PUBLIC_BACKEND_URL=http://localhost:8001
NUXT_PORT=3001
```

## 9. スコープ外（最初は作らない）

- 認証・ユーザー管理
- Docker化
- ジョブの自動削除cron
- LLM注釈の精度チューニング

> 「改修済みファイルの自動生成」「複数ファイルの差分比較」はスコープ外から
> ロードマップ (P1〜) へ移動した。段階分け・根拠は
> [docs/VISION.ja.md](./VISION.ja.md) を参照。
