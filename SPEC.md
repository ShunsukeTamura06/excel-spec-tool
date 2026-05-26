# Excelツール改修支援AI 仕様書

## 1. 概要

長年使われてきた業務用Excelツール（VBA/数式/シート構造を含む `.xlsm` `.xls`）を、VBAに不慣れなユーザーでも安全に改修できるよう支援するWebアプリケーション。

### 1.1 ユーザーフロー

1. ユーザーがExcelツールをアップロードする
2. アプリがExcelツールの**統合設計書**（シート、VBA、数式、セル、参照関係）を生成する
3. ユーザーが設計書を文脈に乗せたチャットで「ここを改修したい」と伝える
4. LLMが改修手順と**波及範囲**（変更で影響を受ける他のセル・VBA箇所）を回答する
5. ユーザーは指示に従って手作業でExcelを改修する

### 1.2 想定ユーザー

- 主担当: 部内のツール保守担当者（VBAは詳しくない）
- 副担当: 部内の事務スタッフ5名（基本的なExcel操作のみ）

### 1.3 制約

- **外部クラウドへのデータ送信は禁止**。LLM呼び出しは社内LLM API経由のみ
- インターネット接続はプロキシ経由でURLホワイトリスト登録が必要
- 社内AWS環境（EC2）にデプロイする
- 利用可能な社内LLM: GPT-5.2, Gemini-3.1-pro（OpenAI互換APIを想定）

## 2. アーキテクチャ

3層構成で、依存方向は `frontend → backend → core` の単方向のみ。

```
┌─────────────────────────┐         ┌─────────────────────────┐
│  Frontend (Nuxt 3 SPA)  │  HTTP   │  Backend (FastAPI)      │
│  - アップロードUI       │ ──────> │  - /extract             │
│  - 設計書表示 + 図解    │ <────── │  - /analyze             │
│  - チャット画面         │         │  - /spec, /references   │
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

### 2.2 ディレクトリ構成

```
excel-spec-tool/
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
│   └── llm_client.py           # 社内LLM API呼び出し
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
├── SPEC.md
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

## 5. Backend層の仕様 (FastAPI)

### 5.1 エンドポイント

| Method | Path | リクエスト | レスポンス | 説明 |
|---|---|---|---|---|
| POST | `/extract` | multipart/form-data: `file` | `{"job_id": "..."}` | アップロード+構造抽出 |
| POST | `/analyze/{job_id}` | - | `{"status": "ok"}` | LLM注釈を付与し設計書生成 |
| GET | `/spec/{job_id}` | - | `{"spec_md": "...", "meta": {...}}` | 設計書取得 |
| GET | `/references/{job_id}` | query: `target` | `{"refs": [...]}` | 特定セルへの参照検索 |
| POST | `/chat/{job_id}` | `{"message": "..."}` | `{"reply": "...", "history": [...]}` | 改修対話 |
| GET | `/chat/{job_id}/history` | - | `{"history": [...]}` | チャット履歴取得 |
| DELETE | `/jobs/{job_id}` | - | `{"deleted": true}` | ジョブ削除 |

### 5.2 ストレージ (backend/storage.py)

ローカルファイル方式。

```
/var/excel-spec-tool/jobs/
└── {job_id}/                       # job_id = UUIDv4
    ├── original.xlsm               # アップロード原本
    ├── extracted.json              # Workbookモデル (Coreの抽出結果)
    ├── spec.md                     # 生成済み設計書
    ├── references.json             # ReferenceIndexモデル
    ├── chat_history.jsonl          # 1行1メッセージで追記
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

**実APIへの接続実装はShunが社内仕様を確認したのちに追記する。CLAUDE.mdに従いモック実装を先に置く。**

### 5.4 `/chat` の挙動

1. `chat_history.jsonl` から既存履歴を読む
2. 設計書 (`spec.md`) を system prompt に固定
3. 参照インデックス検索ツール（function calling）をLLMに渡す
4. 応答後、user/assistantメッセージをjsonlに追記
5. 応答には必ず以下のセクションを含めるようsystem promptで指示:
   - 改修手順（ユーザーの操作レベル）
   - 波及範囲（参照インデックスから引いた、影響を受けるセル/VBA）

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

- `pages/index.vue`: ホーム — ジョブ一覧 + 新規アップロード + analyze 進捗
- `pages/spec/[jobId].vue`: 設計書ページ。タブ構成:
  - 概要 (メトリクスダッシュボード)
  - シート (シート選択 + 数式/名前付き範囲/プレビュー)
  - VBA (モジュール/プロシージャ詳細)
  - ダイアグラム (シート依存図 / VBA コールグラフ)
  - 参照検索 (逆引き)
- `pages/chat/[jobId].vue`: 改修対話チャット
- 共通レイアウトに current job indicator + ナビゲーション

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
openai>=2.36   # 社内LLM (OpenAI互換)

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
JOBS_DIR=/var/excel-spec-tool/jobs    # 開発時は ./jobs
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
- 改修済みファイルの自動生成
- 複数ファイルの差分比較
- ジョブの自動削除cron
- LLM注釈の精度チューニング
