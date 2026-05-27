# Excelツール改修支援AI

VBA / 数式 / 参照関係を含む `.xlsm` `.xls` 業務 Excel を、VBA に不慣れな
担当者でも安全に改修できるよう支援する Web アプリ。

ファイルをアップロードすると、シート構造・VBA・数式・参照関係を抽出して
**統合設計書 (Markdown) + インタラクティブな依存グラフ** を生成し、
**LLM とのチャット** で「ここを直したい」と尋ねれば改修手順と波及範囲を
回答する。

詳細仕様は [SPEC.md](./SPEC.md) を、開発ルールは [CLAUDE.md](./CLAUDE.md) を参照。

## 主な機能

- **設計書生成** — シート / VBA / 数式 / 名前付き範囲 / 条件付き書式を
  抽出し、LLM 注釈つきの Markdown にまとめる
- **シート依存グラフ** — どのシートがどのシートを参照しているかを Vue Flow
  で可視化 (重みつきエッジ、ズーム・ドラッグ・MiniMap)
- **VBA コールグラフ** — プロシージャ間の呼び出し関係を同様に可視化
- **Excel オブジェクト棚卸し** — グラフ系列参照 / ピボット元データ /
  Power Query・外部接続を抽出し、設計書と参照検索に反映
- **未解析リスク検出** — 動的 VBA 参照 / イベント処理 / INDIRECT・OFFSET /
  外部接続など、影響なしと断定できない箇所を明示
- **参照逆引き** — 「このセルを参照しているのは誰か」を全件検索
- **改修チャット** — LLM が設計書 + ツール
  (cells / references / VBA 詳細 / 未解析リスク) を引きながら、
  確認済み事実・波及範囲・手動確認チェックリストを分けて提示

## アーキテクチャ

```
Frontend (Nuxt 3 SPA, TypeScript)         Backend (FastAPI)
├── Nuxt UI v3 + Tailwind v4         <──> ├── /extract, /analyze
├── Vue Flow + dagre (図解)               ├── /spec, /workbook
├── @nuxtjs/mdc (Markdown 描画)           ├── /references, /diagrams
└── Pinia                                 ├── /chat (function calling)
                                          └── /jobs

                                          Core (純 Python)
                                          ├── olevba / openpyxl 抽出
                                          ├── reference_index
                                          ├── spec_generator
                                          └── diagrams (sheet/VBA graph)
```

依存方向は `frontend → backend → core` の単方向のみ。

## セットアップ

### 必要なツール

- Python 3.10+ と [uv](https://docs.astral.sh/uv/)
- Node 20+ と pnpm (corepack 経由)

### 初回

```bash
# Backend (Python)
uv sync --group dev      # .venv 作成 + pytest/ruff/mypy も含めて同期

# Frontend (Node)
corepack enable
cd frontend && pnpm install
```

### 依存追加

```bash
# Python
uv add <package>             # 本体依存
uv add --group dev <pkg>     # 開発用

# Node
cd frontend && pnpm add <pkg>
cd frontend && pnpm add -D <pkg>  # devDependencies
```

## 起動 (開発時)

初回はフロントエンドの `.env` を用意する:

```bash
cd frontend
cp .env.example .env
# 必要に応じて NUXT_PUBLIC_BACKEND_URL 等を編集
```

ターミナルを 2 つ開く:

```bash
# Backend — http://localhost:8001
uv run uvicorn backend.main:app --reload --port 8001

# Frontend — http://localhost:3001
cd frontend && pnpm dev
```

### ローカル本番系 / 開発系を分離して同時起動

ユーザー検証用の本番系と、改修作業用の開発系を同じPC上で分けて起動できる。
ポート、ジョブ保存先、ログ、PID はすべて分離される。

| 系統 | Frontend | Backend | Jobs | Backend | Frontend |
|---|---:|---:|---|---|---|
| prod | 3001 | 8001 | `runtime/prod/jobs` | reload なし | `pnpm generate` + `pnpm preview` |
| dev | 3002 | 8002 | `runtime/dev/jobs` | reload あり | `pnpm dev` |

```bash
# ユーザーに触ってもらう本番系
scripts/local_stack.sh start prod
# http://127.0.0.1:3001

# 開発しながら確認する開発系
scripts/local_stack.sh start dev
# http://127.0.0.1:3002

# 状態確認
scripts/local_stack.sh status prod
scripts/local_stack.sh status dev

# 停止
scripts/local_stack.sh stop prod
scripts/local_stack.sh stop dev
```

本番系と開発系は別オリジン (`127.0.0.1:3001` / `127.0.0.1:3002`) なので、
ブラウザの `localStorage` に保存される現在ジョブIDも分離される。

同じ作業ツリーで両方を起動した場合も、prod は backend reload なし・frontend
静的 preview のため、起動中のユーザー検証は開発中のファイル変更から切り離される。
ただし、prod を再起動するとその時点の作業ツリーのコードを読む。

コード変更まで完全に分ける場合は Git worktree を使う。

```bash
git worktree add worktrees/prod-main main
git worktree add worktrees/dev-codex-chat-evidence-cards codex-chat-evidence-cards

cd worktrees/prod-main
scripts/local_stack.sh start prod

cd ../dev-codex-chat-evidence-cards
scripts/local_stack.sh start dev
```

dev で確認後に PR/merge し、prod 側 worktree で `git pull` してから
`scripts/local_stack.sh stop prod`、`scripts/local_stack.sh start prod` を実行する。

### サーバーIP配下で本番 / 開発をパス分離する

クライアントPCから `8001` / `8002` へ直接アクセスさせず、サーバーの標準HTTP(S)
ポートでパス分離する。プロキシ制約がある環境ではこの形を標準とする。

| URL | 用途 | 内部転送先 |
|---|---|---|
| `http://<server-ip>/` | 本番 Frontend | prod の静的ファイル |
| `http://<server-ip>/api/` | 本番 Backend API | `127.0.0.1:8001` |
| `http://<server-ip>/dev/` | 開発 Frontend | dev の静的ファイル |
| `http://<server-ip>/dev/api/` | 開発 Backend API | `127.0.0.1:8002` |

Frontend は相対パスで API を呼ぶ。本番ビルドは `/api`、開発ビルドは `/dev/api`
を指定する。

```bash
# prod frontend
cd frontend
NUXT_APP_BASE_URL=/ NUXT_PUBLIC_BACKEND_URL=/api pnpm generate

# dev frontend
cd frontend
NUXT_APP_BASE_URL=/dev/ NUXT_PUBLIC_BACKEND_URL=/dev/api pnpm generate
```

Nginx の設定例は `deploy/nginx-path-split.conf` を参照。Backend の `8001` / `8002`
はサーバー内部の待受にし、クライアントPCには公開しない。

サンプル `.xlsx` が手元になければ、ホーム画面の「サンプルをダウンロード」
ボタンから `inventory_sample.xlsx` (Input / Calc / Output の 3 シート構成、
名前付き範囲・条件付き書式入り) を取得して試せる。

## 開発コマンド

```bash
# Python
uv run pytest                # 全テスト (現在 400 件)
uv run ruff check            # lint
uv run ruff format           # フォーマット
uv run mypy core             # 型チェック (core/ のみ --strict)

# Frontend
cd frontend && pnpm dev      # 開発サーバ
cd frontend && pnpm generate # 静的書き出し (本番 SPA 用)
cd frontend && pnpm build    # SSR ビルド
```

## 環境変数

| 変数 | 既定 | 説明 |
|---|---|---|
| `JOBS_DIR` | `./jobs` | アップロードファイル / 抽出結果の置き場所 |
| `LLM_BASE_URL` | (未設定) | 社内 LLM API のベース URL (OpenAI 互換) |
| `LLM_API_KEY` | (未設定) | LLM API キー |
| `LLM_MODEL` | (未設定) | LLM のデフォルトモデル名 |
| `LLM_MODEL_PRO` | `LLM_MODEL` と同じ | チャット用モデル (キャッシュ重視) |
| `LLM_MODEL_FAST` | `LLM_MODEL` と同じ | 注釈バッチ用モデル (大量呼び出し) |
| `CORS_ALLOW_ORIGINS` | `localhost:3001,3000` 等 | カンマ区切りで上書き可 |
| `CHAT_HISTORY_LIMIT_PAIRS` | `10` | LLM 文脈に乗せる直近往復数 |
| `NUXT_PUBLIC_BACKEND_URL` | `/api` | フロントが叩く Backend URL。本番は `/api`、パス分離devは `/dev/api` |
| `NUXT_APP_BASE_URL` | `/` | Frontend の配信ベースパス。パス分離devは `/dev/` |
| `NUXT_PORT` | `3001` | フロント dev サーバの待受ポート |

LLM 系の環境変数が未設定の場合、Backend は `MockLLMClient` で起動する
(チャットは決まり文句しか返さないが、抽出 / 設計書生成 / 図解は動く)。

## 制限

- `.xls` (旧バイナリ形式) は VBA のみ抽出可能。openpyxl 非対応のため
  シート構造は抽出されない
- LLM 呼び出しは社内 LLM API (OpenAI 互換) のみ対応。外部クラウドへの
  データ送信は仕様上禁止 (SPEC §1.3)
- 想定上限: 1 ファイル 50MB / 5000 行程度。それを超える場合は警告ログ。
  上限は `MAX_UPLOAD_BYTES` 環境変数で上書き可能
- VBA のパスワード保護プロジェクトは抽出不可 (olevba の制約)
- Power Query は接続定義・出力先の棚卸しが中心。M コード本文や認証情報の
  解析は初期対応では対象外で、接続文字列内の秘匿値はマスクする

## 本番デプロイ (閉鎖ネットワーク向け)

本番環境がインターネットに出られない前提でフロントを構成している:

- **アイコン**: `@iconify-json/lucide` をビルドに焼き込み、`fallbackToApi: false`
  で `api.iconify.design` への runtime fetch を完全に遮断
- **フォント**: `@fontsource-variable/inter` / `@fontsource-variable/jetbrains-mono`
  を `node_modules` から `@import` し、Google Fonts / Bunny / jsDelivr 非依存
- **テレメトリ**: `NUXT_TELEMETRY_DISABLED=1` を推奨

デプロイ手順 (開発機 → 閉鎖網):

```bash
# 1. 開発機 (ネット可) でビルド
cd frontend
pnpm install
NUXT_APP_BASE_URL=/ NUXT_PUBLIC_BACKEND_URL=/api pnpm generate

# 2. 出力された .output/public/ を閉鎖網ホストへ転送し、
#    任意の静的配信サーバ (Nginx / IIS / `python -m http.server` 等) で配信
```

Backend (Python) 側は `uv sync` 済みの環境で `uv run uvicorn backend.main:app`
を直接起動する。フロントの SPA とは別プロセス。クライアントPCには Backend
ポートを公開せず、reverse proxy で `/api` から内部 Backend へ転送する。

## ディレクトリ構成 (概略)

```
excel-spec-tool/
├── core/             # 純 Python: 抽出 / 参照 / 設計書 / 図解
├── backend/          # FastAPI ルート + ストレージ + LLM クライアント
├── frontend/         # Nuxt 3 SPA
├── tests/            # pytest (core / backend)
├── scripts/          # ユーティリティ (サンプル生成等)
├── SPEC.md           # 仕様書 (正)
├── CLAUDE.md         # 開発エージェント向け作業ルール
└── pyproject.toml
```

詳細な責務分担・モジュール構成は [SPEC.md](./SPEC.md) を参照。
