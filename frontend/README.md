# Excelツール改修支援AI — Frontend (Nuxt 3)

Vue 3 / Nuxt 3 / TypeScript / Nuxt UI / Vue Flow / Pinia による SPA。

## 開発

```bash
# 初回のみ
corepack enable

# 依存インストール + Nuxt 準備
pnpm install

# 開発サーバ起動 (http://localhost:3001)
pnpm dev
```

Backend (FastAPI) は別ターミナルで `uv run uvicorn backend.main:app --reload --port 8001` を起動しておくこと。

## ローカル環境の分離

リポジトリルートの `scripts/local_stack.sh` を使うと、本番系と開発系を
別ポート・別ジョブ保存先で同時起動できる。

```bash
# 本番系: Frontend 3001 / Backend 8001 / runtime/prod/jobs
../scripts/local_stack.sh start prod

# 開発系: Frontend 3002 / Backend 8002 / runtime/dev/jobs
../scripts/local_stack.sh start dev
```

## 環境変数

| 変数 | 既定 | 用途 |
|---|---|---|
| `NUXT_PUBLIC_BACKEND_URL` | `http://localhost:8001` | Backend ベース URL |
| `NUXT_PORT` | `3001` | dev サーバの待受ポート |

## ビルド

```bash
pnpm build      # SSR ビルド
pnpm generate   # 静的書き出し (SPA)
pnpm preview    # ビルド成果物プレビュー
```

## ディレクトリ構成

```
frontend/
├── app.vue              # ルート. UApp で Nuxt UI を初期化
├── nuxt.config.ts       # SPA モード設定 + モジュール登録
├── assets/css/main.css  # Tailwind v4 + テーマトークン + Vue Flow スタイル
├── layouts/             # 共通レイアウト (sidebar + content)
├── pages/               # ルーティング
│   ├── index.vue        # ホーム (ジョブ一覧 + アップロード)
│   ├── spec/[jobId].vue # 設計書ページ (タブ構成)
│   └── chat/[jobId].vue # チャットページ
├── components/          # 再利用 UI 部品
├── composables/         # useBackend (typed API)
├── stores/              # Pinia
└── types/               # Backend Pydantic と対応する TS 型
```
