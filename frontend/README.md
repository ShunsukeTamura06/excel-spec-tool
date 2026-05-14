# Excel 改修支援ツール — Frontend (Nuxt 3)

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
