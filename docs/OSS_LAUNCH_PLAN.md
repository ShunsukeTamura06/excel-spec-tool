# OSS 公開計画 (Excelツール改修支援AI)

> 作業中ドラフト。ユーザー決定待ちの項目があるので Phase 0 から進める。
> 各 Phase は独立した PR にできるサイズを意識する。

## 現状サマリ (記録時点: 2026-05-28)

- リポジトリは GitHub (`ShunsukeTamura06/excel-spec-tool`) に存在するが Private 想定の作り
- **ライセンス: `Proprietary`** (pyproject.toml:11) — 公開ブロッカー
- README / SPEC / CLAUDE すべて日本語のみ
- LICENSE / CONTRIBUTING / CODE_OF_CONDUCT / SECURITY.md / ISSUE_TEMPLATE 未整備
- CI ワークフロー無し (`.github/workflows/` 未作成)
- 製品名「Excelツール改修支援AI」は日本語のみ
- CLAUDE.md に「社内LLM」「Shun」記載あり (公開時に一般化が必要)

---

## Phase 0: 公開可否の判断 (ユーザー決定が必須)

ここを決めないと先に進めない。

1. **ライセンス選定** — MIT / Apache-2.0 / BSL / AGPL のどれにするか
2. **製品名 (英名)** — 例: `excel-spec-tool`, `xlsm-spec`, `ExcelInsight`, `VBADocAI` 等
3. **対応言語ポリシー** — README は英語を主・日本語を従？ それとも併記？
4. **メンテ体制** — 個人主体 / 募集する / Issue は日本語可？
5. **コードオーナーシップ / コミット履歴のクリーニング** — 履歴に社内情報や個人情報が混入していないか確認

## Phase 1: 法務・公開最小要件 (公開ブロッカー)

- [ ] `LICENSE` ファイル追加 (選定したものに対応)
- [ ] `pyproject.toml` の `license` を更新、`authors` の `"Shun"` を本名 or GitHub handle に
- [ ] `frontend/package.json` にも `license` フィールド追加
- [ ] 全ソースから社内情報・PII・APIキー・社内URLが残っていないか grep で再確認 (`LLM_BASE_URL` 既定値など)
- [ ] CLAUDE.md の「社内LLM」「Shun」記述を一般化、もしくは内部用と公開用を分離
- [ ] サンプル xlsm / ジョブ成果物に実データが混じっていないか確認 (`jobs/`, `tests/fixtures/`)

## Phase 1.5: UX 第一波 — 第一印象の事故をなくす

OSS 公開時の "first touch" を破壊する 3 件を先に潰す。
詳細は 2026-05-28 の実機レビュー所感を参照。

- [ ] **mock LLM 注釈の raw prompt leak を撲滅** (`backend/llm_client.py:248`)
  - 現状: LLM 未設定で `MockLLMClient` を使うと `[mock annotation] prompt='...'` が
    設計書本文・シートカード・ダイアグラムノードに大量露出する
  - 対応: mock 時は注釈を空文字または明示シグナルにし、UI 側で
    「LLM 未設定」バッジ等の中立表示にする
- [ ] **設計書 (概要タブ) の MDC スタイル事故を直す**
  - H1 がページ H1 と重複 (`spec_generator.py` の `# 設計書: ...`)
  - 見出しがすべて下線リンク風スタイル (MDC の prose CSS が link 化している)
  - シート一覧テーブル列ヘッダが縦書き状態 (列幅・word-break)
- [ ] **ダイアグラムノードを「読める」内容に** (`frontend/components/SpecNode.vue`)
  - 現状: シート名 + 寸法 + mock annotation 切れ端のみ
  - 対応: 主要数式 1-2 件 / 最頻参照先 / 未解析リスクバッジ を出す
  - 合わせて MiniMap のダークテーマ追従と fitView padding 微調整

### 第二波 (公開後すぐ、1-3 日)

- [ ] **概要タブをインサイトダッシュボードに作り変える** (`SpecOverview.vue`)
  - 現状: spec.md の生 Markdown ダンプ。価値が伝わらない
  - 対応: TL;DR + シートランキング (最も参照される / 最も複雑) + 未解析リスク
    サマリを上に出し、詳細 Markdown は折りたたみへ
  - これが Show HN / Twitter デモ GIF の主役画面になる想定
- [ ] **プレビュー表に行番号と列文字** (`SheetExplorer.vue`)
  - 現状: 値だけ並ぶ。Excel ユーザーが「A4 を見て」と言えない
  - 対応: A B C... / 1 2 3... の固定ヘッダを出して座標で会話できるように
- [ ] **LLM 未設定時の onboarding カード**
  - 現状: 環境変数なしだと MockClient が静かに動き、チャットが意味のない応答を返す
  - 対応: ホーム / チャット空状態で「LLM 未接続 / 設定方法」カードを出す
  - Ollama / OpenAI 互換 API 接続の最短手順を提示

## Phase 2: 国際化 (英語ドキュメント)

- [ ] `README.md` を英語版に置換、日本語版は `README.ja.md` へ
  - Hero (1行価値訴求) / スクリーンショット2-3枚 / Quick Start / Features / Architecture / Roadmap / License
- [ ] `SPEC.md` は日本語のまま `docs/SPEC.ja.md` へ移し、英語要約版を `docs/architecture.md` として新規作成
- [ ] エラーメッセージ・UIテキストの i18n 対応有無を決める (やるなら `@nuxtjs/i18n`)
- [ ] バナー画像 or ロゴ (任意だが映える)

## Phase 3: コミュニティ整備

- [ ] `CONTRIBUTING.md` — セットアップ、PR フロー、Conventional Commits、テスト要件
- [ ] `CODE_OF_CONDUCT.md` — Contributor Covenant 2.1
- [ ] `SECURITY.md` — 脆弱性報告先・公開ポリシー (xlsm パーサが攻撃面なので必須級)
- [ ] `.github/ISSUE_TEMPLATE/` — bug / feature / question
- [ ] `.github/PULL_REQUEST_TEMPLATE.md`
- [ ] `.github/FUNDING.yml` (任意)

## Phase 4: 品質ゲート (CI/CD)

- [ ] `.github/workflows/ci.yml` — Python: pytest + ruff + mypy / Frontend: pnpm typecheck + build を PR トリガで
- [ ] OS マトリクス: Ubuntu + Windows
- [ ] `dependabot.yml` で `pip` / `npm` / `github-actions` の更新
- [ ] CodeQL もしくは ruff の security ルール追加
- [ ] バッジを README に貼る (build / license / python version)

## Phase 5: 配布・体験向上

- [ ] **Docker Compose** — `docker compose up` で backend + frontend が立つようにする
- [ ] スクリーンショット撮影 (設計書ビュー / 依存グラフ / チャット)
- [ ] 1分デモ GIF or 動画
- [ ] **ライブデモ** — Hugging Face Space / Render / Fly.io 等 (LLMキー扱いをどうするか別途検討)
- [ ] サンプル xlsm を `examples/` 配下に複数用意
- [ ] `CHANGELOG.md` + `v0.1.0` タグ + GitHub Release

## Phase 6: 露出・初動

- [ ] GitHub Topics 設定 (`excel`, `vba`, `openpyxl`, `fastapi`, `nuxt`, `llm`, `legacy-modernization`)
- [ ] Show HN / Reddit (`r/programming`, `r/excel`, `r/vba`) / Zenn / Qiita 告知記事
- [ ] `awesome-*` リスト系へPR (例: `awesome-fastapi`, `awesome-nuxt`)
- [ ] 訴求軸: 「Legacy Excel modernization with LLM」
