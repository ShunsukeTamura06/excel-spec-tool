# Excel改修支援ツール

VBA/数式/参照関係を含む `.xlsm` `.xls` 業務ツールの統合設計書を生成し、改修対話を支援するWebアプリ。

詳細仕様は [SPEC.md](./SPEC.md) を、開発ルールは [CLAUDE.md](./CLAUDE.md) を参照。

## セットアップ

[uv](https://docs.astral.sh/uv/) を使う。

```bash
# 依存と仮想環境をまとめて同期 (.venv/ を作成)
uv sync

# 開発用ツール (pytest, ruff, mypy) も含めて同期
uv sync --group dev
```

依存追加:

```bash
uv add <package>          # 本体依存
uv add --group dev <pkg>  # 開発用
```

## 起動 (開発時)

```bash
# Backend
uv run uvicorn backend.main:app --reload --port 8000

# Frontend (別ターミナル) — Nuxt 3
cd frontend && pnpm install && pnpm dev   # http://localhost:3000
```

## 開発コマンド

```bash
uv run pytest
uv run ruff check
uv run ruff format
uv run mypy core
```

## 制限

- `.xls` (旧バイナリ形式) は VBA のみ抽出可能。シート構造は抽出されない (openpyxl 非対応)
- LLM 呼び出しは社内LLM API (OpenAI互換) のみ対応。外部クラウド送信は不可
