# CLAUDE.md - 作業ルール

このリポジトリは Claude Code を主開発エージェントとして進めるプロジェクト。
SPEC.md を仕様の正とし、本ファイルは Claude Code の進め方を規定する。

## 0. 大原則

- **SPEC.md を正とする**。SPEC.md と矛盾する実装はしない
- **依存方向を守る**: `frontend → backend → core` の単方向のみ
  - core は backend / frontend を import しない
  - backend は frontend を import しない
- **作業はステップ単位で進め、各ステップ完了時にユーザー確認を求める**
- 不明点・SPEC.md の曖昧さを発見したら**勝手に解釈せず質問する**

## 1. 環境

- Python 3.10+
- パッケージ管理: `pip` + `venv`（uvが使えるなら uv 推奨）
- フォーマッタ: `ruff format`
- リンタ: `ruff check`
- 型チェック: `mypy --strict` を core/ に対して
- テスト: `pytest`

セットアップ:
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
pip install -e .
```

## 2. 実装順序

下記の順で進める。各ステップ完了時にユーザーに報告し、次に進む承認を得る。

1. **プロジェクト初期化** — `pyproject.toml`, `requirements.txt`, ディレクトリ作成
2. **`core/models.py`** — Pydanticモデル定義 + テスト
3. **`core/extractors/vba.py`** — olevba ラッパー + テスト
4. **`core/extractors/workbook.py`** — openpyxl ラッパー + テスト
5. **`core/reference_index.py`** — 参照インデックス構築 + テスト
6. **`core/spec_generator.py`** — 設計書Markdown生成 + テスト
7. **`backend/storage.py`** — ローカルファイル永続化 + テスト
8. **`backend/llm_client.py`** — モック実装 + インターフェース
9. **`backend/main.py` + `backend/routes/`** — FastAPIエンドポイント
10. **`frontend/api_client.py` + `frontend/app.py`** — Streamlit UI

各ステップの完了基準:
- 該当機能のpytestが通る
- `ruff check` `ruff format --check` が通る
- core/ は `mypy --strict` が通る
- 動作確認手順をコミットメッセージに書く

## 3. やってよいこと / やってはいけないこと

### やってよい
- 自分でテストデータ（小さなxlsm）を作る
- SPEC.mdに書かれていない**実装の細部**は妥当な判断で進めてよい（コミットメッセージで明示）
- リファクタは対象ステップに閉じる範囲で

### やってはいけない
- SPEC.md にない**機能**を勝手に追加する
- 認証、Docker化、ユーザー管理を実装する（スコープ外）
- 外部APIに実際に接続する（社内LLM接続実装はShunが手動で追記する）
- 巨大なPRや一気に複数ステップを完了させる
- 既存ファイルを大幅に書き換える（影響範囲を聞く）

## 4. テストの書き方

- 各 `core/*.py` に対応する `tests/core/test_*.py` を必ず作る
- `tests/fixtures/` に小さな xlsm サンプルを置く（ダミーデータ可）
- backend のテストは `TestClient` を使い、`tmp_path` で `JOBS_DIR` を差し替える
- LLM呼び出しはモックする（`backend/llm_client.py` のモック実装を再利用）
- カバレッジ80%を目指す（厳密ルールではない）

### 必須テストケース

#### `core/extractors/vba.py`
- 単純な `Sub` / `Function` / `Property` の検出
- 複数モジュール
- VBAなしxlsm（空リスト返却）
- パスワード保護VBA（最初はskipしてよいがTODO残す）

#### `core/extractors/workbook.py`
- 数式セルの抽出（`=SUM(A1:A10)` レベル）
- 複数シートまたぎ参照（`Sheet2!A1`）
- 名前付き範囲の抽出
- 条件付き書式の抽出
- `.xls` を渡したときに空シートで返ること

#### `core/reference_index.py`
- 数式 `=SUMIF(Input!A:A, A2, Input!E:E)` から3つの参照を抽出
- VBA `Range("A1:J100")`, `Worksheets("Calc").Range("A1")`, `Sheets("Calc").Cells(2, 8)` の3パターン
- 該当なしクエリで空配列が返る

## 5. コーディング規約

- 関数には型ヒント必須
- public関数にはdocstring（Googleスタイル）
- ログは `logging` モジュールを使う。`print` は禁止
- 例外は素通しせず、適切な層でwrap（Coreは独自例外、Backendは HTTPException に変換）
- 設定値はハードコードせず環境変数 or デフォルト引数

例外クラスは `core/exceptions.py` に定義:
```python
class CoreError(Exception): ...
class ExtractionError(CoreError): ...
class UnsupportedFormatError(CoreError): ...
```

## 6. コミット

- 1コミット = 1論理的変更
- メッセージ規約: Conventional Commits
  - `feat: add VBA extractor`
  - `test: add tests for reference index`
  - `fix: handle empty workbook`
  - `chore: update requirements`
- ステップ完了時には `feat(step-N): {description}` 形式

## 7. ユーザーへの確認・質問の仕方

以下の場合は実装を止めて確認:
- SPEC.md と矛盾する状況を発見した
- SPEC.md に書かれていない機能を実装したくなった
- 外部APIに実際に接続する必要が出てきた
- 大きな設計判断（DBを入れるか、別ライブラリに切り替えるか等）

確認のフォーマット:
```
[確認] {状況}
- 選択肢A: ...
- 選択肢B: ...
推奨: A（理由: ...）
進めてよいですか？
```

## 8. 既知の難所

| 箇所 | 難所 | 対処方針 |
|---|---|---|
| `extract_vba` | パスワード保護VBA | 初版はskip、TODOコメント |
| `extract_workbook` | `.xls` 非対応 | `Workbook.sheets=[]` で返し、警告ログ |
| `extract_workbook` | 外部リンク取得 | `wb._external_links` がprivate APIなので `try/except` で包む |
| `reference_index` | VBA正規表現の捕捉漏れ | 完璧を目指さない。テストで最低限を担保 |
| `streamlit` UploadedFile → backend | バイナリ送信 | `httpx.post(..., files={"file": (name, bytes, mime)})` |
| 大きなxlsm | メモリ・時間 | 初版は1ファイル50MB、5000行を上限と想定。超えたら警告 |

## 9. 質問テンプレート

```
[質問] {タイトル}
状況: {何をしようとしていて、何に詰まったか}
SPEC.md該当箇所: {あれば}
選択肢:
  A) ...
  B) ...
  C) ...
私の推奨: {} (理由: ...)
```

## 10. 完了の定義 (Definition of Done)

- [ ] 実装ファイルが追加されている
- [ ] 対応するテストファイルがあり、pytestが通る
- [ ] `ruff check` が通る
- [ ] core/ なら `mypy --strict` が通る
- [ ] サンプル入力での動作確認ログをコミットメッセージに含める
- [ ] SPEC.md の該当箇所と矛盾がない
- [ ] ユーザーに完了報告し、次ステップへの承認を得た
