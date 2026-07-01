# COM / テスト化 実現可能性スパイク

`docs/VISION.ja.md` の最大の死因候補 ——「現状挙動をテスト化できるか」「企業実機で
Excel COM / VBIDE が止まらないか」—— を **持ち帰り1回** で判定するための診断スクリプト。

このPC (Mac/Linux・Excel なし) では動きません。**会社端末 (Windows + Excel)** で実行します。

## これは何を確かめるのか

| 診断ステップ | 何が分かるか |
|---|---|
| `launch_excel` / `open_workbook` | ヘッドレス (画面非表示・ダイアログ抑制) で開けるか |
| `snapshot_baseline` | 全シートの計算結果を取得できるか (= 回帰テストの土台) |
| `determinism_recalc` | 2回再計算して値が一致するか。**非決定セル (揮発関数等) の検出** |
| `blast_radius` (任意) | 入力を1つ変えると、どのセルに波及するか |
| `vbide_access` | **VBA を VBIDE 経由で読めるか** (GPO で `AccessVBOM` が無効化されていないか) |
| `run_macro` (任意) | 任意マクロを実行できるか |
| `accessvbom_registry` | レジストリ上の VBA 信頼設定 (1=許可 / 0=禁止) |

## 実行手順 (会社端末)

1. このリポジトリを pull する。
2. 対象の Excel ツール (.xlsm 推奨) を用意する。**原本は変更されません** (常に一時コピーを開く)。
3. 1コマンドで実行 (uv が pywin32 をその場で用意する。本体の依存には追加しない):

```powershell
uv run --with pywin32 python spikes/com_probe/probe.py --workbook "C:\path\to\tool.xlsm"
```

波及やマクロも見る場合 (任意):

```powershell
uv run --with pywin32 python spikes/com_probe/probe.py --workbook "C:\path\to\tool.xlsm" `
  --input-cell "設定!B5=0.10" `
  --run-macro "Module1.RecalcAll"
```

4. 実行後、`spike_out\spike_bundle_<日時>.zip` が出力される。**この zip を1つだけこのPCに持ち帰る。**

## オプション

| フラグ | 意味 |
|---|---|
| `--workbook PATH` | 診断対象 (必須) |
| `--input-cell "Sheet!Addr=値"` | 入力を1つ書き換えて波及を観測 |
| `--run-macro "名前"` | 指定マクロを実行 |
| `--enable-events` | 入力/マクロ時にイベントマクロを有効化 (既定は無効=暴発防止) |
| `--out-dir DIR` | 出力先 (既定 `./spike_out`) |

## 持ち帰り物のプライバシー

- バンドルには **セルの生値を含めません**。値は短いハッシュにして「同じ/違う」だけ判定し、
  レポートにはセルアドレス・変化の有無・件数のみ残します。
- トークン/パスワード等のシークレットも記録しません。
- 含まれるのは `report.json` (診断結果) と `run.log` (実行ログ) の2点のみ。

## バンドルの読み方 (持ち帰り後)

- `report.json` の各 `steps[].ok` が `true` なら成功、`false` なら `error` に原因。
- `determinism_recalc.detail.nondeterministic_cell_count` が大きい → 揮発関数/外部依存が多く、
  入力固定なしでは安定テストにならない (テスト化の設計に影響)。
- `vbide_access.ok = false` かつ `accessvbom_registry.access_vbom = 0` → GPO で VBA 操作が
  封じられている。S3 (VBA 自律修正) の前提が崩れるので、回避策 (運用での信頼許可可否) を要検討。

## よくある詰まり

| 症状 | 見立て / 対処 |
|---|---|
| `vbide_access` が失敗 | `AccessVBOM=0`。Excel「マクロのセキュリティ」で信頼を許可できるか (GPO 次第) を確認 |
| 開いた直後にハングする | リンク更新/マクロ警告/保護ビュー等のモーダルダイアログ。`report.json` が途中で切れていたらこれ |
| `spike_out` にゾンビ Excel が残る | スクリプトは終了時に強制終了を試みるが、残ったらタスクマネージャで `EXCEL.EXE` を終了 |
| `uv` が無い | 会社端末に uv を導入するか、`pip install pywin32` 後に `python spikes/com_probe/probe.py ...` |

## 合否の目安 (MVP 受け入れベンチの前段)

- ヘッドレスで開けて、スナップショットが取れ、`determinism_recalc` が現実的な件数に収まる
  → S1/S2 の検証ループは成立する見込み。
- `vbide_access` が通る → S3 (VBA) も射程内。
- どれかが恒常的に失敗 → 設計を見直す (例: VBA を諦め数式レベルに限定、等)。
