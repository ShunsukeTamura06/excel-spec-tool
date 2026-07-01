# xlsx 正規化 diff ノイズ除去 実現可能性スパイク

`docs/VISION.ja.md` §6 の前提3「xlsm 正規化 diff のノイズ除去
（calcChain/sharedStrings/style/rel id/pivot cache）が実用精度に届くか」を検証する。

**Windows / Excel は不要。このPC（Mac/Linux）でそのまま完結する。**

## 検証する仮説

xlsx の生 XML（zip 内パーツ）をそのまま diff すると、意味的な変更が無くても
保存のたびに `calcChain.xml` / `sharedStrings.xml` / `styles.xml` / `.rels` の
id 等が変化してノイズだらけになる。一方、既存の
`core.extractors.cells.extract_cells_to_sqlite`（セル単位の正規化抽出: シート・
座標・値・数式・表示形式）を経由して比較すれば、このノイズを素通りできるのでは
ないか。

## やること

1. リポジトリ同梱のサンプル `frontend/public/samples/retail_monthly_ops.xlsx`
   （`scripts/make_sample.py` で生成済みのもの）をベースラインとしてコピーする
2. **ケースA（ノイズのみ）**: 意味的な変更を一切加えず openpyxl で読み込んで保存し
   直すだけのファイルを作り、ベースラインと比較する
3. **ケースB（本物の変更1件）**: 1セルの数式だけを書き換えたファイルを作り、
   ベースラインと比較する
4. 各ケースについて「生 zip パーツのバイト単位 diff」と「`extract_cells_to_sqlite`
   経由のセル単位 diff」の両方を取り、結果を比較する

## 実行

```bash
uv run python spikes/xlsx_diff_noise/probe.py
```

`spike_out/report.json` に各ケースの結果と `verdict`（仮説が成立したかの機械判定）
が出力される。

## 結果の読み方

- `verdict.hypothesis_confirmed` が `true` なら:
  - ノイズのみのケースで生XML diffは変化を検出するが、セル単位diffは0件
  - 本物の変更1件のケースで、セル単位diffが**その1セルだけ**を指す
  - → 「正規化抽出済みの表現を比較すれば、生XMLの保存ノイズを気にする必要はない」
    という設計判断の裏付けになる

## 既知の限界（未検証事項）

- **openpyxl の保存ノイズしか見ていない**。VISION が本来懸念していたのは
  **Excel 本体で開いて保存し直したときの**ノイズ（calcChain 再構築・sharedStrings
  の頻度順並び替え・.rels id の大幅な振り直し等）で、openpyxl より激しい可能性が
  あった。→ **実機検証済み**（`spikes/com_probe/` の `excel_resave_noise` ステップ、
  `docs/VISION.ja.md` §6.3）。実業務 .xlsm で生XMLパーツは12箇所変化した（openpyxlの
  5〜6箇所より多い）が、計算値ベースでは既知の揮発セル以外に変化はなく、仮説は
  実機でも成立した。
- `extract_cells_to_sqlite` は**セルの値・数式・表示形式・結合範囲のみ**を見る。
  名前付き範囲・条件付き書式・データ検証・グラフ・ピボットテーブル・印刷範囲・
  マクロの副作用など、VISION §4.1 が挙げている他の差分対象は
  `core.extractors.workbook.extract_workbook()` 側の構造化データを別途比較する
  必要があり、本スパイクでは検証していない。
- 大規模ファイルでの性能（`extract_cells_to_sqlite` はワークブックを2回ロードする）
  は未計測。
