"""S2 需要調査 (docs/s2_demand_survey.ja.md) の記入済み CSV を集計する.

安全 op (名前定義・数式参照・セル値等) でカバーできる改修依頼の割合を出し、
EditPlan への投資判断材料にする。実データを含む CSV は git に含めないこと
(docs/s2_demand_survey.ja.md 参照)。

使い方:
    uv run python scripts/tally_demand_survey.py runtime/s2_demand_survey/data.csv
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter
from pathlib import Path

_CATEGORY_LABELS: dict[str, str] = {
    "A": "名前定義の参照先変更",
    "B": "数式の参照先変更",
    "C": "数式範囲の拡張・縮小",
    "D": "セル固定値の変更",
    "E": "VBA中の参照先・定数のみの変更",
    "F": "VBAロジックの変更",
    "G": "シート/列/行の追加・削除等の構造変更",
    "H": "外部接続・ピボット・グラフの設定変更",
    "I": "書式・表示形式のみの変更",
    "J": "その他・複合",
}

_REQUIRED_COLUMNS = {"id", "カテゴリ", "安全opで対応可能か"}
_COVERAGE_WEIGHTS: dict[str, float] = {"Yes": 1.0, "一部": 0.5, "No": 0.0}
_HYPOTHESIS_LABEL = "仮説"


def _load_rows(csv_path: Path) -> list[dict[str, str]]:
    """CSV を読み込み、必須列の存在を検証してから行のリストを返す.

    Args:
        csv_path: 記入済み調査 CSV のパス.

    Returns:
        各行を dict にしたリスト.

    Raises:
        ValueError: 必須列が欠けている場合.
    """
    with csv_path.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        fieldnames = set(reader.fieldnames or [])
        missing = _REQUIRED_COLUMNS - fieldnames
        if missing:
            raise ValueError(f"CSV に必須列がありません: {sorted(missing)}")
        return list(reader)


def _coverage_pct(rows: list[dict[str, str]], *, weighted: bool) -> float:
    """安全opカバー率を計算する.

    Args:
        rows: 集計対象の行.
        weighted: True なら「一部」を 0.5 件として加重集計する
            (codex レビュー: 「一部」を満額カウントすると楽観的に出るため).

    Returns:
        0 件なら 0.0、それ以外はカバー率 (%).
    """
    if not rows:
        return 0.0
    total = 0.0
    for row in rows:
        value = row["安全opで対応可能か"].strip()
        total += (
            _COVERAGE_WEIGHTS.get(value, 0.0)
            if weighted
            else (1.0 if value in ("Yes", "一部") else 0.0)
        )
    return 100 * total / len(rows)


def _print_report(rows: list[dict[str, str]]) -> None:
    """カテゴリ別件数と安全 op カバー率をコンソールに出力する."""
    total = len(rows)
    if total == 0:
        print("0 件です。docs/s2_demand_survey.ja.md の手順で記入してください。")
        return

    category_counts = Counter(row["カテゴリ"].strip() for row in rows)
    coverage_counts = Counter(row["安全opで対応可能か"].strip() for row in rows)
    has_evidence_column = "根拠" in rows[0]

    print(f"総件数: {total}\n")

    print("## カテゴリ別件数")
    for code in sorted(category_counts, key=lambda c: -category_counts[c]):
        label = _CATEGORY_LABELS.get(code, "(未定義コード)")
        count = category_counts[code]
        pct = 100 * count / total
        print(f"  {code} ({label}): {count} 件 ({pct:.0f}%)")

    unknown_codes = sorted(set(category_counts) - set(_CATEGORY_LABELS))
    if unknown_codes:
        print(f"\n  警告: 未定義のカテゴリコードがあります: {unknown_codes}")

    print("\n## 安全opカバー率")
    for label in ("Yes", "一部", "No"):
        count = coverage_counts.get(label, 0)
        pct = 100 * count / total
        print(f"  {label}: {count} 件 ({pct:.0f}%)")

    other_labels = sorted(set(coverage_counts) - {"Yes", "一部", "No"})
    if other_labels:
        print(
            f"\n  警告: 想定外の値があります (Yes/一部/No のいずれかにしてください): {other_labels}"
        )

    print(f"\n安全opで対応可能 (Yes+一部、単純集計): {_coverage_pct(rows, weighted=False):.0f}%")
    print(f"安全opで対応可能 (加重集計、一部=0.5件換算): {_coverage_pct(rows, weighted=True):.0f}%")
    print("  → 「一部」を満額カウントすると楽観的に出るため、判断には加重集計を優先すること。")

    if has_evidence_column:
        hypothesis_rows = [r for r in rows if r.get("根拠", "").strip() == _HYPOTHESIS_LABEL]
        real_rows = [r for r in rows if r.get("根拠", "").strip() != _HYPOTHESIS_LABEL]
        print(f"\n## 根拠の内訳 (実データ vs {_HYPOTHESIS_LABEL})")
        print(f"  実データ (実体験/ヒアリング等): {len(real_rows)} 件")
        print(f"  {_HYPOTHESIS_LABEL}: {len(hypothesis_rows)} 件")
        if real_rows:
            print(
                f"  実データのみのカバー率 (加重): {_coverage_pct(real_rows, weighted=True):.0f}%"
                f"  ← 判断はこちらを優先する"
            )
        else:
            print("  警告: 実データが0件です。仮説だけの集計は判断材料として使わないこと。")
        if hypothesis_rows and len(real_rows) < 20:
            print(
                f"  警告: 実データが {len(real_rows)} 件しかありません"
                "(VISION §6.7 の目安は20〜50件)。仮説の混入率が高いほど結果の信頼度は下がる。"
            )
    else:
        print(
            "\n  警告: 「根拠」列がありません。実データと仮説の区別ができないため、"
            "上のカバー率をそのまま判断材料にしないこと。"
        )

    print("\n解釈の目安は docs/s2_demand_survey.ja.md 「結果の解釈」を参照。")


def main() -> int:
    """CLI エントリポイント."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("csv_path", type=Path, help="記入済み調査 CSV のパス")
    args = parser.parse_args()

    if not args.csv_path.exists():
        print(f"error: ファイルが見つかりません: {args.csv_path}", file=sys.stderr)
        return 1

    try:
        rows = _load_rows(args.csv_path)
    except ValueError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1

    _print_report(rows)
    return 0


if __name__ == "__main__":
    sys.exit(main())
