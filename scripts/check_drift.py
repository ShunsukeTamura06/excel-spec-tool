"""xlblueprint の整合性ガード — 禁止文字列・構造の機械的検証.

CI と pre-commit から呼ばれる。違反が 1 件でもあれば exit 1。

詳細な「なぜ禁止か」は CLAUDE.md §11 (Invariants) を参照。本スクリプトは
その規約を機械的に保証するためのもので、規約自体ではない。

使い方:
    uv run python scripts/check_drift.py            # 違反一覧 + exit 1 or 0
    uv run python scripts/check_drift.py --explain  # 各ルールの理由も表示

ホワイトリスト:
    歴史的記録ファイル (docs/OSS_LAUNCH_PLAN.md / docs/SPEC.ja.md /
    README.ja.md / CHANGELOG / git に対する自己参照) は引用の文脈で旧用語が
    出ても通す。本スクリプトの ALLOW_FILES_PER_RULE に明示する。
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]

# .gitignore に乗っている / バイナリ / 第三者 / 生成物 は最初から除外する.
EXCLUDED_DIRS = {
    ".git",
    ".github",  # workflow 内の文字列言及は OK にする (CI 自体)
    ".venv",
    "venv",
    "node_modules",
    ".nuxt",
    ".output",
    ".pytest_cache",
    ".mypy_cache",
    ".ruff_cache",
    ".uv-cache",
    "__pycache__",
    "jobs",
    "dist",
    "build",
    "htmlcov",
    "xlblueprint.egg-info",
    "excel_spec_tool.egg-info",
    ".claude",
}

# テキストとして scan する拡張子. それ以外は無視 (バイナリ / lockfile 等).
TEXT_EXTS = {
    ".py",
    ".vue",
    ".ts",
    ".js",
    ".tsx",
    ".jsx",
    ".md",
    ".mdx",
    ".yml",
    ".yaml",
    ".toml",
    ".cfg",
    ".ini",
    ".sh",
    ".ps1",
    ".cmd",
    ".bas",
    ".cls",
    ".css",
    ".html",
}

# 個別の例外的ファイル名 (拡張子無視).
ALWAYS_EXCLUDE_FILES = {
    "LICENSE",
    "uv.lock",
    "pnpm-lock.yaml",
}


@dataclass(frozen=True)
class Rule:
    """1 件の禁止ルール."""

    id: str
    pattern: re.Pattern[str]
    message: str
    why: str
    # このパス (REPO_ROOT 基準) は当該ルールについてはホワイトリスト扱い.
    allow_paths: frozenset[str] = field(default_factory=frozenset)


# ----------------------------------------------------------------- ルール定義 ----

# Note: 「歴史的記録 / 政策文書 / 自己定義ファイル」では引用として旧名や禁止
# 文字列がそのまま登場するため、ルール適用から除外する。
HISTORICAL_DOCS = frozenset(
    {
        "CLAUDE.md",  # Invariants 表に禁止文字列を列挙する政策文書
        "AGENTS.md",  # CLAUDE.md の他エージェント向けミラー (同じ Invariants 表を含む)
        "docs/OSS_LAUNCH_PLAN.md",  # 過去フェーズの記録に旧名が引用として残る
        "scripts/check_drift.py",  # ルール定義自身
        "CHANGELOG.md",  # 将来用意
    }
)


RULES: list[Rule] = [
    Rule(
        id="brand-jp-name",
        pattern=re.compile(r"Excelツール改修支援AI"),
        message="旧日本語ブランド名「Excelツール改修支援AI」を使わないでください。"
        "現行ブランド名は xlblueprint です。",
        why="ブランドは 1 つに統一する (Phase 0 決定)。二重表記は UI / docs の情報密度を下げる。",
        allow_paths=HISTORICAL_DOCS,
    ),
    Rule(
        id="legacy-package-name",
        pattern=re.compile(r"\bexcel[-_]spec[-_]tool\b"),
        message="旧パッケージ名 'excel-spec-tool' / 'excel_spec_tool' を新規参照しないでください。"
        "パッケージ名は xlblueprint です。",
        why="PyPI / GitHub の正式名を一本化する。git の origin URL のみ "
        "GitHub リネーム前は残るが、コード・docs での参照は禁止。",
        allow_paths=HISTORICAL_DOCS,
    ),
    Rule(
        id="company-internal-shanai",
        pattern=re.compile(r"社内\s*(?:LLM|AWS|ネットワーク|プロキシ|システム)"),
        message="「社内 LLM」「社内 AWS」等の社内固有記述を使わないでください。"
        "OpenAI 互換 / セルフホスト LLM 等の中立表現にしてください。",
        why="OSS リポでは特定組織の運用前提を出さない。"
        "ユーザーは Ollama / OpenAI / 任意のセルフホストを選べる設計。",
        allow_paths=HISTORICAL_DOCS,
    ),
    Rule(
        id="developer-personal-name",
        # "Shun" の直後に日本語助詞 (が/は/の/に/を/から/まで) または句読点が
        # 来る場合のみマッチ. 間に半角スペース 1 個までは許容。
        # "Shunsuke" / "shunsuke.tamura06" / camelCase 識別子はマッチさせない
        # (lookbehind で英字直前を弾く + 後続条件で日本語限定)。
        pattern=re.compile(r"(?<![A-Za-z])Shun\s?(?:が|は|の|に|を|から|まで|、|。)"),
        message="個人名「Shun」の "
        "「Shun が〜」式の引用は使わないでください。`the maintainer` / "
        "`@ShunsukeTamura06` 等に書き換えてください。",
        why="OSS では特定個人を実装メモに残さない。"
        "メンテナー記述は LICENSE / README の Authors 欄で十分。",
        allow_paths=HISTORICAL_DOCS,
    ),
    Rule(
        id="legacy-spec-path-root",
        # 行頭 SPEC.md (./SPEC.md, リンク `[...](./SPEC.md)` も検出).
        # ファイル先頭の `[SPEC]` 等は誤検出を避けるため `.md` 拡張子必須。
        # 既存のパス `docs/SPEC.ja.md` には反応しないよう lookbehind で除外。
        pattern=re.compile(r"(?<!\.ja)(?<!/)(?<!docs/)SPEC\.md\b"),
        message="ルート直下の SPEC.md は廃止しました。docs/SPEC.ja.md を参照してください。",
        why="日本語仕様は docs/SPEC.ja.md、英語要約は docs/architecture.md に分離。",
        allow_paths=HISTORICAL_DOCS,
    ),
]


# ----------------------------------------------------------------- 走査本体 ----


def iter_text_files(root: Path) -> list[Path]:
    """検査対象のテキストファイル一覧を返す."""
    found: list[Path] = []
    for dirpath, dirnames, filenames in os.walk(root):
        # 除外ディレクトリを in-place で剪定 (os.walk の流儀)
        dirnames[:] = [d for d in dirnames if d not in EXCLUDED_DIRS]
        for name in filenames:
            if name in ALWAYS_EXCLUDE_FILES:
                continue
            ext = Path(name).suffix.lower()
            if ext and ext in TEXT_EXTS:
                found.append(Path(dirpath) / name)
            elif not ext and name in {"Dockerfile", "Makefile"}:
                # 拡張子無いがテキストとして扱いたいもの (将来用)
                found.append(Path(dirpath) / name)
    return found


@dataclass(frozen=True)
class Violation:
    rule_id: str
    path: Path
    line_no: int
    line: str
    message: str


def scan_file(path: Path, rules: list[Rule]) -> list[Violation]:
    """1 ファイルを scan して違反一覧を返す."""
    rel = path.relative_to(REPO_ROOT).as_posix()
    try:
        text = path.read_text(encoding="utf-8")
    except (UnicodeDecodeError, OSError):
        return []

    violations: list[Violation] = []
    for rule in rules:
        if rel in rule.allow_paths:
            continue
        for line_no, line in enumerate(text.splitlines(), start=1):
            if rule.pattern.search(line):
                violations.append(
                    Violation(
                        rule_id=rule.id,
                        path=path.relative_to(REPO_ROOT),
                        line_no=line_no,
                        line=line.strip()[:140],
                        message=rule.message,
                    )
                )
    return violations


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--explain",
        action="store_true",
        help="各ルールの why も表示する",
    )
    parser.add_argument(
        "--rules",
        nargs="*",
        help="特定の rule id だけを実行する (省略時は全件)",
    )
    args = parser.parse_args()

    rules = RULES
    if args.rules:
        wanted = set(args.rules)
        unknown = wanted - {r.id for r in RULES}
        if unknown:
            print(f"error: unknown rule id(s): {sorted(unknown)}", file=sys.stderr)
            return 2
        rules = [r for r in RULES if r.id in wanted]

    violations: list[Violation] = []
    for path in iter_text_files(REPO_ROOT):
        violations.extend(scan_file(path, rules))

    if not violations:
        print(f"OK -- {len(rules)} rule(s) checked, 0 violations.")
        return 0

    # 違反一覧 (rule_id でグループ化)
    grouped: dict[str, list[Violation]] = {}
    for v in violations:
        grouped.setdefault(v.rule_id, []).append(v)

    rule_by_id = {r.id: r for r in rules}
    print(f"FAIL -- {len(violations)} violation(s) across {len(grouped)} rule(s):\n")
    for rule_id, vs in grouped.items():
        rule = rule_by_id[rule_id]
        print(f"[{rule_id}] {rule.message}")
        if args.explain:
            print(f"   why: {rule.why}")
        for v in vs:
            print(f"   {v.path}:{v.line_no}: {v.line}")
        print()

    print(f"Total: {len(violations)} violation(s). See CLAUDE.md §11 (Invariants).")
    return 1


if __name__ == "__main__":
    sys.exit(main())
