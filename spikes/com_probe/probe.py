"""Excel COM / テスト化 実現可能性スパイク.

会社端末 (Windows + Excel) で実行し、ビジョン (docs/VISION.ja.md) の
最大の死因候補を「持ち帰り1回」で診断する。検証する前提:

1. ヘッドレスで .xlsm/.xlsx を開けるか (ダイアログを抑制できるか)
2. 全シートの計算結果スナップショットを取得できるか (= 回帰テストの基盤)
3. 2回再計算して差分 → 非決定セル (揮発関数等) を検出できるか (= テスト化の決定性)
4. 入力セル書込→再計算→出力読取で「波及 (blast radius)」を観測できるか
5. VBIDE 経由で VBA モジュール/プロシージャを列挙できるか (GPO で禁止されていないか)
6. 任意マクロを実行できるか

結果は JSON レポートとログにまとめ、1つの zip バンドルに固める。

重要な制約:
- このPC (Mac/Linux) では実行不可。win32com (pywin32) は Windows 専用。
- 原本は決して書き換えない (常に一時コピーを開く)。
- 既定では業務データ (セルの生値) をバンドルに含めない。アドレス・変化の有無・
  値の型・ハッシュのみ記録する。生値が必要なときだけ --include-values を付ける。

実行例 (会社端末):
    uv run --with pywin32 python spikes/com_probe/probe.py --workbook "C:\\path\\tool.xlsm"
    uv run --with pywin32 python spikes/com_probe/probe.py --workbook tool.xlsm \\
        --input-cell "設定!B5=0.10" --run-macro "Module1.RecalcAll"
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import logging
import platform
import shutil
import sys
import tempfile
import time
import traceback
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

logger = logging.getLogger("com_probe")

# xlCalculation 列挙の逆引き (診断レポート用)
_CALC_MODE = {-4135: "manual", -4105: "automatic", 2: "semiautomatic"}
# 値ハッシュの桁数 (生値を出さずに「同じ/違う」を判定できれば十分)
_HASH_LEN = 12
# スナップショットするセル数の上限 (巨大ブックでのメモリ/時間暴走を防ぐ)
_MAX_CELLS = 300_000


@dataclass
class Step:
    """1つの診断ステップの結果."""

    name: str
    ok: bool | None = None
    elapsed_s: float = 0.0
    detail: dict[str, Any] = field(default_factory=dict)
    error: str | None = None

    def as_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "ok": self.ok,
            "elapsed_s": round(self.elapsed_s, 3),
            "detail": self.detail,
            "error": self.error,
        }


def _hash_value(value: Any) -> str:
    """セルの値を生のまま残さずに比較できるよう短いハッシュにする."""

    raw = repr(value).encode("utf-8", errors="replace")
    return hashlib.sha256(raw).hexdigest()[:_HASH_LEN]


def _col_letter(col: int) -> str:
    """1始まりの列番号を Excel の列記号に変換する."""

    s = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        s = chr(65 + rem) + s
    return s


def _snapshot(workbook: Any) -> dict[str, str]:
    """全シートの計算結果をハッシュ化したスナップショットを返す.

    Returns:
        "シート名!A1" -> 値ハッシュ の辞書。生値は保持しない。
    """

    snap: dict[str, str] = {}
    capped = False
    for ws in workbook.Worksheets:
        used = ws.UsedRange
        rows = used.Rows.Count
        cols = used.Columns.Count
        if rows * cols <= 0:
            continue
        first_row = used.Row
        first_col = used.Column
        values = used.Value2
        # 単一セルのときは Value2 がタプルにならないので正規化する
        if rows == 1 and cols == 1:
            values = ((values,),)
        for r in range(rows):
            for c in range(cols):
                if len(snap) >= _MAX_CELLS:
                    capped = True
                    break
                v = values[r][c]
                if v is None:
                    continue
                addr = f"{ws.Name}!{_col_letter(first_col + c)}{first_row + r}"
                snap[addr] = _hash_value(v)
            if capped:
                break
        if capped:
            break
    if capped:
        logger.warning("スナップショットが上限 %d セルで打ち切られました", _MAX_CELLS)
    return snap


def _diff(before: dict[str, str], after: dict[str, str]) -> list[str]:
    """2つのスナップショット間で値が変わったセルのアドレス一覧."""

    changed: list[str] = []
    for addr, h in after.items():
        if before.get(addr) != h:
            changed.append(addr)
    for addr in before:
        if addr not in after:
            changed.append(addr)
    return sorted(set(changed))


def _read_accessvbom() -> dict[str, Any]:
    """レジストリから「VBA プロジェクトへのアクセスを信頼する」設定を読む.

    AccessVBOM=1 でないと VBIDE 経由の VBA 操作はブロックされる (GPO で 0 固定の企業が多い)。
    """

    result: dict[str, Any] = {"checked_versions": [], "access_vbom": None}
    try:
        import winreg  # type: ignore
    except ImportError:
        result["error"] = "winreg 利用不可 (非 Windows)"
        return result
    for ver in ("16.0", "15.0", "14.0"):
        path = rf"Software\Microsoft\Office\{ver}\Excel\Security"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                val, _ = winreg.QueryValueEx(key, "AccessVBOM")
                result["checked_versions"].append({"version": ver, "AccessVBOM": val})
                if result["access_vbom"] is None:
                    result["access_vbom"] = int(val)
        except FileNotFoundError:
            continue
        except OSError as exc:  # noqa: PERF203 - 診断目的でそのまま記録
            result["checked_versions"].append({"version": ver, "error": str(exc)})
    return result


def _capture_env(app: Any) -> dict[str, Any]:
    """Excel / OS の環境情報を集める (決定性の前提条件になる)."""

    env: dict[str, Any] = {
        "python": sys.version.split()[0],
        "platform": platform.platform(),
    }
    try:
        env["excel_version"] = app.Version
        env["excel_build"] = str(app.Build)
        env["calculation_mode"] = _CALC_MODE.get(app.Calculation, app.Calculation)
        # International 定数: 1=xlCountryCode, 等。代表値だけ拾う。
        env["country_setting"] = app.International(1)
        env["decimal_separator"] = app.International(3)
    except Exception as exc:  # noqa: BLE001 - 環境差で落ちても診断は続ける
        env["env_error"] = str(exc)
    return env


def _log_step(step: Step) -> None:
    """ステップ結果をファイルログにも残す (ハング時の途中経過診断のため)."""

    level = logging.INFO if step.ok else logging.WARNING
    logger.log(level, "[%s] ok=%s (%.2fs) %s", step.name, step.ok, step.elapsed_s, step.error or "")


def _probe_vbide(workbook: Any, max_modules: int = 200) -> Step:
    """VBIDE 経由で VBA モジュール/プロシージャを列挙できるか試す."""

    step = Step("vbide_access")
    t0 = time.perf_counter()
    try:
        project = workbook.VBProject  # AccessVBOM 無効ならここで com_error
        modules = []
        for comp in project.VBComponents:
            code = comp.CodeModule
            lines = code.CountOfLines
            procs = []
            line_no = 1
            while line_no <= lines and len(procs) < 50:
                try:
                    proc_name = code.ProcOfLine(line_no, 0)  # 0=vbext_pk_Proc
                except Exception:  # noqa: BLE001
                    break
                if proc_name and proc_name not in procs:
                    procs.append(proc_name)
                line_no += 1
            modules.append(
                {"name": comp.Name, "type": comp.Type, "lines": lines, "procedures": procs}
            )
            if len(modules) >= max_modules:
                break
        step.ok = True
        step.detail = {"module_count": len(modules), "modules": modules}
    except Exception as exc:  # noqa: BLE001 - これ自体が重要な所見
        step.ok = False
        step.error = f"{type(exc).__name__}: {exc}"
        step.detail = {
            "hint": "AccessVBOM=1 (Excel: 開発タブ→マクロのセキュリティ→"
            "「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」) が必要。"
            "GPO で無効化されている可能性。"
        }
    step.elapsed_s = time.perf_counter() - t0
    return step


def run_probe(args: argparse.Namespace, report: dict[str, Any]) -> None:
    """COM を起動して一連の診断ステップを実行する (Windows 専用)."""

    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    steps: list[Step] = []
    report["steps"] = steps

    src = Path(args.workbook).expanduser().resolve()
    if not src.exists():
        raise FileNotFoundError(f"ワークブックが見つかりません: {src}")

    # 原本を絶対に触らないため一時コピーを開く
    tmp_dir = Path(tempfile.mkdtemp(prefix="com_probe_"))
    work_copy = tmp_dir / src.name
    shutil.copy2(src, work_copy)
    report["workbook"] = {"name": src.name, "suffix": src.suffix, "size_bytes": src.stat().st_size}

    pythoncom.CoInitialize()
    app = None
    app_pid: int | None = None
    workbook = None
    try:
        logger.info("Excel 起動中...")
        st = Step("launch_excel")
        t0 = time.perf_counter()
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.AskToUpdateLinks = False
        app.EnableEvents = False  # 起動時イベントマクロの暴発を防ぐ
        st.ok = True
        st.elapsed_s = time.perf_counter() - t0
        steps.append(st)
        _log_step(st)
        report["env"] = _capture_env(app)
        try:
            import win32process  # type: ignore

            _, app_pid = win32process.GetWindowThreadProcessId(app.Hwnd)
            report["excel_pid"] = app_pid
        except Exception:  # noqa: BLE001
            pass

        # 1) 開く
        logger.info("ワークブックを開いています: %s", work_copy)
        st = Step("open_workbook")
        t0 = time.perf_counter()
        workbook = app.Workbooks.Open(str(work_copy), UpdateLinks=0, ReadOnly=False)
        st.ok = True
        st.detail = {"sheets": workbook.Worksheets.Count}
        st.elapsed_s = time.perf_counter() - t0
        steps.append(st)
        _log_step(st)

        # 2) スナップショット (= 回帰テストの基盤)
        logger.info("ベーススナップショット取得中...")
        st = Step("snapshot_baseline")
        t0 = time.perf_counter()
        snap1 = _snapshot(workbook)
        st.ok = True
        st.detail = {"cell_count": len(snap1)}
        st.elapsed_s = time.perf_counter() - t0
        steps.append(st)
        _log_step(st)

        # 3) 2回再計算して決定性を確認
        logger.info("再計算して決定性を確認中...")
        st = Step("determinism_recalc")
        t0 = time.perf_counter()
        app.CalculateFull()
        snap2 = _snapshot(workbook)
        nondeterministic = _diff(snap1, snap2)
        st.ok = len(nondeterministic) == 0
        st.detail = {
            "nondeterministic_cell_count": len(nondeterministic),
            "sample": nondeterministic[:30],
            "note": "0 件が理想。多い場合は揮発関数 (NOW/RAND/OFFSET 等) や"
            "外部依存があり、入力固定なしでは安定したテストにならない。",
        }
        st.elapsed_s = time.perf_counter() - t0
        steps.append(st)
        _log_step(st)

        # 4) 入力書込→再計算→波及観測 (任意)
        if args.input_cell:
            logger.info("入力セルを書き換えて波及を観測中: %s", args.input_cell)
            st = Step("blast_radius")
            t0 = time.perf_counter()
            try:
                target, _, value = args.input_cell.partition("=")
                sheet_name, _, addr = target.partition("!")
                ws = workbook.Worksheets(sheet_name)
                try:
                    parsed: Any = float(value)
                except ValueError:
                    parsed = value
                app.EnableEvents = bool(args.enable_events)
                ws.Range(addr).Value2 = parsed
                app.CalculateFull()
                snap3 = _snapshot(workbook)
                changed = _diff(snap2, snap3)
                st.ok = True
                st.detail = {
                    "input": args.input_cell,
                    "events_enabled": bool(args.enable_events),
                    "changed_cell_count": len(changed),
                    "sample": changed[:50],
                }
            except Exception as exc:  # noqa: BLE001
                st.ok = False
                st.error = f"{type(exc).__name__}: {exc}"
            finally:
                app.EnableEvents = False
            st.elapsed_s = time.perf_counter() - t0
            steps.append(st)
            _log_step(st)

        # 5) VBIDE アクセス
        logger.info("VBIDE 経由で VBA アクセスを確認中...")
        vbide_step = _probe_vbide(workbook)
        steps.append(vbide_step)
        _log_step(vbide_step)

        # 6) マクロ実行 (任意)
        if args.run_macro:
            logger.info("マクロを実行中: %s", args.run_macro)
            st = Step("run_macro")
            t0 = time.perf_counter()
            try:
                app.EnableEvents = bool(args.enable_events)
                app.Run(args.run_macro)
                st.ok = True
                st.detail = {"macro": args.run_macro}
            except Exception as exc:  # noqa: BLE001
                st.ok = False
                st.error = f"{type(exc).__name__}: {exc}"
            finally:
                app.EnableEvents = False
            st.elapsed_s = time.perf_counter() - t0
            steps.append(st)
            _log_step(st)

        # オプション: 生値を含める場合のみ (既定は含めない)
        if args.include_values:
            report["warning_values_included"] = (
                "業務データ (セル値ハッシュではなく生値) は含めていません。"
                "現状の実装は常にハッシュのみです。"
            )
    finally:
        # 後始末: ゾンビ Excel を残さない
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception:  # noqa: BLE001
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:  # noqa: BLE001
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass
        # それでも残ったら強制終了
        if app_pid:
            try:
                import ctypes

                handle = ctypes.windll.kernel32.OpenProcess(0x0001, False, app_pid)
                if handle:
                    ctypes.windll.kernel32.TerminateProcess(handle, 0)
                    ctypes.windll.kernel32.CloseHandle(handle)
            except Exception:  # noqa: BLE001
                pass
        shutil.rmtree(tmp_dir, ignore_errors=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Excel COM / テスト化 実現可能性スパイク")
    parser.add_argument(
        "--workbook",
        required=True,
        help="診断対象の .xlsm/.xlsx パス (原本は変更しない。コピーを開く)",
    )
    parser.add_argument(
        "--input-cell",
        default=None,
        help='波及確認用の入力。例: "設定!B5=0.10"',
    )
    parser.add_argument("--run-macro", default=None, help='実行するマクロ名。例: "Module1.Foo"')
    parser.add_argument(
        "--enable-events",
        action="store_true",
        help="入力書込/マクロ実行時にイベントマクロを有効化する (既定は無効)",
    )
    parser.add_argument(
        "--include-values",
        action="store_true",
        help="(将来用) 生値の同梱フラグ。現状は常にハッシュのみ。",
    )
    parser.add_argument(
        "--out-dir",
        default="spike_out",
        help="バンドル出力先ディレクトリ (既定: ./spike_out)",
    )
    args = parser.parse_args(argv)

    timestamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = Path(args.out_dir) / timestamp
    out_dir.mkdir(parents=True, exist_ok=True)
    log_path = out_dir / "run.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_path, encoding="utf-8"),
        ],
    )

    report: dict[str, Any] = {
        "timestamp": timestamp,
        "argv": sys.argv[1:],
        "accessvbom_registry": _read_accessvbom(),
    }

    if not sys.platform.startswith("win"):
        logger.error("このスパイクは Windows + Excel 専用です (現在: %s)", sys.platform)
        report["fatal"] = f"非 Windows 環境: {sys.platform}"
    else:
        try:
            run_probe(args, report)
        except Exception as exc:  # noqa: BLE001 - 失敗内容こそ診断材料
            logger.error("致命的エラー: %s", exc)
            report["fatal"] = f"{type(exc).__name__}: {exc}"
            report["traceback"] = traceback.format_exc()

    # レポート書き出し
    import json

    def _default(obj: Any) -> Any:
        return obj.as_dict() if isinstance(obj, Step) else str(obj)

    report_path = out_dir / "report.json"
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, default=_default), encoding="utf-8"
    )

    # バンドル (zip) 化。run.log に収録されるよう、zip 作成前にログを書き切る。
    bundle = Path(args.out_dir) / f"spike_bundle_{timestamp}.zip"
    logger.info("=" * 60)
    logger.info("診断バンドル: %s", bundle.resolve())
    logger.info("このPCへはこの zip を1つ持ち帰ってください (生のセル値は含みません)。")
    logger.info("=" * 60)
    logging.shutdown()

    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(report_path, report_path.name)
        if log_path.exists():
            zf.write(log_path, log_path.name)

    return 0 if "fatal" not in report else 1


if __name__ == "__main__":
    raise SystemExit(main())
