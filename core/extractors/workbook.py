"""ワークブック構造抽出モジュール.

openpyxl で .xlsx / .xlsm を開き、数式セル・名前付き範囲・条件付き書式・外部リンクを
Pydantic モデルに詰める。VBA は本モジュールでは扱わない (core.extractors.vba を別途呼ぶ).

SPEC.md §4.2 参照。
"""

from __future__ import annotations

import logging
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from core.exceptions import ExtractionError
from core.external_functions import detect_in_formula
from core.models import (
    CellFormula,
    ConditionalFormat,
    DataValidation,
    ExcelTable,
    FormControl,
    NamedRange,
    SheetInfo,
    Workbook,
)

# プレビュー範囲（先頭 N 行 × M 列）. 解釈は加えず literal に出力するだけ.
PREVIEW_MAX_ROWS = 20
PREVIEW_MAX_COLS = 20

logger = logging.getLogger(__name__)


# `Sheet1!$A$1` や `'My Sheet'!A1:B10` の先頭シート名部分を取り出す
_SHEET_PREFIX_RE = re.compile(r"^(?:'([^']+)'|([^!'\s]+))!")


def _extract_sheet_name(refers_to: str) -> str | None:
    """`Sheet1!$A$1` 形式の参照から先頭のシート名を抜き出す.

    ヒットしない (シート名なし) 場合は None.
    """
    m = _SHEET_PREFIX_RE.match(refers_to.strip())
    if not m:
        return None
    return m.group(1) or m.group(2)


def _extract_formula_refs(formula: str) -> list[str]:
    """数式から RANGE トークンを抽出する.

    openpyxl の Tokenizer に頼る。`=SUMIF(Input!A:A, A2, Input!E:E)` から
    `["Input!A:A", "A2", "Input!E:E"]` を得る想定。
    """
    from openpyxl.formula.tokenizer import Tokenizer

    if not formula:
        return []
    if not formula.startswith("="):
        formula = "=" + formula

    try:
        tok = Tokenizer(formula)
    except Exception:  # noqa: BLE001 - 構文エラーは握りつぶしてログ
        logger.debug("Failed to tokenize formula: %r", formula)
        return []

    refs: list[str] = []
    for t in tok.items:
        if getattr(t, "subtype", None) == "RANGE":
            refs.append(t.value)
    return refs


def _extract_formulas(ws: object) -> list[CellFormula]:
    """1シートから数式セルを抽出する."""
    formulas: list[CellFormula] = []
    sheet_title: str = ws.title  # type: ignore[attr-defined]

    for row in ws.iter_rows():  # type: ignore[attr-defined]
        for cell in row:
            if cell.data_type != "f":
                continue
            value = cell.value
            if value is None:
                continue
            formula_str = str(value)
            coord = f"{sheet_title}!{cell.coordinate}"
            refs = _extract_formula_refs(formula_str)
            external_fns = detect_in_formula(formula_str)
            formulas.append(
                CellFormula(
                    coord=coord,
                    formula=formula_str,
                    refs=refs,
                    external_functions=external_fns,
                )
            )
    return formulas


def _extract_conditional_formats(ws: object) -> list[ConditionalFormat]:
    """1シートから条件付き書式を抽出する."""
    cfs: list[ConditionalFormat] = []
    cf_list = getattr(ws, "conditional_formatting", None)
    if cf_list is None:
        return cfs
    try:
        for rng, rules in cf_list._cf_rules.items():
            range_str = str(rng.sqref) if hasattr(rng, "sqref") else str(rng)
            for rule in rules:
                rule_str = _summarize_cf_rule(rule)
                cfs.append(ConditionalFormat(range=range_str, rule=rule_str))
    except Exception:  # noqa: BLE001
        logger.debug(
            "Failed to extract conditional formats from %s",
            sheet_title := getattr(ws, "title", "?"),
        )
        _ = sheet_title
    return cfs


def _summarize_cf_rule(rule: object) -> str:
    """ConditionalFormatRule を人間が読める短い文字列に."""
    rtype = getattr(rule, "type", None) or "rule"
    operator = getattr(rule, "operator", None)
    formula = getattr(rule, "formula", None)
    parts = [str(rtype)]
    if operator:
        parts.append(str(operator))
    if formula:
        formulas_list = list(formula) if not isinstance(formula, str) else [formula]
        parts.append(",".join(str(f) for f in formulas_list))
    return " ".join(parts)


def _extract_external_links(wb: object) -> list[str]:
    """外部リンクを抽出する.

    `wb._external_links` は private API なので失敗を許容する。
    """
    links: list[str] = []
    raw = getattr(wb, "_external_links", None)
    if not raw:
        return links
    try:
        for link in raw:
            file_link = getattr(link, "file_link", None)
            target = getattr(file_link, "Target", None) if file_link else None
            if target:
                links.append(str(target))
    except Exception:  # noqa: BLE001
        logger.debug("Failed to enumerate external links")
    return links


def _attach_named_ranges(wb: object, sheets_by_name: dict[str, SheetInfo]) -> None:
    """ワークブック定義名を、参照先シートの SheetInfo.named_ranges に振り分ける.

    シートが特定できない (refers_to が解析不能 or 該当シート無し) 場合は
    最初のシートに付ける。
    """
    defined_names = getattr(wb, "defined_names", None)
    if not defined_names:
        return

    try:
        items = list(defined_names.items())
    except Exception:  # noqa: BLE001
        logger.debug("Failed to enumerate defined_names")
        return

    for name, defn in items:
        refers_to = getattr(defn, "value", None) or ""
        sheet_name = _extract_sheet_name(refers_to)
        target_sheet: SheetInfo | None = None
        if sheet_name and sheet_name in sheets_by_name:
            target_sheet = sheets_by_name[sheet_name]
        elif sheets_by_name:
            target_sheet = next(iter(sheets_by_name.values()))
        if target_sheet is not None:
            target_sheet.named_ranges.append(NamedRange(name=name, refers_to=refers_to))


def _extract_excel_tables(ws: object) -> list[ExcelTable]:
    """ws.tables から Excel テーブル (ListObject) を抽出する.

    `ws.tables` は openpyxl が解釈した確定情報なのでヒューリスティック不要。
    """
    result: list[ExcelTable] = []
    tables_attr = getattr(ws, "tables", None)
    if tables_attr is None:
        return result
    try:
        # openpyxl 3.1 の TableList:
        # - keys() / __iter__ で table 名一覧
        # - items() は (name, ref_str) を返す (Table オブジェクト本体ではない)
        # - get(name) で Table オブジェクト本体を取れる
        names: list[str]
        if hasattr(tables_attr, "keys"):
            names = list(tables_attr.keys())
        else:
            names = [getattr(t, "name", "") for t in tables_attr]
        for name in names:
            table = tables_attr.get(name) if hasattr(tables_attr, "get") else None
            if table is None:
                continue
            display_name = (
                getattr(table, "displayName", None) or getattr(table, "name", None) or name
            )
            ref = getattr(table, "ref", "") or ""
            header_row_count = getattr(table, "headerRowCount", 1) or 1
            if not display_name or not ref:
                continue
            result.append(
                ExcelTable(
                    name=str(display_name),
                    ref=str(ref),
                    header_row_count=int(header_row_count),
                )
            )
    except Exception:  # noqa: BLE001
        logger.debug("Failed to enumerate tables on sheet")
    return result


def _extract_data_validations(ws: object) -> list[DataValidation]:
    """openpyxl の data_validations を抽出する.

    各 DataValidation は適用範囲 (sqref) ごとに 1 行 (sqref に複数範囲が
    含まれていれば、それを space 区切りでまとめて格納). type / formula1 /
    prompt / error message も拾う.
    """
    out: list[DataValidation] = []
    dv_attr = getattr(ws, "data_validations", None)
    if dv_attr is None:
        return out
    try:
        items = list(getattr(dv_attr, "dataValidation", []) or [])
    except Exception:  # noqa: BLE001
        logger.debug("Failed to enumerate data_validations on sheet")
        return out

    for dv in items:
        try:
            sqref = getattr(dv, "sqref", None)
            range_str = str(sqref) if sqref is not None else ""
            dv_type = str(getattr(dv, "type", "") or "")
            formula1 = getattr(dv, "formula1", "") or ""
            operator = str(getattr(dv, "operator", "") or "")
            prompt = str(getattr(dv, "prompt", "") or "")
            error = str(getattr(dv, "error", "") or "")
            allow_blank = bool(getattr(dv, "allowBlank", True))
            if not range_str or not dv_type:
                continue
            out.append(
                DataValidation(
                    range=range_str,
                    type=dv_type,
                    formula=str(formula1),
                    operator=operator,
                    prompt=prompt,
                    error=error,
                    allow_blank=allow_blank,
                )
            )
        except Exception:  # noqa: BLE001
            logger.debug("Skipping malformed data_validation entry")
            continue
    return out


# ----- フォームコントロール (VML-based) ----------------------------------

# VML 名前空間
_VML_NS = {
    "v": "urn:schemas-microsoft-com:vml",
    "x": "urn:schemas-microsoft-com:office:excel",
    "o": "urn:schemas-microsoft-com:office:office",
}
# rels 名前空間
_REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}


def _resolve_package_part(base_dir: str, target: str) -> str:
    """Open XML パッケージ内の Relationship Target を正規化する.

    Args:
        base_dir: relationship 所有パーツのディレクトリ。例: ``xl/worksheets``。
        target: rels に書かれた Target。相対パスまたは ``/xl/...`` 形式。

    Returns:
        zip 内で読める、先頭スラッシュなしの正規化パス。
    """
    if not target:
        return ""
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base_dir, target))


def _xlsm_sheet_to_vml_map(zf: zipfile.ZipFile) -> dict[str, list[str]]:
    """xlsm zip から「シート名 → vmlDrawing*.vml のフルパス一覧」を返す.

    Args:
        zf: xlsm/xltm/xlsb を開いた ZipFile。

    Returns:
        シート名をキー、関連する VML パーツ一覧を値にしたマップ。
        rels の解決にコケた場合は、取れた範囲だけ返す。
    """
    sheet_to_vml: dict[str, list[str]] = {}
    try:
        wb_xml = zf.read("xl/workbook.xml")
        wb_rels = zf.read("xl/_rels/workbook.xml.rels")
    except KeyError:
        return sheet_to_vml

    # r:id → sheet file
    rid_to_target: dict[str, str] = {}
    try:
        root = ET.fromstring(wb_rels)
        for rel in root.iter(
            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        ):
            rid_to_target[rel.attrib.get("Id", "")] = rel.attrib.get("Target", "")
    except ET.ParseError:
        return sheet_to_vml

    # sheet name → r:id → sheet file path
    name_to_sheet_path: dict[str, str] = {}
    try:
        root = ET.fromstring(wb_xml)
        for sheet in root.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"):
            name = sheet.attrib.get("name", "")
            rid_key = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            rid = sheet.attrib.get(rid_key, "")
            target = rid_to_target.get(rid, "")
            if not name or not target:
                continue
            # workbook.xml.rels の Target は通常 "worksheets/sheet1.xml"。
            sheet_path = _resolve_package_part("xl", target)
            name_to_sheet_path[name] = sheet_path
    except ET.ParseError:
        return sheet_to_vml

    # 各 sheet の _rels を辿って vmlDrawing を見つける
    for name, sheet_path in name_to_sheet_path.items():
        # 例: xl/worksheets/sheet1.xml → xl/worksheets/_rels/sheet1.xml.rels
        if "/" in sheet_path:
            dir_part, file_part = sheet_path.rsplit("/", 1)
        else:
            dir_part, file_part = "", sheet_path
        rels_path = f"{dir_part}/_rels/{file_part}.rels"
        try:
            rels_xml = zf.read(rels_path)
        except KeyError:
            continue
        try:
            rels_root = ET.fromstring(rels_xml)
        except ET.ParseError:
            continue
        for rel in rels_root.iter(
            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        ):
            target = rel.attrib.get("Target", "")
            if "vmlDrawing" in target and target.endswith(".vml"):
                vml_path = _resolve_package_part(dir_part, target)
                sheet_to_vml.setdefault(name, []).append(vml_path)

    return sheet_to_vml


def _parse_vml_form_controls(vml_bytes: bytes) -> list[FormControl]:
    """VML 1 ファイル分から FormControl を抽出する.

    各 v:shape の中の x:ClientData (type 属性) と x:FmlaMacro を見る:
      - <x:ClientData ObjectType="Button|Checkbox|Drop|Spin|...">
      - <x:FmlaMacro>マクロ名</x:FmlaMacro>
      - <v:textbox> 配下のテキスト = ボタン表面文字
      - <x:Anchor>FromCol,FromColOff,FromRow,FromRowOff,ToCol,...</x:Anchor>
    """
    out: list[FormControl] = []
    try:
        root = ET.fromstring(vml_bytes)
    except ET.ParseError:
        return out

    # shape 要素は v: 名前空間配下. ET は名前空間付きで {ns}local の形でアクセス.
    v_shape_tag = f"{{{_VML_NS['v']}}}shape"
    x_cd_tag = f"{{{_VML_NS['x']}}}ClientData"
    x_fmla_tag = f"{{{_VML_NS['x']}}}FmlaMacro"
    x_anchor_tag = f"{{{_VML_NS['x']}}}Anchor"
    v_textbox_tag = f"{{{_VML_NS['v']}}}textbox"

    object_type_map = {
        "Button": "button",
        "Checkbox": "checkbox",
        "Radio": "radio",
        "Drop": "dropdown",
        "List": "listbox",
        "Spin": "spinner",
        "Scroll": "scrollbar",
        "GBox": "groupbox",
        "Label": "label",
    }

    for shape in root.iter(v_shape_tag):
        name = shape.attrib.get(f"{{{_VML_NS['o']}}}spid", "") or shape.attrib.get("id", "")
        cd = shape.find(x_cd_tag)
        if cd is None:
            continue
        obj_type = cd.attrib.get("ObjectType", "")
        kind = object_type_map.get(obj_type, obj_type.lower() if obj_type else "control")

        fmla = cd.find(x_fmla_tag)
        macro = (fmla.text or "").strip() if fmla is not None and fmla.text else ""

        anchor_el = cd.find(x_anchor_tag)
        anchor_text = ""
        if anchor_el is not None and anchor_el.text:
            anchor_text = _anchor_to_cell(anchor_el.text)

        # ボタンの表面テキストを v:textbox から拾う (best-effort)
        text = ""
        textbox = shape.find(v_textbox_tag)
        if textbox is not None:
            # textbox 配下の全テキストを連結
            text = "".join(textbox.itertext()).strip()
            # 改行・余分な空白を 1 つに圧縮
            text = re.sub(r"\s+", " ", text)[:120]

        # マクロも表示テキストも空のコントロールは情報量が薄いのでスキップ
        if not macro and not text:
            continue

        out.append(
            FormControl(
                kind=kind or "control",
                name=name,
                text=text,
                macro=macro,
                anchor=anchor_text,
            )
        )
    return out


def _anchor_to_cell(anchor_text: str) -> str:
    """`x:Anchor` の `FromCol,FromColOff,FromRow,FromRowOff,...` を "A1" 形式へ.

    パース失敗時は空文字を返す.
    """
    try:
        from openpyxl.utils import get_column_letter

        parts = [p.strip() for p in anchor_text.split(",")]
        if len(parts) < 4:
            return ""
        from_col = int(parts[0])
        from_row = int(parts[2])
        return f"{get_column_letter(from_col + 1)}{from_row + 1}"
    except (ValueError, IndexError):
        return ""


def _extract_form_controls(file_path: Path, sheet_names: list[str]) -> dict[str, list[FormControl]]:
    """xlsm 内の VML ドローイングから (シート名 → フォームコントロール) を抽出.

    .xls / .xlsx は VBA を含まないので空を返す. xlsm でも VML がない (= ボタンが
    一つも置かれていない) なら空. zipfile / XML パース失敗は警告ログのみで握り潰し
    結果は best-effort.
    """
    out: dict[str, list[FormControl]] = {sn: [] for sn in sheet_names}
    if file_path.suffix.lower() not in {".xlsm", ".xltm", ".xlsb"}:
        return out
    try:
        with zipfile.ZipFile(file_path) as zf:
            sheet_to_vml = _xlsm_sheet_to_vml_map(zf)
            for sheet_name, vml_paths in sheet_to_vml.items():
                if sheet_name not in out:
                    continue
                for vml_path in vml_paths:
                    try:
                        vml_bytes = zf.read(vml_path)
                    except KeyError:
                        continue
                    out[sheet_name].extend(_parse_vml_form_controls(vml_bytes))
    except (zipfile.BadZipFile, OSError) as e:
        logger.warning("failed to read form controls from %s: %s", file_path.name, e)
    return out


def _extract_merged_ranges(ws: object) -> list[str]:
    """シートのマージ範囲を文字列リストで返す."""
    result: list[str] = []
    merged_attr = getattr(ws, "merged_cells", None)
    if merged_attr is None:
        return result
    try:
        for rng in merged_attr.ranges:
            result.append(str(rng))
    except Exception:  # noqa: BLE001
        logger.debug("Failed to enumerate merged_cells")
    return result


def _extract_preview(ws: object) -> tuple[list[list[str | None]], str]:
    """シート冒頭の N 行 × M 列を literal に取得する.

    Returns:
        (preview_rows, origin) のタプル. preview_rows は等長の2次元リスト、
        各要素は文字列化したセル値か None (空セル). origin は "A1" 等の起点座標.
    """
    from openpyxl.utils import get_column_letter

    max_row = min(getattr(ws, "max_row", 0) or 0, PREVIEW_MAX_ROWS)
    max_col = min(getattr(ws, "max_column", 0) or 0, PREVIEW_MAX_COLS)
    if max_row == 0 or max_col == 0:
        return [], ""

    rows: list[list[str | None]] = []
    try:
        for raw_row in ws.iter_rows(  # type: ignore[attr-defined]
            min_row=1, max_row=max_row, min_col=1, max_col=max_col, values_only=True
        ):
            row_vals: list[str | None] = []
            for v in raw_row:
                row_vals.append(None if v is None else str(v))
            rows.append(row_vals)
    except Exception:  # noqa: BLE001
        logger.debug("Failed to extract preview rows")
        return [], ""

    origin = f"A1:{get_column_letter(max_col)}{max_row}"
    return rows, origin


def extract_workbook(file_path: Path) -> Workbook:
    """Excelファイルからシート構造を抽出する.

    Args:
        file_path: 対象ファイルのパス. .xlsx / .xlsm を想定.

    Returns:
        Workbook モデル. VBA モジュールは含めない (vba_modules は空).
        `.xls` (旧バイナリ形式) が渡された場合は sheets=[] で返し警告ログを出す.

    Raises:
        ExtractionError: ファイルが存在しない、または openpyxl が開けない場合.
    """
    if not file_path.exists():
        raise ExtractionError(f"File not found: {file_path}")

    if file_path.suffix.lower() == ".xls":
        logger.warning(
            ".xls (legacy binary) is not supported by openpyxl; "
            "returning empty workbook structure for %s",
            file_path.name,
        )
        return Workbook(filename=file_path.name)

    from openpyxl import load_workbook

    try:
        wb = load_workbook(filename=str(file_path), keep_vba=True, data_only=False)
    except Exception as e:  # noqa: BLE001
        raise ExtractionError(f"openpyxl failed to open {file_path}: {e}") from e

    sheets: list[SheetInfo] = []
    sheets_by_name: dict[str, SheetInfo] = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        preview_rows, preview_origin = _extract_preview(ws)
        info = SheetInfo(
            name=sn,
            rows=ws.max_row or 0,
            cols=ws.max_column or 0,
            formulas=_extract_formulas(ws),
            conditional_formats=_extract_conditional_formats(ws),
            tables=_extract_excel_tables(ws),
            merged_ranges=_extract_merged_ranges(ws),
            data_validations=_extract_data_validations(ws),
            preview_rows=preview_rows,
            preview_origin=preview_origin,
        )
        sheets.append(info)
        sheets_by_name[sn] = info

    _attach_named_ranges(wb, sheets_by_name)

    # フォームコントロール (ボタン → マクロ紐付け) は xlsm の VML を直接読む.
    # openpyxl は VML を解釈しないので zipfile + XML パースで自前抽出.
    fc_map = _extract_form_controls(file_path, list(sheets_by_name.keys()))
    for sn, controls in fc_map.items():
        if sn in sheets_by_name and controls:
            sheets_by_name[sn].form_controls = controls

    return Workbook(
        filename=file_path.name,
        sheets=sheets,
        external_links=_extract_external_links(wb),
    )
