from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

DEFAULT_ACTION = "use_b"
VALID_ACTIONS = {"use_a", "use_b", "manual"}
VALID_DIFFERENCE_TYPES = {
    "value_changed",
    "formula_changed",
    "cached_result_changed",
    "style_changed",
    "comment_changed",
}


@dataclass
class CellDiff:
    sheet: str
    row: int
    column: int
    value_a: Optional[object]
    value_b: Optional[object]
    difference_types: tuple[str, ...] = field(default_factory=tuple)
    formula_a: Optional[object] = None
    formula_b: Optional[object] = None
    cached_value_a: Optional[object] = None
    cached_value_b: Optional[object] = None

    @property
    def coordinate(self) -> str:
        letters = ""
        col = self.column
        while col > 0:
            col, rem = divmod(col - 1, 26)
            letters = chr(65 + rem) + letters
        return f"{letters}{self.row}"

    @property
    def difference_types_label(self) -> str:
        return ", ".join(self.difference_types)


@dataclass
class WorkbookDiff:
    only_in_a: List[str]
    only_in_b: List[str]
    common_sheets: List[str]
    differences: Dict[str, List[CellDiff]]

    def all_differences(self) -> List[CellDiff]:
        output: List[CellDiff] = []
        for sheet in self.common_sheets:
            output.extend(self.differences.get(sheet, []))
        return output


@dataclass
class CompareOptions:
    strip_strings: bool = True
    case_sensitive: bool = True
    ignore_empty_string_vs_none: bool = True
    compare_values: bool = True
    compare_formulas: bool = False
    compare_cached_results: bool = False
    compare_styles: bool = False
    compare_comments: bool = False

    def validate(self) -> None:
        if not any(
            (
                self.compare_values,
                self.compare_formulas,
                self.compare_cached_results,
                self.compare_styles,
                self.compare_comments,
            )
        ):
            raise ValueError("Debe activarse al menos un tipo de comparación")


def _normalize(value: object, options: CompareOptions) -> object:
    if isinstance(value, str):
        normalized = value.strip() if options.strip_strings else value
        return normalized if options.case_sensitive else normalized.lower()

    if options.ignore_empty_string_vs_none and value == "":
        return None

    return value


def _is_formula(value: object) -> bool:
    return isinstance(value, str) and value.startswith("=")


def _style_signature(cell) -> tuple[object, ...]:
    return (
        cell.number_format,
        repr(cell.font),
        repr(cell.fill),
        repr(cell.border),
        repr(cell.alignment),
        repr(cell.protection),
    )


def _comment_signature(comment: Comment | None) -> tuple[object, object] | None:
    if comment is None:
        return None
    return comment.text, comment.author


def _build_difference_types(cell_a, cell_b, cached_cell_a, cached_cell_b, options: CompareOptions) -> tuple[str, ...]:
    difference_types: List[str] = []

    raw_a = cell_a.value
    raw_b = cell_b.value
    raw_a_is_formula = _is_formula(raw_a)
    raw_b_is_formula = _is_formula(raw_b)
    has_formula = raw_a_is_formula or raw_b_is_formula

    if options.compare_values:
        should_compare_as_value = not (raw_a_is_formula and raw_b_is_formula)
        if should_compare_as_value and _normalize(raw_a, options) != _normalize(raw_b, options):
            difference_types.append("value_changed")

    if options.compare_formulas and has_formula:
        formula_a = raw_a if raw_a_is_formula else None
        formula_b = raw_b if raw_b_is_formula else None
        if _normalize(formula_a, options) != _normalize(formula_b, options):
            difference_types.append("formula_changed")

    if options.compare_cached_results and has_formula:
        cached_a = cached_cell_a.value if cached_cell_a is not None else None
        cached_b = cached_cell_b.value if cached_cell_b is not None else None
        if _normalize(cached_a, options) != _normalize(cached_b, options):
            difference_types.append("cached_result_changed")

    if options.compare_styles and _style_signature(cell_a) != _style_signature(cell_b):
        difference_types.append("style_changed")

    if options.compare_comments and _comment_signature(cell_a.comment) != _comment_signature(cell_b.comment):
        difference_types.append("comment_changed")

    return tuple(difference_types)


def compare_workbooks(
    path_a: str | Path,
    path_b: str | Path,
    options: CompareOptions | None = None,
) -> WorkbookDiff:
    options = options or CompareOptions()
    options.validate()

    wb_a = load_workbook(path_a, data_only=False)
    wb_b = load_workbook(path_b, data_only=False)
    cached_wb_a = load_workbook(path_a, data_only=True) if options.compare_cached_results else None
    cached_wb_b = load_workbook(path_b, data_only=True) if options.compare_cached_results else None

    sheets_a = set(wb_a.sheetnames)
    sheets_b = set(wb_b.sheetnames)

    only_in_a = sorted(sheets_a - sheets_b)
    only_in_b = sorted(sheets_b - sheets_a)
    common = sorted(sheets_a & sheets_b)

    differences: Dict[str, List[CellDiff]] = {}

    for sheet_name in common:
        ws_a = wb_a[sheet_name]
        ws_b = wb_b[sheet_name]
        cached_ws_a = cached_wb_a[sheet_name] if cached_wb_a is not None else None
        cached_ws_b = cached_wb_b[sheet_name] if cached_wb_b is not None else None

        max_row = max(ws_a.max_row, ws_b.max_row)
        max_col = max(ws_a.max_column, ws_b.max_column)

        sheet_diffs: List[CellDiff] = []

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell_a = ws_a.cell(row=row, column=col)
                cell_b = ws_b.cell(row=row, column=col)
                cached_cell_a = cached_ws_a.cell(row=row, column=col) if cached_ws_a is not None else None
                cached_cell_b = cached_ws_b.cell(row=row, column=col) if cached_ws_b is not None else None
                difference_types = _build_difference_types(
                    cell_a,
                    cell_b,
                    cached_cell_a,
                    cached_cell_b,
                    options,
                )

                if difference_types:
                    sheet_diffs.append(
                        CellDiff(
                            sheet=sheet_name,
                            row=row,
                            column=col,
                            value_a=cell_a.value,
                            value_b=cell_b.value,
                            difference_types=difference_types,
                            formula_a=cell_a.value if _is_formula(cell_a.value) else None,
                            formula_b=cell_b.value if _is_formula(cell_b.value) else None,
                            cached_value_a=cached_cell_a.value if cached_cell_a is not None else None,
                            cached_value_b=cached_cell_b.value if cached_cell_b is not None else None,
                        )
                    )

        differences[sheet_name] = sheet_diffs

    return WorkbookDiff(
        only_in_a=only_in_a,
        only_in_b=only_in_b,
        common_sheets=common,
        differences=differences,
    )


def diffs_to_dataframe(diffs: Iterable[CellDiff]) -> pd.DataFrame:
    rows = []
    for d in diffs:
        rows.append(
            {
                "sheet": d.sheet,
                "cell": d.coordinate,
                "row": d.row,
                "column": d.column,
                "difference_types": d.difference_types_label,
                "formula_a": d.formula_a,
                "formula_b": d.formula_b,
                "cached_value_a": d.cached_value_a,
                "cached_value_b": d.cached_value_b,
                "value_a": d.value_a,
                "value_b": d.value_b,
                "action": DEFAULT_ACTION,
                "manual_value": None,
            }
        )

    return pd.DataFrame(rows)


def export_decision_template(
    diff: WorkbookDiff,
    output_path: str | Path,
    default_action: str = DEFAULT_ACTION,
) -> Path:
    if default_action not in VALID_ACTIONS:
        raise ValueError(f"default_action debe ser uno de {sorted(VALID_ACTIONS)}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Decisiones"

    headers = [
        "sheet",
        "cell",
        "row",
        "column",
        "difference_types",
        "formula_a",
        "formula_b",
        "cached_value_a",
        "cached_value_b",
        "value_a",
        "value_b",
        "action",
        "manual_value",
    ]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type="solid", start_color="1F4E78", end_color="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)

    for d in diff.all_differences():
        ws.append(
            [
                d.sheet,
                d.coordinate,
                d.row,
                d.column,
                d.difference_types_label,
                d.formula_a,
                d.formula_b,
                d.cached_value_a,
                d.cached_value_b,
                d.value_a,
                d.value_b,
                default_action,
                None,
            ]
        )

    dv = DataValidation(type="list", formula1='"use_a,use_b,manual"', allow_blank=False)
    ws.add_data_validation(dv)
    if ws.max_row >= 2:
        dv.add(f"L2:L{ws.max_row}")

    ws.freeze_panes = "A2"

    summary = wb.create_sheet("Resumen")
    summary.append(["Métrica", "Valor"])
    summary.append(["Hojas en común", len(diff.common_sheets)])
    summary.append(["Hojas solo en A", len(diff.only_in_a)])
    summary.append(["Hojas solo en B", len(diff.only_in_b)])
    summary.append(["Diferencias de celdas", len(diff.all_differences())])
    difference_counter: Dict[str, int] = {diff_type: 0 for diff_type in VALID_DIFFERENCE_TYPES}
    for cell_diff in diff.all_differences():
        for diff_type in cell_diff.difference_types:
            difference_counter[diff_type] = difference_counter.get(diff_type, 0) + 1
    for diff_type in sorted(difference_counter):
        if difference_counter[diff_type]:
            summary.append([f"Tipo {diff_type}", difference_counter[diff_type]])
    if diff.only_in_a:
        summary.append(["Lista solo en A", ", ".join(diff.only_in_a)])
    if diff.only_in_b:
        summary.append(["Lista solo en B", ", ".join(diff.only_in_b)])

    output_path = Path(output_path)
    wb.save(output_path)
    return output_path


def decisions_from_excel(path: str | Path, sheet_name: str = "Decisiones") -> pd.DataFrame:
    wb = load_workbook(path, data_only=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"No existe la hoja '{sheet_name}' en el archivo de decisiones")

    ws = wb[sheet_name]
    rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))
    if not rows:
        return pd.DataFrame(
            columns=[
                "sheet",
                "cell",
                "row",
                "column",
                "difference_types",
                "formula_a",
                "formula_b",
                "cached_value_a",
                "cached_value_b",
                "value_a",
                "value_b",
                "action",
                "manual_value",
            ]
        )

    headers = [str(h) if h is not None else "" for h in rows[0]]
    data = rows[1:]
    df = pd.DataFrame(data, columns=headers)
    if "action" not in df.columns:
        raise ValueError("La hoja de decisiones debe contener la columna 'action'")

    df["action"] = df["action"].fillna(DEFAULT_ACTION).astype(str).str.strip()
    invalid = sorted(set(df[~df["action"].isin(VALID_ACTIONS)]["action"].tolist()))
    if invalid:
        raise ValueError(f"Acciones no válidas: {invalid}. Válidas: {sorted(VALID_ACTIONS)}")

    return df


def apply_decisions(
    base_workbook: str | Path,
    decisions: pd.DataFrame,
    output_path: str | Path,
    workbook_b: str | Path,
    include_sheets_only_in_b: bool = True,
) -> Path:
    wb_out = load_workbook(base_workbook)
    wb_b = load_workbook(workbook_b)

    for _, row in decisions.iterrows():
        sheet_name = row["sheet"]
        r = int(row["row"])
        c = int(row["column"])
        action = row["action"]

        if sheet_name not in wb_out.sheetnames:
            wb_out.create_sheet(sheet_name)
        ws_out = wb_out[sheet_name]

        if action == "use_a":
            continue
        if action == "use_b":
            if sheet_name in wb_b.sheetnames:
                ws_out.cell(row=r, column=c).value = wb_b[sheet_name].cell(row=r, column=c).value
        elif action == "manual":
            ws_out.cell(row=r, column=c).value = row.get("manual_value")

    if include_sheets_only_in_b:
        for sheet in wb_b.sheetnames:
            if sheet not in wb_out.sheetnames:
                src = wb_b[sheet]
                ws_new = wb_out.create_sheet(sheet)
                for source_row in src.iter_rows():
                    for cell in source_row:
                        ws_new.cell(row=cell.row, column=cell.column).value = cell.value

    output_path = Path(output_path)
    wb_out.save(output_path)
    return output_path
