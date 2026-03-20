from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Literal, Optional

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

DEFAULT_ACTION = "use_b"
VALID_ACTIONS = {"use_a", "use_b", "manual"}
WorkbookSide = Literal["a", "b"]


@dataclass
class CellDiff:
    sheet: str
    row: int
    column: int
    value_a: Optional[object]
    value_b: Optional[object]

    @property
    def coordinate(self) -> str:
        letters = ""
        col = self.column
        while col > 0:
            col, rem = divmod(col - 1, 26)
            letters = chr(65 + rem) + letters
        return f"{letters}{self.row}"


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


def _normalize(value: object, options: CompareOptions) -> object:
    if isinstance(value, str):
        normalized = value.strip() if options.strip_strings else value
        return normalized if options.case_sensitive else normalized.lower()

    if options.ignore_empty_string_vs_none and value == "":
        return None

    return value


def _resolve_direction(base: WorkbookSide) -> tuple[str, str]:
    if base not in {"a", "b"}:
        raise ValueError("base debe ser 'a' o 'b'")
    return ("a", "b") if base == "a" else ("b", "a")


def compare_workbooks(
    path_a: str | Path,
    path_b: str | Path,
    options: CompareOptions | None = None,
) -> WorkbookDiff:
    options = options or CompareOptions()
    wb_a = load_workbook(path_a, data_only=False)
    wb_b = load_workbook(path_b, data_only=False)

    sheets_a = set(wb_a.sheetnames)
    sheets_b = set(wb_b.sheetnames)

    only_in_a = sorted(sheets_a - sheets_b)
    only_in_b = sorted(sheets_b - sheets_a)
    common = sorted(sheets_a & sheets_b)

    differences: Dict[str, List[CellDiff]] = {}

    for sheet_name in common:
        ws_a = wb_a[sheet_name]
        ws_b = wb_b[sheet_name]

        max_row = max(ws_a.max_row, ws_b.max_row)
        max_col = max(ws_a.max_column, ws_b.max_column)

        sheet_diffs: List[CellDiff] = []

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                val_a = ws_a.cell(row=row, column=col).value
                val_b = ws_b.cell(row=row, column=col).value

                if _normalize(val_a, options) != _normalize(val_b, options):
                    sheet_diffs.append(
                        CellDiff(
                            sheet=sheet_name,
                            row=row,
                            column=col,
                            value_a=val_a,
                            value_b=val_b,
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

    headers = ["sheet", "cell", "row", "column", "value_a", "value_b", "action", "manual_value"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type="solid", start_color="1F4E78", end_color="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)

    for d in diff.all_differences():
        ws.append([d.sheet, d.coordinate, d.row, d.column, d.value_a, d.value_b, default_action, None])

    dv = DataValidation(type="list", formula1='"use_a,use_b,manual"', allow_blank=False)
    ws.add_data_validation(dv)
    if ws.max_row >= 2:
        dv.add(f"G2:G{ws.max_row}")

    ws.freeze_panes = "A2"

    summary = wb.create_sheet("Resumen")
    summary.append(["Métrica", "Valor"])
    summary.append(["Hojas en común", len(diff.common_sheets)])
    summary.append(["Hojas solo en A", len(diff.only_in_a)])
    summary.append(["Hojas solo en B", len(diff.only_in_b)])
    summary.append(["Diferencias de celdas", len(diff.all_differences())])
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
        return pd.DataFrame(columns=["sheet", "cell", "row", "column", "value_a", "value_b", "action", "manual_value"])

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
    workbook_a: str | Path,
    decisions: pd.DataFrame,
    output_path: str | Path,
    workbook_b: str | Path,
    base: WorkbookSide = "a",
    include_sheets_from_source_only: bool = True,
) -> Path:
    base_key, source_key = _resolve_direction(base)
    workbook_paths = {"a": workbook_a, "b": workbook_b}

    wb_out = load_workbook(workbook_paths[base_key])
    wb_source = load_workbook(workbook_paths[source_key])

    for _, row in decisions.iterrows():
        sheet_name = row["sheet"]
        r = int(row["row"])
        c = int(row["column"])
        action = row["action"]

        if sheet_name not in wb_out.sheetnames:
            wb_out.create_sheet(sheet_name)
        ws_out = wb_out[sheet_name]

        if action == f"use_{base_key}":
            continue
        if action == f"use_{source_key}":
            if sheet_name in wb_source.sheetnames:
                ws_out.cell(row=r, column=c).value = wb_source[sheet_name].cell(row=r, column=c).value
        elif action == "manual":
            ws_out.cell(row=r, column=c).value = row.get("manual_value")

    if include_sheets_from_source_only:
        for sheet in wb_source.sheetnames:
            if sheet not in wb_out.sheetnames:
                src = wb_source[sheet]
                ws_new = wb_out.create_sheet(sheet)
                for row in src.iter_rows():
                    for cell in row:
                        ws_new[cell.coordinate] = cell.value

    output_path = Path(output_path)
    wb_out.save(output_path)
    return output_path
