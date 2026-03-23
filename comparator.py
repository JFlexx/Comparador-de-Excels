from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Literal, Optional

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

DEFAULT_ACTION = "use_b"
VALID_ACTIONS = {"use_a", "use_b", "manual"}
VALID_COMPARE_MODES = {"coordinate", "row-based"}


def column_letter(column: int) -> str:
    letters = ""
    current = column
    while current > 0:
        current, rem = divmod(current - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


@dataclass
class CellDiff:
    sheet: str
    row: int
    column: int
    value_a: Optional[object]
    value_b: Optional[object]
    diff_type: str = "modified"
    key: Optional[str] = None
    header: Optional[str] = None

    @property
    def coordinate(self) -> str:
        return f"{column_letter(self.column)}{self.row}"


@dataclass
class SheetDiffSummary:
    sheet: str
    differences: List[CellDiff]
    total_differences: int = field(init=False)
    columns: List[str] = field(init=False)
    row_numbers: List[int] = field(init=False)

    def __post_init__(self) -> None:
        self.total_differences = len(self.differences)
        self.columns = sorted({column_letter(diff.column) for diff in self.differences})
        self.row_numbers = sorted({diff.row for diff in self.differences})


@dataclass
class WorkbookDiff:
    only_in_a: List[str]
    only_in_b: List[str]
    common_sheets: List[str]
    differences: Dict[str, List[CellDiff]]
    grouped_differences: Dict[str, SheetDiffSummary] = field(init=False)
    total_differences: int = field(init=False)
    _all_differences: List[CellDiff] = field(init=False, repr=False)

    def __post_init__(self) -> None:
        self.grouped_differences = {}
        self._all_differences = []

        for sheet in self.common_sheets:
            sheet_diffs = list(self.differences.get(sheet, []))
            self.differences[sheet] = sheet_diffs
            self.grouped_differences[sheet] = SheetDiffSummary(sheet=sheet, differences=sheet_diffs)
            self._all_differences.extend(sheet_diffs)

        self.total_differences = len(self._all_differences)

    def all_differences(self) -> List[CellDiff]:
        return list(self._all_differences)

    def summary_rows(self) -> List[dict[str, object]]:
        return [
            {
                "sheet": summary.sheet,
                "differences": summary.total_differences,
                "columns": ", ".join(summary.columns),
                "rows": len(summary.row_numbers),
            }
            for summary in self.grouped_differences.values()
        ]

    def to_dataframe(self, default_action: str = DEFAULT_ACTION) -> pd.DataFrame:
        return diffs_to_dataframe(self._all_differences, default_action=default_action)


@dataclass
class CompareOptions:
    strip_strings: bool = True
    case_sensitive: bool = True
    ignore_empty_string_vs_none: bool = True
    compare_mode: str = "coordinate"
    sheet_keys: Dict[str, List[str]] = field(default_factory=dict)
    header_row: int = 1

    def __post_init__(self) -> None:
        if self.compare_mode not in VALID_COMPARE_MODES:
            raise ValueError(f"compare_mode debe ser uno de {sorted(VALID_COMPARE_MODES)}")
        self.sheet_keys = {sheet: list(keys) for sheet, keys in self.sheet_keys.items()}
        if self.header_row < 1:
            raise ValueError("header_row debe ser >= 1")


def _normalize(value: object, options: CompareOptions) -> object:
    if isinstance(value, str):
        normalized = value.strip() if options.strip_strings else value
        return normalized if options.case_sensitive else normalized.lower()

    if options.ignore_empty_string_vs_none and value == "":
        return None

    return value


def _stringify_key_part(value: object) -> str:
    return "<vacío>" if value is None else str(value)


def _build_key_label(key_headers: List[str], key_values: tuple[object, ...], fallback_row: int) -> str:
    if key_headers:
        parts = [f"{header}={_stringify_key_part(value)}" for header, value in zip(key_headers, key_values)]
        return ", ".join(parts)
    return f"fila={fallback_row}"


def _read_sheet_rows(ws: Worksheet, options: CompareOptions) -> dict[str, object]:
    header_row = options.header_row
    raw_headers = [ws.cell(row=header_row, column=col).value for col in range(1, ws.max_column + 1)]

    headers: List[str] = []
    header_to_column: Dict[str, int] = {}
    for idx, raw_header in enumerate(raw_headers, start=1):
        header = str(raw_header).strip() if raw_header is not None else f"__col_{idx}"
        if not header:
            header = f"__col_{idx}"
        if header in header_to_column:
            header = f"{header}__{idx}"
        headers.append(header)
        header_to_column[header] = idx

    rows: List[dict[str, object]] = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        values_by_header = {
            header: ws.cell(row=row_idx, column=col_idx).value
            for col_idx, header in enumerate(headers, start=1)
        }
        if all(value is None for value in values_by_header.values()):
            continue
        rows.append({"row_idx": row_idx, "values": values_by_header})

    return {"headers": headers, "header_to_column": header_to_column, "rows": rows}


def _record_key(
    sheet_name: str,
    record: dict[str, object],
    key_headers: List[str],
    comparable_headers: List[str],
    options: CompareOptions,
) -> tuple[object, ...]:
    values = record["values"]
    if key_headers:
        missing = [header for header in key_headers if header not in values]
        if missing:
            raise ValueError(
                f"La hoja '{sheet_name}' no contiene las columnas clave requeridas: {missing}"
            )
        return tuple(_normalize(values[header], options) for header in key_headers)

    return tuple((header, _normalize(values.get(header), options)) for header in comparable_headers)


def _compare_sheet_by_rows(sheet_name: str, ws_a: Worksheet, ws_b: Worksheet, options: CompareOptions) -> List[CellDiff]:
    parsed_a = _read_sheet_rows(ws_a, options)
    parsed_b = _read_sheet_rows(ws_b, options)

    headers_a = parsed_a["headers"]
    headers_b = parsed_b["headers"]
    comparable_headers = list(dict.fromkeys([*headers_a, *headers_b]))
    key_headers = options.sheet_keys.get(sheet_name, [])

    grouped_a: dict[tuple[object, ...], List[dict[str, object]]] = defaultdict(list)
    grouped_b: dict[tuple[object, ...], List[dict[str, object]]] = defaultdict(list)

    for record in parsed_a["rows"]:
        grouped_a[_record_key(sheet_name, record, key_headers, comparable_headers, options)].append(record)
    for record in parsed_b["rows"]:
        grouped_b[_record_key(sheet_name, record, key_headers, comparable_headers, options)].append(record)

    diffs: List[CellDiff] = []

    for key in sorted(set(grouped_a) | set(grouped_b), key=lambda item: repr(item)):
        records_a = grouped_a.get(key, [])
        records_b = grouped_b.get(key, [])
        shared_count = min(len(records_a), len(records_b))

        for index in range(shared_count):
            record_a = records_a[index]
            record_b = records_b[index]
            key_label = _build_key_label(key_headers, key, record_b["row_idx"])
            for header in comparable_headers:
                value_a = record_a["values"].get(header)
                value_b = record_b["values"].get(header)
                if _normalize(value_a, options) == _normalize(value_b, options):
                    continue
                column = (
                    parsed_b["header_to_column"].get(header)
                    or parsed_a["header_to_column"].get(header)
                    or 1
                )
                diffs.append(
                    CellDiff(
                        sheet=sheet_name,
                        row=record_b["row_idx"],
                        column=column,
                        value_a=value_a,
                        value_b=value_b,
                        diff_type="modified",
                        key=key_label,
                        header=header,
                    )
                )

        for record_a in records_a[shared_count:]:
            key_label = _build_key_label(key_headers, key, record_a["row_idx"])
            for header in comparable_headers:
                value_a = record_a["values"].get(header)
                if _normalize(value_a, options) is None:
                    continue
                column = parsed_a["header_to_column"].get(header) or 1
                diffs.append(
                    CellDiff(
                        sheet=sheet_name,
                        row=record_a["row_idx"],
                        column=column,
                        value_a=value_a,
                        value_b=None,
                        diff_type="deleted",
                        key=key_label,
                        header=header,
                    )
                )

        for record_b in records_b[shared_count:]:
            key_label = _build_key_label(key_headers, key, record_b["row_idx"])
            for header in comparable_headers:
                value_b = record_b["values"].get(header)
                if _normalize(value_b, options) is None:
                    continue
                column = parsed_b["header_to_column"].get(header) or 1
                diffs.append(
                    CellDiff(
                        sheet=sheet_name,
                        row=record_b["row_idx"],
                        column=column,
                        value_a=None,
                        value_b=value_b,
                        diff_type="added",
                        key=key_label,
                        header=header,
                    )
                )

    return diffs


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

        if options.compare_mode == "row-based":
            differences[sheet_name] = _compare_sheet_by_rows(sheet_name, ws_a, ws_b, options)
            continue

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


def diffs_to_dataframe(diffs: Iterable[CellDiff], default_action: str = DEFAULT_ACTION) -> pd.DataFrame:
    if default_action not in VALID_ACTIONS:
        raise ValueError(f"default_action debe ser uno de {sorted(VALID_ACTIONS)}")

    rows = []
    for d in diffs:
        rows.append(
            {
                "sheet": d.sheet,
                "cell": d.coordinate,
                "row": d.row,
                "column": d.column,
                "column_letter": column_letter(d.column),
                "context": f"{d.sheet}!{d.coordinate} · fila {d.row} · columna {column_letter(d.column)} ({d.column})",
                "value_a": d.value_a,
                "value_b": d.value_b,
                "action": default_action,
                "manual_value": None,
                "reviewed": False,
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
        "column_letter",
        "context",
        "value_a",
        "value_b",
        "action",
        "manual_value",
        "reviewed",
    ]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type="solid", start_color="1F4E78", end_color="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)

    for row in diff.to_dataframe(default_action=default_action).itertuples(index=False):
        ws.append(list(row))

    dv = DataValidation(type="list", formula1='"use_a,use_b,manual"', allow_blank=False)
    ws.add_data_validation(dv)
    if ws.max_row >= 2:
        dv.add(f"I2:I{ws.max_row}")

    ws.freeze_panes = "A2"

    summary = wb.create_sheet("Resumen")
    summary.append(["Métrica", "Valor"])
    summary.append(["Hojas en común", len(diff.common_sheets)])
    summary.append(["Hojas solo en A", len(diff.only_in_a)])
    summary.append(["Hojas solo en B", len(diff.only_in_b)])
    summary.append(["Diferencias de celdas", diff.total_differences])
    if diff.only_in_a:
        summary.append(["Lista solo en A", ", ".join(diff.only_in_a)])
    if diff.only_in_b:
        summary.append(["Lista solo en B", ", ".join(diff.only_in_b)])

    for row in diff.summary_rows():
        summary.append([f"Diferencias en {row['sheet']}", row["differences"]])

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
                "column_letter",
                "context",
                "value_a",
                "value_b",
                "action",
                "manual_value",
                "reviewed",
            ]
        )

    headers = [str(h) if h is not None else "" for h in rows[0]]
    data = rows[1:]
    df = pd.DataFrame(data, columns=headers)
    if "action" not in df.columns:
        raise ValueError("La hoja de decisiones debe contener la columna 'action'")

    if "reviewed" not in df.columns:
        df["reviewed"] = False

    df["action"] = df["action"].fillna(DEFAULT_ACTION).astype(str).str.strip()
    df["reviewed"] = df["reviewed"].map(
        lambda value: value
        if isinstance(value, bool)
        else str(value).strip().lower() in {"1", "true", "sí", "si", "yes", "x"}
    ).fillna(False)
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
