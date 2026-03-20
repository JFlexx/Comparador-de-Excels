from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

DEFAULT_ACTION = "use_b"
VALID_ACTIONS = {"use_a", "use_b", "manual"}
VALID_COMPARE_MODES = {"coordinate", "row-based"}


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
        if self.column <= 0:
            return f"ROW{self.row}"

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


def diffs_to_dataframe(diffs: Iterable[CellDiff]) -> pd.DataFrame:
    rows = []
    for d in diffs:
        rows.append(
            {
                "sheet": d.sheet,
                "cell": d.coordinate,
                "row": d.row,
                "column": d.column,
                "header": d.header,
                "diff_type": d.diff_type,
                "key": d.key,
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
        "header",
        "diff_type",
        "key",
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
                d.header,
                d.diff_type,
                d.key,
                d.value_a,
                d.value_b,
                default_action,
                None,
            ]
        )

    dv = DataValidation(type="list", formula1='"use_a,use_b,manual"', allow_blank=False)
    ws.add_data_validation(dv)
    if ws.max_row >= 2:
        dv.add(f"J2:J{ws.max_row}")

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
        return pd.DataFrame(
            columns=[
                "sheet",
                "cell",
                "row",
                "column",
                "header",
                "diff_type",
                "key",
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
