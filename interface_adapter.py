from __future__ import annotations

from pathlib import Path
from typing import Iterable, Mapping

import pandas as pd

from adapter_models import MergeDescriptor
from comparator import (
    CompareOptions,
    ComparisonRequest,
    DecisionTemplateRequest,
    DecisionsLoadRequest,
    MergeRequest,
    SERVICE,
    WorkbookDiff,
    WorkbookSide,
    source_action_for_base,
)


def parse_sheet_keys(raw_entries: Iterable[str], *, item_separator: str) -> dict[str, list[str]]:
    sheet_keys: dict[str, list[str]] = {}
    for raw_entry in raw_entries:
        cleaned = raw_entry.strip()
        if not cleaned:
            continue
        if item_separator not in cleaned:
            raise ValueError(
                f"Cada entrada debe usar el formato Hoja{item_separator}columna1,columna2"
            )
        sheet_name, raw_columns = cleaned.split(item_separator, 1)
        columns = [column.strip() for column in raw_columns.split(",") if column.strip()]
        if not sheet_name.strip() or not columns:
            raise ValueError("Cada entrada debe incluir nombre de hoja y al menos una columna clave")
        sheet_keys[sheet_name.strip()] = columns
    return sheet_keys


def build_compare_options(
    *,
    compare_mode: str,
    ignore_case: bool,
    keep_spaces: bool,
    empty_string_is_value: bool,
    header_row: int,
    sheet_keys: Mapping[str, list[str]],
) -> CompareOptions:
    return CompareOptions(
        strip_strings=not keep_spaces,
        case_sensitive=not ignore_case,
        ignore_empty_string_vs_none=not empty_string_is_value,
        compare_mode=compare_mode,
        sheet_keys={sheet: list(columns) for sheet, columns in sheet_keys.items()},
        header_row=int(header_row),
    )


def compare_files(path_a: str | Path, path_b: str | Path, options: CompareOptions) -> WorkbookDiff:
    return SERVICE.compare(ComparisonRequest(path_a=path_a, path_b=path_b, options=options))


def describe_merge(base: WorkbookSide) -> MergeDescriptor:
    source_key: WorkbookSide = "b" if base == "a" else "a"
    return MergeDescriptor(
        base_key=base,
        source_key=source_key,
        base_label=base.upper(),
        source_label=source_key.upper(),
        default_action=source_action_for_base(base),
    )


def export_template(diff: WorkbookDiff, output_path: str | Path, *, default_action: str) -> Path:
    return SERVICE.export_decision_template(
        DecisionTemplateRequest(diff=diff, output_path=output_path, default_action=default_action)
    )


def load_decisions(path: str | Path) -> pd.DataFrame:
    return SERVICE.decisions_from_excel(DecisionsLoadRequest(path=path))


def merge_workbooks(
    *,
    workbook_a: str | Path,
    workbook_b: str | Path,
    decisions: pd.DataFrame,
    output_path: str | Path,
    base: WorkbookSide,
    include_sheets_from_source_only: bool,
) -> Path:
    return SERVICE.apply_decisions(
        MergeRequest(
            workbook_a=workbook_a,
            workbook_b=workbook_b,
            decisions=decisions,
            output_path=output_path,
            base=base,
            include_sheets_from_source_only=include_sheets_from_source_only,
        )
    )
