from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Mapping

import pandas as pd

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


@dataclass(frozen=True)
class MergeLabels:
    base: str
    source: str
    base_key: WorkbookSide


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


def parse_sheet_keys_block(raw_value: str) -> dict[str, list[str]]:
    return parse_sheet_keys(raw_value.splitlines(), item_separator=":")


def parse_sheet_keys_args(values: list[str] | None) -> dict[str, list[str]]:
    return parse_sheet_keys(values or [], item_separator="=")


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


def default_action_for_base(base: WorkbookSide) -> str:
    return source_action_for_base(base)


def merge_labels(base: WorkbookSide) -> MergeLabels:
    if base == "a":
        return MergeLabels(base="A", source="B", base_key="a")
    return MergeLabels(base="B", source="A", base_key="b")


def build_review_dataframe(diff: WorkbookDiff, *, default_action: str) -> pd.DataFrame:
    decisions_df = diff.to_dataframe(default_action=default_action)
    if decisions_df.empty:
        return decisions_df

    decisions_df = decisions_df.copy()
    decisions_df["value_a_display"] = decisions_df["value_a"].map(format_display_value)
    decisions_df["value_b_display"] = decisions_df["value_b"].map(format_display_value)
    decisions_df["preview"] = decisions_df.apply(
        lambda row: f"🅰️ {row['value_a_display']} ⟶ 🅱️ {row['value_b_display']}",
        axis=1,
    )
    return decisions_df


def export_template(diff: WorkbookDiff, output_path: str | Path, *, base: WorkbookSide, default_action: str | None = None) -> Path:
    action = default_action or default_action_for_base(base)
    return SERVICE.export_decision_template(
        DecisionTemplateRequest(diff=diff, output_path=output_path, default_action=action)
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
    compare_mode = decisions.attrs.get("compare_mode", "coordinate")
    header_row = int(decisions.attrs.get("header_row", 1))
    sheet_keys = decisions.attrs.get("sheet_keys", {})
    return SERVICE.apply_decisions(
        MergeRequest(
            workbook_a=workbook_a,
            workbook_b=workbook_b,
            decisions=decisions,
            output_path=output_path,
            base=base,
            include_sheets_from_source_only=include_sheets_from_source_only,
            compare_mode=compare_mode,
            header_row=header_row,
            sheet_keys={sheet: list(columns) for sheet, columns in sheet_keys.items()},
        )
    )


def format_display_value(value: object) -> str:
    if value is None:
        return "∅"
    if value == "":
        return "''"
    return str(value)
