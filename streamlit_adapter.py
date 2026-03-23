from __future__ import annotations

from pathlib import Path

import pandas as pd

from adapter_models import MergeDescriptor, ReviewTable
from comparator import CompareOptions, WorkbookDiff, WorkbookSide
from interface_adapter import (
    build_compare_options,
    compare_files,
    describe_merge,
    export_template,
    load_decisions,
    merge_workbooks,
    parse_sheet_keys,
)

DISPLAY_COLUMNS = (
    "sheet",
    "cell",
    "header",
    "key",
    "diff_type",
    "value_a_display",
    "value_b_display",
    "preview",
    "action",
    "manual_value",
    "reviewed",
)
EDITABLE_COLUMNS = ("action", "manual_value", "reviewed")


def parse_sheet_keys_block(raw_value: str) -> dict[str, list[str]]:
    return parse_sheet_keys(raw_value.splitlines(), item_separator=":")


def format_display_value(value: object) -> str:
    if value is None:
        return "∅"
    if value == "":
        return "''"
    return str(value)


def build_review_table(diff: WorkbookDiff, *, default_action: str) -> ReviewTable:
    decisions_df = diff.to_dataframe(default_action=default_action)
    if decisions_df.empty:
        return ReviewTable(
            dataframe=decisions_df,
            editable_columns=EDITABLE_COLUMNS,
            display_columns=DISPLAY_COLUMNS,
        )

    reviewed_df = decisions_df.copy()
    reviewed_df["value_a_display"] = reviewed_df["value_a"].map(format_display_value)
    reviewed_df["value_b_display"] = reviewed_df["value_b"].map(format_display_value)
    reviewed_df["preview"] = reviewed_df.apply(
        lambda row: f"A: {row['value_a_display']} -> B: {row['value_b_display']}",
        axis=1,
    )
    return ReviewTable(
        dataframe=reviewed_df,
        editable_columns=EDITABLE_COLUMNS,
        display_columns=DISPLAY_COLUMNS,
    )


def build_streamlit_context(base: WorkbookSide) -> MergeDescriptor:
    return describe_merge(base)


__all__ = [
    "CompareOptions",
    "DISPLAY_COLUMNS",
    "EDITABLE_COLUMNS",
    "MergeDescriptor",
    "Path",
    "ReviewTable",
    "WorkbookDiff",
    "WorkbookSide",
    "build_compare_options",
    "build_review_table",
    "build_streamlit_context",
    "compare_files",
    "export_template",
    "load_decisions",
    "merge_workbooks",
    "parse_sheet_keys_block",
]
