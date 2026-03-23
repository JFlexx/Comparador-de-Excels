from __future__ import annotations

from pathlib import Path

import pandas as pd

from comparator import WorkbookDiff, WorkbookSide
from interface_adapter import export_template, load_decisions, merge_workbooks


def export_decisions_workbook(
    diff: WorkbookDiff,
    output_path: str | Path,
    *,
    default_action: str,
) -> Path:
    return export_template(diff=diff, output_path=output_path, default_action=default_action)


def import_decisions_workbook(path: str | Path) -> pd.DataFrame:
    return load_decisions(path)


def merge_from_decisions_workbook(
    *,
    workbook_a: str | Path,
    workbook_b: str | Path,
    decisions: pd.DataFrame,
    output_path: str | Path,
    base: WorkbookSide,
    include_sheets_from_source_only: bool,
) -> Path:
    return merge_workbooks(
        workbook_a=workbook_a,
        workbook_b=workbook_b,
        decisions=decisions,
        output_path=output_path,
        base=base,
        include_sheets_from_source_only=include_sheets_from_source_only,
    )
