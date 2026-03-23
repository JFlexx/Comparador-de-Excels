from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from comparator import WorkbookSide


@dataclass(frozen=True)
class MergeDescriptor:
    base_key: WorkbookSide
    source_key: WorkbookSide
    base_label: str
    source_label: str
    default_action: str


@dataclass(frozen=True)
class ReviewTable:
    dataframe: pd.DataFrame
    editable_columns: tuple[str, ...]
    display_columns: tuple[str, ...]


@dataclass(frozen=True)
class CliCompareReport:
    template_path: Path
    compare_mode: str
    common_sheets: int
    only_in_a: int
    only_in_b: int
    total_differences: int
    default_action: str
    merge: MergeDescriptor


@dataclass(frozen=True)
class CliMergeReport:
    output_path: Path
    merge: MergeDescriptor
