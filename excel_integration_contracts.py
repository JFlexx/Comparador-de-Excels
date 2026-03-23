from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Literal

from comparator import CompareOptions, WorkbookSide


def _string_path(value: str | Path) -> str:
    return str(Path(value))


@dataclass(frozen=True)
class ExcelWorkbookSelection:
    """Contrato serializable para la selección de libros en Excel Desktop."""

    base_workbook_path: str
    source_workbook_path: str
    base_side: WorkbookSide = "a"

    @classmethod
    def create(
        cls,
        *,
        base_workbook_path: str | Path,
        source_workbook_path: str | Path,
        base_side: WorkbookSide = "a",
    ) -> "ExcelWorkbookSelection":
        return cls(
            base_workbook_path=_string_path(base_workbook_path),
            source_workbook_path=_string_path(source_workbook_path),
            base_side=base_side,
        )

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ExcelWorkbookSelection":
        return cls.create(
            base_workbook_path=payload["base_workbook_path"],
            source_workbook_path=payload["source_workbook_path"],
            base_side=payload.get("base_side", "a"),
        )

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class ExcelCompareContract:
    """Entrada serializable para solicitar una comparación desde el add-in."""

    selection: ExcelWorkbookSelection
    compare_mode: Literal["coordinate", "row-based"] = "coordinate"
    header_row: int = 1
    sheet_keys: dict[str, list[str]] = field(default_factory=dict)
    strip_strings: bool = True
    case_sensitive: bool = True
    ignore_empty_string_vs_none: bool = True

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ExcelCompareContract":
        return cls(
            selection=ExcelWorkbookSelection.from_dict(payload["selection"]),
            compare_mode=payload.get("compare_mode", "coordinate"),
            header_row=int(payload.get("header_row", 1)),
            sheet_keys={
                str(sheet): [str(column) for column in columns]
                for sheet, columns in payload.get("sheet_keys", {}).items()
            },
            strip_strings=bool(payload.get("strip_strings", True)),
            case_sensitive=bool(payload.get("case_sensitive", True)),
            ignore_empty_string_vs_none=bool(payload.get("ignore_empty_string_vs_none", True)),
        )

    def to_compare_options(self) -> CompareOptions:
        return CompareOptions(
            compare_mode=self.compare_mode,
            header_row=self.header_row,
            sheet_keys={sheet: list(columns) for sheet, columns in self.sheet_keys.items()},
            strip_strings=self.strip_strings,
            case_sensitive=self.case_sensitive,
            ignore_empty_string_vs_none=self.ignore_empty_string_vs_none,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            **asdict(self),
            "selection": self.selection.to_dict(),
        }


@dataclass(frozen=True)
class ExcelDecisionRow:
    """Fila serializable del contrato de decisiones consumible por un add-in."""

    sheet: str
    cell: str
    row: int
    column: int
    column_letter: str
    context: str
    header: str | None
    key: str | None
    diff_type: str
    value_a: Any
    value_b: Any
    action: str
    manual_value: Any
    reviewed: bool

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ExcelDecisionRow":
        return cls(**payload)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class ExcelSheetSummary:
    sheet: str
    differences: int
    columns: str
    rows: int

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ExcelSheetSummary":
        return cls(**payload)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class ExcelComparisonResult:
    """Salida serializable que oculta WorkbookDiff al add-in de Excel."""

    route: str
    selection: ExcelWorkbookSelection
    compare_mode: str
    header_row: int
    total_differences: int
    common_sheets: list[str]
    only_in_base: list[str]
    only_in_source: list[str]
    default_action: str
    summary_rows: list[ExcelSheetSummary]
    decision_rows: list[ExcelDecisionRow]

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ExcelComparisonResult":
        return cls(
            route=payload["route"],
            selection=ExcelWorkbookSelection.from_dict(payload["selection"]),
            compare_mode=payload["compare_mode"],
            header_row=int(payload["header_row"]),
            total_differences=int(payload["total_differences"]),
            common_sheets=[str(value) for value in payload.get("common_sheets", [])],
            only_in_base=[str(value) for value in payload.get("only_in_base", [])],
            only_in_source=[str(value) for value in payload.get("only_in_source", [])],
            default_action=str(payload["default_action"]),
            summary_rows=[ExcelSheetSummary.from_dict(row) for row in payload.get("summary_rows", [])],
            decision_rows=[ExcelDecisionRow.from_dict(row) for row in payload.get("decision_rows", [])],
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "route": self.route,
            "selection": self.selection.to_dict(),
            "compare_mode": self.compare_mode,
            "header_row": self.header_row,
            "total_differences": self.total_differences,
            "common_sheets": list(self.common_sheets),
            "only_in_base": list(self.only_in_base),
            "only_in_source": list(self.only_in_source),
            "default_action": self.default_action,
            "summary_rows": [row.to_dict() for row in self.summary_rows],
            "decision_rows": [row.to_dict() for row in self.decision_rows],
        }


@dataclass(frozen=True)
class ExcelDecisionSheetContract:
    """Entrada serializable para materializar la tabla de decisiones en un workbook."""

    workbook_path: str
    sheet_name: str = "DecisionTable"
    clear_sheet: bool = True

    @classmethod
    def create(
        cls,
        workbook_path: str | Path,
        sheet_name: str = "DecisionTable",
        clear_sheet: bool = True,
    ) -> "ExcelDecisionSheetContract":
        return cls(
            workbook_path=_string_path(workbook_path),
            sheet_name=sheet_name,
            clear_sheet=clear_sheet,
        )

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ExcelDecisionSheetContract":
        return cls.create(
            workbook_path=payload["workbook_path"],
            sheet_name=payload.get("sheet_name", "DecisionTable"),
            clear_sheet=bool(payload.get("clear_sheet", True)),
        )

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class ExcelDecisionSheetLoadResult:
    workbook_path: str
    sheet_name: str
    rows_written: int

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class ExcelDecisionsReadResult:
    sheet_name: str
    rows_loaded: int
    decisions: list[ExcelDecisionRow]

    def to_dict(self) -> dict[str, Any]:
        return {
            "sheet_name": self.sheet_name,
            "rows_loaded": self.rows_loaded,
            "decisions": [decision.to_dict() for decision in self.decisions],
        }


@dataclass(frozen=True)
class ExcelMergeContract:
    selection: ExcelWorkbookSelection
    decisions_workbook_path: str
    decisions_sheet_name: str = "DecisionTable"
    output_path: str = "resultado_combinado.xlsx"
    include_sheets_from_source_only: bool = True

    @classmethod
    def create(
        cls,
        *,
        selection: ExcelWorkbookSelection,
        decisions_workbook_path: str | Path,
        decisions_sheet_name: str = "DecisionTable",
        output_path: str | Path = "resultado_combinado.xlsx",
        include_sheets_from_source_only: bool = True,
    ) -> "ExcelMergeContract":
        return cls(
            selection=selection,
            decisions_workbook_path=_string_path(decisions_workbook_path),
            decisions_sheet_name=decisions_sheet_name,
            output_path=_string_path(output_path),
            include_sheets_from_source_only=include_sheets_from_source_only,
        )

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ExcelMergeContract":
        return cls.create(
            selection=ExcelWorkbookSelection.from_dict(payload["selection"]),
            decisions_workbook_path=payload["decisions_workbook_path"],
            decisions_sheet_name=payload.get("decisions_sheet_name", "DecisionTable"),
            output_path=payload.get("output_path", "resultado_combinado.xlsx"),
            include_sheets_from_source_only=bool(payload.get("include_sheets_from_source_only", True)),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            **asdict(self),
            "selection": self.selection.to_dict(),
        }


@dataclass(frozen=True)
class ExcelMergeResult:
    output_path: str
    base_workbook_path: str
    source_workbook_path: str
    decisions_sheet_name: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)
