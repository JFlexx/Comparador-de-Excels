from pathlib import Path

from openpyxl import Workbook, load_workbook

from comparator import (
    CompareOptions,
    apply_decisions,
    compare_workbooks,
    decisions_from_excel,
    diffs_to_dataframe,
    export_decision_template,
)


def _create_wb(path: Path, sheets: dict[str, dict[str, object]]):
    wb = Workbook()
    wb.remove(wb.active)
    for sheet, cells in sheets.items():
        ws = wb.create_sheet(sheet)
        for coord, value in cells.items():
            ws[coord] = value
    wb.save(path)


def test_compare_workbooks_detects_sheet_and_cell_differences(tmp_path: Path):
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"

    _create_wb(
        a,
        {
            "Resumen": {"A1": "ID", "A2": 1, "B2": "Ana"},
            "SoloA": {"A1": "x"},
        },
    )
    _create_wb(
        b,
        {
            "Resumen": {"A1": "ID", "A2": 1, "B2": "Anita"},
            "SoloB": {"A1": "y"},
        },
    )

    diff = compare_workbooks(a, b)

    assert diff.only_in_a == ["SoloA"]
    assert diff.only_in_b == ["SoloB"]
    assert diff.common_sheets == ["Resumen"]
    assert len(diff.differences["Resumen"]) == 1
    assert diff.differences["Resumen"][0].coordinate == "B2"


def test_compare_options_ignore_case_and_spaces(tmp_path: Path):
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"

    _create_wb(a, {"Datos": {"A1": "  Ana  "}})
    _create_wb(b, {"Datos": {"A1": "ana"}})

    default_diff = compare_workbooks(a, b)
    assert len(default_diff.all_differences()) == 1

    relaxed_diff = compare_workbooks(
        a,
        b,
        options=CompareOptions(strip_strings=True, case_sensitive=False),
    )
    assert relaxed_diff.all_differences() == []


def test_compare_workbooks_row_based_avoids_cascade_and_detects_row_states(tmp_path: Path):
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"

    _create_wb(
        a,
        {
            "Datos": {
                "A1": "ID",
                "B1": "Nombre",
                "C1": "Estado",
                "A2": 1,
                "B2": "Ana",
                "C2": "OK",
                "A3": 2,
                "B3": "Luis",
                "C3": "OK",
            }
        },
    )
    _create_wb(
        b,
        {
            "Datos": {
                "A1": "ID",
                "B1": "Nombre",
                "C1": "Estado",
                "A2": 1,
                "B2": "Ana",
                "C2": "Actualizado",
                "A3": 3,
                "B3": "Marta",
                "C3": "Nueva",
                "A4": 2,
                "B4": "Luis",
                "C4": "OK",
            }
        },
    )

    coordinate_diff = compare_workbooks(a, b)
    assert len(coordinate_diff.all_differences()) == 6

    row_diff = compare_workbooks(
        a,
        b,
        options=CompareOptions(compare_mode="row-based", sheet_keys={"Datos": ["ID"]}),
    )

    row_diffs = row_diff.differences["Datos"]
    assert len(row_diffs) == 4
    assert {diff.diff_type for diff in row_diffs} == {"modified", "added"}
    modified = [diff for diff in row_diffs if diff.diff_type == "modified"]
    added = [diff for diff in row_diffs if diff.diff_type == "added"]
    assert len(modified) == 1
    assert modified[0].header == "Estado"
    assert modified[0].key == "ID=1"
    assert {diff.header for diff in added} == {"ID", "Nombre", "Estado"}
    assert {diff.key for diff in added} == {"ID=3"}


def test_apply_decisions_merges_changes(tmp_path: Path):
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"

    _create_wb(a, {"Datos": {"A1": "hola", "A2": "base"}})
    _create_wb(b, {"Datos": {"A1": "hola", "A2": "nuevo"}, "Nueva": {"A1": 99}})

    diff = compare_workbooks(a, b)
    df = diffs_to_dataframe(diff.all_differences())

    output = tmp_path / "out.xlsx"
    apply_decisions(a, df, output, b, base="a")

    wb = load_workbook(output)
    assert wb["Datos"]["A2"].value == "nuevo"
    assert wb["Nueva"]["A1"].value == 99



def test_apply_decisions_merges_changes_from_a_onto_b(tmp_path: Path):
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"

    _create_wb(a, {"Datos": {"A1": "hola", "A2": "nuevo"}, "NuevaA": {"A1": 42}})
    _create_wb(b, {"Datos": {"A1": "hola", "A2": "base"}})

    diff = compare_workbooks(a, b)
    df = diffs_to_dataframe(diff.all_differences())

    output = tmp_path / "out_b.xlsx"
    apply_decisions(a, df, output, b, base="b")

    wb = load_workbook(output)
    assert wb["Datos"]["A2"].value == "nuevo"
    assert wb["NuevaA"]["A1"].value == 42


def test_export_and_read_decisions_template(tmp_path: Path):
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"

    _create_wb(a, {"Datos": {"A1": "x"}})
    _create_wb(b, {"Datos": {"A1": "y"}})

    diff = compare_workbooks(a, b)
    template = tmp_path / "decisiones.xlsx"
    export_decision_template(diff, template)

    loaded = decisions_from_excel(template)
    assert loaded.shape[0] == 1
    assert loaded.iloc[0]["action"] == "use_b"
    assert loaded.iloc[0]["sheet"] == "Datos"
    assert loaded.iloc[0]["diff_type"] == "modified"
