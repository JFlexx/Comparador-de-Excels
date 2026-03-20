from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font

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
    assert diff.differences["Resumen"][0].difference_types == ("value_changed",)



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



def test_compare_workbooks_can_focus_on_formulas_styles_and_comments(tmp_path: Path):
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"

    wb_a = Workbook()
    ws_a = wb_a.active
    ws_a.title = "Datos"
    ws_a["A1"] = "=SUM(1,1)"
    ws_a["A1"].font = Font(bold=True)
    ws_a["A1"].comment = Comment("comentario A", "QA")
    wb_a.save(a)

    wb_b = Workbook()
    ws_b = wb_b.active
    ws_b.title = "Datos"
    ws_b["A1"] = "=SUM(1,2)"
    ws_b["A1"].font = Font(italic=True)
    ws_b["A1"].comment = Comment("comentario B", "QA")
    wb_b.save(b)

    diff = compare_workbooks(
        a,
        b,
        options=CompareOptions(
            compare_values=False,
            compare_formulas=True,
            compare_styles=True,
            compare_comments=True,
        ),
    )

    assert len(diff.all_differences()) == 1
    assert diff.all_differences()[0].difference_types == (
        "formula_changed",
        "style_changed",
        "comment_changed",
    )



def test_compare_options_require_at_least_one_enabled_mode(tmp_path: Path):
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"

    _create_wb(a, {"Datos": {"A1": "x"}})
    _create_wb(b, {"Datos": {"A1": "y"}})

    try:
        compare_workbooks(
            a,
            b,
            options=CompareOptions(
                compare_values=False,
                compare_formulas=False,
                compare_cached_results=False,
                compare_styles=False,
                compare_comments=False,
            ),
        )
    except ValueError as exc:
        assert "al menos un tipo" in str(exc)
    else:
        raise AssertionError("Se esperaba ValueError cuando no hay modos activos")



def test_apply_decisions_merges_changes(tmp_path: Path):
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"

    _create_wb(a, {"Datos": {"A1": "hola", "A2": "base"}})
    _create_wb(b, {"Datos": {"A1": "hola", "A2": "nuevo"}, "Nueva": {"A1": 99}})

    diff = compare_workbooks(a, b)
    df = diffs_to_dataframe(diff.all_differences())

    output = tmp_path / "out.xlsx"
    apply_decisions(a, df, output, b)

    wb = load_workbook(output)
    assert wb["Datos"]["A2"].value == "nuevo"
    assert wb["Nueva"]["A1"].value == 99



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
    assert loaded.iloc[0]["difference_types"] == "value_changed"
    assert "formula_a" in loaded.columns
    assert "cached_value_a" in loaded.columns
