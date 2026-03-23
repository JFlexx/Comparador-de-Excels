from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from comparator import (
    CompareOptions,
    VALID_COMPARE_MODES,
    apply_decisions,
    compare_workbooks,
    decisions_from_excel,
    export_decision_template,
)


def _parse_sheet_keys(raw_value: str) -> dict[str, list[str]]:
    sheet_keys: dict[str, list[str]] = {}
    for line in raw_value.splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue
        if ':' not in cleaned:
            raise ValueError(
                "Cada línea debe usar el formato Hoja:columna1,columna2"
            )
        sheet_name, raw_columns = cleaned.split(':', 1)
        columns = [column.strip() for column in raw_columns.split(',') if column.strip()]
        if not sheet_name.strip() or not columns:
            raise ValueError(
                "Cada línea debe incluir nombre de hoja y al menos una columna clave"
            )
        sheet_keys[sheet_name.strip()] = columns
    return sheet_keys


st.set_page_config(page_title="Comparador de Excels", layout="wide")
st.title("📘 Comparador de libros Excel (multi-hoja)")
st.caption(
    "Compara dos libros completos y resuelve diferencias por celda o por registros. "
    "Incluye flujo Web (Streamlit) y flujo Excel nativo (plantilla editable)."
)


DISPLAY_COLUMNS = [
    "context",
    "value_a_display",
    "value_b_display",
    "preview",
    "action",
    "manual_value",
    "reviewed",
]
EDITABLE_COLUMNS = ["action", "manual_value", "reviewed"]


def _format_value(value: object) -> str:
    if value is None:
        return "∅"
    if value == "":
        return "''"
    return str(value)


with st.sidebar:
    st.header("Opciones de comparación")
    compare_mode = st.selectbox(
        "Modo de comparación",
        options=list(VALID_COMPARE_MODES),
        format_func=lambda value: "Por coordenadas" if value == "coordinate" else "Por filas/estructura",
        help=(
            "Por coordenadas compara celda contra celda. "
            "Por filas usa encabezados y columnas clave para detectar altas, bajas y cambios sin cascadas."
        ),
    )
    ignore_case = st.checkbox("Ignorar mayúsculas/minúsculas", value=False)
    keep_spaces = st.checkbox("No recortar espacios", value=False)
    empty_string_is_value = st.checkbox("Distinguir '' y None", value=False)
    header_row = st.number_input("Fila de encabezados", min_value=1, value=1, step=1)
    sheet_keys_raw = st.text_area(
        "Claves por hoja",
        value="",
        height=120,
        help=(
            "Una hoja por línea con formato Hoja:columna1,columna2. "
            "Si una hoja no tiene clave configurada, se comparará usando la fila completa."
        ),
    )

try:
    sheet_keys = _parse_sheet_keys(sheet_keys_raw)
except ValueError as exc:
    st.sidebar.error(str(exc))
    st.stop()

options = CompareOptions(
    strip_strings=not keep_spaces,
    case_sensitive=not ignore_case,
    ignore_empty_string_vs_none=not empty_string_is_value,
    compare_mode=compare_mode,
    sheet_keys=sheet_keys,
    header_row=int(header_row),
)

col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("Excel A (base)", type=["xlsx", "xlsm"], key="a")
with col2:
    file_b = st.file_uploader("Excel B (a comparar)", type=["xlsx", "xlsm"], key="b")

if not (file_a and file_b):
    st.info("Carga ambos archivos para comenzar.")
    st.stop()

temp_dir = Path(tempfile.mkdtemp(prefix="excel_compare_"))
path_a = temp_dir / file_a.name
path_b = temp_dir / file_b.name
path_a.write_bytes(file_a.getbuffer())
path_b.write_bytes(file_b.getbuffer())

diff = compare_workbooks(path_a, path_b, options=options)
comparison_signature = (
    file_a.name,
    file_a.size,
    file_b.name,
    file_b.size,
    options.strip_strings,
    options.case_sensitive,
    options.ignore_empty_string_vs_none,
)

st.subheader("Resumen")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Hojas en común", len(diff.common_sheets))
m2.metric("Hojas solo en A", len(diff.only_in_a))
m3.metric("Hojas solo en B", len(diff.only_in_b))
m4.metric("Diferencias", diff.total_differences)

if diff.only_in_a:
    st.warning(f"Hojas solo en A: {', '.join(diff.only_in_a)}")
if diff.only_in_b:
    st.info(f"Hojas solo en B: {', '.join(diff.only_in_b)}")

summary_df = pd.DataFrame(diff.summary_rows())
if not summary_df.empty:
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

web_tab, excel_tab = st.tabs(["🖥️ Resolver en web", "📗 Resolver en Excel"])

with web_tab:
    if st.session_state.get("comparison_signature") != comparison_signature:
        decisions_df = diff.to_dataframe()
        if not decisions_df.empty:
            decisions_df["value_a_display"] = decisions_df["value_a"].map(_format_value)
            decisions_df["value_b_display"] = decisions_df["value_b"].map(_format_value)
            decisions_df["preview"] = decisions_df.apply(
                lambda row: f"🅰️ {row['value_a_display']} ⟶ 🅱️ {row['value_b_display']}", axis=1
            )
        st.session_state["comparison_signature"] = comparison_signature
        st.session_state["decisions_df"] = decisions_df

    master_df = st.session_state["decisions_df"]

    if master_df.empty:
        st.success("No hay diferencias entre las hojas comunes.")
    else:
        pending_count = int((~master_df["reviewed"]).sum())
        manual_count = int((master_df["action"] == "manual").sum())
        use_b_count = int((master_df["action"] == "use_b").sum())

        st.write(
            "Edita las decisiones agrupadas por hoja. Puedes marcar una fila como revisada "
            "sin cambiar todavía la acción final."
        )
        info1, info2, info3 = st.columns(3)
        info1.metric("Pendientes de revisión", pending_count)
        info2.metric("Acción use_b", use_b_count)
        info3.metric("Acción manual", manual_count)

        filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
        available_sheets = sorted(master_df["sheet"].unique().tolist())
        available_columns = sorted(master_df["column_letter"].unique().tolist())

        with filter_col1:
            selected_sheets = st.multiselect(
                "Filtrar por hoja",
                options=available_sheets,
                default=available_sheets,
            )
        with filter_col2:
            selected_columns = st.multiselect(
                "Filtrar por columna",
                options=available_columns,
                default=available_columns,
            )
        with filter_col3:
            action_filter = st.selectbox(
                "Filtrar por acción",
                options=["all", "use_a", "use_b", "manual"],
                format_func=lambda value: {
                    "all": "Todas",
                    "use_a": "Solo use_a",
                    "use_b": "Solo use_b",
                    "manual": "Solo manual",
                }[value],
            )
        with filter_col4:
            only_pending = st.checkbox("Solo pendientes de revisión", value=False)

        filtered_df = master_df[
            master_df["sheet"].isin(selected_sheets) & master_df["column_letter"].isin(selected_columns)
        ].copy()
        if action_filter != "all":
            filtered_df = filtered_df[filtered_df["action"] == action_filter]
        if only_pending:
            filtered_df = filtered_df[~filtered_df["reviewed"]]

        if filtered_df.empty:
            st.info("No hay diferencias que coincidan con los filtros seleccionados.")
        else:
            for sheet_name in selected_sheets:
                sheet_rows = filtered_df[filtered_df["sheet"] == sheet_name]
                if sheet_rows.empty:
                    continue

                group_summary = diff.grouped_differences[sheet_name]
                pending_sheet = int((~sheet_rows["reviewed"]).sum())
                expander_label = (
                    f"{sheet_name} · {len(sheet_rows)} visibles / {group_summary.total_differences} totales"
                    f" · columnas: {', '.join(group_summary.columns) or '—'}"
                    f" · pendientes visibles: {pending_sheet}"
                )
                with st.expander(expander_label, expanded=True):
                    edited_sheet = st.data_editor(
                        sheet_rows[DISPLAY_COLUMNS],
                        key=f"editor_{sheet_name}",
                        use_container_width=True,
                        num_rows="fixed",
                        column_config={
                            "context": st.column_config.TextColumn("Coordenada y contexto", disabled=True),
                            "value_a_display": st.column_config.TextColumn("Valor en A", disabled=True),
                            "value_b_display": st.column_config.TextColumn("Valor en B", disabled=True),
                            "preview": st.column_config.TextColumn("Previsualización", disabled=True),
                            "action": st.column_config.SelectboxColumn(
                                "Acción", options=["use_a", "use_b", "manual"], required=True
                            ),
                            "manual_value": st.column_config.TextColumn("Valor manual"),
                            "reviewed": st.column_config.CheckboxColumn("Revisada"),
                        },
                        hide_index=True,
                    )
                    master_df.loc[edited_sheet.index, EDITABLE_COLUMNS] = edited_sheet[EDITABLE_COLUMNS]

        include_extra = st.checkbox("Copiar hojas solo existentes en B", value=True)
        output_name = st.text_input("Nombre de salida", "resultado_combinado.xlsx")

        if st.button("Generar Excel combinado (web)"):
            output_path = temp_dir / output_name
            result = apply_decisions(path_a, master_df, output_path, path_b, include_sheets_only_in_b=include_extra)
            st.success("Archivo combinado generado.")
            st.download_button(
                label="Descargar resultado",
                data=result.read_bytes(),
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

with excel_tab:
    st.write(
        "Flujo recomendado si quieres trabajar dentro de Excel: "
        "1) descarga plantilla de decisiones, 2) edítala en Excel, 3) súbela y genera resultado."
    )

    template_name = st.text_input("Nombre plantilla", "decisiones.xlsx")
    template_path = temp_dir / template_name
    export_decision_template(diff, template_path)
    st.download_button(
        label="Descargar plantilla de decisiones",
        data=template_path.read_bytes(),
        file_name=template_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    decisions_file = st.file_uploader(
        "Sube la plantilla de decisiones ya editada",
        type=["xlsx", "xlsm"],
        key="decisions_excel",
    )
    include_extra_excel = st.checkbox("Copiar hojas solo en B (flujo Excel)", value=True)
    output_name_excel = st.text_input("Nombre salida (flujo Excel)", "resultado_excel_flow.xlsx")

    if compare_mode == "row-based":
        st.info(
            "La plantilla exportada en modo por filas sirve para revisar diferencias por registro. "
            "La fusión automática desde plantilla queda reservada al modo por coordenadas."
        )
    elif decisions_file and st.button("Generar Excel combinado (desde plantilla Excel)"):
        decisions_path = temp_dir / decisions_file.name
        decisions_path.write_bytes(decisions_file.getbuffer())

        decisions_df = decisions_from_excel(decisions_path)
        output_path = temp_dir / output_name_excel
        result = apply_decisions(
            path_a,
            decisions_df,
            output_path,
            path_b,
            include_sheets_only_in_b=include_extra_excel,
        )
        st.success("Archivo combinado generado desde plantilla Excel.")
        st.download_button(
            label="Descargar resultado (flujo Excel)",
            data=result.read_bytes(),
            file_name=output_name_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
