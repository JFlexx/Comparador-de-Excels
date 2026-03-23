from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from comparator import VALID_COMPARE_MODES
from excel_adapter import export_decisions_workbook, import_decisions_workbook, merge_from_decisions_workbook
from streamlit_adapter import (
    build_compare_options,
    build_review_table,
    build_streamlit_context,
    compare_files,
    parse_sheet_keys_block,
)

st.set_page_config(page_title="Comparador de Excels", layout="wide")
st.title("📘 Comparador de libros Excel (multi-hoja)")
st.caption(
    "La interfaz Streamlit consume adaptadores específicos y deja el motor de dominio aislado en comparator.py."
)

with st.sidebar:
    st.header("Opciones de comparación")
    compare_mode = st.selectbox(
        "Modo de comparación",
        options=sorted(VALID_COMPARE_MODES),
        format_func=lambda value: "Por coordenadas" if value == "coordinate" else "Por filas/estructura",
        help=(
            "Por coordenadas compara celda contra celda. "
            "Por filas usa encabezados y columnas clave para detectar altas, bajas y cambios sin cascadas."
        ),
    )
    merge_direction = st.radio(
        "Libro base del merge final",
        options=["a", "b"],
        format_func=lambda value: "Construir resultado sobre A" if value == "a" else "Construir resultado sobre B",
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
    sheet_keys = parse_sheet_keys_block(sheet_keys_raw)
    options = build_compare_options(
        compare_mode=compare_mode,
        ignore_case=ignore_case,
        keep_spaces=keep_spaces,
        empty_string_is_value=empty_string_is_value,
        header_row=int(header_row),
        sheet_keys=sheet_keys,
    )
except ValueError as exc:
    st.sidebar.error(str(exc))
    st.stop()

merge_context = build_streamlit_context(merge_direction)

col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("Excel A", type=["xlsx", "xlsm"], key="a")
with col2:
    file_b = st.file_uploader("Excel B", type=["xlsx", "xlsm"], key="b")

if not (file_a and file_b):
    st.info("Carga ambos archivos para comenzar.")
    st.stop()

temp_dir = Path(tempfile.mkdtemp(prefix="excel_compare_"))
path_a = temp_dir / file_a.name
path_b = temp_dir / file_b.name
path_a.write_bytes(file_a.getbuffer())
path_b.write_bytes(file_b.getbuffer())

diff = compare_files(path_a, path_b, options)
comparison_signature = (
    file_a.name,
    file_a.size,
    file_b.name,
    file_b.size,
    options.compare_mode,
    options.header_row,
    tuple(sorted((sheet, tuple(columns)) for sheet, columns in options.sheet_keys.items())),
    options.strip_strings,
    options.case_sensitive,
    options.ignore_empty_string_vs_none,
    merge_direction,
)

st.subheader("Resumen")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Hojas en común", len(diff.common_sheets))
m2.metric("Hojas solo en A", len(diff.only_in_a))
m3.metric("Hojas solo en B", len(diff.only_in_b))
m4.metric("Diferencias", diff.total_differences)

st.caption(
    f"Merge objetivo: traer cambios de {merge_context.source_label} hacia {merge_context.base_label}. "
    f"Acción por defecto de la plantilla: {merge_context.default_action}."
)

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
        st.session_state["comparison_signature"] = comparison_signature
        st.session_state["review_table"] = build_review_table(
            diff,
            default_action=merge_context.default_action,
        )

    review_table = st.session_state["review_table"]
    master_df = review_table.dataframe

    if master_df.empty:
        st.success("No hay diferencias entre las hojas comunes.")
    else:
        pending_count = int((~master_df["reviewed"]).sum())
        manual_count = int((master_df["action"] == "manual").sum())
        source_count = int((master_df["action"] == merge_context.default_action).sum())

        st.write(
            "Edita las decisiones agrupadas por hoja. La UI solo modifica el DataFrame de revisión; "
            "el merge final sigue delegándose al adaptador Excel."
        )
        info1, info2, info3 = st.columns(3)
        info1.metric("Pendientes de revisión", pending_count)
        info2.metric(f"Acción {merge_context.default_action}", source_count)
        info3.metric("Acción manual", manual_count)

        filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
        available_sheets = sorted(master_df["sheet"].unique().tolist())
        available_columns = sorted(master_df["column_letter"].unique().tolist())

        with filter_col1:
            selected_sheets = st.multiselect("Filtrar por hoja", options=available_sheets, default=available_sheets)
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
                        sheet_rows[list(review_table.display_columns)],
                        key=f"editor_{sheet_name}",
                        use_container_width=True,
                        num_rows="fixed",
                        column_config={
                            "sheet": st.column_config.TextColumn("Hoja", disabled=True),
                            "cell": st.column_config.TextColumn("Celda", disabled=True),
                            "header": st.column_config.TextColumn("Header", disabled=True),
                            "key": st.column_config.TextColumn("Clave registro", disabled=True),
                            "diff_type": st.column_config.TextColumn("Tipo diff", disabled=True),
                            "value_a_display": st.column_config.TextColumn("Valor en A", disabled=True),
                            "value_b_display": st.column_config.TextColumn("Valor en B", disabled=True),
                            "preview": st.column_config.TextColumn("Previsualización", disabled=True),
                            "action": st.column_config.SelectboxColumn(
                                "Acción",
                                options=["use_a", "use_b", "manual"],
                                required=True,
                            ),
                            "manual_value": st.column_config.TextColumn("Valor manual"),
                            "reviewed": st.column_config.CheckboxColumn("Revisada"),
                        },
                        hide_index=True,
                    )
                    master_df.loc[edited_sheet.index, list(review_table.editable_columns)] = edited_sheet[
                        list(review_table.editable_columns)
                    ]

        if compare_mode == "row-based":
            st.info(
                "El modo row-based ahora permite merge por registro: "
                "el motor usa key/header/diff_type para ubicar cada decisión aunque las filas se hayan movido."
            )

        include_extra = st.checkbox(f"Copiar hojas solo existentes en {labels.source}", value=True)
        output_name = st.text_input("Nombre de salida", "resultado_combinado.xlsx")

        if st.button("Generar Excel combinado (web)"):
            output_path = temp_dir / output_name
            result = merge_workbooks(
                workbook_a=path_a,
                workbook_b=path_b,
                decisions=master_df,
                output_path=output_path,
                base=merge_direction,
                include_sheets_from_source_only=include_extra,
            )
            st.success("Archivo combinado generado.")
            st.download_button(
                label="Descargar resultado",
                data=result.read_bytes(),
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            include_extra = st.checkbox(
                f"Copiar hojas solo existentes en {merge_context.source_label}",
                value=True,
            )
            output_name = st.text_input("Nombre de salida", "resultado_combinado.xlsx")

            if st.button("Generar Excel combinado (web)"):
                output_path = temp_dir / output_name
                result = merge_from_decisions_workbook(
                    workbook_a=path_a,
                    workbook_b=path_b,
                    decisions=master_df,
                    output_path=output_path,
                    base=merge_direction,
                    include_sheets_from_source_only=include_extra,
                )
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
        f"1) comparar, 2) descargar plantilla con acción por defecto para traer cambios de {merge_context.source_label} hacia {merge_context.base_label}, "
        "3) editar decisiones, 4) subir plantilla y solicitar el merge final."
    )
    if compare_mode == "row-based":
        st.caption(
            "En row-based, la plantilla conserva metadata del diff para que el merge use la identidad lógica del registro."
        )

    template_name = st.text_input("Nombre plantilla", "decisiones.xlsx")
    template_path = temp_dir / template_name
    export_decisions_workbook(diff, template_path, default_action=merge_context.default_action)
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
    include_extra_excel = st.checkbox(
        f"Copiar hojas solo en {merge_context.source_label} (flujo Excel)",
        value=True,
    )
    output_name_excel = st.text_input("Nombre salida (flujo Excel)", "resultado_excel_flow.xlsx")

    if compare_mode == "row-based":
        st.info(
            "La plantilla exportada en modo row-based sirve para revisar diferencias por registro. "
            "La fusión automática desde plantilla queda reservada al modo coordinate."
        )
    elif decisions_file and st.button("Generar Excel combinado (desde plantilla Excel)"):
        decisions_path = temp_dir / decisions_file.name
        decisions_path.write_bytes(decisions_file.getbuffer())

        decisions_df = import_decisions_workbook(decisions_path)
        output_path = temp_dir / output_name_excel
        result = merge_from_decisions_workbook(
            workbook_a=path_a,
            workbook_b=path_b,
            decisions=decisions_df,
            output_path=output_path,
            base=merge_direction,
            include_sheets_from_source_only=include_extra_excel,
        )
        st.success("Archivo combinado generado desde plantilla Excel.")
        st.download_button(
            label="Descargar resultado (flujo Excel)",
            data=result.read_bytes(),
            file_name=output_name_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
