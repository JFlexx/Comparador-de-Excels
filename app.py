from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from comparator import (
    CompareOptions,
    apply_decisions,
    compare_workbooks,
    decisions_from_excel,
    diffs_to_dataframe,
    export_decision_template,
)

st.set_page_config(page_title="Comparador de Excels", layout="wide")
st.title("📘 Comparador de libros Excel (multi-hoja)")
st.caption(
    "Compara dos libros completos y resuelve diferencias por celda. "
    "Incluye flujo Web (Streamlit) y flujo Excel nativo (plantilla editable)."
)

with st.sidebar:
    st.header("Opciones de comparación")
    ignore_case = st.checkbox("Ignorar mayúsculas/minúsculas", value=False)
    keep_spaces = st.checkbox("No recortar espacios", value=False)
    empty_string_is_value = st.checkbox("Distinguir '' y None", value=False)

options = CompareOptions(
    strip_strings=not keep_spaces,
    case_sensitive=not ignore_case,
    ignore_empty_string_vs_none=not empty_string_is_value,
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
total_diffs = sum(len(diff.differences[s]) for s in diff.common_sheets)

st.subheader("Resumen")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Hojas en común", len(diff.common_sheets))
m2.metric("Hojas solo en A", len(diff.only_in_a))
m3.metric("Hojas solo en B", len(diff.only_in_b))
m4.metric("Diferencias", total_diffs)

if diff.only_in_a:
    st.warning(f"Hojas solo en A: {', '.join(diff.only_in_a)}")
if diff.only_in_b:
    st.info(f"Hojas solo en B: {', '.join(diff.only_in_b)}")

web_tab, excel_tab = st.tabs(["🖥️ Resolver en web", "📗 Resolver en Excel"])

with web_tab:
    df = diffs_to_dataframe(diff.all_differences())

    if df.empty:
        st.success("No hay diferencias entre las hojas comunes.")
    else:
        st.write("Edita la acción por fila para generar el libro combinado.")
        edited = st.data_editor(
            df,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "action": st.column_config.SelectboxColumn(
                    "action", options=["use_a", "use_b", "manual"], required=True
                ),
                "manual_value": st.column_config.TextColumn("manual_value"),
                "row": st.column_config.NumberColumn(disabled=True),
                "column": st.column_config.NumberColumn(disabled=True),
                "sheet": st.column_config.TextColumn(disabled=True),
                "cell": st.column_config.TextColumn(disabled=True),
                "value_a": st.column_config.TextColumn(disabled=True),
                "value_b": st.column_config.TextColumn(disabled=True),
            },
        )

        include_extra = st.checkbox("Copiar hojas solo existentes en B", value=True)
        output_name = st.text_input("Nombre de salida", "resultado_combinado.xlsx")

        if st.button("Generar Excel combinado (web)"):
            output_path = temp_dir / output_name
            result = apply_decisions(path_a, edited, output_path, path_b, include_sheets_only_in_b=include_extra)
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

    if decisions_file and st.button("Generar Excel combinado (desde plantilla Excel)"):
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
