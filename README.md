# Comparador de Excels (arquitectura Excel-first)

Herramienta interna en Python para comparar dos libros Excel completos, revisar diferencias y generar un merge final usando `comparator.py` como **motor único**.

## Interfaz principal objetivo

La interfaz principal pasa a ser **Excel Desktop** mediante un add-in con **`xlwings` + Ribbon + runtime local de Python**. El add-in no habla con `WorkbookDiff` ni con objetos internos del motor: consume contratos serializables del adaptador `excel_addin_adapter.py` y deja toda la lógica de negocio en `comparator.py`.

### Ruta concreta de integración

**Opción elegida:** `xlwings` / Ribbon + Python local.

Motivos:
- permite trabajar directamente sobre libros abiertos en Excel Desktop,
- evita exponer el motor como servicio adicional para un caso inicialmente local,
- simplifica la carga y lectura de la tabla de decisiones en hojas reales,
- reutiliza el mismo runtime Python que ya ejecuta el comparador y sus dependencias.

Flujo previsto:
1. El usuario pulsa acciones del Ribbon para seleccionar **libro base** y **libro origen**.
2. El add-in construye un payload serializable y llama a `ExcelAddinAdapter.compare` o `compare_payload`.
3. El adaptador usa `ComparatorService` para comparar y devuelve un resultado serializable con resumen + filas de decisión.
4. El add-in llama a `load_decision_table_into_workbook` para materializar la tabla en una hoja de Excel.
5. Tras la edición manual, el add-in llama a `read_decisions_from_workbook`.
6. Finalmente invoca `execute_merge` para generar el libro combinado.

---

## Arquitectura objetivo

### Motor único

- `comparator.py`: núcleo estable del dominio. Expone `ComparatorService`, contratos de comparación, exportación/lectura de decisiones y merge final.
- `excel_addin_adapter.py`: adaptador Excel-first orientado a Excel Desktop. Traduce operaciones de alto nivel del usuario a llamadas al motor.
- `excel_integration_contracts.py`: contratos serializables para que el add-in trabaje con `dict`/JSON-friendly payloads sin conocer `WorkbookDiff`.
- `app.py`: interfaz **legacy/demo** en Streamlit; ya no es la interfaz principal.
- `excel_tool.py`: CLI auxiliar para automatización o soporte.

### Principio de diseño

Las interfaces solo deben:
1. capturar rutas, opciones y decisiones del usuario,
2. convertirlas a contratos serializables,
3. delegar en `ComparatorService` a través del adaptador,
4. mostrar o persistir resultados.

La lógica de diff, validación, normalización de decisiones y merge final debe seguir residiendo en `comparator.py`.

- Estás comparando listados tabulares con encabezados (`ID`, `Código`, `Email`, etc.).
- Puede haber **altas, bajas o inserciones intermedias** de filas.
- Quieres evitar el efecto cascada típico de los diffs por coordenadas cuando una fila nueva desplaza todas las siguientes.
- Puedes definir una o varias columnas clave por hoja, por ejemplo `Clientes:ID` o `Pedidos:Empresa,NúmeroPedido`.
- También quieres **fusionar** esos cambios al libro final sin depender de la coordenada original de Excel.

## Contratos serializables para Excel

- Toma la fila de encabezados (por defecto la fila 1).
- Empareja filas por las columnas clave configuradas para esa hoja.
- Si una hoja no tiene clave configurada, usa el contenido completo de la fila como identidad implícita.
- Reporta diferencias con tipo `added`, `deleted` o `modified`.
- La plantilla de decisiones y el merge consumen `sheet`, `key`, `header` y `diff_type` para reubicar el registro al aplicar cambios.

> **Nota:** para que el merge `row-based` sea estable y predecible, conviene definir `sheet_keys` por hoja. Si no se define clave, el motor puede caer en coordenadas/fila de apoyo cuando no pueda reconstruir una identidad lógica suficiente.

`ExcelWorkbookSelection`
- `base_workbook_path`
- `source_workbook_path`
- `base_side`

### Solicitud de comparación

`ExcelCompareContract`
- `selection`
- `compare_mode`
- `header_row`
- `sheet_keys`
- banderas de normalización (`strip_strings`, `case_sensitive`, `ignore_empty_string_vs_none`)

### Resultado de comparación

`ExcelComparisonResult`
- `route`
- `selection`
- `total_differences`
- `common_sheets`
- `only_in_base`
- `only_in_source`
- `default_action`
- `summary_rows`
- `decision_rows`

### Tabla de decisiones en hoja de Excel

`ExcelDecisionSheetContract`
- `workbook_path`
- `sheet_name`
- `clear_sheet`

### Merge final

`ExcelMergeContract`
- `selection`
- `decisions_workbook_path`
- `decisions_sheet_name`
- `output_path`
- `include_sheets_from_source_only`

Todos estos contratos tienen representación `dict`, por lo que un add-in puede serializarlos fácilmente.

Esto genera `decisiones.xlsx` con:
- Hoja `Decisiones`: una fila por diferencia.
- Columnas auxiliares como `header`, `diff_type` y `key` para identificar el registro afectado.
- Metadata del diff en la hoja `Resumen` (`compare_mode`, `header_row`, `sheet_keys`) para permitir merge posterior también en `row-based`.
- Columna `action` con lista desplegable (`use_a`, `use_b`, `manual`).
- Columna `manual_value` para casos manuales.

### 2) Editar decisiones en Excel

## Operaciones de alto nivel del adaptador Excel

`ExcelAddinAdapter` expone el flujo pedido para Excel:

1. **Seleccionar libro base y libro origen**
   - `select_workbooks(...)`
2. **Comparar**
   - `compare(contract)` o `compare_payload(payload)`
3. **Cargar tabla de decisiones en una hoja de trabajo**
   - `load_decision_table_into_workbook(comparison, target)`
4. **Volver a leer decisiones desde Excel**
   - `read_decisions_from_workbook(target)`
5. **Ejecutar merge final**
   - `execute_merge(contract)`

---

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

En la interfaz web (`streamlit run app.py`) ahora también puedes elegir visualmente la dirección del merge antes de generar el archivo final. Esto aplica tanto a `coordinate` como a `row-based`.

---

## Uso recomendado: Excel Desktop

### 1) Comparar desde el add-in

El add-in construye un payload similar a este:

```python
payload = {
    "selection": {
        "base_workbook_path": "libro_base.xlsx",
        "source_workbook_path": "libro_origen.xlsx",
        "base_side": "a",
    },
    "compare_mode": "coordinate",
    "header_row": 1,
    "sheet_keys": {},
}
```

Y llama al adaptador:

- `comparator.py`: núcleo estable del motor. Expone la API de servicio (`ComparatorService`) y las funciones públicas `compare_workbooks`, `export_decision_template`, `decisions_from_excel` y `apply_decisions`. También documenta los contratos de entrada/salida para rutas, DataFrames de decisiones, acciones válidas y modos `coordinate` / `row-based`.
- `interface_adapter.py`: adaptador compartido con parsing de opciones, DTOs de dirección de merge y acceso estable al servicio.
- `streamlit_adapter.py`: adaptador específico de Streamlit; prepara tablas de revisión y wording visual sin contaminar el motor.
- `cli_adapter.py`: adaptador específico de CLI; transforma argumentos/reportes sin mezclar reglas de negocio.
- `excel_adapter.py`: adaptador específico para exportar/importar decisiones y ejecutar merges desde flujos Excel.
- `app.py`: interfaz web en Streamlit; consume `streamlit_adapter.py` y `excel_adapter.py`.
- `excel_tool.py`: CLI para flujo Excel-first; consume `cli_adapter.py`.
- `tests/test_comparator.py`: pruebas unitarias del núcleo y de los contratos principales.

adapter = ExcelAddinAdapter()
comparison = adapter.compare_payload(payload)
```

### 2) Volcar la tabla de decisiones a una hoja

```python
adapter.load_decision_table_payload(
    comparison,
    {
        "workbook_path": "host_decisiones.xlsx",
        "sheet_name": "DecisionTable",
        "clear_sheet": True,
    },
)
```

### 3) Leer decisiones editadas y ejecutar merge

```python
adapter.execute_merge_payload(
    {
        "selection": payload["selection"],
        "decisions_workbook_path": "host_decisiones.xlsx",
        "decisions_sheet_name": "DecisionTable",
        "output_path": "resultado_combinado.xlsx",
        "include_sheets_from_source_only": True,
    }
)
```

---

## Modos de comparación

### Usa `coordinate` cuando...
- importa la posición exacta de la celda,
- quieres aplicar merge automático sobre coordenadas concretas,
- estás comparando plantillas, reportes o formatos relativamente estables.

### Usa `row-based` cuando...
- comparas tablas con encabezados,
- puede haber altas, bajas o inserciones intermedias,
- quieres evitar cascadas por desplazamiento de filas,
- puedes definir columnas clave por hoja como `Clientes:ID`.

> Nota: `row-based` sigue siendo especialmente útil para revisión y auditoría. El merge automático final sigue siendo la ruta más directa en `coordinate`.

---

## Interfaz secundaria: Streamlit (legacy/demo)

La UI web se mantiene solo como interfaz opcional de apoyo:

```bash
streamlit run app.py
```

Úsala para demos internas, validación rápida o troubleshooting. La dirección principal del producto es Excel Desktop.

- **Entrada:** un `WorkbookDiff`, una ruta de salida y una acción por defecto válida.
- **Salida:** archivo Excel con hoja `Decisiones` y hoja `Resumen`.
- **Contrato de decisiones:** la hoja `Decisiones` usa columnas estables como `sheet`, `row`, `column`, `header`, `key`, `diff_type`, `action`, `manual_value` y `reviewed`.
- **Metadata complementaria:** la hoja `Resumen` persiste `compare_mode`, `header_row` y `sheet_keys` para que el merge pueda reconstruir el contexto lógico del diff.

## CLI auxiliar (opcional)

- **Entrada:** ruta a una plantilla editada.
- **Salida:** `pandas.DataFrame` normalizado, con columnas estándar del motor.
- **Acciones válidas:** `use_a`, `use_b`, `manual`.
- **Metadata preservada:** cuando existe hoja `Resumen`, el DataFrame resultante conserva `compare_mode`, `header_row` y `sheet_keys` en `DataFrame.attrs`.

#### `apply_decisions(workbook_a, decisions, output_path, workbook_b, base, compare_mode, header_row, sheet_keys)`

- **Entrada:** dos rutas de libros, un DataFrame de decisiones y la base de merge (`a` o `b`).
- **Salida:** libro combinado en `output_path`.
- **Semántica de acciones:**
  - `use_a`: conserva el valor de A.
  - `use_b`: conserva el valor de B.
  - `manual`: escribe `manual_value`.
- **Contrato `row-based`:** cuando `compare_mode="row-based"`, el merge resuelve decisiones por identidad lógica de registro usando `sheet`, `key`, `header` y `diff_type`, y se apoya en `header_row` / `sheet_keys` para localizar filas aunque hayan cambiado de posición.

## Ejecutar pruebas

```bash
pytest -q
```

---

## Notas

- Soporta `.xlsx` y `.xlsm`.
- Compara valores de celda (no formato, estilos, comentarios, validaciones avanzadas).
- En modo `row-based`, usa encabezados para comparar estructura y detectar registros agregados/eliminados/modificados, y puede aplicar merge sobre esos registros desde la web o desde la plantilla Excel.
- Si quieres auditoría, puedes añadir columnas como usuario/fecha/comentario en la plantilla y extender `apply_decisions`.
