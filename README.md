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

---

## Contratos serializables para Excel

El add-in debe trabajar con contratos estables en vez de consumir `WorkbookDiff`.

### Selección de libros

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

---

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

```python
from excel_addin_adapter import ExcelAddinAdapter

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

---

## CLI auxiliar (opcional)

Sigue existiendo un flujo por línea de comandos para soporte y automatización:

```bash
python excel_tool.py compare --a libro_a.xlsx --b libro_b.xlsx --base a --template decisiones.xlsx
python excel_tool.py merge --a libro_a.xlsx --b libro_b.xlsx --decisions decisiones.xlsx --apply-onto a --output resultado.xlsx
```

---

## Ejecutar pruebas

```bash
pytest -q
```

---

## Notas

- Soporta `.xlsx` y `.xlsm`.
- Compara valores de celda; no busca preservar estilos complejos, comentarios ni reglas avanzadas de Excel.
- La hoja de decisiones usa un contrato estable con columnas como `sheet`, `row`, `column`, `action`, `manual_value` y `reviewed`.
- Si en el futuro se necesitara telemetría, auditoría o integración corporativa adicional, debería añadirse en nuevos adaptadores o capas de orquestación, no en el motor base.
