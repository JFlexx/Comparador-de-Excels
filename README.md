# Comparador de Excels (libro completo)

Herramienta interna en Python para comparar dos archivos Excel a nivel de **libro completo** (todas las hojas), resolver diferencias y generar un archivo combinado.

Incluye **dos interfaces**:
- **Web (Streamlit)** para uso guiado.
- **Excel-first** para equipos que prefieren editar decisiones directamente dentro de Excel.

## ¿Qué problema resuelve?

- Comparación multi-hoja (no solo una sheet).
- Resolución de diferencias por celda (`use_a`, `use_b`, `manual`).
- Flujo de fusión para "traer" cambios de un libro a otro.
- Alternativa interna sin depender de productos de terceros para editar/combinar.

## Características

- Compara todas las hojas comunes entre dos libros.
- Detecta hojas exclusivas de A y de B.
- Soporta dos modos de diff:
  - **`coordinate`**: compara celda contra celda por posición.
  - **`row-based`**: compara registros por filas usando encabezados y columnas clave opcionales por hoja.
- Permite reglas de comparación:
  - ignorar mayúsculas/minúsculas,
  - recortar o no espacios,
  - tratar `""` y `None` como iguales o distintos.
- Genera resultado combinado en cualquiera de las dos direcciones (A→B o B→A).
- Copia opcional de hojas que existen solo en el libro origen elegido.

---

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## ¿Cuándo usar cada modo?

### Usa `coordinate` cuando...

- Quieres comparar plantillas o reportes donde **la posición exacta de la celda importa**.
- Una diferencia en `B12` debe seguir mostrándose como cambio en `B12`.
- Vas a usar el flujo de combinación automática para aplicar decisiones al libro final.

### Usa `row-based` cuando...

- Estás comparando listados tabulares con encabezados (`ID`, `Código`, `Email`, etc.).
- Puede haber **altas, bajas o inserciones intermedias** de filas.
- Quieres evitar el efecto cascada típico de los diffs por coordenadas cuando una fila nueva desplaza todas las siguientes.
- Puedes definir una o varias columnas clave por hoja, por ejemplo `Clientes:ID` o `Pedidos:Empresa,NúmeroPedido`.
- También quieres **fusionar** esos cambios al libro final sin depender de la coordenada original de Excel.

### Cómo funciona `row-based`

- Toma la fila de encabezados (por defecto la fila 1).
- Empareja filas por las columnas clave configuradas para esa hoja.
- Si una hoja no tiene clave configurada, usa el contenido completo de la fila como identidad implícita.
- Reporta diferencias con tipo `added`, `deleted` o `modified`.
- La plantilla de decisiones y el merge consumen `sheet`, `key`, `header` y `diff_type` para reubicar el registro al aplicar cambios.

> **Nota:** para que el merge `row-based` sea estable y predecible, conviene definir `sheet_keys` por hoja. Si no se define clave, el motor puede caer en coordenadas/fila de apoyo cuando no pueda reconstruir una identidad lógica suficiente.

## Interfaz 1: Web (Streamlit)

```bash
streamlit run app.py
```

Luego abre la URL mostrada por Streamlit (normalmente `http://localhost:8501`).

En la barra lateral puedes elegir:
- modo de comparación,
- fila de encabezados,
- claves por hoja con formato `Hoja:columna1,columna2`.

## Interfaz 2: Flujo Excel (CLI + plantilla editable)

### 1) Crear plantilla de decisiones

#### Modo por coordenadas

```bash
python excel_tool.py compare --a libro_a.xlsx --b libro_b.xlsx --base a --template decisiones.xlsx
```

#### Modo por filas con claves por hoja

```bash
python excel_tool.py compare   --a libro_a.xlsx   --b libro_b.xlsx   --compare-mode row-based   --sheet-key Clientes=ID   --sheet-key Pedidos=Empresa,NumeroPedido   --header-row 1   --template decisiones.xlsx
```

Esto genera `decisiones.xlsx` con:
- Hoja `Decisiones`: una fila por diferencia.
- Columnas auxiliares como `header`, `diff_type` y `key` para identificar el registro afectado.
- Metadata del diff en la hoja `Resumen` (`compare_mode`, `header_row`, `sheet_keys`) para permitir merge posterior también en `row-based`.
- Columna `action` con lista desplegable (`use_a`, `use_b`, `manual`).
- Columna `manual_value` para casos manuales.

### 2) Editar decisiones en Excel

Abre `decisiones.xlsx` en Excel y cambia acciones.

### 3) Generar libro combinado

#### Traer cambios de B hacia A

```bash
python excel_tool.py compare --a libro_a.xlsx --b libro_b.xlsx --base a --template decisiones_b_hacia_a.xlsx
python excel_tool.py merge --a libro_a.xlsx --b libro_b.xlsx --decisions decisiones_b_hacia_a.xlsx --apply-onto a --output resultado_b_hacia_a.xlsx
```

#### Traer cambios de A hacia B

```bash
python excel_tool.py compare --a libro_a.xlsx --b libro_b.xlsx --base b --template decisiones_a_hacia_b.xlsx
python excel_tool.py merge --a libro_a.xlsx --b libro_b.xlsx --decisions decisiones_a_hacia_b.xlsx --apply-onto b --output resultado_a_hacia_b.xlsx
```

En la interfaz web (`streamlit run app.py`) ahora también puedes elegir visualmente la dirección del merge antes de generar el archivo final. Esto aplica tanto a `coordinate` como a `row-based`.

---

## Ejecutar pruebas

```bash
pytest -q
```

## Arquitectura

- `comparator.py`: núcleo estable del motor. Expone la API de servicio (`ComparatorService`) y las funciones públicas `compare_workbooks`, `export_decision_template`, `decisions_from_excel` y `apply_decisions`. También documenta los contratos de entrada/salida para rutas, DataFrames de decisiones, acciones válidas y modos `coordinate` / `row-based`, incluyendo merge lógico por registro.
- `interface_adapter.py`: capa adaptadora para interfaces. Centraliza parsing de opciones, etiquetas de dirección de merge, DataFrames de revisión y llamadas al servicio.
- `app.py`: interfaz web en Streamlit; solo orquesta UI y delega en la capa adaptadora.
- `excel_tool.py`: CLI para flujo Excel-first; solo interpreta argumentos y delega en la capa adaptadora.
- `tests/test_comparator.py`: pruebas unitarias del núcleo y de los contratos principales.

## Arquitectura objetivo

El objetivo es que **cualquier interfaz corporativa futura** consuma el núcleo de `comparator.py` en lugar de reimplementar lógica de negocio. Eso incluye:

- add-ins de Excel,
- APIs internas,
- automatizaciones batch,
- portales web o backoffices.

### Regla de diseño

Las interfaces deben limitarse a:

1. capturar archivos y parámetros del usuario,
2. convertir esos parámetros al contrato del motor,
3. invocar el servicio del comparador,
4. presentar o persistir el resultado.

La comparación, la construcción del DataFrame de decisiones válido, la exportación/importación de plantillas y la ejecución del merge final deben residir en `comparator.py`.

### Contratos del núcleo

#### `compare_workbooks(path_a, path_b, options)`

- **Entrada:** rutas de libros Excel y un `CompareOptions`.
- **Salida:** `WorkbookDiff` con `only_in_a`, `only_in_b`, `common_sheets`, `differences`, `grouped_differences` y `total_differences`.
- **Modo `coordinate`:** compara celda a celda por fila/columna exacta.
- **Modo `row-based`:** compara registros usando `header_row` y `sheet_keys`. Si una hoja no tiene clave configurada, usa la fila completa como identidad implícita.

#### `export_decision_template(diff, output_path, default_action)`

- **Entrada:** un `WorkbookDiff`, una ruta de salida y una acción por defecto válida.
- **Salida:** archivo Excel con hoja `Decisiones` y hoja `Resumen`.
- **Contrato de decisiones:** la hoja `Decisiones` usa columnas estables como `sheet`, `row`, `column`, `header`, `key`, `diff_type`, `action`, `manual_value` y `reviewed`.
- **Metadata complementaria:** la hoja `Resumen` persiste `compare_mode`, `header_row` y `sheet_keys` para que el merge pueda reconstruir el contexto lógico del diff.

#### `decisions_from_excel(path)`

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

### Puntos de integración para un futuro add-in

Un add-in de Excel o una API interna necesitaría, como mínimo, estos puntos de integración:

1. **Cargar libro base y libro fuente** desde disco, memoria o almacenamiento corporativo.
2. **Invocar la comparación** con `compare_workbooks` o `ComparatorService.compare`.
3. **Mostrar/editar decisiones** sobre el DataFrame estándar o una representación equivalente de la hoja `Decisiones`.
4. **Persistir o rehidratar decisiones** usando `export_decision_template` / `decisions_from_excel`.
5. **Solicitar el merge final** con `apply_decisions` o `ComparatorService.apply_decisions`, indicando si el resultado se construye sobre A o sobre B.
6. **Gestionar hojas exclusivas** del libro origen según la política del canal (`include_sheets_from_source_only`).

## Notas

- Soporta `.xlsx` y `.xlsm`.
- Compara valores de celda (no formato, estilos, comentarios, validaciones avanzadas).
- En modo `row-based`, usa encabezados para comparar estructura y detectar registros agregados/eliminados/modificados, y puede aplicar merge sobre esos registros desde la web o desde la plantilla Excel.
- Si quieres auditoría, puedes añadir columnas como usuario/fecha/comentario en la plantilla y extender `apply_decisions`.
