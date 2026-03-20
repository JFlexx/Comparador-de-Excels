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
- Permite comparar por tipo de cambio:
  - valores literales,
  - fórmulas,
  - resultados cacheados de fórmulas,
  - estilos,
  - comentarios.
- Etiqueta cada diferencia con tipos como `value_changed`, `formula_changed`, `cached_result_changed`, `style_changed` o `comment_changed`.
- Genera una plantilla Excel con columnas adicionales para entender qué cambió (`difference_types`, fórmulas y resultados cacheados de cada lado).
- Genera resultado combinado.
- Copia opcional de hojas que existen solo en B.

---

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Interfaz 1: Web (Streamlit)

```bash
streamlit run app.py
```

Luego abre la URL mostrada por Streamlit (normalmente `http://localhost:8501`).

En la barra lateral puedes elegir si quieres comparar valores, fórmulas, resultados cacheados, estilos o comentarios. Puedes combinar varios criterios a la vez.

## Interfaz 2: Flujo Excel (CLI + plantilla editable)

### 1) Crear plantilla de decisiones

```bash
python excel_tool.py compare --a libro_a.xlsx --b libro_b.xlsx --template decisiones.xlsx
```

Esto genera `decisiones.xlsx` con:
- Hoja `Decisiones`: una fila por diferencia.
- Columna `difference_types` para explicar el cambio detectado.
- Columnas `formula_a`, `formula_b`, `cached_value_a`, `cached_value_b` para auditoría adicional.
- Columna `action` con lista desplegable (`use_a`, `use_b`, `manual`).
- Columna `manual_value` para casos manuales.

### Modos de comparación desde CLI

Por defecto el CLI compara solo valores literales. Para cambiarlo, repite `--compare-mode`:

```bash
python excel_tool.py compare \
  --a libro_a.xlsx \
  --b libro_b.xlsx \
  --template decisiones.xlsx \
  --compare-mode formulas \
  --compare-mode styles
```

Modos disponibles:
- `values`
- `formulas`
- `cached_results`
- `styles`
- `comments`

### 2) Editar decisiones en Excel

Abre `decisiones.xlsx` en Excel y cambia acciones.

### 3) Generar libro combinado

```bash
python excel_tool.py merge --a libro_a.xlsx --b libro_b.xlsx --decisions decisiones.xlsx --output resultado.xlsx
```

---

## Fórmulas vs. resultados calculados (`data_only`)

La comparación usa dos lecturas distintas de `openpyxl` según lo que hayas elegido:

- **Comparar fórmulas**: se carga el libro con `load_workbook(..., data_only=False)` para leer el texto real de la fórmula, por ejemplo `=SUM(A1:A5)`.
- **Comparar resultados cacheados**: se carga además con `load_workbook(..., data_only=True)` para leer el último valor calculado y guardado dentro del archivo.

### Limitaciones importantes

- `openpyxl` **no recalcula fórmulas**.
- `data_only=True` **no ejecuta Excel** ni actualiza fórmulas; solo lee el valor cacheado ya guardado en el archivo.
- Si el archivo nunca fue recalculado/guardado por Excel, LibreOffice u otra herramienta que escriba ese caché, el resultado puede aparecer como `None` o vacío.
- Si comparas solo `values`, dos fórmulas distintas no aparecerán como cambio salvo que una de las celdas deje de ser fórmula y pase a un valor literal.

Recomendación:
- usa `formulas` cuando quieras detectar cambios en la lógica;
- usa `cached_results` cuando quieras revisar el último resultado visible guardado;
- combina ambos si quieres saber si cambió tanto la lógica como el resultado persistido.

## Ejecutar pruebas

```bash
pytest -q
```

## Arquitectura

- `comparator.py`: lógica de comparación, export/import de decisiones y aplicación de merge.
- `app.py`: interfaz web en Streamlit.
- `excel_tool.py`: CLI para flujo Excel-first.
- `tests/test_comparator.py`: pruebas unitarias.

## Notas

- Soporta `.xlsx` y `.xlsm`.
- La fusión final copia valores elegidos desde B hacia A; actualmente no replica estilos/comentarios durante `apply_decisions`.
- Si quieres auditoría adicional, puedes extender la plantilla con columnas como usuario/fecha/comentario y ajustar `apply_decisions`.
