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
- Permite reglas de comparación:
  - ignorar mayúsculas/minúsculas,
  - recortar o no espacios,
  - tratar `""` y `None` como iguales o distintos.
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

## Interfaz 2: Flujo Excel (CLI + plantilla editable)

### 1) Crear plantilla de decisiones

```bash
python excel_tool.py compare --a libro_a.xlsx --b libro_b.xlsx --template decisiones.xlsx
```

Esto genera `decisiones.xlsx` con:
- Hoja `Decisiones`: una fila por diferencia.
- Columna `action` con lista desplegable (`use_a`, `use_b`, `manual`).
- Columna `manual_value` para casos manuales.

### 2) Editar decisiones en Excel

Abre `decisiones.xlsx` en Excel y cambia acciones.

### 3) Generar libro combinado

```bash
python excel_tool.py merge --a libro_a.xlsx --b libro_b.xlsx --decisions decisiones.xlsx --output resultado.xlsx
```

---

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
- Compara valores de celda (no formato, estilos, comentarios, validaciones avanzadas).
- Si quieres auditoría, puedes añadir columnas como usuario/fecha/comentario en la plantilla y extender `apply_decisions`.
