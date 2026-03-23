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

### Cómo funciona `row-based`

- Toma la fila de encabezados (por defecto la fila 1).
- Empareja filas por las columnas clave configuradas para esa hoja.
- Si una hoja no tiene clave configurada, usa el contenido completo de la fila como identidad implícita.
- Reporta diferencias con tipo `added`, `deleted` o `modified`.

> **Nota:** la comparación `row-based` está orientada a auditoría y revisión de registros. La combinación automática del libro final sigue siendo la opción recomendada en modo `coordinate`.

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

En la interfaz web (`streamlit run app.py`) ahora también puedes elegir visualmente la dirección del merge antes de generar el archivo final.

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
- En modo `row-based`, usa encabezados para comparar estructura y detectar registros agregados/eliminados/modificados.
- Si quieres auditoría, puedes añadir columnas como usuario/fecha/comentario en la plantilla y extender `apply_decisions`.
