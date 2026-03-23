# Comparador de Excels

Herramienta en Python orientada a **comparar, revisar y fusionar libros Excel** sin sacar al usuario de su flujo habitual de trabajo. El objetivo del producto es que la experiencia principal ocurra en Excel —o en una integración/add-in conectada a Excel— mientras el motor reutiliza el núcleo de comparación y merge ya existente.

## Producto objetivo

La visión del proyecto es una solución **Excel-first**:

- comparar dos libros `.xlsx` o `.xlsm`,
- revisar diferencias con decisiones explícitas,
- aplicar un merge controlado,
- guardar un libro final listo para continuar trabajando en Excel.

La prioridad ya no es presentar el proyecto como una app web generalista, sino como un **motor de comparación de Excel con integración operativa en Excel**.

## Estado del producto y transición

### Qué se mantiene del motor actual

Se conserva como base estable el motor Python ya implementado:

- comparación multi-hoja entre dos libros,
- detección de hojas exclusivas de A y de B,
- reglas de normalización de valores (`case_sensitive`, espacios, `""` vs `None`),
- modos de comparación `coordinate` y `row-based`,
- exportación e importación de plantillas de decisiones,
- merge final controlado por acciones `use_a`, `use_b` y `manual`.

El punto de continuidad es `comparator.py`, que sigue siendo el contrato principal para cualquier interfaz futura.

### Qué interfaz queda soportada

La interfaz que se considera **objetivo y soportada** es la integración centrada en Excel:

- flujo con plantilla de decisiones dentro de un libro Excel,
- automatización o add-in que invoque el núcleo Python,
- decisiones revisadas desde Excel antes de generar el merge final.

En términos prácticos, hoy esto se materializa mediante el flujo de `excel_tool.py`, pensado como puente hacia una experiencia Excel-first más integrada.

### Qué interfaz queda en desuso

La interfaz web con Streamlit pasa a considerarse **legado / compatibilidad**:

- sigue disponible para pruebas internas o soporte operativo,
- no define la arquitectura objetivo del producto,
- no debe tomarse como la experiencia principal a futuro.

## Capacidades del motor actual

- Compara todas las hojas comunes entre dos libros.
- Detecta hojas exclusivas en A y en B.
- Permite elegir la dirección del merge final (construir sobre A o sobre B).
- Exporta una plantilla Excel editable con decisiones por diferencia.
- Aplica decisiones manuales o automáticas para producir un archivo combinado.
- Puede copiar hojas que solo existen en el libro origen elegido.

## Modos de comparación

### `coordinate`

Úsalo cuando importa la **posición exacta de la celda**:

- plantillas,
- reportes,
- hojas donde `B12` debe seguir siendo `B12`,
- escenarios en los que el merge final debe aplicarse sobre coordenadas concretas.

### `row-based`

Úsalo cuando el contenido representa **registros tabulares**:

- altas y bajas de filas,
- inserciones intermedias,
- comparación por encabezados,
- identificación por columnas clave como `ID`, `Código`, `Email`, etc.

Este modo ayuda a revisar cambios de estructura y registros sin el efecto cascada típico del diff por coordenadas.

## Flujo real del usuario en Excel

La experiencia objetivo del usuario dentro de Excel es la siguiente:

### 1. Abrir la integración

El usuario abre Excel y lanza la integración disponible para su entorno:

- una plantilla conectada,
- un script/launcher corporativo,
- o un futuro add-in que invoque el motor Python.

Desde ahí selecciona el **libro A** y el **libro B**, además del modo de comparación y, si aplica, claves por hoja.

### 2. Comparar

La integración ejecuta el núcleo de comparación y genera una vista o plantilla de decisiones:

- detecta hojas comunes y exclusivas,
- identifica diferencias,
- prepara acciones por defecto (`use_a`, `use_b` o `manual`),
- deja lista la revisión dentro del libro o del add-in.

### 3. Revisar decisiones

El usuario revisa cada diferencia desde Excel:

- acepta el valor de A,
- acepta el valor de B,
- o escribe un valor manual.

En el flujo actual, esto ocurre editando la hoja `Decisiones` de un archivo Excel generado por la herramienta.

### 4. Aplicar merge

Una vez revisadas las decisiones, la integración ejecuta el merge sobre el libro base elegido:

- construir resultado sobre A para traer cambios desde B,
- o construir resultado sobre B para traer cambios desde A.

El merge reutiliza las decisiones registradas y genera un único libro resultante.

### 5. Guardar resultado

El usuario guarda el archivo combinado y continúa trabajando en Excel con ese libro como nuevo resultado consolidado.

## Arquitectura objetivo

La arquitectura objetivo se organiza en tres piezas.

### 1. Núcleo Python: `comparator.py`

`comparator.py` es el **núcleo de negocio** y debe seguir concentrando:

- contratos de comparación (`compare_workbooks` / `ComparatorService.compare`),
- estructura estándar de diferencias,
- exportación de plantilla de decisiones,
- lectura de decisiones editadas,
- aplicación del merge final.

La lógica de negocio debe vivir aquí para que no se replique en cada interfaz.

### 2. Adaptador Excel

Sobre el núcleo se sitúa un **adaptador de integración con Excel**. Hoy ese papel lo cumplen la CLI `excel_tool.py` y la capa `interface_adapter.py`; a futuro puede ser un add-in, un launcher o una automatización corporativa.

Su responsabilidad es:

- capturar archivos y parámetros,
- traducirlos al contrato del núcleo,
- lanzar la comparación,
- transportar decisiones desde/hacia Excel,
- ejecutar el merge final.

### 3. Flujo de decisiones dentro del libro o add-in

La capa visible para el usuario debe residir en Excel:

- hoja de decisiones dentro del libro,
- panel lateral de add-in,
- acciones guiadas para revisar y confirmar cambios.

La experiencia objetivo no es “subir dos Excels a una web”, sino **resolver diferencias desde el propio contexto de Excel**.

## Interfaz soportada hoy: flujo Excel-first

### Crear plantilla de decisiones

#### Comparación por coordenadas

```bash
python excel_tool.py compare --a libro_a.xlsx --b libro_b.xlsx --base a --template decisiones.xlsx
```

#### Comparación por filas con claves por hoja

```bash
python excel_tool.py compare \
  --a libro_a.xlsx \
  --b libro_b.xlsx \
  --compare-mode row-based \
  --sheet-key Clientes=ID \
  --sheet-key Pedidos=Empresa,NumeroPedido \
  --header-row 1 \
  --template decisiones.xlsx
```

Esto genera un Excel con:

- hoja `Decisiones`,
- columnas como `header`, `diff_type` y `key`,
- columna `action` con `use_a`, `use_b` y `manual`,
- columna `manual_value` para intervención manual.

### Revisar decisiones en Excel

```bash
# abrir decisiones.xlsx manualmente en Excel y editar la hoja Decisiones
```

### Generar el libro combinado

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

## Streamlit como compatibilidad / legado

La interfaz web sigue disponible como herramienta secundaria para soporte interno o validación funcional.

```bash
streamlit run app.py
```

Úsala solo cuando necesites:

- revisar diferencias desde navegador en lugar de Excel,
- depurar el comportamiento del motor,
- validar rápidamente opciones de comparación.

No representa la dirección objetivo del producto.

## Limitaciones actuales

### 1. Estilos y formato no se comparan

La herramienta compara **valores** y decisiones de merge, pero no resuelve de forma completa:

- estilos,
- formatos,
- comentarios,
- validaciones avanzadas,
- otros metadatos visuales del libro.

### 2. Diferencias entre `coordinate` y `row-based`

Los dos modos no significan lo mismo ni ofrecen la misma capacidad operativa:

| Aspecto | `coordinate` | `row-based` |
|---|---|---|
| Unidad comparada | Celda por posición | Registro por fila |
| Supuesto principal | Importa la coordenada exacta | Importa la identidad lógica del registro |
| Requiere encabezados | No | Sí, o al menos una estructura tabular interpretable |
| Claves por hoja | No | Recomendadas para emparejar registros |
| Altas/bajas de filas | Tienden a generar cascadas de diferencias | Se modelan mejor como added/deleted/modified |
| Merge final automático | Es el flujo recomendado | Actualmente queda orientado sobre todo a revisión/auditoría |

### 3. Dependencia de entorno local si la integración usa Python

Mientras la integración dependa de ejecutar Python localmente, el entorno del usuario o del equipo necesita:

- Python instalado,
- dependencias del proyecto disponibles,
- acceso a los archivos Excel desde ese entorno,
- mecanismos corporativos para distribuir/actualizar el runtime si se quiere una experiencia masiva.

Esto es relevante para cualquier integración en Excel basada en scripts, launcher local o add-in que delegue en un backend Python local.

## Tabla de migración para usuarios actuales

| Si hoy usas... | Estado | Qué cambia | Qué debes hacer ahora |
|---|---|---|---|
| `app.py` | Legado / compatibilidad | La web deja de ser la experiencia principal; pasa a ser un visor/editor auxiliar | Mantenerlo solo para soporte, pruebas o validación interna |
| `excel_tool.py compare` | Soportado | Se consolida como puente hacia la integración Excel-first | Seguir generando la plantilla de decisiones desde aquí |
| `excel_tool.py merge` | Soportado | Sigue siendo el mecanismo operativo para materializar el merge final | Seguir usándolo tras revisar decisiones en Excel |
| Integración futura con add-in | Objetivo | Debe reutilizar `comparator.py` y el adaptador, no duplicar lógica de negocio | Diseñar la UX dentro de Excel y conectar con el núcleo existente |

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Ejecutar pruebas

```bash
pytest -q
```

## Componentes principales

- `comparator.py`: núcleo estable del motor y contrato principal de comparación/merge.
- `interface_adapter.py`: capa adaptadora entre el núcleo y las interfaces.
- `excel_tool.py`: interfaz operativa actual para el flujo Excel-first.
- `app.py`: interfaz Streamlit en modo legado/compatibilidad.
- `tests/test_comparator.py`: pruebas unitarias del núcleo.

## Notas finales

- Soporta `.xlsx` y `.xlsm`.
- El núcleo está preparado para ser reutilizado por un add-in, automatización o integración corporativa.
- La dirección del producto es **Excel como experiencia principal** y **Python como motor reutilizable**.
