# AQF_12 — Base de datos de equipos médicos (Programa de Mantenciones Preventivas 2026)

Este repositorio importa la planilla **`Programacion_MP_2026.xlsm`** y crea una
base de datos **SQLite** (`equipos.db`) con el inventario de equipos médicos.

## Columnas de la base de datos

La tabla `equipos` contiene exactamente estas columnas (origen: hoja
`PMP_2026` de la planilla):

| Columna en la planilla | Columna en SQLite    | Tipo    | Notas |
|------------------------|----------------------|---------|-------|
| ID                     | `id`                 | INTEGER | Orden dentro de la planilla (clave primaria) |
| N° Carpeta             | `n_carpeta`          | INTEGER | |
| N° Inventario          | `n_inventario`       | TEXT    | **Se conserva tal cual** (preserva ceros) |
| Equipo                 | `equipo`             | TEXT    | |
| Servicio               | `servicio`           | TEXT    | |
| Unidad                 | `unidad`             | TEXT    | |
| Ubicación              | `ubicacion`          | TEXT    | |
| Procedencia            | `procedencia`        | TEXT    | |
| Marca                  | `marca`              | TEXT    | |
| Modelo                 | `modelo`             | TEXT    | |
| Serie                  | `serie`              | TEXT    | **Se conserva tal cual** (preserva ceros) |
| Año Instalación        | `anio_instalacion`   | INTEGER | |
| Vida Útil Residual     | `vida_util_residual` | TEXT    | Número o `"Disponible"` (es una fórmula en Excel) |
| Clasificación          | `clasificacion`      | TEXT    | |
| ENU / Baja             | `enu_baja`           | TEXT    | |

> También existe la vista **`vista_equipos`**, que expone los mismos datos con
> los **encabezados originales** (`"N° Carpeta"`, `"Año Instalación"`, etc.) por
> si prefieres consultar con esos nombres.

## Reglas de identidad

- **`id`** corresponde al **orden dentro de la planilla Excel**; es solo un
  número de fila, no un identificador de negocio estable.
- Cada equipo se **identifica de forma única por su `serie` o su
  `n_inventario`**. Por eso ambos campos tienen un índice `UNIQUE` (parcial:
  solo sobre los valores presentes, de modo que las filas sin identificador no
  chocan entre sí).

## Preservación de ceros a la izquierda

Los números de serie e inventario se guardan como **TEXTO** y la importación
**nunca los convierte a número**. Así, `"0024"` se mantiene como `"0024"` y
jamás se transforma en `"24"`. Lo mismo ocurre con valores como `"02516."`,
`"00888"`, `"2-013518"`, etc.

## Cómo regenerar la base de datos

Requisitos: Python 3.9+.

```bash
# 1. Instalar dependencias
pip install -r requirements.txt

# 2. Importar la planilla y crear equipos.db
python scripts/importar_excel.py
```

Parámetros opcionales:

```bash
python scripts/importar_excel.py \
    --excel data/Programacion_MP_2026.xlsm \
    --db    equipos.db \
    --hoja  PMP_2026
```

El script vuelve a crear la tabla desde cero en cada ejecución (es
idempotente) e informa cuántos equipos se importaron y si hubo conflictos de
unicidad.

## Consultas de ejemplo

```sql
-- Buscar un equipo por número de serie (conservando ceros)
SELECT * FROM equipos WHERE serie = '00888';

-- Buscar por número de inventario
SELECT * FROM equipos WHERE n_inventario = '2-013518';

-- Equipos por servicio
SELECT servicio, COUNT(*) AS total
FROM equipos
GROUP BY servicio
ORDER BY total DESC;

-- Ver los datos con los encabezados originales
SELECT * FROM vista_equipos LIMIT 20;
```

Desde la terminal:

```bash
sqlite3 equipos.db "SELECT id, equipo, marca, serie FROM equipos WHERE serie='00888';"
```

## Tabla con filtros tipo Excel (HTML)

El archivo **`equipos_filtrable.html`** es una tabla autónoma (no necesita
servidor ni internet) en la que **cada columna es filtrable al estilo Excel**:

- Haz clic en el **▾** del encabezado para abrir un desplegable con casillas y
  marcar **uno o más valores**.
- El desplegable se ajusta a los filtros de las otras columnas (como Excel) y
  trae búsqueda de valores y orden A→Z / Z→A.
- Búsqueda global, indicador de columnas filtradas, contador de resultados y
  **Exportar CSV** de la vista actual.
- Los ceros a la izquierda de `Serie` y `N° Inventario` se conservan.

Para regenerarlo a partir de la base:

```bash
python scripts/generar_html.py     # equipos.db -> equipos_filtrable.html
```

### Escribir notas y observaciones

Las columnas **Observaciones** (precargadas con la columna *Observación* de la
planilla) y **Notas** (campo libre) son **editables**: haz clic en la celda y
escribe. Cómo se guardan:

1. **Automático en el navegador** (localStorage): se conservan al recargar.
2. **💾 Guardar notas**: descarga `notas_equipos.json`, un respaldo portable
   (para llevarlas a otro computador o devolverlas a la base de datos).
3. **📂 Cargar notas**: importa un `notas_equipos.json` previo.

> El guardado en el navegador depende del equipo/navegador. Para no perder
> trabajo, usa **Guardar notas** periódicamente. Si hay cambios sin respaldar,
> aparece el aviso *“● notas sin respaldar en archivo”*.

Para incorporar las notas a la base de datos (que sea la fuente de verdad):

```bash
python scripts/aplicar_notas.py notas_equipos.json   # vuelca el JSON a equipos.db
python scripts/generar_html.py                       # regenera el HTML ya con las notas
```

Cada equipo se identifica por `Serie` → `N° Inventario` → `#ID` (el primero no
vacío), tanto en el HTML como al volcar a la base.

## Estructura del repositorio

```
.
├── data/
│   └── Programacion_MP_2026.xlsm   # Planilla de origen
├── scripts/
│   ├── importar_excel.py           # Importador Excel -> SQLite
│   ├── generar_html.py             # Genera la tabla HTML con filtros tipo Excel
│   └── aplicar_notas.py            # Vuelca notas_equipos.json a la base de datos
├── schema.sql                      # Definición de la tabla, índices y vista
├── equipos.db                      # Base de datos generada (966 equipos)
├── equipos_filtrable.html          # Tabla interactiva: filtros + notas editables
├── requirements.txt                # Dependencias (openpyxl)
└── README.md
```
