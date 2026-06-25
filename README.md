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

El archivo **`equipos_filtrable.html`** es una interfaz autónoma (no necesita
servidor ni internet) para explorar y anotar los equipos:

- **Filtro por columna estilo Excel**: clic en el **▾** del encabezado para
  marcar **uno o más valores** (con buscador y orden A→Z / Z→A). Los filtros se
  combinan y cada lista se ajusta a lo ya filtrado.
- **Chips de filtros activos** removibles, con resumen y *Limpiar todo*.
- **Tablero accionable**: los contadores *con pendientes*, *vencidos* y *vida
  útil vencida* filtran la tabla al hacer clic.
- **Señal de vida útil**: la columna *Vida Útil Residual* se colorea (rojo si
  está agotada, ámbar si quedan ≤2 años, verde si vigente).
- **Vista despejada por defecto**: parte mostrando las columnas más útiles; el
  resto se activa en **▦ Columnas** (la preferencia se recuerda).
- **Columnas ID y Equipo congeladas** al hacer scroll horizontal.
- **Ficha de detalle**: clic en una fila abre un panel lateral con todos los
  datos del equipo, texto libre y el registro de intervenciones.
- **Intervenciones, pendientes y recordatorios** (ver abajo): columnas
  **Estado** (Vencido / Pendiente / Al día) y **Vence**, con resaltado de
  vencidos (rojo) y próximos (ámbar).
- **Densidad** cómoda/compacta, **búsqueda global**, **Exportar CSV** de la
  vista actual y estado vacío cuando no hay coincidencias.
- **Panel Resumen** (📊): equipos y pendientes por servicio, distribución de
  vida útil y de estado de mantención; cada barra/segmento filtra al hacer clic.
- **Vista de tarjetas** (▤) además de la tabla, para escanear los equipos.
- **Modo oscuro** (🌙) con interruptor, recordado entre sesiones.
- **Vista Programa de mantención** (📅, ver abajo): una fila por equipo y mes
  programado, con la fecha de ejecución editable.
- **Atajos de teclado**: `/` buscar · `t` tabla/tarjetas · `d` resumen ·
  `p` programa · `o` claro/oscuro · `?` ayuda · `Esc` cerrar.
- Los ceros a la izquierda de `Serie` y `N° Inventario` se conservan.

Para regenerarlo a partir de la base:

```bash
python scripts/generar_html.py     # equipos.db -> equipos_filtrable.html
```

### Cuatro formas de anotar (y en qué se diferencian)

| Concepto | Qué responde | Ejemplo |
|----------|--------------|---------|
| **Observación** | *¿Cómo está el equipo? ¿Qué se constató?* — hecho **técnico / de estado** (viene de la columna *Observación* de la planilla). | "Batería al 40%", "Equipo operativo" |
| **Nota** | *¿Qué quiero comentar/recordar yo?* — anotación **libre o administrativa**. | "Equipo prioritario", "En arriendo Mediplex" |
| **Pendiente** | Una **tarea por hacer**: fecha + descripción + estado pendiente/listo. | "25-06: cambiar palas — pendiente" |
| **Recordatorio** | Un pendiente con **fecha de vencimiento** que la app resalta cuando se acerca o ya pasó. | "Vence 01-08: renovar garantía" |

**Intervenciones / pendientes / recordatorios** se gestionan en la ficha de
detalle (clic en una fila): agrega cada intervención con su **fecha**,
**descripción**, marca de **pendiente** y, opcionalmente, una **fecha de
vencimiento** (recordatorio). La tabla deriva de ahí:

- **Estado**: `Vencido` (hay un pendiente con vencimiento ya pasado),
  `Pendiente`, `Al día` (tiene intervenciones, ninguna pendiente) o vacío.
- **Vence**: el próximo vencimiento pendiente; se pinta rojo si está vencido y
  ámbar si vence dentro de 30 días. La fila también se marca con un borde de
  color.
- Filtra al instante con *⏳ Pendientes* / *⚠ Vencidos* y revisa los contadores
  de la barra de estadísticas.

### Programa de mantención (vista)

El botón **📅 Programa MP** (o la tecla `p`) abre una vista con el programa de
mantención del año, leído de la hoja `Registro_MP-2026` del Excel. Cada fila es
una mantención (equipo × mes programado) con estas columnas:

**ID · N° Carpeta · N° Inventario · Equipo · Servicio · Ubicación · Marca ·
Modelo · Serie · Mes · Programa · Resultado · Fecha de ejecución · Última
actualización**.

Los códigos de **Programa** y **Resultado** se muestran decodificados (con el
código original visible al pasar el cursor):

| Programa | | Resultado | |
|----|----|----|----|
| `X`  | Mantención Preventiva Programada | `Si`      | Realizada |
| `R`  | Reprogramada                     | `C1`–`C8` | Reprogramada (ver causales) |
| `RA` | Reprogramada de Año Anterior     | `Si-RA`   | Realizada de Año Anterior |
| `PM` | Puesta en Marcha                 | `FS`      | Fuera de Servicio |
|      |                                  | `No`      | No Realizada |
|      |                                  | `NU`      | No Ubicable |
|      |                                  | `Baja`    | Equipo Dado de Baja |
|      |                                  | *(vacío)* | Pendiente |

La **Fecha de ejecución** es **editable** (se completa al ejecutar la mantención)
y la **Última actualización** se pone **sola** cada vez que editas la fecha.

**Filtros estilo Excel**: cada encabezado tiene un botón **▾** que abre un menú
con buscador y casillas para elegir **uno o varios valores** de esa columna
(con "(Seleccionar todo)"). Los filtros de distintas columnas se **combinan**, el
encabezado filtrado se resalta y **✕ Limpiar filtros** los quita todos. La
búsqueda global también aplica. La fecha de ejecución se guarda igual que las
notas (localStorage + respaldo `notas_equipos.json`).

**En la base de datos**: al importar se crea la tabla `mantenciones` (una fila
por equipo y mes) y la vista SQL **`vista_programa`** con esas columnas ya
legibles. La interfaz lee el programa desde esa tabla (fuente única). Para
volcar las fechas de ejecución que editaste en el HTML a la base:

```bash
python scripts/aplicar_notas.py notas_equipos.json   # escribe en la tabla mantenciones
python scripts/generar_html.py                       # regenera el HTML con las fechas
```

Consultas SQL de ejemplo:

```sql
SELECT * FROM vista_programa WHERE "Resultado" = 'Pendiente' AND "Mes" = 'Agosto';
SELECT "Resultado", COUNT(*) FROM vista_programa GROUP BY "Resultado";
-- También se puede consultar la tabla base directamente:
SELECT mes, COUNT(*) FROM mantenciones WHERE resultado = 'Si' GROUP BY mes ORDER BY mes;
```

### Escribir notas y observaciones

Las columnas **Observaciones** (precargadas con la columna *Observación* de la
planilla) y **Notas** (campo libre) son **editables**: haz clic en la celda (o
usa la ficha de detalle) y escribe. Cómo se guardan:

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
