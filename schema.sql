-- =====================================================================
--  Esquema de la base de datos de equipos médicos (Programa de
--  Mantenciones Preventivas 2026).
--
--  Origen: data/Programacion_MP_2026.xlsm  (hoja "PMP_2026").
--
--  REGLAS DE IDENTIDAD
--  -------------------
--  * `id`           -> corresponde al ORDEN dentro de la planilla Excel.
--                      No es un identificador de negocio estable.
--  * `serie` y
--    `n_inventario` -> identifican de forma ÚNICA a cada equipo. Por eso
--                      ambos tienen un índice UNIQUE (parcial: solo sobre
--                      los valores no nulos, de modo que las filas sin
--                      identificador no entran en conflicto entre sí).
--
--  PRESERVACIÓN DE CEROS A LA IZQUIERDA
--  ------------------------------------
--  `serie` y `n_inventario` se declaran como TEXT y la importación nunca
--  los convierte a número. Así "0024" se mantiene como "0024" y nunca se
--  transforma en "24".
-- =====================================================================

DROP VIEW  IF EXISTS vista_equipos;
DROP TABLE IF EXISTS equipos;

CREATE TABLE equipos (
    id                  INTEGER PRIMARY KEY,   -- ID = orden en la planilla Excel
    n_carpeta           INTEGER,               -- N° Carpeta
    n_inventario        TEXT,                  -- N° Inventario  (texto: preserva ceros)
    equipo              TEXT,                  -- Equipo
    servicio            TEXT,                  -- Servicio
    unidad              TEXT,                  -- Unidad
    ubicacion           TEXT,                  -- Ubicación
    procedencia         TEXT,                  -- Procedencia
    marca               TEXT,                  -- Marca
    modelo              TEXT,                  -- Modelo
    serie               TEXT,                  -- Serie  (texto: preserva ceros a la izquierda)
    anio_instalacion    INTEGER,               -- Año Instalación
    vida_util_residual  TEXT,                  -- Vida Útil Residual (número o "Disponible")
    clasificacion       TEXT,                  -- Clasificación
    enu_baja            TEXT,                  -- ENU / Baja
    observaciones       TEXT,                  -- Observaciones (precargadas de la planilla; editables)
    notas               TEXT,                  -- Notas (campo libre del usuario; editables)
    registros           TEXT,                  -- Intervenciones: JSON [{f:fecha, d:descripción, p:pendiente, v:vencimiento}]
    pendiente           INTEGER,               -- 1 si hay alguna intervención pendiente, 0 si no, NULL si no hay registros
    ultima_intervencion TEXT,                  -- Fecha (YYYY-MM-DD) de la intervención más reciente
    proximo_vencimiento TEXT                   -- Fecha (YYYY-MM-DD) del próximo recordatorio/vencimiento pendiente
);

-- Identidad única del equipo (solo sobre valores presentes).
CREATE UNIQUE INDEX ux_equipos_serie
    ON equipos (serie)        WHERE serie IS NOT NULL;
CREATE UNIQUE INDEX ux_equipos_inventario
    ON equipos (n_inventario) WHERE n_inventario IS NOT NULL;

-- Índices de apoyo para búsquedas frecuentes.
CREATE INDEX ix_equipos_equipo   ON equipos (equipo);
CREATE INDEX ix_equipos_servicio ON equipos (servicio);

-- Vista con los nombres de columna EXACTOS de la planilla original,
-- para quien prefiera consultar con los encabezados tal cual.
CREATE VIEW vista_equipos AS
SELECT
    id                 AS "ID",
    n_carpeta          AS "N° Carpeta",
    n_inventario       AS "N° Inventario",
    equipo             AS "Equipo",
    servicio           AS "Servicio",
    unidad             AS "Unidad",
    ubicacion          AS "Ubicación",
    procedencia        AS "Procedencia",
    marca              AS "Marca",
    modelo             AS "Modelo",
    serie              AS "Serie",
    anio_instalacion   AS "Año Instalación",
    vida_util_residual AS "Vida Útil Residual",
    clasificacion      AS "Clasificación",
    enu_baja           AS "ENU / Baja",
    observaciones      AS "Observaciones",
    notas              AS "Notas",
    registros          AS "Registros",
    pendiente          AS "Pendiente",
    ultima_intervencion AS "Última intervención",
    proximo_vencimiento AS "Próximo vencimiento"
FROM equipos
ORDER BY id;

-- =====================================================================
--  Programa de mantención (origen: hoja "Registro_MP-2026").
--  Una fila por equipo y mes con programa o resultado registrado.
--   Programa : 'X'  = Mantención Preventiva Programada
--              'R'  = Mantención Preventiva Reprogramada
--              'RA' = Reprogramada de Año Anterior
--              'PM' = Puesta en Marcha
--   Resultado: 'Si'    = Realizada;          vacío  = Pendiente
--              'C1'-'C8'= Reprogramada (ver causales)
--              'Si-RA'  = Realizada año ant.; 'FS'  = Fuera de Servicio
--              'No'     = No Realizada;       'NU'  = No Ubicable
--              'Baja'   = Equipo Dado de Baja
--   fecha_ejecucion / ultima_actualizacion: los completa el usuario desde la
--     interfaz (vía aplicar_notas.py). No vienen en el Excel.
--  `clave` es la misma identidad del equipo: Serie -> N° Inventario -> "#"+ID.
-- =====================================================================
DROP VIEW  IF EXISTS vista_programa;
DROP TABLE IF EXISTS mantenciones;

CREATE TABLE mantenciones (
    id                   INTEGER PRIMARY KEY,
    clave                TEXT,       -- identidad del equipo (Serie/N° Inventario/#ID)
    id_planilla          TEXT,       -- ID del equipo en la planilla
    n_inventario         TEXT,
    equipo               TEXT,
    servicio             TEXT,
    ubicacion            TEXT,
    marca                TEXT,
    modelo               TEXT,
    serie                TEXT,
    mes                  INTEGER,    -- 1 = Enero … 12 = Diciembre
    programa             TEXT,
    resultado            TEXT,
    fecha_ejecucion      TEXT,       -- editable (la completa el usuario)
    ultima_actualizacion TEXT        -- se actualiza al editar la fecha de ejecución
);

CREATE INDEX ix_mant_clave ON mantenciones (clave);
CREATE UNIQUE INDEX ux_mant_clave_mes ON mantenciones (clave, mes);

-- Vista legible: mes con nombre y programa/resultado decodificados.
CREATE VIEW vista_programa AS
SELECT
    id_planilla  AS "ID",
    n_inventario AS "N° Inventario",
    equipo       AS "Equipo",
    servicio     AS "Servicio",
    ubicacion    AS "Ubicación",
    marca        AS "Marca",
    modelo       AS "Modelo",
    serie        AS "Serie",
    CASE mes WHEN 1 THEN 'Enero' WHEN 2 THEN 'Febrero' WHEN 3 THEN 'Marzo'
             WHEN 4 THEN 'Abril' WHEN 5 THEN 'Mayo' WHEN 6 THEN 'Junio'
             WHEN 7 THEN 'Julio' WHEN 8 THEN 'Agosto' WHEN 9 THEN 'Septiembre'
             WHEN 10 THEN 'Octubre' WHEN 11 THEN 'Noviembre' WHEN 12 THEN 'Diciembre'
             ELSE mes END                                          AS "Mes",
    CASE upper(programa)
         WHEN 'X'  THEN 'Programada'
         WHEN 'R'  THEN 'Reprogramada'
         WHEN 'RA' THEN 'Reprogramada (año anterior)'
         WHEN 'PM' THEN 'Puesta en marcha'
         WHEN 'BAJA' THEN 'Baja'
         ELSE programa END                                        AS "Programa",
    CASE
         WHEN resultado IS NULL OR resultado = '' THEN 'Pendiente'
         WHEN upper(resultado) = 'SI'    THEN 'Realizada'
         WHEN upper(resultado) = 'SI-RA' THEN 'Realizada (año anterior)'
         WHEN upper(resultado) IN ('C1','C2','C3','C4','C5','C6','C7','C8')
              THEN 'Reprogramada (' || resultado || ')'
         WHEN upper(resultado) = 'FS'   THEN 'Fuera de servicio'
         WHEN upper(resultado) = 'NO'   THEN 'No realizada'
         WHEN upper(resultado) = 'NU'   THEN 'No ubicable'
         WHEN upper(resultado) = 'BAJA' THEN 'Baja'
         ELSE resultado END                                       AS "Resultado",
    fecha_ejecucion      AS "Fecha de ejecución",
    ultima_actualizacion AS "Última actualización"
FROM mantenciones
ORDER BY clave, mes;
