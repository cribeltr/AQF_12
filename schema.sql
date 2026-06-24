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
    notas               TEXT                   -- Notas (campo libre del usuario; editables)
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
    notas              AS "Notas"
FROM equipos
ORDER BY id;
