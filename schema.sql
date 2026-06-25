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
    n_carpeta            TEXT,
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

-- ===== Causales de reprogramación (referencia fija, provista por SEC) =====
--  plazo_dias = 30  -> debe reprogramarse dentro de 30 días.
--  plazo_dias = NULL -> sin nueva fecha; se registra en el mes real de ejecución
--                       una vez que el equipo se reintegra.
DROP TABLE IF EXISTS causales;
CREATE TABLE causales (
    codigo      TEXT PRIMARY KEY,   -- C1 … C8
    descripcion TEXT,
    plazo_dias  INTEGER,
    regla       TEXT
);
INSERT INTO causales (codigo, descripcion, plazo_dias, regla) VALUES
 ('C1','Imposibilidad de desocupar el equipo del paciente por indicación clínica',30,'Reprogramar dentro de 30 días'),
 ('C2','Equipo en servicio técnico',NULL,'Sin nueva fecha; se registra en el mes real de ejecución al reintegrarse'),
 ('C3','Equipo no operativo, a la espera de repuestos o accesorios',NULL,'Sin nueva fecha; se registra en el mes real de ejecución al reintegrarse'),
 ('C4','Equipo en préstamo a otro hospital o institución',NULL,'Sin nueva fecha; se registra en el mes real de ejecución al reintegrarse'),
 ('C5','No disponibilidad de horas hombre del funcionario SEC por alta carga laboral',30,'Reprogramar dentro de 30 días'),
 ('C6','No disponibilidad de horas hombre del servicio técnico externo',30,'Reprogramar dentro de 30 días'),
 ('C7','Ausencia justificada del funcionario SEC superior a 15 días',30,'Reprogramar dentro de 30 días'),
 ('C8','Contingencia hospitalaria',30,'Reprogramar dentro de 30 días');

-- Vista legible: mes con nombre, programa/resultado decodificados y la causal.
CREATE VIEW vista_programa AS
SELECT
    m.id_planilla  AS "ID",
    m.n_carpeta    AS "N° Carpeta",
    m.n_inventario AS "N° Inventario",
    m.equipo       AS "Equipo",
    m.servicio     AS "Servicio",
    m.ubicacion    AS "Ubicación",
    m.marca        AS "Marca",
    m.modelo       AS "Modelo",
    m.serie        AS "Serie",
    CASE m.mes WHEN 1 THEN 'Enero' WHEN 2 THEN 'Febrero' WHEN 3 THEN 'Marzo'
             WHEN 4 THEN 'Abril' WHEN 5 THEN 'Mayo' WHEN 6 THEN 'Junio'
             WHEN 7 THEN 'Julio' WHEN 8 THEN 'Agosto' WHEN 9 THEN 'Septiembre'
             WHEN 10 THEN 'Octubre' WHEN 11 THEN 'Noviembre' WHEN 12 THEN 'Diciembre'
             ELSE m.mes END                                        AS "Mes",
    CASE upper(m.programa)
         WHEN 'X'  THEN 'Programada'
         WHEN 'R'  THEN 'Reprogramada'
         WHEN 'RA' THEN 'Reprogramada (año anterior)'
         WHEN 'PM' THEN 'Puesta en marcha'
         WHEN 'BAJA' THEN 'Baja'
         ELSE m.programa END                                      AS "Programa",
    CASE
         WHEN m.resultado IS NULL OR m.resultado = '' THEN 'Pendiente'
         WHEN upper(m.resultado) = 'SI'    THEN 'Realizada'
         WHEN upper(m.resultado) = 'SI-RA' THEN 'Realizada (año anterior)'
         WHEN upper(m.resultado) IN ('C1','C2','C3','C4','C5','C6','C7','C8')
              THEN 'Reprogramada (' || m.resultado || ')'
         WHEN upper(m.resultado) = 'FS'   THEN 'Fuera de servicio'
         WHEN upper(m.resultado) = 'NO'   THEN 'No realizada'
         WHEN upper(m.resultado) = 'NU'   THEN 'No ubicable'
         WHEN upper(m.resultado) = 'BAJA' THEN 'Baja'
         ELSE m.resultado END                                     AS "Resultado",
    c.descripcion        AS "Causal",
    m.fecha_ejecucion      AS "Fecha de ejecución",
    m.ultima_actualizacion AS "Última actualización"
FROM mantenciones m
LEFT JOIN causales c ON upper(m.resultado) = c.codigo
ORDER BY m.clave, m.mes;
