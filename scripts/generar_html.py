#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera un archivo HTML autónomo (sin servidor) para explorar y gestionar los
equipos médicos:

  * Filtro por columna estilo Excel (casillas, uno o más valores).
  * Columnas ID y Equipo congeladas; gestor de columnas; chips de filtros.
  * Ficha lateral por equipo con:
      - Observaciones (estado técnico, precargado de la planilla) y Notas (libre).
      - Registro de Intervenciones / Pendientes / Recordatorios:
        cada uno con fecha, descripción, estado pendiente y fecha de
        vencimiento opcional. La tabla muestra el Estado (Vencido / Pendiente /
        Al día) y el próximo vencimiento, resaltando vencidos y próximos.
  * Persistencia local (localStorage) + respaldo a notas_equipos.json
    (Guardar/Cargar). scripts/aplicar_notas.py vuelca ese JSON a equipos.db.

Los datos se incrustan leyéndolos de equipos.db; los números de serie/inventario
conservan sus ceros a la izquierda (texto).

Uso:
    python scripts/generar_html.py            # equipos.db -> equipos_filtrable.html
"""
from __future__ import annotations

import argparse
import json
import sqlite3
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
DB_POR_DEFECTO = RAIZ / "equipos.db"
SALIDA_POR_DEFECTO = RAIZ / "equipos_filtrable.html"

# Columnas de SOLO LECTURA (orden y encabezados exactos de la planilla).
COLUMNAS_BASE = [
    ("id", "ID"), ("n_carpeta", "N° Carpeta"), ("n_inventario", "N° Inventario"),
    ("equipo", "Equipo"), ("servicio", "Servicio"), ("unidad", "Unidad"),
    ("ubicacion", "Ubicación"), ("procedencia", "Procedencia"), ("marca", "Marca"),
    ("modelo", "Modelo"), ("serie", "Serie"), ("anio_instalacion", "Año Instalación"),
    ("vida_util_residual", "Vida Útil Residual"), ("clasificacion", "Clasificación"),
    ("enu_baja", "ENU / Baja"),
]
# Columnas de texto EDITABLES (vienen de la BD como ROWS).
COLUMNAS_EDIT = [("observaciones", "Observaciones"), ("notas", "Notas")]
# Columnas VIRTUALES (calculadas en el navegador a partir de los registros).
COLUMNAS_VIRT = ["Estado", "Vence"]

COLUMNAS_ROW = COLUMNAS_BASE + COLUMNAS_EDIT          # lo que se incrusta como ROWS
NUMERICAS = [0, 1, 11]


def clave(serie, inventario, idv):
    s = str(serie).strip() if serie not in (None, "") else ""
    if s: return s
    i = str(inventario).strip() if inventario not in (None, "") else ""
    if i: return i
    return "#" + str(idv).strip()


def leer_datos(db: Path):
    con = sqlite3.connect(db)
    cols = ", ".join(c for c, _ in COLUMNAS_ROW)
    q = con.execute(f"SELECT {cols}, registros FROM equipos ORDER BY id").fetchall()
    con.close()
    rows, db_reg = [], {}
    n = len(COLUMNAS_ROW)
    for r in q:
        base = [("" if v is None else v) for v in r[:n]]
        rows.append(base)
        regtxt = r[n]
        if regtxt:
            try:
                lst = json.loads(regtxt)
            except Exception:
                lst = []
            if lst:
                db_reg[clave(base[10], base[2], base[0])] = lst
    return rows, db_reg


def leer_programa(db: Path):
    """Lee la tabla `mantenciones` de la base y devuelve las filas del programa.

    Cada elemento: {k: clave equipo, eq, se(serie), sv(servicio), m: 1-12,
    p: Programa, r: Resultado, fe: fecha de ejecución, ua: última actualización}.
    Devuelve [] si la tabla no existe todavía.
    """
    con = sqlite3.connect(db)
    try:
        q = con.execute(
            "SELECT clave, id_planilla, n_carpeta, n_inventario, equipo, servicio, ubicacion, marca, modelo, "
            "serie, mes, programa, resultado, fecha_ejecucion, ultima_actualizacion "
            "FROM mantenciones ORDER BY clave, mes"
        ).fetchall()
    except sqlite3.OperationalError:
        con.close()
        return []
    con.close()
    return [{"k": r[0], "id": r[1] or "", "car": r[2] or "", "inv": r[3] or "", "eq": r[4] or "", "sv": r[5] or "",
             "ub": r[6] or "", "ma": r[7] or "", "mo": r[8] or "", "se": r[9] or "",
             "m": r[10], "p": r[11] or "", "r": r[12] or "", "fe": r[13] or "", "ua": r[14] or ""}
            for r in q]


PLANTILLA = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Equipos médicos · Mantenciones 2026</title>
<style>
  :root{
    --bg:#f1f5f9; --surface:#ffffff; --surface-2:#f8fafc;
    --text:#0f172a; --muted:#64748b; --faint:#94a3b8; --border:#e2e8f0;
    --primary:#0d9488; --primary-strong:#0f766e; --primary-050:#f0fdfa; --primary-100:#ccfbf1;
    --edit:#fffdf3; --edit-border:#fde68a;
    --red:#b91c1c; --red-bg:#fee2e2; --amber:#92400e; --amber-bg:#fef3c7; --green:#166534; --green-bg:#dcfce7;
    --danger:#e11d48; --row-alt:#f8fafc; --row-hover:#ecfeff;
    --shadow-md:0 6px 20px rgba(15,23,42,.12);
    --radius:9px; --gap:8px; --rowpad:7px;
  }
  *{ box-sizing:border-box; }
  body{ margin:0; background:var(--bg); color:var(--text);
        font-family:"Inter",-apple-system,"Segoe UI",Roboto,Arial,sans-serif; font-size:13px;
        -webkit-font-smoothing:antialiased; display:flex; flex-direction:column; height:100vh; }
  button{ font:inherit; }
  ::selection{ background:var(--primary-100); }

  .appbar{ background:linear-gradient(180deg,#0f766e,#0d9488); color:#fff; padding:12px 20px; display:flex; align-items:center; gap:12px; }
  .appbar .logo{ width:34px; height:34px; border-radius:9px; background:rgba(255,255,255,.16); display:grid; place-items:center; font-size:18px; }
  .appbar h1{ margin:0; font-size:16px; font-weight:650; }
  .appbar .sub{ font-size:12px; color:#d1faf3; margin-top:1px; }

  .toolbar{ background:var(--surface); border-bottom:1px solid var(--border); padding:10px 20px; display:flex; gap:var(--gap); align-items:center; flex-wrap:wrap; }
  .search{ position:relative; }
  .search input{ padding:8px 12px 8px 32px; border:1px solid var(--border); border-radius:8px; min-width:230px; background:var(--surface-2); }
  .search input:focus{ outline:none; border-color:var(--primary); background:#fff; box-shadow:0 0 0 3px var(--primary-100); }
  .search svg{ position:absolute; left:10px; top:50%; transform:translateY(-50%); color:var(--faint); }
  .btn{ padding:8px 13px; border:1px solid var(--border); background:var(--surface); color:var(--text); border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; transition:.15s; }
  .btn:hover{ background:var(--surface-2); border-color:#cbd5e1; }
  .btn:active{ transform:translateY(1px); }
  .btn.primary{ background:var(--primary); color:#fff; border-color:var(--primary-strong); }
  .btn.primary:hover{ background:var(--primary-strong); }
  .btn.on{ background:var(--amber-bg); border-color:#f59e0b; color:var(--amber); }
  .btn.on.venc{ background:var(--red-bg); border-color:#ef4444; color:var(--red); }
  .spacer{ flex:1; }
  .dirty{ color:var(--danger); font-weight:600; font-size:12px; display:none; align-items:center; gap:6px; }
  .dirty.on{ display:inline-flex; }
  .dirty .pulse{ width:8px; height:8px; border-radius:50%; background:var(--danger); box-shadow:0 0 0 0 rgba(225,29,72,.5); animation:pulse 1.8s infinite; }
  @keyframes pulse{ 70%{ box-shadow:0 0 0 7px rgba(225,29,72,0); } 100%{ box-shadow:0 0 0 0 rgba(225,29,72,0); } }

  .substrip{ background:var(--surface); border-bottom:1px solid var(--border); padding:8px 20px; display:flex; gap:14px; align-items:center; flex-wrap:wrap; }
  .stats{ display:flex; gap:8px; flex-wrap:wrap; }
  .stat{ background:var(--surface-2); border:1px solid var(--border); border-radius:999px; padding:4px 11px; font-size:12px; color:var(--muted); transition:.12s; }
  .stat b{ color:var(--text); font-weight:650; }
  .stat.venc b{ color:var(--red); } .stat.pend b{ color:var(--amber); }
  .stat.click{ cursor:pointer; } .stat.click:hover{ background:#eef2f6; border-color:#cbd5e1; }
  .stat.act{ background:var(--primary-050); border-color:var(--primary-100); }
  .stat.act.venc{ background:var(--red-bg); border-color:#fecaca; }
  .stat.act.pend{ background:var(--amber-bg); border-color:#fde68a; }
  .chips{ display:flex; gap:6px; flex-wrap:wrap; align-items:center; }
  .chip{ background:var(--primary-050); border:1px solid var(--primary-100); color:var(--primary-strong); border-radius:999px; padding:4px 6px 4px 11px; font-size:12px; display:inline-flex; align-items:center; gap:7px; }
  .chip button{ border:none; background:rgba(13,148,136,.14); color:var(--primary-strong); cursor:pointer; width:17px; height:17px; border-radius:50%; line-height:1; }
  .chip button:hover{ background:rgba(13,148,136,.28); }
  .chips .clear-all{ color:var(--muted); font-size:12px; background:none; border:none; cursor:pointer; text-decoration:underline; }

  .table-wrap{ flex:1; overflow:auto; }
  table{ border-collapse:separate; border-spacing:0; width:100%; }
  thead th{ position:sticky; top:0; z-index:8; background:#0f766e; color:#fff; padding:0; border-right:1px solid rgba(255,255,255,.14); white-space:nowrap; font-weight:600; }
  thead th.edit{ background:#0e7490; } thead th.virt{ background:#155e63; }
  .th-inner{ display:flex; align-items:center; gap:6px; padding:9px 10px; }
  .th-label{ cursor:pointer; user-select:none; flex:1; display:inline-flex; align-items:center; gap:5px; }
  .th-label:hover{ color:#d1faf3; }
  .sort-ind{ font-size:10px; opacity:.9; min-width:9px; }
  .lapiz{ font-size:11px; opacity:.85; }
  .filtro-btn{ cursor:pointer; border:1px solid rgba(255,255,255,.4); background:rgba(255,255,255,.14); color:#fff; border-radius:6px; font-size:10px; line-height:1; padding:4px 5px; opacity:.62; transition:opacity .15s, background .15s; }
  .th-inner:hover .filtro-btn{ opacity:1; }
  .filtro-btn:hover{ background:rgba(255,255,255,.25); }
  .filtro-btn.activo{ background:#fde68a; color:#7c2d12; border-color:#f59e0b; opacity:1; }
  tbody td{ border-bottom:1px solid var(--border); border-right:1px solid var(--border); padding:var(--rowpad) 10px; white-space:nowrap; max-width:340px; overflow:hidden; text-overflow:ellipsis; vertical-align:top; background:var(--surface); }
  tbody tr:nth-child(even) td{ background:var(--row-alt); }
  tbody tr:hover td{ background:var(--row-hover); }
  tbody tr{ cursor:pointer; }
  td.num{ text-align:right; font-variant-numeric:tabular-nums; color:#334155; }
  .badge{ display:inline-block; padding:2px 9px; border-radius:999px; font-size:11px; font-weight:600; background:#e2e8f0; color:#334155; }
  .badge.inv{ background:#fee2e2; color:#b91c1c; } .badge.noinv{ background:#dcfce7; color:#166534; }
  .badge.fija{ background:#e0e7ff; color:#3730a3; } .badge.transp{ background:#fef3c7; color:#92400e; }
  .badge.venc{ background:var(--red-bg); color:var(--red); } .badge.pend{ background:var(--amber-bg); color:var(--amber); } .badge.ok{ background:var(--green-bg); color:var(--green); }
  .fecha-venc{ color:var(--red); font-weight:600; } .fecha-prox{ color:var(--amber); font-weight:600; }
  td.editable{ background:var(--edit); white-space:pre-wrap; min-width:170px; max-width:280px; overflow:visible; text-overflow:clip; cursor:text; }
  tbody tr:hover td.editable{ background:#fffdf0; }
  td.editable:focus{ outline:2px solid var(--primary); outline-offset:-2px; background:#fff; }
  td.editable:empty::before{ content:"✎ escribir…"; color:#c7b377; }
  td.freeze, th.freeze{ position:sticky; z-index:6; }
  thead th.freeze{ z-index:12; }
  .col-id{ left:0; width:56px; min-width:56px; max-width:56px; text-align:right; }
  .col-eq{ left:56px; min-width:150px; }
  th.col-id, th.col-eq{ background:#0f766e; }
  td.freeze-last{ box-shadow:6px 0 8px -8px rgba(15,23,42,.35); }
  td.col-id.urg-venc{ border-left:4px solid #ef4444; } td.col-id.urg-prox{ border-left:4px solid #f59e0b; }
  .empty{ padding:48px 20px; text-align:center; color:var(--muted); }
  .empty .big{ font-size:15px; color:var(--text); font-weight:600; margin-bottom:4px; }

  .menu{ position:fixed; z-index:1000; background:var(--surface); border:1px solid #cbd5e1; border-radius:10px; box-shadow:var(--shadow-md); font-size:13px; }
  .dropdown{ width:268px; }
  .dropdown .dd-top{ padding:10px; border-bottom:1px solid var(--border); }
  .dd-search{ width:100%; padding:7px 9px; border:1px solid var(--border); border-radius:7px; background:var(--surface-2); }
  .dd-search:focus{ outline:none; border-color:var(--primary); box-shadow:0 0 0 3px var(--primary-100); }
  .dd-ord{ display:flex; gap:6px; margin-top:8px; }
  .dd-ord button{ flex:1; padding:6px; border:1px solid var(--border); background:var(--surface-2); border-radius:7px; cursor:pointer; font-size:12px; color:var(--muted); }
  .dd-ord button:hover{ background:#eef2f6; color:var(--text); }
  .dd-list{ max-height:248px; overflow:auto; padding:6px 8px; }
  .dd-list label{ display:flex; align-items:center; gap:8px; padding:5px 4px; cursor:pointer; border-radius:6px; }
  .dd-list label:hover{ background:var(--surface-2); }
  .dd-list label.todos{ border-bottom:1px solid var(--border); margin-bottom:4px; padding-bottom:7px; font-weight:600; }
  .dd-val{ overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .dd-acc{ display:flex; gap:8px; padding:10px; border-top:1px solid var(--border); }
  .dd-acc button{ flex:1; padding:7px; border-radius:7px; border:1px solid var(--border); cursor:pointer; background:#fff; }
  .dd-acc .ok{ background:var(--primary); color:#fff; border-color:var(--primary-strong); }
  .dd-acc .ok:hover{ background:var(--primary-strong); }
  .colmenu{ width:240px; padding:8px; max-height:60vh; overflow:auto; }
  .colmenu .ti{ font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.5px; padding:4px 6px; }
  .colmenu label{ display:flex; align-items:center; gap:8px; padding:6px; border-radius:6px; cursor:pointer; }
  .colmenu label:hover{ background:var(--surface-2); }
  .colmenu label.lock{ opacity:.55; cursor:not-allowed; }

  .scrim{ position:fixed; inset:0; background:rgba(15,23,42,.38); opacity:0; pointer-events:none; transition:.2s; z-index:1100; }
  .scrim.on{ opacity:1; pointer-events:auto; }
  .drawer{ position:fixed; top:0; right:0; height:100%; width:min(470px,100%); background:var(--surface); box-shadow:-12px 0 32px rgba(15,23,42,.18); transform:translateX(100%); transition:transform .22s ease; z-index:1101; display:flex; flex-direction:column; }
  .drawer.on{ transform:translateX(0); }
  .dh{ padding:16px 18px; border-bottom:1px solid var(--border); background:linear-gradient(180deg,#0f766e,#0d9488); color:#fff; position:relative; }
  .dh .eyebrow{ font-size:11px; text-transform:uppercase; letter-spacing:.6px; color:#bff3ea; }
  .dh h2{ margin:3px 0 8px; font-size:18px; font-weight:680; }
  .dh .tags{ display:flex; gap:6px; flex-wrap:wrap; }
  .dh .tag{ background:rgba(255,255,255,.16); border-radius:999px; padding:3px 10px; font-size:12px; }
  .dh .x{ position:absolute; top:14px; right:14px; background:rgba(255,255,255,.16); border:none; color:#fff; width:30px; height:30px; border-radius:8px; cursor:pointer; font-size:16px; }
  .dh .x:hover{ background:rgba(255,255,255,.3); }
  .dbody{ padding:16px 18px; overflow:auto; flex:1; }
  .sechead{ font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:.4px; color:var(--muted); margin:4px 0 10px; padding-bottom:6px; border-bottom:1px solid var(--border); }
  .dl{ display:grid; grid-template-columns:120px 1fr; gap:8px 12px; margin:0 0 18px; background:var(--surface-2); border:1px solid var(--border); border-radius:10px; padding:13px 15px; }
  .dl dt{ color:var(--muted); font-size:12px; } .dl dd{ margin:0; font-size:13px; word-break:break-word; }
  .field{ margin-bottom:16px; }
  .field > label{ display:flex; justify-content:space-between; align-items:center; font-weight:650; margin-bottom:6px; }
  .field .saved{ font-size:11px; color:var(--primary); opacity:0; transition:.2s; }
  .field .saved.on{ opacity:1; }
  .field textarea{ width:100%; min-height:74px; resize:vertical; padding:10px; border:1px solid var(--edit-border); border-radius:8px; background:var(--edit); font:inherit; line-height:1.45; }
  .field textarea:focus{ outline:none; border-color:var(--primary); box-shadow:0 0 0 3px var(--primary-100); background:#fff; }
  .reg-add{ display:grid; grid-template-columns:1fr 1fr; gap:8px; margin-bottom:12px; }
  .reg-add input, .reg-add button{ padding:8px; border:1px solid var(--border); border-radius:7px; background:#fff; }
  .reg-add .desc{ grid-column:1/3; }
  .reg-add .chk{ display:flex; align-items:center; gap:7px; font-size:13px; padding:8px; background:var(--surface-2); border:1px solid var(--border); border-radius:7px; cursor:pointer; }
  .reg-add input[type=date]:focus, .reg-add input[type=text]:focus{ outline:none; border-color:var(--primary); box-shadow:0 0 0 3px var(--primary-100); }
  .reg-add button{ grid-column:1/3; background:var(--primary); color:#fff; border-color:var(--primary-strong); cursor:pointer; font-weight:600; }
  .reg-add button:hover{ background:var(--primary-strong); }
  .reg-add .lblv{ grid-column:1/3; font-size:11px; color:var(--muted); margin-top:-4px; }
  .reg-list{ list-style:none; margin:0; padding:0; display:flex; flex-direction:column; gap:7px; }
  .reg-item{ display:grid; grid-template-columns:auto 1fr auto; gap:9px; align-items:center; padding:9px 11px; border:1px solid var(--border); border-radius:9px; background:var(--surface-2); }
  .reg-item.pend{ border-left:4px solid #f59e0b; } .reg-item.venc{ border-left:4px solid #ef4444; }
  .reg-item .rf{ font-variant-numeric:tabular-nums; color:var(--muted); font-size:12px; white-space:nowrap; }
  .reg-item .rd{ word-break:break-word; }
  .reg-item .actions{ display:flex; gap:7px; align-items:center; }
  .due{ font-size:11px; padding:2px 8px; border-radius:999px; background:#e2e8f0; color:#334155; white-space:nowrap; }
  .due.venc{ background:var(--red-bg); color:var(--red); } .due.prox{ background:var(--amber-bg); color:var(--amber); }
  .pill{ cursor:pointer; border:none; border-radius:999px; padding:3px 11px; font-size:12px; font-weight:600; }
  .pill.pend{ background:var(--amber-bg); color:var(--amber); } .pill.ok{ background:var(--green-bg); color:var(--green); }
  .reg-item .del{ border:none; background:none; cursor:pointer; color:var(--faint); font-size:14px; }
  .reg-item .del:hover{ color:var(--danger); }
  .reg-empty{ color:var(--muted); font-size:12px; padding:6px 2px; }
  .dfoot{ padding:12px 18px; border-top:1px solid var(--border); display:flex; gap:8px; }

  #toasts{ position:fixed; right:18px; bottom:18px; z-index:1200; display:flex; flex-direction:column; gap:8px; }
  .toast{ background:#0f172a; color:#fff; padding:10px 14px; border-radius:9px; box-shadow:var(--shadow-md); font-size:13px; }
  .toast .ok{ color:#34d399; }
  .table-wrap::-webkit-scrollbar{ width:13px; height:13px; }
  .table-wrap::-webkit-scrollbar-thumb{ background:#cbd5e1; border-radius:9px; border:3px solid var(--bg); }
  .table-wrap::-webkit-scrollbar-thumb:hover{ background:#94a3b8; }
  :focus-visible{ outline:2px solid var(--primary); outline-offset:1px; border-radius:4px; }
  body.compact{ --rowpad:3px; font-size:12.5px; }

  /* ---------- Modo oscuro ---------- */
  body[data-theme="oscuro"]{
    --bg:#0b1220; --surface:#101a2e; --surface-2:#16223a; --text:#e6edf6; --muted:#9fb0c6; --faint:#6b7c93;
    --border:#26344f; --row-alt:#13203a; --row-hover:#1a2b48;
    --edit:#23241a; --edit-border:#5a4d1d;
    --primary-050:#0c2b27; --primary-100:#15534b;
    --red:#f87171; --red-bg:#3a1620; --amber:#fbbf24; --amber-bg:#39290f; --green:#4ade80; --green-bg:#0f2e1f;
    --shadow-md:0 8px 24px rgba(0,0,0,.55);
  }
  body[data-theme="oscuro"] td.num{ color:#b8c4d6; }
  body[data-theme="oscuro"] .badge{ background:#26344f; color:#cdd9ea; }
  body[data-theme="oscuro"] .due{ background:#26344f; color:#cdd9ea; }
  body[data-theme="oscuro"] .badge.inv{ background:#3a1620; color:#f6a9a9; }
  body[data-theme="oscuro"] .badge.noinv{ background:#0f2e1f; color:#86efac; }
  body[data-theme="oscuro"] .badge.fija{ background:#1e2553; color:#aab4f5; }
  body[data-theme="oscuro"] .badge.transp{ background:#39290f; color:#f4cf86; }
  body[data-theme="oscuro"] .reg-add input{ background:var(--surface); color:var(--text); }
  body[data-theme="oscuro"] .dd-acc button{ background:var(--surface); color:var(--text); }
  body[data-theme="oscuro"] td.editable:focus, body[data-theme="oscuro"] .field textarea:focus{ background:var(--surface-2); }
  body[data-theme="oscuro"] .table-wrap::-webkit-scrollbar-thumb{ background:#33405a; border-color:var(--bg); }
  body[data-theme="oscuro"] .cards::-webkit-scrollbar-thumb{ background:#33405a; }

  /* ---------- Panel Resumen (dashboard) ---------- */
  .dashboard{ background:var(--surface); border-bottom:1px solid var(--border); padding:0 20px; max-height:0; overflow:hidden; transition:max-height .25s ease, padding .2s ease; }
  .dashboard.on{ max-height:60vh; padding:14px 20px; overflow:auto; }
  .dash-grid{ display:grid; grid-template-columns:1.5fr 1fr 1fr; gap:14px; }
  @media (max-width:1000px){ .dash-grid{ grid-template-columns:1fr; } }
  .card-d{ background:var(--surface-2); border:1px solid var(--border); border-radius:12px; padding:14px 15px; }
  .card-d h3{ margin:0 0 12px; font-size:13px; display:flex; justify-content:space-between; gap:8px; align-items:baseline; }
  .card-d h3 small{ color:var(--muted); font-weight:500; font-size:11px; }
  .bar-row{ display:grid; grid-template-columns:116px 1fr 52px; gap:9px; align-items:center; margin-bottom:7px; cursor:pointer; font-size:12px; }
  .bar-row .nm{ overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .bar-row .ct{ text-align:right; font-variant-numeric:tabular-nums; color:var(--muted); }
  .bar-track{ background:var(--bg); border-radius:6px; height:15px; position:relative; overflow:hidden; }
  .bar-fill{ background:var(--primary); height:100%; border-radius:6px; }
  .bar-pend{ position:absolute; top:0; left:0; height:100%; background:repeating-linear-gradient(45deg,var(--amber),var(--amber) 4px,transparent 4px,transparent 8px); }
  .bar-row:hover .bar-track{ outline:2px solid var(--primary-100); }
  .seg{ display:flex; height:20px; border-radius:7px; overflow:hidden; border:1px solid var(--border); }
  .seg span{ height:100%; }
  .legend{ margin-top:10px; display:flex; flex-direction:column; gap:6px; font-size:12px; }
  .legend .li{ display:flex; align-items:center; gap:8px; cursor:pointer; }
  .legend .li:hover{ color:var(--primary-strong); }
  .legend .sw{ width:12px; height:12px; border-radius:3px; flex:none; }
  .legend .lc{ margin-left:auto; color:var(--muted); font-variant-numeric:tabular-nums; }

  /* ---------- Vista de tarjetas ---------- */
  .cards{ display:none; grid-template-columns:repeat(auto-fill,minmax(270px,1fr)); gap:12px; padding:16px 20px; overflow:auto; flex:1; align-content:start; }
  .cards.on{ display:grid; }
  .eqcard{ background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:13px 14px; cursor:pointer; transition:.12s; }
  .eqcard:hover{ border-color:var(--primary); box-shadow:var(--shadow-md); transform:translateY(-2px); }
  .eqcard.urg-venc{ border-left:4px solid #ef4444; } .eqcard.urg-prox{ border-left:4px solid #f59e0b; }
  .eqcard h4{ margin:0 0 3px; font-size:14px; display:flex; justify-content:space-between; gap:8px; align-items:flex-start; }
  .eqcard .meta{ color:var(--muted); font-size:12px; margin-bottom:7px; }
  .eqcard .row{ display:flex; gap:6px; flex-wrap:wrap; align-items:center; font-size:12px; }

  /* ---------- Vista Programa de mantención ---------- */
  .badge.repro{ background:#e0e7ff; color:#3730a3; }
  .badge.prog-x{ background:#dbeafe; color:#1e40af; }
  .badge.prog-pm{ background:#ccfbf1; color:#0f766e; }
  .badge.nu{ background:#ffedd5; color:#9a3412; }
  body[data-theme="oscuro"] .badge.repro{ background:#1e2553; color:#aab4f5; }
  body[data-theme="oscuro"] .badge.prog-x{ background:#15294d; color:#9dc0f7; }
  body[data-theme="oscuro"] .badge.prog-pm{ background:#0d2e2a; color:#7fe0d2; }
  body[data-theme="oscuro"] .badge.nu{ background:#3a230f; color:#f4bd86; }
  .programa{ display:none; flex-direction:column; flex:1; overflow:hidden; }
  .programa.on{ display:flex; }
  .prog-bar{ display:flex; gap:10px; align-items:center; padding:10px 20px; border-bottom:1px solid var(--border); background:var(--surface); flex-wrap:wrap; }
  .prog-bar strong{ font-size:13px; }
  .prog-sel{ padding:7px 10px; border:1px solid var(--border); border-radius:8px; background:var(--surface-2); color:var(--text); }
  .prog-sel:focus{ outline:none; border-color:var(--primary); box-shadow:0 0 0 3px var(--primary-100); }
  .prog-info{ margin-left:auto; color:var(--muted); font-size:12px; }
  .prog-info b{ color:var(--text); }
  .prog-table-wrap{ overflow:auto; flex:1; }
  .prog-table{ border-collapse:separate; border-spacing:0; width:100%; }
  .prog-table thead th{ position:sticky; top:0; z-index:5; background:#0f766e; color:#fff; padding:0; text-align:left; white-space:nowrap; border-right:1px solid rgba(255,255,255,.14); font-size:12px; }
  .prog-table thead .th-inner{ padding:8px 9px 8px 11px; }
  .prog-table thead .th-label{ flex:1; }
  .prog-table td{ border-bottom:1px solid var(--border); border-right:1px solid var(--border); padding:6px 11px; white-space:nowrap; font-size:12.5px; }
  .prog-table tbody tr:nth-child(even) td{ background:var(--row-alt); }
  .prog-table tbody tr:hover td{ background:var(--row-hover); }
  .prog-table .eqn{ font-weight:600; }
  .prog-table input.fe{ padding:4px 7px; border:1px solid var(--edit-border); border-radius:6px; background:var(--edit); font:inherit; color:var(--text); }
  .prog-table input.fe:focus{ outline:none; border-color:var(--primary); box-shadow:0 0 0 3px var(--primary-100); }
  .prog-table .ua{ color:var(--muted); font-variant-numeric:tabular-nums; }
  .btn.sm{ padding:5px 10px; font-size:12px; }
  .prog-empty{ padding:40px 20px; text-align:center; color:var(--muted); }

  /* ---------- Ayuda (atajos) ---------- */
  .help-scrim{ position:fixed; inset:0; background:rgba(15,23,42,.55); display:none; place-items:center; z-index:1300; }
  .help-scrim.on{ display:grid; }
  .help{ background:var(--surface); color:var(--text); border-radius:14px; padding:22px 24px; width:min(420px,92vw); box-shadow:var(--shadow-md); }
  .help h3{ margin:0 0 14px; font-size:15px; }
  .help dl{ display:grid; grid-template-columns:auto 1fr; gap:9px 14px; margin:0; }
  .help dt{ text-align:right; } .help dd{ margin:0; color:var(--muted); }
  kbd{ background:var(--surface-2); border:1px solid var(--border); border-bottom-width:2px; border-radius:6px; padding:2px 7px; font-family:ui-monospace,monospace; font-size:12px; color:var(--text); }

  @media (max-width:720px){ .appbar .sub{ display:none; } .search input{ min-width:150px; } }
</style>
</head>
<body>
<div class="appbar">
  <div class="logo">🩺</div>
  <div>
    <h1>Equipos médicos — Mantenciones Preventivas 2026</h1>
    <div class="sub">Filtra por columna, abre una fila para ver su ficha y registra intervenciones, pendientes y recordatorios.</div>
  </div>
</div>

<div class="toolbar">
  <div class="search">
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>
    <input type="search" id="busqueda" placeholder="Buscar en todas las columnas…">
  </div>
  <button class="btn" id="btnDash">📊 Resumen</button>
  <button class="btn" id="btnPrograma">📅 Programa MP</button>
  <button class="btn" id="btnVista">▤ Tarjetas</button>
  <button class="btn" id="btnCols">▦ Columnas</button>
  <button class="btn" id="btnDensidad">≣ Densidad</button>
  <button class="btn" id="limpiar">✕ Limpiar</button>
  <div class="spacer"></div>
  <span class="dirty" id="dirty"><span class="pulse"></span> notas sin respaldar</span>
  <button class="btn" id="btnTema" title="Modo oscuro">🌙</button>
  <button class="btn" id="btnAyuda" title="Atajos de teclado (?)">⌨</button>
  <button class="btn primary" id="guardar">💾 Guardar</button>
  <button class="btn" id="cargar">📂 Cargar</button>
  <button class="btn" id="exportar">⤓ CSV</button>
  <input type="file" id="archivo" accept=".json,application/json" style="display:none">
</div>

<div class="substrip" id="substrip">
  <div class="stats" id="stats"></div>
  <div class="chips" id="chips"></div>
</div>

<div class="dashboard" id="dashboard"></div>

<div class="programa" id="programa">
  <div class="prog-bar">
    <button class="btn" id="progBack">← Equipos</button>
    <strong>Programa de mantención</strong>
    <button class="btn sm" id="progClear" hidden>✕ Limpiar filtros</button>
    <span class="prog-info" id="progInfo"></span>
  </div>
  <div class="prog-table-wrap">
    <table class="prog-table">
      <thead><tr id="progHead"></tr></thead>
      <tbody id="progBody"></tbody>
    </table>
  </div>
</div>

<div class="table-wrap" id="table-wrap">
  <table><thead><tr id="encabezado"></tr></thead><tbody id="cuerpo"></tbody></table>
  <div class="empty" id="vacio" style="display:none"><div class="big">Sin resultados</div><div>Ningún equipo coincide con los filtros actuales.</div></div>
</div>
<div class="cards" id="cards"></div>

<div class="scrim" id="scrim"></div>
<aside class="drawer" id="drawer" role="dialog" aria-modal="true" aria-label="Ficha del equipo">
  <div class="dh">
    <button class="x" id="cerrarDrawer" aria-label="Cerrar">✕</button>
    <div class="eyebrow" id="dEyebrow"></div>
    <h2 id="dTitulo"></h2>
    <div class="tags" id="dTags"></div>
  </div>
  <div class="dbody">
    <div class="sechead">Intervenciones · pendientes · recordatorios</div>
    <div id="dRegistros"></div>
    <div class="sechead">Datos del equipo</div>
    <dl class="dl" id="dDatos"></dl>
    <div class="sechead">Texto libre</div>
    <div id="dCampos"></div>
  </div>
  <div class="dfoot">
    <button class="btn primary" id="dGuardar" style="flex:1">💾 Guardar notas en archivo</button>
    <button class="btn" id="dCerrar2">Cerrar</button>
  </div>
</aside>

<div class="help-scrim" id="help">
  <div class="help" role="dialog" aria-modal="true" aria-label="Atajos de teclado">
    <h3>⌨ Atajos de teclado</h3>
    <dl>
      <dt><kbd>/</kbd></dt><dd>Buscar</dd>
      <dt><kbd>t</kbd></dt><dd>Cambiar entre Tabla y Tarjetas</dd>
      <dt><kbd>d</kbd></dt><dd>Mostrar/ocultar el panel Resumen</dd>
      <dt><kbd>p</kbd></dt><dd>Vista Programa de mantención</dd>
      <dt><kbd>o</kbd></dt><dd>Modo claro / oscuro</dd>
      <dt><kbd>?</kbd></dt><dd>Mostrar esta ayuda</dd>
      <dt><kbd>Esc</kbd></dt><dd>Cerrar ficha, menús y ayuda</dd>
    </dl>
  </div>
</div>

<div id="toasts"></div>

<script>
const HEADERS = __HEADERS__;
const ROWS = __ROWS__;
const DB_REGISTROS = __DBREG__;
const PROGRAMA = __PROGRAMA__;
const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const NUMERICAS = new Set(__NUMERICAS__);
const NB = __NB__;
const IDX_OBS = __IDX_OBS__, IDX_NOTAS = __IDX_NOTAS__, IDX_ESTADO = __IDX_ESTADO__, IDX_VENCE = __IDX_VENCE__;
const EDIT = { [IDX_OBS]:'observaciones', [IDX_NOTAS]:'notas' };
const EDIT_COLS = new Set([IDX_OBS, IDX_NOTAS]);
const DERIV = new Set([IDX_ESTADO, IDX_VENCE]);
const FREEZE = { 0:'col-id', 3:'col-eq' };
const LS_KEY = 'aqf12_notas_v1', LS_COLS = 'aqf12_cols_v1';
const VACIO = "(Vacías)", DIAS_PROX = 30;

let notas = {};
try { notas = JSON.parse(localStorage.getItem(LS_KEY) || '{}') || {}; } catch(e){ notas = {}; }
const DEFAULT_OCULTO = [1, 5, 7, 11, 14];   // N° Carpeta, Unidad, Procedencia, Año Instalación, ENU/Baja
let oculto;
try { const g = localStorage.getItem(LS_COLS); oculto = new Set(g ? JSON.parse(g) : DEFAULT_OCULTO); } catch(e){ oculto = new Set(DEFAULT_OCULTO); }
let dirty = false, ordenCol = null, ordenDir = 1, menuAbierto = null, drawerRow = null, filtroVida = false;
const filtros = new Map();
const LS_THEME = 'aqf12_theme_v1', LS_VIEW = 'aqf12_view_v1';
let vista = 'tabla', dashAbierto = false, tema = 'claro';
const PCOLS = [
  {k:'id',t:'ID'},{k:'car',t:'N° Carpeta'},{k:'inv',t:'N° Inventario'},{k:'eq',t:'Equipo'},{k:'sv',t:'Servicio'},
  {k:'ub',t:'Ubicación'},{k:'ma',t:'Marca'},{k:'mo',t:'Modelo'},{k:'se',t:'Serie'},
  {k:'mes',t:'Mes'},{k:'prog',t:'Programa'},{k:'res',t:'Resultado'},
  {k:'fe',t:'Fecha de ejecución'},{k:'ua',t:'Última actualización'}
];
const progFiltros = new Map();   // k -> Set de valores permitidos (estilo Excel)
try { vista = localStorage.getItem(LS_VIEW) || 'tabla'; } catch(e){}
try { tema = localStorage.getItem(LS_THEME) || 'claro'; } catch(e){}

const $ = s => document.querySelector(s);
const norm = v => (v === null || v === undefined) ? "" : String(v);
const claveFiltro = v => { const s = norm(v).trim(); return s === "" ? VACIO : s; };
const visible = ci => !oculto.has(ci);
function fechaISO(off=0){ const d = new Date(); d.setDate(d.getDate()+off);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
const HOY = fechaISO(0), PROX = fechaISO(DIAS_PROX);

function claveEquipo(i){
  const serie = norm(ROWS[i][10]).trim(); if (serie) return serie;
  const inv = norm(ROWS[i][2]).trim();    if (inv)   return inv;
  return "#" + norm(ROWS[i][0]).trim();
}
function registrosDe(i){ const k = claveEquipo(i);
  if (notas[k] && Array.isArray(notas[k].registros)) return notas[k].registros;
  return DB_REGISTROS[k] || []; }
function regsActuales(i){ const k = claveEquipo(i);
  if (!notas[k] || !Array.isArray(notas[k].registros)){
    const base = (DB_REGISTROS[k] || []).map(x => ({...x}));
    (notas[k] || (notas[k] = {})).registros = base; }
  return notas[k].registros; }
function guardar(){ try { localStorage.setItem(LS_KEY, JSON.stringify(notas)); } catch(e){} marcarDirty(true); }
function marcarDirty(v){ dirty = v; $('#dirty').classList.toggle('on', v); }

function estadoDe(i){ const r = registrosDe(i); if (!r.length) return "";
  const pend = r.filter(x => x.p); if (!pend.length) return "Al día";
  if (pend.some(x => x.v && x.v < HOY)) return "Vencido"; return "Pendiente"; }
function venceDe(i){ const p = registrosDe(i).filter(x => x.p && x.v);
  return p.length ? p.reduce((m,x) => (m==="" || x.v<m) ? x.v : m, "") : ""; }
function urgenciaDe(i){ const e = estadoDe(i); if (e === "Vencido") return "venc";
  const v = venceDe(i); if (v && v <= PROX) return "prox"; return ""; }
function claseVenc(r){ if (!r.p || !r.v) return ""; if (r.v < HOY) return "venc"; if (r.v <= PROX) return "prox"; return ""; }

function getCell(i, ci){
  if (ci === IDX_ESTADO) return estadoDe(i);
  if (ci === IDX_VENCE)  return venceDe(i);
  if (EDIT_COLS.has(ci)){ const f = EDIT[ci], k = claveEquipo(i), ov = notas[k];
    if (ov && Object.prototype.hasOwnProperty.call(ov, f)) return ov[f]; return norm(ROWS[i][ci]); }
  return ROWS[i][ci];
}
function setCell(i, ci, valor){ const f = EDIT[ci], k = claveEquipo(i);
  (notas[k] || (notas[k] = {}))[f] = valor; guardar(); }

function comparar(col, a, b){
  if (NUMERICAS.has(col)){ const na = parseFloat(a), nb = parseFloat(b);
    return (isNaN(na)?Infinity:na) - (isNaN(nb)?Infinity:nb); }
  return String(a).localeCompare(String(b), 'es', {numeric:true, sensitivity:'base'});
}
function pasaExcepto(i, excepto){ for (const [col, set] of filtros){ if (col === excepto) continue;
  if (!set.has(claveFiltro(getCell(i, col)))) return false; } return true; }
function coincideBusqueda(i){ const q = $('#busqueda').value.trim().toLowerCase(); if (!q) return true;
  for (let c = 0; c < HEADERS.length; c++) if (norm(getCell(i, c)).toLowerCase().includes(q)) return true; return false; }
function indicesVisibles(){ let v = [];
  for (let i = 0; i < ROWS.length; i++){
    if (filtroVida){ const n = parseFloat(norm(getCell(i, 12))); if (isNaN(n) || n >= 0) continue; }
    if (pasaExcepto(i, -1) && coincideBusqueda(i)) v.push(i); }
  if (ordenCol !== null) v.sort((x, y) => ordenDir * comparar(ordenCol, getCell(x, ordenCol), getCell(y, ordenCol))); return v; }
function valoresDisponibles(col){ const set = new Set();
  for (let i = 0; i < ROWS.length; i++) if (pasaExcepto(i, col) && coincideBusqueda(i)) set.add(claveFiltro(getCell(i, col)));
  const arr = [...set]; arr.sort((a, b) => { if (a===VACIO) return 1; if (b===VACIO) return -1; return comparar(col, a, b); }); return arr; }
function valoresTotal(col){ const s = new Set(); for (let i=0;i<ROWS.length;i++) s.add(claveFiltro(getCell(i,col))); return s.size; }

function celdaBadge(ci, s){
  if (ci === 13){ const t = s.toLowerCase();
    const cls = t.includes('invasiv') && !t.includes('no') ? 'inv' : (t==='no invasivo'?'noinv':(t.includes('fija')?'fija':'transp'));
    return s ? `<span class="badge ${cls}">${s}</span>` : ''; }
  if (ci === IDX_ESTADO){ if (!s) return '';
    const cls = s==='Vencido'?'venc':(s==='Pendiente'?'pend':'ok'); return `<span class="badge ${cls}">${s}</span>`; }
  if (ci === 12){ if (!s) return null;
    if (/disponible/i.test(s)) return `<span class="badge" title="Sin uso registrado">${s}</span>`;
    const n = parseFloat(s); if (isNaN(n)) return null;
    const cls = n < 0 ? 'venc' : (n <= 2 ? 'pend' : 'ok');
    const t = n < 0 ? 'Vida útil agotada' : (n <= 2 ? 'Por agotar' : 'Vida útil vigente');
    return `<span class="badge ${cls}" title="${t} (años)">${s}</span>`; }
  return null;
}

function mostrarVista(){
  const prog = vista === 'programa';
  $('#table-wrap').style.display = (vista === 'tabla') ? '' : 'none';
  $('#cards').classList.toggle('on', vista === 'tarjetas');
  $('#programa').classList.toggle('on', prog);
  $('#substrip').style.display = prog ? 'none' : '';
  $('#btnPrograma').classList.toggle('on', prog);
  if (prog){ dashAbierto = false; $('#dashboard').classList.remove('on'); $('#btnDash').classList.remove('on'); }
}
function render(){
  if (vista === 'programa'){ mostrarVista(); renderPrograma(); return; }
  const vis = indicesVisibles();
  if (vista === 'tarjetas') renderTarjetas(vis); else renderTabla(vis);
  mostrarVista();
  renderStats(vis.length); renderChips();
  if (dashAbierto) renderDashboard();
  document.querySelectorAll('.filtro-btn').forEach(b => b.classList.toggle('activo', filtros.has(+b.dataset.col)));
  document.querySelectorAll('.sort-ind').forEach(el => { const c = +el.dataset.col;
    el.textContent = (ordenCol === c) ? (ordenDir === 1 ? '▲' : '▼') : ''; });
}
function renderTabla(vis){
  const cuerpo = $('#cuerpo');
  const frag = document.createDocumentFragment();
  const cols = [...Array(HEADERS.length).keys()].filter(visible);
  for (const i of vis){
    const tr = document.createElement('tr'); tr.dataset.i = i;
    const urg = urgenciaDe(i);
    for (const ci of cols){
      const td = document.createElement('td');
      if (FREEZE[ci]){ td.classList.add('freeze', FREEZE[ci]); if (ci === 3) td.classList.add('freeze-last');
        if (ci === 0 && urg) td.classList.add(urg==='venc'?'urg-venc':'urg-prox'); }
      if (EDIT_COLS.has(ci)){
        td.classList.add('editable'); td.contentEditable = 'true'; td.spellcheck = false;
        td.dataset.i = i; td.dataset.ci = ci; td.textContent = norm(getCell(i, ci));
        td.addEventListener('input', onInlineEdit);
      } else if (ci === IDX_VENCE){
        const v = venceDe(i); if (v){ const c = v<HOY?'fecha-venc':(v<=PROX?'fecha-prox':'');
          td.innerHTML = `<span class="${c}">${v}</span>`; }
      } else {
        const s = norm(getCell(i, ci)); const b = celdaBadge(ci, s);
        if (b !== null){ td.innerHTML = b; }
        else { if (NUMERICAS.has(ci)) td.classList.add('num'); td.textContent = s; }
      }
      tr.appendChild(td);
    }
    frag.appendChild(tr);
  }
  cuerpo.replaceChildren(frag);
  $('#vacio').style.display = vis.length ? 'none' : 'block';
}
function renderTarjetas(vis){
  const cont = $('#cards');
  if (!vis.length){ cont.innerHTML = '<div class="empty"><div class="big">Sin resultados</div><div>Ningún equipo coincide con los filtros.</div></div>'; return; }
  const frag = document.createDocumentFragment();
  for (const i of vis){
    const c = document.createElement('div'); c.className = 'eqcard'; c.dataset.i = i;
    const urg = urgenciaDe(i); if (urg) c.classList.add(urg==='venc'?'urg-venc':'urg-prox');
    const estado = estadoDe(i), vence = venceDe(i);
    const eb = estado ? `<span class="badge ${estado==='Vencido'?'venc':(estado==='Pendiente'?'pend':'ok')}">${estado}</span>` : '';
    const vu = celdaBadge(12, norm(getCell(i,12))) || '';
    const cl = celdaBadge(13, norm(getCell(i,13))) || '';
    const marca = [norm(getCell(i,8)), norm(getCell(i,9))].filter(Boolean).join(' ');
    const serie = norm(getCell(i,10)) ? '· Serie ' + norm(getCell(i,10)) : '';
    const venceChip = vence ? `<span class="due ${vence<HOY?'venc':(vence<=PROX?'prox':'')}">vence ${vence}</span>` : '';
    c.innerHTML = `<h4><span>${norm(getCell(i,3)) || '—'}</span>${eb}</h4>
      <div class="meta">${[norm(getCell(i,4)), norm(getCell(i,6))].filter(Boolean).join(' · ') || '—'}</div>
      <div class="row">${marca} ${serie}</div>
      <div class="row" style="margin-top:9px">${vu} ${cl} ${venceChip}</div>`;
    frag.appendChild(c);
  }
  cont.replaceChildren(frag);
}
function onInlineEdit(ev){ const i = +ev.target.dataset.i, ci = +ev.target.dataset.ci;
  setCell(i, ci, ev.target.innerText);
  if (drawerRow === i){ const ta = $('#ta_'+ci); if (ta && ta.value !== ev.target.innerText) ta.value = ev.target.innerText; }
  tostarGuardado(); }
function renderStats(filtrados){
  let pend = 0, venc = 0, vu = 0;
  for (let i = 0; i < ROWS.length; i++){ const e = estadoDe(i); if (e==='Vencido'){ venc++; pend++; } else if (e==='Pendiente') pend++;
    const n = parseFloat(norm(getCell(i, 12))); if (!isNaN(n) && n < 0) vu++; }
  const cont = $('#stats'); cont.replaceChildren();
  const mk = (html, o={}) => { const el = document.createElement('span');
    el.className = 'stat' + (o.cls ? ' '+o.cls : '') + (o.click ? ' click' : '') + (o.act ? ' act' : '');
    if (o.click) el.onclick = o.click; el.innerHTML = html; return el; };
  cont.append(
    mk(`<b>${ROWS.length.toLocaleString('es')}</b> equipos`),
    mk(`<b>${filtrados.toLocaleString('es')}</b> en pantalla`),
    mk(`<b>${pend}</b> con pendientes`, {cls:'pend', click:()=>toggleEstado(['Pendiente','Vencido']), act:estadoFiltrado(['Pendiente','Vencido'])}),
    mk(`<b>${venc}</b> vencidos`, {cls:'venc', click:()=>toggleEstado(['Vencido']), act:estadoFiltrado(['Vencido'])}),
    mk(`<b>${vu}</b> vida útil vencida`, {cls:'venc', click:toggleVida, act:filtroVida}),
  );
}
function renderChips(){ const cont = $('#chips'); cont.innerHTML = ''; if (!filtros.size && !filtroVida) return;
  for (const [col, set] of filtros){ const txt = set.size === 1 ? [...set][0] : `${set.size} de ${valoresTotal(col)} valores`;
    const chip = document.createElement('span'); chip.className = 'chip';
    chip.innerHTML = `<span><b>${HEADERS[col]}:</b> ${txt}</span>`;
    const x = document.createElement('button'); x.textContent = '✕'; x.title = 'Quitar filtro';
    x.onclick = () => { filtros.delete(col); render(); }; chip.appendChild(x); cont.appendChild(chip); }
  if (filtroVida){ const chip = document.createElement('span'); chip.className = 'chip';
    chip.innerHTML = '<span><b>Vida útil:</b> vencida</span>';
    const x = document.createElement('button'); x.textContent = '✕'; x.title = 'Quitar filtro';
    x.onclick = () => { filtroVida = false; render(); }; chip.appendChild(x); cont.appendChild(chip); }
  const clr = document.createElement('button'); clr.className = 'clear-all'; clr.textContent = 'Limpiar todo';
  clr.onclick = () => { filtros.clear(); filtroVida = false; render(); }; cont.appendChild(clr); }

/* ---------- Tema / vista / ayuda ---------- */
function aplicarTema(){ document.body.setAttribute('data-theme', tema==='oscuro'?'oscuro':'claro');
  $('#btnTema').textContent = tema==='oscuro'?'☀️':'🌙'; $('#btnTema').title = tema==='oscuro'?'Modo claro':'Modo oscuro'; }
function toggleTema(){ tema = tema==='oscuro'?'claro':'oscuro'; try{ localStorage.setItem(LS_THEME, tema); }catch(e){} aplicarTema(); }
function toggleVista(){ vista = vista==='tabla'?'tarjetas':'tabla'; try{ localStorage.setItem(LS_VIEW, vista); }catch(e){}
  $('#btnVista').textContent = vista==='tabla'?'▤ Tarjetas':'▦ Tabla'; cerrarMenu(); render(); }
function toggleDash(){ if (vista==='programa'){ vista='tabla'; render(); }
  dashAbierto = !dashAbierto; $('#dashboard').classList.toggle('on', dashAbierto);
  $('#btnDash').classList.toggle('on', dashAbierto); if (dashAbierto) renderDashboard(); }
function toggleAyuda(){ $('#help').classList.toggle('on'); }
function cerrarAyuda(){ $('#help').classList.remove('on'); }

/* ---------- Vista Programa de mantención ---------- */
function togglePrograma(){ vista = (vista==='programa') ? 'tabla' : 'programa';
  try{ localStorage.setItem(LS_VIEW, vista); }catch(e){} render(); }
function ahora(){ const d = new Date(), p = n => String(n).padStart(2,'0');
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`; }
function resultadoEstado(r){ const s = (r||'').toString().trim().toUpperCase();
  if (s==='') return 'pendiente';
  if (s==='SI' || s==='SI-RA') return 'realizado';
  if (/^C[1-8]$/.test(s)) return 'reprogramado';
  if (s==='NO') return 'no'; if (s==='NU') return 'nu'; if (s==='FS') return 'fs';
  if (s==='BAJA') return 'baja'; return 'otro'; }
function progLabel(p){ const s = (p||'').toString().trim().toUpperCase();
  if (s==='') return '';
  return ({X:'Programada', R:'Reprogramada', RA:'Reprog. año anterior', PM:'Puesta en marcha', BAJA:'Baja'})[s] || (p||'').toString().trim(); }
function resLabel(r){ const raw = (r||'').toString().trim(), s = raw.toUpperCase();
  if (s==='') return 'Pendiente';
  if (s==='SI') return 'Realizada';
  if (s==='SI-RA') return 'Realizada año ant.';
  if (/^C[1-8]$/.test(s)) return `Reprogramada (${raw})`;
  if (s==='FS') return 'Fuera de servicio';
  if (s==='NO') return 'No realizada';
  if (s==='NU') return 'No ubicable';
  if (s==='BAJA') return 'Baja';
  return raw; }
function etiquetaResultado(r){ const s = (r||'').toString().trim().toUpperCase(), L = resLabel(r);
  if (s==='') return '<span class="badge pend">Pendiente</span>';
  if (s==='SI'||s==='SI-RA') return `<span class="badge ok" title="${s==='SI'?'Mantención Preventiva Realizada':'Mantención de Año Anterior Realizada'}">${L}</span>`;
  if (/^C[1-8]$/.test(s)) return `<span class="badge repro" title="Mantención Preventiva Reprogramada (ver causales)">${L}</span>`;
  if (s==='FS'||s==='NO') return `<span class="badge venc" title="${s==='FS'?'Fuera de Servicio':'No Realizada'}">${L}</span>`;
  if (s==='NU') return `<span class="badge nu" title="No Ubicable">${L}</span>`;
  if (s==='BAJA') return '<span class="badge" title="Equipo Dado de Baja">Baja</span>';
  return `<span class="badge">${L||'—'}</span>`; }
function etiquetaPrograma(p){ const s = (p||'').toString().trim().toUpperCase(), L = progLabel(p);
  if (s==='') return '—';
  if (s==='X') return `<span class="badge prog-x" title="Mantención Preventiva Programada">${L}</span>`;
  if (s==='R') return `<span class="badge repro" title="Mantención Preventiva Reprogramada">${L}</span>`;
  if (s==='RA') return `<span class="badge repro" title="Reprogramada de Año Anterior">${L}</span>`;
  if (s==='PM') return `<span class="badge prog-pm" title="Puesta en Marcha">${L}</span>`;
  if (s==='BAJA') return '<span class="badge" title="Equipo Dado de Baja">Baja</span>';
  return `<span class="badge">${L}</span>`; }
function progFE(x){ const ov=(notas[x.k]&&notas[x.k].prog&&notas[x.k].prog[x.m])||null; return ov?(ov.fe||''):(x.fe||''); }
function progUA(x){ const ov=(notas[x.k]&&notas[x.k].prog&&notas[x.k].prog[x.m])||null; return ov?(ov.ua||''):(x.ua||''); }
function colVal(x,k){ switch(k){
    case 'id': return norm(x.id); case 'car': return norm(x.car); case 'inv': return norm(x.inv); case 'eq': return norm(x.eq);
    case 'sv': return norm(x.sv); case 'ub': return norm(x.ub); case 'ma': return norm(x.ma);
    case 'mo': return norm(x.mo); case 'se': return norm(x.se); case 'mes': return MESES[x.m-1];
    case 'prog': return progLabel(x.p); case 'res': return resLabel(x.r);
    case 'fe': return progFE(x); case 'ua': return progUA(x); default: return ''; } }
function colKeyVal(x,k){ const v = colVal(x,k); return v===''? VACIO : v; }
function valoresProg(k){ const set = new Set();
  for (const x of PROGRAMA) set.add(colKeyVal(x,k));
  let arr = [...set];
  if (k==='mes') arr.sort((a,b)=> MESES.indexOf(a)-MESES.indexOf(b));
  else if (k==='id') arr.sort((a,b)=> a===VACIO?1 : b===VACIO?-1 : ((parseFloat(a)||0)-(parseFloat(b)||0)) || a.localeCompare(b,'es',{numeric:true}));
  else arr.sort((a,b)=> a===VACIO?-1 : b===VACIO?1 : a.localeCompare(b,'es',{numeric:true}));
  return arr; }
function buildProgHead(){ const tr = $('#progHead'); tr.replaceChildren();
  for (const c of PCOLS){ const th = document.createElement('th');
    const inner = document.createElement('div'); inner.className = 'th-inner';
    const lab = document.createElement('span'); lab.className = 'th-label'; lab.style.cursor = 'default'; lab.innerHTML = `<span>${c.t}</span>`;
    const fb = document.createElement('span'); fb.className = 'filtro-btn' + (progFiltros.has(c.k)?' activo':''); fb.dataset.pcol = c.k; fb.textContent = '▾'; fb.title = 'Filtrar ' + c.t;
    fb.onclick = e => { e.stopPropagation(); abrirDropdownProg(c.k, fb); };
    inner.append(lab, fb); th.appendChild(inner); tr.appendChild(th); } }
function abrirDropdownProg(k, btn){
  if (menuAbierto && menuAbierto.dataset.pcol === k){ cerrarMenu(); return; } cerrarMenu();
  const valores = valoresProg(k);
  const seleccion = progFiltros.has(k) ? new Set(progFiltros.get(k)) : new Set(valores);
  const dd = document.createElement('div'); dd.className = 'menu dropdown'; dd.dataset.pcol = k;
  dd.innerHTML = `<div class="dd-top"><input type="text" class="dd-search" placeholder="Buscar valor…"></div>
    <div class="dd-list"></div>
    <div class="dd-acc"><button class="ok">Aplicar</button><button class="cancel">Cancelar</button></div>`;
  document.body.appendChild(dd); menuAbierto = dd;
  const lista = dd.querySelector('.dd-list');
  const dibujar = (f="") => { lista.replaceChildren(); const q = f.trim().toLowerCase();
    const vis = valores.filter(v => !q || v.toLowerCase().includes(q));
    const tl = document.createElement('label'); tl.className = 'todos';
    const tc = document.createElement('input'); tc.type = 'checkbox';
    const mv = vis.filter(v => seleccion.has(v)).length;
    tc.checked = vis.length>0 && mv===vis.length; tc.indeterminate = mv>0 && mv<vis.length;
    tc.onchange = () => { vis.forEach(v => tc.checked ? seleccion.add(v) : seleccion.delete(v)); dibujar(f); };
    const ts = document.createElement('span'); ts.className='dd-val'; ts.textContent='(Seleccionar todo)';
    tl.append(tc, ts); lista.appendChild(tl);
    for (const v of vis){ const l = document.createElement('label');
      const c = document.createElement('input'); c.type='checkbox'; c.checked=seleccion.has(v);
      c.onchange = () => { c.checked?seleccion.add(v):seleccion.delete(v);
        const m = vis.filter(x=>seleccion.has(x)).length; tc.checked=m===vis.length; tc.indeterminate=m>0&&m<vis.length; };
      const s = document.createElement('span'); s.className='dd-val'; s.textContent=v;
      if (v===VACIO){ s.style.color='var(--faint)'; s.style.fontStyle='italic'; }
      l.append(c, s); lista.appendChild(l); } };
  dibujar();
  dd.querySelector('.dd-search').oninput = e => dibujar(e.target.value);
  dd.querySelector('.ok').onclick = () => { if (seleccion.size===valores.length) progFiltros.delete(k); else progFiltros.set(k, seleccion); cerrarMenu(); renderPrograma(); };
  dd.querySelector('.cancel').onclick = cerrarMenu;
  posicionar(dd, btn); dd.querySelector('.dd-search').focus(); }
function getProg(k,m){ return (notas[k] && notas[k].prog && notas[k].prog[m]) || {}; }
function setProg(k,m,fe){ if(!notas[k]) notas[k]={}; if(!notas[k].prog) notas[k].prog={};
  const ua = ahora(); notas[k].prog[m] = {fe: fe||'', ua}; try{ localStorage.setItem(LS_KEY, JSON.stringify(notas)); }catch(e){} marcarDirty(true); return ua; }
function programaVisibles(){ const q = $('#busqueda').value.trim().toLowerCase();
  return PROGRAMA.filter(x => {
    for (const [k,sel] of progFiltros){ if (!sel.has(colKeyVal(x,k))) return false; }
    if (q && !PCOLS.some(c => norm(colVal(x,c.k)).toLowerCase().includes(q))) return false;
    return true; }); }
function renderPrograma(){
  buildProgHead();
  $('#progClear').hidden = progFiltros.size === 0;
  const vis = programaVisibles();
  let real=0, pend=0; for (const x of PROGRAMA){ const e = resultadoEstado(x.r); if (e==='realizado') real++; else if (e==='pendiente') pend++; }
  $('#progInfo').innerHTML = `<b>${PROGRAMA.length.toLocaleString('es')}</b> mantenciones · <b>${real.toLocaleString('es')}</b> realizadas · <b>${pend.toLocaleString('es')}</b> pendientes · <b>${vis.length.toLocaleString('es')}</b> en pantalla`;
  const body = $('#progBody');
  if (!vis.length){ body.innerHTML = '<tr><td colspan="14" class="prog-empty">Sin mantenciones que coincidan con los filtros.</td></tr>'; return; }
  const frag = document.createDocumentFragment();
  for (const x of vis){
    const fe = progFE(x), ua = progUA(x);
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${norm(x.id)||'—'}</td><td>${norm(x.car)||'—'}</td><td>${norm(x.inv)||'—'}</td>
      <td class="eqn">${norm(x.eq)||'—'}</td><td>${norm(x.sv)||'—'}</td><td>${norm(x.ub)||'—'}</td>
      <td>${norm(x.ma)||'—'}</td><td>${norm(x.mo)||'—'}</td><td>${norm(x.se)||'—'}</td>
      <td>${MESES[x.m-1]}</td><td>${etiquetaPrograma(x.p)}</td><td>${etiquetaResultado(x.r)}</td>
      <td><input type="date" class="fe" data-k="${(x.k||'').replace(/"/g,'&quot;')}" data-m="${x.m}" value="${fe}"></td>
      <td class="ua">${fe ? ua : '—'}</td>`;
    frag.appendChild(tr); }
  body.replaceChildren(frag);
}

/* ---------- Panel Resumen ---------- */
function resumenServicios(){ const m = new Map();
  for (let i=0;i<ROWS.length;i++){ const s = claveFiltro(getCell(i,4));
    const o = m.get(s) || {t:0,p:0}; o.t++; const e = estadoDe(i); if (e==='Pendiente'||e==='Vencido') o.p++; m.set(s,o); }
  return [...m.entries()].sort((a,b)=>b[1].t-a[1].t); }
function resumenVida(){ let vig=0,pa=0,ag=0,disp=0;
  for (let i=0;i<ROWS.length;i++){ const s = norm(getCell(i,12)), n = parseFloat(s);
    if (s===''||/disponible/i.test(s)||isNaN(n)) disp++; else if (n<0) ag++; else if (n<=2) pa++; else vig++; }
  return {vig,pa,ag,disp}; }
function resumenEstado(){ let ald=0,pen=0,ven=0,sin=0;
  for (let i=0;i<ROWS.length;i++){ const e = estadoDe(i);
    if (e==='Vencido') ven++; else if (e==='Pendiente') pen++; else if (e==='Al día') ald++; else sin++; }
  return {ald,pen,ven,sin}; }
function pc(n, tot){ return tot ? (n/tot*100) : 0; }
function renderDashboard(){
  const tot = ROWS.length || 1;
  const serv = resumenServicios(), maxT = serv.length ? serv[0][1].t : 1;
  const filasServ = serv.slice(0,12).map(([nm,o]) => {
    const esc = nm.replace(/"/g,'&quot;');
    return `<div class="bar-row" data-serv="${esc}"><span class="nm" title="${esc}">${nm}</span>
      <span class="bar-track"><span class="bar-fill" style="width:${pc(o.t,maxT)}%"></span><span class="bar-pend" style="width:${pc(o.p,maxT)}%"></span></span>
      <span class="ct">${o.t}${o.p?` · ${o.p}⏳`:''}</span></div>`; }).join('');
  const v = resumenVida();
  const segV = `<div class="seg">
    <span style="width:${pc(v.vig,tot)}%;background:var(--green)" title="Vigente: ${v.vig}"></span>
    <span style="width:${pc(v.pa,tot)}%;background:var(--amber)" title="Por agotar: ${v.pa}"></span>
    <span style="width:${pc(v.ag,tot)}%;background:var(--red)" title="Agotada: ${v.ag}"></span>
    <span style="width:${pc(v.disp,tot)}%;background:var(--faint)" title="Disponible/sin dato: ${v.disp}"></span></div>
    <div class="legend">
      <span class="li" data-vida="vig"><span class="sw" style="background:var(--green)"></span>Vigente (&gt;2 años)<span class="lc">${v.vig}</span></span>
      <span class="li" data-vida="pa"><span class="sw" style="background:var(--amber)"></span>Por agotar (≤2)<span class="lc">${v.pa}</span></span>
      <span class="li" data-vida="ag"><span class="sw" style="background:var(--red)"></span>Agotada<span class="lc">${v.ag}</span></span>
      <span class="li"><span class="sw" style="background:var(--faint)"></span>Disponible / sin dato<span class="lc">${v.disp}</span></span></div>`;
  const e = resumenEstado();
  const segE = `<div class="seg">
    <span style="width:${pc(e.ald,tot)}%;background:var(--green)" title="Al día: ${e.ald}"></span>
    <span style="width:${pc(e.pen,tot)}%;background:var(--amber)" title="Pendiente: ${e.pen}"></span>
    <span style="width:${pc(e.ven,tot)}%;background:var(--red)" title="Vencido: ${e.ven}"></span>
    <span style="width:${pc(e.sin,tot)}%;background:var(--faint)" title="Sin registros: ${e.sin}"></span></div>
    <div class="legend">
      <span class="li" data-est="Al día"><span class="sw" style="background:var(--green)"></span>Al día<span class="lc">${e.ald}</span></span>
      <span class="li" data-est="Pendiente"><span class="sw" style="background:var(--amber)"></span>Pendiente<span class="lc">${e.pen}</span></span>
      <span class="li" data-est="Vencido"><span class="sw" style="background:var(--red)"></span>Vencido<span class="lc">${e.ven}</span></span>
      <span class="li"><span class="sw" style="background:var(--faint)"></span>Sin registros<span class="lc">${e.sin}</span></span></div>`;
  $('#dashboard').innerHTML = `<div class="dash-grid">
    <div class="card-d"><h3>Equipos por servicio <small>top 12 · ⏳ con pendientes</small></h3>${filasServ}</div>
    <div class="card-d"><h3>Vida útil</h3>${segV}</div>
    <div class="card-d"><h3>Estado de mantención</h3>${segE}</div></div>`;
  $('#dashboard').querySelectorAll('.bar-row').forEach(r => r.onclick = () => { filtros.set(4, new Set([r.dataset.serv])); render(); });
  $('#dashboard').querySelectorAll('.li[data-vida="ag"]').forEach(el => el.onclick = () => { filtroVida = true; render(); });
  $('#dashboard').querySelectorAll('.li[data-est]').forEach(el => el.onclick = () => { filtros.set(IDX_ESTADO, new Set([el.dataset.est])); render(); });
}

function construirEncabezado(){ const tr = $('#encabezado'); tr.replaceChildren();
  HEADERS.forEach((h, ci) => { if (!visible(ci)) return;
    const th = document.createElement('th');
    if (FREEZE[ci]) th.classList.add('freeze', FREEZE[ci]);
    if (EDIT_COLS.has(ci)) th.classList.add('edit'); if (DERIV.has(ci)) th.classList.add('virt');
    const inner = document.createElement('div'); inner.className = 'th-inner';
    const lab = document.createElement('span'); lab.className = 'th-label';
    const lapiz = EDIT_COLS.has(ci) ? ' <span class="lapiz" title="editable">✎</span>' : '';
    lab.innerHTML = `<span>${h}${lapiz}</span> <span class="sort-ind" data-col="${ci}"></span>`;
    lab.title = 'Ordenar por ' + h;
    lab.onclick = () => { if (ordenCol === ci) ordenDir = -ordenDir; else { ordenCol = ci; ordenDir = 1; } render(); };
    const fb = document.createElement('span'); fb.className = 'filtro-btn'; fb.dataset.col = ci;
    fb.textContent = '▾'; fb.title = 'Filtrar ' + h;
    fb.onclick = e => { e.stopPropagation(); abrirDropdown(ci, fb); };
    inner.append(lab, fb); th.appendChild(inner); tr.appendChild(th); });
}

function cerrarMenu(){ if (menuAbierto){ menuAbierto.remove(); menuAbierto = null; } }
function posicionar(el, btn){ const r = btn.getBoundingClientRect(); el.style.top = (r.bottom + 6) + 'px';
  const w = el.offsetWidth || 260; let left = r.left;
  if (left + w > window.innerWidth - 8) left = window.innerWidth - w - 8; el.style.left = Math.max(8, left) + 'px'; }
function abrirDropdown(col, btn){
  if (menuAbierto && menuAbierto.dataset.col == col){ cerrarMenu(); return; } cerrarMenu();
  const valores = valoresDisponibles(col);
  const seleccion = filtros.has(col) ? new Set(filtros.get(col)) : new Set(valores);
  const dd = document.createElement('div'); dd.className = 'menu dropdown'; dd.dataset.col = col;
  dd.innerHTML = `<div class="dd-top"><input type="text" class="dd-search" placeholder="Buscar valor…">
    <div class="dd-ord"><button data-dir="1">▲ A → Z</button><button data-dir="-1">▼ Z → A</button></div></div>
    <div class="dd-list"></div>
    <div class="dd-acc"><button class="ok">Aplicar</button><button class="cancel">Cancelar</button></div>`;
  document.body.appendChild(dd); menuAbierto = dd;
  const lista = dd.querySelector('.dd-list');
  const dibujar = (f="") => { lista.replaceChildren(); const q = f.trim().toLowerCase();
    const vis = valores.filter(v => !q || v.toLowerCase().includes(q));
    const tl = document.createElement('label'); tl.className = 'todos';
    const tc = document.createElement('input'); tc.type = 'checkbox';
    const mv = vis.filter(v => seleccion.has(v)).length;
    tc.checked = vis.length>0 && mv===vis.length; tc.indeterminate = mv>0 && mv<vis.length;
    tc.onchange = () => { vis.forEach(v => tc.checked ? seleccion.add(v) : seleccion.delete(v)); dibujar(f); };
    const ts = document.createElement('span'); ts.className='dd-val'; ts.textContent='(Seleccionar todo)';
    tl.append(tc, ts); lista.appendChild(tl);
    for (const v of vis){ const l = document.createElement('label');
      const c = document.createElement('input'); c.type='checkbox'; c.checked=seleccion.has(v);
      c.onchange = () => { c.checked?seleccion.add(v):seleccion.delete(v);
        const m = vis.filter(x=>seleccion.has(x)).length; tc.checked=m===vis.length; tc.indeterminate=m>0&&m<vis.length; };
      const s = document.createElement('span'); s.className='dd-val'; s.textContent=v;
      if (v===VACIO){ s.style.color='var(--faint)'; s.style.fontStyle='italic'; }
      l.append(c, s); lista.appendChild(l); } };
  dibujar();
  dd.querySelector('.dd-search').oninput = e => dibujar(e.target.value);
  dd.querySelectorAll('.dd-ord button').forEach(b => b.onclick = () => { ordenCol=col; ordenDir=+b.dataset.dir; render(); cerrarMenu(); });
  dd.querySelector('.ok').onclick = () => { if (seleccion.size===valores.length) filtros.delete(col); else filtros.set(col, seleccion); cerrarMenu(); render(); };
  dd.querySelector('.cancel').onclick = cerrarMenu;
  posicionar(dd, btn); dd.querySelector('.dd-search').focus();
}
function abrirColumnas(btn){ if (menuAbierto && menuAbierto.dataset.cols){ cerrarMenu(); return; } cerrarMenu();
  const m = document.createElement('div'); m.className = 'menu colmenu'; m.dataset.cols = '1';
  m.innerHTML = '<div class="ti">Mostrar columnas</div>';
  HEADERS.forEach((h, ci) => { const lock = (ci===0 || ci===3);
    const l = document.createElement('label'); if (lock) l.className = 'lock';
    const c = document.createElement('input'); c.type='checkbox'; c.checked=visible(ci); c.disabled=lock;
    c.onchange = () => { if (c.checked) oculto.delete(ci); else oculto.add(ci);
      try{ localStorage.setItem(LS_COLS, JSON.stringify([...oculto])); }catch(e){}
      construirEncabezado(); render(); };
    const s = document.createElement('span'); s.textContent = h + (lock?' (fija)':'');
    l.append(c, s); m.appendChild(l); });
  document.body.appendChild(m); menuAbierto = m; posicionar(m, btn); }

/* ---------- Drawer ---------- */
function abrirDrawer(i){ drawerRow = i;
  $('#dEyebrow').textContent = `ID ${norm(getCell(i,0))} · ${norm(getCell(i,4)) || 'Sin servicio'}`;
  $('#dTitulo').textContent = norm(getCell(i,3)) || 'Equipo sin nombre';
  const tags = [];
  if (norm(getCell(i,8))||norm(getCell(i,9))) tags.push(`${norm(getCell(i,8))} ${norm(getCell(i,9))}`.trim());
  if (norm(getCell(i,10))) tags.push('Serie ' + norm(getCell(i,10)));
  if (norm(getCell(i,2))) tags.push('Inv. ' + norm(getCell(i,2)));
  $('#dTags').innerHTML = tags.map(t => `<span class="tag">${t}</span>`).join('');
  pintarRegistros(i);
  const dl = $('#dDatos'); dl.replaceChildren();
  for (let ci = 0; ci < NB; ci++){ const dt = document.createElement('dt'); dt.textContent = HEADERS[ci];
    const dd = document.createElement('dd'); dd.textContent = norm(getCell(i,ci)) || '—'; dl.append(dt, dd); }
  const cont = $('#dCampos'); cont.replaceChildren();
  [IDX_OBS, IDX_NOTAS].forEach(ci => { const wrap = document.createElement('div'); wrap.className = 'field';
    const lab = document.createElement('label');
    lab.innerHTML = `<span>${HEADERS[ci]}</span><span class="saved" id="sv_${ci}">✓ guardado</span>`;
    const ta = document.createElement('textarea'); ta.id = 'ta_'+ci; ta.value = norm(getCell(i,ci));
    ta.placeholder = ci===IDX_OBS ? 'Estado o condición técnica del equipo…' : 'Comentario libre o administrativo…';
    ta.addEventListener('input', () => { setCell(i, ci, ta.value);
      const cell = document.querySelector(`td.editable[data-i="${i}"][data-ci="${ci}"]`);
      if (cell && cell.textContent !== ta.value) cell.textContent = ta.value;
      const sv = $('#sv_'+ci); sv.classList.add('on'); clearTimeout(sv._t); sv._t = setTimeout(()=>sv.classList.remove('on'), 1200); });
    wrap.append(lab, ta); cont.appendChild(wrap); });
  $('#scrim').classList.add('on'); $('#drawer').classList.add('on');
}
function ordenRegistros(a, b){ const ra=a.r, rb=b.r;
  if (!!ra.p !== !!rb.p) return ra.p ? -1 : 1;
  if (ra.p){ const va=ra.v||'9999-99-99', vb=rb.v||'9999-99-99'; if (va!==vb) return va<vb?-1:1; return (rb.f||'').localeCompare(ra.f||''); }
  return (rb.f||'').localeCompare(ra.f||''); }
function pintarRegistros(i){
  const cont = $('#dRegistros');
  cont.innerHTML = `<div class="reg-add">
      <input type="date" id="regFecha">
      <label class="chk"><input type="checkbox" id="regPend"> Pendiente</label>
      <input type="text" id="regDesc" class="desc" placeholder="Descripción de la intervención…">
      <input type="date" id="regVence" class="desc" title="Vencimiento / recordatorio (opcional)">
      <div class="lblv">Fecha · ☐ pendiente · descripción · y vencimiento (opcional, para recordatorio)</div>
      <button id="regAdd">+ Agregar intervención</button>
    </div><ul class="reg-list" id="regList"></ul>`;
  $('#regFecha').value = HOY;
  const arr = registrosDe(i);
  const ul = $('#regList');
  if (!arr.length){ ul.innerHTML = '<li class="reg-empty">Sin intervenciones registradas todavía.</li>'; }
  else { const orden = arr.map((r, idx) => ({r, idx})).sort(ordenRegistros);
    for (const {r, idx} of orden){ const li = document.createElement('li');
      const cv = claseVenc(r); li.className = 'reg-item' + (r.p ? (cv==='venc'?' venc':' pend') : '');
      const due = r.v ? `<span class="due ${cv}">vence ${r.v}</span>` : '';
      li.innerHTML = `<span class="rf">${r.f||'—'}</span><span class="rd"></span>
        <span class="actions">${due}<button class="pill ${r.p?'pend':'ok'}">${r.p?'Pendiente':'Listo'}</button>
        <button class="del" title="Eliminar">✕</button></span>`;
      li.querySelector('.rd').textContent = r.d || '';
      li.querySelector('.pill').onclick = () => { const a = regsActuales(i); a[idx].p = !a[idx].p; guardar(); pintarRegistros(i); render(); };
      li.querySelector('.del').onclick = () => { const a = regsActuales(i); a.splice(idx,1); guardar(); pintarRegistros(i); render(); };
      ul.appendChild(li); } }
  $('#regAdd').onclick = () => { const d = $('#regDesc').value.trim(); if (!d){ $('#regDesc').focus(); return; }
    const reg = { f: $('#regFecha').value || HOY, d, p: $('#regPend').checked };
    const v = $('#regVence').value; if (v) reg.v = v;
    const a = regsActuales(i); a.push(reg); guardar(); pintarRegistros(i); render(); toast('Intervención agregada'); };
}
function cerrarDrawer(){ $('#scrim').classList.remove('on'); $('#drawer').classList.remove('on'); drawerRow = null; }

function toast(msg, ok=true){ const t = document.createElement('div'); t.className = 'toast';
  t.innerHTML = (ok?'<span class="ok">✓</span> ':'') + msg; $('#toasts').appendChild(t);
  setTimeout(() => { t.style.opacity='0'; t.style.transition='.3s'; setTimeout(()=>t.remove(),300); }, 2200); }
let _tg; function tostarGuardado(){ clearTimeout(_tg); _tg = setTimeout(()=>toast('Guardado en el navegador'), 700); }

function toggleEstado(vals, btn){ const cur = filtros.get(IDX_ESTADO);
  const igual = cur && cur.size===vals.length && vals.every(v => cur.has(v));
  if (igual) filtros.delete(IDX_ESTADO); else filtros.set(IDX_ESTADO, new Set(vals)); render(); }
function estadoFiltrado(vals){ const c = filtros.get(IDX_ESTADO); return !!(c && c.size===vals.length && vals.every(v => c.has(v))); }
function toggleVida(){ filtroVida = !filtroVida; render(); }
document.addEventListener('mousedown', e => { if (menuAbierto && !menuAbierto.contains(e.target) && !e.target.classList.contains('filtro-btn') && e.target.id!=='btnCols') cerrarMenu(); });
document.addEventListener('keydown', e => {
  if (e.key === 'Escape'){ cerrarMenu(); cerrarDrawer(); cerrarAyuda(); return; }
  const t = e.target, tag = (t.tagName||'').toLowerCase();
  if (tag==='input' || tag==='textarea' || t.isContentEditable) return;
  if (e.key === '?' || (e.key === '/' && e.shiftKey)){ e.preventDefault(); toggleAyuda(); }
  else if (e.key === '/'){ e.preventDefault(); $('#busqueda').focus(); }
  else if (e.key === 't'){ toggleVista(); }
  else if (e.key === 'd'){ toggleDash(); }
  else if (e.key === 'p'){ togglePrograma(); }
  else if (e.key === 'o'){ toggleTema(); }
});
$('.table-wrap').addEventListener('scroll', cerrarMenu);
window.addEventListener('resize', cerrarMenu);
$('#cuerpo').addEventListener('click', e => { if (e.target.closest('.editable')) return;
  const tr = e.target.closest('tr'); if (tr) abrirDrawer(+tr.dataset.i); });
$('#cards').addEventListener('click', e => { const c = e.target.closest('.eqcard'); if (c) abrirDrawer(+c.dataset.i); });
$('#btnTema').addEventListener('click', toggleTema);
$('#btnVista').addEventListener('click', toggleVista);
$('#btnDash').addEventListener('click', toggleDash);
$('#btnPrograma').addEventListener('click', togglePrograma);
$('#progBack').addEventListener('click', () => { vista='tabla'; try{ localStorage.setItem(LS_VIEW, vista); }catch(e){} render(); });
$('#progClear').addEventListener('click', () => { progFiltros.clear(); renderPrograma(); });
$('.prog-table-wrap').addEventListener('scroll', cerrarMenu);
$('#progBody').addEventListener('change', e => { const inp = e.target.closest('input.fe'); if (!inp) return;
  const ua = setProg(inp.dataset.k, +inp.dataset.m, inp.value);
  const cell = inp.closest('tr').querySelector('.ua'); if (cell) cell.textContent = inp.value ? ua : '—'; });
$('#btnAyuda').addEventListener('click', toggleAyuda);
$('#help').addEventListener('click', e => { if (e.target.id === 'help') cerrarAyuda(); });
$('#scrim').addEventListener('click', cerrarDrawer);
$('#cerrarDrawer').addEventListener('click', cerrarDrawer);
$('#dCerrar2').addEventListener('click', cerrarDrawer);
$('#busqueda').addEventListener('input', render);
$('#btnCols').addEventListener('click', e => { e.stopPropagation(); abrirColumnas(e.currentTarget); });
$('#btnDensidad').addEventListener('click', () => document.body.classList.toggle('compact'));
$('#limpiar').addEventListener('click', () => { filtros.clear(); progFiltros.clear(); filtroVida=false; $('#busqueda').value=''; ordenCol=null; render(); });

function descargarNotas(){ const blob = new Blob([JSON.stringify({version:2, notas}, null, 2)], {type:'application/json'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'notas_equipos.json';
  a.click(); URL.revokeObjectURL(a.href); marcarDirty(false); toast('Respaldo descargado: notas_equipos.json'); }
$('#guardar').addEventListener('click', descargarNotas);
$('#dGuardar').addEventListener('click', descargarNotas);
$('#cargar').addEventListener('click', () => $('#archivo').click());
$('#archivo').addEventListener('change', ev => { const file = ev.target.files[0]; if (!file) return;
  const fr = new FileReader();
  fr.onload = () => { try { const obj = JSON.parse(fr.result); const ent = obj && obj.notas ? obj.notas : obj;
      if (typeof ent !== 'object') throw new Error('formato no válido');
      let n = 0; for (const k in ent){ notas[k] = Object.assign(notas[k]||{}, ent[k]); n++; }
      try { localStorage.setItem(LS_KEY, JSON.stringify(notas)); } catch(e){}
      marcarDirty(false); render(); toast(`Notas cargadas (${n} equipos)`);
    } catch(e){ toast('No se pudo leer el archivo: ' + e.message, false); } };
  fr.readAsText(file); ev.target.value = ''; });
$('#exportar').addEventListener('click', () => { const vis = indicesVisibles(); const esc = s => '"' + norm(s).replace(/"/g,'""') + '"';
  const cols = [...Array(HEADERS.length).keys()].filter(visible);
  const lin = [cols.map(c=>esc(HEADERS[c])).join(',')];
  for (const i of vis) lin.push(cols.map(c=>esc(getCell(i,c))).join(','));
  const blob = new Blob(["﻿"+lin.join("\r\n")], {type:'text/csv;charset=utf-8;'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'equipos_filtrado.csv';
  a.click(); URL.revokeObjectURL(a.href); toast(`CSV exportado (${vis.length} filas)`); });
window.addEventListener('beforeunload', e => { if (dirty){ e.preventDefault(); e.returnValue = ''; } });

aplicarTema();
$('#btnVista').textContent = vista==='tabla' ? '▤ Tarjetas' : '▦ Tabla';
construirEncabezado();
render();
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--db", type=Path, default=DB_POR_DEFECTO)
    ap.add_argument("--salida", type=Path, default=SALIDA_POR_DEFECTO)
    args = ap.parse_args()

    rows, db_reg = leer_datos(args.db)
    programa = leer_programa(args.db)
    headers = [h for _, h in COLUMNAS_ROW] + COLUMNAS_VIRT
    nb = len(COLUMNAS_BASE)
    html = (PLANTILLA
            .replace("__PROGRAMA__", json.dumps(programa, ensure_ascii=False, separators=(",", ":")))
            .replace("__HEADERS__", json.dumps(headers, ensure_ascii=False))
            .replace("__ROWS__", json.dumps(rows, ensure_ascii=False, separators=(",", ":")))
            .replace("__DBREG__", json.dumps(db_reg, ensure_ascii=False, separators=(",", ":")))
            .replace("__NUMERICAS__", json.dumps(NUMERICAS))
            .replace("__NB__", str(nb))
            .replace("__IDX_OBS__", str(nb))
            .replace("__IDX_NOTAS__", str(nb + 1))
            .replace("__IDX_ESTADO__", str(nb + 2))
            .replace("__IDX_VENCE__", str(nb + 3)))
    args.salida.write_text(html, encoding="utf-8")
    print(f"Generado: {args.salida}  ({len(rows)} equipos, {len(programa)} mantenciones, {len(html)//1024} KB)")


if __name__ == "__main__":
    main()
