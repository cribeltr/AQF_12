#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera un archivo HTML autónomo (sin servidor) con una tabla de los equipos
en la que CADA COLUMNA es filtrable al estilo Excel: un desplegable con
casillas para seleccionar uno o más valores, además de búsqueda global,
ordenamiento y exportación a CSV.

Los datos se incrustan en el propio HTML leyéndolos de equipos.db, por lo que
los números de serie/inventario conservan sus ceros a la izquierda (texto).

Uso:
    python scripts/generar_html.py            # usa equipos.db -> equipos_filtrable.html
    python scripts/generar_html.py --db x.db --salida tabla.html
"""
from __future__ import annotations

import argparse
import json
import sqlite3
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
DB_POR_DEFECTO = RAIZ / "equipos.db"
SALIDA_POR_DEFECTO = RAIZ / "equipos_filtrable.html"

# Orden y encabezados EXACTOS de la planilla.
COLUMNAS = [
    ("id", "ID"),
    ("n_carpeta", "N° Carpeta"),
    ("n_inventario", "N° Inventario"),
    ("equipo", "Equipo"),
    ("servicio", "Servicio"),
    ("unidad", "Unidad"),
    ("ubicacion", "Ubicación"),
    ("procedencia", "Procedencia"),
    ("marca", "Marca"),
    ("modelo", "Modelo"),
    ("serie", "Serie"),
    ("anio_instalacion", "Año Instalación"),
    ("vida_util_residual", "Vida Útil Residual"),
    ("clasificacion", "Clasificación"),
    ("enu_baja", "ENU / Baja"),
]
# Índices de columnas numéricas (para ordenar numéricamente).
NUMERICAS = [0, 1, 11]


def leer_datos(db: Path):
    con = sqlite3.connect(db)
    cols = ", ".join(c for c, _ in COLUMNAS)
    filas = con.execute(f"SELECT {cols} FROM equipos ORDER BY id").fetchall()
    con.close()
    # None -> "" para el JSON; los textos (serie, inventario) ya vienen tal cual.
    return [[("" if v is None else v) for v in fila] for fila in filas]


PLANTILLA = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Equipos médicos — Programa de Mantenciones 2026</title>
<style>
  :root{ --verde:#217346; --verde2:#1a5c38; --borde:#d6d6d6; --sel:#cfe8d8; }
  *{ box-sizing:border-box; }
  body{ margin:0; font-family:"Segoe UI",Roboto,Arial,sans-serif; font-size:13px; color:#222; }
  header{ background:var(--verde); color:#fff; padding:10px 16px; }
  header h1{ margin:0; font-size:16px; font-weight:600; }
  header .sub{ font-size:12px; opacity:.9; margin-top:2px; }
  .toolbar{ position:sticky; top:0; z-index:30; display:flex; gap:8px; align-items:center;
            flex-wrap:wrap; padding:8px 16px; background:#f3f3f3; border-bottom:1px solid var(--borde); }
  .toolbar input[type=search]{ padding:6px 10px; border:1px solid #bbb; border-radius:4px; min-width:240px; }
  .toolbar button{ padding:6px 12px; border:1px solid #bbb; background:#fff; border-radius:4px; cursor:pointer; }
  .toolbar button:hover{ background:#eaeaea; }
  .toolbar .conteo{ margin-left:auto; font-weight:600; color:#333; }
  .table-wrap{ overflow:auto; max-height:calc(100vh - 120px); }
  table{ border-collapse:collapse; width:100%; }
  thead th{ position:sticky; top:0; z-index:10; background:var(--verde); color:#fff;
            padding:0; border:1px solid var(--verde2); white-space:nowrap; }
  .th-inner{ display:flex; align-items:center; gap:6px; padding:6px 8px; }
  .th-label{ cursor:pointer; user-select:none; flex:1; }
  .th-label:hover{ text-decoration:underline; }
  .sort-ind{ font-size:10px; opacity:.85; }
  .filtro-btn{ cursor:pointer; border:1px solid rgba(255,255,255,.5); background:rgba(255,255,255,.12);
               color:#fff; border-radius:3px; font-size:11px; line-height:1; padding:3px 5px; }
  .filtro-btn:hover{ background:rgba(255,255,255,.28); }
  .filtro-btn.activo{ background:#ffd24d; color:#1a1a1a; border-color:#e6b800; }
  tbody td{ border:1px solid var(--borde); padding:4px 8px; white-space:nowrap;
            max-width:340px; overflow:hidden; text-overflow:ellipsis; }
  tbody tr:nth-child(even){ background:#f7faf8; }
  tbody tr:hover{ background:#eef6f0; }
  td.num{ text-align:right; font-variant-numeric:tabular-nums; }
  .vacia{ color:#999; font-style:italic; }
  /* Desplegable de filtro */
  .dropdown{ position:fixed; z-index:1000; width:260px; background:#fff; border:1px solid #9b9b9b;
             border-radius:5px; box-shadow:0 8px 24px rgba(0,0,0,.22); font-size:13px; color:#222; }
  .dropdown .dd-top{ padding:8px; border-bottom:1px solid #eee; }
  .dropdown .dd-search{ width:100%; padding:5px 8px; border:1px solid #bbb; border-radius:4px; }
  .dropdown .dd-ord{ display:flex; gap:6px; margin-top:8px; }
  .dropdown .dd-ord button{ flex:1; padding:4px; border:1px solid #ccc; background:#fafafa; border-radius:4px; cursor:pointer; font-size:12px; }
  .dropdown .dd-ord button:hover{ background:#eee; }
  .dropdown .dd-list{ max-height:240px; overflow:auto; padding:6px 8px; }
  .dropdown label{ display:flex; align-items:center; gap:7px; padding:3px 2px; cursor:pointer; }
  .dropdown label:hover{ background:#f0f0f0; }
  .dropdown label.todos{ border-bottom:1px solid #eee; margin-bottom:4px; padding-bottom:6px; font-weight:600; }
  .dropdown .dd-val{ overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .dropdown .dd-acc{ display:flex; gap:6px; padding:8px; border-top:1px solid #eee; }
  .dropdown .dd-acc button{ flex:1; padding:6px; border-radius:4px; border:1px solid #bbb; cursor:pointer; }
  .dropdown .dd-acc .ok{ background:var(--verde); color:#fff; border-color:var(--verde2); }
  .dropdown .dd-acc .ok:hover{ background:var(--verde2); }
  .dropdown .dd-acc .cancel:hover{ background:#eee; }
</style>
</head>
<body>
<header>
  <h1>Equipos médicos — Programa de Mantenciones Preventivas 2026</h1>
  <div class="sub">Filtra cada columna como en Excel: haz clic en el embudo ▾ del encabezado y marca uno o más valores.</div>
</header>

<div class="toolbar">
  <input type="search" id="busqueda" placeholder="Buscar en todas las columnas…">
  <button id="limpiar">Limpiar filtros</button>
  <button id="exportar">Exportar CSV (vista actual)</button>
  <span class="conteo" id="conteo"></span>
</div>

<div class="table-wrap">
  <table>
    <thead><tr id="encabezado"></tr></thead>
    <tbody id="cuerpo"></tbody>
  </table>
</div>

<script>
const HEADERS = __HEADERS__;
const ROWS = __ROWS__;
const NUMERICAS = new Set(__NUMERICAS__);
const VACIO = "(Vacías)";

const filtros = new Map();      // colIndex -> Set(valores seleccionados)
let ordenCol = null, ordenDir = 1;
let ddAbierto = null;           // elemento dropdown activo

const norm = v => (v === null || v === undefined) ? "" : String(v);
const clave = v => { const s = norm(v).trim(); return s === "" ? VACIO : s; };

function comparar(col, a, b){
  if (NUMERICAS.has(col)){
    const na = parseFloat(a), nb = parseFloat(b);
    const va = isNaN(na) ? Infinity : na, vb = isNaN(nb) ? Infinity : nb;
    return va - vb;
  }
  return String(a).localeCompare(String(b), 'es', {numeric:true, sensitivity:'base'});
}

function pasaExcepto(fila, excepto){
  for (const [col, set] of filtros){
    if (col === excepto) continue;
    if (!set.has(clave(fila[col]))) return false;
  }
  return true;
}
function coincideBusqueda(fila){
  const q = document.getElementById('busqueda').value.trim().toLowerCase();
  if (!q) return true;
  return fila.some(c => norm(c).toLowerCase().includes(q));
}
function filasVisibles(){
  let v = ROWS.filter(f => pasaExcepto(f, -1) && coincideBusqueda(f));
  if (ordenCol !== null){
    v = v.slice().sort((x, y) => ordenDir * comparar(ordenCol, x[ordenCol], y[ordenCol]));
  }
  return v;
}

function valoresDisponibles(col){
  // Valores distintos considerando los filtros de las OTRAS columnas (como Excel).
  const set = new Set();
  for (const f of ROWS){ if (pasaExcepto(f, col) && coincideBusqueda(f)) set.add(clave(f[col])); }
  const arr = [...set];
  arr.sort((a, b) => {
    if (a === VACIO) return 1; if (b === VACIO) return -1;
    return comparar(col, a, b);
  });
  return arr;
}

function render(){
  const cuerpo = document.getElementById('cuerpo');
  const vis = filasVisibles();
  const frag = document.createDocumentFragment();
  for (const fila of vis){
    const tr = document.createElement('tr');
    fila.forEach((val, ci) => {
      const td = document.createElement('td');
      const s = norm(val);
      if (NUMERICAS.has(ci)) td.className = 'num';
      if (s === ''){ td.textContent = ''; }
      else td.textContent = s;
      tr.appendChild(td);
    });
    frag.appendChild(tr);
  }
  cuerpo.replaceChildren(frag);
  document.getElementById('conteo').textContent =
    `${vis.length.toLocaleString('es')} de ${ROWS.length.toLocaleString('es')} equipos`;
  // marcar encabezados con filtro activo / indicador de orden
  document.querySelectorAll('.filtro-btn').forEach(btn => {
    const c = +btn.dataset.col;
    btn.classList.toggle('activo', filtros.has(c));
  });
  document.querySelectorAll('.sort-ind').forEach(el => {
    const c = +el.dataset.col;
    el.textContent = (ordenCol === c) ? (ordenDir === 1 ? '▲' : '▼') : '';
  });
}

function construirEncabezado(){
  const tr = document.getElementById('encabezado');
  HEADERS.forEach((h, ci) => {
    const th = document.createElement('th');
    const inner = document.createElement('div'); inner.className = 'th-inner';
    const lab = document.createElement('span'); lab.className = 'th-label';
    lab.innerHTML = `${h} <span class="sort-ind" data-col="${ci}"></span>`;
    lab.title = 'Ordenar por ' + h;
    lab.onclick = () => { if (ordenCol === ci) ordenDir = -ordenDir; else { ordenCol = ci; ordenDir = 1; } render(); };
    const fb = document.createElement('span'); fb.className = 'filtro-btn'; fb.dataset.col = ci;
    fb.textContent = '▾'; fb.title = 'Filtrar ' + h;
    fb.onclick = (e) => { e.stopPropagation(); abrirDropdown(ci, fb); };
    inner.appendChild(lab); inner.appendChild(fb);
    th.appendChild(inner); tr.appendChild(th);
  });
}

function cerrarDropdown(){ if (ddAbierto){ ddAbierto.remove(); ddAbierto = null; } }

function abrirDropdown(col, btn){
  if (ddAbierto && ddAbierto.dataset.col == col){ cerrarDropdown(); return; }
  cerrarDropdown();
  const valores = valoresDisponibles(col);
  const seleccion = filtros.has(col) ? filtros.get(col) : new Set(valores); // sin filtro = todos

  const dd = document.createElement('div'); dd.className = 'dropdown'; dd.dataset.col = col;
  dd.innerHTML = `
    <div class="dd-top">
      <input type="text" class="dd-search" placeholder="Buscar valor…">
      <div class="dd-ord">
        <button data-dir="1">▲ Ordenar A→Z</button>
        <button data-dir="-1">▼ Ordenar Z→A</button>
      </div>
    </div>
    <div class="dd-list"></div>
    <div class="dd-acc">
      <button class="ok">Aceptar</button>
      <button class="cancel">Cancelar</button>
    </div>`;
  document.body.appendChild(dd);
  ddAbierto = dd;

  const lista = dd.querySelector('.dd-list');
  const dibujarLista = (filtro="") => {
    lista.replaceChildren();
    const f = filtro.trim().toLowerCase();
    const visibles = valores.filter(v => !f || v.toLowerCase().includes(f));
    // (Seleccionar todo)
    const todosL = document.createElement('label'); todosL.className = 'todos';
    const todosC = document.createElement('input'); todosC.type = 'checkbox';
    const marcadosVis = visibles.filter(v => seleccion.has(v)).length;
    todosC.checked = visibles.length > 0 && marcadosVis === visibles.length;
    todosC.indeterminate = marcadosVis > 0 && marcadosVis < visibles.length;
    todosC.onchange = () => { visibles.forEach(v => todosC.checked ? seleccion.add(v) : seleccion.delete(v)); dibujarLista(filtro); };
    const todosS = document.createElement('span'); todosS.className = 'dd-val'; todosS.textContent = '(Seleccionar todo)';
    todosL.appendChild(todosC); todosL.appendChild(todosS); lista.appendChild(todosL);
    // valores
    for (const v of visibles){
      const l = document.createElement('label');
      const c = document.createElement('input'); c.type = 'checkbox'; c.checked = seleccion.has(v);
      c.onchange = () => { c.checked ? seleccion.add(v) : seleccion.delete(v);
        // recalcular estado del "todos" sin redibujar toda la lista
        const mv = visibles.filter(x => seleccion.has(x)).length;
        todosC.checked = mv === visibles.length; todosC.indeterminate = mv > 0 && mv < visibles.length; };
      const s = document.createElement('span'); s.className = 'dd-val';
      s.textContent = v; if (v === VACIO) s.classList.add('vacia');
      l.appendChild(c); l.appendChild(s); lista.appendChild(l);
    }
  };
  dibujarLista();

  dd.querySelector('.dd-search').oninput = (e) => dibujarLista(e.target.value);
  dd.querySelectorAll('.dd-ord button').forEach(b => b.onclick = () => {
    ordenCol = col; ordenDir = +b.dataset.dir; render(); cerrarDropdown();
  });
  dd.querySelector('.ok').onclick = () => {
    if (seleccion.size === valores.length) filtros.delete(col);     // todo marcado = sin filtro
    else filtros.set(col, new Set(seleccion));
    cerrarDropdown(); render();
  };
  dd.querySelector('.cancel').onclick = () => cerrarDropdown();

  // posicionar
  const r = btn.getBoundingClientRect();
  dd.style.top = (r.bottom + 4) + 'px';
  let left = r.left; const w = 260;
  if (left + w > window.innerWidth - 8) left = window.innerWidth - w - 8;
  dd.style.left = Math.max(8, left) + 'px';
  dd.querySelector('.dd-search').focus();
}

// cerrar al hacer clic fuera / Esc / scroll
document.addEventListener('mousedown', (e) => { if (ddAbierto && !ddAbierto.contains(e.target) && !e.target.classList.contains('filtro-btn')) cerrarDropdown(); });
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') cerrarDropdown(); });
document.querySelector('.table-wrap').addEventListener('scroll', cerrarDropdown);
window.addEventListener('resize', cerrarDropdown);

document.getElementById('busqueda').addEventListener('input', () => render());
document.getElementById('limpiar').addEventListener('click', () => {
  filtros.clear(); document.getElementById('busqueda').value = ''; ordenCol = null; render();
});
document.getElementById('exportar').addEventListener('click', () => {
  const vis = filasVisibles();
  const esc = s => '"' + norm(s).replace(/"/g, '""') + '"';
  const lineas = [HEADERS.map(esc).join(',')];
  for (const f of vis) lineas.push(f.map(esc).join(','));
  const blob = new Blob(["﻿" + lineas.join("\r\n")], {type:'text/csv;charset=utf-8;'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = 'equipos_filtrado.csv';
  a.click(); URL.revokeObjectURL(a.href);
});

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

    datos = leer_datos(args.db)
    headers = [h for _, h in COLUMNAS]
    html = (PLANTILLA
            .replace("__HEADERS__", json.dumps(headers, ensure_ascii=False))
            .replace("__ROWS__", json.dumps(datos, ensure_ascii=False, separators=(",", ":")))
            .replace("__NUMERICAS__", json.dumps(NUMERICAS)))
    args.salida.write_text(html, encoding="utf-8")
    print(f"Generado: {args.salida}  ({len(datos)} equipos, {len(html)//1024} KB)")


if __name__ == "__main__":
    main()
