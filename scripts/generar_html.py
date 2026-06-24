#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera un archivo HTML autónomo (sin servidor) con una tabla de los equipos
en la que CADA COLUMNA es filtrable al estilo Excel (desplegable con casillas
para seleccionar uno o más valores) y donde las columnas **Observaciones** y
**Notas** son EDITABLES: el usuario puede escribir en ellas directamente.

Persistencia de las notas (sin servidor):
  * Se guardan automáticamente en el navegador (localStorage).
  * Botón "Guardar notas" -> descarga un archivo notas_equipos.json (respaldo
    portable, sirve para llevarlas a otro equipo o devolverlas a la base).
  * Botón "Cargar notas" -> importa un notas_equipos.json previo.
  * scripts/aplicar_notas.py vuelca ese JSON a equipos.db (fuente de verdad).

Los datos se incrustan leyéndolos de equipos.db; los números de serie/inventario
conservan sus ceros a la izquierda (texto).

Uso:
    python scripts/generar_html.py            # equipos.db -> equipos_filtrable.html
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

# Columnas de SOLO LECTURA (orden y encabezados exactos de la planilla).
COLUMNAS_BASE = [
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
# Columnas EDITABLES (se pueden escribir desde la tabla).
COLUMNAS_EDIT = [
    ("observaciones", "Observaciones"),
    ("notas", "Notas"),
]
COLUMNAS = COLUMNAS_BASE + COLUMNAS_EDIT
# Índices de columnas numéricas (para ordenar numéricamente).
NUMERICAS = [0, 1, 11]


def leer_datos(db: Path):
    con = sqlite3.connect(db)
    cols = ", ".join(c for c, _ in COLUMNAS)
    filas = con.execute(f"SELECT {cols} FROM equipos ORDER BY id").fetchall()
    con.close()
    return [[("" if v is None else v) for v in fila] for fila in filas]


PLANTILLA = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Equipos médicos — Programa de Mantenciones 2026</title>
<style>
  :root{ --verde:#217346; --verde2:#1a5c38; --borde:#d6d6d6; --edit:#fffbe6; }
  *{ box-sizing:border-box; }
  body{ margin:0; font-family:"Segoe UI",Roboto,Arial,sans-serif; font-size:13px; color:#222; }
  header{ background:var(--verde); color:#fff; padding:10px 16px; }
  header h1{ margin:0; font-size:16px; font-weight:600; }
  header .sub{ font-size:12px; opacity:.92; margin-top:2px; }
  .toolbar{ position:sticky; top:0; z-index:30; display:flex; gap:8px; align-items:center;
            flex-wrap:wrap; padding:8px 16px; background:#f3f3f3; border-bottom:1px solid var(--borde); }
  .toolbar input[type=search]{ padding:6px 10px; border:1px solid #bbb; border-radius:4px; min-width:220px; }
  .toolbar button{ padding:6px 12px; border:1px solid #bbb; background:#fff; border-radius:4px; cursor:pointer; }
  .toolbar button:hover{ background:#eaeaea; }
  .toolbar button.primario{ background:var(--verde); color:#fff; border-color:var(--verde2); }
  .toolbar button.primario:hover{ background:var(--verde2); }
  .toolbar .conteo{ margin-left:auto; font-weight:600; color:#333; }
  .dirty{ color:#b00; font-weight:600; display:none; }
  .dirty.on{ display:inline; }
  .table-wrap{ overflow:auto; max-height:calc(100vh - 130px); }
  table{ border-collapse:collapse; width:100%; }
  thead th{ position:sticky; top:0; z-index:10; background:var(--verde); color:#fff;
            padding:0; border:1px solid var(--verde2); white-space:nowrap; }
  thead th.edit{ background:#0f6b46; }
  .th-inner{ display:flex; align-items:center; gap:6px; padding:6px 8px; }
  .th-label{ cursor:pointer; user-select:none; flex:1; }
  .th-label:hover{ text-decoration:underline; }
  .sort-ind{ font-size:10px; opacity:.85; }
  .lapiz{ font-size:11px; opacity:.85; }
  .filtro-btn{ cursor:pointer; border:1px solid rgba(255,255,255,.5); background:rgba(255,255,255,.12);
               color:#fff; border-radius:3px; font-size:11px; line-height:1; padding:3px 5px; }
  .filtro-btn:hover{ background:rgba(255,255,255,.28); }
  .filtro-btn.activo{ background:#ffd24d; color:#1a1a1a; border-color:#e6b800; }
  tbody td{ border:1px solid var(--borde); padding:4px 8px; white-space:nowrap;
            max-width:340px; overflow:hidden; text-overflow:ellipsis; vertical-align:top; }
  tbody tr:nth-child(even){ background:#f7faf8; }
  tbody tr:hover{ background:#eef6f0; }
  td.num{ text-align:right; font-variant-numeric:tabular-nums; }
  td.editable{ background:var(--edit); white-space:pre-wrap; min-width:200px; max-width:420px;
               overflow:visible; text-overflow:clip; cursor:text; }
  td.editable:focus{ outline:2px solid var(--verde); background:#fff; }
  td.editable:empty::before{ content:"✎ escribir…"; color:#b9a93a; }
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
  <div class="sub">Filtra cada columna con el embudo ▾. Escribe en <b>Observaciones</b> y <b>Notas</b> (se guardan solas); usa <b>Guardar notas</b> para respaldarlas en un archivo.</div>
</header>

<div class="toolbar">
  <input type="search" id="busqueda" placeholder="Buscar en todas las columnas…">
  <button id="limpiar">Limpiar filtros</button>
  <button id="guardar" class="primario">💾 Guardar notas</button>
  <button id="cargar">📂 Cargar notas</button>
  <button id="exportar">Exportar CSV</button>
  <input type="file" id="archivo" accept=".json,application/json" style="display:none">
  <span class="dirty" id="dirty">● notas sin respaldar en archivo</span>
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
const NB = __NB__;                 // n.º de columnas de solo lectura
const EDIT = { __EDIT_OBS__: 'observaciones', __EDIT_NOTAS__: 'notas' };
const EDIT_COLS = new Set(Object.keys(EDIT).map(Number));
const LS_KEY = 'aqf12_notas_v1';
const VACIO = "(Vacías)";

// notas[clave] = {observaciones?:str, notas?:str}  (solo lo que el usuario edita)
let notas = {};
try { notas = JSON.parse(localStorage.getItem(LS_KEY) || '{}') || {}; } catch(e){ notas = {}; }
let dirty = false;
const filtros = new Map();      // colIndex -> Set(valores seleccionados)
let ordenCol = null, ordenDir = 1;
let ddAbierto = null;

const norm = v => (v === null || v === undefined) ? "" : String(v);
const claveFiltro = v => { const s = norm(v).trim(); return s === "" ? VACIO : s; };

function claveEquipo(i){
  const serie = norm(ROWS[i][10]).trim();           // Serie
  if (serie) return serie;
  const inv = norm(ROWS[i][2]).trim();              // N° Inventario
  if (inv) return inv;
  return "#" + norm(ROWS[i][0]).trim();             // #ID
}
function getCell(i, ci){
  if (EDIT_COLS.has(ci)){
    const f = EDIT[ci], k = claveEquipo(i), ov = notas[k];
    if (ov && Object.prototype.hasOwnProperty.call(ov, f)) return ov[f];
    return norm(ROWS[i][ci]);
  }
  return ROWS[i][ci];
}
function setCell(i, ci, valor){
  const f = EDIT[ci], k = claveEquipo(i);
  (notas[k] || (notas[k] = {}))[f] = valor;
  try { localStorage.setItem(LS_KEY, JSON.stringify(notas)); } catch(e){}
  marcarDirty(true);
}
function marcarDirty(v){ dirty = v; document.getElementById('dirty').classList.toggle('on', v); }

function comparar(col, a, b){
  if (NUMERICAS.has(col)){
    const na = parseFloat(a), nb = parseFloat(b);
    const va = isNaN(na) ? Infinity : na, vb = isNaN(nb) ? Infinity : nb;
    return va - vb;
  }
  return String(a).localeCompare(String(b), 'es', {numeric:true, sensitivity:'base'});
}
function pasaExcepto(i, excepto){
  for (const [col, set] of filtros){
    if (col === excepto) continue;
    if (!set.has(claveFiltro(getCell(i, col)))) return false;
  }
  return true;
}
function coincideBusqueda(i){
  const q = document.getElementById('busqueda').value.trim().toLowerCase();
  if (!q) return true;
  for (let c = 0; c < HEADERS.length; c++) if (norm(getCell(i, c)).toLowerCase().includes(q)) return true;
  return false;
}
function indicesVisibles(){
  let v = [];
  for (let i = 0; i < ROWS.length; i++) if (pasaExcepto(i, -1) && coincideBusqueda(i)) v.push(i);
  if (ordenCol !== null) v.sort((x, y) => ordenDir * comparar(ordenCol, getCell(x, ordenCol), getCell(y, ordenCol)));
  return v;
}
function valoresDisponibles(col){
  const set = new Set();
  for (let i = 0; i < ROWS.length; i++) if (pasaExcepto(i, col) && coincideBusqueda(i)) set.add(claveFiltro(getCell(i, col)));
  const arr = [...set];
  arr.sort((a, b) => { if (a === VACIO) return 1; if (b === VACIO) return -1; return comparar(col, a, b); });
  return arr;
}

function render(){
  const cuerpo = document.getElementById('cuerpo');
  const vis = indicesVisibles();
  const frag = document.createDocumentFragment();
  for (const i of vis){
    const tr = document.createElement('tr');
    for (let ci = 0; ci < HEADERS.length; ci++){
      const td = document.createElement('td');
      if (EDIT_COLS.has(ci)){
        td.className = 'editable';
        td.contentEditable = 'true';
        td.spellcheck = false;
        td.dataset.i = i; td.dataset.ci = ci;
        td.textContent = norm(getCell(i, ci));
        td.addEventListener('input', ev => setCell(+ev.target.dataset.i, +ev.target.dataset.ci, ev.target.innerText));
      } else {
        const s = norm(getCell(i, ci));
        if (NUMERICAS.has(ci)) td.className = 'num';
        td.textContent = s;
      }
      tr.appendChild(td);
    }
    frag.appendChild(tr);
  }
  cuerpo.replaceChildren(frag);
  document.getElementById('conteo').textContent =
    `${vis.length.toLocaleString('es')} de ${ROWS.length.toLocaleString('es')} equipos`;
  document.querySelectorAll('.filtro-btn').forEach(b => b.classList.toggle('activo', filtros.has(+b.dataset.col)));
  document.querySelectorAll('.sort-ind').forEach(el => {
    const c = +el.dataset.col; el.textContent = (ordenCol === c) ? (ordenDir === 1 ? '▲' : '▼') : '';
  });
}

function construirEncabezado(){
  const tr = document.getElementById('encabezado');
  HEADERS.forEach((h, ci) => {
    const th = document.createElement('th');
    if (EDIT_COLS.has(ci)) th.className = 'edit';
    const inner = document.createElement('div'); inner.className = 'th-inner';
    const lab = document.createElement('span'); lab.className = 'th-label';
    const lapiz = EDIT_COLS.has(ci) ? ' <span class="lapiz" title="editable">✎</span>' : '';
    lab.innerHTML = `${h}${lapiz} <span class="sort-ind" data-col="${ci}"></span>`;
    lab.title = 'Ordenar por ' + h;
    lab.onclick = () => { if (ordenCol === ci) ordenDir = -ordenDir; else { ordenCol = ci; ordenDir = 1; } render(); };
    const fb = document.createElement('span'); fb.className = 'filtro-btn'; fb.dataset.col = ci;
    fb.textContent = '▾'; fb.title = 'Filtrar ' + h;
    fb.onclick = e => { e.stopPropagation(); abrirDropdown(ci, fb); };
    inner.appendChild(lab); inner.appendChild(fb);
    th.appendChild(inner); tr.appendChild(th);
  });
}

function cerrarDropdown(){ if (ddAbierto){ ddAbierto.remove(); ddAbierto = null; } }

function abrirDropdown(col, btn){
  if (ddAbierto && ddAbierto.dataset.col == col){ cerrarDropdown(); return; }
  cerrarDropdown();
  const valores = valoresDisponibles(col);
  const seleccion = filtros.has(col) ? new Set(filtros.get(col)) : new Set(valores);

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
    <div class="dd-acc"><button class="ok">Aceptar</button><button class="cancel">Cancelar</button></div>`;
  document.body.appendChild(dd);
  ddAbierto = dd;

  const lista = dd.querySelector('.dd-list');
  const dibujar = (filtro="") => {
    lista.replaceChildren();
    const f = filtro.trim().toLowerCase();
    const visibles = valores.filter(v => !f || v.toLowerCase().includes(f));
    const todosL = document.createElement('label'); todosL.className = 'todos';
    const todosC = document.createElement('input'); todosC.type = 'checkbox';
    const mv = visibles.filter(v => seleccion.has(v)).length;
    todosC.checked = visibles.length > 0 && mv === visibles.length;
    todosC.indeterminate = mv > 0 && mv < visibles.length;
    todosC.onchange = () => { visibles.forEach(v => todosC.checked ? seleccion.add(v) : seleccion.delete(v)); dibujar(filtro); };
    const todosS = document.createElement('span'); todosS.className = 'dd-val'; todosS.textContent = '(Seleccionar todo)';
    todosL.appendChild(todosC); todosL.appendChild(todosS); lista.appendChild(todosL);
    for (const v of visibles){
      const l = document.createElement('label');
      const c = document.createElement('input'); c.type = 'checkbox'; c.checked = seleccion.has(v);
      c.onchange = () => { c.checked ? seleccion.add(v) : seleccion.delete(v);
        const m = visibles.filter(x => seleccion.has(x)).length;
        todosC.checked = m === visibles.length; todosC.indeterminate = m > 0 && m < visibles.length; };
      const s = document.createElement('span'); s.className = 'dd-val'; s.textContent = v;
      if (v === VACIO){ s.style.color = '#999'; s.style.fontStyle = 'italic'; }
      l.appendChild(c); l.appendChild(s); lista.appendChild(l);
    }
  };
  dibujar();
  dd.querySelector('.dd-search').oninput = e => dibujar(e.target.value);
  dd.querySelectorAll('.dd-ord button').forEach(b => b.onclick = () => { ordenCol = col; ordenDir = +b.dataset.dir; render(); cerrarDropdown(); });
  dd.querySelector('.ok').onclick = () => {
    if (seleccion.size === valores.length) filtros.delete(col); else filtros.set(col, seleccion);
    cerrarDropdown(); render();
  };
  dd.querySelector('.cancel').onclick = () => cerrarDropdown();

  const r = btn.getBoundingClientRect();
  dd.style.top = (r.bottom + 4) + 'px';
  let left = r.left; const w = 260;
  if (left + w > window.innerWidth - 8) left = window.innerWidth - w - 8;
  dd.style.left = Math.max(8, left) + 'px';
  dd.querySelector('.dd-search').focus();
}

document.addEventListener('mousedown', e => { if (ddAbierto && !ddAbierto.contains(e.target) && !e.target.classList.contains('filtro-btn')) cerrarDropdown(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') cerrarDropdown(); });
document.querySelector('.table-wrap').addEventListener('scroll', cerrarDropdown);
window.addEventListener('resize', cerrarDropdown);

document.getElementById('busqueda').addEventListener('input', render);
document.getElementById('limpiar').addEventListener('click', () => {
  filtros.clear(); document.getElementById('busqueda').value = ''; ordenCol = null; render();
});

// Guardar / cargar notas como archivo JSON (respaldo portable)
document.getElementById('guardar').addEventListener('click', () => {
  const blob = new Blob([JSON.stringify({version:1, notas}, null, 2)], {type:'application/json'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = 'notas_equipos.json';
  a.click(); URL.revokeObjectURL(a.href);
  marcarDirty(false);
});
document.getElementById('cargar').addEventListener('click', () => document.getElementById('archivo').click());
document.getElementById('archivo').addEventListener('change', ev => {
  const file = ev.target.files[0]; if (!file) return;
  const fr = new FileReader();
  fr.onload = () => {
    try {
      const obj = JSON.parse(fr.result);
      const entrantes = obj && obj.notas ? obj.notas : obj;
      if (typeof entrantes !== 'object') throw new Error('formato');
      for (const k in entrantes) notas[k] = Object.assign(notas[k] || {}, entrantes[k]);
      try { localStorage.setItem(LS_KEY, JSON.stringify(notas)); } catch(e){}
      marcarDirty(false); render();
      alert('Notas cargadas correctamente.');
    } catch(e){ alert('No se pudo leer el archivo de notas: ' + e.message); }
  };
  fr.readAsText(file); ev.target.value = '';
});

document.getElementById('exportar').addEventListener('click', () => {
  const vis = indicesVisibles();
  const esc = s => '"' + norm(s).replace(/"/g, '""') + '"';
  const lineas = [HEADERS.map(esc).join(',')];
  for (const i of vis){ const row = []; for (let c = 0; c < HEADERS.length; c++) row.push(esc(getCell(i, c))); lineas.push(row.join(',')); }
  const blob = new Blob(["﻿" + lineas.join("\r\n")], {type:'text/csv;charset=utf-8;'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = 'equipos_filtrado.csv';
  a.click(); URL.revokeObjectURL(a.href);
});

window.addEventListener('beforeunload', e => { if (dirty){ e.preventDefault(); e.returnValue = ''; } });

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
    nb = len(COLUMNAS_BASE)
    html = (PLANTILLA
            .replace("__HEADERS__", json.dumps(headers, ensure_ascii=False))
            .replace("__ROWS__", json.dumps(datos, ensure_ascii=False, separators=(",", ":")))
            .replace("__NUMERICAS__", json.dumps(NUMERICAS))
            .replace("__NB__", str(nb))
            .replace("__EDIT_OBS__", str(nb))        # índice de "Observaciones"
            .replace("__EDIT_NOTAS__", str(nb + 1)))  # índice de "Notas"
    args.salida.write_text(html, encoding="utf-8")
    print(f"Generado: {args.salida}  ({len(datos)} equipos, {len(html)//1024} KB)")


if __name__ == "__main__":
    main()
