#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vuelca a equipos.db las notas/observaciones que se exportaron desde la tabla
HTML (archivo notas_equipos.json), de modo que la base de datos quede como
fuente de verdad. Al volver a generar el HTML, esas notas aparecerán ya
incorporadas.

La clave de cada equipo es la misma que usa el HTML:
    Serie  ->  N° Inventario  ->  "#" + ID   (el primero no vacío)

Uso:
    python scripts/aplicar_notas.py notas_equipos.json
    python scripts/aplicar_notas.py notas_equipos.json --db equipos.db
"""
from __future__ import annotations

import argparse
import json
import sqlite3
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
DB_POR_DEFECTO = RAIZ / "equipos.db"


def clave_equipo(serie, inventario, id_):
    s = (serie or "").strip()
    if s:
        return s
    i = (inventario or "").strip()
    if i:
        return i
    return "#" + str(id_).strip()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("json", type=Path, help="Archivo notas_equipos.json exportado desde el HTML")
    ap.add_argument("--db", type=Path, default=DB_POR_DEFECTO)
    args = ap.parse_args()

    if not args.json.exists():
        return print(f"ERROR: no existe {args.json}") or 1

    obj = json.loads(args.json.read_text(encoding="utf-8"))
    entrantes = obj.get("notas", obj) if isinstance(obj, dict) else {}
    if not isinstance(entrantes, dict) or not entrantes:
        return print("No hay notas que aplicar en el archivo.") or 0

    con = sqlite3.connect(args.db)
    con.row_factory = sqlite3.Row

    # Mapa clave -> id (misma lógica que el HTML).
    mapa = {}
    for r in con.execute("SELECT id, serie, n_inventario FROM equipos"):
        mapa[clave_equipo(r["serie"], r["n_inventario"], r["id"])] = r["id"]

    aplicadas, sin_match = 0, []
    for k, campos in entrantes.items():
        if k not in mapa or not isinstance(campos, dict):
            if k not in mapa:
                sin_match.append(k)
            continue
        sets, valores = [], []
        # Campos de texto libre.
        for campo in ("observaciones", "notas"):
            if campo in campos:
                sets.append(f"{campo} = ?")
                v = campos[campo]
                valores.append(v if (v is None or str(v).strip() != "") else None)
        # Historial de intervenciones (lista) + columnas de consulta derivadas.
        if "registros" in campos:
            regs = campos["registros"] if isinstance(campos["registros"], list) else []
            pendiente = 1 if any(r.get("p") for r in regs) else (0 if regs else None)
            fechas = [r.get("f") for r in regs if r.get("f")]
            ultima = max(fechas) if fechas else None
            vencs = [r.get("v") for r in regs if r.get("p") and r.get("v")]
            proximo = min(vencs) if vencs else None
            sets += ["registros = ?", "pendiente = ?", "ultima_intervencion = ?", "proximo_vencimiento = ?"]
            valores += [json.dumps(regs, ensure_ascii=False) if regs else None, pendiente, ultima, proximo]
        if sets:
            valores.append(mapa[k])
            con.execute(f"UPDATE equipos SET {', '.join(sets)} WHERE id = ?", valores)
            aplicadas += 1
    con.commit()
    con.close()

    print(f"Notas aplicadas a {aplicadas} equipo(s).")
    if sin_match:
        print(f"Sin coincidencia ({len(sin_match)}): {', '.join(sin_match[:10])}"
              + (" …" if len(sin_match) > 10 else ""))
    print("Listo. Vuelve a generar el HTML con: python scripts/generar_html.py")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
