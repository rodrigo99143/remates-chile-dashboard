# -*- coding: utf-8 -*-
"""
Script de SOLO INVESTIGACIÓN (2026-11-05): revisar, con datos reales de la
API de la TGR, qué valores trae el campo "tipoDeuda" (para la columna nueva
Fiscal/Territorial que pidió Álvaro) y confirmar el caso de Lastarria 70
(rol 070-1713-60) con los datos crudos tal cual los entrega la TGR.

No modifica nada del sistema real. Uso: python investigar_campos_tgr.py
"""
import json

import requests

URL_TGR = "https://remates.tgr.cl/v1/getListaRematesActivos"
HEADERS = {"User-Agent": "Mozilla/5.0"}


def main():
    r = requests.get(URL_TGR, headers=HEADERS, timeout=60)
    print("status", r.status_code)
    data = r.json()["data"]
    print("total filas:", len(data))

    tipos_deuda = {}
    for row in data:
        td = row.get("tipoDeuda")
        tipos_deuda[td] = tipos_deuda.get(td, 0) + 1
    print("valores distintos de 'tipoDeuda' y cuántas veces aparece cada uno:")
    for k, v in sorted(tipos_deuda.items(), key=lambda kv: -kv[1]):
        print(f"  {k!r}: {v}")

    print()
    print("Buscando el caso de Lastarria 70 (rol 070-1713-60) para confirmar campos crudos...")
    encontrado = False
    for row in data:
        rol_fmt = str(row.get("rolFormato") or "")
        rol_11 = str(row.get("rol") or "")
        direccion = str(row.get("direccionRol") or "").upper()
        if "1713" in rol_fmt or "1713" in rol_11 or "LASTARRIA" in direccion:
            encontrado = True
            print(json.dumps(row, indent=2, ensure_ascii=False))
    if not encontrado:
        print("  No se encontró (puede que este remate ya no esté 'activo' en la lista de hoy).")

    print()
    print("Mostrando 3 filas cualquiera completas, para ver todos los campos que trae la TGR:")
    for row in data[:3]:
        print(json.dumps(row, indent=2, ensure_ascii=False))
        print("-" * 60)


if __name__ == "__main__":
    main()
