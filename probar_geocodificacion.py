# -*- coding: utf-8 -*-
"""
Script de una sola vez (2026-10-28) para revisar, con Nominatim real, las 12
direcciones del reporte acumulado que quedaron sin coordenadas - y confirmar
cuáles se arreglan con la mejora de limpieza de direcciones (ver
_limpiar_para_geocodificar en generar_pdf.py) y cuáles no (porque son
direcciones rurales/informales sin una calle real que mapear - eso no tiene
arreglo con solo "limpiar texto").

No corre el programa completo (no necesita TGR, Macal, el catastro, etc.) -
solo geocodifica estas 12 direcciones puntuales, así que es rápido.

Uso: python probar_geocodificacion.py
"""
import time
from pathlib import Path

DIRECCIONES = [
    ("LT B- 1 PROL CALLE EJERC. LIBER. 2735", "RIO BUENO"),
    ("CANAL SARMIENTO 6081 -F DP- 304", "PUERTO MONTT"),
    ("V MACKENNA EDIF C 7735 DP 702", "LA FLORIDA"),
    ("HJ. 2 LOS MAITENES LT C", "MULCHEN"),
    ("AV G J ARTIGAS 3121 DP 813", "NUNOA"),
    ("PJ PARTICULAR 64 -A", "QUILPUE"),
    ("PJE. GLACIAR NEF 41 LT D 7", "QUILPUE"),
    ("LOTE A 2 C 2", "VILLA ALEMANA"),
    ("HDA RENACA PC 154", "VINA DEL MAR"),
    ("AV LAS AMERICAS REMANENTE 1", "VILLA ALEMANA"),
    ("LAS HIJUELAS 2615 B 6", "QUILPUE"),
    ("PJE. ISLA MOCHA 892 CS 3", "QUILPUE"),
]


def main():
    import generar_pdf
    from geopy.geocoders import Nominatim

    geocoder = Nominatim(user_agent=generar_pdf.GEOCODER_UA)
    cache = {}

    print(f"{'DIRECCIÓN ORIGINAL':45} | {'CONSULTA LIMPIA':35} | RESULTADO")
    print("-" * 110)
    arreglados = []
    siguen_sin = []
    for direccion, comuna in DIRECCIONES:
        consulta = generar_pdf._limpiar_para_geocodificar(direccion)
        coords = None
        try:
            coords = generar_pdf._geocodificar(direccion, comuna, cache, geocoder)
        except Exception as e:
            print(f"  (error geocodificando: {e})")
        if coords:
            resultado = f"OK -> {coords[0]:.5f}, {coords[1]:.5f}"
            arreglados.append((direccion, comuna, coords))
        else:
            resultado = "sigue sin encontrarse"
            siguen_sin.append((direccion, comuna))
        print(f"{direccion:45} | {consulta:35} | {resultado}")
        time.sleep(1.1)

    print()
    print(f"RESUMEN: {len(arreglados)} de {len(DIRECCIONES)} se arreglaron con la mejora de limpieza.")
    print()
    print("ARREGLADAS:")
    for d, c, coords in arreglados:
        print(f"  - {d} ({c}) -> {coords[0]:.5f}, {coords[1]:.5f}")
    print()
    print("SIGUEN SIN COORDENADAS (dirección rural/informal, sin calle real que mapear):")
    for d, c in siguen_sin:
        print(f"  - {d} ({c})")


if __name__ == "__main__":
    main()
