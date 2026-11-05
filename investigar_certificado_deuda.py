# -*- coding: utf-8 -*-
"""
Script de SOLO INVESTIGACIÓN (2026-11-05), pedido por Rodrigo: revisar si es
posible automatizar la obtención del "certificado de pago de deuda de
contribuciones" de la TGR (lo pidió Álvaro, para descartar remates
TERRITORIALES con menos de 3 años de deuda - ver conversación del 04-11).

Esto NO modifica remates_v2.py ni el flujo de correo. Es un sondeo aparte:
- Descarga la página de ChileAtiende y de la TGR que describen el trámite.
- Busca en el HTML/JS de la TGR cualquier pista de una API tipo
  "remates.tgr.cl/v1/..." (que es como ya se obtiene la lista de remates en
  remates_v2.py) u otro endpoint JSON que el formulario use por debajo.
- Si encuentra algo prometedor, intenta un pedido de prueba para la
  propiedad de Lastarria 70 (rol 070-1713-60, Región Metropolitana, comuna
  Santiago) para ver si se puede obtener el dato sin notario/clave única.

Todo lo que encuentre (o no encuentre) queda impreso para revisarlo en el
log de la corrida - no se guarda nada más, no se toca ninguna otra parte del
sistema.

Uso: python investigar_certificado_deuda.py
"""
import re
import sys

import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0 Safari/537.36",
}

PAGINAS_A_REVISAR = [
    "https://www.chileatiende.gob.cl/fichas/4288-certificado-de-pago-de-deuda-de-contribuciones-de-bienes-raices",
    "https://www.tgr.cl/tramites-tgr/certificado-de-pago-de-deudas-de-contribuciones/",
    "https://www.tgr.cl/tramites-tgr/certificado-de-deuda-de-contribuciones/",
]

# Propiedad de prueba conocida (del reporte del 02-11, caso Lastarria/Álvaro)
PRUEBA = {"region": "Región Metropolitana de Santiago", "comuna": "Santiago", "rol": "070-1713-60"}

RE_ENDPOINT = re.compile(
    r"""["'](?:https?://)?[a-zA-Z0-9_.\-]*tgr\.cl[a-zA-Z0-9_\-/.]*(?:v1|api|certificado|deuda)[a-zA-Z0-9_\-/.]*["']""",
    re.IGNORECASE,
)
RE_SCRIPT_SRC = re.compile(r'<script[^>]+src=["\']([^"\']+)["\']', re.IGNORECASE)


def descargar(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
        print(f"  GET {url} -> HTTP {r.status_code} ({len(r.text)} bytes)")
        return r
    except Exception as e:
        print(f"  GET {url} -> FALLÓ: {e}")
        return None


def buscar_pistas_endpoint(texto, origen):
    hallados = set(m.group(0).strip("\"'") for m in RE_ENDPOINT.finditer(texto))
    if hallados:
        print(f"  [{origen}] posibles endpoints/strings con 'tgr.cl' + v1/api/certificado/deuda:")
        for h in sorted(hallados):
            print(f"     - {h}")
    return hallados


def main():
    print("=" * 90)
    print("PASO 1: revisar las páginas públicas que describen el trámite")
    print("=" * 90)
    pistas = set()
    scripts_js = set()
    for url in PAGINAS_A_REVISAR:
        r = descargar(url)
        if r is None or r.status_code != 200:
            continue
        pistas |= buscar_pistas_endpoint(r.text, url)
        for m in RE_SCRIPT_SRC.finditer(r.text):
            src = m.group(1)
            if "tgr.cl" in src or src.startswith("/"):
                scripts_js.add(src if src.startswith("http") else "https://www.tgr.cl" + src)

    print()
    print(f"  Scripts JS encontrados en la(s) página(s) de la TGR: {len(scripts_js)}")
    for s in sorted(scripts_js)[:15]:
        print(f"     - {s}")

    print()
    print("=" * 90)
    print("PASO 1b: revisar links/botones de la página de la TGR (a dónde lleva el trámite)")
    print("=" * 90)
    RE_LINK = re.compile(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', re.IGNORECASE | re.DOTALL)
    for url in PAGINAS_A_REVISAR:
        r = descargar(url)
        if r is None or r.status_code != 200:
            continue
        print(f"  [{url}] links que mencionan certificado/deuda/tramite/clave/login/e-services/sii:")
        vistos = set()
        for m in RE_LINK.finditer(r.text):
            href, texto_link = m.group(1), re.sub(r"<[^>]+>", " ", m.group(2)).strip()
            junto = (href + " " + texto_link).lower()
            if any(p in junto for p in ("certificad", "deuda", "tramite", "clave", "login",
                                          "e-services", "eservices", "sii.cl", "contribuc")):
                clave = (href, texto_link[:80])
                if clave in vistos:
                    continue
                vistos.add(clave)
                print(f"     - href={href!r}  texto={texto_link[:80]!r}")

    print()
    print("=" * 90)
    print("PASO 2: revisar esos scripts JS en busca de endpoints (bundle de Angular/React)")
    print("=" * 90)
    for s in sorted(scripts_js)[:15]:
        r = descargar(s)
        if r is None or r.status_code != 200:
            continue
        pistas |= buscar_pistas_endpoint(r.text, s)

    print()
    print("=" * 90)
    print(f"RESUMEN PASO 1-2: {len(pistas)} pistas de endpoint encontradas en total.")
    print("=" * 90)
    for p in sorted(pistas):
        print(f"  - {p}")

    if not pistas:
        print()
        print("  No se encontró ninguna pista clara de un endpoint JSON/API para el")
        print("  certificado (a diferencia de 'getListaRematesActivos', que sí es una API")
        print("  abierta). Es probable que este trámite funcione con un formulario que")
        print("  renderiza un PDF en el servidor (no una API fácil de llamar directo), o")
        print("  que requiera pasos intermedios (sesión, token, captcha) no visibles solo")
        print("  mirando el HTML/JS público.")

    print()
    print("=" * 90)
    print("PASO 3: probar un pedido directo con la propiedad de Lastarria 70 (070-1713-60)")
    print("=" * 90)
    print(f"  Propiedad de prueba: {PRUEBA}")
    # Intentos "educados" de variantes de endpoint, basados en el patrón ya
    # conocido de remates.tgr.cl/v1/... - se prueban aunque no se haya
    # encontrado una pista exacta, por si acaso siguen la misma convención.
    intentos = [
        "https://remates.tgr.cl/v1/getCertificadoDeudaContribuciones",
        "https://remates.tgr.cl/v1/certificadoDeuda",
        "https://www.tgr.cl/v1/getCertificadoDeudaContribuciones",
    ]
    for url in intentos:
        try:
            r = requests.get(url, headers=HEADERS, params=PRUEBA, timeout=20)
            print(f"  GET {url} -> HTTP {r.status_code} | primeros 200 caracteres: {r.text[:200]!r}")
        except Exception as e:
            print(f"  GET {url} -> FALLÓ: {e}")

    print()
    print("=" * 90)
    print("CONCLUSIÓN (a completar por Claude leyendo este log)")
    print("=" * 90)


if __name__ == "__main__":
    main()
