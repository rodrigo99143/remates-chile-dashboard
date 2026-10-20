# -*- coding: utf-8 -*-
"""
Genera un PDF con una ficha por propiedad (estilo "informe de propiedad" como
el de avaluos.cl que compartió Rodrigo): portada + una página por propiedad,
con mapa de ubicación, datos clave y alertas en lenguaje simple.

Pensado para correr en GitHub Actions (tiene internet completo para geocodificar
direcciones y bajar el mapa). Si algo de red falla, la ficha se arma igual, solo
que sin el mapa - nunca debe hacer fallar la corrida completa.

Requiere (se instalan solas si faltan, igual que el resto del programa):
  reportlab, geopy, staticmap, Pillow
"""
import importlib
import subprocess
import sys
import time
import datetime as dt
from pathlib import Path


def instalar(paquete, modulo=None):
    modulo = modulo or paquete
    try:
        importlib.import_module(modulo)
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", paquete], check=True)


for paquete, modulo in [("reportlab", "reportlab"), ("geopy", "geopy"),
                         ("staticmap", "staticmap"), ("Pillow", "PIL")]:
    instalar(paquete, modulo)

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
                                 Image as RLImage, PageBreak)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

AZUL = colors.HexColor("#1F4E79")
AZUL_CLARO = colors.HexColor("#EAF1F8")
NARANJA = colors.HexColor("#E8793A")
GRIS = colors.HexColor("#555555")

ESTILOS = getSampleStyleSheet()
TITULO = ParagraphStyle("titulo", parent=ESTILOS["Title"], textColor=AZUL, fontSize=17, spaceAfter=2)
SUBTITULO = ParagraphStyle("subtitulo", parent=ESTILOS["Normal"], textColor=GRIS, fontSize=10, leading=12)
SECCION = ParagraphStyle("seccion", parent=ESTILOS["Normal"], textColor=colors.white, fontSize=10,
                          leading=12, leftIndent=4)
OPORTUNIDAD_GRANDE = ParagraphStyle("oport", parent=ESTILOS["Normal"], fontSize=20, textColor=NARANJA, leading=23)
# OJO (auditoría 2026-10-20): "Normal" trae leading=14 por defecto sin importar
# el fontSize que uno le ponga - con fontSize=9 eso dejaba cada fila de la
# tabla ~40% más alta de lo necesario, y era la causa principal de que la
# ficha no entrara en una sola página (más que el tamaño del mapa). Por eso
# acá SÍ se fija el leading a mano, un poco más que el fontSize (9 -> 11).
CELDA = ParagraphStyle("celda", parent=ESTILOS["Normal"], fontSize=9, leading=11, textColor=colors.black)
PORTADA_TXT = ParagraphStyle("portada_txt", parent=ESTILOS["Normal"], fontSize=11, textColor=GRIS, leading=16)

# Alertas en lenguaje simple para quien lee el PDF (no programador). Las que no
# están acá simplemente se muestran con su nombre tal cual, en vez de caerse.
ALERTAS_PDF = {
    "DERECHOS_PARCIALES_O_CUOTA": "Se rematan derechos o una cuota del inmueble, no la propiedad completa",
    "MULTIPLES_MINIMOS": "El aviso trae más de un mínimo (ej. inmueble y derechos de agua por separado)",
    "UF_SIN_CONVERTIR": "No se pudo convertir el mínimo (estaba en UF) a pesos",
    "ROLES_PARCIALES": "Hay varios roles involucrados y no todos se encontraron en el catastro",
    "DIRECCION_NO_COINCIDE": "La dirección del aviso no coincide exactamente con la del catastro",
    "RATIO_TASACION_AVALUO_ANORMAL": "La tasación de la TGR se ve fuera de lo normal respecto del avalúo",
    "ROL_ES_MATRIZ_DE_CONDOMINIO": "El rol corresponde a todo el edificio/condominio, no a una unidad",
    "OPORTUNIDAD_SOSPECHOSA": "El % de oportunidad es inusualmente alto: conviene revisar el dato a mano",
    "MINIMO_ESTIMADO_TGR": "La TGR no publica el mínimo: se estimó a partir de la tasación",
    "OCUPADA": "La propiedad figura como ocupada: verificar antes de participar",
    "VARIAS_UNIDADES": "El remate incluye más de una unidad (ej. depto + bodega + estacionamiento)",
    "POSIBLE_PREDIO_RURAL": "Podría tratarse de un predio rural (parcela, hijuela, fundo)",
    "RURAL_O_AGRICOLA": "Catastrado como predio rural o agrícola",
    "VARIOS_ROLES": "El aviso menciona más de un rol de avalúo",
    "SIN_AVALUO_SII": "No se encontró avalúo SII para este rol",
    "INCLUYE_DERECHOS_DE_AGUA": "El remate incluye derechos de aprovechamiento de agua",
    "ROL_DEDUCIDO_POR_DIRECCION": "El rol se dedujo por la dirección (no venía explícito en el aviso)",
    "REMATE_YA_REALIZADO_O_VENCIDO": "La fecha de remate ya pasó: verificar vigencia antes de actuar",
}

GEOCODER_UA = "remates-chile-dashboard/1.0 (contacto: rbastias1991@gmail.com)"

# Auditoría 2026-10-20 (corrida de validación): de 30 fichas, 11 salieron sin
# mapa. Las 11 direcciones tenían algo en común - casi todas traen un
# descriptor de unidad/lote PEGADO después del número de la calle (ej. "AV
# DEL VALLE 662 OF 601", "SANTA VICTORIA 411 DP 719", "AMSTERDAM ORIENTE 4401
# M 8 L 23"), y Nominatim (el geocodificador gratuito que usamos) no logra
# encontrar la calle cuando la consulta incluye ese texto extra - busca "calle
# número OF 601" como si fuera parte de la dirección, no la encuentra, y
# devuelve vacío. La solución: para geocodificar (solo para ESO, no para lo
# que se muestra en la ficha) se recorta la dirección hasta el primer número,
# descartando cualquier descriptor de depto/oficina/lote que venga después -
# igual que ya hace remates_v2.py para cruzar contra el catastro del SII (ver
# MARCAS_UNIDAD / MARCAS_LOTE ahí). Un par de direcciones (hijuelas/parcelas
# rurales, sin un "calle + número" reconocible) van a seguir sin poder
# geocodificarse - eso es una limitación real de los mapas gratuitos con
# direcciones rurales, no un error del programa.
_MARCAS_CORTE = {
    "DP", "DPTO", "DEPTO", "OF", "OFICINA", "BX", "BD", "LC", "ES", "BO", "EST",
    "LT", "LO", "LOTE", "MZ", "MZA", "M", "ST", "SI", "CA", "CS", "SITIO", "KM",
    "EDIF", "EDIFICIO", "BL", "BLOCK", "BLOQUE", "TORRE", "HJ", "HIJUELA",
    "CHACRA", "RESTO", "EX",
}
import re as _re
_RE_NUMERO = _re.compile(r"\d+")


def _limpiar_para_geocodificar(direccion):
    """Deja la dirección hasta el primer número + lo que sigue pegado (ej. '662'
    o '411'), cortando antes del primer descriptor de depto/oficina/lote que
    venga después (ver _MARCAS_CORTE arriba). Si no hay número reconocible,
    devuelve la dirección tal cual (es probablemente rural; no hay mucho más
    que limpiar)."""
    m = _RE_NUMERO.search(direccion)
    if not m:
        return direccion
    palabras = direccion[m.end():].split()
    corte = len(direccion)
    for palabra in palabras:
        p = palabra.strip(".,").upper()
        if p in _MARCAS_CORTE:
            # recorta desde donde empieza esa palabra en el string original
            idx = direccion.index(palabra, m.end())
            corte = idx
            break
    return direccion[:corte].strip(" ,.-")


def _vacio(v):
    """True si v es None, NaN (float u objeto pandas) o string vacío. OJO: un
    NaN de pandas es *truthy* en Python ('v or "—"' NO lo detecta), por eso
    todo el módulo pasa los valores por esta función antes de mostrarlos."""
    if v is None:
        return True
    if isinstance(v, float) and v != v:
        return True
    try:
        import pandas as pd
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(v, str) and not v.strip():
        return True
    return False


def texto(v, defecto="—"):
    return defecto if _vacio(v) else str(v)


def peso(v):
    if _vacio(v):
        return "—"
    return "$" + f"{v:,.0f}".replace(",", ".")


def num(v, sufijo=""):
    if _vacio(v):
        return "—"
    try:
        return f"{float(v):,.0f}{sufijo}".replace(",", ".")
    except (TypeError, ValueError):
        return f"{v}{sufijo}"


def texto_alertas(alertas_txt):
    if _vacio(alertas_txt):
        return "Sin alertas registradas"
    codigos = [a.strip() for a in str(alertas_txt).split(";") if a.strip()]
    return " · ".join(ALERTAS_PDF.get(c, c) for c in codigos)


def _fila_seccion(titulo_txt):
    return Table([[Paragraph(titulo_txt, SECCION)]], colWidths=[170 * mm],
                 style=TableStyle([
                     ("BACKGROUND", (0, 0), (-1, -1), AZUL),
                     ("TOPPADDING", (0, 0), (-1, -1), 3),
                     ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                     ("LEFTPADDING", (0, 0), (-1, -1), 6),
                 ]))


def _tabla_datos(pares):
    # Padding de celda reducido (antes 4/4) a propósito - ver nota junto al
    # mapa más abajo: con padding 4 + mapa a tamaño completo, la ficha no
    # entraba en una sola página (2026-10-20).
    data = [[k, Paragraph(str(v), CELDA)] for k, v in pares]
    t = Table(data, colWidths=[62 * mm, 108 * mm])
    t.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TEXTCOLOR", (0, 0), (0, -1), GRIS),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DDDDDD")),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, AZUL_CLARO]),
    ]))
    return t


def _geocodificar(direccion, comuna, cache, geocoder):
    clave = f"{direccion}, {comuna}, Chile"
    if clave in cache:
        return cache[clave]
    direccion_limpia = _limpiar_para_geocodificar(direccion)
    consulta = f"{direccion_limpia}, {comuna}, Chile"
    resultado = None
    motivo = None
    try:
        loc = geocoder.geocode(consulta, timeout=10)
        if loc:
            resultado = (loc.latitude, loc.longitude)
        else:
            motivo = "Nominatim no encontró resultados"
    except Exception as e:
        motivo = f"error al geocodificar: {e}"
    if resultado is None:
        # Se deja impreso en el log de la corrida (GitHub Actions) para poder
        # revisar más adelante qué direcciones quedan sin mapa y por qué -
        # antes esto fallaba en silencio y no había forma de saber la causa.
        print(f"  [PDF-MAPA] sin mapa para '{direccion}' ({comuna}): {motivo} "
              f"(consulta usada: '{consulta}')")
    cache[clave] = resultado
    time.sleep(1.1)  # política de uso de Nominatim: máx. 1 solicitud por segundo
    return resultado


def _mapa_png(lat, lon, ruta_png):
    from staticmap import StaticMap, CircleMarker
    m = StaticMap(600, 350, url_template="https://a.tile.openstreetmap.org/{z}/{x}/{y}.png")
    m.add_marker(CircleMarker((lon, lat), "#E8793A", 14))
    m.add_marker(CircleMarker((lon, lat), "#FFFFFF", 6))
    imagen = m.render(zoom=15)
    imagen.save(ruta_png)


def _ficha_propiedad(fila, cache_geo, geocoder, carpeta_tmp, idx):
    direccion = texto(fila.get("direccion_final"), "(sin dirección)")
    comuna = texto(fila.get("comuna_propiedad"), "")
    region = texto(fila.get("region"), "")
    elementos = []
    elementos.append(Paragraph("Ficha de Oportunidad de Remate", TITULO))
    elementos.append(Paragraph(f"{direccion} — {comuna}, {region}", SUBTITULO))
    elementos.append(Spacer(1, 3 * mm))

    pct = fila.get("oportunidad_pct")
    if not _vacio(pct):
        elementos.append(Paragraph(f"Oportunidad estimada: {pct:+.0f}% vs. valor de referencia", OPORTUNIDAD_GRANDE))
    else:
        elementos.append(Paragraph("Oportunidad estimada: sin datos suficientes", OPORTUNIDAD_GRANDE))
    elementos.append(Spacer(1, 2 * mm))

    # ---- Mapa (si se logra geocodificar la dirección) ----
    # Tamaño reducido (antes 170x99mm) a propósito: con el mapa a tamaño
    # completo la ficha no alcanzaba a entrar en una sola página y la sección
    # REMATE se pasaba a una segunda hoja (detectado en la auditoría del
    # 2026-10-20: 19 de 30 fichas con mapa ocupaban 2 páginas). A este tamaño
    # entra todo en una página, con o sin mapa.
    try:
        coords = _geocodificar(direccion, comuna, cache_geo, geocoder)
        if coords:
            ruta_png = carpeta_tmp / f"mapa_{idx}.png"
            _mapa_png(coords[0], coords[1], ruta_png)
            elementos.append(RLImage(str(ruta_png), width=130 * mm, height=76 * mm))
            elementos.append(Spacer(1, 2 * mm))
    except Exception as e:
        print(f"  [PDF-MAPA] no se pudo generar la imagen del mapa para '{direccion}' ({comuna}): {e}")

    tipo = fila.get("tipo_propiedad")
    tipo = tipo if not _vacio(tipo) else fila.get("destino_desc")
    confianza = fila.get("confianza_mercado_m2")
    confianza = confianza if not _vacio(confianza) else fila.get("confianza")
    link = fila.get("url")
    link = link if not _vacio(link) else fila.get("link_poder_judicial")

    elementos.append(_fila_seccion("IDENTIFICACIÓN DE LA PROPIEDAD"))
    elementos.append(_tabla_datos([
        ("Dirección", direccion),
        ("Comuna", comuna),
        ("Región", region),
        ("Rol de avalúo (Tesorería)", texto(fila.get("rol_completo"))),
        ("Tipo", texto(tipo)),
        ("Dormitorios / Baños", f"{texto(fila.get('dormitorios'))} / {texto(fila.get('banos'))}"),
        ("Estado de ocupación", texto(fila.get("ocupacion"))),
        ("Superficie útil / terreno", f"{num(fila.get('superficie_util_m2'), ' m2')} / "
                                       f"{num(fila.get('superficie_total_m2'), ' m2')}"),
        ("Año de construcción", num(fila.get("anio_construccion"))),
    ]))
    elementos.append(Spacer(1, 3 * mm))

    elementos.append(_fila_seccion("VALORES Y OPORTUNIDAD"))
    elementos.append(_tabla_datos([
        ("Avalúo fiscal (SII)", peso(fila.get("avaluo_sii"))),
        ("Mínimo de remate", f"{peso(fila.get('minimo_pesos'))}  ({texto(fila.get('tipo_minimo'))})"),
        ("Valor de referencia usado", peso(fila.get("valor_ref_pesos"))),
        ("Confianza del dato de mercado", texto(confianza)),
    ]))
    elementos.append(Spacer(1, 3 * mm))

    elementos.append(_fila_seccion("REMATE"))
    fecha_remate = fila.get("fecha_remate")
    fecha_txt = fecha_remate.strftime("%d-%m-%Y") if hasattr(fecha_remate, "strftime") else texto(fecha_remate)
    elementos.append(_tabla_datos([
        ("Fecha", fecha_txt),
        ("Hora", texto(fila.get("hora_remate"))),
        ("Modalidad", texto(fila.get("modalidad"))),
        ("Fuente", texto(fila.get("fuente"))),
        ("Tribunal", texto(fila.get("tribunal"))),
        ("Acreedor / Mandante", texto(fila.get("mandante_acreedor"))),
        ("Alertas", texto_alertas(fila.get("alertas"))),
        ("Link de la publicación", texto(link)),
    ]))
    elementos.append(PageBreak())
    return elementos


def construir_pdf(df_filas, ruta_salida):
    """df_filas: DataFrame con las propiedades a incluir (ya filtradas y
    ordenadas por quien llama). ruta_salida: Path del .pdf a generar."""
    import tempfile
    if df_filas is None or len(df_filas) == 0:
        return None

    from geopy.geocoders import Nominatim
    geocoder = Nominatim(user_agent=GEOCODER_UA)
    cache_geo = {}

    carpeta_tmp = Path(tempfile.mkdtemp(prefix="mapas_pdf_"))
    elementos = []

    # ---- Portada ----
    elementos.append(Spacer(1, 40 * mm))
    elementos.append(Paragraph("Oportunidades de Remate", TITULO))
    elementos.append(Spacer(1, 4 * mm))
    elementos.append(Paragraph(
        f"{len(df_filas)} propiedades nuevas desde el último reporte — "
        f"generado el {dt.datetime.now().strftime('%d-%m-%Y %H:%M')}", PORTADA_TXT))
    elementos.append(PageBreak())

    for idx, (_, fila) in enumerate(df_filas.iterrows()):
        elementos.extend(_ficha_propiedad(fila, cache_geo, geocoder, carpeta_tmp, idx))

    doc = SimpleDocTemplate(str(ruta_salida), pagesize=A4, topMargin=16 * mm, bottomMargin=16 * mm,
                             leftMargin=20 * mm, rightMargin=20 * mm)
    doc.build(elementos)
    return ruta_salida
