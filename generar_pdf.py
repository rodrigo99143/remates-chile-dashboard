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
SUBTITULO = ParagraphStyle("subtitulo", parent=ESTILOS["Normal"], textColor=GRIS, fontSize=10)
SECCION = ParagraphStyle("seccion", parent=ESTILOS["Normal"], textColor=colors.white, fontSize=10,
                          leading=14, leftIndent=4)
OPORTUNIDAD_GRANDE = ParagraphStyle("oport", parent=ESTILOS["Normal"], fontSize=22, textColor=NARANJA, leading=26)
CELDA = ParagraphStyle("celda", parent=ESTILOS["Normal"], fontSize=9, textColor=colors.black)
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
                     ("TOPPADDING", (0, 0), (-1, -1), 4),
                     ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                     ("LEFTPADDING", (0, 0), (-1, -1), 6),
                 ]))


def _tabla_datos(pares):
    data = [[k, Paragraph(str(v), CELDA)] for k, v in pares]
    t = Table(data, colWidths=[62 * mm, 108 * mm])
    t.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TEXTCOLOR", (0, 0), (0, -1), GRIS),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DDDDDD")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, AZUL_CLARO]),
    ]))
    return t


def _geocodificar(direccion, comuna, cache, geocoder):
    clave = f"{direccion}, {comuna}, Chile"
    if clave in cache:
        return cache[clave]
    resultado = None
    try:
        loc = geocoder.geocode(clave, timeout=10)
        if loc:
            resultado = (loc.latitude, loc.longitude)
    except Exception:
        resultado = None
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
    elementos.append(Spacer(1, 6 * mm))

    pct = fila.get("oportunidad_pct")
    if not _vacio(pct):
        elementos.append(Paragraph(f"Oportunidad estimada: {pct:+.0f}% vs. valor de referencia", OPORTUNIDAD_GRANDE))
    else:
        elementos.append(Paragraph("Oportunidad estimada: sin datos suficientes", OPORTUNIDAD_GRANDE))
    elementos.append(Spacer(1, 4 * mm))

    # ---- Mapa (si se logra geocodificar la dirección) ----
    try:
        coords = _geocodificar(direccion, comuna, cache_geo, geocoder)
        if coords:
            ruta_png = carpeta_tmp / f"mapa_{idx}.png"
            _mapa_png(coords[0], coords[1], ruta_png)
            elementos.append(RLImage(str(ruta_png), width=170 * mm, height=99 * mm))
            elementos.append(Spacer(1, 4 * mm))
    except Exception:
        pass  # sin mapa esta vez, no es razón para botar la ficha completa

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
