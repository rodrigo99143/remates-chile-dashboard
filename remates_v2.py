#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
REMATES DE PROPIEDADES EN CHILE  -  BARRIDO COMPLETO + OPORTUNIDADES  (v2)
================================================================================
Qué hace, en cada ejecución y todo automático:
  1. Instala solo las bibliotecas que falten.
  2. Barre EN VIVO las fuentes: TGR (remates de Tesorería), MACAL (macal.cl, todas las
     propiedades publicadas) y los AVISOS (.docx de resúmenes semanales en la carpeta "avisos").
  3. Usa el catastro nacional del SII (archivos BRORGA...N_NAC_... y ...A_NAC_...) para:
       - entregar dirección y avalúo fiscal a partir del ROL (comuna + manzana + predio), o
       - encontrar el ROL a partir de la dirección cuando la fuente no lo trae.
  4. Calcula la OPORTUNIDAD de cada propiedad y elige las 30 mejores del país, con cupos
     por zona (por defecto 2 de la provincia de Concepción, 2 de la Región Metropolitana
     y 2 de la Región de Valparaíso). Todo se cambia en la sección CONFIGURACIÓN.
  5. Guarda Excel y CSV en la carpeta "salidas".

Cómo se ejecuta (Terminal, dentro de la carpeta del proyecto):
    /usr/local/bin/python3 remates_v2.py
Pruebas sin internet:      /usr/local/bin/python3 remates_v2.py --prueba
Solo preparar catastro:    /usr/local/bin/python3 remates_v2.py --indexar
Saltar una fuente:         --sin-macal   --sin-avisos   --sin-tgr

Reglas de este programa:
  * Datos 100% en vivo. Si una fuente falla, se avisa en pantalla y en la hoja RESUMEN
    (nunca se rellena con datos antiguos). La TGR es obligatoria: si falla, el programa se detiene.
  * Nunca inventa valores. Lo estimado se rotula "ESTIMADO"; lo publicado, "REAL".
  * El ROL siempre se busca con COMUNA + manzana + predio.
  * "Oportunidad" es un indicador para REVISAR candidatas, no una tasación ni una promesa.
"""
import sys
import subprocess
import importlib


# ------------------------------------------------------------------------------
# 0. INSTALACIÓN AUTOMÁTICA DE BIBLIOTECAS
# ------------------------------------------------------------------------------
def instalar(paquete):
    base = [sys.executable, "-m", "pip", "install", "--quiet", paquete]
    for extra in ([], ["--break-system-packages"], ["--user"]):
        try:
            subprocess.check_call(base + extra, stdout=subprocess.DEVNULL)
            return True
        except Exception:
            continue
    return False


def asegurar_bibliotecas():
    necesarias = [("pandas", "pandas"), ("requests", "requests"), ("openpyxl", "openpyxl"),
                  ("docx", "python-docx"), ("bs4", "beautifulsoup4")]
    # ^ precio_mercado.py (módulo complementario) reutiliza requests y bs4, ya asegurados aquí.
    for modulo, paquete in necesarias:
        try:
            importlib.import_module(modulo)
        except ImportError:
            print(f"  Instalando '{paquete}' (solo la primera vez)...")
            if not instalar(paquete):
                print(f"\nNO pude instalar '{paquete}'. Escribe en Terminal:\n"
                      f"    {sys.executable} -m pip install {paquete}\n"
                      f"y vuelve a ejecutar el programa.")
                sys.exit(1)
            importlib.invalidate_caches()


asegurar_bibliotecas()

import os
import re
import json
import time
import sqlite3
import zipfile
import unicodedata
import traceback
import datetime as dt
from pathlib import Path

import pandas as pd
import requests

# ---- Módulo complementario de precio de mercado (debe estar en la misma carpeta) ----
try:
    import precio_mercado
    PRECIO_MERCADO_DISPONIBLE = True
except ImportError:
    PRECIO_MERCADO_DISPONIBLE = False

# ------------------------------------------------------------------------------
# CONFIGURACIÓN  (aquí se cambian los criterios; no hace falta tocar nada más)
# ------------------------------------------------------------------------------
URL_TGR = "https://remates.tgr.cl/v1/getListaRematesActivos"

# ---- Fuentes que se barren en cada ejecución (True = sí) ----
FUENTES = {"TGR": True, "MACAL": True, "AVISOS": True, "PRECIO_MERCADO": True}
FUENTES_OBLIGATORIAS = set()        # NINGUNA detiene el programa (decisión explícita: el barrido
                                     # siempre debe llegar hasta el final, pase lo que pase).
FUENTES_CRITICAS_ALERTA = {"TGR"}   # si una de estas falla, no se detiene el programa, pero se
                                     # lanza una alerta fuerte en pantalla y (si está configurado
                                     # más abajo) un correo, porque sin ella el resultado del día
                                     # queda incompleto en lo más importante.
TGR_MAXIMO_INTENTOS = 3
TGR_ESPERA_ENTRE_INTENTOS_SEG = 3

# ---- Correo de alerta cuando falla una fuente crítica (opcional, gratis vía Gmail) ----
# Para activarlo: sacar una "contraseña de aplicación" en https://myaccount.google.com/apppasswords
# (necesita verificación en 2 pasos activada en la cuenta) y completar los 3 datos de abajo.
# Si se deja vacío, el aviso solo aparece en pantalla, no se manda correo.
CORREO_ALERTA_REMITENTE = ""
CORREO_ALERTA_CONTRASEÑA_APP = ""
CORREO_ALERTA_DESTINO = ""

# ---- TGR: la TGR NO publica la "postura mínima"; publica "tasacion" y "avaluo" ----
# Mientras no se confirme la regla real, el mínimo de la TGR se ESTIMA así (rotulado ESTIMADO):
TGR_MINIMO_FACTOR_SOBRE_TASACION = 2 / 3     # mínimo estimado = tasación x este factor

# ---- Selección de las mejores oportunidades ----
N_TOP = 30
# Cupos por zona: se toman PRIMERO las mejores de cada zona, y el resto de los 30 por ranking nacional.
#   "PROV:082" = provincia de Concepción | "REG:RM" = Región Metropolitana | "REG:05" = Región de Valparaíso
CUPOS = {"PROV:082": 2, "REG:RM": 2, "REG:05": 2}

# ---- Doble ranking (avalúo fiscal vs. precio de mercado) - decisión explícita con Rodrigo ----
# Ranking A: mejores por oportunidad vs. AVALÚO FISCAL (como siempre).
# Ranking B: mejores por oportunidad vs. PRECIO DE MERCADO (solo propiedades con ese dato).
# Comodín: el resto de los cupos, las mejores oportunidades de cualquiera de los dos (sin repetir).
# Si el Ranking B no llega a su cupo (poca data de mercado todavía), los cupos que le faltan se
# los "presta" al Ranking A, para que el consolidado final se acerque al máximo igual.
# La cuota por zona (CUPOS) se exige POR SEPARADO dentro de cada ranking (A y B cada uno con la suya).
N_TOP_FISCAL = 20
N_TOP_MERCADO = 20
N_TOP_COMODIN = 10
N_TOP_MAXIMO = N_TOP_FISCAL + N_TOP_MERCADO + N_TOP_COMODIN   # 50

# ---- Cuota mínima de la TGR (pedido explícito de Rodrigo, 2026-10-09) ----
# Si el ranking de arriba no junta por sí solo al menos este número de
# propiedades de la TGR, se "fuerzan" las mejores candidatas de la TGR que
# falten, reemplazando a las propiedades NO-TGR (o TGR sin causa) con peor
# oportunidad. El TOP nunca cambia de tamaño (sigue en 50 como máximo):
# solo se intercambian cupos.
MINIMO_TGR_EN_TOP = 10
# De esas (al menos) 10 de la TGR, al menos este número deben tener
# tribunal Y rol de causa (= "carpeta judicial" disponible para consultar
# en el Poder Judicial).
MINIMO_TGR_CON_CAUSA_EN_TOP = 5
NOMBRE_ZONA = {"PROV:082": "Provincia de Concepción", "REG:RM": "Región Metropolitana",
               "REG:05": "Región de Valparaíso"}
VALOR_MINIMO_REF = 20_000_000       # se descartan propiedades cuyo avalúo (referencia) sea menor, en pesos
OPORTUNIDAD_MIN_PCT = 1.0           # solo entran al ranking oportunidades >= este % (0 = mínimo igual al avalúo)
OPORTUNIDAD_MAX_PCT = 400.0         # sobre este % se considera dato dudoso y va a la hoja REVISAR
SOLO_VIGENTES = True                # solo remates con fecha de hoy en adelante
EXCLUIR_DESTINOS_TOP = {"Z", "L"}   # estacionamientos y bodegas sueltos no entran al top
EXIGIR_CRUCE_SII = True             # solo entran al top propiedades verificadas en el catastro SII
INCLUIR_DATOS_DEUDOR = False        # True = agrega nombre y RUT del dueño (datos de la TGR)

# ---- Macal ----
MACAL_BASE = "https://macal.cl"
MACAL_MAX_PAGINAS = 60              # tope de seguridad del listado
MACAL_PAUSA_SEG = 0.3               # pausa de cortesía entre consultas

# ---- Carpeta de avisos (.docx de resúmenes semanales) ----
AVISOS_CARPETA_NOMBRE = "avisos"

# Destinos del SII que se confirmaron con los datos. Los demás se muestran solo con su letra.
DESTINOS_CONFIRMADOS = {"H": "Habitacional", "L": "Bodega", "Z": "Estacionamiento"}

# ------------------------------------------------------------------------------
# RUTAS
# ------------------------------------------------------------------------------
try:
    BASE = Path(__file__).resolve().parent
except NameError:
    BASE = Path.cwd()
DATOS = BASE / "datos"
SALIDAS = BASE / "salidas"
HOME = Path.home()
CARPETAS_BUSQUEDA = [BASE, DATOS, HOME / "Downloads", HOME / "Descargas",
                     HOME / "Desktop", HOME / "Escritorio", HOME / "python"]

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
    "Accept": "application/json",
}

# ------------------------------------------------------------------------------
# TABLA DE COMUNAS (código SII  |  nombre  |  código Tesorería)
# El ROL de la TGR trae el código de TESORERÍA; el catastro usa el código SII.
# ------------------------------------------------------------------------------
COMUNAS = [
    ("01101", "ARICA", "001"),
    ("01106", "CAMARONES", "295"),
    ("01201", "IQUIQUE", "002"),
    ("01203", "PICA", "004"),
    ("01204", "POZO ALMONTE", "005"),
    ("01206", "HUARA", "003"),
    ("01208", "CAMINA", "296"),
    ("01210", "COLCHANE", "297"),
    ("01211", "ALTO HOSPICIO", "347"),
    ("01301", "PUTRE", "294"),
    ("01302", "GENERAL LAGOS", "293"),
    ("02101", "TOCOPILLA", "006"),
    ("02103", "MARIA ELENA", "298"),
    ("02201", "ANTOFAGASTA", "007"),
    ("02202", "TALTAL", "009"),
    ("02203", "MEJILLONES", "008"),
    ("02206", "SIERRA GORDA", "299"),
    ("02301", "CALAMA", "010"),
    ("02302", "OLLAGUE", "300"),
    ("02303", "SAN PEDRO DE ATACAMA", "301"),
    ("03101", "CHANARAL", "011"),
    ("03102", "DIEGO DE ALMAGRO", "012"),
    ("03201", "COPIAPO", "013"),
    ("03202", "CALDERA", "014"),
    ("03203", "TIERRA AMARILLA", "015"),
    ("03301", "VALLENAR", "016"),
    ("03302", "FREIRINA", "017"),
    ("03303", "HUASCO", "018"),
    ("03304", "ALTO DEL CARMEN", "302"),
    ("04101", "LA SERENA", "019"),
    ("04102", "LA HIGUERA", "020"),
    ("04103", "COQUIMBO", "021"),
    ("04104", "ANDACOLLO", "022"),
    ("04105", "VICUNA", "023"),
    ("04106", "PAIHUANO", "024"),
    ("04201", "OVALLE", "025"),
    ("04203", "MONTE PATRIA", "026"),
    ("04204", "PUNITAQUI", "027"),
    ("04205", "COMBARBALA", "029"),
    ("04206", "RIO HURTADO", "028"),
    ("04301", "ILLAPEL", "030"),
    ("04302", "SALAMANCA", "032"),
    ("04303", "LOS VILOS", "033"),
    ("04304", "CANELA", "031"),
    ("05101", "ISLA DE PASCUA", "041"),
    ("05201", "LA LIGUA", "059"),
    ("05202", "PETORCA", "055"),
    ("05203", "CABILDO", "056"),
    ("05204", "ZAPALLAR", "058"),
    ("05205", "PAPUDO", "057"),
    ("05301", "VALPARAISO", "034"),
    ("05302", "VINA DEL MAR", "037"),
    ("05303", "VILLA ALEMANA", "039"),
    ("05304", "QUILPUE", "038"),
    ("05305", "CASABLANCA", "040"),
    ("05306", "QUINTERO", "035"),
    ("05307", "PUCHUNCAVI", "036"),
    ("05308", "JUAN FERNANDEZ", "321"),
    ("05309", "CONCON", "340"),
    ("05401", "SAN ANTONIO", "042"),
    ("05402", "SANTO DOMINGO", "043"),
    ("05403", "CARTAGENA", "046"),
    ("05404", "EL TABO", "047"),
    ("05405", "EL QUISCO", "045"),
    ("05406", "ALGARROBO", "044"),
    ("05501", "QUILLOTA", "048"),
    ("05502", "NOGALES", "052"),
    ("05503", "HIJUELAS", "051"),
    ("05504", "LA CALERA", "050"),
    ("05505", "LA CRUZ", "049"),
    ("05506", "LIMACHE", "053"),
    ("05507", "OLMUE", "054"),
    ("05601", "SAN FELIPE", "060"),
    ("05602", "PANQUEHUE", "062"),
    ("05603", "CATEMU", "063"),
    ("05604", "PUTAENDO", "061"),
    ("05605", "SANTA MARIA", "064"),
    ("05606", "LLAY-LLAY", "065"),
    ("05701", "LOS ANDES", "066"),
    ("05702", "CALLE LARGA", "067"),
    ("05703", "SAN ESTEBAN", "069"),
    ("05704", "RINCONADA", "068"),
    ("06101", "RANCAGUA", "105"),
    ("06102", "MACHALI", "106"),
    ("06103", "GRANEROS", "107"),
    ("06104", "SAN FRANCISCO DE MOSTAZAL", "111"),
    ("06105", "DONIHUE", "112"),
    ("06106", "COLTAUCO", "113"),
    ("06107", "CODEGUA", "110"),
    ("06108", "PEUMO", "115"),
    ("06109", "LAS CABRAS", "116"),
    ("06110", "SAN VICENTE", "117"),
    ("06111", "PICHIDEGUA", "118"),
    ("06112", "RENGO", "121"),
    ("06113", "REQUINOA", "119"),
    ("06114", "OLIVAR", "120"),
    ("06115", "MALLOA", "122"),
    ("06116", "COINCO", "114"),
    ("06117", "QUINTA DE TILCOCO", "123"),
    ("06201", "SAN FERNANDO", "124"),
    ("06202", "CHIMBARONGO", "125"),
    ("06203", "NANCAGUA", "126"),
    ("06204", "PLACILLA", "127"),
    ("06205", "SANTA CRUZ", "128"),
    ("06206", "LOLOL", "129"),
    ("06207", "PALMILLA", "130"),
    ("06208", "PERALILLO", "131"),
    ("06209", "CHEPICA", "132"),
    ("06214", "PUMANQUE", "135"),
    ("06301", "PICHILEMU", "137"),
    ("06302", "NAVIDAD", "138"),
    ("06303", "LITUECHE", "136"),
    ("06304", "LA ESTRELLA", "139"),
    ("06305", "MARCHIGUE", "134"),
    ("06306", "PAREDONES", "133"),
    ("07101", "CURICO", "140"),
    ("07102", "TENO", "142"),
    ("07103", "ROMERAL", "141"),
    ("07104", "RAUCO", "143"),
    ("07105", "LICANTEN", "145"),
    ("07106", "VICHUQUEN", "146"),
    ("07107", "HUALANE", "144"),
    ("07108", "MOLINA", "147"),
    ("07109", "SAGRADA FAMILIA", "148"),
    ("07201", "TALCA", "150"),
    ("07202", "SAN CLEMENTE", "151"),
    ("07203", "PELARCO", "152"),
    ("07204", "RIO CLARO", "149"),
    ("07205", "PENCAHUE", "153"),
    ("07206", "MAULE", "154"),
    ("07207", "CUREPTO", "155"),
    ("07208", "CONSTITUCION", "157"),
    ("07209", "EMPEDRADO", "158"),
    ("07210", "SAN RAFAEL", "341"),
    ("07301", "LINARES", "159"),
    ("07302", "YERBAS BUENAS", "160"),
    ("07303", "COLBUN", "161"),
    ("07304", "LONGAVI", "162"),
    ("07305", "PARRAL", "164"),
    ("07306", "RETIRO", "165"),
    ("07309", "VILLA ALEGRE", "163"),
    ("07310", "SAN JAVIER", "156"),
    ("07401", "CAUQUENES", "166"),
    ("07402", "PELLUHUE", "320"),
    ("07403", "CHANCO", "167"),
    ("08101", "CHILLAN", "168"),
    ("08102", "PINTO", "169"),
    ("08103", "COIHUECO", "170"),
    ("08104", "QUIRIHUE", "172"),
    ("08105", "NINHUE", "174"),
    ("08106", "PORTEZUELO", "171"),
    ("08107", "COBQUECURA", "175"),
    ("08108", "TREHUACO", "173"),
    ("08109", "SAN CARLOS", "176"),
    ("08110", "NIQUEN", "177"),
    ("08111", "SAN FABIAN", "178"),
    ("08112", "SAN NICOLAS", "179"),
    ("08113", "BULNES", "180"),
    ("08114", "SAN IGNACIO", "181"),
    ("08115", "QUILLON", "182"),
    ("08116", "YUNGAY", "183"),
    ("08117", "PEMUCO", "184"),
    ("08118", "EL CARMEN", "185"),
    ("08119", "RANQUIL", "187"),
    ("08120", "COELEMU", "186"),
    ("08121", "CHILLAN VIEJO", "342"),
    ("08201", "CONCEPCION", "188"),
    ("08202", "PENCO", "191"),
    ("08203", "HUALQUI", "192"),
    ("08204", "FLORIDA", "193"),
    ("08205", "TOME", "190"),
    ("08206", "TALCAHUANO", "189"),
    ("08207", "CORONEL", "194"),
    ("08208", "LOTA", "195"),
    ("08209", "SANTA JUANA", "196"),
    ("08210", "SAN PEDRO DE LA PAZ", "343"),
    ("08211", "CHIGUAYANTE", "344"),
    ("08212", "HUALPEN", "346"),
    ("08301", "ARAUCO", "198"),
    ("08302", "CURANILAHUE", "197"),
    ("08303", "LEBU", "199"),
    ("08304", "LOS ALAMOS", "200"),
    ("08305", "CANETE", "201"),
    ("08306", "CONTULMO", "202"),
    ("08307", "TIRUA", "203"),
    ("08401", "LOS ANGELES", "204"),
    ("08402", "SANTA BARBARA", "205"),
    ("08403", "LAJA", "210"),
    ("08404", "QUILLECO", "206"),
    ("08405", "NACIMIENTO", "212"),
    ("08406", "NEGRETE", "213"),
    ("08407", "MULCHEN", "214"),
    ("08408", "QUILACO", "215"),
    ("08409", "YUMBEL", "207"),
    ("08410", "CABRERO", "208"),
    ("08411", "SAN ROSENDO", "211"),
    ("08412", "TUCAPEL", "209"),
    ("08413", "ANTUCO", "303"),
    ("08414", "ALTO BIOBIO", "349"),
    ("09101", "ANGOL", "216"),
    ("09102", "PUREN", "217"),
    ("09103", "LOS SAUCES", "218"),
    ("09104", "RENAICO", "219"),
    ("09105", "COLLIPULLI", "220"),
    ("09106", "ERCILLA", "221"),
    ("09107", "TRAIGUEN", "222"),
    ("09108", "LUMACO", "223"),
    ("09109", "VICTORIA", "224"),
    ("09110", "CURACAUTIN", "225"),
    ("09111", "LONQUIMAY", "226"),
    ("09201", "TEMUCO", "227"),
    ("09202", "VILCUN", "228"),
    ("09203", "FREIRE", "229"),
    ("09204", "CUNCO", "230"),
    ("09205", "LAUTARO", "231"),
    ("09206", "PERQUENCO", "233"),
    ("09207", "GALVARINO", "232"),
    ("09208", "NUEVA IMPERIAL", "234"),
    ("09209", "CARAHUE", "235"),
    ("09210", "SAAVEDRA", "236"),
    ("09211", "PITRUFQUEN", "237"),
    ("09212", "GORBEA", "238"),
    ("09213", "TOLTEN", "239"),
    ("09214", "LONCOCHE", "240"),
    ("09215", "VILLARRICA", "241"),
    ("09216", "PUCON", "242"),
    ("09217", "MELIPEUCO", "304"),
    ("09218", "CURARREHUE", "305"),
    ("09219", "TEODORO SCHMIDT", "306"),
    ("09220", "PADRE LAS CASAS", "345"),
    ("09221", "CHOLCHOL", "348"),
    ("10101", "VALDIVIA", "243"),
    ("10102", "MARIQUINA", "245"),
    ("10103", "LANCO", "249"),
    ("10104", "LOS LAGOS", "247"),
    ("10105", "FUTRONO", "248"),
    ("10106", "CORRAL", "244"),
    ("10107", "MAFIL", "246"),
    ("10108", "PANGUIPULLI", "250"),
    ("10109", "LA UNION", "251"),
    ("10110", "PAILLACO", "252"),
    ("10111", "RIO BUENO", "253"),
    ("10112", "LAGO RANCO", "254"),
    ("10201", "OSORNO", "255"),
    ("10202", "SAN PABLO", "257"),
    ("10203", "PUERTO OCTAY", "258"),
    ("10204", "PUYEHUE", "256"),
    ("10205", "RIO NEGRO", "259"),
    ("10206", "PURRANQUE", "260"),
    ("10207", "SAN JUAN DE LA COSTA", "307"),
    ("10301", "PUERTO MONTT", "261"),
    ("10302", "COCHAMO", "262"),
    ("10303", "PUERTO VARAS", "266"),
    ("10304", "FRESIA", "268"),
    ("10305", "FRUTILLAR", "269"),
    ("10306", "LLANQUIHUE", "267"),
    ("10307", "MAULLIN", "263"),
    ("10308", "LOS MUERMOS", "264"),
    ("10309", "CALBUCO", "265"),
    ("10401", "CASTRO", "270"),
    ("10402", "CHONCHI", "271"),
    ("10403", "QUEILEN", "272"),
    ("10404", "QUELLON", "273"),
    ("10405", "PUQUELDON", "274"),
    ("10406", "ANCUD", "277"),
    ("10407", "QUEMCHI", "278"),
    ("10408", "DALCAHUE", "279"),
    ("10410", "CURACO DE VELEZ", "276"),
    ("10415", "QUINCHAO", "275"),
    ("10501", "CHAITEN", "280"),
    ("10502", "HUALAIHUE", "308"),
    ("10503", "FUTALEUFU", "281"),
    ("10504", "PALENA", "282"),
    ("11101", "AYSEN", "285"),
    ("11102", "CISNES", "286"),
    ("11104", "GUAITECAS", "309"),
    ("11201", "CHILE CHICO", "287"),
    ("11203", "RIO IBANEZ", "288"),
    ("11301", "COCHRANE", "289"),
    ("11302", "OHIGGINS", "310"),
    ("11303", "TORTEL", "311"),
    ("11401", "COYHAIQUE", "284"),
    ("11402", "LAGO VERDE", "312"),
    ("12101", "NATALES", "291"),
    ("12103", "TORRES DEL PAINE", "313"),
    ("12202", "RIO VERDE", "314"),
    ("12204", "SAN GREGORIO", "315"),
    ("12205", "PUNTA ARENAS", "290"),
    ("12206", "LAGUNA BLANCA", "316"),
    ("12301", "PORVENIR", "292"),
    ("12302", "PRIMAVERA", "317"),
    ("12304", "TIMAUKEL", "318"),
    ("12401", "CABO DE HORNOS", "319"),
    ("13101", "SANTIAGO", "070"),
    ("13134", "SANTIAGO OESTE", "073"),
    ("13135", "SANTIAGO SUR", "084"),
    ("13159", "RECOLETA", "329"),
    ("13167", "INDEPENDENCIA", "330"),
    ("14107", "QUINTA NORMAL", "081"),
    ("14109", "MAIPU", "094"),
    ("14111", "PUDAHUEL", "082"),
    ("14113", "RENCA", "077"),
    ("14114", "QUILICURA", "079"),
    ("14127", "CONCHALI", "075"),
    ("14155", "LO PRADO", "325"),
    ("14156", "CERRO NAVIA", "324"),
    ("14157", "ESTACION CENTRAL", "328"),
    ("14158", "HUECHURABA", "334"),
    ("14166", "CERRILLOS", "333"),
    ("14201", "COLINA", "076"),
    ("14202", "LAMPA", "078"),
    ("14203", "TIL-TIL", "080"),
    ("14501", "TALAGANTE", "086"),
    ("14502", "ISLA DE MAIPO", "087"),
    ("14503", "EL MONTE", "089"),
    ("14504", "PENAFLOR", "085"),
    ("14505", "PADRE HURTADO", "339"),
    ("14601", "MELIPILLA", "088"),
    ("14602", "MARIA PINTO", "090"),
    ("14603", "CURACAVI", "083"),
    ("14604", "SAN PEDRO", "108"),
    ("14605", "ALHUE", "109"),
    ("15103", "PROVIDENCIA", "072"),
    ("15105", "NUNOA", "091"),
    ("15108", "LAS CONDES", "071"),
    ("15128", "LA FLORIDA", "093"),
    ("15132", "LA REINA", "092"),
    ("15151", "MACUL", "323"),
    ("15152", "PENALOLEN", "322"),
    ("15160", "VITACURA", "331"),
    ("15161", "LO BARNECHEA", "332"),
    ("16106", "SAN MIGUEL", "095"),
    ("16110", "LA CISTERNA", "096"),
    ("16131", "LA GRANJA", "097"),
    ("16153", "SAN RAMON", "326"),
    ("16154", "LA PINTANA", "327"),
    ("16162", "PEDRO AGUIRRE CERDA", "336"),
    ("16163", "SAN JOAQUIN", "335"),
    ("16164", "LO ESPEJO", "337"),
    ("16165", "EL BOSQUE", "338"),
    ("16301", "PUENTE ALTO", "100"),
    ("16302", "PIRQUE", "101"),
    ("16303", "SAN JOSE DE MAIPO", "102"),
    ("16401", "SAN BERNARDO", "098"),
    ("16402", "CALERA DE TANGO", "099"),
    ("16403", "BUIN", "103"),
    ("16404", "PAINE", "104"),
]
TES2SII = {t: s for s, _, t in COMUNAS}
NOMBRE2SII = {n: s for s, n, _ in COMUNAS}
SII2NOMBRE = {s: n for s, n, _ in COMUNAS}

GENERICAS = {"", "s/i", "s/d", "sin informacion", "sin direccion", "null", "none", "nan"}

LOG = []


def decir(msg=""):
    print(msg, flush=True)
    LOG.append(str(msg))


def alerta_fuerte(mensaje):
    """Alerta imposible de no ver en pantalla, para cuando falla una fuente
    marcada en FUENTES_CRITICAS_ALERTA. Nunca detiene el programa."""
    decir("")
    decir("!" * 70)
    decir("!!  ALERTA IMPORTANTE")
    decir("!" * 70)
    for renglon in str(mensaje).split("\n"):
        decir(f"!!  {renglon}")
    decir("!" * 70)
    decir("")
    enviar_correo_alerta(f"[AVISO] remates_v2.py: {mensaje.splitlines()[0]}", mensaje)


def enviar_correo_alerta(asunto, cuerpo):
    if not CORREO_ALERTA_REMITENTE or not CORREO_ALERTA_CONTRASEÑA_APP or not CORREO_ALERTA_DESTINO:
        decir("  (Correo de alerta no configurado - el aviso queda solo en pantalla. "
              "Ver CORREO_ALERTA_* en la sección CONFIGURACIÓN si se quiere activar.)")
        return False
    try:
        import smtplib
        from email.mime.text import MIMEText
        msg = MIMEText(cuerpo)
        msg["Subject"] = asunto
        msg["From"] = CORREO_ALERTA_REMITENTE
        msg["To"] = CORREO_ALERTA_DESTINO
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as servidor:
            servidor.login(CORREO_ALERTA_REMITENTE, CORREO_ALERTA_CONTRASEÑA_APP)
            servidor.sendmail(CORREO_ALERTA_REMITENTE, [CORREO_ALERTA_DESTINO], msg.as_string())
        decir(f"  Correo de alerta enviado a {CORREO_ALERTA_DESTINO}.")
        return True
    except Exception as e:
        decir(f"  No se pudo enviar el correo de alerta: {e}")
        return False


# ------------------------------------------------------------------------------
# UTILIDADES
# ------------------------------------------------------------------------------
def normalizar(texto):
    t = unicodedata.normalize("NFD", str(texto)).encode("ascii", "ignore").decode()
    t = t.upper().replace("'", "")
    return re.sub(r"\s+", " ", t).strip()


def direccion_valida(glosa):
    if glosa is None or (isinstance(glosa, float) and pd.isna(glosa)):
        return False
    t = normalizar(glosa).lower()
    return t not in GENERICAS and len(t) > 3 and re.search(r"[a-z]", t) is not None


def parsear_rol_11(rol):
    """'15004530019' -> ('150', 4530, 19): código Tesorería(3) + manzana(5) + predio(3)."""
    d = re.sub(r"\D", "", str(rol if rol is not None else ""))
    if not d or len(d) > 11:
        return None, None, None
    d = d.zfill(11)
    return d[:3], int(d[3:8]), int(d[8:])


def parsear_rol_formato(rf):
    """'LEBU-00454-009' -> ('LEBU', 454, 9). Admite nombres con guion (LLAY-LLAY)."""
    m = re.fullmatch(r"(.+)-(\d+)-(\d+)", str(rf if rf is not None else "").strip())
    if not m:
        return None, None, None
    return m.group(1).strip(), int(m.group(2)), int(m.group(3))


def pesos(n):
    try:
        return f"${int(n):,}".replace(",", ".")
    except Exception:
        return "-"


# ------------------------------------------------------------------------------
# 1. DESCARGA EN VIVO DE LA TGR
# ------------------------------------------------------------------------------
def descargar_tgr():
    errores = []
    try:
        r = requests.get(URL_TGR, headers=HEADERS, timeout=60)
        if r.status_code == 200:
            return r.json()
        errores.append(f"requests: la TGR respondió código {r.status_code}")
    except Exception as e:
        errores.append(f"requests: {e}")

    # Segundo intento: imitar un Chrome real (sirve si la TGR bloquea a los programas)
    try:
        try:
            from curl_cffi import requests as cr
        except ImportError:
            decir("  Instalando 'curl_cffi' (solo si la TGR bloquea el primer intento)...")
            instalar("curl_cffi")
            importlib.invalidate_caches()
            from curl_cffi import requests as cr
        r = cr.get(URL_TGR, headers=HEADERS, impersonate="chrome", timeout=60)
        if r.status_code == 200:
            return r.json()
        errores.append(f"curl_cffi: la TGR respondió código {r.status_code}")
    except Exception as e:
        errores.append(f"curl_cffi: {e}")

    raise RuntimeError("No pude leer los remates en vivo de la TGR.\n  - " + "\n  - ".join(errores) +
                       "\n(Por regla del programa NO se usan datos antiguos de respaldo.)")


def construir_tabla_tgr(payload):
    if not isinstance(payload, dict) or not isinstance(payload.get("data"), list):
        llaves = list(payload)[:10] if isinstance(payload, dict) else type(payload).__name__
        raise ValueError(f"La TGR respondió con un formato inesperado. Llaves: {llaves}")
    filas = []
    for r in payload["data"]:
        tes, man_rol, pred_rol = parsear_rol_11(r.get("rol"))
        nom_rf, man_rf, pred_rf = parsear_rol_formato(r.get("rolFormato"))
        # El ROL con formato puede traer manzana/predio completos; se prefiere ese.
        manzana = man_rf if man_rf is not None else man_rol
        predio = pred_rf if pred_rf is not None else pred_rol
        sii_por_codigo = TES2SII.get(tes) if tes else None
        sii_por_nombre = NOMBRE2SII.get(normalizar(nom_rf)) if nom_rf else None
        comuna_sii = sii_por_codigo or sii_por_nombre
        if sii_por_codigo and sii_por_nombre and sii_por_codigo != sii_por_nombre:
            comuna_estado = "INCONSISTENTE"
        elif comuna_sii:
            comuna_estado = "OK"
        else:
            comuna_estado = "NO RECONOCIDA"
        rol_consistente = None
        if man_rol is not None and man_rf is not None:
            rol_consistente = (man_rol, pred_rol) == (man_rf, pred_rf)
        fila = {
            "fecha_remate": r.get("fechaRemate"),
            "fecha_primera_publicacion": r.get("fechaPrimeraPublicacion"),
            "comuna_sii": comuna_sii,
            "comuna_propiedad": SII2NOMBRE.get(comuna_sii) if comuna_sii else (nom_rf or None),
            "comuna_estado": comuna_estado,
            "rol_formato": r.get("rolFormato"),
            "manzana": manzana,
            "predio": predio,
            "rol_consistente": rol_consistente,
            "direccion_tgr": r.get("direccionRol"),
            "tasacion": r.get("tasacion"),
            "avaluo_tgr": r.get("avaluo"),
            "tipo_deuda": r.get("tipoDeuda"),
            "tribunal": r.get("nombreJuzgado"),
            "comuna_tribunal": r.get("comunaJuzgado"),
            "direccion_tribunal": r.get("direccionJuzgado"),
            "expediente": f"{r.get('nroExpJud')}-{r.get('agnoExpjud')}",
            "cod_demanda": r.get("codDemanda"),
            "modalidad": modalidad_desde_texto(r.get("datosSubasta")),
            "rol_causa": (str(r.get("SK") or "").split("#")[1:2] or [None])[0],
            "fojas": r.get("conservadorFojas"),
            "fojas_numero": r.get("conservadorNumero"),
            "fojas_anio": r.get("conservadorAno"),
            "cbr": r.get("cbr"),
            "datos_subasta": r.get("datosSubasta"),
        }
        if INCLUIR_DATOS_DEUDOR:
            fila["nombre_dueno"] = r.get("nombreDuegno")
            fila["rut_dueno"] = r.get("rutDuegno")
        filas.append(fila)
    df = pd.DataFrame(filas)
    if df.empty:
        return df
    for c in ("tasacion", "avaluo_tgr"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ("fecha_remate", "fecha_primera_publicacion"):
        f = pd.to_datetime(df[c], errors="coerce", utc=True)
        df[c] = f.dt.tz_convert("America/Santiago").dt.tz_localize(None)
    df["manzana"] = pd.to_numeric(df["manzana"], errors="coerce").astype("Int64")
    df["predio"] = pd.to_numeric(df["predio"], errors="coerce").astype("Int64")
    return df


# ------------------------------------------------------------------------------
# 2. CATASTRO NACIONAL DEL SII  ->  BASE LOCAL (SQLite)
# ------------------------------------------------------------------------------
PATRON_ARCHIVO = re.compile(r"^BRORGA\d+([NA])_NAC_(\d{4})_(\d)$")
PATRON_ZIP = re.compile(r"^BRORGA\d+[NA]_NAC_.*\.zip$", re.IGNORECASE)


def descomprimir_si_hace_falta():
    """Si bajaste el catastro como .zip, lo abre solo y deja el archivo en la carpeta 'datos'."""
    for carpeta in CARPETAS_BUSQUEDA:
        if not carpeta.is_dir():
            continue
        for z in carpeta.glob("*.zip"):
            if not PATRON_ZIP.match(z.name):
                continue
            try:
                with zipfile.ZipFile(z) as zf:
                    for info in zf.infolist():
                        nombre = Path(info.filename).name
                        if "__MACOSX" in info.filename or nombre.startswith("._"):
                            continue
                        if not PATRON_ARCHIVO.match(nombre):
                            continue
                        destino = DATOS / nombre
                        if destino.exists() and destino.stat().st_size == info.file_size:
                            continue
                        DATOS.mkdir(exist_ok=True)
                        decir(f"  Descomprimiendo {z.name} ({info.file_size / 1e6:,.0f} MB)...")
                        with zf.open(info) as src, open(destino, "wb") as dst:
                            while True:
                                bloque = src.read(8 * 1024 * 1024)
                                if not bloque:
                                    break
                                dst.write(bloque)
            except Exception as e:
                decir(f"  Aviso: no pude abrir {z.name}: {e}")


def buscar_archivos_catastro():
    """Devuelve {'N': (Path, (año, semestre)), 'A': (...)} con la versión más reciente."""
    descomprimir_si_hace_falta()
    hallados = {}
    for carpeta in CARPETAS_BUSQUEDA:
        if not carpeta.is_dir():
            continue
        try:
            for p in carpeta.iterdir():
                m = PATRON_ARCHIVO.match(p.name)
                if m and p.is_file():
                    serie, anio, sem = m.group(1), int(m.group(2)), int(m.group(3))
                    version = (anio, sem)
                    if serie not in hallados or version > hallados[serie][1]:
                        hallados[serie] = (p, version)
        except PermissionError:
            continue
    return hallados


def _firma(p):
    return f"{p.name}|{p.stat().st_size}"


def preparar_catastro(forzar=False):
    """Devuelve la ruta de la base SQLite del catastro (la crea si no existe). None si no hay archivos."""
    hallados = buscar_archivos_catastro()
    if not hallados:
        return None
    version = max(v for _, v in hallados.values())
    db_path = DATOS / f"catastro_{version[0]}_{version[1]}.sqlite"
    firmas = {s: _firma(p) for s, (p, _) in hallados.items()}
    if db_path.exists() and not forzar:
        try:
            con = sqlite3.connect(db_path)
            meta = dict(con.execute("SELECT clave, valor FROM meta").fetchall())
            con.close()
            if meta.get("completo") == "1" and all(meta.get(f"firma_{s}") == f for s, f in firmas.items()):
                decir(f"  Catastro ya preparado: {db_path.name}")
                return db_path
        except Exception:
            pass
    construir_base(hallados, db_path, firmas)
    return db_path


def construir_base(hallados, db_path, firmas):
    DATOS.mkdir(exist_ok=True)
    tmp = db_path.with_suffix(".tmp")
    if tmp.exists():
        tmp.unlink()
    decir("  Preparando la base del catastro (se hace UNA sola vez; puede tardar unos minutos)...")
    con = sqlite3.connect(tmp)
    con.execute("PRAGMA journal_mode=OFF")
    con.execute("PRAGMA synchronous=OFF")
    con.execute("""CREATE TABLE predios(
        comuna TEXT, manzana INTEGER, predio INTEGER, direccion TEXT,
        avaluo INTEGER, exento INTEGER, destino TEXT, tipo TEXT, fuente TEXT)""")
    con.execute("CREATE TABLE meta(clave TEXT, valor TEXT)")
    t0 = time.time()
    total = 0
    for serie in ("N", "A"):
        if serie not in hallados:
            continue
        ruta = hallados[serie][0]
        with open(ruta, "r", encoding="latin-1") as f:
            primera = f.readline().rstrip("\n").split("|")
        n_campos = len(primera)
        esperado = (19, 20) if serie == "N" else (9, 10)
        if n_campos not in esperado:
            raise ValueError(f"El archivo {ruta.name} tiene {n_campos} columnas y esperaba {esperado}. "
                             f"El SII pudo cambiar el formato; envíame una captura de este mensaje.")
        # N: tipo U/R en la columna 15 ; A: tipo U/R en la columna 8
        usecols = [0, 1, 2, 3, 4, 5, 6, 15 if serie == "N" else 8]
        decir(f"  Leyendo {ruta.name} ({ruta.stat().st_size / 1e6:,.0f} MB)...")
        lector = pd.read_csv(ruta, sep="|", header=None, dtype=str, encoding="latin-1",
                             usecols=usecols, index_col=False, chunksize=500_000,
                             na_filter=False)
        for trozo in lector:
            trozo.columns = ["comuna", "manzana", "predio", "direccion", "avaluo", "exento", "destino", "tipo"]
            for c in ("manzana", "predio", "avaluo", "exento"):
                trozo[c] = pd.to_numeric(trozo[c], errors="coerce").fillna(0).astype("int64")
            trozo["direccion"] = trozo["direccion"].str.strip()
            trozo["comuna"] = trozo["comuna"].str.strip()
            trozo["fuente"] = serie
            con.executemany(
                "INSERT INTO predios VALUES (?,?,?,?,?,?,?,?,?)",
                trozo[["comuna", "manzana", "predio", "direccion", "avaluo", "exento",
                       "destino", "tipo", "fuente"]].itertuples(index=False, name=None))
            total += len(trozo)
            print(f"\r    {total:,} predios cargados...", end="", flush=True)
        print()
    decir("  Creando índice de búsqueda...")
    con.execute("CREATE INDEX ix_rol ON predios(comuna, manzana, predio)")
    for s, f in firmas.items():
        con.execute("INSERT INTO meta VALUES (?,?)", (f"firma_{s}", f))
    con.execute("INSERT INTO meta VALUES ('predios', ?)", (str(total),))
    con.execute("INSERT INTO meta VALUES ('completo', '1')")
    con.commit()
    con.close()
    if db_path.exists():
        db_path.unlink()
    tmp.rename(db_path)
    decir(f"  Base lista: {total:,} predios en {time.time() - t0:,.0f} segundos.")


def cruzar_con_catastro(df, db_path):
    """Agrega dirección y avalúo del SII a cada remate (por comuna SII + manzana + predio)."""
    columnas_nuevas = ["direccion_sii", "avaluo_sii", "exento_sii", "destino_sii", "tipo_sii", "fuente_sii"]
    for c in columnas_nuevas:
        df[c] = None
    llaves = df.loc[df["comuna_sii"].notna() & df["manzana"].notna() & df["predio"].notna(),
                    ["comuna_sii", "manzana", "predio"]].drop_duplicates()
    encontrados = pd.DataFrame()
    if len(llaves):
        con = sqlite3.connect(db_path)
        con.execute("CREATE TEMP TABLE k(comuna TEXT, manzana INTEGER, predio INTEGER)")
        con.executemany("INSERT INTO k VALUES (?,?,?)",
                        [(str(a), int(b), int(c)) for a, b, c in llaves.itertuples(index=False, name=None)])
        encontrados = pd.read_sql_query(
            """SELECT k.comuna AS comuna_sii, k.manzana, k.predio,
                      p.direccion AS direccion_sii, p.avaluo AS avaluo_sii, p.exento AS exento_sii,
                      p.destino AS destino_sii, p.tipo AS tipo_sii, p.fuente AS fuente_sii
               FROM k JOIN predios p
                 ON p.comuna = k.comuna AND p.manzana = k.manzana AND p.predio = k.predio""", con)
        con.close()
        encontrados = encontrados.drop_duplicates(subset=["comuna_sii", "manzana", "predio"])
    base = df.drop(columns=columnas_nuevas)
    if len(encontrados):
        base["manzana"] = base["manzana"].astype("Int64")
        base["predio"] = base["predio"].astype("Int64")
        encontrados["manzana"] = encontrados["manzana"].astype("Int64")
        encontrados["predio"] = encontrados["predio"].astype("Int64")
        out = base.merge(encontrados, on=["comuna_sii", "manzana", "predio"], how="left")
    else:
        out = base.copy()
        for c in columnas_nuevas:
            out[c] = None

    def estado(fila):
        if fila["comuna_estado"] == "NO RECONOCIDA":
            return "COMUNA NO RECONOCIDA"
        if pd.isna(fila["manzana"]) or pd.isna(fila["predio"]):
            return "SIN ROL (no informado ni deducible)"
        if pd.isna(fila["avaluo_sii"]):
            return "ROL NO ENCONTRADO EN CATASTRO"
        if fila["comuna_estado"] == "INCONSISTENTE":
            return "OK (REVISAR: código y nombre de comuna no coinciden)"
        return "OK"

    out["cruce_estado"] = out.apply(estado, axis=1)
    return out



# ------------------------------------------------------------------------------
# 3. ZONAS (región / provincia) A PARTIR DEL CÓDIGO SII DE LA COMUNA
# ------------------------------------------------------------------------------
REGIONES = {
    "01": "Tarapacá / Arica y Parinacota", "02": "Región de Antofagasta", "03": "Región de Atacama",
    "04": "Región de Coquimbo", "05": "Región de Valparaíso", "06": "Región de O'Higgins",
    "07": "Región del Maule", "08": "Región del Biobío", "09": "Región de La Araucanía",
    "10": "Región de Los Lagos", "11": "Región de Aysén", "12": "Región de Magallanes",
    "13": "Región Metropolitana", "14": "Región Metropolitana", "15": "Región Metropolitana",
    "16": "Región Metropolitana",
}
PROVINCIAS_CONOCIDAS = {"082": "Provincia de Concepción", "083": "Provincia de Arauco",
                        "084": "Provincia del Biobío", "053": "Provincia de Valparaíso",
                        "055": "Provincia de Quillota", "054": "Provincia de San Antonio",
                        "131": "Provincia de Santiago"}


def zona_de(cod_sii):
    """Devuelve (región, provincia, conjunto_de_zonas) para un código SII de comuna."""
    if not cod_sii or not isinstance(cod_sii, str) or len(cod_sii) < 3:
        return None, None, set()
    p2, p3 = cod_sii[:2], cod_sii[:3]
    region = REGIONES.get(p2)
    if p3 == "081":
        region = "Región de Ñuble"
    elif p3 == "101":
        region = "Región de Los Ríos"
    elif p3 in ("011", "013"):
        region = "Región de Arica y Parinacota"
    elif p3 == "012":
        region = "Región de Tarapacá"
    prov = PROVINCIAS_CONOCIDAS.get(p3, f"Provincia (código {p3})")
    zonas = {f"PROV:{p3}", f"REG:{p2}"}
    if p2 in ("13", "14", "15", "16"):
        zonas.add("REG:RM")
    return region, prov, zonas


# ------------------------------------------------------------------------------
# 4. UTILIDADES DE TEXTO Y NÚMEROS
# ------------------------------------------------------------------------------
MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
         "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}


def a_numero_chileno(s):
    """'83.952.908' -> 83952908 ; '7.785,48' -> 7785.48 ; '5.299' -> 5299"""
    s = str(s).strip().rstrip(".-").strip()
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(".", "")
    try:
        v = float(s)
        return int(v) if v == int(v) else v
    except ValueError:
        return None


def modalidad_desde_texto(texto):
    """REMOTO / PRESENCIAL / MIXTA-REVISAR / NO INDICADA, a partir del texto del remate."""
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return "NO INDICADA"
    t = normalizar(texto).lower()
    remoto = re.search(r"zoom|video ?conferencia|via remota|modalidad remota|forma remota|manera remota|"
                       r"remoto|en linea|plataforma|telematic|meet\.google|teams", t)
    presencial = re.search(r"(forma|manera|modo|modalidad|subasta|remate) presencial|presencialmente|"
                           r"en (las )?dependencias del tribunal|en la sala de remates|sala de audiencias", t)
    if remoto and presencial:
        return "MIXTA/REVISAR"
    if remoto:
        return "REMOTO"
    if presencial:
        return "PRESENCIAL"
    return "NO INDICADA"


# ------------------------------------------------------------------------------
# 5. VALOR DE LA UF  (en vivo desde mindicador.cl; nunca inventado)
# ------------------------------------------------------------------------------
_UF_CACHE = {}


def valor_uf(fecha):
    """Devuelve (valor, nota). nota: 'EXACTA' o 'UF DE HOY (aproximada)'. (None, motivo) si no se pudo."""
    if fecha is None or pd.isna(fecha):
        fecha = dt.datetime.now()
    fecha = pd.Timestamp(fecha).to_pydatetime()
    clave = fecha.strftime("%d-%m-%Y")
    if clave in _UF_CACHE:
        return _UF_CACHE[clave]
    resultado = (None, "no se pudo consultar la UF")
    try:
        r = requests.get(f"https://mindicador.cl/api/uf/{clave}", headers=HEADERS, timeout=30)
        if r.status_code == 200:
            serie = r.json().get("serie") or []
            if serie:
                resultado = (float(serie[0]["valor"]), "EXACTA")
        if resultado[0] is None:
            r = requests.get("https://mindicador.cl/api/uf", headers=HEADERS, timeout=30)
            if r.status_code == 200:
                serie = r.json().get("serie") or []
                for s in serie:
                    if str(s.get("fecha", ""))[:10] == fecha.strftime("%Y-%m-%d"):
                        resultado = (float(s["valor"]), "EXACTA")
                        break
                if resultado[0] is None and serie:
                    resultado = (float(serie[0]["valor"]), "UF DE HOY (aproximada)")
    except Exception as e:
        resultado = (None, f"no se pudo consultar la UF: {e}")
    _UF_CACHE[clave] = resultado
    return resultado


# ------------------------------------------------------------------------------
# 6. DIRECCIÓN  ->  ROL  (búsqueda en el catastro del SII)
# ------------------------------------------------------------------------------
TIPOS_VIA = {"CALLE", "CL", "AV", "AVDA", "AVENIDA", "PJE", "PSJE", "PSJ", "PASAJE", "CAMINO", "CAM",
             "RUTA", "DIAGONAL", "DIAG", "CALLEJON", "CJN", "PASEO", "POBL", "POBLACION", "PARCELA"}
MARCAS_UNIDAD = {"DP", "BX", "BD", "LC", "OF", "ES", "BO", "DPTO", "DEPTO", "EST"}     # departamento, bodega, etc.
MARCAS_LOTE = {"LT", "LO", "MZ", "M", "ST", "SI", "CA", "CS", "SITIO", "LOTE", "PARCELA", "KM", "MC", "MZA"}  # descriptores de loteo
ABREVIATURAS = {"GRAL": "GENERAL", "GRL": "GENERAL", "CDTE": "COMANDANTE", "PDTE": "PRESIDENTE", "CAP": "CAPITAN",
                "TTE": "TENIENTE", "ALM": "ALMIRANTE", "DR": "DOCTOR", "STA": "SANTA", "STO": "SANTO", "SN": "SAN",
                "SGTO": "SARGENTO", "MCAL": "MARISCAL", "CARD": "CARDENAL", "OBPO": "OBISPO", "AVDA": "AV",
                "PJE": "PJE", "PSJE": "PJE", "PASAJE": "PJE", "MED": "MEDIA"}
NUM_PALABRAS = {"UNO": "1", "UN": "1", "DOS": "2", "TRES": "3", "CUATRO": "4", "CINCO": "5", "SEIS": "6", "SIETE": "7",
                "OCHO": "8", "NUEVE": "9", "DIEZ": "10", "ONCE": "11", "DOCE": "12", "TRECE": "13", "CATORCE": "14",
                "QUINCE": "15", "PRIMERO": "1", "PRIMER": "1", "SEGUNDO": "2", "TERCERO": "3", "TERCER": "3"}
PALABRAS_COMUNES = {"GENERAL", "DOCTOR", "PRESIDENTE", "SANTA", "SANTO", "SAN", "NORTE", "SUR", "ORIENTE", "PONIENTE",
                    "LOS", "LAS", "DEL", "DE", "LA", "EL", "AV", "PJE", "CALLE", "CAPITAN", "COMANDANTE", "ALMIRANTE"}

RE_VIA_CON_MARCA = re.compile(
    r"(?i)\b(calle|avenida|avda\.?|av\.?|pasaje|psje\.?|pje\.?|camino|diagonal|callej[oó]n|paseo)\s+"
    r"([^,;]{1,55}?)\s*(?:N[º°o]\.?\s*|n[uú]mero\s+|#\s*)(\d[\d.]*)")
RE_VIA_SIN_MARCA = re.compile(
    r"(?i)\b(calle|avenida|avda\.?|av\.?|pasaje|psje\.?|pje\.?|camino|diagonal|callej[oó]n|paseo)\s+"
    r"([^,;\d]{1,55}?)\s+(\d[\d.]*)\b")
RE_UNIDAD = re.compile(r"(?i)\b(?:departamento|depto\.?|dpto\.?|dp\.?)\s*(?:n[º°o.]*\s*|n[uú]mero\s+)?"
                       r"((?:[A-Za-z]\s?-\s?)?\d{1,5}(?:\s?-\s?[A-Za-z]\b|[A-Za-z]\b)?)")
PALABRAS_PROPIEDAD = re.compile(r"(?i)inmueble|propiedad|departamento|depto|casa|sitio|lote|vivienda|"
                                r"bien ra[ií]z|parcela|local comercial|unidad")
PALABRAS_TRIBUNAL = re.compile(r"(?i)juzgado|tribunal|secretar[ií]a|domicilio")


def clave_calle(nombre):
    """Normaliza el nombre de una calle: sin tipo (PJE, AV...), abreviaturas expandidas,
    números en palabras como dígitos y plurales simplificados (Garrochas = Garrocha)."""
    toks = normalizar(nombre).replace(".", " ").split()
    while toks and toks[0] in TIPOS_VIA:
        toks = toks[1:]
    out = []
    for t in toks:
        t = ABREVIATURAS.get(t, t)
        t = NUM_PALABRAS.get(t, t)
        if len(t) > 4 and t.endswith("S") and not t.isdigit():
            t = t[:-1]
        out.append(t)
    return " ".join(out)


def nombres_compatibles(a, b):
    """Igual, o iguales salvo iniciales ('C GOYENECHEA' = 'CANDELARIA GOYENECHEA'), o uno contenido en el otro (>=2 palabras)."""
    if a == b:
        return True
    ta, tb = a.split(), b.split()
    if len(ta) == len(tb) and all(x == y or (len(x) == 1 and y.startswith(x)) or (len(y) == 1 and x.startswith(y))
                                  for x, y in zip(ta, tb)):
        return True
    corto, largo = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    return len(corto) >= 2 and all(w in largo for w in corto)


def parsear_unidad(txt):
    """'1207-C' / 'C-403' / '101 -A' -> ('1207','C') ; sin letra -> ('803','')"""
    nums = [x.lstrip("0") for x in re.findall(r"\d+", txt or "")]
    nums = [x for x in nums if x]
    if not nums:
        return None
    letras = "".join(re.findall(r"[A-Za-z]", txt or "")).upper()
    return (nums[0], letras[:2])


def extraer_direccion_texto(texto):
    """Busca una calle y número de la PROPIEDAD en un texto libre (descarta la dirección del tribunal).
    Devuelve dict(nombre, numero, unidad) o None."""
    for regex in (RE_VIA_CON_MARCA, RE_VIA_SIN_MARCA):
        for m in regex.finditer(texto):
            ctx = texto[max(0, m.start() - 110):m.start()]
            ult = None
            for t in PALABRAS_TRIBUNAL.finditer(ctx):
                ult = t.end()
            if ult is not None and not PALABRAS_PROPIEDAD.search(ctx[ult:]):
                continue            # es la dirección del tribunal, no de la propiedad
            numero = a_numero_chileno(m.group(3))
            nombre = clave_calle(m.group(2))
            if numero is None or not nombre:
                continue
            u = RE_UNIDAD.search(texto[m.end():m.end() + 80]) or RE_UNIDAD.search(texto[max(0, m.start() - 70):m.end() + 80])
            unidad = parsear_unidad(u.group(1)) if u else None
            return {"nombre": nombre, "numero": int(numero), "unidad": unidad}
    return None


def parsear_dir_catastro(direccion):
    """'SANTA ISABEL 171 199 DP 801' -> ('SANTA ISABEL', 171, 199, ({'801'}, '')). 'PJE 9 SUR 3486' -> ('9 SUR', 3486, 3486, None)"""
    toks = normalizar(direccion).replace(".", " ").replace(",", " ").split()
    unidad = None
    corte = len(toks)
    for i, t in enumerate(toks):
        if t in MARCAS_UNIDAD and i > 0:
            corte = i
            resto = " ".join(toks[i + 1:i + 5])
            nums = [x.lstrip("0") for x in re.findall(r"\d+", resto)]
            nums = [x for x in nums if x]
            letras = "".join(re.findall(r"\b([A-Z])\b", resto))
            unidad = (set(nums), letras[:2]) if nums else (set(), letras[:2])
            break
    for i, t in enumerate(toks[:corte]):
        if t in MARCAS_LOTE and i > 1 and any(x.isdigit() for x in toks[:i]):
            corte = i
            break
    cabeza = toks[:corte]
    while len(cabeza) > 1 and len(cabeza[-1]) == 1 and cabeza[-1].isalpha() and cabeza[-2].isdigit():
        cabeza = cabeza[:-1]            # '1141 A' -> '1141'
    nums = []
    while cabeza and cabeza[-1].isdigit() and len(nums) < 2:
        nums.insert(0, int(cabeza.pop()))
    if not nums:
        return None
    nombre = clave_calle(" ".join(cabeza))
    return nombre, nums[0], nums[-1], unidad


def buscar_rol_por_direccion(con, comuna_sii, dir_tx):
    """dir_tx = dict(nombre, numero, unidad). Devuelve (estado, candidatos[(manzana,predio,direccion,avaluo,destino)])."""
    palabras = [p for p in dir_tx["nombre"].split() if len(p) >= 3 and not p.isdigit()]
    utiles = [p for p in palabras if p not in PALABRAS_COMUNES] or palabras or dir_tx["nombre"].split()
    if not utiles:
        return "SIN DIRECCION UTIL", []
    ancla = max(utiles, key=len)
    if len(ancla) > 5 and ancla.endswith("S"):
        ancla = ancla[:-1]
    filas = con.execute(
        "SELECT manzana, predio, direccion, avaluo, destino FROM predios "
        "WHERE comuna=? AND REPLACE(direccion,'Ñ','N') LIKE ?", (comuna_sii, f"%{ancla}%")).fetchall()
    cands = []
    for man, pre, d, av, dest in filas:
        if pre >= 90000 or dest == "K":
            continue            # rol "matriz" del condominio (avalúo de todo el edificio): no es una unidad
        p = parsear_dir_catastro(d)
        if not p:
            continue
        nombre, n1, n2, unidad = p
        if nombres_compatibles(nombre, dir_tx["nombre"]) and n1 <= dir_tx["numero"] <= max(n1, n2):
            cands.append((man, pre, d, av, dest, unidad, nombre))
    exactos = [c for c in cands if c[6] == dir_tx["nombre"]]
    if exactos:
        cands = exactos             # si hay calles de nombre idéntico, se prefieren a las "parecidas"
    if dir_tx.get("unidad"):
        d, letras = dir_tx["unidad"]
        cands = [c for c in cands if c[5] is not None and d in c[5][0] and (not letras or letras == c[5][1])]
    else:
        sin_unidad = [c for c in cands if c[5] is None]
        if sin_unidad or not cands:
            cands = sin_unidad
        else:
            return "AMBIGUO (edificio: hay varias unidades)", [c[:5] for c in cands[:6]]
    if not cands:
        return "SIN COINCIDENCIA", []
    if len(cands) == 1:
        return "UNICO", [cands[0][:5]]
    return f"AMBIGUO ({len(cands)} candidatos)", [c[:5] for c in cands[:6]]


def resolver_roles_faltantes(df, db_path):
    """Para filas sin ROL pero con comuna y dirección, busca el ROL en el catastro.
    Solo se asigna cuando la coincidencia es ÚNICA; si hay dudas, se listan candidatos y NO se asigna."""
    for c in ("rol_origen", "rol_candidatos", "estado_busqueda_direccion"):
        if c not in df.columns:
            df[c] = None
    df.loc[df["manzana"].notna() & df["rol_origen"].isna(), "rol_origen"] = "INFORMADO POR LA FUENTE"
    if db_path is None:
        return df
    con = sqlite3.connect(db_path)
    n_ok = n_amb = n_no = 0
    for i, f in df[df["manzana"].isna() & df["comuna_sii"].notna() & (df["fuente"] != "AVISO")].iterrows():
        txt = f.get("direccion_texto_busqueda")
        dtx = extraer_direccion_texto(txt) if isinstance(txt, str) and txt else None
        if dtx is None and isinstance(f.get("direccion_tgr"), str):
            dtx = extraer_direccion_texto("calle " + f["direccion_tgr"] + " ")  # intento básico
        if dtx is None:
            df.at[i, "estado_busqueda_direccion"] = "NO SE PUDO LEER LA DIRECCION"
            n_no += 1
            continue
        estado, cands = buscar_rol_por_direccion(con, f["comuna_sii"], dtx)
        df.at[i, "estado_busqueda_direccion"] = estado
        if estado == "UNICO":
            df.at[i, "manzana"] = cands[0][0]
            df.at[i, "predio"] = cands[0][1]
            df.at[i, "rol_origen"] = "DEDUCIDO POR DIRECCION (verificar)"
            n_ok += 1
        else:
            df.at[i, "rol_candidatos"] = "; ".join(f"{m}-{p} ({d})" for m, p, d, _, _ in cands) or None
            n_amb += 1 if cands else 0
            n_no += 0 if cands else 1
    con.close()
    decir(f"  ROL deducido por dirección: {n_ok} únicos | {n_amb} ambiguos (no se asignan) | {n_no} sin coincidencia/lectura")
    df["manzana"] = pd.to_numeric(df["manzana"], errors="coerce").astype("Int64")
    df["predio"] = pd.to_numeric(df["predio"], errors="coerce").astype("Int64")
    return df


# ------------------------------------------------------------------------------
# 7. FUENTE MACAL  (barrido completo del sitio macal.cl)
# ------------------------------------------------------------------------------
RE_LINK_FICHA = re.compile(r"(/venta/[A-Za-z0-9%_\-./]+?/P\d+-\d+)")
ETIQUETAS_MACAL = ["Precio mínimo", "Tipo", "Superficie útil", "Superficie total", "Dormitorios", "Baños",
                   "Estado de ocupación", "Uso", "Rol avalúo", "Plazo de pago", "Vendedor",
                   "Fecha de Subasta", "Subasta", "Garantía Requerida"]


class LectorWeb:
    """Descarga páginas; si el sitio exige un navegador real, usa Playwright (se instala solo)."""

    def __init__(self):
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": HEADERS["User-Agent"], "Accept": "text/html,*/*"})
        self.navegador = None
        self.pw = None
        self.page = None

    def get(self, url):
        ultimo = None
        for intento in range(3):
            try:
                if self.page is not None:
                    self.page.goto(url, wait_until="networkidle", timeout=60000)
                    return self.page.content()
                r = self.s.get(url, timeout=45)
                if r.status_code == 200:
                    return r.text
                ultimo = f"código {r.status_code}"
            except Exception as e:
                ultimo = str(e)
            time.sleep(1.5 * (intento + 1))
        raise RuntimeError(f"No pude leer {url}: {ultimo}")

    def activar_navegador(self):
        if self.page is not None:
            return
        decir("  Macal pide un navegador real: instalando/abriendo Playwright (solo la primera vez)...")
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            instalar("playwright")
            importlib.invalidate_caches()
            from playwright.sync_api import sync_playwright
        subprocess.call([sys.executable, "-m", "playwright", "install", "chromium"],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        self.pw = sync_playwright().start()
        self.navegador = self.pw.chromium.launch(headless=True)
        self.page = self.navegador.new_page()

    def cerrar(self):
        try:
            if self.navegador:
                self.navegador.close()
            if self.pw:
                self.pw.stop()
        except Exception:
            pass


# ==============================================================================
# PODER JUDICIAL DE CHILE (oficinajudicialvirtual.pjud.cl) - "Consulta Unificada
# de Causas". Es PUBLICA, GRATUITA y no pide login ni captcha (verificado a
# mano el 07-10-2026; confirmado también contra las Condiciones de Uso del sitio,
# que no restringen este tipo de consulta para causas civiles no reservadas).
#
# Para cada propiedad del TOP del ranking (no todas las analizadas - solo las
# mejores, porque cada consulta real toma cerca de 30-40 segundos) que tenga
# tribunal + rol de causa, se entra al sitio y se lee: estado del procedimiento,
# etapa, y el link al PDF de "Texto Demanda" si está disponible.
#
# Mientras se está probando que todo funcione bien, LIMITE_PRUEBA_TOP limita la
# consulta a solo esta cantidad de propiedades (las primeras del ranking, las
# de mayor oportunidad). Cuando ya se confíe en que anda bien, cambiar este
# número a 50 (o a None para "todas las del TOP, sin límite") y listo - no hay
# que tocar nada más en el resto del programa.
# ==============================================================================
CARPETA_DEBUG_PJUD = DATOS.parent / "debug_pjud"

URL_PJUD_INICIO = "https://oficinajudicialvirtual.pjud.cl/indexN.php"

LIMITE_PRUEBA_TOP = 5

# ---- Holgura para que cargue la página de inicio del Poder Judicial ----
# La nube (GitHub Actions) es más lenta/compartida que un Mac, así que el
# sitio puede tardar más en mostrar el botón "Consulta causas". Ajustado el
# 2026-10-09 porque en la nube las 5 consultas fallaron por timeout.
PJUD_TIMEOUT_BOTON_MS = 45_000      # antes: 25.000 ms
PJUD_TIMEOUT_CARGA_INICIAL_MS = 90_000   # antes: 60.000 ms
PJUD_REINTENTOS_PAGINA_INICIO = 4   # antes: 3
PJUD_PAUSA_ENTRE_REINTENTOS_MS = 6_000   # antes: 4.000 ms


_PJUD_ORDINALES_EN_PALABRAS = {
    "PRIMER": "1", "PRIMERO": "1",
    "SEGUNDO": "2",
    "TERCER": "3", "TERCERO": "3",
    "CUARTO": "4",
    "QUINTO": "5",
    "SEXTO": "6",
    "SEPTIMO": "7", "SETIMO": "7",
    "OCTAVO": "8",
    "NOVENO": "9",
    "DECIMO": "10",
}


def _pjud_normalizar(texto):
    """Quita tildes, pasa a mayúsculas y unifica símbolos de grado (° vs º),
    para poder comparar nombres de tribunal aunque estén escritos distinto.
    Además, convierte un ordinal escrito en palabras al inicio del texto
    (ej. 'Segundo Juzgado...') a su número (ej. '2 Juzgado...'), porque el
    sitio del Poder Judicial siempre usa el número, pero los datos de origen
    (TGR, avisos de diario) a veces traen el ordinal escrito en palabras -
    eso fue justo lo que pasó con 'Segundo Juzgado de Letras de Quilpué'."""
    if not texto:
        return ""
    texto = texto.replace("º", "°").replace("deg", "°")
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r"\s+", " ", texto).strip().upper()
    primera_palabra = texto.split(" ", 1)[0] if texto else ""
    if primera_palabra in _PJUD_ORDINALES_EN_PALABRAS:
        resto = texto[len(primera_palabra):].lstrip()
        texto = f"{_PJUD_ORDINALES_EN_PALABRAS[primera_palabra]} {resto}".strip()
    return texto


def _pjud_elegir_opcion_similar(select_locator, texto_buscado):
    """Un <select> HTML normal, pero no sabemos si el nombre del tribunal está
    escrito EXACTAMENTE igual en el menú desplegable (tildes, º/°, etc.). Esta
    función mira todas las opciones disponibles y elige la más parecida al
    texto que buscamos, en vez de exigir una coincidencia exacta."""
    opciones = select_locator.locator("option").all_text_contents()
    objetivo = _pjud_normalizar(texto_buscado)
    mejor = None
    for op in opciones:
        if _pjud_normalizar(op) == objetivo:
            mejor = op
            break
    if mejor is None:
        for op in opciones:
            if objetivo in _pjud_normalizar(op) or _pjud_normalizar(op) in objetivo:
                mejor = op
                break
    if mejor is None:
        raise ValueError(
            f"No encontré una opción parecida a '{texto_buscado}' en el menú. "
            f"Opciones disponibles (primeras 15): {opciones[:15]}"
        )
    select_locator.select_option(label=mejor)
    return mejor


def _pjud_campo_despues_de(pagina, etiqueta, tipo="select"):
    """Busca el campo (select o input) que viene justo después del texto de su
    etiqueta (ej. 'Competencia', 'Rol'), sin depender de ids internos que no
    conocemos. Para <input>, además ignora los que estén escondidos
    (type=hidden), o que sean radio/checkbox de OTRA sección de la página que
    no se ve pero que igual existe en el HTML."""
    etiqueta_loc = pagina.locator(f"text={etiqueta}").first
    if tipo == "input":
        xpath = ("xpath=following::input["
                 "not(@type='hidden') and not(@type='radio') and not(@type='checkbox')]")
    else:
        xpath = f"xpath=following::{tipo}"
    candidatos = etiqueta_loc.locator(xpath)
    total = candidatos.count()
    for i in range(total):
        c = candidatos.nth(i)
        if c.is_visible():
            return c
    return candidatos.first


def _pjud_consultar_una_causa(pw, tribunal, rol, anio):
    """Devuelve un diccionario con el resultado de consultar UNA causa. Abre un
    NAVEGADOR COMPLETAMENTE NUEVO (no solo una pestaña) para esta causa, y lo
    cierra al terminar - el sitio frena las consultas seguidas desde el mismo
    navegador, y un navegador nuevo cada vez imita mejor a una persona distinta
    entrando a consultar una causa puntual."""
    CARPETA_DEBUG_PJUD.mkdir(exist_ok=True)
    resultado = {
        "tribunal_buscado": tribunal, "rol": rol, "anio": anio,
        "encontrada": False, "error": "", "estado_proc": "", "etapa": "",
        "caratulado": "", "historia_resumen": "", "link_demanda_pdf": "",
    }
    navegador = pw.chromium.launch(
        headless=True,
        args=["--disable-blink-features=AutomationControlled"],
    )
    contexto = navegador.new_context(
        user_agent=("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36"),
        viewport={"width": 1440, "height": 900},
        locale="es-CL",
    )
    contexto.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    pagina = contexto.new_page()
    pagina_trabajo = pagina
    try:
        pagina.goto(URL_PJUD_INICIO, wait_until="domcontentloaded", timeout=PJUD_TIMEOUT_CARGA_INICIAL_MS)
        pagina.wait_for_timeout(3000)

        boton_listo = False
        ultimo_error = None
        for intento in range(1, PJUD_REINTENTOS_PAGINA_INICIO + 1):
            try:
                pagina.wait_for_selector("text=Consulta causas", timeout=PJUD_TIMEOUT_BOTON_MS)
                boton_listo = True
                break
            except Exception as e:
                ultimo_error = e
                pagina.screenshot(path=str(CARPETA_DEBUG_PJUD / f"0_inicio_lento_intento{intento}_{rol}.png"))
                pagina.wait_for_timeout(PJUD_PAUSA_ENTRE_REINTENTOS_MS)
                pagina.goto(URL_PJUD_INICIO, wait_until="domcontentloaded", timeout=PJUD_TIMEOUT_CARGA_INICIAL_MS)
                pagina.wait_for_timeout(3000)
        if not boton_listo:
            resultado["error"] = (
                f"La página de inicio no cargó el botón 'Consulta causas' tras "
                f"{PJUD_REINTENTOS_PAGINA_INICIO} intentos. Último error: {ultimo_error}"
            )
            return resultado

        pestanas_antes = len(contexto.pages)
        pagina.get_by_text("Consulta causas", exact=False).first.click()
        pagina.wait_for_timeout(3000)
        if len(contexto.pages) > pestanas_antes:
            pagina_trabajo = contexto.pages[-1]
            pagina_trabajo.wait_for_timeout(2000)
        pagina_trabajo.screenshot(path=str(CARPETA_DEBUG_PJUD / f"1_entrada_{rol}.png"))

        try:
            pagina_trabajo.wait_for_selector("text=Competencia", timeout=20000)
        except Exception:
            pagina_trabajo.screenshot(path=str(CARPETA_DEBUG_PJUD / f"1b_formulario_no_aparecio_{rol}.png"))
            resultado["error"] = (
                "El formulario de búsqueda no apareció después de 20 segundos "
                "(pantalla en blanco). Revisar captura 1b_formulario_no_aparecio."
            )
            return resultado

        sel_competencia = _pjud_campo_despues_de(pagina_trabajo, "Competencia")
        sel_competencia.select_option(label="Civil", timeout=15000)
        pagina_trabajo.wait_for_timeout(1000)

        sel_corte = _pjud_campo_despues_de(pagina_trabajo, "Corte")
        _pjud_elegir_opcion_similar(sel_corte, "Todos")

        # ---- Al elegir Corte="Todos", el sitio recarga por AJAX la lista de
        # TODOS los tribunales del país en el menú "Tribunal" - esa lista es
        # grande y a veces tarda varios segundos en llegar completa. Si
        # buscamos el tribunal apenas se selecciona "Todos", puede que el
        # menú todavía tenga cargada solo una lista parcial (ej. solo los
        # tribunales del norte) y entonces no lo encuentre, aunque el
        # tribunal sí exista en la lista completa. Por eso reintentamos
        # varias veces con pausas, en vez de intentarlo una sola vez. ----
        sel_tribunal = _pjud_campo_despues_de(pagina_trabajo, "Tribunal")
        ultimo_error_tribunal = None
        for intento_tribunal in range(1, 6):
            try:
                _pjud_elegir_opcion_similar(sel_tribunal, tribunal)
                ultimo_error_tribunal = None
                break
            except ValueError as e:
                ultimo_error_tribunal = e
                pagina_trabajo.wait_for_timeout(2000)
                sel_tribunal = _pjud_campo_despues_de(pagina_trabajo, "Tribunal")
        if ultimo_error_tribunal is not None:
            raise ultimo_error_tribunal
        pagina_trabajo.wait_for_timeout(500)

        sel_libro = _pjud_campo_despues_de(pagina_trabajo, "Libro/Tipo")
        _pjud_elegir_opcion_similar(sel_libro, "C")

        campo_rol = _pjud_campo_despues_de(pagina_trabajo, "Rol", tipo="input")
        campo_rol.fill(str(rol))

        campo_anio = _pjud_campo_despues_de(pagina_trabajo, "Año", tipo="input")
        campo_anio.fill(str(anio))

        pagina_trabajo.screenshot(path=str(CARPETA_DEBUG_PJUD / f"2_formulario_lleno_{rol}.png"))

        pagina_trabajo.get_by_role("button", name="Buscar").click()
        pagina_trabajo.wait_for_timeout(3000)
        pagina_trabajo.screenshot(path=str(CARPETA_DEBUG_PJUD / f"3_resultado_busqueda_{rol}.png"))

        texto_pagina = pagina_trabajo.inner_text("body")
        if "Total de registros: 0" in texto_pagina or f"C-{rol}-{anio}" not in texto_pagina.replace(" ", ""):
            resultado["error"] = "No apareció ningún resultado para ese Rol/Año/Tribunal."
            return resultado

        fila = pagina_trabajo.locator("tr", has_text=str(rol)).first
        fila.locator("td").first.click()
        pagina_trabajo.wait_for_timeout(2500)
        pagina_trabajo.screenshot(path=str(CARPETA_DEBUG_PJUD / f"4_detalle_{rol}.png"))

        resultado["encontrada"] = True
        detalle_texto = pagina_trabajo.inner_text("body")

        m = re.search(r"Estado Proc\.?:?\s*(.+?)(?:\s{2,}|\s+Etapa:|\s+Tribunal:|$)", detalle_texto)
        if m:
            resultado["estado_proc"] = m.group(1).strip()
        m = re.search(r"\bEtapa:?\s*(.+?)(?:\s{2,}|\s+Tribunal:|$)", detalle_texto)
        if m:
            resultado["etapa"] = m.group(1).strip()

        filas_historia = pagina_trabajo.locator("table tr").all_text_contents()
        resultado["historia_resumen"] = " | ".join(
            re.sub(r"\s+", " ", f.strip()) for f in filas_historia[1:8] if f.strip()
        )

        try:
            with pagina_trabajo.expect_popup(timeout=5000) as popup_info:
                pagina_trabajo.get_by_text("Texto Demanda", exact=False).first.click()
            nueva_pagina = popup_info.value
            nueva_pagina.wait_for_timeout(1500)
            resultado["link_demanda_pdf"] = nueva_pagina.url
            nueva_pagina.close()
        except Exception:
            pass

    except Exception as e:
        resultado["error"] = f"{type(e).__name__}: {e}"
        try:
            pagina_trabajo.screenshot(path=str(CARPETA_DEBUG_PJUD / f"ERROR_{rol}.png"))
        except Exception:
            pass

    finally:
        try:
            contexto.close()
        except Exception:
            pass
        try:
            navegador.close()
        except Exception:
            pass

    return resultado


def _pjud_separar_rol_causa(rol_causa):
    """'C-1586-2025' -> ('1586', '2025'). Si no tiene el formato esperado,
    devuelve (None, None) y esa fila simplemente no se consulta."""
    if not rol_causa:
        return None, None
    m = re.match(r"^[A-Za-z]{1,2}-([\d.]+)-(\d{4})$", str(rol_causa).strip())
    if not m:
        return None, None
    return m.group(1).replace(".", ""), m.group(2)


def agregar_info_poder_judicial(df, decir=print, pausa_entre_causas_seg=20):
    """Para las propiedades del TOP (columna EN_TOP) que tengan 'tribunal' y
    'rol_causa', consulta su causa en el Poder Judicial (público, sin login) y
    agrega columnas nuevas al df: estado_causa_pjud, etapa_causa_pjud,
    link_demanda_pjud, error_causa_pjud.

    Se consulta SOLO el TOP (máximo ~50 propiedades), no las miles analizadas,
    porque cada consulta real toma cerca de 30-40 segundos. Si algo falla en
    una causa puntual, esa fila simplemente queda vacía y se sigue con las
    demás - nunca se cae el programa completo por esto."""
    for col in ("estado_causa_pjud", "etapa_causa_pjud", "link_demanda_pjud", "error_causa_pjud"):
        if col not in df.columns:
            df[col] = None

    candidatas = df[df.get("EN_TOP", False) == True]
    if "tribunal" in df.columns and "rol_causa" in df.columns:
        candidatas = candidatas[candidatas["tribunal"].notna() & candidatas["rol_causa"].notna()]
    else:
        decir("  [PODER_JUDICIAL] el df no tiene columnas 'tribunal'/'rol_causa' - no hay nada que consultar.")
        return df

    if "oportunidad_pct" in candidatas.columns:
        candidatas = candidatas.sort_values(["oportunidad_pct"], ascending=False)

    if LIMITE_PRUEBA_TOP is not None:
        candidatas = candidatas.head(LIMITE_PRUEBA_TOP)
        decir(f"  [PODER_JUDICIAL] LIMITE_PRUEBA_TOP={LIMITE_PRUEBA_TOP}: consultando solo las primeras "
              f"{len(candidatas)} del ranking (sube este número, al inicio del archivo, cuando quieras "
              f"probar con más, o con las 50).")

    total = len(candidatas)
    if total == 0:
        decir("  [PODER_JUDICIAL] ninguna propiedad del TOP tiene tribunal + rol de causa para consultar.")
        return df

    decir("  [PODER_JUDICIAL] instalando/abriendo Playwright si falta (solo la primera vez)...")
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        instalar("playwright")
        importlib.invalidate_caches()
        from playwright.sync_api import sync_playwright
    subprocess.call([sys.executable, "-m", "playwright", "install", "chromium"],
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    decir(f"  [PODER_JUDICIAL] consultando {total} causas en oficinajudicialvirtual.pjud.cl "
          f"(con pausas entre cada una - esto puede tardar bastante)...")

    with sync_playwright() as pw:
        hechas = 0
        for i, fila in candidatas.iterrows():
            rol, anio = _pjud_separar_rol_causa(fila["rol_causa"])
            if not rol:
                continue
            hechas += 1
            decir(f"    [PODER_JUDICIAL] ({hechas}/{total}) {fila['tribunal']} - {fila['rol_causa']}...")
            r = _pjud_consultar_una_causa(pw, fila["tribunal"], rol, anio)
            if r["encontrada"]:
                df.at[i, "estado_causa_pjud"] = r["estado_proc"]
                df.at[i, "etapa_causa_pjud"] = r["etapa"]
                df.at[i, "link_demanda_pjud"] = r["link_demanda_pdf"]
                decir(f"      -> OK: estado='{r['estado_proc']}', etapa='{r['etapa']}'")
            else:
                df.at[i, "error_causa_pjud"] = r["error"] or "no encontrada"
                decir(f"      -> NO: {r['error'] or 'no encontrada'}")
            if hechas < total:
                time.sleep(pausa_entre_causas_seg)

    encontradas = df["estado_causa_pjud"].notna().sum()
    decir(f"  [PODER_JUDICIAL] OK: {encontradas} de {total} causas consultadas con éxito.")
    return df


def macal_links_desde_html(html):
    vistos, orden = set(), []
    for m in RE_LINK_FICHA.finditer(html.replace("\\u002F", "/").replace("\\/", "/")):
        u = m.group(1)
        if u not in vistos:
            vistos.add(u)
            orden.append(u)
    return orden


def macal_descubrir_fichas(lector):
    """Recorre el listado (?pagina=1,2,3...) hasta que no aparezcan fichas nuevas."""
    todas, vistos = [], set()
    for pag in range(1, MACAL_MAX_PAGINAS + 1):
        html = lector.get(f"{MACAL_BASE}/venta?pagina={pag}")
        links = macal_links_desde_html(html)
        if not links and pag == 1 and lector.page is None:
            lector.activar_navegador()
            html = lector.get(f"{MACAL_BASE}/venta?pagina={pag}")
            links = macal_links_desde_html(html)
        nuevos = [u for u in links if u not in vistos]
        if not nuevos:
            break
        for u in nuevos:
            vistos.add(u)
            todas.append(u)
        time.sleep(MACAL_PAUSA_SEG)
    return todas


def _texto_lineas(html):
    from bs4 import BeautifulSoup
    sopa = BeautifulSoup(html, "html.parser")
    for t in sopa(["script", "style", "noscript"]):
        t.decompose()
    lineas = [re.sub(r"\s+", " ", x).strip() for x in sopa.get_text("\n").split("\n")]
    return [x for x in lineas if x], (sopa.find("h1").get_text(" ", strip=True) if sopa.find("h1") else None)


def _valor_etiqueta(lineas, etiqueta):
    """Busca 'Etiqueta: valor' en una línea, o 'Etiqueta' en una línea y el valor en la siguiente."""
    e = normalizar(etiqueta)
    for i, l in enumerate(lineas):
        nl = normalizar(l)
        if nl.startswith(e):
            resto = l[len(etiqueta):].lstrip(" :-–").strip() if normalizar(l[:len(etiqueta)]) == e else ""
            if resto:
                return resto
            if i + 1 < len(lineas):
                return lineas[i + 1]
    return None


def parsear_precio(txt):
    """'UF 450' -> (450, 'UF') ; '$ 45.000.000' -> (45000000, 'CLP')"""
    if not txt:
        return None, None
    m = re.search(r"(?i)UF\s*([\d.]+(?:,\d+)?)", txt)
    if m:
        return a_numero_chileno(m.group(1)), "UF"
    m = re.search(r"\$\s*([\d.]+)", txt)
    if m:
        return a_numero_chileno(m.group(1)), "CLP"
    return None, None


def parsear_ficha_macal(html, url):
    lineas, h1 = _texto_lineas(html)
    texto = "\n".join(lineas)
    r = {"url": MACAL_BASE + url if url.startswith("/") else url}
    m = re.search(r"(P\d+-\d+)\s*$", url)
    r["id_fuente"] = m.group(1) if m else url
    r["titulo"] = h1 or (lineas[0] if lineas else None)
    ub = re.search(r"([^\n,]{3,40}),\s*Regi[oó]n\s+([^\n]{1,40})", texto)
    r["comuna_texto"] = ub.group(1).strip() if ub else None
    r["region_texto"] = ub.group(2).strip() if ub else None
    v = {e: _valor_etiqueta(lineas, e) for e in ETIQUETAS_MACAL}
    r["minimo_valor"], r["minimo_unidad"] = parsear_precio(v["Precio mínimo"])
    r["tipo_propiedad"] = v["Tipo"]
    r["superficie_util_m2"] = a_numero_chileno(re.sub(r"[^\d.,]", "", v["Superficie útil"] or "")) if v["Superficie útil"] else None
    r["superficie_total_m2"] = a_numero_chileno(re.sub(r"[^\d.,]", "", v["Superficie total"] or "")) if v["Superficie total"] else None
    r["dormitorios"] = v["Dormitorios"]
    r["banos"] = v["Baños"]
    r["ocupacion"] = v["Estado de ocupación"]
    r["plazo_pago"] = v["Plazo de pago"]
    r["vendedor"] = v["Vendedor"]
    fs = v["Fecha de Subasta"] or v["Subasta"]
    fecha = None
    if fs:
        mm = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", fs)
        if mm:
            try:
                fecha = dt.datetime(int(mm.group(3)), int(mm.group(2)), int(mm.group(1)))
            except ValueError:
                fecha = None
    r["fecha_remate"] = fecha
    r["garantia_pesos"], _ = parsear_precio(v["Garantía Requerida"])
    roles = []
    for a, b in re.findall(r"(\d{1,5})\s*-\s*(\d{1,4})", v["Rol avalúo"] or ""):
        roles.append((int(a), int(b)))
    r["roles"] = roles
    return r


def descargar_macal():
    lector = LectorWeb()
    try:
        links = macal_descubrir_fichas(lector)
        if not links:
            raise RuntimeError("No encontré ninguna ficha en el listado de Macal (¿cambió el sitio?)")
        decir(f"  Macal: {len(links)} fichas encontradas en el listado. Leyendo cada una...")
        fichas, fallos = [], []
        for k, u in enumerate(links, 1):
            try:
                html = lector.get(MACAL_BASE + u)
                f = parsear_ficha_macal(html, u)
                if f["minimo_valor"] is None and lector.page is None and k == 1:
                    lector.activar_navegador()
                    f = parsear_ficha_macal(lector.get(MACAL_BASE + u), u)
                fichas.append(f)
            except Exception as e:
                fallos.append((u, str(e)))
            if k % 25 == 0:
                print(f"\r    {k}/{len(links)} fichas...", end="", flush=True)
            time.sleep(MACAL_PAUSA_SEG)
        print()
        if fallos:
            decir(f"  AVISO Macal: {len(fallos)} fichas no se pudieron leer (se informan en RESUMEN).")
        return fichas, fallos
    finally:
        lector.cerrar()


def tabla_macal(fichas):
    filas = []
    for f in fichas:
        nom = normalizar(f["comuna_texto"]) if f.get("comuna_texto") else None
        sii = NOMBRE2SII.get(nom) if nom else None
        titulo = f.get("titulo") or ""
        busq = titulo if re.match(r"(?i)\s*(calle|avenida|avda|av\.?|pasaje|psje|pje|camino|diagonal|paseo)\b", titulo) \
            else "calle " + titulo
        man, pre = (f["roles"][0] if f.get("roles") else (None, None))
        filas.append({
            "fuente": "MACAL", "id_fuente": f["id_fuente"], "url": f["url"], "fecha_remate": f["fecha_remate"],
            "comuna_sii": sii, "comuna_propiedad": SII2NOMBRE.get(sii) if sii else f.get("comuna_texto"),
            "comuna_estado": "OK" if sii else "NO RECONOCIDA",
            "manzana": man, "predio": pre,
            "rol_formato": f"{man}-{pre}" if man is not None else None,
            "direccion_tgr": titulo, "direccion_texto_busqueda": busq + " ",
            "minimo_valor": f["minimo_valor"], "minimo_unidad": f["minimo_unidad"],
            "modalidad": "NO INDICADA", "tribunal": f.get("vendedor"),
            "tipo_juicio": "REMATE PRIVADO (Macal)", "tipo_propiedad": f.get("tipo_propiedad"),
            "superficie_util_m2": f.get("superficie_util_m2"), "superficie_total_m2": f.get("superficie_total_m2"),
            "dormitorios": f.get("dormitorios"), "banos": f.get("banos"), "ocupacion": f.get("ocupacion"),
            "plazo_pago": f.get("plazo_pago"), "garantia_pesos": f.get("garantia_pesos"),
            "roles_texto": ", ".join(f"{a}-{b}" for a, b in f.get("roles", [])),
            "n_roles": len(f.get("roles", [])),
        })
    df = pd.DataFrame(filas)
    if not df.empty:
        df["manzana"] = pd.to_numeric(df["manzana"], errors="coerce").astype("Int64")
        df["predio"] = pd.to_numeric(df["predio"], errors="coerce").astype("Int64")
        df["fecha_remate"] = pd.to_datetime(df["fecha_remate"], errors="coerce")
    return df


# ------------------------------------------------------------------------------
# 8. FUENTE AVISOS  (.docx de resúmenes semanales de avisos de remate)
# ------------------------------------------------------------------------------
NOMBRES_COMUNA_NORM = sorted({n for _, n, _ in COMUNAS}, key=len, reverse=True)
_ALT_COMUNAS = "|".join(re.escape(n) for n in NOMBRES_COMUNA_NORM)
RE_MARCA_COMUNA = re.compile(
    r"(?:COMUNA(?: Y CIUDAD| DE LA CIUDAD)? DE|CIUDAD Y COMUNA DE|UBICAD[OA]S? EN(?: LA)?(?: CIUDAD| COMUNA)?(?: DE)?|"
    r"SITO EN|SITA EN|EN ESTA CIUDAD DE|, DE)\s+(" + _ALT_COMUNAS + r")\b")
RE_CBR = re.compile(r"(?:CONSERVADOR(?: DE BIENES RAICES)?|CBR) DE\s+(" + _ALT_COMUNAS + r")\b")
RE_ROL_AVALUO = re.compile(
    r"\bROLE?S?\s+(?:DE\s+)?(?:AVALUOS?\s+)?(?:FISCAL(?:ES)?\s+)?(?:N\s*|NUMERO\s*|NRO\.?\s*|NO\s+)?[:\s]*"
    r"((?:\d{1,5}\s*-\s*\d{1,4}\b[\s,;]*(?:Y\s+|E\s+)?)+)")
RE_PRIMERA_PALABRA_TRIB = r"(?:\d{1,2}\s*[ºO°]?|PRIMER|SEGUND|TERCER|CUART|QUINT|SEXT|SEPTIM|OCTAV|NOVEN|DECIM|VIGESIM)\w*"
RE_TRIBUNAL = re.compile(
    r"(?i)((?:\d{1,2}\s*[º°o]?\s*|(?:primer|segund|tercer|cuart|quint|sext|s[eé]ptim|octav|noven|d[eé]cim|vig[eé]sim)\w*\s+)?"
    r"juzgado\s+(?:civil|de\s+letras|letras|de\s+garant[ií]a|de\s+familia|de\s+cobranza)?[^,.;()“\"]{0,60})")
CORTES_TRIBUNAL = re.compile(r"(?i)\s+(ubicado|en causa|causa|rol|rematar|rematara|ordeno|ha |se |autos|que |dictada|sobre|"
                             r"juicio|domicilio|con domicilio|de fecha|fijado|ante)\b")
MESES_UP = {k.upper(): v for k, v in MESES.items()}
RE_FECHA = re.compile(r"(\d{1,2})\s+(?:DE\s+)?(" + "|".join(MESES_UP) + r")\s+(?:DEL?\s+)?(?:ANO\s+)?(\d{4})")
RE_HORA = re.compile(r"(?:(\d{1,2})\s*[:.]\s*(\d{2})\s*(?:HORAS|HRS|H\b))|(?:A LAS\s+(\d{1,2})\s*(?:HORAS|HRS))")
RE_MINIMO = re.compile(r"(?i)m[ií]nimo\b([^$]{0,90}?)(\$\s?[\d.]+(?:,\d+)?|UF\s?[\d.]+(?:,\d+)?)")
ORDINALES = {"PRIMER": 1, "PRIMERO": 1, "PRIMERA": 1, "1": 1, "SEGUNDO": 2, "SEGUNDA": 2, "2": 2,
             "TERCER": 3, "TERCERO": 3, "TERCERA": 3, "3": 3, "CUARTO": 4, "CUARTA": 4, "4": 4}


def cabecera_fecha(texto):
    m = re.fullmatch(r"\s*(\d{1,2})\s+DE\s+([A-ZÁÉÍÓÚÑ]+)\s*", texto.upper())
    if m and normalizar(m.group(2)).lower() in MESES:
        return int(m.group(1)), MESES[normalizar(m.group(2)).lower()]
    return None


def extraer_tribunal(texto):
    m = RE_TRIBUNAL.search(texto)
    if not m:
        if re.search(r"(?i)juez\s+(partidor|[aá]rbitro)|jueza\s+(partidora|[aá]rbitro)", texto):
            return "JUEZ PARTIDOR / ARBITRO"
        return None
    t = m.group(1)
    c = CORTES_TRIBUNAL.search(t)
    if c:
        t = t[:c.start()]
    return re.sub(r"\s+", " ", t).strip(" ,.;")


def extraer_fecha_hora(texto_norm, cab):
    """Devuelve (datetime|None, nota)."""
    fechas = [(m, int(m.group(1)), MESES_UP[m.group(2)], int(m.group(3))) for m in RE_FECHA.finditer(texto_norm)]
    for m in re.finditer(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", texto_norm):
        if 1 <= int(m.group(2)) <= 12:
            fechas.append((m, int(m.group(1)), int(m.group(2)), int(m.group(3))))
    fechas = sorted([f for f in fechas if 2024 <= f[3] <= 2035], key=lambda f: f[0].start())
    elegida = None
    if cab:
        for f in fechas:
            if (f[1], f[2]) == cab:
                elegida = f
                break
    if elegida is None:
        for f in fechas:
            pre = texto_norm[max(0, f[0].start() - 120):f[0].start()]
            if re.search(r"REMAT|SUBAST", pre):
                elegida = f
                break
    if elegida is None and fechas:
        elegida = fechas[0]
    if elegida is None:
        return None, "SIN FECHA EN EL TEXTO"
    m = elegida[0]
    h = RE_HORA.search(texto_norm[m.end():m.end() + 90])
    hh, mm = 0, 0
    nota = "FECHA SIN HORA"
    if h:
        if h.group(1):
            hh, mm = int(h.group(1)), int(h.group(2))
        else:
            hh = int(h.group(3))
        nota = "OK"
    try:
        return dt.datetime(elegida[3], elegida[2], elegida[1], hh, mm), nota
    except ValueError:
        return None, "FECHA INVALIDA"


def extraer_roles(texto_norm):
    roles = []
    for m in RE_ROL_AVALUO.finditer(texto_norm):
        for a, b in re.findall(r"(\d{1,5})\s*-\s*(\d{1,4})", m.group(1)):
            if len(b) == 4 and 1990 <= int(b) <= 2035:
                continue            # parece un año (rol de causa), no un predio
            par = (int(a), int(b))
            if par not in roles:
                roles.append(par)
    return roles


def extraer_minimo(texto):
    """Devuelve (valor, unidad, n_distintos). Ignora la garantía ('10% del mínimo, esto es $...')."""
    vistos = []
    for m in RE_MINIMO.finditer(texto):
        antes = texto[max(0, m.start() - 45):m.start()].lower()
        if re.search(r"10\s*%|diez por ciento|garant|consign|vale vista|dep[oó]sito", antes):
            continue
        if re.search(r"(?i)esto es|10\s*%|garant", m.group(1)) and "equivalente a la tasaci" not in m.group(1).lower():
            continue
        if re.search(r"\.\s+[A-ZÁÉÍÓÚ]", m.group(1)):
            continue
        val = m.group(2)
        unidad = "UF" if val.upper().startswith("UF") else "CLP"
        num = a_numero_chileno(re.sub(r"(?i)uf|\$", "", val))
        if num is not None and (num, unidad) not in vistos:
            vistos.append((num, unidad))
    if not vistos:
        return None, None, 0
    return vistos[0][0], vistos[0][1], len(vistos)


def extraer_tipo_juicio(texto):
    t = normalizar(texto)
    if re.search(r"PARTIDOR|PARTICION|ARBITRO|COMUNIDAD|LIQUIDACION DE (LA )?(COMUNIDAD|SOCIEDAD CONYUGAL)", t):
        return "PARTICIONAL"
    if re.search(r"QUIEBRA|LIQUIDACION CONCURSAL|CONCURSAL|INSOLVENCIA|REORGANIZACION", t):
        return "CONCURSAL"
    if re.search(r"EJECUTIV|DESPOSEIMIENTO|HIPOTECARI|PRENDARI|COBRO", t):
        return "CIVIL - EJECUTIVO"
    return "NO INDICADO"


def extraer_n_remate(texto):
    t = normalizar(texto)
    m = re.search(r"\b(PRIMER[OA]?|SEGUND[OA]|TERCER[OA]?|CUART[OA]|1|2|3|4)\s*[º°O]?\s+(?:REMATE|SUBASTA|LLAMADO)\b", t)
    if m:
        return ORDINALES.get(m.group(1))
    m = re.search(r"\b(?:REMATE|SUBASTA|LLAMADO)\s+(?:N\s*|NUMERO\s*)?[º°O]?\s*([1-4])\b", t)
    return int(m.group(1)) if m else None


def extraer_fojas(texto):
    m = re.search(r"(?i)\b(?:fojas|fs\.?)\s*([\d.]+)\s*(vta\.?|vuelta)?\s*(?:n[º°o.]*|n[uú]mero)\s*([\d.]+)", texto)
    if not m:
        return None, None, None, None
    anio = None
    ma = re.search(r"(?i)(?:a[ñn]o\s*)?\b(19\d{2}|20[0-3]\d)\b", texto[m.end():m.end() + 160])
    if ma:
        anio = int(ma.group(1))
    cbr = None
    mc = RE_CBR.search(normalizar(texto[m.end():m.end() + 260]))
    if mc:
        cbr = mc.group(1)
    return a_numero_chileno(m.group(1)), a_numero_chileno(m.group(3)), anio, cbr


def candidatos_comuna(texto):
    """Lista ordenada de (comuna_sii, origen). Prioriza lo que el aviso dice de la PROPIEDAD."""
    tn = normalizar(texto)
    cands = []

    def agregar(nombre, origen):
        sii = NOMBRE2SII.get(nombre)
        if sii and sii not in [c for c, _ in cands]:
            cands.append((sii, origen))

    for m in RE_MARCA_COMUNA.finditer(tn):
        ctx = tn[max(0, m.start() - 110):m.start()]
        ult = None
        for t in re.finditer(r"JUZGADO|TRIBUNAL|SECRETARIA|DOMICILIO", ctx):
            ult = t.end()
        if ult is not None and not re.search(r"INMUEBLE|PROPIEDAD|DEPARTAMENTO|CASA|SITIO|LOTE|VIVIENDA|BIEN RAIZ|PARCELA", ctx[ult:]):
            continue
        agregar(m.group(1), "TEXTO")
    for m in RE_ROL_AVALUO.finditer(tn):
        mm = re.match(r"[\s,.]*(?:COMUNA DE|DE LA COMUNA DE|DE)\s+(" + _ALT_COMUNAS + r")\b", tn[m.end():m.end() + 60])
        if mm:
            agregar(mm.group(1), "ROL")
    for m in RE_CBR.finditer(tn):
        agregar(m.group(1), "CBR")
    trib = extraer_tribunal(texto)
    if trib:
        for n in NOMBRES_COMUNA_NORM:
            if re.search(r"\b" + re.escape(n) + r"\b", normalizar(trib)):
                agregar(n, "TRIBUNAL")
                break
    return cands


def parsear_aviso(texto, cab=None):
    tn = normalizar(texto)
    fecha, nota_fecha = extraer_fecha_hora(tn, cab)
    val, uni, n_min = extraer_minimo(texto)
    roles = extraer_roles(tn)
    fojas, numero, anio, cbr = extraer_fojas(texto)
    mrc = re.search(r"(?i)\b(?:rol|causa)\s*(?:n[º°o.]*\s*)?([A-Z]{1,2})\s*[-–]\s*([\d.]+)\s*[-–]\s*(\d{4})", texto)
    rol_causa = f"{mrc.group(1).upper()}-{mrc.group(2).replace('.', '')}-{mrc.group(3)}" if mrc else None
    alertas = []
    if re.search(r"(?i)\b(derechos?\s+(?:inscritos|equivalentes|proindiviso|del?\s+\d)|cuota|proindiviso|\d+\s*/\s*\d+\s+avas?)", texto):
        alertas.append("DERECHOS_PARCIALES_O_CUOTA")
    if re.search(r"(?i)aprovechamiento de aguas?|derechos de agua", texto):
        alertas.append("INCLUYE_DERECHOS_DE_AGUA")
    if re.search(r"(?i)\b(estacionamiento|bodega)\b", texto) and re.search(r"(?i)\bdepartamento\b", texto):
        alertas.append("VARIAS_UNIDADES")
    if len(roles) > 1:
        alertas.append("VARIOS_ROLES")
    if n_min > 1:
        alertas.append("MULTIPLES_MINIMOS")
    if re.search(r"(?i)rural|parcela|hijuela|hect[aá]rea|\bfundo\b", texto):
        alertas.append("POSIBLE_PREDIO_RURAL")
    return {
        "tribunal": extraer_tribunal(texto), "modalidad": modalidad_desde_texto(texto),
        "fecha_remate": fecha, "nota_fecha": nota_fecha, "rol_causa": rol_causa,
        "tipo_juicio": extraer_tipo_juicio(texto), "n_remate": extraer_n_remate(texto),
        "fojas": fojas, "fojas_numero": numero, "fojas_anio": anio, "cbr": cbr,
        "minimo_valor": val, "minimo_unidad": uni, "roles": roles,
        "candidatos": candidatos_comuna(texto), "alertas_aviso": alertas,
    }


def leer_avisos_docx(ruta):
    import docx
    d = docx.Document(str(ruta))
    cab = None
    filas = []
    n = 0
    for p in d.paragraphs:
        tx = p.text.strip()
        if not tx:
            continue
        c = cabecera_fecha(tx)
        if c:
            cab = c
            continue
        if len(tx) < 150 or not re.search(r"(?i)remat|subast", tx):
            continue
        n += 1
        a = parsear_aviso(tx, cab)
        a.update({"texto": tx, "archivo": Path(ruta).name, "n_aviso": n})
        filas.append(a)
    return filas


def buscar_docx_avisos():
    carpeta = BASE / AVISOS_CARPETA_NOMBRE
    rutas = []
    if carpeta.is_dir():
        rutas += [p for p in carpeta.glob("*.docx") if not p.name.startswith("~$")]
    for c in CARPETAS_BUSQUEDA:          # también resúmenes sueltos llamados RESUMEN*.docx
        if c.is_dir() and c != carpeta:
            rutas += [p for p in c.glob("*RESUMEN*.docx") if not p.name.startswith("~$")]
    unicas = {}
    for p in rutas:
        unicas[p.resolve()] = p
    return sorted(unicas.values())


def tabla_avisos(rutas):
    filas = []
    for r in rutas:
        for a in leer_avisos_docx(r):
            filas.append({
                "fuente": "AVISO", "id_fuente": f"{a['archivo']}#{a['n_aviso']}", "url": None,
                "fecha_remate": a["fecha_remate"], "modalidad": a["modalidad"], "tribunal": a["tribunal"],
                "rol_causa": a["rol_causa"], "tipo_juicio": a["tipo_juicio"], "n_remate": a["n_remate"],
                "fojas": a["fojas"], "fojas_numero": a["fojas_numero"], "fojas_anio": a["fojas_anio"], "cbr": a["cbr"],
                "minimo_valor": a["minimo_valor"], "minimo_unidad": a["minimo_unidad"],
                "roles_lista": a["roles"], "candidatos_comuna": a["candidatos"],
                "alertas_fuente": "; ".join(a["alertas_aviso"]), "nota_fecha": a["nota_fecha"],
                "direccion_texto_busqueda": a["texto"], "texto_aviso": a["texto"],
                "roles_texto": ", ".join(f"{x}-{y}" for x, y in a["roles"]), "n_roles": len(a["roles"]),
                "manzana": None, "predio": None, "comuna_sii": None, "comuna_estado": "NO RECONOCIDA",
            })
    df = pd.DataFrame(filas)
    if not df.empty:
        df["fecha_remate"] = pd.to_datetime(df["fecha_remate"], errors="coerce")
    return df


def resolver_avisos(df, db_path):
    """Para cada aviso: elige la comuna y el ROL (explícito o deducido por dirección). Sin catastro: solo comuna."""
    if df.empty:
        return df
    for c in ("rol_origen", "rol_candidatos", "estado_busqueda_direccion", "verif_direccion_catastro",
              "avaluo_roles_total", "n_roles_encontrados"):
        if c not in df.columns:
            df[c] = None
    con = sqlite3.connect(db_path) if db_path else None
    for i, f in df[df["fuente"] == "AVISO"].iterrows():
        cands = f["candidatos_comuna"] or []
        roles = f["roles_lista"] or []
        dtx = extraer_direccion_texto(f["texto_aviso"])
        elegido = None
        if roles and con is not None:
            man, pre = roles[0]
            aciertos = []
            for sii, origen in cands:
                r = con.execute("SELECT direccion, avaluo FROM predios WHERE comuna=? AND manzana=? AND predio=?",
                                (sii, man, pre)).fetchone()
                if r:
                    ok_calle = None
                    if dtx:
                        p = parsear_dir_catastro(r[0])
                        ok_calle = bool(p and nombres_compatibles(p[0], dtx["nombre"]))
                    aciertos.append((sii, origen, ok_calle))
            elegido = next((a for a in aciertos if a[2]), None)
            if elegido is None and len(aciertos) == 1:
                elegido = aciertos[0]
            if elegido is None and aciertos and aciertos[0][1] in ("TEXTO", "ROL"):
                elegido = aciertos[0]
            if elegido:
                sii, origen, ok_calle = elegido
                df.at[i, "comuna_sii"], df.at[i, "manzana"], df.at[i, "predio"] = sii, man, pre
                df.at[i, "rol_origen"] = "INFORMADO POR LA FUENTE (aviso)"
                df.at[i, "verif_direccion_catastro"] = ("COINCIDE" if ok_calle else
                                                        "NO COINCIDE" if ok_calle is False else "SIN DATO")
                tot, enc = 0, 0
                for (m2, p2) in roles:
                    rr = con.execute("SELECT avaluo FROM predios WHERE comuna=? AND manzana=? AND predio=?",
                                     (sii, m2, p2)).fetchone()
                    if rr:
                        tot += rr[0]
                        enc += 1
                df.at[i, "avaluo_roles_total"] = tot if enc else None
                df.at[i, "n_roles_encontrados"] = enc
        if elegido is None:
            # sin ROL utilizable: intentar por dirección en cada comuna candidata
            if con is not None and dtx:
                for sii, origen in cands:
                    estado, cs = buscar_rol_por_direccion(con, sii, dtx)
                    if estado == "UNICO":
                        df.at[i, "comuna_sii"], df.at[i, "manzana"], df.at[i, "predio"] = sii, cs[0][0], cs[0][1]
                        df.at[i, "rol_origen"] = "DEDUCIDO POR DIRECCION (verificar)"
                        df.at[i, "estado_busqueda_direccion"] = "UNICO"
                        elegido = (sii, origen, None)
                        break
                    df.at[i, "estado_busqueda_direccion"] = estado
                    if cs:
                        df.at[i, "rol_candidatos"] = "; ".join(f"{m}-{p} ({d})" for m, p, d, _, _ in cs)
            elif dtx is None:
                df.at[i, "estado_busqueda_direccion"] = "NO SE PUDO LEER LA DIRECCION"
        if elegido is None and cands:
            sii, origen = cands[0]
            df.at[i, "comuna_sii"] = sii
            if origen == "TRIBUNAL":
                df.at[i, "estado_busqueda_direccion"] = (str(df.at[i, "estado_busqueda_direccion"] or "")
                                                          + " | COMUNA ASUMIDA DEL TRIBUNAL")
        sii_final = df.at[i, "comuna_sii"]
        if sii_final:
            df.at[i, "comuna_estado"] = "OK"
            df.at[i, "comuna_propiedad"] = SII2NOMBRE.get(sii_final)
        if dtx:
            df.at[i, "direccion_tgr"] = f"{dtx['nombre']} {dtx['numero']}" + (f" (unidad {dtx['unidad'][0]}{dtx['unidad'][1]})" if dtx["unidad"] else "")
        df.at[i, "rol_formato"] = (f"{df.at[i, 'manzana']}-{df.at[i, 'predio']}"
                                   if pd.notna(df.at[i, "manzana"]) else None)
    if con is not None:
        con.close()
    df["manzana"] = pd.to_numeric(df["manzana"], errors="coerce").astype("Int64")
    df["predio"] = pd.to_numeric(df["predio"], errors="coerce").astype("Int64")
    sub = df[df["fuente"] == "AVISO"]
    decir(f"  Avisos: {len(sub)} leídos | con ROL: {int(sub['manzana'].notna().sum())} "
          f"(informado: {int((sub['rol_origen'] == 'INFORMADO POR LA FUENTE (aviso)').sum())}, "
          f"deducido por dirección: {int((sub['rol_origen'] == 'DEDUCIDO POR DIRECCION (verificar)').sum())})")
    return df


# ------------------------------------------------------------------------------
# 9. UNIFICACIÓN, VALORES Y OPORTUNIDAD
# ------------------------------------------------------------------------------
BLOQUEANTES = {
    "DERECHOS_PARCIALES_O_CUOTA": "derechos parciales/cuota: el avalúo del predio completo no es comparable",
    "MULTIPLES_MINIMOS": "el aviso trae más de un mínimo (p.ej. inmueble y derechos de agua)",
    "UF_SIN_CONVERTIR": "no se pudo convertir el mínimo en UF a pesos",
    "ROLES_PARCIALES": "hay varios roles y no todos se encontraron en el catastro",
    "DIRECCION_NO_COINCIDE": "la calle del aviso no coincide con la del catastro para ese ROL",
    "RATIO_TASACION_AVALUO_ANORMAL": "tasación/avalúo de la TGR fuera de lo normal (1,3)",
    "ROL_ES_MATRIZ_DE_CONDOMINIO": "el ROL es la matriz del condominio (avalúo de todo el edificio), no una unidad",
    "OPORTUNIDAD_SOSPECHOSA": "ratio sobre el máximo plausible: probable error de dato",
}
PRIORIDAD_FUENTE = {"AVISO": 0, "MACAL": 1, "TGR": 2}


def cargar_valores_mercado():
    """Gancho para precios de mercado. Si existe datos/valores_mercado.csv
    (columnas: comuna_sii,manzana,predio,valor_mercado_pesos) se usa; si no, no hay valor de mercado.
    Aquí es donde se conectaría un proveedor automático (p.ej. Portal Inmobiliario o Data Inmobiliaria)."""
    ruta = DATOS / "valores_mercado.csv"
    if not ruta.exists():
        return None
    v = pd.read_csv(ruta, dtype={"comuna_sii": str})
    v["manzana"] = pd.to_numeric(v["manzana"], errors="coerce").astype("Int64")
    v["predio"] = pd.to_numeric(v["predio"], errors="coerce").astype("Int64")
    return v[["comuna_sii", "manzana", "predio", "valor_mercado_pesos"]].drop_duplicates(
        ["comuna_sii", "manzana", "predio"])


def aplicar_minimos(df):
    """Mínimo en pesos: REAL si la fuente lo publica (convierte UF con la UF en vivo); ESTIMADO para la TGR."""
    df["minimo_pesos"] = pd.NA
    df["minimo_nota"] = None
    df["tipo_minimo"] = "NO DISPONIBLE"
    df["minimo_pesos"] = df["minimo_pesos"].astype("object")
    for i, f in df.iterrows():
        val, uni = f.get("minimo_valor"), f.get("minimo_unidad")
        if pd.notna(val) and uni == "CLP":
            df.at[i, "minimo_pesos"] = float(val)
            df.at[i, "tipo_minimo"] = "REAL (publicado en la fuente)"
        elif pd.notna(val) and uni == "UF":
            uf, nota = valor_uf(f.get("fecha_remate"))
            df.at[i, "tipo_minimo"] = "REAL (publicado en la fuente)"
            if uf is None:
                df.at[i, "minimo_nota"] = "UF_SIN_CONVERTIR"
            else:
                df.at[i, "minimo_pesos"] = float(val) * uf
                df.at[i, "minimo_nota"] = "UF_APROXIMADA" if "aprox" in nota else "UF_EXACTA"
        elif f["fuente"] == "TGR" and pd.notna(f.get("tasacion")):
            df.at[i, "minimo_pesos"] = float(f["tasacion"]) * TGR_MINIMO_FACTOR_SOBRE_TASACION
            df.at[i, "tipo_minimo"] = "ESTIMADO (tasación x factor; la TGR no publica el mínimo)"
    df["minimo_pesos"] = pd.to_numeric(df["minimo_pesos"], errors="coerce")
    return df


def tabla_precio_por_region_tipo(tabla_precio_mercado):
    """Junta la tabla de precio de mercado (que calcula precio_mercado.py por
    COMUNA) a nivel de REGIÓN, para usar como respaldo cuando una comuna
    puntual se queda corta de publicaciones propias (algo normal: muchas
    comunas chicas nunca van a juntar 15-30 publicaciones solas, pero su
    región sí). Usa las mismas 16 regiones que ya usa el resto del ranking
    (REGIONES / zona_de), así que calza exacto con la columna 'region' del
    resto del programa - no hay que mantener una lista de regiones aparte."""
    acumulado = {}  # (region, tipo) -> [suma_ponderada_por_n_publicaciones, suma_n]
    for fila in tabla_precio_mercado:
        cod_sii = NOMBRE2SII.get(str(fila["comuna"]).strip().upper())
        region, _prov, _zonas = zona_de(cod_sii) if cod_sii else (None, None, set())
        if not region:
            continue
        clave = (region, fila["tipo_propiedad"])
        n = fila["n_publicaciones"]
        if clave not in acumulado:
            acumulado[clave] = [0.0, 0]
        acumulado[clave][0] += fila["precio_m2_estimado"] * n
        acumulado[clave][1] += n

    resultado = {}
    for (region, tipo), (suma_pond, suma_n) in acumulado.items():
        if suma_n <= 0:
            continue
        resultado[(region, tipo)] = {
            "precio_m2_estimado": suma_pond / suma_n,
            "n_publicaciones": suma_n,
            "confianza_mercado": precio_mercado.nivel_confianza(suma_n) if PRECIO_MERCADO_DISPONIBLE else "BAJA",
        }
    return resultado


def estimar_valor_mercado_por_comuna_tipo(df, tabla_precio_mercado):
    """Para propiedades sin un valor de mercado EXACTO por ROL (datos/valores_mercado.csv),
    estima uno aproximado usando el precio promedio del m2 que calculó precio_mercado.py,
    multiplicado por la superficie de la propiedad. Se intenta primero por COMUNA+tipo
    (más preciso); si esa comuna no tiene suficientes publicaciones propias (confianza
    BAJA o SIN DATOS SUFICIENTES, o directamente no hay ninguna publicación de esa comuna),
    se usa como respaldo el promedio de toda la REGIÓN+tipo (tabla_precio_por_region_tipo).
    Solo se usa un valor si su confianza es MEDIA o ALTA (nunca BAJA ni SIN DATOS
    SUFICIENTES, para no meter ruido al ranking con estimaciones poco sólidas)."""
    df["valor_mercado_estimado_m2_pesos"] = pd.NA
    df["confianza_mercado_m2"] = None
    df["n_publicaciones_mercado_m2"] = pd.NA
    df["nivel_mercado_m2"] = None  # "COMUNA" o "REGION", para que quede claro de dónde salió
    if not tabla_precio_mercado:
        return df

    indice_comuna = {}
    for fila in tabla_precio_mercado:
        if fila["confianza_mercado"] in ("ALTA", "MEDIA"):
            indice_comuna[(str(fila["comuna"]).strip().lower(), fila["tipo_propiedad"])] = fila

    indice_region = {
        (region, tipo): fila
        for (region, tipo), fila in tabla_precio_por_region_tipo(tabla_precio_mercado).items()
        if fila["confianza_mercado"] in ("ALTA", "MEDIA")
    }

    def _tipo_desde_destino(f):
        # Traducción simple del tipo de propiedad del SII/fuente al vocabulario del
        # módulo de precio de mercado. Se puede refinar con el tiempo.
        d = str(f.get("tipo_propiedad") or f.get("destino_desc") or "").lower()
        if "depart" in d or "depto" in d:
            return "departamento"
        if "casa" in d:
            return "casa"
        if "sitio" in d or "parcela" in d or "terreno" in d or "rural" in d:
            return "sitio_eriazo_rural"
        return None

    for i, f in df.iterrows():
        comuna_nombre = SII2NOMBRE.get(f.get("comuna_sii")) if isinstance(f.get("comuna_sii"), str) else None
        tipo = _tipo_desde_destino(f)
        if not tipo:
            continue
        superficie = f.get("superficie_util_m2") or f.get("superficie_total_m2")
        if not superficie or pd.isna(superficie) or superficie <= 0:
            continue

        nivel = None
        hallazgo = indice_comuna.get((comuna_nombre.strip().lower(), tipo)) if comuna_nombre else None
        if hallazgo:
            nivel = "COMUNA"
        else:
            region_propiedad = f.get("region")
            if region_propiedad:
                hallazgo = indice_region.get((region_propiedad, tipo))
                if hallazgo:
                    nivel = "REGION"
        if not hallazgo:
            continue

        df.at[i, "valor_mercado_estimado_m2_pesos"] = hallazgo["precio_m2_estimado"] * float(superficie)
        df.at[i, "confianza_mercado_m2"] = hallazgo["confianza_mercado"]
        df.at[i, "n_publicaciones_mercado_m2"] = hallazgo["n_publicaciones"]
        df.at[i, "nivel_mercado_m2"] = nivel
    return df


def calcular(df, hay_catastro, tabla_precio_mercado=None):
    for c in ("minimo_valor", "minimo_unidad", "alertas_fuente", "ocupacion", "rol_origen", "minimo_nota",
              "verif_direccion_catastro", "estado_busqueda_direccion", "direccion_tgr", "fuente_sii", "tipo_sii",
              "destino_sii", "cruce_estado"):
        if c not in df.columns:
            df[c] = None
    df["minimo_valor"] = pd.to_numeric(df["minimo_valor"], errors="coerce")
    if not hay_catastro:
        df["cruce_estado"] = "SIN CATASTRO CARGADO"
        for c in ("direccion_sii", "avaluo_sii", "exento_sii", "destino_sii", "tipo_sii", "fuente_sii"):
            df[c] = None
    for c in ("avaluo_sii", "tasacion", "avaluo_tgr", "avaluo_roles_total", "n_roles_encontrados", "n_roles"):
        if c not in df.columns:
            df[c] = None
        df[c] = pd.to_numeric(df[c], errors="coerce")
    z = df["comuna_sii"].map(lambda c: zona_de(c) if isinstance(c, str) else (None, None, set()))
    df["region"] = z.map(lambda t: t[0])
    df["provincia"] = z.map(lambda t: t[1])
    df["zonas"] = z.map(lambda t: t[2])
    df["direccion_final"] = df.apply(
        lambda f: f["direccion_sii"] if direccion_valida(f["direccion_sii"]) else f.get("direccion_tgr"), axis=1)
    df["origen_direccion"] = df.apply(
        lambda f: "SII" if direccion_valida(f["direccion_sii"])
        else ("FUENTE" if direccion_valida(f.get("direccion_tgr")) else "SIN DIRECCION"), axis=1)
    df["destino_desc"] = df["destino_sii"].map(
        lambda d: DESTINOS_CONFIRMADOS.get(d, f"(código {d})") if isinstance(d, str) and d else None)
    df["rol_completo"] = df.apply(
        lambda f: f"{f['comuna_sii']}-{int(f['manzana'])}-{int(f['predio'])}"
        if pd.notna(f["manzana"]) and pd.notna(f["predio"]) and f["comuna_sii"] else None, axis=1)
    df["dif_avaluo_tgr_vs_sii_pct"] = (df["avaluo_tgr"] / df["avaluo_sii"] - 1) * 100
    df["ratio_tasacion_avaluo_tgr"] = df["tasacion"] / df["avaluo_tgr"]
    df["fecha_remate"] = pd.to_datetime(df["fecha_remate"], errors="coerce")

    # ---- Deduplicación: la misma propiedad y fecha en varias fuentes/demandas -> una fila ----
    antes = len(df)
    df["_prio"] = df["fuente"].map(PRIORIDAD_FUENTE).fillna(9) + df["minimo_valor"].isna().astype(int) * 10
    df = df.sort_values(["_prio"]).reset_index(drop=True)
    df["_clave"] = df.apply(
        lambda f: f"{f['rol_completo']}|{f['fecha_remate'].date() if pd.notna(f['fecha_remate']) else 'sf'}"
        if f["rol_completo"] else f"unico|{f.name}", axis=1)
    df["tambien_en"] = df.groupby("_clave")["fuente"].transform(lambda s: ", ".join(sorted(set(s))))
    df["n_demandas"] = df.groupby("_clave")["fuente"].transform("size")
    df = df.drop_duplicates("_clave", keep="first").drop(columns=["_prio", "_clave"]).reset_index(drop=True)
    decir(f"  Filas recibidas: {antes:,} | propiedades únicas (misma propiedad y fecha): {len(df):,}")

    df = aplicar_minimos(df)

    # ---- Valor de referencia ----
    # Orden de preferencia: 1) valor de mercado EXACTO por ROL (datos/valores_mercado.csv,
    # si existe) > 2) estimado por comuna+tipo a partir de lo que recorrió precio_mercado.py
    # (solo si confianza MEDIA o ALTA) > 3) avalúo fiscal del SII.
    df["valor_mercado_pesos"] = pd.NA
    vm = cargar_valores_mercado()
    if vm is not None:
        df = df.drop(columns=["valor_mercado_pesos"]).merge(vm, on=["comuna_sii", "manzana", "predio"], how="left")
    df["valor_mercado_pesos"] = pd.to_numeric(df["valor_mercado_pesos"], errors="coerce")

    df = estimar_valor_mercado_por_comuna_tipo(df, tabla_precio_mercado)

    multi_ok = (df["n_roles"] > 1) & (df["n_roles_encontrados"] == df["n_roles"]) & df["avaluo_roles_total"].notna()
    avaluo_ref = df["avaluo_sii"].where(~multi_ok, df["avaluo_roles_total"])

    valor_mercado_final = df["valor_mercado_pesos"].fillna(df["valor_mercado_estimado_m2_pesos"])
    df["valor_ref_pesos"] = valor_mercado_final.fillna(avaluo_ref)

    def _origen_valor_ref(f):
        if pd.notna(f["valor_mercado_pesos"]):
            return "MERCADO (exacto por ROL)"
        if pd.notna(f["valor_mercado_estimado_m2_pesos"]):
            return (f"MERCADO (estimado por m2 - {f['nivel_mercado_m2']}, "
                    f"confianza {f['confianza_mercado_m2']})")
        if pd.notna(f["avaluo_sii"]) or pd.notna(f.get("avaluo_roles_total")):
            return "AVALUO FISCAL SII"
        return None

    df["base_valor_ref"] = df.apply(_origen_valor_ref, axis=1)

    # ---- Oportunidad (general: usa lo mejor disponible - mercado si hay, si no avalúo) ----
    costo = df["minimo_pesos"]
    df["margen_pesos"] = df["valor_ref_pesos"] - costo
    df["oportunidad_pct"] = (df["valor_ref_pesos"] / costo - 1) * 100
    df.loc[~(costo > 0), ["oportunidad_pct", "margen_pesos"]] = None
    df["oportunidad_pct"] = pd.to_numeric(df["oportunidad_pct"], errors="coerce")
    df["margen_pesos"] = pd.to_numeric(df["margen_pesos"], errors="coerce")

    # ---- Oportunidad SEPARADA por tipo de referencia, para el doble ranking ----
    # Ranking A: siempre contra el avalúo fiscal del SII (como funcionaba hasta ahora).
    df["margen_fiscal_pesos"] = avaluo_ref - costo
    df["oportunidad_fiscal_pct"] = (avaluo_ref / costo - 1) * 100
    df.loc[~(costo > 0) | avaluo_ref.isna(), ["oportunidad_fiscal_pct", "margen_fiscal_pesos"]] = None
    df["oportunidad_fiscal_pct"] = pd.to_numeric(df["oportunidad_fiscal_pct"], errors="coerce")
    df["margen_fiscal_pesos"] = pd.to_numeric(df["margen_fiscal_pesos"], errors="coerce")

    # Ranking B: solo contra un precio de mercado (exacto por ROL, o estimado con confianza ALTA/MEDIA).
    df["margen_mercado_pesos"] = valor_mercado_final - costo
    df["oportunidad_mercado_pct"] = (valor_mercado_final / costo - 1) * 100
    df.loc[~(costo > 0) | valor_mercado_final.isna(), ["oportunidad_mercado_pct", "margen_mercado_pesos"]] = None
    df["oportunidad_mercado_pct"] = pd.to_numeric(df["oportunidad_mercado_pct"], errors="coerce")
    df["margen_mercado_pesos"] = pd.to_numeric(df["margen_mercado_pesos"], errors="coerce")

    # ---- Alertas ----
    hoy = pd.Timestamp(dt.date.today())

    def alertas(f):
        a = [x for x in str(f.get("alertas_fuente") or "").split("; ") if x]
        if f["fuente"] == "TGR":
            a.append("MINIMO_ESTIMADO_TGR")
            r = f.get("ratio_tasacion_avaluo_tgr")
            if pd.notna(r) and abs(r - 1.3) > 0.05:
                a.append("RATIO_TASACION_AVALUO_ANORMAL")
        if f.get("rol_origen") == "DEDUCIDO POR DIRECCION (verificar)":
            a.append("ROL_DEDUCIDO_POR_DIRECCION")
        if f.get("minimo_nota") in ("UF_APROXIMADA", "UF_SIN_CONVERTIR"):
            a.append(f["minimo_nota"])
        if pd.isna(f.get("avaluo_sii")):
            a.append("SIN_AVALUO_SII")
        if f.get("fuente_sii") == "A" or f.get("tipo_sii") == "R":
            a.append("RURAL_O_AGRICOLA")
        if isinstance(f.get("ocupacion"), str) and re.search(r"(?i)\bocupad", f["ocupacion"]) \
                and not re.search(r"(?i)desocup", f["ocupacion"]):
            a.append("OCUPADA")
        if f.get("destino_sii") == "K" or (pd.notna(f.get("predio")) and f["predio"] >= 90000):
            a.append("ROL_ES_MATRIZ_DE_CONDOMINIO")
        if f.get("verif_direccion_catastro") == "NO COINCIDE":
            a.append("DIRECCION_NO_COINCIDE")
        if "COMUNA ASUMIDA" in str(f.get("estado_busqueda_direccion") or ""):
            a.append("COMUNA_ASUMIDA_DEL_TRIBUNAL")
        if f.get("n_roles", 0) and f["n_roles"] > 1 and f.get("n_roles_encontrados") != f["n_roles"]:
            a.append("ROLES_PARCIALES")
        if pd.isna(f["fecha_remate"]):
            a.append("SIN_FECHA")
        elif f["fecha_remate"].normalize() < hoy:
            a.append("REMATE_YA_REALIZADO_O_VENCIDO")
        if pd.notna(f.get("oportunidad_pct")) and f["oportunidad_pct"] > OPORTUNIDAD_MAX_PCT:
            a.append("OPORTUNIDAD_SOSPECHOSA")
        vistos = []
        for x in a:
            if x not in vistos:
                vistos.append(x)
        return vistos

    df["_alertas"] = df.apply(alertas, axis=1)
    df["alertas"] = df["_alertas"].map(lambda l: "; ".join(l))

    def confianza(f):
        if pd.isna(f["oportunidad_pct"]):
            return "NO EVALUABLE"
        al = set(f["_alertas"])
        base = str(f["base_valor_ref"] or "")
        es_mercado_exacto = base == "MERCADO (exacto por ROL)"
        es_mercado_estimado_alta = base.startswith("MERCADO (estimado por m2") and "ALTA" in base
        if (es_mercado_exacto or es_mercado_estimado_alta) and not (al & {"ROL_DEDUCIDO_POR_DIRECCION", "UF_APROXIMADA"}):
            return "ALTA"
        if base.startswith("MERCADO (estimado por m2") and "MEDIA" in base \
                and not (al & {"ROL_DEDUCIDO_POR_DIRECCION", "UF_APROXIMADA"}):
            return "MEDIA"
        if f["tipo_minimo"].startswith("ESTIMADO") or al & {"ROL_DEDUCIDO_POR_DIRECCION", "RURAL_O_AGRICOLA",
                                                              "UF_APROXIMADA", "COMUNA_ASUMIDA_DEL_TRIBUNAL"}:
            return "BAJA"
        return "MEDIA"

    df["confianza"] = df.apply(confianza, axis=1)

    def motivo(f):
        m = []
        if SOLO_VIGENTES and ("REMATE_YA_REALIZADO_O_VENCIDO" in f["_alertas"] or "SIN_FECHA" in f["_alertas"]):
            m.append("remate vencido o sin fecha")
        if pd.isna(f["oportunidad_pct"]):
            m.append("no evaluable (falta mínimo o avalúo)")
        for k, txt in BLOQUEANTES.items():
            if k in f["_alertas"]:
                m.append(txt)
        if f.get("destino_sii") in EXCLUIR_DESTINOS_TOP and f.get("n_roles", 0) <= 1:
            m.append("estacionamiento/bodega suelto")
        if pd.notna(f["valor_ref_pesos"]) and f["valor_ref_pesos"] < VALOR_MINIMO_REF:
            m.append("valor de referencia bajo el mínimo configurado")
        if pd.notna(f["oportunidad_pct"]) and f["oportunidad_pct"] < OPORTUNIDAD_MIN_PCT:
            m.append("oportunidad bajo el mínimo configurado")
        if EXIGIR_CRUCE_SII and not str(f.get("cruce_estado", "")).startswith("OK"):
            m.append("sin verificación en catastro SII")
        return "; ".join(m)

    df["motivo_no_elegible"] = df.apply(motivo, axis=1)
    df["ELEGIBLE"] = df["motivo_no_elegible"] == ""

    # ---- Elegibilidad SEPARADA para el doble ranking (mismos bloqueos de fondo,
    # pero cada ranking evalúa su propia oportunidad/valor de referencia) ----
    def _motivo_generico(f, oportunidad_pct, valor_ref):
        m = []
        if SOLO_VIGENTES and ("REMATE_YA_REALIZADO_O_VENCIDO" in f["_alertas"] or "SIN_FECHA" in f["_alertas"]):
            m.append("remate vencido o sin fecha")
        if pd.isna(oportunidad_pct):
            m.append("no evaluable (falta mínimo o valor de referencia)")
        for k, txt in BLOQUEANTES.items():
            if k in f["_alertas"]:
                m.append(txt)
        if f.get("destino_sii") in EXCLUIR_DESTINOS_TOP and f.get("n_roles", 0) <= 1:
            m.append("estacionamiento/bodega suelto")
        if pd.notna(valor_ref) and valor_ref < VALOR_MINIMO_REF:
            m.append("valor de referencia bajo el mínimo configurado")
        if pd.notna(oportunidad_pct) and oportunidad_pct < OPORTUNIDAD_MIN_PCT:
            m.append("oportunidad bajo el mínimo configurado")
        if EXIGIR_CRUCE_SII and not str(f.get("cruce_estado", "")).startswith("OK"):
            m.append("sin verificación en catastro SII")
        return "; ".join(m)

    df["motivo_no_elegible_fiscal"] = df.apply(
        lambda f: _motivo_generico(f, f["oportunidad_fiscal_pct"], avaluo_ref.loc[f.name]), axis=1)
    df["ELEGIBLE_FISCAL"] = df["motivo_no_elegible_fiscal"] == ""

    df["motivo_no_elegible_mercado"] = df.apply(
        lambda f: _motivo_generico(f, f["oportunidad_mercado_pct"], valor_mercado_final.loc[f.name]), axis=1)
    df["ELEGIBLE_MERCADO"] = df["motivo_no_elegible_mercado"] == ""
    return df.drop(columns=["_alertas"])


# ------------------------------------------------------------------------------
# 10. SELECCIÓN DE LAS MEJORES (con cupos por zona)
# ------------------------------------------------------------------------------
def seleccionar_top(df, n_top=None, cupos=None):
    n_top = N_TOP if n_top is None else n_top
    cupos = CUPOS if cupos is None else cupos
    df = df.copy()
    df["ranking_nacional"] = pd.NA
    df["razon_seleccion"] = None
    df["EN_TOP"] = False
    el = df[df["ELEGIBLE"]].sort_values(["oportunidad_pct", "margen_pesos"], ascending=False)
    for k, i in enumerate(el.index, 1):
        df.at[i, "ranking_nacional"] = k
    elegidos = []
    faltantes = []
    for zona, n in cupos.items():
        cand = el[el["zonas"].map(lambda z: zona in z) & ~el.index.isin(elegidos)].head(n)
        for k, i in enumerate(cand.index, 1):
            df.at[i, "razon_seleccion"] = f"CUPO {NOMBRE_ZONA.get(zona, zona)} (#{k} de la zona)"
            elegidos.append(i)
        if len(cand) < n:
            faltantes.append(f"{NOMBRE_ZONA.get(zona, zona)}: solo {len(cand)} de {n} cupos (no hay más elegibles)")
    for i in el.index:
        if len(elegidos) >= n_top:
            break
        if i not in elegidos:
            df.at[i, "razon_seleccion"] = f"RANKING NACIONAL (#{int(df.at[i, 'ranking_nacional'])})"
            elegidos.append(i)
    df.loc[elegidos, "EN_TOP"] = True
    df["ranking_nacional"] = pd.to_numeric(df["ranking_nacional"], errors="coerce").astype("Int64")
    return df, faltantes


def _llenar_con_cupos_por_zona(candidatas, cupos, objetivo, ya_elegidos, etiqueta_ranking):
    """Toma primero los cupos por zona, y después completa con las siguientes mejores
    (sin importar zona) hasta llegar a 'objetivo'. 'candidatas' debe venir YA ordenada
    de mejor a peor oportunidad. Devuelve (lista_de_indices_elegidos, razones, faltantes)."""
    elegidos, razones, faltantes = [], {}, []
    disponibles = candidatas[~candidatas.index.isin(ya_elegidos)]
    for zona, n in cupos.items():
        cand = disponibles[disponibles["zonas"].map(lambda z: zona in z) & ~disponibles.index.isin(elegidos)].head(n)
        for k, i in enumerate(cand.index, 1):
            razones[i] = f"{etiqueta_ranking} - CUPO {NOMBRE_ZONA.get(zona, zona)} (#{k} de la zona)"
            elegidos.append(i)
        if len(cand) < n:
            faltantes.append(f"{etiqueta_ranking} - {NOMBRE_ZONA.get(zona, zona)}: solo {len(cand)} de {n} cupos")
    for i in disponibles.index:
        if len(elegidos) >= objetivo:
            break
        if i not in elegidos:
            razones[i] = f"{etiqueta_ranking} (ranking #{len(elegidos) + 1})"
            elegidos.append(i)
    if len(elegidos) < objetivo:
        faltantes.append(f"{etiqueta_ranking}: solo {len(elegidos)} de {objetivo} elegibles en total")
    return elegidos, razones, faltantes


def _forzar_cuota_tgr(df, elegidos, razones, universo_elegible, minimo_tgr, minimo_tgr_con_causa):
    """Garantiza que el TOP tenga al menos 'minimo_tgr' propiedades de la TGR, y que al
       menos 'minimo_tgr_con_causa' de esas tengan tribunal + rol de causa (= carpeta
       judicial disponible para el Poder Judicial). Si el ranking normal no las trae
       solas, se cambian las propiedades NO-TGR (o TGR sin causa) con peor oportunidad
       por las mejores candidatas de la TGR que falten. El TOP nunca cambia de tamaño:
       solo se intercambian cupos, uno por uno."""
    elegidos = list(elegidos)
    razones = dict(razones)
    faltantes = []

    def tiene_causa(i):
        return pd.notna(df.at[i, "tribunal"]) and pd.notna(df.at[i, "rol_causa"])

    def oportunidad(i):
        v = df.at[i, "oportunidad_pct"]
        return v if pd.notna(v) else -1

    def candidatos_tgr(con_causa=None):
        pool = [i for i in universo_elegible if i not in elegidos and df.at[i, "fuente"] == "TGR"]
        if con_causa is True:
            pool = [i for i in pool if tiene_causa(i)]
        elif con_causa is False:
            pool = [i for i in pool if not tiene_causa(i)]
        return sorted(pool, key=oportunidad, reverse=True)

    def peor_no_tgr():
        no_tgr = [i for i in elegidos if df.at[i, "fuente"] != "TGR"]
        return min(no_tgr, key=oportunidad) if no_tgr else None

    # ---- Paso 1: al menos 'minimo_tgr' propiedades de la TGR en el TOP ----
    faltan = minimo_tgr - sum(1 for i in elegidos if df.at[i, "fuente"] == "TGR")
    for nuevo in candidatos_tgr():
        if faltan <= 0:
            break
        peor = peor_no_tgr()
        if peor is None:
            faltantes.append("Cuota mínima de la TGR: no quedan propiedades no-TGR que reemplazar en el TOP.")
            break
        elegidos.remove(peor)
        razones.pop(peor, None)
        elegidos.append(nuevo)
        razones[nuevo] = f"FORZADO: cuota mínima de la TGR (al menos {minimo_tgr} en el TOP)"
        faltan -= 1
    if faltan > 0:
        faltantes.append(f"Cuota mínima de la TGR: solo se alcanzaron {minimo_tgr - faltan} de {minimo_tgr} "
                          f"(no hay más propiedades elegibles de la TGR esta semana).")

    # ---- Paso 2: de esas, al menos 'minimo_tgr_con_causa' con tribunal + rol de causa ----
    con_causa = sum(1 for i in elegidos if df.at[i, "fuente"] == "TGR" and tiene_causa(i))
    faltan_causa = minimo_tgr_con_causa - con_causa
    for nuevo in candidatos_tgr(con_causa=True):
        if faltan_causa <= 0:
            break
        tgr_sin_causa = [i for i in elegidos if df.at[i, "fuente"] == "TGR" and not tiene_causa(i)]
        peor = min(tgr_sin_causa, key=oportunidad) if tgr_sin_causa else peor_no_tgr()
        if peor is None:
            faltantes.append("Cuota de carpetas judiciales de la TGR: no hay más propiedades que reemplazar.")
            break
        elegidos.remove(peor)
        razones.pop(peor, None)
        elegidos.append(nuevo)
        razones[nuevo] = (f"FORZADO: cuota mínima de la TGR con carpeta judicial "
                           f"(al menos {minimo_tgr_con_causa} con tribunal + rol de causa)")
        faltan_causa -= 1
    if faltan_causa > 0:
        faltantes.append(f"Cuota de carpetas judiciales de la TGR: solo se alcanzaron "
                          f"{minimo_tgr_con_causa - faltan_causa} de {minimo_tgr_con_causa} "
                          f"(no hay más propiedades de la TGR con tribunal + rol de causa esta semana).")

    return elegidos, razones, faltantes


def seleccionar_top_dual(df, n_fiscal=None, n_mercado=None, n_comodin=None, cupos=None):
    """Doble ranking acordado con Rodrigo:
       - hasta 20 mejores por oportunidad vs. AVALÚO FISCAL (Ranking A)
       - hasta 20 mejores por oportunidad vs. PRECIO DE MERCADO (Ranking B)
       - hasta 10 'comodín': las mejores que queden, de cualquiera de los dos
       - máximo 50 en total; si el Ranking B no llega a su cupo, el Ranking A
         se queda con los cupos que sobraron.
       - la cuota por zona (2 Concepción, 2 RM, 2 Valparaíso) se exige por
         separado en el Ranking A y en el Ranking B."""
    n_fiscal = N_TOP_FISCAL if n_fiscal is None else n_fiscal
    n_mercado = N_TOP_MERCADO if n_mercado is None else n_mercado
    n_comodin = N_TOP_COMODIN if n_comodin is None else n_comodin
    cupos = CUPOS if cupos is None else cupos

    df = df.copy()
    df["razon_seleccion"] = None
    df["EN_TOP"] = False
    df["ranking_fiscal"] = pd.NA
    df["ranking_mercado"] = pd.NA

    cand_fiscal = df[df["ELEGIBLE_FISCAL"]].sort_values(
        ["oportunidad_fiscal_pct", "margen_fiscal_pesos"], ascending=False)
    for k, i in enumerate(cand_fiscal.index, 1):
        df.at[i, "ranking_fiscal"] = k

    cand_mercado = df[df["ELEGIBLE_MERCADO"]].sort_values(
        ["oportunidad_mercado_pct", "margen_mercado_pesos"], ascending=False)
    for k, i in enumerate(cand_mercado.index, 1):
        df.at[i, "ranking_mercado"] = k

    faltantes = []

    # ---- Ranking B primero (precio de mercado) ----
    elegidos_b, razones_b, falt_b = _llenar_con_cupos_por_zona(
        cand_mercado, cupos, n_mercado, [], "OPORTUNIDAD VS. MERCADO")
    faltantes += falt_b
    cupos_prestados_a_fiscal = n_mercado - len(elegidos_b)
    if cupos_prestados_a_fiscal > 0:
        faltantes.append(
            f"El Ranking de mercado solo llenó {len(elegidos_b)} de {n_mercado} cupos; "
            f"los {cupos_prestados_a_fiscal} que faltaron se le prestaron al Ranking de avalúo fiscal.")

    # ---- Ranking A (avalúo fiscal), con los cupos propios + lo que le prestó B ----
    objetivo_a = n_fiscal + cupos_prestados_a_fiscal
    elegidos_a, razones_a, falt_a = _llenar_con_cupos_por_zona(
        cand_fiscal, cupos, objetivo_a, elegidos_b, "OPORTUNIDAD VS. AVALUO FISCAL")
    faltantes += falt_a

    ya_elegidos = set(elegidos_b) | set(elegidos_a)

    # ---- Comodín: las mejores que queden, de cualquiera de los dos rankings ----
    restantes = df[df.index.isin(set(cand_fiscal.index) | set(cand_mercado.index)) & ~df.index.isin(ya_elegidos)]
    # para comparar oportunidades de distinta naturaleza, se usa la mejor disponible de cada una (oportunidad_pct general)
    restantes = restantes.sort_values(["oportunidad_pct", "margen_pesos"], ascending=False)
    elegidos_comodin = list(restantes.index)[:n_comodin]
    if len(elegidos_comodin) < n_comodin:
        faltantes.append(f"Comodín: solo {len(elegidos_comodin)} de {n_comodin} cupos (no quedan más elegibles)")

    razones = {}
    for i in elegidos_b:
        razones[i] = razones_b[i]
    for i in elegidos_a:
        razones[i] = razones_a[i]
    for k, i in enumerate(elegidos_comodin, 1):
        razones[i] = f"COMODIN (mejor oportunidad disponible, #{k})"

    todos_elegidos = elegidos_b + elegidos_a + elegidos_comodin
    universo_elegible = set(cand_fiscal.index) | set(cand_mercado.index)
    todos_elegidos, razones, faltantes_tgr = _forzar_cuota_tgr(
        df, todos_elegidos, razones, universo_elegible,
        MINIMO_TGR_EN_TOP, MINIMO_TGR_CON_CAUSA_EN_TOP)
    faltantes += faltantes_tgr

    for i, razon in razones.items():
        df.at[i, "razon_seleccion"] = razon

    df.loc[todos_elegidos, "EN_TOP"] = True
    df["ranking_fiscal"] = pd.to_numeric(df["ranking_fiscal"], errors="coerce").astype("Int64")
    df["ranking_mercado"] = pd.to_numeric(df["ranking_mercado"], errors="coerce").astype("Int64")
    return df, faltantes


# ------------------------------------------------------------------------------
# 11. EXPORTACIÓN
# ------------------------------------------------------------------------------
COLUMNAS_TOP = ["razon_seleccion", "ranking_fiscal", "ranking_mercado", "ranking_nacional",
                "fuente", "tambien_en", "fecha_remate", "modalidad",
                "region", "provincia", "comuna_propiedad", "rol_formato", "rol_completo", "rol_origen",
                "direccion_final", "origen_direccion", "destino_desc", "tipo_propiedad",
                "superficie_util_m2", "superficie_total_m2", "dormitorios", "banos", "ocupacion",
                "avaluo_sii", "valor_ref_pesos", "base_valor_ref", "minimo_pesos", "tipo_minimo",
                "oportunidad_pct", "margen_pesos",
                "oportunidad_fiscal_pct", "margen_fiscal_pesos",
                "oportunidad_mercado_pct", "margen_mercado_pesos",
                "confianza", "alertas",
                "tribunal", "rol_causa", "tipo_juicio", "n_remate", "fojas", "fojas_numero", "fojas_anio", "cbr",
                "estado_causa_pjud", "etapa_causa_pjud", "link_demanda_pjud",
                "garantia_pesos", "plazo_pago", "url", "id_fuente", "motivo_no_elegible"]
COLUMNAS_OCULTAS = {"zonas", "texto_aviso", "direccion_texto_busqueda", "roles_lista", "candidatos_comuna",
                    "datos_subasta"}


def ordenar_columnas(df):
    primeras = [c for c in COLUMNAS_TOP if c in df.columns]
    resto = [c for c in df.columns if c not in primeras and c not in COLUMNAS_OCULTAS]
    return df[primeras + resto]


def resumen_texto(df, db_path, hay_catastro, estado_fuentes, faltantes):
    top = df[df["EN_TOP"]]
    n_de_mercado = int(top["razon_seleccion"].str.startswith("OPORTUNIDAD VS. MERCADO", na=False).sum())
    n_de_fiscal = int(top["razon_seleccion"].str.startswith("OPORTUNIDAD VS. AVALUO FISCAL", na=False).sum())
    n_comodin = int(top["razon_seleccion"].str.startswith("COMODIN", na=False).sum())
    filas = [
        ("Fecha y hora de la ejecución", dt.datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Propiedades únicas analizadas", len(df)),
        ("Elegibles - ranking avalúo fiscal", int(df["ELEGIBLE_FISCAL"].sum())),
        ("Elegibles - ranking precio de mercado", int(df["ELEGIBLE_MERCADO"].sum())),
        (f"En el consolidado final (máx. {N_TOP_MAXIMO})", len(top)),
        (f"  - de los cuales, por avalúo fiscal", n_de_fiscal),
        (f"  - de los cuales, por precio de mercado", n_de_mercado),
        (f"  - de los cuales, comodín (mejor disponible)", n_comodin),
        ("Catastro SII usado", db_path.name if (hay_catastro and db_path) else "NO HAY CATASTRO CARGADO"),
        ("Cupos por zona (se exigen por separado en cada ranking)",
         "; ".join(f"{NOMBRE_ZONA.get(k, k)}={v}" for k, v in CUPOS.items())),
        ("Mínimo TGR", f"ESTIMADO = tasación x {TGR_MINIMO_FACTOR_SOBRE_TASACION:.3f} (la TGR no publica el mínimo)"),
        ("Valor de referencia", "Ranking A: avalúo fiscal SII siempre. Ranking B: precio de mercado "
                                "(exacto por ROL, o estimado por comuna/tipo con confianza ALTA/MEDIA)."),
    ]
    for f, e in estado_fuentes.items():
        filas.append((f"Fuente {f}", e))
    for fu, n in df["fuente"].value_counts().items():
        filas.append((f"Filas de {fu}", int(n)))
    for estado, n in df["cruce_estado"].value_counts().items():
        filas.append((f"Cruce SII: {estado}", int(n)))
    for k, n in df["confianza"].value_counts().items():
        filas.append((f"Confianza {k}", int(n)))
    sin_senal = int(pd.to_numeric(df["oportunidad_pct"], errors="coerce").abs().lt(1).sum())
    if sin_senal:
        filas.append(("SIN SEÑAL", f"{sin_senal} propiedades tienen mínimo = avalúo fiscal (oportunidad 0%): "
                                   "con avalúo no se pueden distinguir; se necesita precio de mercado."))
    for t in faltantes:
        filas.append(("AVISO DE CUPO", t))
    filas.append(("NOTA 1", "La oportunidad compara el valor de referencia con el mínimo de entrada. "
                            "Con avalúo fiscal NO es precio de mercado: es una señal para revisar, no una tasación."))
    filas.append(("NOTA 2", "El mínimo REAL viene publicado por la fuente (Macal, avisos). El de la TGR es ESTIMADO "
                            "y por eso todas sus propiedades tienen la misma oportunidad teórica."))
    return pd.DataFrame(filas, columns=["concepto", "valor"])


def exportar(df, db_path, hay_catastro, estado_fuentes, faltantes):
    SALIDAS.mkdir(exist_ok=True)
    df = ordenar_columnas(df)
    top = df[df["EN_TOP"]].sort_values(["oportunidad_pct"], ascending=False)
    cand = df[df["ELEGIBLE_FISCAL"] | df["ELEGIBLE_MERCADO"]].sort_values(["oportunidad_pct"], ascending=False)
    rev = df[~df["ELEGIBLE"] & df["motivo_no_elegible"].ne("")].copy()
    rev = rev[~rev["motivo_no_elegible"].str.contains("vencido", na=False)]
    resumen = resumen_texto(df, db_path, hay_catastro, estado_fuentes, faltantes)
    df.drop(columns=["EN_TOP"]).to_csv(SALIDAS / "todas_las_propiedades.csv", index=False, encoding="utf-8-sig")
    top.to_csv(SALIDAS / "TOP_OPORTUNIDADES.csv", index=False, encoding="utf-8-sig")
    ruta = SALIDAS / "OPORTUNIDADES.xlsx"
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        hojas = [("TOP_CONSOLIDADO", top), ("CANDIDATAS", cand), ("REVISAR", rev), ("TODAS", df), ("RESUMEN", resumen)]
        for nombre, d in hojas:
            d.to_excel(w, sheet_name=nombre, index=False)
            dar_formato(w.sheets[nombre])
    return ruta


def dar_formato(ws):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    relleno = PatternFill("solid", fgColor="1F3864")
    for celda in ws[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = relleno
        celda.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A2"
    if ws.max_row > 1 and ws.max_column > 0:
        ws.auto_filter.ref = ws.dimensions
    dinero = {"tasacion", "avaluo_tgr", "avaluo_sii", "exento_sii", "valor_ref_pesos", "minimo_pesos",
              "margen_pesos", "garantia_pesos", "avaluo_roles_total", "valor_mercado_pesos",
              "margen_fiscal_pesos", "margen_mercado_pesos", "valor_mercado_estimado_m2_pesos"}
    anchos45 = {"direccion_final", "direccion_tgr", "direccion_sii", "tribunal", "valor", "concepto",
                "alertas", "razon_seleccion", "motivo_no_elegible", "url", "rol_candidatos"}
    for i, celda in enumerate(ws[1], start=1):
        nombre = str(celda.value or "")
        ws.column_dimensions[get_column_letter(i)].width = 45 if nombre in anchos45 else max(12, min(30, len(nombre) + 4))
        fmt = None
        if nombre in dinero:
            fmt = "#,##0"
        elif nombre == "fecha_remate":
            fmt = "yyyy-mm-dd hh:mm"
        elif nombre in ("oportunidad_pct", "oportunidad_fiscal_pct", "oportunidad_mercado_pct",
                        "dif_avaluo_tgr_vs_sii_pct", "ratio_tasacion_avaluo_tgr"):
            fmt = "0.0"
        if fmt:
            for fila in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                fila[0].number_format = fmt


# ------------------------------------------------------------------------------
# 12. PROGRAMA PRINCIPAL
# ------------------------------------------------------------------------------
def guardar_log():
    try:
        SALIDAS.mkdir(exist_ok=True)
        (SALIDAS / "log_ultima_ejecucion.txt").write_text("\n".join(LOG), encoding="utf-8")
    except Exception:
        pass


def reunir_fuentes():
    """Barre las fuentes activadas. Devuelve (lista_de_tablas, estado_por_fuente)."""
    tablas, estado = [], {}
    SALIDAS.mkdir(exist_ok=True)
    sello = dt.datetime.now().strftime("%Y%m%d_%H%M")

    if FUENTES.get("TGR"):
        decir("\n  [TGR] descargando remates EN VIVO...")
        ultimo_error_tgr = None
        exito_tgr = False
        for intento in range(1, TGR_MAXIMO_INTENTOS + 1):
            try:
                decir(f"    Intento {intento}/{TGR_MAXIMO_INTENTOS}...")
                payload = descargar_tgr()
                (SALIDAS / f"tgr_respuesta_{sello}.json").write_text(
                    json.dumps(payload, ensure_ascii=False), encoding="utf-8")
                t = construir_tabla_tgr(payload)
                if t.empty:
                    raise RuntimeError("La TGR respondió, pero sin ningún remate.")
                t["fuente"] = "TGR"
                t["id_fuente"] = t["cod_demanda"].astype(str)
                tablas.append(t)
                estado["TGR"] = f"OK ({len(t):,} registros, {dt.datetime.now():%H:%M})"
                decir(f"  [TGR] OK: {len(t):,} registros")
                exito_tgr = True
                break
            except Exception as e:
                ultimo_error_tgr = str(e)
                decir(f"    -> Falló: {ultimo_error_tgr}")
                if intento < TGR_MAXIMO_INTENTOS:
                    time.sleep(TGR_ESPERA_ENTRE_INTENTOS_SEG)
        if not exito_tgr:
            estado["TGR"] = f"FALLÓ tras {TGR_MAXIMO_INTENTOS} intentos: {ultimo_error_tgr}"
            decir(f"  [TGR] FALLÓ tras {TGR_MAXIMO_INTENTOS} intentos. El programa SIGUE con las demás fuentes.")
            if "TGR" in FUENTES_CRITICAS_ALERTA:
                alerta_fuerte(
                    f"La fuente CRITICA 'TGR' no pudo obtener datos hoy, tras {TGR_MAXIMO_INTENTOS} intentos.\n"
                    f"Motivo: {ultimo_error_tgr}\n"
                    f"El programa SIGUE funcionando con las demas fuentes (Macal, Avisos, Precio de Mercado),\n"
                    f"pero los resultados de hoy NO incluyen remates de la TGR.\n"
                    f"Revisa si el sitio de la TGR cambio o esta caido."
                )
            if "TGR" in FUENTES_OBLIGATORIAS:
                raise RuntimeError(ultimo_error_tgr)
    if FUENTES.get("MACAL"):
        decir("\n  [MACAL] barrido completo de macal.cl...")
        try:
            fichas, fallos = descargar_macal()
            t = tabla_macal(fichas)
            tablas.append(t)
            estado["MACAL"] = (f"OK ({len(t)} fichas leídas, {len(fallos)} con error, {dt.datetime.now():%H:%M})")
            if fallos:
                estado["MACAL (fichas con error)"] = "; ".join(u for u, _ in fallos[:20])
            decir(f"  [MACAL] OK: {len(t)} propiedades")
        except Exception as e:
            estado["MACAL"] = f"FALLÓ: {e}"
            decir(f"  [MACAL] FALLÓ (el programa sigue con las demás fuentes): {e}")
            if "MACAL" in FUENTES_OBLIGATORIAS:
                raise
    if FUENTES.get("AVISOS"):
        decir("\n  [AVISOS] leyendo resúmenes .docx...")
        try:
            rutas = buscar_docx_avisos()
            if not rutas:
                estado["AVISOS"] = f"SIN ARCHIVOS: deja los .docx en {BASE / AVISOS_CARPETA_NOMBRE}"
                decir(f"  [AVISOS] no hay .docx en {BASE / AVISOS_CARPETA_NOMBRE}")
            else:
                t = tabla_avisos(rutas)
                tablas.append(t)
                estado["AVISOS"] = f"OK ({len(t)} avisos de {len(rutas)} archivo(s): " + ", ".join(p.name for p in rutas) + ")"
                decir(f"  [AVISOS] OK: {len(t)} avisos")
        except Exception as e:
            estado["AVISOS"] = f"FALLÓ: {e}"
            decir(f"  [AVISOS] FALLÓ: {e}")
            if "AVISOS" in FUENTES_OBLIGATORIAS:
                raise

    tabla_precio_mercado = None
    if FUENTES.get("PRECIO_MERCADO"):
        decir("\n  [PRECIO_MERCADO] actualizando base propia de precios de mercado...")
        if not PRECIO_MERCADO_DISPONIBLE:
            estado["PRECIO_MERCADO"] = "NO DISPONIBLE: falta el archivo precio_mercado.py en esta carpeta"
            decir("  [PRECIO_MERCADO] no encontré 'precio_mercado.py' en la carpeta. Sigo sin precio de mercado "
                  "(se usará solo el avalúo fiscal del SII como referencia).")
        else:
            try:
                tabla_precio_mercado, estado_proveedores = precio_mercado.actualizar_precio_mercado(DATOS, decir=decir)
                estado.update(estado_proveedores)
                estado["PRECIO_MERCADO"] = f"OK ({len(tabla_precio_mercado)} combinaciones comuna/tipo con datos)"
            except Exception as e:
                estado["PRECIO_MERCADO"] = f"FALLÓ: {e}"
                decir(f"  [PRECIO_MERCADO] FALLÓ por completo (el programa sigue sin precio de mercado): {e}")
    return tablas, estado, tabla_precio_mercado


def procesar(tablas, db_path, tabla_precio_mercado=None):
    hay_catastro = db_path is not None
    df = pd.concat([t for t in tablas if t is not None and not t.empty], ignore_index=True, sort=False)
    df = resolver_avisos(df, db_path)
    df = resolver_roles_faltantes(df, db_path)
    if hay_catastro:
        df = cruzar_con_catastro(df, db_path)
    df = calcular(df, hay_catastro, tabla_precio_mercado)
    return df, hay_catastro


def principal():
    decir("=" * 70)
    decir(" REMATES EN CHILE: BARRIDO COMPLETO + OPORTUNIDADES (v2)")
    decir("=" * 70)
    decir("\nPaso 1 de 4: catastro nacional del SII")
    db_path = preparar_catastro()
    if db_path is None:
        decir("  No encontré el archivo del catastro (BRORGA...N_NAC_...).")
        decir(f"  Déjalo (o su .zip) en: {DATOS}")
        decir("  Sigo SIN catastro: no habrá avalúo SII ni ROL por dirección (las oportunidades quedarán 'no evaluables').")
    decir("\nPaso 2 de 4: barrido en vivo de las fuentes")
    tablas, estado, tabla_precio_mercado = reunir_fuentes()
    if not tablas:
        raise RuntimeError("Ninguna fuente entregó datos.")
    decir("\nPaso 3 de 4: ROL, avalúo SII, mínimos, oportunidad y selección")
    df, hay_catastro = procesar(tablas, db_path, tabla_precio_mercado)
    df, faltantes = seleccionar_top_dual(df)
    for estado_c, n in df["cruce_estado"].value_counts().items():
        decir(f"    Cruce SII - {estado_c}: {n:,}")

    decir("\n  [PODER_JUDICIAL] consultando el estado de la causa de cada propiedad del TOP...")
    try:
        df = agregar_info_poder_judicial(df, decir=decir)
        estado["PODER_JUDICIAL"] = "OK"
    except Exception as e:
        estado["PODER_JUDICIAL"] = f"FALLÓ: {e}"
        decir(f"  [PODER_JUDICIAL] FALLÓ por completo (el programa sigue, solo sin esta info): {e}")

    decir("\nPaso 4 de 4: guardando archivos")
    ruta = exportar(df, db_path, hay_catastro, estado, faltantes)
    decir(f"\nLISTO. En el consolidado final (máx. {N_TOP_MAXIMO}): {int(df['EN_TOP'].sum())} propiedades | "
          f"elegibles fiscal: {int(df['ELEGIBLE_FISCAL'].sum())} | elegibles mercado: {int(df['ELEGIBLE_MERCADO'].sum())} "
          f"de {len(df):,} analizadas")
    for t in faltantes:
        decir(f"  AVISO: {t}")
    for f, e in estado.items():
        decir(f"  Fuente {f}: {e}")
    decir(f"  Archivos en: {SALIDAS}")
    decir(f"  Excel: {ruta.name}")
    guardar_log()


# ------------------------------------------------------------------------------
# 13. PRUEBAS SIN INTERNET  (python remates_v2.py --prueba)
# ------------------------------------------------------------------------------
FICHA_MACAL_PRUEBA = """<html><body><h1>Trinidad Ramírez 01010, Depto. 1207-C</h1>
<div>La Cisterna, Región RM</div><div>Precio mínimo</div><div>UF 450</div>
<ul><li>Tipo: Departamento</li><li>Superficie útil: 43 m²</li><li>Superficie total: 46 m²</li>
<li>Dormitorios: 3</li><li>Baños: 1</li><li>Estado de ocupación: DESOCUPADA</li>
<li>Rol avalúo: 3046-690</li><li>Plazo de pago: 30 días</li></ul>
<p>Vendedor: Banco BCI</p><p>Fecha de Subasta: 14-10-2026</p><p>Garantía Requerida: $4.000.000</p>
<script>var x = "ruido";</script></body></html>"""

AVISO_PRUEBA_1 = ("AVISO DE REMATE. El 2º Juzgado de Letras La Serena, ubicado en Rengifo Nº 240, rematará el 14 de "
                  "septiembre de 2026, a las 12:00 hrs., inmueble ubicado en La Serena, Pasaje El Sauco Nº 5.299, del "
                  "Loteo Altos de La Florida VI. Inscrito a fojas 4829 Nº 3439 del Registro de Propiedad del CBR de La "
                  "Serena del año 2016. Rol de avalúo Nº 3106-87, de La Serena. Mínimo para la subasta $83.952.908.- "
                  "Garantía: vale vista equivalente al 10% del mínimo, esto es $8.395.290. Demás condiciones autos "
                  "“BANCO X con PEREZ”, Rol C-434-2026, del tribunal citado. Subasta presencial en el tribunal.")
AVISO_PRUEBA_2 = ("Remate. Primer Juzgado Civil de Santiago, Huérfanos 1409, en causa Rol C-1283-2025, juicio ejecutivo, "
                  "se rematará por videoconferencia, mediante plataforma Zoom, el día 14 de septiembre de 2026, a las "
                  "13:40 horas, el departamento 803 del piso 8, con acceso por calle David Arellano Nº 1810, comuna de "
                  "Independencia, y la bodega Nº 12. Mínimo para las posturas UF 3.200,50. Segundo remate. Garantía 10% "
                  "del mínimo.")


def prueba():
    fallas = []

    def comprobar(nombre, condicion):
        decir(f"  [{'OK ' if condicion else 'FALLA'}] {nombre}")
        if not condicion:
            fallas.append(nombre)

    decir("PRUEBAS DE LÓGICA (no usan internet)")
    comprobar("ROL con formato", parsear_rol_formato("LEBU-00454-009") == ("LEBU", 454, 9))
    comprobar("ROL con guion en la comuna", parsear_rol_formato("LLAY-LLAY-00010-002") == ("LLAY-LLAY", 10, 2))
    comprobar("ROL de 11 dígitos", parsear_rol_11("19900454009") == ("199", 454, 9))
    comprobar("Tesorería 188 es Concepción (SII 08201)", TES2SII.get("188") == "08201")
    comprobar("347 comunas cargadas", len(COMUNAS) == 347)
    comprobar("Número chileno 83.952.908.-", a_numero_chileno("83.952.908.-") == 83952908)
    comprobar("Número chileno 7.785,48", a_numero_chileno("7.785,48") == 7785.48)
    comprobar("Zona: Concepción 08201 -> PROV:082", "PROV:082" in zona_de("08201")[2])
    comprobar("Zona: Providencia (15xxx) es Región Metropolitana", "REG:RM" in zona_de("15108")[2])
    comprobar("Zona: Viña del Mar es REG:05", "REG:05" in zona_de("05109")[2])
    comprobar("Modalidad REMOTO", modalidad_desde_texto("Subasta por plataforma Zoom ID 99") == "REMOTO")
    comprobar("Modalidad PRESENCIAL", modalidad_desde_texto("se hará de forma presencial en el tribunal") == "PRESENCIAL")

    # --- Avisos ---
    a1 = parsear_aviso(AVISO_PRUEBA_1, (14, 9))
    comprobar("Aviso: mínimo REAL $83.952.908 (no la garantía $8.395.290)",
              a1["minimo_valor"] == 83952908 and a1["minimo_unidad"] == "CLP")
    comprobar("Aviso: rol de avalúo 3106-87", a1["roles"] == [(3106, 87)])
    comprobar("Aviso: rol de causa C-434-2026", a1["rol_causa"] == "C-434-2026")
    comprobar("Aviso: fecha y hora 2026-09-14 12:00", a1["fecha_remate"] == dt.datetime(2026, 9, 14, 12, 0))
    comprobar("Aviso: fojas 4829 número 3439 año 2016 CBR LA SERENA",
              (a1["fojas"], a1["fojas_numero"], a1["fojas_anio"], a1["cbr"]) == (4829, 3439, 2016, "LA SERENA"))
    comprobar("Aviso: juzgado leído", a1["tribunal"] and "Juzgado de Letras La Serena" in a1["tribunal"])
    comprobar("Aviso: modalidad PRESENCIAL", a1["modalidad"] == "PRESENCIAL")
    comprobar("Aviso: comuna La Serena detectada", a1["candidatos"] and a1["candidatos"][0][0] == NOMBRE2SII["LA SERENA"])
    d1 = extraer_direccion_texto(AVISO_PRUEBA_1)
    comprobar("Aviso: dirección de la PROPIEDAD (El Sauco 5299), no la del tribunal (Rengifo)",
              d1 == {"nombre": "EL SAUCO", "numero": 5299, "unidad": None})
    a2 = parsear_aviso(AVISO_PRUEBA_2, (14, 9))
    comprobar("Aviso 2: mínimo en UF 3200,50", a2["minimo_valor"] == 3200.5 and a2["minimo_unidad"] == "UF")
    comprobar("Aviso 2: remoto, juicio ejecutivo, segundo remate",
              a2["modalidad"] == "REMOTO" and a2["tipo_juicio"] == "CIVIL - EJECUTIVO" and a2["n_remate"] == 2)
    comprobar("Aviso 2: alerta de varias unidades", "VARIAS_UNIDADES" in a2["alertas_aviso"])
    comprobar("Aviso 2: comuna = Independencia (no la del tribunal)",
              a2["candidatos"] and a2["candidatos"][0][0] == NOMBRE2SII["INDEPENDENCIA"])
    comprobar("Aviso partición detectado",
              extraer_tipo_juicio("ante el juez partidor se rematará") == "PARTICIONAL")
    comprobar("Un rol de causa sin letra no se confunde con rol de avalúo", extraer_roles("ROL 2583-2023 DEL JUZGADO") == [])
    comprobar("Cabecera de fecha", cabecera_fecha("14 DE SEPTIEMBRE") == (14, 9))

    # --- Macal ---
    f = parsear_ficha_macal(FICHA_MACAL_PRUEBA, "/venta/departamento/la-cisterna/trinidad-ramirez-01010/P3170-1")
    comprobar("Macal: precio UF 450", (f["minimo_valor"], f["minimo_unidad"]) == (450, "UF"))
    comprobar("Macal: ROL 3046-690", f["roles"] == [(3046, 690)])
    comprobar("Macal: comuna y fecha", f["comuna_texto"] == "La Cisterna" and f["fecha_remate"] == dt.datetime(2026, 10, 14))
    comprobar("Macal: ocupación, superficie, garantía",
              f["ocupacion"] == "DESOCUPADA" and f["superficie_util_m2"] == 43 and f["garantia_pesos"] == 4000000)
    comprobar("Macal: links del listado",
              macal_links_desde_html('<a href="/venta/casa/x/y/P1-1">a</a><a href="/venta/casa/x/y/P1-1">b</a>'
                                     '<a href="/venta/depto/z/w/P22-3">c</a>') ==
              ["/venta/casa/x/y/P1-1", "/venta/depto/z/w/P22-3"])
    tm = tabla_macal([f])
    comprobar("Macal: comuna SII 16110 (La Cisterna)", tm.loc[0, "comuna_sii"] == "16110")

    # --- Oportunidad y selección (datos sintéticos, UF simulada) ---
    global valor_uf
    original_uf = valor_uf
    valor_uf = lambda fecha: (40000.0, "EXACTA")
    try:
        base = []
        hoy = dt.datetime.now() + dt.timedelta(days=10)
        zonas_cod = {"con": "08201", "rm": "13101", "val": "05101", "otra": "04101"}
        k = 0
        for z, cod in zonas_cod.items():
            for j in range(5):
                k += 1
                base.append({"fuente": "MACAL", "id_fuente": f"T{k}", "fecha_remate": hoy, "comuna_sii": cod,
                             "manzana": 100 + k, "predio": 1, "comuna_estado": "OK",
                             "minimo_valor": 1000.0 + 100 * j, "minimo_unidad": "UF",
                             "avaluo_sii": 80_000_000 * (1 + 0.1 * k) + 0, "cruce_estado": "OK",
                             "destino_sii": "H", "fuente_sii": "N", "tipo_sii": "U", "n_roles": 1,
                             "n_roles_encontrados": 1, "direccion_sii": "CALLE X 1", "direccion_tgr": None,
                             "alertas_fuente": ""})
        # una con derechos parciales (no debe entrar), una muy cara (oportunidad negativa)
        base[0]["alertas_fuente"] = "DERECHOS_PARCIALES_O_CUOTA"
        base[0]["avaluo_sii"] = 9e9
        base[1]["minimo_valor"] = 9000.0
        t = pd.DataFrame(base)
        t["avaluo_roles_total"] = None
        r = calcular(t, True)
        comprobar("Conversión UF -> pesos (1200 UF x 40000)", abs(r.loc[r["id_fuente"] == "T3", "minimo_pesos"].iloc[0] - 48_000_000) < 1)
        comprobar("Derechos parciales no es elegible", not r.loc[r["id_fuente"] == "T1", "ELEGIBLE"].iloc[0])
        comprobar("Mínimo caro (oportunidad negativa) no es elegible", not r.loc[r["id_fuente"] == "T2", "ELEGIBLE"].iloc[0])
        r2, falt = seleccionar_top(r, n_top=8, cupos={"PROV:082": 2, "REG:RM": 2, "REG:05": 2})
        top = r2[r2["EN_TOP"]]
        comprobar("El top tiene 8 propiedades", len(top) == 8)
        cuenta = lambda z: int(top["zonas"].map(lambda s: z in s).sum())
        comprobar("Cupo Concepción >= 2", cuenta("PROV:082") >= 2)
        comprobar("Cupo RM >= 2", cuenta("REG:RM") >= 2)
        comprobar("Cupo Valparaíso >= 2", cuenta("REG:05") >= 2)
        comprobar("Cada elegida tiene razón de selección", top["razon_seleccion"].notna().all())

        # --- Doble ranking (avalúo fiscal vs. precio de mercado), con cupos chicos para la prueba ---
        r_dual = r.copy()
        # a la mitad de las propiedades les damos un valor de mercado (más alto que el avalúo,
        # para que destaquen en el Ranking B); a la otra mitad las dejamos solo con avalúo fiscal.
        r_dual["valor_mercado_pesos"] = pd.NA
        con_mercado = r_dual.index[::2]
        r_dual.loc[con_mercado, "valor_mercado_pesos"] = r_dual.loc[con_mercado, "avaluo_sii"] * 3
        r_dual["valor_mercado_estimado_m2_pesos"] = pd.NA
        r_dual["margen_mercado_pesos"] = r_dual["valor_mercado_pesos"] - r_dual["minimo_pesos"]
        r_dual["oportunidad_mercado_pct"] = (r_dual["valor_mercado_pesos"] / r_dual["minimo_pesos"] - 1) * 100
        r_dual.loc[r_dual["valor_mercado_pesos"].isna(), ["oportunidad_mercado_pct", "margen_mercado_pesos"]] = None
        r_dual["ELEGIBLE_MERCADO"] = r_dual["ELEGIBLE"] & r_dual["valor_mercado_pesos"].notna()
        r_dual["ELEGIBLE_FISCAL"] = r_dual["ELEGIBLE"]

        r3, falt3 = seleccionar_top_dual(r_dual, n_fiscal=4, n_mercado=4, n_comodin=2,
                                          cupos={"PROV:082": 1, "REG:RM": 1, "REG:05": 1})
        top3 = r3[r3["EN_TOP"]]
        comprobar("Doble ranking: no pasa del máximo (4+4+2=10)", len(top3) <= 10)
        comprobar("Doble ranking: hay seleccionadas por mercado",
                  top3["razon_seleccion"].str.startswith("OPORTUNIDAD VS. MERCADO", na=False).any())
        comprobar("Doble ranking: hay seleccionadas por avalúo fiscal",
                  top3["razon_seleccion"].str.startswith("OPORTUNIDAD VS. AVALUO FISCAL", na=False).any())
        comprobar("Doble ranking: ninguna propiedad se repite", not top3.index.duplicated().any())

        # ahora simulamos que CASI no hay datos de mercado (solo 1 propiedad) -> el Ranking A debe
        # recibir los cupos que le sobraron al Ranking B ("préstamo" de cupos)
        r_pocos = r_dual.copy()
        r_pocos.loc[r_pocos.index[1:], "ELEGIBLE_MERCADO"] = False
        r4, falt4 = seleccionar_top_dual(r_pocos, n_fiscal=4, n_mercado=4, n_comodin=2,
                                          cupos={"PROV:082": 1, "REG:RM": 1, "REG:05": 1})
        top4 = r4[r4["EN_TOP"]]
        n_fiscal_real = int(top4["razon_seleccion"].str.startswith("OPORTUNIDAD VS. AVALUO FISCAL", na=False).sum())
        comprobar("Doble ranking: si el de mercado está corto, el fiscal recibe los cupos prestados",
                  n_fiscal_real > 4)
        comprobar("Doble ranking: queda registrado el préstamo de cupos en los avisos",
                  any("prestaron" in t for t in falt4))

        import tempfile
        global SALIDAS
        salidas_original = SALIDAS
        SALIDAS = Path(tempfile.mkdtemp())
        try:
            ruta = exportar(r2, None, True, {"PRUEBA": "OK"}, falt)
            import openpyxl
            hojas = openpyxl.load_workbook(ruta).sheetnames
            comprobar("Excel con hojas TOP_CONSOLIDADO, CANDIDATAS, REVISAR, TODAS, RESUMEN",
                      hojas == ["TOP_CONSOLIDADO", "CANDIDATAS", "REVISAR", "TODAS", "RESUMEN"])
            comprobar("CSV del top guardado", (SALIDAS / "TOP_OPORTUNIDADES.csv").exists())
        finally:
            SALIDAS = salidas_original
        tg = pd.DataFrame([{"fuente": "TGR", "id_fuente": "1", "fecha_remate": hoy, "comuna_sii": "08201",
                            "manzana": 1, "predio": 1, "comuna_estado": "OK", "tasacion": 130_000_000,
                            "avaluo_tgr": 100_000_000, "avaluo_sii": 100_000_000, "cruce_estado": "OK",
                            "destino_sii": "H", "fuente_sii": "N", "tipo_sii": "U", "direccion_sii": "X 1",
                            "direccion_tgr": None}])
        rt = calcular(tg, True)
        comprobar("TGR: mínimo estimado = 2/3 de la tasación", abs(rt.loc[0, "minimo_pesos"] - 130e6 * 2 / 3) < 1)
        comprobar("TGR: tipo de mínimo dice ESTIMADO y confianza BAJA",
                  rt.loc[0, "tipo_minimo"].startswith("ESTIMADO") and rt.loc[0, "confianza"] == "BAJA")
    finally:
        valor_uf = original_uf

    # --- Pruebas con el catastro real (si está) ---
    hallados = buscar_archivos_catastro()
    db_path = preparar_catastro() if hallados else None
    if db_path:
        decir("\n  Pruebas con el catastro real:")
        con = sqlite3.connect(db_path)
        n = con.execute("SELECT COUNT(*) FROM predios").fetchone()[0]
        comprobar(f"El catastro tiene muchos predios ({n:,})", n > 1_000_000)
        e, c = buscar_rol_por_direccion(con, "16110", {"nombre": "TRINIDAD RAMIREZ", "numero": 1010, "unidad": ("1207", "C")})
        comprobar(f"Dirección -> ROL: Trinidad Ramírez 1010 depto 1207 ({e})", e == "UNICO" and (c[0][0], c[0][1]) == (3046, 690))
        e, c = buscar_rol_por_direccion(con, "07201", {"nombre": "9 SUR", "numero": 3486, "unidad": None})
        decir(f"    (Talca calle 9 Sur 3486 -> {e} {[(x[0], x[1]) for x in c]})")
        comprobar("Talca 9 Sur 3486 -> ROL 4530-19", e == "UNICO" and (c[0][0], c[0][1]) == (4530, 19))
        e, c = buscar_rol_por_direccion(con, "13167", {"nombre": "DAVID ARELLANO", "numero": 1810, "unidad": None})
        comprobar("Edificio sin unidad: no se confunde con el rol matriz (90032)", not (e == "UNICO" and c[0][1] >= 90000))
        e, _ = buscar_rol_por_direccion(con, "16110", {"nombre": "CALLE QUE NO EXISTE", "numero": 1, "unidad": None})
        comprobar("Calle inexistente -> SIN COINCIDENCIA", e == "SIN COINCIDENCIA")
        con.close()
        # Aviso completo de punta a punta con el catastro
        t = tabla_avisos([]) if False else pd.DataFrame([{
            "fuente": "AVISO", "id_fuente": "p#1", "fecha_remate": dt.datetime.now() + dt.timedelta(days=5),
            "texto_aviso": AVISO_PRUEBA_1, "direccion_texto_busqueda": AVISO_PRUEBA_1,
            "roles_lista": [(3106, 87)], "candidatos_comuna": candidatos_comuna(AVISO_PRUEBA_1),
            "minimo_valor": 83952908, "minimo_unidad": "CLP", "alertas_fuente": "",
            "manzana": None, "predio": None, "comuna_sii": None, "comuna_estado": "NO RECONOCIDA",
            "n_roles": 1}])
        r = resolver_avisos(t, db_path)
        decir(f"    Aviso La Serena rol 3106-87 -> comuna {r.loc[0, 'comuna_sii']}, verificación calle: {r.loc[0, 'verif_direccion_catastro']}")
        comprobar("Aviso: ROL explícito encontrado en el catastro de La Serena", r.loc[0, "comuna_sii"] == "04101" and r.loc[0, "manzana"] == 3106)
        # Prueba de punta a punta: remates TGR simulados con ROL reales
        payload = {"message": "success", "data": [
            {"rol": "15004530019", "rolFormato": "TALCA-04530-019", "direccionRol": "PJE 9 SUR 3486",
             "tasacion": 25_000_000, "avaluo": 19_000_000, "fechaRemate": "2026-12-20T14:00:00.000Z",
             "nombreJuzgado": "PRUEBA", "comunaJuzgado": "TALCA", "nroExpJud": 1, "agnoExpjud": 2026, "codDemanda": 1},
            {"rol": "15099999999", "rolFormato": "TALCA-99999-999", "direccionRol": "CALLE FICTICIA 1",
             "tasacion": 30_000_000, "avaluo": 20_000_000, "fechaRemate": "2026-12-22T14:00:00.000Z",
             "nombreJuzgado": "PRUEBA", "comunaJuzgado": "TALCA", "nroExpJud": 3, "agnoExpjud": 2026, "codDemanda": 3},
        ]}
        tg = construir_tabla_tgr(payload)
        tg["fuente"] = "TGR"
        tg["id_fuente"] = tg["cod_demanda"].astype(str)
        df, hay = procesar([tg, tabla_macal([f])], db_path)
        estados = dict(zip(df["id_fuente"], df["cruce_estado"]))
        comprobar("Talca 4530-19 se cruza con el catastro", estados.get("1", "").startswith("OK"))
        comprobar("ROL inexistente queda 'no encontrado'", estados.get("3") == "ROL NO ENCONTRADO EN CATASTRO")
        comprobar("Macal La Cisterna se cruza (avalúo del catastro)", estados.get("P3170-1", "").startswith("OK"))
    else:
        decir("\n  (No hay archivo de catastro en las carpetas; se omiten las pruebas con datos reales.)")

    # --- Avisos reales (si hay .docx), solo informativo ---
    rutas = buscar_docx_avisos()
    if rutas:
        decir("\n  Lectura de avisos reales (informativo):")
        ta = tabla_avisos(rutas)
        if len(ta):
            decir(f"    {len(ta)} avisos | con fecha {int(ta['fecha_remate'].notna().sum())} | con mínimo "
                  f"{int(ta['minimo_valor'].notna().sum())} | con rol de causa {int(ta['rol_causa'].notna().sum())} | "
                  f"con tribunal {int(ta['tribunal'].notna().sum())} | con ROL de avalúo {int((ta['n_roles'] > 0).sum())}")
            if db_path:
                ra = resolver_avisos(ta.copy(), db_path)
                exp = ra[ra["rol_origen"] == "INFORMADO POR LA FUENTE (aviso)"]
                if len(exp):
                    decir(f"    ROL explícito verificado contra la calle del catastro: "
                          f"{int((exp['verif_direccion_catastro'] == 'COINCIDE').sum())} coinciden, "
                          f"{int((exp['verif_direccion_catastro'] == 'NO COINCIDE').sum())} no coinciden, "
                          f"{int((exp['verif_direccion_catastro'] == 'SIN DATO').sum())} sin dato")
    decir("")
    if fallas:
        decir(f"RESULTADO: {len(fallas)} prueba(s) fallaron: {fallas}")
        guardar_log()
        sys.exit(1)
    decir("RESULTADO: todas las pruebas pasaron.")
    guardar_log()


if __name__ == "__main__":
    try:
        if "--sin-macal" in sys.argv:
            FUENTES["MACAL"] = False
        if "--sin-avisos" in sys.argv:
            FUENTES["AVISOS"] = False
        if "--sin-tgr" in sys.argv:
            FUENTES["TGR"] = False
            FUENTES_OBLIGATORIAS.discard("TGR")
        if "--prueba" in sys.argv:
            prueba()
        elif "--indexar" in sys.argv:
            ruta = preparar_catastro(forzar=("--forzar" in sys.argv))
            decir("Catastro listo." if ruta else f"No encontré el catastro. Déjalo en: {DATOS}")
            guardar_log()
        else:
            principal()
    except SystemExit:
        raise
    except Exception as e:
        decir("\n" + "!" * 70)
        decir(" EL PROGRAMA SE DETUVO")
        decir("!" * 70)
        decir(f"Motivo: {e}")
        decir("\nDetalle técnico (envíame una captura de esto):")
        decir("".join(traceback.format_exc().splitlines(True)[-8:]))
        guardar_log()
        sys.exit(1)
