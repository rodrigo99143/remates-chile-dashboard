#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PRECIO DE MERCADO POR COMUNA Y TIPO DE PROPIEDAD
==================================================================
Este archivo es un MÓDULO COMPLEMENTARIO de remates_v2.py. Debe guardarse
en LA MISMA CARPETA que remates_v2.py (no en una subcarpeta). remates_v2.py
lo importa automáticamente; no se corre este archivo por sí solo.

QUÉ HACE:
  1. Recorre un grupo de sitios de propiedades ya revisados y autorizados
     por su robots.txt (nunca se agrega un sitio sin haber revisado eso
     primero).
  2. Extrae de cada publicación: comuna, tipo de propiedad, precio y
     metros cuadrados.
  3. Guarda todo en una base de datos propia (archivo .db en la carpeta
     "datos"), sin duplicar publicaciones ya vistas.
  4. Descarta automáticamente publicaciones de más de 2 años.
  5. Calcula el precio promedio del m² por comuna y tipo de propiedad,
     ponderando: año actual = 100%, hace 2 años = 80%.
  6. Le entrega ese resultado a remates_v2.py para que lo use como
     referencia de mercado (en vez de, o además de, el avalúo fiscal).

CONTINGENCIA: si un sitio falla (cambió de estructura, está caído, etc.),
este módulo NUNCA detiene el programa completo. Solo anota que esa fuente
falló hoy y sigue con las demás. La TGR y el resto de remates_v2.py no
dependen de que esto funcione perfecto.

CÓMO SE AMPLÍA A FUTURO: cada fuente nueva se agrega como una función más
en el diccionario PROVEEDORES, al final del archivo. No hace falta tocar
nada más.
"""

import re
import sys
import time
import sqlite3
import subprocess
import importlib
import datetime as dt
import smtplib
from email.mime.text import MIMEText
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# ------------------------------------------------------------------------------
# CONFIGURACIÓN
# ------------------------------------------------------------------------------

# ---- Ponderación por antigüedad (acordado con Rodrigo) ----
PESO_DENTRO_DEL_AÑO = 1.0     # publicaciones de este año o el anterior
PESO_HACE_DOS_AÑOS = 0.8      # publicaciones de hace 2 años
# más de 2 años: se descartan directamente, ni se guardan como válidas para el promedio

# ---- Umbrales de confianza (según cantidad de publicaciones encontradas) ----
MINIMO_CONFIANZA_ALTA = 30
MINIMO_CONFIANZA_MEDIA = 15
MINIMO_CONFIANZA_BAJA = 5
# menos de 5 publicaciones: "SIN DATOS SUFICIENTES" (no se usa)

# ---- Reintentos por fuente ----
MAXIMO_INTENTOS_POR_FUENTE = 3
ESPERA_ENTRE_INTENTOS_SEG = 2
PAUSA_CORTESIA_SEG = 0.5   # pausa entre páginas del mismo sitio, para no sobrecargarlo

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
}

# ---- Correo de alerta (opcional, gratis vía Gmail). Dejar vacío = no se manda correo ----
CORREO_REMITENTE = ""
CONTRASEÑA_DE_APLICACION = ""
CORREO_DESTINO = ""


# ------------------------------------------------------------------------------
# BASE DE DATOS PROPIA (acumula con el tiempo, nunca se pierde si una fuente cae)
# ------------------------------------------------------------------------------

def conectar_bd_mercado(carpeta_datos):
    carpeta_datos.mkdir(exist_ok=True)
    ruta = carpeta_datos / "precio_mercado.db"
    conexion = sqlite3.connect(ruta)
    cursor = conexion.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS publicaciones_mercado (
            id TEXT PRIMARY KEY,
            fuente TEXT,
            comuna TEXT,
            tipo_propiedad TEXT,
            precio_m2 REAL,
            año_publicacion INTEGER,
            url TEXT,
            primera_vez_vista TEXT,
            ultima_vez_vista TEXT
        )
    """)
    conexion.commit()
    return conexion, cursor


def generar_id(*partes):
    import hashlib
    texto = "|".join(str(p) for p in partes)
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()[:20]


# ------------------------------------------------------------------------------
# AVISOS SEMANALES YA PROCESADOS (para la regla "forzar el top 3 de avisos
# NUEVOS", pedida por Rodrigo el 2026-10-09)
# ------------------------------------------------------------------------------

def registrar_avisos_procesados(carpeta_datos, rutas):
    """Para cada archivo .docx de avisos de esta corrida, revisa si YA se
    procesó antes (mismo nombre + mismo contenido, usando un hash). Devuelve
    el conjunto de nombres de archivo que son NUEVOS esta corrida (nunca
    vistos, o con contenido distinto a la última vez), y deja registrados
    TODOS los archivos de esta corrida como "ya procesados" para la próxima
    vez. Vive en la misma base de datos que precio_mercado.db, que ya se
    sincroniza sola con la nube en cada corrida."""
    import hashlib
    conexion, cursor = conectar_bd_mercado(carpeta_datos)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS avisos_procesados (
            archivo TEXT PRIMARY KEY,
            hash TEXT,
            primera_vez_procesado TEXT,
            ultima_vez_visto TEXT
        )
    """)
    conexion.commit()
    hoy = dt.date.today().isoformat()
    nuevos = set()
    for ruta in rutas:
        try:
            contenido = ruta.read_bytes()
        except OSError:
            continue
        hash_archivo = hashlib.sha256(contenido).hexdigest()
        nombre = ruta.name
        cursor.execute("SELECT hash FROM avisos_procesados WHERE archivo = ?", (nombre,))
        fila = cursor.fetchone()
        if fila is None or fila[0] != hash_archivo:
            nuevos.add(nombre)
            cursor.execute("""
                INSERT INTO avisos_procesados (archivo, hash, primera_vez_procesado, ultima_vez_visto)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(archivo) DO UPDATE SET hash=excluded.hash, ultima_vez_visto=excluded.ultima_vez_visto
            """, (nombre, hash_archivo, hoy, hoy))
        else:
            cursor.execute("UPDATE avisos_procesados SET ultima_vez_visto = ? WHERE archivo = ?", (hoy, nombre))
    conexion.commit()
    conexion.close()
    return nuevos


def guardar_publicacion(cursor, fuente, comuna, tipo_propiedad, precio_m2, año_publicacion, url, hoy):
    if not comuna or not tipo_propiedad or not precio_m2 or precio_m2 <= 0:
        return "descartada_datos_incompletos"
    id_pub = generar_id(fuente, url or f"{comuna}-{tipo_propiedad}-{precio_m2}-{año_publicacion}")
    cursor.execute("SELECT id FROM publicaciones_mercado WHERE id = ?", (id_pub,))
    if cursor.fetchone() is not None:
        cursor.execute("UPDATE publicaciones_mercado SET ultima_vez_vista = ? WHERE id = ?", (hoy, id_pub))
        return "ya_existia"
    cursor.execute(
        """INSERT INTO publicaciones_mercado
           (id, fuente, comuna, tipo_propiedad, precio_m2, año_publicacion, url,
            primera_vez_vista, ultima_vez_vista)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (id_pub, fuente, comuna, tipo_propiedad, precio_m2, año_publicacion, url, hoy, hoy),
    )
    return "nueva"


def calcular_peso(año_publicacion, año_actual):
    antiguedad = año_actual - año_publicacion
    if antiguedad <= 1:
        return PESO_DENTRO_DEL_AÑO
    elif antiguedad == 2:
        return PESO_HACE_DOS_AÑOS
    else:
        return 0.0


def nivel_confianza(n):
    if n >= MINIMO_CONFIANZA_ALTA:
        return "ALTA"
    elif n >= MINIMO_CONFIANZA_MEDIA:
        return "MEDIA"
    elif n >= MINIMO_CONFIANZA_BAJA:
        return "BAJA"
    return "SIN DATOS SUFICIENTES"


def tabla_precio_por_comuna_tipo(cursor):
    """Devuelve una lista de dicts: comuna, tipo_propiedad, precio_m2_estimado,
    n_publicaciones, confianza. Usa TODO lo acumulado en la base (no solo
    lo de hoy), aplicando la ponderación por antigüedad."""
    año_actual = dt.date.today().year
    cursor.execute("SELECT comuna, tipo_propiedad, precio_m2, año_publicacion FROM publicaciones_mercado")
    filas = cursor.fetchall()

    acumulado = {}  # (comuna, tipo) -> [suma_ponderada, suma_pesos, cantidad]
    for comuna, tipo, precio_m2, año in filas:
        peso = calcular_peso(año, año_actual)
        if peso <= 0:
            continue
        clave = (comuna, tipo)
        if clave not in acumulado:
            acumulado[clave] = [0.0, 0.0, 0]
        acumulado[clave][0] += precio_m2 * peso
        acumulado[clave][1] += peso
        acumulado[clave][2] += 1

    resultado = []
    for (comuna, tipo), (suma_pond, suma_pesos, n) in acumulado.items():
        if suma_pesos <= 0:
            continue
        resultado.append({
            "comuna": comuna,
            "tipo_propiedad": tipo,
            "precio_m2_estimado": suma_pond / suma_pesos,
            "n_publicaciones": n,
            "confianza_mercado": nivel_confianza(n),
        })
    return resultado


# ------------------------------------------------------------------------------
# ALERTAS (pantalla + correo opcional)
# ------------------------------------------------------------------------------

def alerta_grande(mensaje):
    print()
    print("!" * 70)
    print("!!  ALERTA - PRECIO DE MERCADO")
    print("!" * 70)
    for renglon in mensaje.split("\n"):
        print(f"!!  {renglon}")
    print("!" * 70)
    print()


def enviar_correo_alerta(asunto, cuerpo):
    if not CORREO_REMITENTE or not CONTRASEÑA_DE_APLICACION or not CORREO_DESTINO:
        return False
    try:
        msg = MIMEText(cuerpo)
        msg["Subject"] = asunto
        msg["From"] = CORREO_REMITENTE
        msg["To"] = CORREO_DESTINO
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as servidor:
            servidor.login(CORREO_REMITENTE, CONTRASEÑA_DE_APLICACION)
            servidor.sendmail(CORREO_REMITENTE, [CORREO_DESTINO], msg.as_string())
        return True
    except Exception as e:
        print(f"  [PRECIO_MERCADO] No se pudo enviar el correo de alerta: {e}")
        return False


# ------------------------------------------------------------------------------
# UTILIDADES DE EXTRACCIÓN (compartidas entre proveedores)
# ------------------------------------------------------------------------------

RE_UF = re.compile(r"UF\s*\$?\s*([\d\.]+(?:,\d+)?)", re.IGNORECASE)
RE_CLP = re.compile(r"\$\s*([\d\.]{6,})")
RE_M2 = re.compile(r"([\d\.,]+)\s*m²|([\d\.,]+)\s*m2\b", re.IGNORECASE)

TIPOS_PROPIEDAD = {
    "departamento": "departamento", "depto": "departamento", "depart": "departamento",
    "casa": "casa",
    "parcela": "sitio_eriazo_rural", "sitio": "sitio_eriazo_rural", "terreno": "sitio_eriazo_rural",
    "oficina": "comercial", "local": "comercial", "bodega": "comercial",
}


def _num(texto):
    if not texto:
        return None
    texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def detectar_tipo_propiedad(texto):
    texto_l = texto.lower()
    for clave, tipo in TIPOS_PROPIEDAD.items():
        if clave in texto_l:
            return tipo
    return None


def detectar_precio_uf_o_clp(texto, valor_uf_actual):
    m = RE_UF.search(texto)
    if m:
        uf = _num(m.group(1))
        if uf and valor_uf_actual:
            return uf * valor_uf_actual
    m = RE_CLP.search(texto)
    if m:
        return _num(m.group(1))
    return None


def detectar_m2(texto):
    m = RE_M2.search(texto)
    if m:
        return _num(m.group(1) or m.group(2))
    return None


# ------------------------------------------------------------------------------
# DETECCIÓN DE COMUNA A PARTIR DEL TEXTO DE UNA TARJETA/AVISO
# ------------------------------------------------------------------------------
# Usa la misma lista de 347 comunas de Chile que ya usa remates_v2.py (sin
# tildes, en mayúsculas), para poder cruzar después por nombre de comuna.

_TABLA_TILDES = str.maketrans("ÁÉÍÓÚÑáéíóúñ", "AEIOUNaeioun")


def _quitar_tildes(texto):
    return texto.translate(_TABLA_TILDES)


NOMBRES_COMUNA = [
    "ARICA", "CAMARONES", "IQUIQUE", "PICA", "POZO ALMONTE", "HUARA", "CAMINA",
    "COLCHANE", "ALTO HOSPICIO", "PUTRE", "GENERAL LAGOS", "ANTOFAGASTA", "MEJILLONES",
    "SIERRA GORDA", "TALTAL", "CALAMA", "OLLAGUE", "SAN PEDRO DE ATACAMA", "TOCOPILLA",
    "MARIA ELENA", "COPIAPO", "CALDERA", "TIERRA AMARILLA", "CHAÑARAL", "DIEGO DE ALMAGRO",
    "VALLENAR", "ALTO DEL CARMEN", "FREIRINA", "HUASCO", "LA SERENA", "COQUIMBO",
    "ANDACOLLO", "LA HIGUERA", "PAIGUANO", "VICUÑA", "ILLAPEL", "CANELA", "LOS VILOS",
    "SALAMANCA", "OVALLE", "COMBARBALA", "MONTE PATRIA", "PUNITAQUI", "RIO HURTADO",
    "VALPARAISO", "CASABLANCA", "CONCON", "JUAN FERNANDEZ", "PUCHUNCAVI", "QUINTERO",
    "VIÑA DEL MAR", "ISLA DE PASCUA", "LOS ANDES", "CALLE LARGA", "RINCONADA",
    "SAN ESTEBAN", "LA LIGUA", "CABILDO", "PAPUDO", "PETORCA", "ZAPALLAR", "QUILLOTA",
    "LA CALERA", "HIJUELAS", "LA CRUZ", "NOGALES", "SAN ANTONIO", "ALGARROBO",
    "CARTAGENA", "EL QUISCO", "EL TABO", "SANTO DOMINGO", "QUILPUE", "LIMACHE",
    "OLMUE", "VILLA ALEMANA", "PUTAENDO", "SANTA MARIA", "SAN FELIPE", "LLAILLAY",
    "CATEMU", "PANQUEHUE", "RINCONADA DE SAN FELIPE", "SANTIAGO", "CERRILLOS",
    "CERRO NAVIA", "CONCHALI", "EL BOSQUE", "ESTACION CENTRAL", "HUECHURABA",
    "INDEPENDENCIA", "LA CISTERNA", "LA FLORIDA", "LA GRANJA", "LA PINTANA",
    "LA REINA", "LAS CONDES", "LO BARNECHEA", "LO ESPEJO", "LO PRADO", "MACUL",
    "MAIPU", "ÑUÑOA", "NUNOA", "PEDRO AGUIRRE CERDA", "PEÑALOLEN", "PENALOLEN",
    "PROVIDENCIA", "PUDAHUEL", "QUILICURA", "QUINTA NORMAL", "RECOLETA", "RENCA",
    "SAN JOAQUIN", "SAN MIGUEL", "SAN RAMON", "VITACURA", "PUENTE ALTO", "PIRQUE",
    "SAN JOSE DE MAIPO", "SAN BERNARDO", "CALERA DE TANGO", "BUIN", "PAINE",
    "COLINA", "LAMPA", "TIL TIL", "MELIPILLA", "ALHUE", "CURACAVI", "MARIA PINTO",
    "SAN PEDRO", "TALAGANTE", "EL MONTE", "ISLA DE MAIPO", "PADRE HURTADO",
    "PEÑAFLOR", "PENAFLOR", "RANCAGUA", "CODEGUA", "COLTAUCO", "DOÑIHUE", "DONIHUE",
    "GRANEROS", "LAS CABRAS", "MACHALI", "MALLOA", "MOSTAZAL", "OLIVAR",
    "PEUMO", "PICHIDEGUA", "QUINTA DE TILCOCO", "RENGO", "REQUINOA", "SAN VICENTE",
    "PICHILEMU", "LA ESTRELLA", "LITUECHE", "MARCHIGUE", "NAVIDAD", "PAREDONES",
    "SAN FERNANDO", "CHEPICA", "CHIMBARONGO", "LOLOL", "NANCAGUA", "PALMILLA",
    "PERALILLO", "PLACILLA", "PUMANQUE", "SANTA CRUZ", "TALCA", "CONSTITUCION",
    "CUREPTO", "EMPEDRADO", "MAULE", "PELARCO", "PENCAHUE", "RIO CLARO",
    "SAN CLEMENTE", "SAN RAFAEL", "CAUQUENES", "CHANCO", "PELLUHUE", "LINARES",
    "COLBUN", "LONGAVI", "PARRAL", "RETIRO", "SAN JAVIER", "VILLA ALEGRE",
    "YERBAS BUENAS", "CHILLAN", "BULNES", "CHILLAN VIEJO", "EL CARMEN",
    "PEMUCO", "PINTO", "QUILLON", "SAN IGNACIO", "YUNGAY", "QUIRIHUE",
    "COBQUECURA", "COELEMU", "NINHUE", "PORTEZUELO", "RANQUIL", "TREGUACO",
    "SAN CARLOS", "COIHUECO", "ÑIQUEN", "NIQUEN", "SAN FABIAN", "SAN NICOLAS",
    "CONCEPCION", "CORONEL", "CHIGUAYANTE", "FLORIDA", "HUALPEN", "HUALQUI",
    "LOTA", "PENCO", "SAN PEDRO DE LA PAZ", "SANTA JUANA", "TALCAHUANO",
    "TOME", "LEBU", "ARAUCO", "CAÑETE", "CANETE", "CONTULMO", "CURANILAHUE",
    "LOS ALAMOS", "TIRUA", "LOS ANGELES", "ANTUCO", "CABRERO", "LAJA",
    "MULCHEN", "NACIMIENTO", "NEGRETE", "QUILACO", "QUILLECO", "SAN ROSENDO",
    "SANTA BARBARA", "TUCAPEL", "YUMBEL", "ALTO BIOBIO", "TEMUCO", "CARAHUE",
    "CUNCO", "CURARREHUE", "FREIRE", "GALVARINO", "GORBEA", "LAUTARO",
    "LONCOCHE", "MELIPEUCO", "NUEVA IMPERIAL", "PADRE LAS CASAS", "PERQUENCO",
    "PITRUFQUEN", "PUCON", "SAAVEDRA", "TEODORO SCHMIDT", "TOLTEN", "VILCUN",
    "VILLARRICA", "CHOLCHOL", "ANGOL", "COLLIPULLI", "CURACAUTIN", "ERCILLA",
    "LONQUIMAY", "LOS SAUCES", "LUMACO", "PUREN", "RENAICO", "TRAIGUEN",
    "VICTORIA", "VALDIVIA", "CORRAL", "LANCO", "LOS LAGOS", "MAFIL",
    "MARIQUINA", "PAILLACO", "PANGUIPULLI", "LA UNION", "FUTRONO", "LAGO RANCO",
    "RIO BUENO", "OSORNO", "PUERTO OCTAY", "PURRANQUE", "PUYEHUE", "RIO NEGRO",
    "SAN JUAN DE LA COSTA", "SAN PABLO", "CHAITEN", "FUTALEUFU", "HUALAIHUE",
    "PALENA", "PUERTO MONTT", "CALBUCO", "COCHAMO", "FRESIA", "FRUTILLAR",
    "LOS MUERMOS", "LLANQUIHUE", "MAULLIN", "PUERTO VARAS", "ANCUD",
    "CASTRO", "CHONCHI", "CURACO DE VELEZ", "DALCAHUE", "PUQUELDON",
    "QUEILEN", "QUELLON", "QUEMCHI", "QUINCHAO", "COYHAIQUE", "LAGO VERDE",
    "AYSEN", "CISNES", "GUAITECAS", "COCHRANE", "OHIGGINS", "TORTEL",
    "CHILE CHICO", "RIO IBAÑEZ", "PUNTA ARENAS", "LAGUNA BLANCA", "RIO VERDE",
    "SAN GREGORIO", "CABO DE HORNOS", "ANTARTICA", "PORVENIR", "PRIMAVERA",
    "TIMAUKEL", "NATALES", "TORRES DEL PAINE", "ISLA DE PASCUA",
]
NOMBRES_COMUNA = sorted({_quitar_tildes(n).upper() for n in NOMBRES_COMUNA}, key=lambda n: -len(n))
_RE_COMUNAS = [(n, re.compile(r"\b" + re.escape(n) + r"\b")) for n in NOMBRES_COMUNA]


def detectar_comuna(texto):
    """Busca el nombre de alguna comuna chilena dentro del texto (título,
    dirección, etc. de un aviso). Devuelve el nombre más largo que calce,
    para evitar falsos positivos cortos (ej. preferir 'SAN PEDRO DE LA PAZ'
    antes que una coincidencia parcial)."""
    texto_norm = _quitar_tildes(texto).upper()
    for nombre, patron in _RE_COMUNAS:
        if patron.search(texto_norm):
            return nombre
    return None


def obtener_valor_uf_aproximado():
    """Valor UF aproximado para convertir precios en UF a pesos. Si falla,
    usa un valor de respaldo razonable (hay que actualizarlo de tanto en
    tanto a mano si se usa el respaldo)."""
    try:
        r = requests.get("https://mindicador.cl/api/uf", headers=HEADERS, timeout=10)
        if r.status_code == 200:
            return r.json()["serie"][0]["valor"]
    except Exception:
        pass
    return 39000.0  # valor de respaldo aproximado - revisar si se usa seguido


# ------------------------------------------------------------------------------
# PROVEEDORES (uno por sitio). Cada uno devuelve una lista de publicaciones:
# [{"comuna": ..., "tipo_propiedad": ..., "precio_m2": ..., "año_publicacion": ..., "url": ...}, ...]
# ------------------------------------------------------------------------------

DIAGNOSTICO_MERCADO = True  # True = imprime, para cada URL, en qué paso se pierden los avisos
                            # (útil mientras arreglamos los 9 proveedores en cero). Una vez que
                            # estén todos funcionando bien, se puede dejar en False.


def _tarjetas_genericas(html, base_url, comuna_fija=None):
    """Extractor genérico: busca bloques de texto que tengan a la vez un
    precio y una superficie en m². Sirve como primera aproximación para
    sitios con HTML simple (sin JavaScript).

    Si 'comuna_fija' viene informado (porque la página ya es de una comuna
    específica, ej. ".../venta/casa/concepcion/0"), se usa directamente.
    Si no, se intenta reconocer el nombre de la comuna dentro del propio
    texto de la tarjeta (título, dirección, etc.)."""
    soup = BeautifulSoup(html, "html.parser")
    valor_uf = obtener_valor_uf_aproximado()
    resultados = []
    año_actual = dt.date.today().year

    # Candidatos: tarjetas/bloques típicos de listados (article, li, div con clase que suene a "card"/"item"/"propiedad")
    posibles = soup.find_all(["article", "li", "div"], limit=400)
    n_candidatos_tamano_ok = n_con_m2_y_precio = n_con_tipo = n_con_comuna = 0
    ejemplos_sin_precio_m2 = []   # textos de ejemplo cuando no se detecta precio+m2
    ejemplos_sin_comuna = []      # textos de ejemplo cuando sí hay tipo pero no comuna
    for bloque in posibles:
        texto = bloque.get_text(" ", strip=True)
        if len(texto) < 20 or len(texto) > 600:
            continue
        if "arriendo" in texto.lower():
            continue  # solo nos interesa precio de VENTA, no de arriendo
        n_candidatos_tamano_ok += 1
        m2 = detectar_m2(texto)
        precio = detectar_precio_uf_o_clp(texto, valor_uf)
        if not m2 or not precio or m2 <= 0:
            if len(ejemplos_sin_precio_m2) < 3:
                ejemplos_sin_precio_m2.append(texto[:200])
            continue
        n_con_m2_y_precio += 1
        tipo = detectar_tipo_propiedad(texto)
        if not tipo:
            continue
        n_con_tipo += 1
        comuna = comuna_fija or detectar_comuna(texto)
        if not comuna:
            if len(ejemplos_sin_comuna) < 3:
                ejemplos_sin_comuna.append(texto[:200])
            continue
        n_con_comuna += 1
        comuna = _quitar_tildes(comuna).upper()  # debe calzar con SII2NOMBRE de remates_v2.py
        enlace = bloque.find("a", href=True)
        url = enlace["href"] if enlace else base_url
        resultados.append({
            "comuna": comuna,
            "tipo_propiedad": tipo,
            "precio_m2": precio / m2,
            "año_publicacion": año_actual,  # por defecto, "recién visto" = año actual
            "url": url,
        })

    if DIAGNOSTICO_MERCADO:
        print(f"    [DIAG] {base_url}")
        print(f"    [DIAG]   HTML recibido: {len(html):,} caracteres | bloques article/li/div en la página: {len(posibles)}")
        print(f"    [DIAG]   de tamaño razonable (20-600 caracteres, sin 'arriendo'): {n_candidatos_tamano_ok}")
        print(f"    [DIAG]   -> con precio Y m2 detectados: {n_con_m2_y_precio}")
        print(f"    [DIAG]   -> de esos, con tipo de propiedad reconocido: {n_con_tipo}")
        print(f"    [DIAG]   -> de esos, con comuna reconocida: {n_con_comuna} (= {len(resultados)} avisos aprovechables)")
        if n_con_m2_y_precio == 0 and ejemplos_sin_precio_m2:
            print(f"    [DIAG]   EJEMPLOS de texto que NO se reconoció como precio+m2 (para ajustar la búsqueda):")
            for i, ej in enumerate(ejemplos_sin_precio_m2, 1):
                print(f"    [DIAG]     #{i}: {ej!r}")
        if n_con_tipo > 0 and n_con_comuna == 0 and ejemplos_sin_comuna:
            print(f"    [DIAG]   EJEMPLOS de texto con tipo reconocido pero SIN comuna detectada:")
            for i, ej in enumerate(ejemplos_sin_comuna, 1):
                print(f"    [DIAG]     #{i}: {ej!r}")
    return resultados


# Comunas usadas para recorrer chilepropiedades.cl (que exige comuna en la URL):
# cubren las 3 zonas con cupo obligatorio (Concepción, RM, Valparaíso) más un
# grupo de ciudades grandes de otras regiones, para tener una base más amplia.
_COMUNAS_OBJETIVO_CHILEPROPIEDADES = [
    ("concepcion", "CONCEPCION"), ("talcahuano", "TALCAHUANO"),
    ("san-pedro-de-la-paz", "SAN PEDRO DE LA PAZ"), ("chiguayante", "CHIGUAYANTE"),
    ("hualpen", "HUALPEN"), ("coronel", "CORONEL"),
    ("santiago", "SANTIAGO"), ("providencia", "PROVIDENCIA"),
    ("las-condes", "LAS CONDES"), ("nunoa", "NUNOA"), ("maipu", "MAIPU"),
    ("puente-alto", "PUENTE ALTO"), ("la-florida", "LA FLORIDA"), ("vitacura", "VITACURA"),
    ("valparaiso", "VALPARAISO"), ("vina-del-mar", "VIÑA DEL MAR"), ("concon", "CONCON"),
    ("quilpue", "QUILPUE"), ("san-antonio", "SAN ANTONIO"),
    ("la-serena", "LA SERENA"), ("antofagasta", "ANTOFAGASTA"), ("temuco", "TEMUCO"),
    ("rancagua", "RANCAGUA"), ("talca", "TALCA"),
]
_TIPOS_CHILEPROPIEDADES = ["casa", "departamento"]
PAUSA_CHILEPROPIEDADES_SEG = 2.0  # robots.txt pide Crawl-delay: 2


def proveedor_chilepropiedades():
    """HTML simple, confirmado viable por robots.txt. La página general
    '/propiedades/venta' no existe (da 404): hay que pedir comuna y tipo
    puntual, ej. '/propiedades/venta/casa/concepcion/0'. Recorre un grupo
    de comunas fijas (ver _COMUNAS_OBJETIVO_CHILEPROPIEDADES) y primera
    página de resultados de cada una."""
    resultados = []
    for slug, nombre in _COMUNAS_OBJETIVO_CHILEPROPIEDADES:
        for tipo in _TIPOS_CHILEPROPIEDADES:
            url = f"https://chilepropiedades.cl/propiedades/venta/{tipo}/{slug}/0"
            try:
                r = requests.get(url, headers=HEADERS, timeout=20)
                r.raise_for_status()
                resultados.extend(_tarjetas_genericas(r.text, url, comuna_fija=nombre))
            except Exception:
                pass  # una comuna que falle no debe tumbar todo el proveedor
            time.sleep(PAUSA_CHILEPROPIEDADES_SEG)
    return resultados


def proveedor_toppropiedades():
    """Agregador de cientos de corredoras chicas - confirmado viable.
    Es un listado nacional (no exige comuna en la URL); la comuna se
    reconoce desde el propio texto de cada aviso."""
    url = "https://www.toppropiedades.cl/propiedades/venta"
    r = requests.get(url, headers=HEADERS, timeout=20)
    r.raise_for_status()
    return _tarjetas_genericas(r.text, url)


def proveedor_procasa():
    """Confirmado viable - ojo: sesgo hacia segmento alto (RM oriente).
    '/propiedades/venta' no existe; la ruta real exige tipo/operación/comuna.
    Se usa 'Venta' (no 'Venta_y_Arriendo') para traer solo propiedades en
    venta y no arriendos.
    Arreglado el 08-10-2026: el diagnóstico mostró que el pedido normal solo
    traía el pie de página (teléfono y correo), no las propiedades - el
    sitio las carga con JavaScript. Se cambió a usar el navegador automático."""
    url = "https://www.procasa.cl/Todos_los_tipos/Venta/Todas_las_comunas"
    return _recorrer_urls_renderizado([url])


def _recorrer_urls(urls, pausa_seg=1.0, timeout_seg=20, comuna_fija=None):
    """Helper compartido: pide una lista de URLs (de un mismo sitio) y junta
    los avisos encontrados en todas. Si una URL puntual falla (404, timeout,
    etc.) no se pierde el resto - se sigue con las demás. Se usa para sitios
    que no tienen una sola página 'todas las propiedades', sino que hay que
    recorrer varias categorías/regiones."""
    resultados = []
    for url in urls:
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout_seg)
            r.raise_for_status()
            resultados.extend(_tarjetas_genericas(r.text, url, comuna_fija=comuna_fija))
        except Exception as e:
            if DIAGNOSTICO_MERCADO:
                print(f"    [DIAG] {url} -> FALLÓ el pedido: {type(e).__name__}: {e}")
        time.sleep(pausa_seg)
    return resultados


def _recorrer_urls_renderizado(urls, espera_ms=4000, comuna_fija=None, pausa_seg=1.0):
    """Igual que _recorrer_urls, pero para sitios que cargan el listado de
    propiedades con JavaScript - el pedido normal (requests) solo trae el
    'cascarón' de la página (menú, pie de página), sin las propiedades.
    Usa el navegador automático (Playwright, el mismo que ya se usaba para
    Property Partners) para esperar a que cargue el contenido real. Es más
    lento, por eso se usa solo en los sitios que de verdad lo necesitan."""
    resultados = []
    for url in urls:
        try:
            html = _html_renderizado(url, espera_ms=espera_ms)
            resultados.extend(_tarjetas_genericas(html, url, comuna_fija=comuna_fija))
        except Exception as e:
            if DIAGNOSTICO_MERCADO:
                print(f"    [DIAG] {url} -> FALLÓ el pedido (navegador): {type(e).__name__}: {e}")
        time.sleep(pausa_seg)
    return resultados


def proveedor_doomos():
    """Portal de clasificados nacional - confirmado viable por robots.txt.
    Arreglado el 08-10-2026: el diagnóstico mostró que el pedido normal solo
    traía el menú de navegación, no las propiedades - el sitio las carga con
    JavaScript. Se cambió a usar el navegador automático."""
    urls = [
        "https://www.doomos.cl/venta-casas",
        "https://www.doomos.cl/venta-departamentos",
        "https://www.doomos.cl/venta-terrenos",
    ]
    return _recorrer_urls_renderizado(urls)


def proveedor_inmoclick():
    """Portal con cobertura en varias regiones - confirmado viable."""
    urls = [
        "https://www.inmoclick.cl/casas-en-venta",
        "https://www.inmoclick.cl/departamentos-en-venta",
    ]
    return _recorrer_urls(urls)


def proveedor_prourbano():
    """Cubre Santiago, Valparaíso, Viña del Mar, Reñaca, Concón.
    Arreglado el 08-10-2026: el diagnóstico mostró que el pedido normal solo
    traía el correo y teléfono de contacto, no las propiedades - el sitio
    las carga con JavaScript. Se cambió a usar el navegador automático."""
    urls = [
        "https://www.prourbano.cl/s/casa/venta?id_property_type=1&business_type%5B0%5D=for_sale",
        "https://www.prourbano.cl/s/departamento/venta?id_property_type=2&business_type%5B0%5D=for_sale",
    ]
    return _recorrer_urls_renderizado(urls)


def proveedor_bernal():
    """Bernal Propiedades - cubre V Región y RM - confirmado viable."""
    urls = [
        "https://bernalpropiedades.cl/resultados-de-la-busqueda/?es_category%5B0%5D=10",
    ]
    return _recorrer_urls(urls)


def proveedor_paz():
    """Inmobiliaria Paz - desarrolladora (proyectos nuevos). Su robots.txt
    autoriza explícitamente a ClaudeBot por nombre - el mejor caso posible.
    OJO: muestra precios 'desde UF x' por proyecto, no siempre m2 por unidad,
    así que puede aportar menos publicaciones que un portal de corredores."""
    urls = ["https://www.paz.cl/proyectos"]
    return _recorrer_urls(urls)


def proveedor_siena():
    """Siena Inmobiliaria - desarrolladora, zona central. Precios por
    proyecto, no siempre por unidad.
    Arreglado el 08-10-2026: el diagnóstico mostró que el pedido normal solo
    traía el menú de navegación, no las propiedades - el sitio las carga con
    JavaScript. Se cambió a usar el navegador automático."""
    urls = [
        "https://www.siena.cl/casas/",
        "https://www.siena.cl/departamentos/",
    ]
    return _recorrer_urls_renderizado(urls)


def proveedor_castro():
    """Castro Propiedades - cobertura real: RM (Quilicura, San Bernardo) y
    Araucanía (Carahue, Villarrica) - no la IV Región que anunciaba la lista
    original. Precios en CLP. URL confirmada por Rodrigo el 08-10-2026 (el
    sitio está montado en una subcarpeta /wordpress/, por eso la URL antigua
    sin esa parte daba 404).
    Arreglado también el 08-10-2026: el diagnóstico mostró que el pedido
    normal solo traía el encabezado de la página (título, "Asesoría en
    Compra y Venta"), no las fichas de propiedades - el listado se carga con
    JavaScript. Se cambió a usar el navegador automático."""
    urls = ["https://castropropiedades.cl/wordpress/index.php/propiedades/"]
    return _recorrer_urls_renderizado(urls)


def proveedor_maitencillo():
    """Maitencillo Propiedades - litoral central (Zapallar, Cachagua,
    Maitencillo) - confirmado viable, cubre una zona que no teníamos.
    Arreglado el 08-10-2026: el diagnóstico mostró que SÍ detectaba precio,
    m2 y tipo de propiedad bien, pero descartaba todo por no encontrar el
    nombre de ninguna comuna escrito en el texto del aviso (el sitio solo
    dice 'Maitencillo', 'Cachagua', etc., que son sectores, no comunas).
    Maitencillo y Cachagua pertenecen a la comuna de Puchuncaví (Región de
    Valparaíso), así que se fija esa comuna directamente."""
    urls = [
        "https://maitencillopropiedades.cl/casa/",
        "https://maitencillopropiedades.cl/departamentos/",
    ]
    return _recorrer_urls(urls, comuna_fija="PUCHUNCAVI")


def proveedor_grupo_premium():
    """Grupo Premium - nacional. El dominio real que responde es gpremium.cl
    (grupopremium.cl redirige ahí).
    Arreglado el 08-10-2026: el diagnóstico mostró que el pedido normal solo
    traía el menú de navegación, no las propiedades - el sitio las carga con
    JavaScript. Se cambió a usar el navegador automático."""
    urls = ["https://gpremium.cl/comprar-una-propiedad/"]
    return _recorrer_urls_renderizado(urls)


def proveedor_coproch():
    """COPROCH (gremio de corredores) - segundo gremio nacional además de
    ACOP, con filtro propio de 'Venta' en el sitio.
    Arreglado el 08-10-2026: el diagnóstico mostró que el pedido normal solo
    traía el pie de página y el formulario de login del tema de WordPress,
    no las propiedades - el listado se carga con JavaScript. Se cambió a
    usar el navegador automático."""
    urls = ["https://www.coproch.cl/propiedades/"]
    return _recorrer_urls_renderizado(urls)


def proveedor_dosam():
    """Dosam Propiedades - foco norte (Arica, Iquique) + Santiago y
    Valparaíso - confirmado viable. Cubre una zona difícil de encontrar
    en otras fuentes."""
    urls = [
        "https://dosampropiedades.cl/?al_product-cat=casas-2",
        "https://dosampropiedades.cl/?al_product-cat=departamentos-2",
    ]
    return _recorrer_urls(urls)


def proveedor_engel_volkers():
    """Engel & Völkers Chile - segmento premium, varias oficinas regionales
    (Concepción, Antofagasta, Chicureo) - confirmado viable por robots.txt."""
    urls = [
        "https://www.engelvoelkers.com/cl/en/properties/res/sale/apartment",
        "https://www.engelvoelkers.com/cl/en/properties/res/sale/house",
    ]
    return _recorrer_urls(urls)


# ------------------------------------------------------------------------------
# PORTAL INMOBILIARIO (portalinmobiliario.com, de Mercado Libre) - el sitio de
# propiedades más grande de Chile, con mucha diferencia. Rodrigo encontró que
# tiene su propia "Referencia de precios" por propiedad, pero esa página
# individual (la "VIP") está BLOQUEADA en su robots.txt (Disallow: /vip/).
# En cambio, las páginas de LISTADO por comuna sí están permitidas (se
# revisó el robots.txt completo el 08-10-2026 y se confirmó el patrón de URL
# en 3 comunas reales: Concepción->biobio, Las Condes->metropolitana,
# Viña del Mar->valparaiso). Así que se usan esas páginas de listado, igual
# que con los demás proveedores, y el promedio por comuna lo calculamos
# nosotros con lo ya acumulado (no se usa el gráfico de ellos).
#
# Las comunas de Biobío, Metropolitana y Valparaíso están confirmadas. Las
# demás (La Serena, Antofagasta, Temuco, Rancagua, Talca) usan el nombre de
# su región en minúsculas como mejor estimación del patrón - si alguna falla
# (404), no tumba el proveedor completo, solo esa comuna puntual; queda
# pendiente confirmar la URL real si el diagnóstico muestra que falló.
# ------------------------------------------------------------------------------
_COMUNAS_OBJETIVO_PORTALINMOBILIARIO = [
    ("concepcion-biobio", "CONCEPCION"), ("talcahuano-biobio", "TALCAHUANO"),
    ("san-pedro-de-la-paz-biobio", "SAN PEDRO DE LA PAZ"), ("chiguayante-biobio", "CHIGUAYANTE"),
    ("hualpen-biobio", "HUALPEN"), ("coronel-biobio", "CORONEL"),
    ("santiago-metropolitana", "SANTIAGO"), ("providencia-metropolitana", "PROVIDENCIA"),
    ("las-condes-metropolitana", "LAS CONDES"), ("nunoa-metropolitana", "NUNOA"),
    ("maipu-metropolitana", "MAIPU"), ("puente-alto-metropolitana", "PUENTE ALTO"),
    ("la-florida-metropolitana", "LA FLORIDA"), ("vitacura-metropolitana", "VITACURA"),
    ("valparaiso-valparaiso", "VALPARAISO"), ("vina-del-mar-valparaiso", "VIÑA DEL MAR"),
    ("concon-valparaiso", "CONCON"), ("quilpue-valparaiso", "QUILPUE"),
    ("san-antonio-valparaiso", "SAN ANTONIO"),
    # ---- no confirmadas todavía (mejor estimación del patrón) ----
    ("la-serena-coquimbo", "LA SERENA"), ("antofagasta-antofagasta", "ANTOFAGASTA"),
    ("temuco-araucania", "TEMUCO"), ("rancagua-ohiggins", "RANCAGUA"),
    ("talca-maule", "TALCA"),
]
_TIPOS_PORTALINMOBILIARIO = ["casa", "departamento"]
PAUSA_PORTALINMOBILIARIO_SEG = 2.0  # cortesía - es un sitio grande, mejor no apurarlo


def proveedor_portalinmobiliario():
    """Portal Inmobiliario (Mercado Libre) - recorre páginas de listado por
    comuna y tipo. Página individual de cada propiedad NO se usa (bloqueada
    por robots.txt)."""
    resultados = []
    for slug, nombre in _COMUNAS_OBJETIVO_PORTALINMOBILIARIO:
        for tipo in _TIPOS_PORTALINMOBILIARIO:
            url = f"https://www.portalinmobiliario.com/venta/{tipo}/{slug}"
            try:
                r = requests.get(url, headers=HEADERS, timeout=20)
                r.raise_for_status()
                resultados.extend(_tarjetas_genericas(r.text, url, comuna_fija=nombre))
            except Exception as e:
                if DIAGNOSTICO_MERCADO:
                    print(f"    [DIAG] {url} -> FALLÓ el pedido: {type(e).__name__}: {e}")
            time.sleep(PAUSA_PORTALINMOBILIARIO_SEG)
    return resultados


# ------------------------------------------------------------------------------
# NAVEGADOR AUTOMÁTICO (Playwright) - solo para sitios que cargan el listado
# con JavaScript, como Property Partners. Es el mismo mecanismo que ya usa
# remates_v2.py para Macal: se instala solo la primera vez, no requiere que
# hagas nada tú. Más lento que un pedido normal, por eso se usa solo cuando
# de verdad hace falta.
# ------------------------------------------------------------------------------

_navegador_pw = None
_navegador_contexto = None
_navegador_pagina = None


def _obtener_pagina_navegador():
    global _navegador_pw, _navegador_contexto, _navegador_pagina
    if _navegador_pagina is not None:
        return _navegador_pagina
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        subprocess.call([sys.executable, "-m", "pip", "install", "--break-system-packages", "-q", "playwright"])
        importlib.invalidate_caches()
        from playwright.sync_api import sync_playwright
    subprocess.call([sys.executable, "-m", "playwright", "install", "chromium"],
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    _navegador_pw = sync_playwright().start()
    navegador = _navegador_pw.chromium.launch(headless=True)
    _navegador_contexto = navegador.new_context(user_agent=HEADERS["User-Agent"])
    _navegador_pagina = _navegador_contexto.new_page()
    return _navegador_pagina


def _cerrar_navegador():
    global _navegador_pw, _navegador_contexto, _navegador_pagina
    try:
        if _navegador_contexto:
            _navegador_contexto.close()
        if _navegador_pw:
            _navegador_pw.stop()
    except Exception:
        pass
    _navegador_pw = _navegador_contexto = _navegador_pagina = None


def _html_renderizado(url, espera_ms=4000):
    """Abre la URL con un navegador real (headless) y devuelve el HTML ya
    renderizado, después de esperar a que cargue el contenido dinámico."""
    pagina = _obtener_pagina_navegador()
    pagina.goto(url, wait_until="networkidle", timeout=60000)
    pagina.wait_for_timeout(espera_ms)  # margen extra para listados que cargan por partes
    return pagina.content()


# Búsquedas de Property Partners que cubren las 3 zonas con cupo obligatorio
# más Biobío (confirmado viable con la ficha de ejemplo que mandó Rodrigo).
_BUSQUEDAS_PROPERTY_PARTNERS = [
    "https://ppartnersgroup.com/es-cl/search/venta/departamento/region-metropolitana-de-santiago",
    "https://ppartnersgroup.com/es-cl/search/venta/casa/region-metropolitana-de-santiago",
    "https://ppartnersgroup.com/es-cl/search/venta/departamento/region-de-valparaiso",
    "https://ppartnersgroup.com/es-cl/search/venta/casa/region-de-valparaiso",
    "https://ppartnersgroup.com/es-cl/search/venta/departamento/region-del-bio-bio",
    "https://ppartnersgroup.com/es-cl/search/venta/casa/region-del-bio-bio",
]


def proveedor_property_partners():
    """Property Partners - segmento premium, cobertura nacional e
    internacional. Su listado de resultados se carga con JavaScript (el
    HTML simple solo trae los filtros, no las propiedades), así que acá SÍ
    usamos el navegador automático (Playwright) en vez de un pedido normal.
    Es más lento que los demás proveedores por eso mismo.
    NOTA: esta fuente no se pudo probar en vivo durante el desarrollo (el
    entorno donde se escribió este código no tiene salida a internet hacia
    sitios externos) - revisar el resultado de la primera corrida real con
    atención."""
    resultados = []
    for url in _BUSQUEDAS_PROPERTY_PARTNERS:
        try:
            html = _html_renderizado(url)
            resultados.extend(_tarjetas_genericas(html, url))
        except Exception:
            pass  # una búsqueda que falle no debe tumbar todo el proveedor
    return resultados


# ---- Revisadas y DESCARTADAS esta ronda (no se agregan, por las razones
#      indicadas - no es una omisión, es una decisión explícita): ----
#
# Urbalia (urbalia.cl): su robots.txt bloquea TODAS las rutas del sitio.
#   Aunque una lista anterior decía "permite sin restricciones", al revisar
#   de nuevo ahora el sitio cambió su robots.txt y deniega todo. Se respeta
#   tal cual, sin excepción.
#
# Remax Chile (remax.cl): la página no trae el listado en el HTML (lo carga
#   con JavaScript, como le pasó a Tattersall Propiedades y a Property
#   Partners) - a diferencia de Property Partners, no se probó todavía con
#   el navegador automático. Queda pendiente para una próxima ronda.
#
# Property Partners: SÍ se agregó (ver proveedor_property_partners), usando
#   el navegador automático (Playwright) porque su listado carga con
#   JavaScript. Es la única fuente de esta lista que lo necesita.
#
# Enlace Inmobiliario (enlaceinmobiliario.cl): SACADO el 08-10-2026. El sitio
#   cambió por completo de enfoque - ya no muestra un listado de propiedades
#   para recorrer, ahora es una herramienta de simulación de crédito
#   hipotecario que pide buscar comuna por comuna. Ya no calza con la forma
#   en que este programa recorre los sitios (confirmado dando 403 Forbidden
#   en las URLs antiguas). Se puede reconsiderar si en el futuro vuelve a
#   tener un listado navegable.


# ---- Pendientes de implementar (ya vetados por robots.txt, falta escribir
#      el extractor específico de cada uno). Se van agregando uno por uno;
#      no hace falta terminarlos todos de una vez. ----
# proveedor_doomos, proveedor_enlace_inmobiliario, proveedor_inmoclick,
# proveedor_urbalia, proveedor_prourbano, proveedor_bernal, proveedor_paz,
# proveedor_siena, proveedor_grupo_premium, proveedor_coproch,
# proveedor_dosam, proveedor_castro, proveedor_maitencillo,
# proveedor_engel_volkers, proveedor_property_partners, proveedor_remax,
# proveedor_acop (requiere JavaScript - más avanzado, dejar para después)

PROVEEDORES = {
    "ChilePropiedades": {"funcion": proveedor_chilepropiedades, "critica": False},
    "TopPropiedades": {"funcion": proveedor_toppropiedades, "critica": False},
    "Procasa": {"funcion": proveedor_procasa, "critica": False},
    "Doomos": {"funcion": proveedor_doomos, "critica": False},
    "Inmoclick": {"funcion": proveedor_inmoclick, "critica": False},
    "ProUrbano": {"funcion": proveedor_prourbano, "critica": False},
    "Bernal": {"funcion": proveedor_bernal, "critica": False},
    "Paz": {"funcion": proveedor_paz, "critica": False},
    "Siena": {"funcion": proveedor_siena, "critica": False},
    "Castro": {"funcion": proveedor_castro, "critica": False},
    "Maitencillo": {"funcion": proveedor_maitencillo, "critica": False},
    "GrupoPremium": {"funcion": proveedor_grupo_premium, "critica": False},
    "COPROCH": {"funcion": proveedor_coproch, "critica": False},
    "Dosam": {"funcion": proveedor_dosam, "critica": False},
    "EngelVolkers": {"funcion": proveedor_engel_volkers, "critica": False},
    "PropertyPartners": {"funcion": proveedor_property_partners, "critica": False},
    # "PortalInmobiliario" se sacó del barrido DIARIO (08-10-2026): esa fuente ya
    # se alimenta del barrido NACIONAL semestral (actualizar_portal_inmobiliario_nacional.py),
    # que es mucho más completo. El chequeo diario de 24 comunas aportaba poco
    # (los precios no cambian mucho en 6 meses) y además está bloqueado en la nube
    # (Portal Inmobiliario detecta el servidor y no entrega datos). La función
    # proveedor_portalinmobiliario() se deja escrita más abajo por si se quiere
    # reactivar en el futuro.
}


# ------------------------------------------------------------------------------
# ORQUESTADOR: corre todos los proveedores, con reintentos, sin detener nunca
# el programa, y actualiza la base de datos.
# ------------------------------------------------------------------------------

def actualizar_precio_mercado(carpeta_datos, decir=print):
    """Punto de entrada que llama remates_v2.py. Devuelve:
       (tabla_comuna_tipo, estado_por_proveedor)
    'tabla_comuna_tipo' es una lista de dicts lista para usar en calcular().
    'estado_por_proveedor' es un dict con el resultado de cada fuente, para
    mostrar en el resumen final (igual que TGR/MACAL/AVISOS)."""
    conexion, cursor = conectar_bd_mercado(carpeta_datos)
    hoy = dt.date.today().isoformat()
    estado = {}

    for nombre, info in PROVEEDORES.items():
        decir(f"  [PRECIO_MERCADO: {nombre}] barriendo...")
        ultimo_error = None
        publicaciones = None
        for intento in range(1, MAXIMO_INTENTOS_POR_FUENTE + 1):
            try:
                publicaciones = info["funcion"]()
                break
            except Exception as e:
                ultimo_error = str(e)
                if intento < MAXIMO_INTENTOS_POR_FUENTE:
                    time.sleep(ESPERA_ENTRE_INTENTOS_SEG)

        if publicaciones is None:
            estado[f"PRECIO_MERCADO:{nombre}"] = f"FALLÓ tras {MAXIMO_INTENTOS_POR_FUENTE} intentos: {ultimo_error}"
            decir(f"  [PRECIO_MERCADO: {nombre}] FALLÓ (el programa sigue sin esta fuente): {ultimo_error}")
            if info["critica"]:
                alerta_grande(
                    f"La fuente de precio de mercado '{nombre}' (marcada como critica) fallo hoy.\n"
                    f"El programa sigue funcionando con las demas fuentes de precio de mercado,\n"
                    f"y con lo que ya habia acumulado en la base de datos de dias anteriores.\n"
                    f"Motivo: {ultimo_error}"
                )
                enviar_correo_alerta(f"[AVISO] Fuente de precio de mercado caída: {nombre}",
                                      f"La fuente '{nombre}' fallo. Motivo: {ultimo_error}")
            continue

        nuevas = ya_vistas = descartadas = 0
        for pub in publicaciones:
            if not pub.get("comuna"):
                descartadas += 1
                continue
            resultado = guardar_publicacion(
                cursor, nombre, pub["comuna"], pub["tipo_propiedad"], pub["precio_m2"],
                pub["año_publicacion"], pub.get("url"), hoy,
            )
            if resultado == "nueva":
                nuevas += 1
            elif resultado == "ya_existia":
                ya_vistas += 1
            else:
                descartadas += 1
        conexion.commit()
        estado[f"PRECIO_MERCADO:{nombre}"] = (
            f"OK ({nuevas} nuevas, {ya_vistas} ya vistas, {descartadas} descartadas)")
        decir(f"  [PRECIO_MERCADO: {nombre}] OK: {nuevas} nuevas, {ya_vistas} ya vistas, {descartadas} descartadas")

    tabla = tabla_precio_por_comuna_tipo(cursor)
    conexion.close()
    _cerrar_navegador()  # si se usó Playwright (Property Partners), lo cerramos al terminar
    return tabla, estado
