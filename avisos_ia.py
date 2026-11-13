# -*- coding: utf-8 -*-
"""
Rescate puntual con IA (Claude Haiku) para avisos donde el regex se quedó
sin ROL de avalúo.  (Pedido de Rodrigo, 11-11-2026, después de la cuenta de
Álvaro de la billetera para esto.)

Por qué existe: en la llamada del 09-11-2026, Álvaro confirmó que el aviso de
Río Negro / Puerto Montt SÍ traía el ROL en el texto original - el programa
no lo encontró porque el guion largo ("–") se borra al normalizar ANTES de
que el regex del ROL lo intente leer. Es un bug real, no un caso sin dato.
Álvaro mismo fue claro en que no exige automatizar todo ("si no tenemos el
rol, no nos sirve de nada" - le parece bien descartar lo que de verdad no
tiene el dato) - pero tampoco hay que perder lo que SÍ estaba escrito solo
porque el regex no lo pudo leer.

Cómo se usa esto (filosofía "nunca se esconde nada por error", igual que el
resto del programa):
  - El regex (extraer_roles / parsear_aviso, en remates_v2.py) sigue siendo
    el método principal y el único que se usa mientras encuentre un ROL.
    Esta IA NUNCA se llama si el regex ya encontró algo.
  - Solo se llama, aviso por aviso, cuando ese aviso puntual quedó con
    roles=[] después del regex.
  - Sin ANTHROPIC_API_KEY configurada (variable de entorno), esta función no
    hace nada - el programa sigue exactamente igual que antes de que esto
    existiera, sin gastar ni un centavo ni cambiar ningún resultado.
  - El resultado se trata EXACTAMENTE igual que si lo hubiera encontrado el
    regex: pasa por las mismas validaciones de resolver_avisos() en
    remates_v2.py (cruce contra el catastro, verificación de la dirección).
    Si la IA se equivoca, ese cruce lo detecta igual que detectaría un error
    del regex - esta función no se salta ningún control que ya existía.
  - Se cachea por el hash del texto del aviso en precio_mercado.db (la misma
    base que ya persiste sola entre corridas de GitHub Actions), para que una
    corrida que no cambió ese aviso no vuelva a gastar en la API.
  - Todo aviso rescatado así queda marcado con la alerta "RESCATADO_CON_IA"
    (se agrega a alertas_fuente) para que quede siempre visible que ese ROL
    no vino del regex, nunca oculto.
"""
import hashlib
import json
import os
import re
import sqlite3

# Modelo más económico disponible en la cuenta de Rodrigo (console de
# Anthropic, noviembre 2026). Si más adelante conviene cambiarlo, basta con
# cambiar esta constante - no queda hardcodeado en ningún otro lado.
MODELO_IA = "claude-haiku-5-5"

# OJO (bug real, detectado el 11-11-2026, segunda corrida del día): la primera
# corrida con este módulo falló SIEMPRE (ver más abajo, el bug de Haiku 5.5
# "pensando" todo el presupuesto) - y ese resultado vacío/fallido se guardó
# en la caché igual que un resultado bueno. Entonces, al corregir el bug y
# correr de nuevo, el programa ni siquiera volvió a llamar a la IA para esos
# 564 avisos: encontró la respuesta vacía YA cacheada de la corrida rota, y
# se quedó con eso - "con ROL" salió exactamente igual (223) en las dos
# corridas. Para que este arreglo (y cualquier arreglo futuro al prompt o al
# modelo) invalide la caché vieja SIN tener que entrar a mano a borrar filas
# de precio_mercado.db, cada resultado se guarda con esta versión - si no
# coincide con la que está guardada, se trata como si no hubiera caché (se
# vuelve a intentar). Subir este número cada vez que se cambie algo que
# pueda hacer que una respuesta vieja ya no sea confiable.
VERSION_CACHE = 2

_cliente_cache = {}

# OJO (12-11-2026): Rodrigo preguntó por qué este módulo, ya pagado y ya
# conectado, parecía no haber hecho nada en ninguna corrida real - y la
# respuesta honesta es que, tal como estaba escrito, un rescate exitoso NO
# deja ningún rastro visible: solo se imprime algo en consola cuando algo
# FALLA (sin cliente, o la llamada a la IA revienta). Si todo funcionaba
# bien, quedaba en silencio total. Estas estadísticas existen para que la
# próxima corrida diga, con números reales, qué pasó - sin tener que
# adivinar a partir de la ausencia de líneas de error.
ESTADISTICAS = {
    "sin_cliente": 0,       # no hay API key (o la librería/clave fallaron) - ni se intentó llamar a la IA
    "cache_con_rol": 0,     # ya se había rescatado antes este mismo aviso (incluso de otra corrida)
    "cache_vacio": 0,       # ya se había consultado antes y la IA de verdad no encontró nada
    "nuevo_con_rol": 0,     # llamada nueva a la IA esta corrida, SÍ encontró ROL
    "nuevo_vacio": 0,       # llamada nueva a la IA esta corrida, no encontró ROL (dato genuinamente ausente)
    "fallido": 0,           # la llamada a la IA en sí falló (red, modelo inválido, etc.) - se reintenta la próxima vez
}


def resumen_estadisticas():
    """Copia de ESTADISTICAS al momento de llamarla - para que remates_v2.py
    la deje en el RESUMEN del Excel y en el log de la corrida."""
    return dict(ESTADISTICAS)


def reiniciar_estadisticas():
    """Vuelve todos los contadores a cero - se llama una vez al principio de
    cada corrida real (tabla_avisos() en remates_v2.py), por si el proceso
    reutiliza el módulo más de una vez (como en --prueba)."""
    for k in ESTADISTICAS:
        ESTADISTICAS[k] = 0


def _cliente():
    """Devuelve el cliente de Anthropic, o None si no hay API key o falla la
    inicialización - en cualquiera de los dos casos, rescatar() simplemente
    no hace nada (ver docstring del módulo)."""
    if "c" in _cliente_cache:
        return _cliente_cache["c"]
    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    c = None
    if api_key:
        try:
            import anthropic
            c = anthropic.Anthropic(api_key=api_key)
        except Exception as e:
            print(f"  [IA-AVISOS] no se pudo preparar el cliente de Anthropic ({e}); sigo sin rescate con IA, como si no hubiera key.")
    _cliente_cache["c"] = c
    return c


def _tabla_cache(con):
    con.execute("""
        CREATE TABLE IF NOT EXISTS cache_ia_avisos (
            hash TEXT PRIMARY KEY,
            resultado TEXT,
            fecha TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)


def _hash_texto(texto):
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()


def _leer_cache(cache_db_path, h):
    if not cache_db_path:
        return None
    try:
        con = sqlite3.connect(cache_db_path)
        _tabla_cache(con)
        fila = con.execute("SELECT resultado FROM cache_ia_avisos WHERE hash=?", (h,)).fetchone()
        con.close()
        if fila:
            datos = json.loads(fila[0])
            if datos.get("_version") != VERSION_CACHE:
                return None  # caché de una versión anterior (posiblemente de un bug ya corregido) - se reintenta
            datos["roles"] = [tuple(r) for r in datos.get("roles", [])]  # JSON guarda listas, no tuplas
            return datos
    except Exception as e:
        print(f"  [IA-AVISOS] no pude leer la cache ({e}); sigo igual, solo que esta vez puede volver a llamar a la IA de más.")
    return None


def _guardar_cache(cache_db_path, h, resultado):
    if not cache_db_path:
        return
    try:
        con = sqlite3.connect(cache_db_path)
        _tabla_cache(con)
        con.execute("INSERT OR REPLACE INTO cache_ia_avisos (hash, resultado) VALUES (?, ?)",
                    (h, json.dumps({**resultado, "_version": VERSION_CACHE}, ensure_ascii=False)))
        con.commit()
        con.close()
    except Exception as e:
        print(f"  [IA-AVISOS] no pude guardar en la cache ({e}); sigo igual, solo que la próxima corrida va a volver a llamar a la IA para este mismo aviso.")


PROMPT = """Este es un aviso de remate judicial chileno publicado en un diario. \
El extractor automático del programa (basado en expresiones regulares) no \
encontró ningún "ROL de avalúo" en este texto (el identificador del predio \
ante el SII - "manzana-predio", suele aparecer como "Rol de Avalúo N.o \
1234-56" o parecido, a veces con el guion mal escrito o cortado por un salto \
de línea). OJO: el "Rol de avalúo" NO es el "Rol de la causa" del tribunal \
(ese es distinto, tipo "C-1234-2026" o "V-99-2026") - ignora ese.

Lee el siguiente texto y responde SOLO con un JSON (sin texto adicional, sin \
bloque de markdown) con esta forma exacta:
{{"roles": [[manzana, predio], ...], "comuna": "NOMBRE COMUNA o null", "confianza": "ALTA|MEDIA|BAJA"}}

- "roles": lista de pares [manzana, predio] (números enteros) de cada ROL de \
avalúo que encuentres, aunque esté mal separado o con un guion distinto al \
normal. Si de verdad no hay ningún ROL de avalúo en el texto, responde \
"roles": [].
- "comuna": el nombre de la comuna de Chile donde está UBICADA la propiedad \
(no la comuna del tribunal, salvo que sea la misma). null si no se puede \
determinar con razonable certeza.
- "confianza": tu propia evaluación de qué tan segura es tu lectura del/los \
ROL(es) que extrajiste.

Texto del aviso:
---
{texto}
---"""


def rescatar(texto, cache_db_path=None):
    """Intenta rescatar el/los ROL de avalúo (y la comuna) de un aviso cuyo
    regex se quedó sin ROL. Devuelve None si no hay API key, si la IA falla,
    o si la IA tampoco encuentra nada - en los tres casos el aviso sigue su
    camino normal, exactamente como si esta función no existiera."""
    cliente = _cliente()
    if cliente is None:
        ESTADISTICAS["sin_cliente"] += 1
        return None

    h = _hash_texto(texto)
    cacheado = _leer_cache(cache_db_path, h)
    if cacheado is not None:
        if cacheado.get("roles"):
            ESTADISTICAS["cache_con_rol"] += 1
        else:
            ESTADISTICAS["cache_vacio"] += 1
        return cacheado if cacheado.get("roles") else None

    # OJO (este fue el bug real del 11-11-2026, segunda parte): antes, si la
    # llamada a la IA fallaba por cualquier motivo técnico (el bug de
    # "pensamiento" de más arriba, un corte de red, lo que sea), igual se
    # guardaba un resultado vacío en la caché - como si la IA de verdad hubiera
    # mirado el aviso y no hubiera encontrado nada. Eso "envenenaba" la caché
    # para siempre: aunque se corrigiera el bug después, la próxima corrida
    # encontraba ese resultado vacío YA guardado y ni siquiera volvía a llamar
    # a la IA. Ahora: un fallo técnico (la llamada en sí falla) NO se guarda en
    # la caché - se reintenta en la próxima corrida. Solo se cachea una
    # respuesta que la IA de verdad llegó a dar (sea "encontré esto" o
    # "de verdad no hay nada"), para no repetir el gasto en avisos que
    # genuinamente no tienen ROL.
    try:
        resp = cliente.messages.create(
            model=MODELO_IA,
            max_tokens=1024,
            output_config={"effort": "low"},
            messages=[{"role": "user", "content": PROMPT.format(texto=texto[:4000])}],
        )
        bruto = "".join(b.text for b in resp.content if getattr(b, "type", None) == "text").strip()
        if not bruto:
            raise ValueError(f"la respuesta no tuvo bloque de texto (stop_reason={resp.stop_reason!r}, "
                              f"tipos de bloque recibidos: {[getattr(b, 'type', None) for b in resp.content]!r})")
        bruto = re.sub(r"^```(json)?|```$", "", bruto, flags=re.MULTILINE).strip()
        datos = json.loads(bruto)
        roles = [(int(m), int(p)) for m, p in (datos.get("roles") or [])
                 if m is not None and p is not None]
        resultado = {"roles": roles, "comuna": datos.get("comuna"), "confianza": datos.get("confianza")}
    except Exception as e:
        ESTADISTICAS["fallido"] += 1
        print(f"  [IA-AVISOS] no se pudo rescatar este aviso con IA ({e}); sigue sin ROL, como antes "
              f"(no se guarda en la caché - se vuelve a intentar en la próxima corrida).")
        return None

    _guardar_cache(cache_db_path, h, resultado)
    if resultado["roles"]:
        ESTADISTICAS["nuevo_con_rol"] += 1
    else:
        ESTADISTICAS["nuevo_vacio"] += 1
    return resultado if resultado["roles"] else None
