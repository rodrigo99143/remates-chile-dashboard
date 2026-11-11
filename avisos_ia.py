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

_cliente_cache = {}


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
                    (h, json.dumps(resultado, ensure_ascii=False)))
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
        return None

    h = _hash_texto(texto)
    cacheado = _leer_cache(cache_db_path, h)
    if cacheado is not None:
        return cacheado if cacheado.get("roles") else None

    resultado = {"roles": [], "comuna": None, "confianza": None}
    try:
        # OJO (bug real, detectado en la corrida del 11-11-2026): Haiku 5.5 piensa
        # "adaptativamente" aunque no se le pida - sin esto, max_tokens=300 se
        # gastaba TODO en el bloque de "pensamiento" y nunca llegaba a escribir la
        # respuesta (content[0] era un ThinkingBlock sin .text, y las 564 llamadas
        # de esa corrida fallaron todas por eso). Ahora: max_tokens más holgado
        # para que le alcance para pensar poco Y responder, output_config con
        # effort "low" para que piense lo mínimo posible (más rápido y barato), y
        # se busca el bloque de tipo "text" por tipo, no por posición - un
        # ThinkingBlock puede seguir apareciendo antes del texto.
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
        print(f"  [IA-AVISOS] no se pudo rescatar este aviso con IA ({e}); sigue sin ROL, como antes.")

    _guardar_cache(cache_db_path, h, resultado)
    return resultado if resultado["roles"] else None
