# -*- coding: utf-8 -*-
"""
Envía dos tipos de correo distintos, a dos destinatarios distintos:

1. CORREO DE RESULTADOS (a CORREO_PARA): la tabla del TOP + el Excel
   adjunto. Este es el que puede llegarle también a Álvaro (u otro
   colaborador) el día que se agregue su correo a CORREO_PARA. NUNCA
   menciona si alguna fuente falló, ni detalles técnicos - solo muestra
   las oportunidades encontradas.

2. CORREO DE ALERTA (a CORREO_ALERTAS): solo se envía si algo salió mal
   esta corrida (una fuente falló, o no se generó Excel). Siempre va a
   CORREO_ALERTAS, que por defecto es la MISMA cuenta que envía
   (CORREO_DESDE) - o sea, Rodrigo se lo manda a sí mismo. Este correo
   NUNCA se mezcla con la lista de CORREO_PARA, así que aunque Álvaro
   esté en CORREO_PARA, jamás recibe este. Pedido explícito de Rodrigo
   (2026-10-09): "quiero que solo me avise a mí en mi correo personal...
   quiero entregarme yo y repararlo en silencio".

No requiere nada especial de Rodrigo para seguir funcionando: si
CORREO_ALERTAS no está configurado, usa CORREO_DESDE automáticamente. Si
algo falla aquí (no hay Excel, falla el envío, etc.), este script NUNCA
debe hacer que el resto del programa se vea como "fallado" - por eso en
el workflow de GitHub se corre con "continue-on-error: true".

Variables de entorno que necesita (se configuran en el workflow):
  GMAIL_APP_PASSWORD  -> la contraseña de aplicación de Gmail (secreto)
  CORREO_DESDE        -> cuenta de Gmail que envía (ej: rbastias1991@gmail.com)
  CORREO_PARA         -> a quién(es) llega el correo de RESULTADOS, separados por coma
  CORREO_ALERTAS      -> (opcional) a quién le llegan las alertas técnicas;
                          si no se configura, se usa CORREO_DESDE (Rodrigo mismo)
"""
import os
import smtplib
import ssl
import sys
from email.message import EmailMessage
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
SALIDAS = BASE / "salidas"
EXCEL = SALIDAS / "OPORTUNIDADES.xlsx"

CORREO_DESDE = os.environ.get("CORREO_DESDE", "").strip()
CORREO_PARA = os.environ.get("CORREO_PARA", "").strip()
CORREO_ALERTAS = os.environ.get("CORREO_ALERTAS", "").strip() or CORREO_DESDE
APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "").strip()

# GitHub Actions deja estas variables puestas solas en cada corrida - sirven
# para armar el link directo a esta corrida, para que Rodrigo pueda ir
# derecho al registro técnico sin tener que buscarlo.
GITHUB_SERVER_URL = os.environ.get("GITHUB_SERVER_URL", "https://github.com")
GITHUB_REPOSITORY = os.environ.get("GITHUB_REPOSITORY", "")
GITHUB_RUN_ID = os.environ.get("GITHUB_RUN_ID", "")
LINK_CORRIDA = (f"{GITHUB_SERVER_URL}/{GITHUB_REPOSITORY}/actions/runs/{GITHUB_RUN_ID}"
                if GITHUB_REPOSITORY and GITHUB_RUN_ID else None)


def v(x):
    # pandas deja NaN (no "") en celdas vacías al leer el Excel con
    # read_excel; "NaN or X" en Python es True (NaN no es falsy), así que
    # sin este filtro una celda vacía terminaba mostrando el texto "nan"
    # en el correo en vez de quedar en blanco. Encontrado en la auditoría
    # del 2026-10-09.
    return "" if pd.isna(x) else x


def construir_resumen_html():
    """Arma la tabla del TOP para el correo de RESULTADOS (el público)."""
    try:
        top = pd.read_excel(EXCEL, sheet_name="TOP_CONSOLIDADO")
    except Exception as e:
        return f"<p>No pude leer la hoja TOP_CONSOLIDADO: {e}</p>", 0

    if top.empty:
        return "<p>Esta corrida no encontró oportunidades que entraran al TOP.</p>", 0

    # Auditoría 2026-10-09: estos nombres de columna quedaron desactualizados
    # (comuna_sii -> comuna_propiedad, link_demanda_pjud -> link_poder_judicial)
    # cuando remates_v2.py cambió de nombre esas columnas; como se leían con
    # r.get(..., "") el correo seguía "funcionando" pero mostrando comuna
    # vacía y "(sin link)" en TODAS las filas, sin ningún error visible. Es
    # justo el tipo de rotura silenciosa que hay que evitar: nadie se entera
    # salvo que mire el correo con lupa. Tribunal/mandante_acreedor son
    # mutuamente excluyentes según la fuente (ver tabla_macal), por eso se
    # muestra el que corresponda en cada fila.
    filas_html = []
    for _, r in top.iterrows():
        direccion = v(r.get("direccion_final", ""))
        comuna = v(r.get("comuna_propiedad", ""))
        pct = v(r.get("oportunidad_pct", ""))
        pct = f"{pct:.1f}" if isinstance(pct, (int, float)) else pct
        tribunal = v(r.get("tribunal", "")) or v(r.get("mandante_acreedor", ""))
        rol = v(r.get("rol_causa", ""))
        link = v(r.get("link_poder_judicial", ""))
        link_html = f'<a href="{link}">Ver causa en el Poder Judicial</a>' if isinstance(link, str) and link.startswith("http") else "(sin link)"
        filas_html.append(
            f"<tr><td>{direccion}</td><td>{comuna}</td><td>{pct}</td>"
            f"<td>{tribunal or ''} {rol or ''}</td><td>{link_html}</td></tr>"
        )

    tabla = (
        "<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-family:sans-serif;font-size:13px'>"
        "<tr style='background:#1F3864;color:white'>"
        "<th>Dirección</th><th>Comuna</th><th>Oportunidad %</th><th>Tribunal / Rol</th><th>Poder Judicial</th></tr>"
        + "".join(filas_html) + "</table>"
    )
    return tabla, len(top)


def detectar_fallas():
    """Lee la hoja RESUMEN del Excel y devuelve la lista de líneas que digan
    que alguna fuente falló esta corrida (p.ej. "Fuente TGR: FALLÓ tras 3
    intentos..."). Esto es solo para el correo de ALERTA (privado) - jamás
    se muestra en el correo de resultados."""
    if not EXCEL.exists():
        return ["No se generó ningún archivo Excel esta corrida (el programa pudo haberse caído antes de llegar a exportar)."]
    try:
        resumen = pd.read_excel(EXCEL, sheet_name="RESUMEN")
    except Exception as e:
        return [f"No pude leer la hoja RESUMEN para revisar si hubo fallas: {e}"]

    fallas = []
    for _, r in resumen.iterrows():
        concepto = str(r.get("concepto", ""))
        valor = str(r.get("valor", ""))
        if concepto.startswith("Fuente") and "FALLÓ" in valor.upper():
            fallas.append(f"{concepto}: {valor}")
    return fallas


def enviar(destinatarios, asunto, html_body, adjuntar_excel):
    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = CORREO_DESDE
    msg["To"] = destinatarios
    msg.set_content("Tu cliente de correo no muestra HTML. Revisa el Excel adjunto si corresponde.")
    msg.add_alternative(html_body, subtype="html")

    if adjuntar_excel and EXCEL.exists():
        msg.add_attachment(
            EXCEL.read_bytes(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="OPORTUNIDADES.xlsx",
        )

    contexto = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as server:
        server.login(CORREO_DESDE, APP_PASSWORD)
        server.send_message(msg)


def main():
    if not APP_PASSWORD or not CORREO_DESDE:
        print("  [CORREO] Faltan variables (CORREO_DESDE / GMAIL_APP_PASSWORD). No envío nada.")
        sys.exit(0)  # no es un error grave, simplemente no manda correo

    fallas = detectar_fallas()

    # --- Correo de ALERTA: solo a Rodrigo (CORREO_ALERTAS), solo si hubo fallas ---
    if fallas:
        detalle = "".join(f"<li>{f}</li>" for f in fallas)
        link_html = f'<p><a href="{LINK_CORRIDA}">Ver el registro completo de esta corrida en GitHub Actions</a></p>' if LINK_CORRIDA else ""
        cuerpo = (
            "<h2>⚠️ Hubo problemas en la corrida de hoy</h2>"
            f"<ul>{detalle}</ul>"
            "<p>Esto no significa que no se haya mandado el correo de resultados - las demás fuentes "
            "que sí funcionaron igual se procesaron. Pero conviene revisar esto pronto.</p>"
            f"{link_html}"
            "<p style='color:#888;font-size:12px'>Correo privado, generado automáticamente - no se envía a nadie más.</p>"
        )
        try:
            enviar(CORREO_ALERTAS, "⚠️ Remates Chile - revisar: hubo fallas en la corrida de hoy", cuerpo, adjuntar_excel=False)
            print(f"  [CORREO-ALERTA] Enviado solo a {CORREO_ALERTAS} ({len(fallas)} falla(s)).")
        except Exception as e:
            print(f"  [CORREO-ALERTA] FALLÓ el envío de la alerta (no detiene el programa): {e}")

    # --- Correo de RESULTADOS: a la lista pública (CORREO_PARA) ---
    if not CORREO_PARA:
        print("  [CORREO] Falta CORREO_PARA. No envío el correo de resultados.")
        return

    if not EXCEL.exists():
        # Nada que mostrar, y no queremos que el correo público diga "no se
        # generó Excel" - eso ya quedó avisado en el correo de alerta privado.
        print("  [CORREO] No se generó Excel esta corrida; no envío el correo de resultados (ya se avisó por alerta privada).")
        return

    tabla_html, n = construir_resumen_html()
    cuerpo = (
        f"<h2>Barrido de remates - resumen</h2>"
        f"<p>Se encontraron <b>{n}</b> propiedades en el TOP de oportunidades.</p>"
        f"{tabla_html}"
        f"<p style='color:#888;font-size:12px'>Correo generado automáticamente.</p>"
    )
    try:
        enviar(CORREO_PARA, f"Remates Chile - {n} oportunidades en el TOP de hoy", cuerpo, adjuntar_excel=True)
        print(f"  [CORREO] Enviado a {CORREO_PARA} ({n} oportunidades).")
    except Exception as e:
        print(f"  [CORREO] FALLÓ el envío (no detiene el programa): {e}")


if __name__ == "__main__":
    main()
