# -*- coding: utf-8 -*-
"""
Envía un correo con el resumen de la corrida y el Excel adjunto.

No requiere nada especial de Rodrigo: lee directo los archivos que ya dejó
"remates_v2.py" en la carpeta "salidas/". Si algo falla aquí (no hay Excel,
falla el envío, etc.), este script NUNCA debe hacer que el resto del
programa se vea como "fallado" - por eso en el workflow de GitHub se corre
con "continue-on-error: true".

Variables de entorno que necesita (se configuran en el workflow):
  GMAIL_APP_PASSWORD  -> la contraseña de aplicación de Gmail (secreto)
  CORREO_DESDE        -> cuenta de Gmail que envía (ej: rbastias1991@gmail.com)
  CORREO_PARA         -> a quién(es) llega, separados por coma
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
APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "").strip()


def construir_resumen_html():
    if not EXCEL.exists():
        return "<p>Esta corrida no generó un Excel (revisa el registro en GitHub Actions).</p>", 0

    try:
        top = pd.read_excel(EXCEL, sheet_name="TOP_CONSOLIDADO")
    except Exception as e:
        return f"<p>No pude leer la hoja TOP_CONSOLIDADO: {e}</p>", 0

    if top.empty:
        return "<p>Esta corrida no encontró oportunidades que entraran al TOP.</p>", 0

    columnas_mostrar = [c for c in [
        "direccion_final", "comuna_sii", "oportunidad_pct", "tribunal",
        "rol_causa", "link_demanda_pjud",
    ] if c in top.columns]

    filas_html = []
    for _, r in top.iterrows():
        direccion = r.get("direccion_final", "")
        comuna = r.get("comuna_sii", "")
        pct = r.get("oportunidad_pct", "")
        tribunal = r.get("tribunal", "")
        rol = r.get("rol_causa", "")
        link = r.get("link_demanda_pjud", "")
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


def main():
    if not APP_PASSWORD or not CORREO_DESDE or not CORREO_PARA:
        print("  [CORREO] Faltan variables (CORREO_DESDE / CORREO_PARA / GMAIL_APP_PASSWORD). No envío nada.")
        sys.exit(0)  # no es un error grave, simplemente no manda correo

    tabla_html, n = construir_resumen_html()

    msg = EmailMessage()
    msg["Subject"] = f"Remates Chile - {n} oportunidades en el TOP de hoy"
    msg["From"] = CORREO_DESDE
    msg["To"] = CORREO_PARA
    msg.set_content("Tu cliente de correo no muestra HTML. Revisa el Excel adjunto.")
    msg.add_alternative(
        f"<h2>Barrido de remates - resumen</h2>"
        f"<p>Se encontraron <b>{n}</b> propiedades en el TOP de oportunidades.</p>"
        f"{tabla_html}"
        f"<p style='color:#888;font-size:12px'>Correo generado automáticamente.</p>",
        subtype="html",
    )

    if EXCEL.exists():
        msg.add_attachment(
            EXCEL.read_bytes(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="OPORTUNIDADES.xlsx",
        )

    try:
        contexto = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as server:
            server.login(CORREO_DESDE, APP_PASSWORD)
            server.send_message(msg)
        print(f"  [CORREO] Enviado a {CORREO_PARA} ({n} oportunidades).")
    except Exception as e:
        print(f"  [CORREO] FALLÓ el envío (no detiene el programa): {e}")


if __name__ == "__main__":
    main()
