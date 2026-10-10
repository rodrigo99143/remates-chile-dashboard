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

    # Auditoría 2026-10-09: estos nombres de columna quedaron desactualizados
    # (comuna_sii -> comuna_propiedad, link_demanda_pjud -> link_poder_judicial)
    # cuando remates_v2.py cambió de nombre esas columnas; como se leían con
    # r.get(..., "") el correo seguía "funcionando" pero mostrando comuna
    # vacía y "(sin link)" en TODAS las filas, sin ningún error visible. Es
    # justo el tipo de rotura silenciosa que hay que evitar: nadie se entera
    # salvo que mire el correo con lupa. Tribunal/mandante_acreedor son
    # mutuamente excluyentes según la fuente (ver tabla_macal), por eso se
    # muestra el que corresponda en cada fila.
    columnas_mostrar = [c for c in [
        "direccion_final", "comuna_propiedad", "oportunidad_pct", "tribunal",
        "mandante_acreedor", "rol_causa", "link_poder_judicial",
    ] if c in top.columns]

    def v(x):
        # pandas deja NaN (no "") en celdas vacías al leer el Excel con
        # read_excel; "NaN or X" en Python es True (NaN no es falsy), así que
        # sin este filtro una celda vacía terminaba mostrando el texto "nan"
        # en el correo en vez de quedar en blanco. Otra rotura silenciosa
        # encontrada en la misma auditoría del 2026-10-09.
        return "" if pd.isna(x) else x

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
