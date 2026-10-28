# -*- coding: utf-8 -*-
"""
Reporte "de una sola vez" (2026-10-28) para ponerse al día: Rodrigo pidió un
Excel y un PDF con TODO lo que apareció como oportunidad GENUINAMENTE NUEVA
desde el 2026-10-20 (la última corrida de la que de verdad llegó a tener
Excel+PDF en la mano) hasta hoy - sin que importe cuántas corridas de prueba
se hicieron en el medio mientras se arreglaba el envío del correo.

Cómo decide qué es "nuevo desde esa fecha": usa la tabla candidatas_historico
(la misma que usa remates_v2.py para TOP_NUEVO) y se queda con las claves
cuya "primera_vez" (la primera vez que esa propiedad, identificada por su
ROL + fecha de remate, apareció en CUALQUIER corrida) sea >= la fecha dada.
Eso es distinto de "nuevo_hoy" (que compara solo contra la corrida
INMEDIATAMENTE anterior) - por eso una corrida normal no sirve para esto: si
se corrió el programa varias veces seguidas hoy, "nuevo_hoy" sale vacío
aunque haya mucho acumulado desde el 20.

Se apoya en los archivos que ya dejó "Correr el programa" en esta misma
corrida (salidas/OPORTUNIDADES.xlsx, hoja CANDIDATAS = TODO el universo de
candidatas elegibles de HOY) y en datos/precio_mercado.db (ya actualizado
por esa misma corrida) - no vuelve a barrer ninguna fuente ni gasta tiempo
extra de internet aparte de geocodificar las direcciones nuevas para el PDF.

Uso: python reporte_acumulado.py 2026-10-20
"""
import os
import smtplib
import ssl
import sqlite3
import sys
import datetime as dt
from email.message import EmailMessage
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
SALIDAS = BASE / "salidas"
DATOS = BASE / "datos"
EXCEL_HOY = SALIDAS / "OPORTUNIDADES.xlsx"
DB = DATOS / "precio_mercado.db"

CORREO_DESDE = os.environ.get("CORREO_DESDE", "").strip()
CORREO_PARA = os.environ.get("CORREO_PARA", "").strip()
APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "").strip()


def _clave_real(f):
    # Misma fórmula exacta que usa remates_v2.py (ver _clave_real ahí) - así
    # las claves calzan con las que ya quedaron guardadas en candidatas_historico.
    rol = f.get("rol_completo")
    fecha_remate = f.get("fecha_remate")
    fecha = fecha_remate.date() if pd.notna(fecha_remate) else "sf"
    if pd.notna(rol) and rol:
        return f"{rol}|{fecha}"
    return f"{f.get('direccion_final')}|{fecha}"


def main():
    if len(sys.argv) != 2:
        print("Uso: python reporte_acumulado.py AAAA-MM-DD")
        sys.exit(1)
    fecha_desde = sys.argv[1]

    if not EXCEL_HOY.exists():
        print(f"  [ACUMULADO] No existe {EXCEL_HOY} (parece que 'Correr el programa' no llegó a generar el Excel esta corrida). Sin esto no se puede armar el reporte acumulado.")
        sys.exit(1)
    if not DB.exists():
        print(f"  [ACUMULADO] No existe {DB} (la base de precio de mercado). Sin esto no se puede saber qué es 'nuevo desde el {fecha_desde}'.")
        sys.exit(1)

    cand = pd.read_excel(EXCEL_HOY, sheet_name="CANDIDATAS")
    print(f"  [ACUMULADO] {len(cand)} candidatas elegibles en el universo de hoy.")

    cand["clave_real"] = cand.apply(_clave_real, axis=1)

    conexion = sqlite3.connect(DB)
    primeras_vez = dict(conexion.execute(
        "SELECT clave, primera_vez FROM candidatas_historico"
    ).fetchall())
    conexion.close()

    cand["primera_vez"] = cand["clave_real"].map(primeras_vez)
    acumulado = cand[cand["primera_vez"] >= fecha_desde].copy()
    acumulado = acumulado.drop(columns=["clave_real"])
    print(f"  [ACUMULADO] {len(acumulado)} de esas son nuevas desde el {fecha_desde} (primera_vez registrada en esa fecha o después).")

    if acumulado.empty:
        print("  [ACUMULADO] No hay ninguna propiedad nueva desde esa fecha. No se genera ni se envía nada.")
        return

    acumulado = acumulado.sort_values(["fecha_remate", "oportunidad_pct"], ascending=[True, False])

    import remates_v2
    import generar_pdf

    acumulado = generar_pdf.geocodificar_top_nuevo(acumulado)

    ruta_excel = SALIDAS / "REPORTE_ACUMULADO.xlsx"
    ruta_pdf = SALIDAS / "REPORTE_ACUMULADO.pdf"
    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as w:
        hoja = remates_v2.ordenar_columnas(acumulado)
        hoja.to_excel(w, sheet_name="ACUMULADO", index=False)
        remates_v2.dar_formato(w.sheets["ACUMULADO"])
    print(f"  [ACUMULADO] Excel: {ruta_excel.name} ({len(acumulado)} filas)")

    try:
        generar_pdf.construir_pdf(acumulado, ruta_pdf)
        print(f"  [ACUMULADO] PDF: {ruta_pdf.name}")
    except Exception as e:
        print(f"  [ACUMULADO] no se pudo generar el PDF ({e}); el Excel se manda igual.")

    if not APP_PASSWORD or not CORREO_DESDE or not CORREO_PARA:
        print("  [ACUMULADO] Faltan variables de correo; el Excel/PDF quedaron guardados en 'salidas/' pero no se mandó nada por correo.")
        return

    tamano_excel = ruta_excel.stat().st_size if ruta_excel.exists() else 0
    tamano_pdf = ruta_pdf.stat().st_size if ruta_pdf.exists() else 0
    limite = 18 * 1024 * 1024
    omitir_pdf = ruta_pdf.exists() and (tamano_excel + tamano_pdf) > limite

    msg = EmailMessage()
    msg["Subject"] = f"Remates Chile - reporte acumulado desde el {fecha_desde} ({len(acumulado)} oportunidades)"
    msg["From"] = CORREO_DESDE
    msg["To"] = CORREO_PARA
    aviso_pdf = ("<p style='color:#A05A00'>El PDF quedó muy pesado para enviarlo por correo; "
                 "queda disponible para descargar en los resultados de esta corrida en GitHub Actions.</p>") if omitir_pdf else ""
    cuerpo = (
        f"<h2>Reporte acumulado - todo lo nuevo desde el {fecha_desde}</h2>"
        f"<p>Se juntaron <b>{len(acumulado)}</b> propiedades que aparecieron como oportunidad "
        f"nueva desde el {fecha_desde} hasta hoy ({dt.date.today().isoformat()}), "
        f"sin importar cuántas corridas se hicieron en el medio.</p>"
        f"{aviso_pdf}"
        f"<p style='color:#888;font-size:12px'>Correo generado automáticamente - reporte de una sola vez, pedido el 2026-10-28.</p>"
    )
    msg.set_content("Tu cliente de correo no muestra HTML. Revisa el Excel adjunto.")
    msg.add_alternative(cuerpo, subtype="html")

    msg.add_attachment(
        ruta_excel.read_bytes(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="REPORTE_ACUMULADO.xlsx",
    )
    if ruta_pdf.exists() and not omitir_pdf:
        msg.add_attachment(
            ruta_pdf.read_bytes(),
            maintype="application",
            subtype="pdf",
            filename="REPORTE_ACUMULADO.pdf",
        )

    try:
        contexto = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as server:
            server.login(CORREO_DESDE, APP_PASSWORD)
            server.send_message(msg)
        print(f"  [ACUMULADO] Enviado a {CORREO_PARA}.")
    except Exception as e:
        print(f"  [ACUMULADO] FALLÓ el envío: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
