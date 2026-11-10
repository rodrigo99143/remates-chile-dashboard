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
    pool = cand[cand["primera_vez"] >= fecha_desde].copy()
    pool = pool.drop(columns=["clave_real"])
    print(f"  [ACUMULADO] {len(pool)} de esas son nuevas desde el {fecha_desde} (primera_vez registrada en esa fecha o después).")

    if pool.empty:
        print("  [ACUMULADO] No hay ninguna propiedad nueva desde esa fecha. No se genera ni se envía nada.")
        return

    # OJO (pedido de Rodrigo/Álvaro, 09-11-2026): a diferencia del primer intento
    # (2026-10-28), esta vez NO se cura a un top de 50 - el software de Álvaro
    # necesita TODAS las elegibles nuevas, no solo las mejores ("una cosa es el
    # ranking de las mejores para decidir, y otra que Álvaro necesita incluir
    # todo lo nuevo, todo, en su software"). El PDF sí sigue acotado a las 30
    # mejores (son fichas con mapa, uno por propiedad - mandar cientas sería un
    # PDF enorme y poco práctico de todas formas), pero el Excel va completo.
    MAX_PDF = 30
    acumulado = pool.sort_values("oportunidad_pct", ascending=False)
    print(f"  [ACUMULADO] las {len(acumulado)} elegibles nuevas van COMPLETAS al Excel (sin tope).")

    import remates_v2
    import generar_pdf

    acumulado = generar_pdf.geocodificar_top_nuevo(acumulado)  # coordenadas (se usan en el Excel también)

    # ---- Hoja REVISION (pedido de Rodrigo, 09-11-2026): las propiedades que
    # quedaron bloqueadas (derechos parciales/cuota, dirección no coincide, dato
    # sospechoso, etc.) no deben perderse - deben quedar visibles en algún lado
    # para que alguien las revise a mano, en vez de desaparecer sin dejar rastro.
    # OJO/LIMITACIÓN conocida: a diferencia de las elegibles (que sí tienen
    # historial de "primera vez vista" desde antes), todavía no existe ese
    # historial para las NO elegibles anterior al 09-11-2026 - así que esta hoja
    # trae TODAS las que hoy quedaron en revisión (no se puede filtrar de forma
    # confiable "solo las nuevas desde el 5-11" para estas todavía; de aquí en
    # adelante sí va a poder, gracias al historial nuevo que ya se empezó a
    # registrar hoy).
    try:
        revision = pd.read_excel(EXCEL_HOY, sheet_name="REVISAR")
        print(f"  [ACUMULADO] {len(revision)} propiedades quedan en la hoja REVISION (bloqueadas o con alerta, "
              f"universo completo de hoy - no se puede acotar por fecha todavía para estas).")
    except Exception as e:
        revision = pd.DataFrame()
        print(f"  [ACUMULADO] no pude leer la hoja REVISAR de hoy ({e}); la hoja REVISION queda vacía.")

    ruta_excel = SALIDAS / "REPORTE_ACUMULADO.xlsx"
    ruta_pdf = SALIDAS / "REPORTE_ACUMULADO.pdf"
    # OJO: "acumulado" y "revision" se leyeron (más arriba) directamente de las hojas
    # CANDIDATAS y REVISAR de OPORTUNIDADES.xlsx - y esas hojas YA vienen en el formato
    # "definitivo" de Álvaro (columnas reducidas y renombradas, pedido 09-11-2026, ver
    # ordenar_columnas_alvaro/RENOMBRES_DISPLAY en remates_v2.py). Por eso aquí NO se
    # vuelve a reordenar/renombrar columnas - eso ya viene hecho, y si se repitiera con
    # los nombres internos (ej. "avaluo_sii") no encontraría nada porque esas columnas
    # ya se llaman "Avalúo fiscal", etc. - y las borraría por error.
    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as w:
        acumulado.to_excel(w, sheet_name="ACUMULADO", index=False)
        remates_v2.dar_formato(w.sheets["ACUMULADO"])
        if not revision.empty:
            revision.to_excel(w, sheet_name="REVISION", index=False)
            remates_v2.dar_formato(w.sheets["REVISION"])
    print(f"  [ACUMULADO] Excel: {ruta_excel.name} ({len(acumulado)} elegibles nuevas + {len(revision)} en REVISION)")

    acumulado_pdf = acumulado.head(MAX_PDF)
    try:
        generar_pdf.construir_pdf(acumulado_pdf, ruta_pdf)
        print(f"  [ACUMULADO] PDF: {ruta_pdf.name} ({len(acumulado_pdf)} fichas, las mejores del Excel)")
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
    msg["Subject"] = (f"Remates Chile - reporte acumulado desde el {fecha_desde} "
                       f"({len(acumulado)} elegibles nuevas + {len(revision)} en revisión)")
    msg["From"] = CORREO_DESDE
    msg["To"] = CORREO_PARA
    aviso_pdf = ("<p style='color:#A05A00'>El PDF quedó muy pesado para enviarlo por correo; "
                 "queda disponible para descargar en los resultados de esta corrida en GitHub Actions.</p>") if omitir_pdf else ""
    cuerpo = (
        f"<h2>Reporte acumulado - todo lo nuevo desde el {fecha_desde}</h2>"
        f"<p>Desde el {fecha_desde} hasta hoy ({dt.date.today().isoformat()}) este reporte trae, en la hoja "
        f"<b>ACUMULADO</b>, las <b>{len(acumulado)}</b> propiedades elegibles nuevas (TODAS, sin tope - no solo "
        f"las mejores), con dirección y coordenadas de cada una. El PDF adjunto trae fichas con mapa solo de "
        f"las mejores {min(len(acumulado), MAX_PDF)}, para no ser demasiado pesado.</p>"
        f"<p>En la hoja <b>REVISION</b> van, aparte, <b>{len(revision)}</b> propiedades que quedaron bloqueadas o "
        f"con alguna alerta (derechos parciales, dirección no coincide, dato sospechoso, etc.) - no son "
        f"oportunidades confirmadas, pero quedan ahí para que alguien las revise a mano en vez de perderse.</p>"
        f"{aviso_pdf}"
        f"<p style='color:#888;font-size:12px'>Correo generado automáticamente - reporte de una sola vez, pedido el 2026-11-09.</p>"
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
