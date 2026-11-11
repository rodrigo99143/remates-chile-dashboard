# -*- coding: utf-8 -*-
"""
Reporte especial "de una sola vez" (pedido de Rodrigo, 11-11-2026): para UN
informe de avisos puntual (el que se pase como argumento), procesa TODOS sus
avisos - completos o no - SIN aplicar el filtro de "no repetir lo que ya se
vio sin cambios" que usa la corrida normal (remates_v2.py). Antes de tener el
rescate con IA, casi ninguno de esos avisos tenía ROL y "no servía de nada"
(como dijo Álvaro); ahora Rodrigo quiere escrutarlos TODOS de una vez, con el
rescate ya puesto, para ver realmente qué se pudo sacar de cada uno.

De paso (mismo pedido, 11-11-2026), agrega una hoja con las propiedades de
TGR que sean GENUINAMENTE nuevas desde una fecha dada (por ejemplo, desde el
lunes o martes) - a diferencia de "nuevo_hoy" (que solo compara contra la
corrida inmediatamente anterior de hoy), esto usa el mismo mecanismo de
"primera_vez" que ya usa reporte_acumulado.py, para no perder TGR nuevo
aunque el programa se haya corrido varias veces entre medio.

No toca el historial real del programa (candidatas_historico / universo
histórico) - es sólo un reporte de lectura, aparte del Excel de todos los
días.

Uso: python reporte_informe_fijo.py avisos/INFORME_06-11-2026_AL_08-11-2026.docx 2026-11-09
"""
import os
import smtplib
import ssl
import sqlite3
import sys
from email.message import EmailMessage
from pathlib import Path

import pandas as pd

import remates_v2

BASE = Path(__file__).resolve().parent
SALIDAS = BASE / "salidas"
DATOS = BASE / "datos"
EXCEL_HOY = SALIDAS / "OPORTUNIDADES.xlsx"
DB_MERCADO = DATOS / "precio_mercado.db"

CORREO_DESDE = os.environ.get("CORREO_DESDE", "").strip()
CORREO_PARA = os.environ.get("CORREO_PARA", "").strip()
APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "").strip()


def _clave_real(f):
    # Misma fórmula exacta que usa remates_v2.py / reporte_acumulado.py.
    rol = f.get("rol_completo")
    fecha_remate = f.get("fecha_remate")
    fecha = fecha_remate.date() if pd.notna(fecha_remate) else "sf"
    if pd.notna(rol) and rol:
        return f"{rol}|{fecha}"
    return f"{f.get('direccion_final')}|{fecha}"


def main():
    if len(sys.argv) != 3:
        print("Uso: python reporte_informe_fijo.py <archivo.docx> <fecha_tgr_desde AAAA-MM-DD>")
        sys.exit(1)
    nombre_docx = sys.argv[1]
    fecha_tgr_desde = sys.argv[2]
    ruta_docx = Path(nombre_docx)
    if not ruta_docx.is_absolute():
        ruta_docx = BASE / ruta_docx
    if not ruta_docx.exists():
        print(f"  [INFORME_FIJO] no encuentro {ruta_docx}")
        sys.exit(1)

    # ---- 1) TODOS los avisos de ESE informe, completos o no ----
    print(f"  [INFORME_FIJO] leyendo {ruta_docx.name} ...")
    tabla = remates_v2.tabla_avisos([ruta_docx], cache_ia_path=DB_MERCADO)
    print(f"  [INFORME_FIJO] {len(tabla)} avisos leídos de {ruta_docx.name}.")

    db_path_catastro = remates_v2.preparar_catastro()
    if db_path_catastro is None:
        print("  [INFORME_FIJO] no hay catastro disponible esta corrida - los avisos quedan sin avalúo SII ni "
              "verificación de dirección, pero se incluyen TODOS igual (completos o no).")

    # OJO: calcular() (dentro de procesar()) trae su propio filtro de "no repetir
    # lo que ya se vio sin cambios" (el mismo "[HISTORICO]" de la corrida normal),
    # que además MARCA esas propiedades como vistas en precio_mercado.db. Para
    # este reporte puntual eso es justo lo que no queremos: Rodrigo pidió ver
    # TODOS los avisos de este informe, completos o no, sin que ese filtro
    # esconda ninguno - y tampoco queremos que este reporte de solo lectura
    # modifique el historial real que usa la corrida normal de todos los días.
    # Se apaga ese filtro solo mientras dura esta llamada.
    estado_previo_precio_mercado = remates_v2.PRECIO_MERCADO_DISPONIBLE
    remates_v2.PRECIO_MERCADO_DISPONIBLE = False
    try:
        # aplicar_dedup=False (pedido de Rodrigo, 11-11-2026, tras encontrar que
        # 78 de 139 avisos del informe del lunes desaparecían sin aviso): este
        # reporte quiere ver TODOS los avisos del informe, uno por fila, SIN que
        # la deduplicación (pensada para la corrida normal, donde sí interesa
        # fusionar la misma propiedad vista por 2 fuentes) fusione ninguno.
        df_avisos, _hay_catastro = remates_v2.procesar([tabla], db_path_catastro, aplicar_dedup=False)
    finally:
        remates_v2.PRECIO_MERCADO_DISPONIBLE = estado_previo_precio_mercado
    n_leidos = len(tabla)
    n_en_informe = len(df_avisos)
    con_rol = int((df_avisos["roles_lista"].apply(lambda r: bool(r))).sum())
    print(f"  [INFORME_FIJO] {n_leidos} avisos leídos del Word | {n_en_informe} en el Excel "
          f"(deben coincidir - no se descarta nada en silencio) | {con_rol} con ROL de avalúo utilizable.")
    if n_leidos != n_en_informe:
        print(f"  [INFORME_FIJO] *** OJO: no coinciden ({n_leidos} leídos vs {n_en_informe} en el Excel) - "
              f"revisar, esto no debería pasar con aplicar_dedup=False. ***")

    ruta_excel = SALIDAS / "REPORTE_INFORME_FIJO.xlsx"
    with pd.ExcelWriter(ruta_excel, engine="openpyxl") as w:
        # ---- Resumen de control (pedido de Rodrigo, 11-11-2026): para que se
        # pueda verificar de un vistazo que no se perdió nada, sin tener que
        # contar filas a mano. ----
        con_direccion = int(df_avisos["direccion_final"].notna().sum()) if "direccion_final" in df_avisos.columns else 0
        resumen_control = pd.DataFrame({
            "Control": [
                "Avisos leídos del Word",
                "Avisos en este Excel (deben ser el mismo número de arriba)",
                "  - con ROL de avalúo utilizable (regex o IA)",
                "  - con dirección final resuelta",
                "  - sin ROL ni dirección (quedan igual, con el texto original a la vista en el aviso)",
            ],
            "Cantidad": [n_leidos, n_en_informe, con_rol, con_direccion, n_en_informe - max(con_rol, con_direccion)],
        })
        resumen_control.to_excel(w, sheet_name="RESUMEN_DE_CONTROL", index=False)
        remates_v2.dar_formato(w.sheets["RESUMEN_DE_CONTROL"])

        hoja = remates_v2.renombrar_para_mostrar(
            remates_v2.ordenar_columnas_alvaro(df_avisos, extra=["oportunidad_pct", "alertas", "motivo_no_elegible"])
        )
        nombre_hoja = ruta_docx.stem[:31] or "AVISOS"
        hoja.to_excel(w, sheet_name=nombre_hoja, index=False)
        remates_v2.dar_formato(w.sheets[nombre_hoja])

        # ---- 2) TGR genuinamente nuevo desde la fecha pedida ----
        n_tgr_nuevo = 0
        if EXCEL_HOY.exists() and DB_MERCADO.exists():
            try:
                cand = pd.read_excel(EXCEL_HOY, sheet_name="CANDIDATAS")
                cand_tgr = cand[cand["fuente"] == "TGR"].copy()
                cand_tgr["clave_real"] = cand_tgr.apply(_clave_real, axis=1)
                con = sqlite3.connect(DB_MERCADO)
                primeras_vez = dict(con.execute("SELECT clave, primera_vez FROM candidatas_historico").fetchall())
                con.close()
                cand_tgr["primera_vez"] = cand_tgr["clave_real"].map(primeras_vez)
                tgr_nuevo = cand_tgr[cand_tgr["primera_vez"] >= fecha_tgr_desde].drop(columns=["clave_real", "primera_vez"])
                n_tgr_nuevo = len(tgr_nuevo)
                tgr_nuevo.to_excel(w, sheet_name="TGR_NUEVO", index=False)
                remates_v2.dar_formato(w.sheets["TGR_NUEVO"])
            except Exception as e:
                print(f"  [INFORME_FIJO] no pude armar la hoja TGR_NUEVO ({e}); queda sin esa hoja.")
        else:
            print("  [INFORME_FIJO] no encuentro el Excel/base de datos de hoy; la hoja TGR_NUEVO no se puede armar.")
        print(f"  [INFORME_FIJO] {n_tgr_nuevo} propiedades de TGR nuevas desde el {fecha_tgr_desde}.")

    print(f"  [INFORME_FIJO] Excel: {ruta_excel.name}")

    if not APP_PASSWORD or not CORREO_DESDE or not CORREO_PARA:
        print("  [INFORME_FIJO] faltan variables de correo; el Excel queda guardado en 'salidas/' pero no se manda nada por correo.")
        return

    msg = EmailMessage()
    msg["Subject"] = f"Remates Chile - informe fijo ({ruta_docx.name}) + TGR nuevo desde {fecha_tgr_desde}"
    msg["From"] = CORREO_DESDE
    msg["To"] = CORREO_PARA
    msg.set_content(
        f"Reporte especial de una sola vez.\n\n"
        f"Hoja '{nombre_hoja[:31]}': los {len(df_avisos)} avisos de {ruta_docx.name}, completos o no "
        f"({con_rol} con ROL de avalúo utilizable).\n"
        f"Hoja 'TGR_NUEVO': {n_tgr_nuevo} propiedades de TGR nuevas desde el {fecha_tgr_desde}.\n"
    )
    msg.add_attachment(
        ruta_excel.read_bytes(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=ruta_excel.name,
    )
    try:
        contexto = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as server:
            server.login(CORREO_DESDE, APP_PASSWORD)
            server.send_message(msg)
        print(f"  [INFORME_FIJO] Enviado a {CORREO_PARA}.")
    except Exception as e:
        print(f"  [INFORME_FIJO] FALLÓ el envío: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
