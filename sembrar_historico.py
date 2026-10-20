# -*- coding: utf-8 -*-
"""
Script de UNA SOLA VEZ (2026-10-20) para "sembrar" la tabla candidatas_historico
de precio_mercado.db con la línea base acordada entre Rodrigo y Álvaro: todo lo
que ya estaba como candidata elegible en la corrida del 2026-10-13 22:12 (antes
del viernes 16-10) se marca como "ya visto" desde esa fecha. Así, la primera
corrida real con la lógica nueva de "nuevo" (universo completo de candidatas,
no solo el Top 50) reconoce como YA VISTAS las ~1195 candidatas de esa línea
base, y solo muestra como nuevas las que de verdad aparecieron después - que
es justo el grupo de ~30 (19 TGR + 8 AVISO + 3 MACAL) que Rodrigo y Álvaro
revisaron y aprobaron antes de correr esto.

Se activa SOLO si el workflow se dispara con el input "sembrar_baseline" en
true (ver .github/workflows/correr-remates.yml) - en corridas normales este
paso no se ejecuta, así que no hay riesgo de "resetear" el histórico por
accidente en una corrida semanal futura.

Uso: python sembrar_historico.py datos/precio_mercado.db datos/seed_candidatas_2026-10-13.csv
"""
import sqlite3
import sys
import csv


def main():
    if len(sys.argv) != 3:
        print("Uso: python sembrar_historico.py <ruta_db> <ruta_csv>")
        sys.exit(1)

    ruta_db, ruta_csv = sys.argv[1], sys.argv[2]
    conexion = sqlite3.connect(ruta_db)
    cursor = conexion.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS candidatas_historico (
            clave TEXT PRIMARY KEY,
            oportunidad_pct REAL,
            primera_vez TEXT,
            ultima_vez TEXT
        )
    """)
    conexion.commit()

    baseline_fecha = "2026-10-13"
    filas = []
    with open(ruta_csv, newline="", encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            clave = fila["clave"]
            pct_txt = fila["oportunidad_pct"]
            pct = float(pct_txt) if pct_txt not in ("", "nan", None) else None
            filas.append((clave, pct, baseline_fecha, baseline_fecha))

    cursor.executemany("""
        INSERT INTO candidatas_historico (clave, oportunidad_pct, primera_vez, ultima_vez)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(clave) DO UPDATE SET
            oportunidad_pct=excluded.oportunidad_pct,
            ultima_vez=excluded.ultima_vez
    """, filas)
    conexion.commit()
    n = cursor.execute("SELECT COUNT(*) FROM candidatas_historico").fetchone()[0]
    print(f"  [SEMBRAR] {len(filas)} claves sembradas desde {ruta_csv}. "
          f"candidatas_historico queda con {n} filas en total.")
    conexion.close()


if __name__ == "__main__":
    main()
