"""
Horas DFIR — CLI de automatización semanal
==========================================
Uso:
    python main.py

Requiere config.json en el mismo directorio con:
    {
        "excel_path": "C:\\ruta\\al\\Horas DFIR FY26_v2.xlsx",
        "output_dir": "C:\\ruta\\carpeta\\salida"
    }
"""

from __future__ import annotations

import json
import sys
from datetime import date, timedelta
from pathlib import Path

import openpyxl  # noqa: F401 — verificamos la dependencia al inicio

from src.excel_reader import detect_semanas, parse_semana_fecha
from src.funcion1 import generar_correos
from src.funcion2 import arrastrar_fds
from src.funcion3 import procesar_aprobacion
from src.funcion4 import generar_resumen

# ── Carga de configuración ────────────────────────────────────────────────────

def cargar_config() -> tuple[Path, Path]:
    config_path = Path(__file__).parent / "config.json"
    if not config_path.exists():
        print(f"ERROR: No se encontró config.json en {config_path.parent}")
        sys.exit(1)
    with open(config_path, encoding="utf-8") as fh:
        cfg = json.load(fh)
    excel_path = Path(cfg["excel_path"])
    output_dir = Path(cfg["output_dir"])
    if not excel_path.exists():
        print(f"ERROR: El archivo Excel no existe:\n  {excel_path}")
        print("Revisa la ruta 'excel_path' en config.json.")
        sys.exit(1)
    output_dir.mkdir(parents=True, exist_ok=True)
    return excel_path, output_dir


# ── Selección de semana ───────────────────────────────────────────────────────

def seleccionar_semana(excel_path: Path) -> str:
    semanas = detect_semanas(excel_path)
    if not semanas:
        print("ERROR: No se encontraron hojas 'Semana DD-MM-YYYY' en el Excel.")
        sys.exit(1)

    # Semana por defecto: la más reciente cuyo lunes ya pasó o es hoy
    hoy   = date.today()
    lunes_semana_actual = hoy - timedelta(days=hoy.weekday())  # lunes de esta semana

    idx_defecto = 0
    for i, (fecha, _) in enumerate(semanas):
        if fecha <= lunes_semana_actual:
            idx_defecto = i

    print("\n" + "=" * 55)
    print("  HOJAS SEMANALES DISPONIBLES")
    print("=" * 55)
    for i, (fecha, nombre) in enumerate(semanas):
        marca = "  <- por defecto" if i == idx_defecto else ""
        print(f"  {i+1:>2}. {nombre}{marca}")
    print("=" * 55)

    while True:
        raw = input(
            f"\n  Elige semana [Enter = {semanas[idx_defecto][1]}]: "
        ).strip()
        if raw == "":
            return semanas[idx_defecto][1]
        try:
            idx = int(raw) - 1
            if 0 <= idx < len(semanas):
                return semanas[idx][1]
            print(f"  Número fuera de rango (1–{len(semanas)}).")
        except ValueError:
            print("  Ingresa un número o presiona Enter.")


# ── Menú principal ────────────────────────────────────────────────────────────

MENU = """
+--------------------------------------------------+
|         HORAS DFIR -- MENU PRINCIPAL             |
+--------------------------------------------------+
|  1. Generar correos a gerentes                   |
|  2. Arrastrar horas pendientes (FDS)             |
|  3. Mensajes al equipo (tras aprobacion)         |
|  4. Resumen para RD                              |
|  0. Salir                                        |
+--------------------------------------------------+"""


def main() -> None:
    print("\n  Cargando configuración...")
    excel_path, output_dir = cargar_config()
    print(f"  Excel : {excel_path}")
    print(f"  Salida: {output_dir}")

    semana_sel = seleccionar_semana(excel_path)
    lunes_sel  = parse_semana_fecha(semana_sel)
    viernes_sel = lunes_sel + timedelta(days=4)

    print(f"\n  Semana activa: {semana_sel}  "
          f"({lunes_sel.strftime('%d/%m/%Y')} - {viernes_sel.strftime('%d/%m/%Y')})")

    while True:
        print(MENU)
        opcion = input("  Opción: ").strip()

        if opcion == "1":
            print("\n── Función 1: Correos a gerentes ──────────────────────")
            generar_correos(excel_path, semana_sel, output_dir)

        elif opcion == "2":
            print("\n── Función 2: Arrastre FDS ────────────────────────────")
            arrastrar_fds(excel_path, semana_sel)

        elif opcion == "3":
            print("\n── Función 3: Mensajes al equipo ──────────────────────")
            procesar_aprobacion(excel_path, semana_sel, output_dir)

        elif opcion == "4":
            print("\n── Función 4: Resumen para RD ────────────────────────")
            generar_resumen(excel_path, semana_sel, output_dir)

        elif opcion == "0":
            print("\n  Hasta luego.\n")
            break

        else:
            print("  Opción inválida. Elige entre 0 y 4.")

        # Pausa antes de volver al menú
        input("\n  [Enter para volver al menú]")


if __name__ == "__main__":
    main()
