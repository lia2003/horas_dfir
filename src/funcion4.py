"""
Función 4 — Resumen semanal para el jefe (formato Teams).

Filtra filas con Estado='Cargado' y Tipo='Cliente' (incluye 1. Lunes (FDS)).
Agrupa por persona → por engagement (código de Jobs FY26).
El nombre de la persona aparece solo en su primera fila.
Avisa si algún proyecto no tiene engagement.
Guarda el resultado en Resumen_Jefe.txt en la carpeta de salida de la semana.
"""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from .excel_reader import (
    PROYECTOS_EXCLUIDOS,
    get_semana_rows,
    load_jobs,
    parse_semana_fecha,
)


def generar_resumen(excel_path: Path, semana_nombre: str, output_dir: Path) -> None:
    parse_semana_fecha(semana_nombre)  # valida el nombre

    filas = get_semana_rows(excel_path, semana_nombre)
    jobs  = load_jobs(excel_path)

    # ── 1. Filtrar: Cargado + Cliente + no excluidos ──────────────────────────
    filas_cargadas = [
        f for f in filas
        if f["estado"] == "Cargado"
        and f["tipo"] == "Cliente"
        and (f["proyecto"] or "").strip() not in PROYECTOS_EXCLUIDOS
        and f["proyecto"]
        and (f["horas"] or 0) > 0
    ]

    if not filas_cargadas:
        print("  No hay filas en Estado='Cargado' para generar el resumen.")
        return

    # ── 2. Agrupar por persona → engagement ──────────────────────────────────
    # resumen[nombre][engagement] = horas_totales
    resumen: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    sin_engagement: list[dict] = []

    for f in filas_cargadas:
        nombre  = f["nombre"] or ""
        proy    = (f["proyecto"] or "").strip()
        info    = jobs.get(proy, {})
        eng     = (info.get("engagement") or "").strip()

        if not eng:
            sin_engagement.append(f)
            eng = f"SIN ENGAGEMENT ({proy})"

        resumen[nombre][eng] += f["horas"]

    # ── 3. Avisar proyectos sin engagement ────────────────────────────────────
    if sin_engagement:
        print("\n  ADVERTENCIA — Las siguientes filas cargadas no tienen engagement "
              "en Jobs FY26:")
        proyectos_vistos: set[str] = set()
        for f in sin_engagement:
            proy = (f["proyecto"] or "").strip()
            if proy not in proyectos_vistos:
                print(f"    - {proy} (gerente: {jobs.get(proy, {}).get('gerente', 'N/A')})")
                proyectos_vistos.add(proy)
        print("  Actualiza el engagement en Jobs FY26 y vuelve a generar el resumen.\n")

    # ── 4. Construir tabla ────────────────────────────────────────────────────
    ANC_NOMBRE = 28
    ANC_ENG    = 22
    ANC_HORAS  = 10

    encabezado = (
        f"{'Nombre':<{ANC_NOMBRE}}  "
        f"{'Proyecto (Engagement)':<{ANC_ENG}}  "
        f"{'Horas':>{ANC_HORAS}}"
    )
    separador = "-" * len(encabezado)

    lineas: list[str] = [encabezado, separador]

    total_global = 0.0
    for nombre in sorted(resumen.keys()):
        engagements = resumen[nombre]
        primera = True
        for eng in sorted(engagements.keys()):
            horas = engagements[eng]
            total_global += horas
            nombre_display = nombre if primera else ""
            lineas.append(
                f"{nombre_display:<{ANC_NOMBRE}}  "
                f"{eng:<{ANC_ENG}}  "
                f"{horas:>{ANC_HORAS}.1f}"
            )
            primera = False

    lineas.append(separador)
    lineas.append(
        f"{'TOTAL':<{ANC_NOMBRE}}  "
        f"{'':<{ANC_ENG}}  "
        f"{total_global:>{ANC_HORAS}.1f}"
    )

    tabla = "\n".join(lineas)

    # ── 5. Mostrar y guardar ──────────────────────────────────────────────────
    print(f"\n  === RESUMEN SEMANAL - {semana_nombre} ===\n")
    print(tabla)
    print()

    carpeta = output_dir / semana_nombre
    carpeta.mkdir(parents=True, exist_ok=True)
    archivo = carpeta / "Resumen_Jefe.txt"
    archivo.write_text(
        f"RESUMEN SEMANAL - {semana_nombre}\n\n{tabla}\n",
        encoding="utf-8",
    )
    print(f"  Guardado en: {archivo}")
