"""
Función 4 — Resumen semanal para RD (Excel agrupado por Engagement ID).

Filtra filas con Estado='Cargado' y Tipo='Cliente' (incluye 1. Lunes (FDS)).
Agrupa por engagement ID → lista de personas con sus horas.
Si Andrea Neira tiene horas, pregunta su prorateo y lo aplica.
Guarda el resultado en Resumen_RD.xlsx en la carpeta de salida de la semana.
"""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .excel_reader import (
    PROYECTOS_EXCLUIDOS,
    get_semana_rows,
    load_jobs,
    parse_semana_fecha,
)


def generar_resumen(excel_path: Path, semana_nombre: str, output_dir: Path) -> None:
    parse_semana_fecha(semana_nombre)

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

    # ── 2. Detectar Andrea Neira y pedir prorateo ─────────────────────────────
    nombres_en_datos = {f["nombre"] for f in filas_cargadas if f["nombre"]}
    nombre_andrea = next(
        (n for n in nombres_en_datos
         if n and "andrea" in n.lower() and "neira" in n.lower()),
        None,
    )
    prorateo_andrea: float | None = None
    if nombre_andrea:
        print(f"\n  Se detectaron horas cargadas para {nombre_andrea}.")
        while True:
            raw = input(
                f"  Prorateo de {nombre_andrea} "
                "(ej: 0.48 — sus horas se multiplicarán): "
            ).strip()
            try:
                prorateo_andrea = float(raw.replace(",", "."))
                break
            except ValueError:
                print("  Ingresa un número válido (ej: 0.48).")

    # ── 3. Agrupar por engagement → persona ──────────────────────────────────
    # resumen[engagement][nombre] = horas_totales (originales)
    resumen: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    sin_engagement: list[dict] = []

    for f in filas_cargadas:
        nombre = f["nombre"] or ""
        proy   = (f["proyecto"] or "").strip()
        info   = jobs.get(proy, {})
        eng    = (info.get("engagement") or "").strip()

        if not eng:
            sin_engagement.append(f)
            eng = f"SIN ENGAGEMENT ({proy})"

        resumen[eng][nombre] += f["horas"]

    # ── 4. Avisar proyectos sin engagement ────────────────────────────────────
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

    # ── 5. Generar Excel ──────────────────────────────────────────────────────
    carpeta = output_dir / semana_nombre
    carpeta.mkdir(parents=True, exist_ok=True)
    archivo = carpeta / "Resumen_RD.xlsx"

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "Resumen"

    # Estilos
    hdr_font   = Font(bold=True, color="FFFFFF")
    hdr_fill   = PatternFill("solid", fgColor="1F4E79")
    eng_font   = Font(bold=True)
    eng_fill   = PatternFill("solid", fgColor="D6E4F0")
    tot_font   = Font(bold=True)
    tot_fill   = PatternFill("solid", fgColor="F2F2F2")
    center_aln = Alignment(horizontal="center", vertical="center")
    left_aln   = Alignment(horizontal="left",   vertical="center")
    right_aln  = Alignment(horizontal="right",  vertical="center")

    # Encabezado
    ws_out.append(["Engagement ID", "Nombre", "Horas"])
    for col_idx in range(1, 4):
        cell = ws_out.cell(row=1, column=col_idx)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.alignment = center_aln

    ws_out.column_dimensions["A"].width = 28
    ws_out.column_dimensions["B"].width = 30
    ws_out.column_dimensions["C"].width = 12

    total_global = 0.0
    fila_actual  = 2

    for eng in sorted(resumen.keys()):
        personas = resumen[eng]

        # Subtotal del engagement (con prorateo aplicado)
        subtotal = sum(
            round(h * prorateo_andrea, 1) if (nombre_andrea and n == nombre_andrea and prorateo_andrea is not None) else h
            for n, h in personas.items()
        )
        total_global += subtotal

        # Fila de cabecera de engagement
        ws_out.cell(row=fila_actual, column=1, value=eng).font      = eng_font
        ws_out.cell(row=fila_actual, column=1).fill      = eng_fill
        ws_out.cell(row=fila_actual, column=1).alignment = left_aln
        ws_out.cell(row=fila_actual, column=2, value="").fill       = eng_fill
        ws_out.cell(row=fila_actual, column=3, value=subtotal)
        ws_out.cell(row=fila_actual, column=3).font      = eng_font
        ws_out.cell(row=fila_actual, column=3).fill      = eng_fill
        ws_out.cell(row=fila_actual, column=3).alignment = right_aln
        ws_out.cell(row=fila_actual, column=3).number_format = "0.0"
        fila_actual += 1

        # Filas de persona
        for nombre in sorted(personas.keys()):
            h_orig = personas[nombre]
            if nombre_andrea and nombre == nombre_andrea and prorateo_andrea is not None:
                h_mostrar = round(h_orig * prorateo_andrea, 1)
                label = f"{nombre} (×{prorateo_andrea})"
            else:
                h_mostrar = h_orig
                label = nombre

            ws_out.cell(row=fila_actual, column=1, value="").alignment = left_aln
            ws_out.cell(row=fila_actual, column=2, value=label).alignment = left_aln
            ws_out.cell(row=fila_actual, column=3, value=h_mostrar)
            ws_out.cell(row=fila_actual, column=3).alignment    = right_aln
            ws_out.cell(row=fila_actual, column=3).number_format = "0.0"
            fila_actual += 1

    # Fila de total global
    ws_out.cell(row=fila_actual, column=1, value="TOTAL").font      = tot_font
    ws_out.cell(row=fila_actual, column=1).fill      = tot_fill
    ws_out.cell(row=fila_actual, column=1).alignment = left_aln
    ws_out.cell(row=fila_actual, column=2, value="").fill            = tot_fill
    ws_out.cell(row=fila_actual, column=3, value=round(total_global, 1))
    ws_out.cell(row=fila_actual, column=3).font      = tot_font
    ws_out.cell(row=fila_actual, column=3).fill      = tot_fill
    ws_out.cell(row=fila_actual, column=3).alignment = right_aln
    ws_out.cell(row=fila_actual, column=3).number_format = "0.0"

    wb_out.save(archivo)

    # ── 6. Mostrar resumen en pantalla ────────────────────────────────────────
    print(f"\n  === RESUMEN SEMANAL - {semana_nombre} ===\n")
    for eng in sorted(resumen.keys()):
        print(f"  {eng}")
        for nombre in sorted(resumen[eng].keys()):
            h_orig = resumen[eng][nombre]
            if nombre_andrea and nombre == nombre_andrea and prorateo_andrea is not None:
                h = round(h_orig * prorateo_andrea, 1)
                print(f"    {nombre:<30} {h:>6.1f}h  (×{prorateo_andrea})")
            else:
                print(f"    {nombre:<30} {h_orig:>6.1f}h")
        print()

    print(f"  TOTAL GLOBAL: {round(total_global, 1)}h")
    print(f"\n  Guardado en: {archivo}")
