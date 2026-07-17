"""
Función 3 — Mensajes al equipo tras aprobación de gerentes + Reporte_RD.

Flujo:
  1. Muestra las horas solicitadas por proyecto / persona (de la hoja semanal).
  2. Permite marcar proyectos que esta semana NO se van a cargar (sin mensaje
     al equipo). Esos proyectos igual se incluyen en el Reporte_RD, con los
     datos tal cual están en el Excel y la columna APROBADAS sin marcar,
     porque no pasan por la aprobación manual de gerente.
  3. Para el resto de los proyectos, el usuario ingresa: engagement, job
     number y horas aprobadas por persona (puede diferir de lo solicitado).
  4. Opcionalmente el usuario indica consideraciones extra por proyecto.
  5. Genera un .txt por integrante con la tabla de horas aprobadas lista para
     enviar (el usuario lo copia/pega o lo envía manualmente).
  6. Genera/actualiza Reporte_RD.xlsx con TODOS los proyectos de la semana
     (excluidos o no), listo para enviar al jefe de RD.

No se modifica el Excel en esta función.
"""

from __future__ import annotations

from collections import defaultdict
from datetime import date
from pathlib import Path

from .excel_reader import (
    PROYECTOS_EXCLUIDOS,
    get_semana_rows,
    load_jobs,
    load_rates,
    parse_semana_fecha,
)
from .reporte_rd import generar_reporte_rd
from .utils import fmt_fecha, sanitizar_nombre_archivo


def procesar_aprobacion(
    excel_path: Path,
    semana_nombre: str,
    output_dir: Path,
    reporte_rd_path: Path | None = None,
) -> None:
    lunes: date = parse_semana_fecha(semana_nombre)

    filas = get_semana_rows(excel_path, semana_nombre)
    jobs  = load_jobs(excel_path)
    rates = load_rates(excel_path)

    # ── 1. Filas Cliente válidas ──────────────────────────────────────────────
    filas_cliente = [
        f for f in filas
        if f["tipo"] == "Cliente"
        and (f["proyecto"] or "").strip() not in PROYECTOS_EXCLUIDOS
        and f["proyecto"]
    ]

    if not filas_cliente:
        print("  No hay filas Cliente en esta semana.")
        return

    # Agrupar por proyecto
    por_proyecto: dict[str, list[dict]] = defaultdict(list)
    for f in filas_cliente:
        por_proyecto[(f["proyecto"] or "").strip()].append(f)

    # Rank por persona (para el Reporte_RD)
    rank_por_nombre: dict[str, str] = {}
    for f in filas_cliente:
        n = f["nombre"] or ""
        if n and n not in rank_por_nombre and f["rank"]:
            rank_por_nombre[n] = f["rank"]

    # ── 1b. Marcar proyectos que esta semana NO se van a cargar ──────────────
    proyectos_lista = sorted(por_proyecto.keys())
    print("\n  === PROYECTOS DE LA SEMANA ===")
    for i, p in enumerate(proyectos_lista, start=1):
        print(f"    {i}. {p}")
    excl_input = input(
        "\n  Proyectos a EXCLUIR del mensaje al equipo esta semana "
        "(numeros separados por coma, Enter = ninguno): "
    ).strip()

    proyectos_excluidos_semana: set[str] = set()
    if excl_input:
        for tok in excl_input.split(","):
            tok = tok.strip()
            if not tok:
                continue
            try:
                idx = int(tok) - 1
                if 0 <= idx < len(proyectos_lista):
                    proyectos_excluidos_semana.add(proyectos_lista[idx])
                else:
                    print(f"    Aviso: número {tok} fuera de rango, ignorado.")
            except ValueError:
                print(f"    Aviso: '{tok}' no es un número válido, ignorado.")
        if proyectos_excluidos_semana:
            print(
                "\n  Proyectos excluidos del mensaje al equipo esta semana "
                "(igual se incluyen en Reporte_RD, sin aprobación manual):"
            )
            for p in sorted(proyectos_excluidos_semana):
                print(f"    - {p}")

    # ── 2. Ingresar respuestas de gerentes por proyecto ───────────────────────
    print("\n  === INGRESO DE RESPUESTAS DE GERENTES ===")
    print("  Para cada proyecto no excluido, ingresa el engagement, job y horas aprobadas.\n")

    datos_aprobados: dict[str, dict] = {}

    for proy, fps in sorted(por_proyecto.items()):
        info_job = jobs.get(proy, {})
        eng_default = info_job.get("engagement") or ""

        # Horas solicitadas por persona (se calcula siempre, se use o no)
        horas_solicitadas: dict[str, float] = defaultdict(float)
        tipo_act_por_persona: dict[str, str] = {}
        for f in fps:
            nombre = f["nombre"] or ""
            horas_solicitadas[nombre] += f["horas"]
            if nombre not in tipo_act_por_persona and f["tipo_actividad"]:
                tipo_act_por_persona[nombre] = f["tipo_actividad"]

        if proy in proyectos_excluidos_semana:
            # No se manda mensaje al equipo por este proyecto; se incluye
            # igual en el Reporte_RD con los datos tal cual, sin pedir
            # aprobación manual de gerente.
            print(f"  {'-' * 55}")
            print(f"  Proyecto : {proy}  (excluido del mensaje -> va a Reporte_RD sin aprobar)")
            datos_aprobados[proy] = {
                "engagement":        eng_default or "SIN ENGAGEMENT",
                "job":               "0000",
                "horas_aprobadas":   dict(horas_solicitadas),
                "tipo_act_persona":  tipo_act_por_persona,
                "extras":            None,
                "aprobada":          False,
            }
            continue

        print(f"  {'-' * 55}")
        print(f"  Proyecto : {proy}")
        print(f"  Gerente  : {info_job.get('gerente', 'N/A')}")
        print()

        print("  Solicitado:")
        for nombre, h in sorted(horas_solicitadas.items()):
            print(f"    {nombre:<25}  {h}h")
        print()

        # Engagement
        eng_input = input(
            f"  Engagement [{eng_default or 'no registrado'}] "
            "(Enter para usar el de Jobs FY26): "
        ).strip()
        engagement = eng_input if eng_input else eng_default or "SIN ENGAGEMENT"

        # Job number
        job_input = input("  Job number (Enter = '0000'): ").strip()
        job_num = job_input if job_input else "0000"

        # Horas aprobadas por persona
        print()
        horas_aprobadas: dict[str, float] = {}
        for nombre in sorted(horas_solicitadas.keys()):
            h_sol = horas_solicitadas[nombre]
            while True:
                raw = input(
                    f"  Horas aprobadas — {nombre:<25} "
                    f"(solicitadas: {h_sol}h, Enter = mismas): "
                ).strip()
                if raw == "":
                    horas_aprobadas[nombre] = h_sol
                    break
                try:
                    horas_aprobadas[nombre] = float(raw.replace(",", "."))
                    break
                except ValueError:
                    print("    Ingresa un número válido (ej: 8 o 7.5).")

        # Consideraciones extra
        print()
        extra = input(
            f"  Consideraciones extra para {proy} "
            "(ej: '4h a E-XXXX y 4h a E-YYYY', Enter para ninguna): "
        ).strip()

        datos_aprobados[proy] = {
            "engagement":        engagement,
            "job":               job_num,
            "horas_aprobadas":   horas_aprobadas,
            "tipo_act_persona":  tipo_act_por_persona,
            "extras":            extra or None,
            "aprobada":          True,
        }
        print()

    # ── 3. Construir tabla por integrante (solo proyectos NO excluidos) ───────
    # { nombre: [ {proyecto, engagement, job, comentario, horas, extras} ] }
    por_integrante: dict[str, list[dict]] = defaultdict(list)

    for proy, datos in datos_aprobados.items():
        if proy in proyectos_excluidos_semana:
            continue
        for nombre, horas in datos["horas_aprobadas"].items():
            por_integrante[nombre].append({
                "proyecto":    proy,
                "engagement":  datos["engagement"],
                "job":         datos["job"],
                "comentario":  datos["tipo_act_persona"].get(nombre, ""),
                "horas":       horas,
                "extras":      datos["extras"],
            })

    # ── 3b. Prorateo de Andrea Neira (sobre todos los proyectos, incluidos
    #        los excluidos del mensaje, ya que igual entran al Reporte_RD) ────
    nombre_andrea = next(
        (n for n in rank_por_nombre
         if n and "andrea" in n.lower() and "neira" in n.lower()),
        None,
    )
    prorateo_andrea: float | None = None
    if nombre_andrea:
        print(f"\n  Se detectaron horas para {nombre_andrea}.")
        while True:
            raw = input(
                f"  Ingresa el prorateo de {nombre_andrea} "
                "(ej: 0.48 — las horas se multiplicarán por este factor): "
            ).strip()
            try:
                prorateo_andrea = float(raw.replace(",", "."))
                break
            except ValueError:
                print("  Ingresa un número válido (ej: 0.48).")

    nombre_daniel = next(
        (n for n in rank_por_nombre
         if n and "daniel" in n.lower() and "cabrera" in n.lower()),
        None,
    )
    prorateo_daniel: float | None = None
    if nombre_daniel:
        print(f"\n  Se detectaron horas para {nombre_daniel}.")
        while True:
            raw = input(
                f"  Ingresa el prorateo de {nombre_daniel} "
                "(ej: 0.48 — las horas se multiplicarán por este factor): "
            ).strip()
            try:
                prorateo_daniel = float(raw.replace(",", "."))
                break
            except ValueError:
                print("  Ingresa un número válido (ej: 0.48).")

    # ── 4. Generar .txt por integrante ────────────────────────────────────────
    carpeta = output_dir / semana_nombre
    carpeta.mkdir(parents=True, exist_ok=True)

    print("  === MENSAJES GENERADOS ===")
    if not por_integrante:
        print("    (Todos los proyectos quedaron excluidos del mensaje al equipo esta semana.)")
    for nombre, entradas in sorted(por_integrante.items()):
        if nombre == nombre_andrea:
            prorateo = prorateo_andrea
        elif nombre == nombre_daniel:
            prorateo = prorateo_daniel
        else:
            prorateo = None
        _generar_txt_integrante(nombre, entradas, lunes, carpeta, prorateo)
        sufijo = f" (prorateo ×{prorateo})" if prorateo is not None else ""
        print(f"    OK: {nombre}{sufijo}")

    print(f"\n  Carpeta de salida: {carpeta}")

    # ── 5. Reporte_RD.xlsx ────────────────────────────────────────────────────
    if reporte_rd_path is not None:
        prorateos: dict[str, float] = {}
        if nombre_andrea and prorateo_andrea is not None:
            prorateos[nombre_andrea] = prorateo_andrea
        if nombre_daniel and prorateo_daniel is not None:
            prorateos[nombre_daniel] = prorateo_daniel

        generar_reporte_rd(
            reporte_rd_path,
            lunes,
            datos_aprobados,
            rank_por_nombre,
            jobs,
            rates,
            prorateos,
            nombre_andrea,
            nombre_daniel,
        )
        print(
            "  Reporte_RD.xlsx queda listo para enviar a tu jefe de RD "
            "(revisa la columna APROBADAS antes de enviarlo)."
        )
    else:
        print("\n  Aviso: no hay 'reporte_rd_path' configurado; no se generó Reporte_RD.xlsx.")


# ── helper ────────────────────────────────────────────────────────────────────

def _generar_txt_integrante(
    nombre: str,
    entradas: list[dict],
    lunes: date,
    carpeta: Path,
    prorateo: float | None = None,
) -> None:
    nombre_corto = nombre.split()[0] if nombre else nombre

    # Anchos de columna
    ANC = [25, 22, 20, 10, 14, 17]
    cols = ["Nombre", "Proyecto", "Engagement", "Job", "Comentario", "Horas aprobadas"]
    sep  = " | "

    encabezado = sep.join(f"{c:<{a}}" for c, a in zip(cols, ANC))
    separador  = "-" * len(encabezado)

    filas_txt: list[str] = [encabezado, separador]
    notas: list[str] = []

    for e in entradas:
        h = e["horas"]
        if prorateo is not None:
            h = round(h * prorateo, 1)
        fila = sep.join([
            f"{nombre:<{ANC[0]}}",
            f"{e['proyecto']:<{ANC[1]}}",
            f"{e['engagement']:<{ANC[2]}}",
            f"{e['job']:<{ANC[3]}}",
            f"{e['comentario']:<{ANC[4]}}",
            f"{h:<{ANC[5]}}",
        ])
        filas_txt.append(fila)
        if e.get("extras"):
            notas.append(f"  * {e['proyecto']}: {e['extras']}")

    tabla = "\n".join(filas_txt)

    partes: list[str] = [
        f"Hola {nombre_corto}, carga tus horas AHORA:",
        "",
        tabla,
    ]
    if notas:
        partes += ["", "Consideraciones:"] + notas

    partes += [
        "",
        "No te olvides de ACTUALIZAR EL EXCEL COMO CARGADO.",
    ]

    contenido = "\n".join(partes)
    nombre_file = sanitizar_nombre_archivo(nombre)
    archivo = carpeta / f"Mensaje_{nombre_file}.txt"
    archivo.write_text(contenido, encoding="utf-8")
