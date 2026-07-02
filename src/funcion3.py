"""
Función 3 — Mensajes al equipo tras aprobación de gerentes.

Flujo:
  1. Muestra las horas solicitadas por proyecto / persona (de la hoja semanal).
  2. El usuario ingresa, por proyecto: engagement, job number y horas aprobadas
     por persona (puede diferir de lo solicitado).
  3. Opcionalmente el usuario indica consideraciones extra por proyecto.
  4. Genera un .txt por integrante con la tabla de horas aprobadas lista para
     enviar (el usuario lo copia/pega o lo envía manualmente).

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
    parse_semana_fecha,
)
from .utils import fmt_fecha, sanitizar_nombre_archivo


def procesar_aprobacion(excel_path: Path, semana_nombre: str, output_dir: Path) -> None:
    lunes: date = parse_semana_fecha(semana_nombre)

    filas = get_semana_rows(excel_path, semana_nombre)
    jobs  = load_jobs(excel_path)

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

    # ── 2. Ingresar respuestas de gerentes por proyecto ───────────────────────
    print("\n  === INGRESO DE RESPUESTAS DE GERENTES ===")
    print("  Para cada proyecto ingresa el engagement, job y horas aprobadas.\n")

    datos_aprobados: dict[str, dict] = {}

    for proy, fps in sorted(por_proyecto.items()):
        info_job = jobs.get(proy, {})
        eng_default = info_job.get("engagement") or ""

        print(f"  {'-' * 55}")
        print(f"  Proyecto : {proy}")
        print(f"  Gerente  : {info_job.get('gerente', 'N/A')}")
        print()

        # Horas solicitadas por persona
        horas_solicitadas: dict[str, float] = defaultdict(float)
        tipo_act_por_persona: dict[str, str] = {}
        for f in fps:
            nombre = f["nombre"] or ""
            horas_solicitadas[nombre] += f["horas"]
            if nombre not in tipo_act_por_persona and f["tipo_actividad"]:
                tipo_act_por_persona[nombre] = f["tipo_actividad"]

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
        }
        print()

    # ── 3. Construir tabla por integrante ─────────────────────────────────────
    # { nombre: [ {proyecto, engagement, job, comentario, horas, extras} ] }
    por_integrante: dict[str, list[dict]] = defaultdict(list)

    for proy, datos in datos_aprobados.items():
        for nombre, horas in datos["horas_aprobadas"].items():
            por_integrante[nombre].append({
                "proyecto":    proy,
                "engagement":  datos["engagement"],
                "job":         datos["job"],
                "comentario":  datos["tipo_act_persona"].get(nombre, ""),
                "horas":       horas,
                "extras":      datos["extras"],
            })

    # ── 3b. Prorateo de Andrea Neira ─────────────────────────────────────────
    nombre_andrea = next(
        (n for n in por_integrante
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
        (n for n in por_integrante
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
        f"Hola {nombre_corto},",
        "",
        "Te compartimos las horas aprobadas para cargar esta semana:",
        "",
        tabla,
    ]
    if notas:
        partes += ["", "Consideraciones:"] + notas

    partes += [
        "",
        "Por favor carga exactamente las horas indicadas.",
        "",
        "Saludos,",
    ]

    contenido = "\n".join(partes)
    nombre_file = sanitizar_nombre_archivo(nombre)
    archivo = carpeta / f"Mensaje_{nombre_file}.txt"
    archivo.write_text(contenido, encoding="utf-8")
