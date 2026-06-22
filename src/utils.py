"""Utilidades compartidas: guardado seguro, confirmaciones, formateo."""

import os
import shutil
import tempfile
import time
from datetime import date
from pathlib import Path


def guardar_excel(wb, excel_path: Path, post_process=None) -> None:
    """
    Guarda el workbook de forma segura para entornos OneDrive compartidos:

    - Escribe en un temporal en %TEMP% (fuera de OneDrive, invisible al equipo).
    - Copia el temporal sobre el original; en OneDrive for Business esto
      funciona aunque compañeros tengan el archivo abierto en co-autoría,
      ya que el archivo no queda con bloqueo exclusivo en ese modo.
    - Si el archivo sí estuviera bloqueado (p.ej. tú misma lo tienes abierto
      localmente en Excel), reintenta cada 3s hasta 60s mostrando un aviso.
    """
    # Temporal en carpeta del sistema, nunca en la carpeta de OneDrive
    fd, tmp_str = tempfile.mkstemp(suffix=excel_path.suffix)
    tmp = Path(tmp_str)
    os.close(fd)

    try:
        wb.save(tmp)
    except Exception as e:
        tmp.unlink(missing_ok=True)
        raise RuntimeError(f"Error preparando el guardado: {e}") from e

    if post_process is not None:
        post_process(tmp)

    max_intentos = 20
    for intento in range(1, max_intentos + 1):
        try:
            shutil.copy2(str(tmp), str(excel_path))
            tmp.unlink(missing_ok=True)
            return
        except PermissionError:
            if intento == 1:
                print(
                    "\n  AVISO: El Excel está abierto en tu computadora."
                    "\n  Ciérralo y el programa continuará automáticamente."
                )
            print(f"  Reintentando... ({intento * 3}s)")
            time.sleep(3)

    tmp.unlink(missing_ok=True)
    raise PermissionError(
        f"No se pudo guardar '{excel_path.name}' tras {max_intentos * 3}s. "
        "Cierra el archivo en Excel e intenta de nuevo."
    )


def confirmar(mensaje: str) -> bool:
    """Pide s/n al usuario. Retorna True si confirma."""
    while True:
        resp = input(f"{mensaje} [s/n]: ").strip().lower()
        if resp in ("s", "si", "sí", "y", "yes"):
            return True
        if resp in ("n", "no"):
            return False
        print("  Responde 's' o 'n'.")


def fmt_fecha(d: date) -> str:
    """date → DD/MM/YYYY"""
    return d.strftime("%d/%m/%Y")


def fmt_fecha_corta(d: date) -> str:
    """date → DD/MM/YY"""
    return d.strftime("%d/%m/%y")


def sanitizar_nombre_archivo(nombre: str) -> str:
    """Elimina caracteres inválidos para nombres de archivo en Windows."""
    import re
    return re.sub(r'[<>:"/\\|?*]', '_', nombre).strip()
