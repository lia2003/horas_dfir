# Horas DFIR — Automatización semanal

CLI en Python para automatizar la gestión de horas del equipo DFIR/Forensics de EY.

## Requisitos

- Python 3.10+
- openpyxl

```bash
pip install -r requirements.txt
```

## Configuración

Edita `config.json` con la ruta del Excel y la carpeta de salida:

```json
{
  "excel_path": "C:\\Users\\liama\\Downloads\\Horas DFIR FY26_v2.xlsx",
  "output_dir": "C:\\Users\\liama\\Documents\\horas_dfir\\output"
}
```

Para cambiar de máquina (ej. PC de trabajo) solo edita `excel_path`.

## Uso

```bash
python main.py
```

Al iniciar, el CLI detecta las hojas semanales y te propone la semana activa por defecto (la más reciente cuyo lunes ya pasó). Puedes elegir otra si lo necesitas.

## Las 4 funciones

### Función 1 — Correos a gerentes
Genera un `.txt` (wording del correo) y un `.xlsx` (tabla formateada) por cada proyecto Cliente de la semana.  
Los archivos se guardan en `output/Semana DD-MM-YYYY/`.

- Excluye automáticamente "LE" y "LE Gestión".
- Si hay proyectos sin match en Jobs FY26, te pregunta si los agregas al catálogo.
- Avisa si un proyecto no tiene engagement registrado.

### Función 2 — Arrastre FDS (al final de semana, idealmente viernes)
Muestra todas las filas Cliente con Estado=Pendiente agrupadas por proyecto.  
Seleccionas cuáles arrastrar, y el código las copia a la hoja de la semana siguiente como `1. Lunes (FDS)` / Estado=Pendiente, respetando el orden de cada persona.

- Si la hoja siguiente no existe, la crea clonando la hoja "Plantilla".
- Siempre hace respaldo del Excel antes de escribir.

### Función 3 — Mensajes al equipo + Reporte RD
Antes de pedir las respuestas de gerentes, te lista los proyectos de la semana y te deja marcar
cuáles NO se van a cargar (algunas semanas hay proyectos sin horas para cargar).  
Luego ingresas las respuestas de los gerentes (engagement, job number, horas aprobadas por
persona) y genera un `.txt` por integrante con su tabla de horas aprobadas lista para enviar.

Con esos mismos datos (ya sin los proyectos excluidos) arma/actualiza `Reporte_RD.xlsx` en la
ruta de `reporte_rd_path`, agregando una hoja nueva por semana (`DD-MM-YYYY`) sin borrar las
anteriores. Ese archivo es el que se envía al jefe de RD — se genera a partir de las horas
aprobadas para el equipo, no de lo que quedó cargado en el Excel.

- El ajuste de horas en el Excel (si el gerente aprobó distinto a lo mapeado) lo haces tú manualmente antes de que el equipo cargue.
- La columna APROBADAS del Reporte_RD queda editable (dropdown ☑/☐) para marcarla vos misma antes de enviar.

### Función 4 — Resumen de horas cargadas (interno)
Genera la tabla de horas cargadas (Estado=Cargado) agrupada por persona y engagement, para pegar en Teams o para verificación propia una vez que el equipo ya cargó.  
Incluye filas `1. Lunes (FDS)` que ya estén en Cargado. Ya no es el reporte que se envía a RD (eso lo genera la Función 3).

## Estructura del proyecto

```
horas_dfir/
├── config.json         ← ruta del Excel y carpeta de salida
├── main.py             ← punto de entrada CLI
├── requirements.txt
├── .gitignore          ← excluye .xlsx, respaldos y output/
├── README.md
└── src/
    ├── excel_reader.py ← lectura del Excel (solo lectura)
    ├── utils.py        ← backup, confirmaciones, formateo
    ├── funcion1.py     ← correos a gerentes
    ├── funcion2.py     ← arrastre FDS
    ├── funcion3.py     ← mensajes al equipo
    └── funcion4.py     ← resumen para el jefe
```

## Notas importantes

- **Nunca se modifica el Excel sin hacer un respaldo previo con timestamp.**
- El código nunca toca la columna J ("Cargado en Job") ni la columna K ("Comentarios").
- La columna F "Tarea" nunca es modificada ni generada por el código; tú la revisas antes.
- Los respaldos quedan en la misma carpeta del Excel con formato `..._backup_YYYYMMDD_HHMMSS.xlsx`.
