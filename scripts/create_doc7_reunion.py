#!/usr/bin/env python3
"""Archivo 7: ejecucion_protocolo_reunion_semanal.docx"""
import sys; sys.path.insert(0, '/home/user/sistema-dies/scripts')
from docx_helpers import *

OUTPUT = "/mnt/user-data/outputs/ejecucion_protocolo_reunion_semanal.docx"

doc = new_doc()
add_header_footer(doc, "Protocolo de Reunión Semanal")

portada(doc,
    "PROTOCOLO DE REUNIÓN SEMANAL",
    "Revisión de Métricas — 30 Minutos Exactos",
    "Pilar 3: Ejecución | Sistema Comercial Operativo")

doc.add_page_break()

h1(doc, "¿PARA QUÉ SIRVE ESTA REUNIÓN?")
body(doc, "La reunión semanal de métricas es el mecanismo que mantiene el sistema vivo. Sin ella, el CRM se convierte en un archivador y el dashboard en decoración.")
body(doc, "Esta reunión tiene UN objetivo: que el equipo y el dueño sepan si van bien o mal, y que salgan con un plan claro para la semana siguiente.")
doc.add_paragraph()

section_box(doc, "Reglas de oro de la reunión",
    [
        "Duración máxima: 30 minutos. Si se extiende, el problema es de preparación, no de tiempo.",
        "Asistencia obligatoria: León + representante del cliente. El dueño: cuando los números lo requieran.",
        "Sin excepciones: si alguien no puede asistir, se hace virtual. La reunión no se cancela.",
        "Se habla de números, no de excusas: el foco es el dato, no la justificación.",
        "Se cierra con compromisos concretos: quién hace qué, antes de cuándo.",
    ],
    bg_hex="FFF9E6", title_color=DORADO, icon="!")

doc.add_page_break()

h1(doc, "AGENDA FIJA — 30 MINUTOS")
body(doc, "Esta agenda no se improvisa. Se usa exactamente así en cada reunión:")
doc.add_paragraph()

agenda_items = [
    ("0–5 min", "REVISIÓN DE MÉTRICAS", "C9973A",
     [
         "Abrir el dashboard en pantalla compartida",
         "Revisar los 5 KPIs de la semana: conversaciones recibidas / tiempo de respuesta / leads a propuesta / cierres / meta vs. real",
         "Comparar con la semana anterior",
         "Identificar el número más preocupante de la semana",
     ]),
    ("5–15 min", "IDENTIFICACIÓN DE CUELLOS DE BOTELLA", "1A2B4A",
     [
         "Preguntar: '¿Qué frenó el avance esta semana?'",
         "Máximo 2 cuellos de botella por reunión — no más",
         "Para cada cuello: ¿Es un problema de proceso, de persona o de herramienta?",
         "No resolver aquí — solo identificar y asignar responsable de solución",
     ]),
    ("15–25 min", "PLAN DE LA SEMANA SIGUIENTE", "27AE60",
     [
         "Definir la meta de la semana en números concretos",
         "Asignar prioridades: ¿qué es lo más importante que debe pasar?",
         "Revisar los leads en riesgo (etapa 4 con más de 7 días sin movimiento)",
         "Definir qué automatizaciones o ajustes se necesitan",
     ]),
    ("25–30 min", "COMPROMISOS Y CIERRE", "C0392B",
     [
         "Cada asistente dice UNA cosa concreta que va a hacer antes de la próxima reunión",
         "León registra los compromisos en las notas de reunión",
         "Se confirma hora y formato de la próxima reunión",
         "Tiempo máximo para compromisos: 5 minutos",
     ]),
]

for tiempo, titulo, color, puntos in agenda_items:
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, color)
    pt = cell.paragraphs[0]
    pt.paragraph_format.space_before = Pt(8)
    pt.paragraph_format.space_after = Pt(4)
    pt.paragraph_format.left_indent = Cm(0.4)
    rt = pt.add_run(f"[{tiempo}] {titulo}")
    rt.font.bold = True
    rt.font.size = Pt(13)
    rt.font.color.rgb = BLANCO

    for punto in puntos:
        pp = cell.add_paragraph()
        pp.paragraph_format.left_indent = Cm(0.8)
        pp.paragraph_format.space_before = Pt(3)
        pp.paragraph_format.space_after = Pt(3)
        rp = pp.add_run(f"• {punto}")
        rp.font.size = Pt(10)
        rp.font.color.rgb = BLANCO

    pe = cell.add_paragraph()
    pe.paragraph_format.space_after = Pt(8)
    doc.add_paragraph()

doc.add_page_break()

h1(doc, "TEMPLATE DE NOTAS DE REUNIÓN")
body(doc, "Complete este template en cada sesión. Guarde las notas en la carpeta del cliente en Google Drive.")
doc.add_paragraph()

tbl_notas = doc.add_table(rows=12, cols=2)
tbl_notas.style = 'Table Grid'
campos_notas = [
    ("Empresa:", "[Nombre]"),
    ("Fecha:", "[DD/MM/AAAA]"),
    ("Semana del proceso:", "Semana [X] de 8"),
    ("Asistentes:", "[Nombres]"),
    ("Facilita:", "León López"),
    ("KPI más preocupante esta semana:", "[Dato concreto]"),
    ("Cuello de botella #1 identificado:", "[Descripción]"),
    ("Cuello de botella #2 identificado:", "[Descripción]"),
    ("Meta de la semana siguiente:", "$[X] COP / [X] cierres"),
    ("Compromiso #1 (quién / qué / cuándo):", "[Nombre] — [Acción] — [Fecha límite]"),
    ("Compromiso #2 (quién / qué / cuándo):", "[Nombre] — [Acción] — [Fecha límite]"),
    ("Compromiso #3 (quién / qué / cuándo):", "[Nombre] — [Acción] — [Fecha límite]"),
]
for i, (k, v) in enumerate(campos_notas):
    c0 = tbl_notas.cell(i, 0)
    c1 = tbl_notas.cell(i, 1)
    set_cell_bg(c0, "1A2B4A")
    r0 = c0.paragraphs[0].add_run(k)
    r0.font.bold = True
    r0.font.size = Pt(10)
    r0.font.color.rgb = DORADO
    r1 = c1.paragraphs[0].add_run(v)
    r1.font.size = Pt(10)
    r1.font.color.rgb = NEGRO

doc.add_paragraph()

h1(doc, "CÓMO ESCALAR UN PROBLEMA AL DUEÑO")
body(doc, "No todo debe llegar al dueño. Use estos criterios para saber cuándo escalar:")
doc.add_paragraph()

criterios_escalar = [
    ("Escalar INMEDIATAMENTE", "ROJO", [
        "La proyección del mes está por debajo del 70% de la meta",
        "Un asesor clave renunció o está en conflicto abierto",
        "El CRM no está siendo usado por el 50% o más del equipo",
        "Se detectó una pérdida de cliente importante por falla del sistema",
    ]),
    ("Escalar en la REUNIÓN SEMANAL", "AMARILLO", [
        "La proyección del mes está entre 70% y 80% de la meta",
        "Un cuello de botella lleva más de 2 semanas sin resolverse",
        "El equipo necesita una decisión que solo el dueño puede tomar",
    ]),
    ("NO escalar — León lo gestiona", "VERDE", [
        "El equipo tardó 1 semana en adoptar una nueva funcionalidad del CRM",
        "Una automatización falló y ya fue corregida",
        "Un asesor no registró correctamente — ya fue corregido en la reunión",
    ]),
]

for titulo, estado, items in criterios_escalar:
    colores = {"ROJO": "C0392B", "AMARILLO": "F39C12", "VERDE": "27AE60"}
    section_box(doc, titulo, items, bg_hex="F5F5F5",
                title_color=RGBColor(*[int(colores[estado][i:i+2], 16) for i in (0,2,4)]))

doc.add_page_break()

h1(doc, "CRITERIOS PARA SABER SI EL SISTEMA ESTÁ SIENDO ADOPTADO")
body(doc, "Al finalizar la semana 5, el sistema debería mostrar estos indicadores:")
doc.add_paragraph()

semaforo_table(doc, [
    ("Registro en CRM", "verde", "80% o más de los leads registrados el mismo día"),
    ("Tiempo de respuesta", "verde", "Promedio semanal por debajo de 15 minutos"),
    ("Asistencia a reuniones", "verde", "100% del equipo presente o recuperación en 24h"),
    ("Meta individual", "verde", "Cada asesor conoce su meta y sabe cómo va"),
    ("Dashboard activo", "verde", "El dueño lo revisa mínimo 3 veces por semana"),
])

body(doc, "Si alguno de estos indicadores está en rojo después de la semana 5, es una señal de alerta: el sistema no está siendo adoptado y se requiere intervención del dueño.", italic=True)

doc.save(OUTPUT)
print(f"✓ Creado: {OUTPUT}")
