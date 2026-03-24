#!/usr/bin/env python3
"""Archivo 8: ejecucion_protocolo_seguimiento_leads.docx"""
import sys; sys.path.insert(0, '/home/user/sistema-dies/scripts')
from docx_helpers import *

OUTPUT = "/mnt/user-data/outputs/ejecucion_protocolo_seguimiento_leads.docx"

doc = new_doc()
add_header_footer(doc, "Protocolo de Seguimiento de Leads")

portada(doc,
    "PROTOCOLO DE SEGUIMIENTO DE LEADS",
    "Manual de Operación para el Equipo Comercial",
    "Pilar 3: Ejecución | Para uso directo de los asesores")

doc.add_page_break()

h1(doc, "EL PROCESO EN 1 PÁGINA")
body(doc, "(Para imprimir y poner en el escritorio de cada asesor)")
doc.add_paragraph()

# Big visual process table
tbl_proceso = doc.add_table(rows=7, cols=3)
tbl_proceso.style = 'Table Grid'
etapas = [
    ("①", "LEAD NUEVO LLEGA", "Entra a WhatsApp / formulario / Instagram.\nEl CRM lo asigna automáticamente.\nTiempo máximo para responder: 5 MINUTOS"),
    ("②", "PRIMER CONTACTO", "Usa el mensaje de bienvenida estandarizado.\nHaz las 3 preguntas de calificación.\nRegistra en el CRM: canal + producto de interés"),
    ("③", "CALIFICACIÓN", "¿Tiene el presupuesto? ¿Tiene la necesidad? ¿Cuándo necesita?\nSi califica: mueve a 'En conversación activa'\nSi no califica: registra motivo y cierra"),
    ("④", "PROPUESTA / COTIZACIÓN", "Envía la cotización en máximo 24 horas.\nUsa el formato estándar aprobado.\nMueve el lead a 'Propuesta enviada'"),
    ("⑤", "SEGUIMIENTO", "Día 1: recordatorio post-cotización\nDía 3: llamada o mensaje de seguimiento\nDía 7: último intento con oferta de valor\nDía 14: cierre definitivo (ganado o perdido)"),
    ("⑥", "CIERRE", "Ganado: registra valor de venta en el CRM.\nPerdido: registra OBLIGATORIAMENTE la razón.\nAmbos: envía mensaje de confirmación al cliente"),
    ("!", "REGLA DE ORO", "Ningún lead se queda sin registro.\nSi no puedes atenderlo ahora:\n→ Regístralo en el CRM y asigna hora de respuesta"),
]

for i, (num, titulo, desc) in enumerate(etapas):
    row = tbl_proceso.rows[i]
    c0 = row.cells[0]
    c1 = row.cells[1]
    c2 = row.cells[2]

    c0.width = Inches(0.5)
    c1.width = Inches(2.0)
    c2.width = Inches(4.0)

    bg = "1A2B4A" if num != "!" else "C0392B"
    set_cell_bg(c0, bg)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run(num)
    r0.font.bold = True
    r0.font.size = Pt(18)
    r0.font.color.rgb = DORADO if num != "!" else BLANCO

    set_cell_bg(c1, bg)
    p1 = c1.paragraphs[0]
    p1.paragraph_format.left_indent = Cm(0.2)
    r1 = p1.add_run(titulo)
    r1.font.bold = True
    r1.font.size = Pt(11)
    r1.font.color.rgb = DORADO if num != "!" else BLANCO

    p2 = c2.paragraphs[0]
    p2.paragraph_format.left_indent = Cm(0.3)
    r2 = p2.add_run(desc)
    r2.font.size = Pt(10)
    r2.font.color.rgb = NEGRO

doc.add_page_break()

h1(doc, "TIEMPOS MÁXIMOS DE RESPUESTA POR CANAL")
body(doc, "Estos no son sugerencias. Son los tiempos máximos del sistema. Si no se cumplen, el CRM genera una alerta.")
doc.add_paragraph()

tbl_tiempos = doc.add_table(rows=6, cols=4)
tbl_tiempos.style = 'Table Grid'
ht = tbl_tiempos.rows[0].cells
headers_t = ["CANAL", "TIEMPO MÁXIMO", "RESPONSABLE", "QUÉ HACER SI NO PUEDES"]
for i, h in enumerate(headers_t):
    set_cell_bg(ht[i], "1A2B4A")
    ht[i].paragraphs[0].add_run(h).font.color.rgb = DORADO
    ht[i].paragraphs[0].runs[0].font.bold = True
    ht[i].paragraphs[0].runs[0].font.size = Pt(10)

tiempos = [
    ("WhatsApp", "5 MINUTOS", "Asesor asignado", "Activar respuesta automática + agendar llamada"),
    ("Instagram / DMs", "15 MINUTOS", "Asesor de redes", "Responder 'Te contacto en X minutos' + registrar"),
    ("Formulario web", "30 MINUTOS", "Asesor de turno", "Llamar directamente si no responde al mensaje"),
    ("Llamada perdida", "10 MINUTOS", "Asesor de turno", "Devolver la llamada inmediatamente"),
    ("Referido presencial", "Mismo día", "Asesor asignado", "Registrar en CRM y contactar antes de las 6pm"),
]
for i, (canal, tiempo, resp, accion) in enumerate(tiempos):
    row = tbl_tiempos.rows[i+1]
    row.cells[0].paragraphs[0].add_run(canal).font.size = Pt(10)
    row.cells[0].paragraphs[0].runs[0].font.bold = True

    set_cell_bg(row.cells[1], "C9973A")
    row.cells[1].paragraphs[0].add_run(tiempo).font.color.rgb = BLANCO
    row.cells[1].paragraphs[0].runs[0].font.bold = True
    row.cells[1].paragraphs[0].runs[0].font.size = Pt(10)
    row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    row.cells[2].paragraphs[0].add_run(resp).font.size = Pt(10)
    row.cells[3].paragraphs[0].add_run(accion).font.size = Pt(9)
    row.cells[3].paragraphs[0].runs[0].font.italic = True

doc.add_page_break()

h1(doc, "PLANTILLAS DE PRIMER CONTACTO")
body(doc, "Use estas plantillas exactas. No improvise. Si necesita personalizarlas, agréguele el nombre del cliente al principio — nada más.")
doc.add_paragraph()

plantillas = [
    ("PARA SPAS Y ESTÉTICA", "WhatsApp",
     "Hola [Nombre] 😊 Soy [Tu nombre] del equipo de [Nombre del spa]. Vi que nos escribiste sobre [servicio]. ¿Cuándo te gustaría agendar tu cita? Tenemos disponibilidad esta semana. ¿Te parece si te cuento las opciones?"),
    ("PARA CLÍNICAS / SALUD", "WhatsApp",
     "Hola [Nombre], con [Tu nombre] de [Nombre clínica]. Gracias por comunicarte con nosotros. Para ayudarte mejor, ¿me puedes contar qué servicio o especialidad necesitas? Así te busco el especialista y la disponibilidad que mejor se adapte a ti."),
    ("PARA DISTRIBUIDORAS / B2B", "WhatsApp o Email",
     "Buenos días [Nombre], le habla [Tu nombre] de [Empresa]. Vi su solicitud de información sobre [producto/servicio]. ¿Podríamos agendar una llamada de 15 minutos esta semana para entender su necesidad y prepararle una propuesta ajustada? ¿Cuándo tiene disponibilidad?"),
    ("PARA SERVICIOS GENERALES", "Cualquier canal",
     "Hola [Nombre], soy [Tu nombre] de [Empresa]. Gracias por escribirnos. Estoy revisando tu consulta ahora mismo. ¿Me puedes contar un poco más sobre lo que necesitas para ayudarte mejor?"),
]

for tipo, canal, mensaje in plantillas:
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, "1A2B4A")

    pt = cell.paragraphs[0]
    pt.paragraph_format.space_before = Pt(6)
    pt.paragraph_format.left_indent = Cm(0.4)
    rt = pt.add_run(f"{tipo}  ·  Canal: {canal}")
    rt.font.bold = True
    rt.font.size = Pt(11)
    rt.font.color.rgb = DORADO

    pm = cell.add_paragraph()
    pm.paragraph_format.left_indent = Cm(0.6)
    pm.paragraph_format.space_before = Pt(4)
    pm.paragraph_format.space_after = Pt(8)
    rm = pm.add_run(f'"{mensaje}"')
    rm.font.size = Pt(10)
    rm.font.color.rgb = BLANCO
    rm.font.italic = True
    doc.add_paragraph()

h1(doc, "PROTOCOLO DE SEGUIMIENTO — DÍA A DÍA")
body(doc, "Para leads que mostraron interés pero no cerraron en el primer contacto:")
doc.add_paragraph()

seguimientos = [
    ("DÍA 1", "post-cotización", "C9973A",
     'Hola [Nombre], te escribo para ver si pudiste revisar la información que te compartí sobre [servicio/producto]. ¿Tienes alguna pregunta? Estoy aquí para ayudarte.'),
    ("DÍA 3", "recordatorio", "1A2B4A",
     'Hola [Nombre], quería hacer un seguimiento. Sé que estás evaluando opciones. ¿Hay algo específico que te ayudaría a tomar la decisión? Puedo agendarte una llamada rápida si prefieres.'),
    ("DÍA 7", "último intento activo", "F39C12",
     'Hola [Nombre], te escribo por última vez sobre tu consulta de [servicio]. Esta semana tenemos [disponibilidad especial / oferta / cupo limitado]. Si te interesa, avísame hoy y lo apartamos para ti.'),
    ("DÍA 14", "cierre definitivo", "27AE60",
     'Hola [Nombre], quiero cerrar tu caso en nuestro sistema. ¿Tomaste una decisión? Si en algún momento necesitas [servicio] en el futuro, con gusto te atiendo. ¡Gracias por tu tiempo!'),
]

for dia, desc, color, mensaje in seguimientos:
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Table Grid'
    c0 = tbl.cell(0, 0)
    c0.width = Inches(1.4)
    set_cell_bg(c0, color)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run(dia)
    r0.font.bold = True
    r0.font.size = Pt(14)
    r0.font.color.rgb = BLANCO
    p0b = c0.add_paragraph()
    p0b.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0b = p0b.add_run(desc)
    r0b.font.size = Pt(8)
    r0b.font.italic = True
    r0b.font.color.rgb = BLANCO

    c1 = tbl.cell(0, 1)
    p1 = c1.paragraphs[0]
    p1.paragraph_format.left_indent = Cm(0.3)
    r1 = p1.add_run(f'"{mensaje}"')
    r1.font.size = Pt(10)
    r1.font.italic = True
    r1.font.color.rgb = NEGRO
    doc.add_paragraph()

doc.add_page_break()

h1(doc, "CÓMO REGISTRAR EN EL CRM — PASO A PASO")
body(doc, "Cada interacción con un lead debe quedar registrada. Sin registro, no existe.")
doc.add_paragraph()

pasos_crm = [
    ("Paso 1: Buscar o crear el lead",
     "Al llegar un mensaje nuevo, abrir GoHighLevel → Buscar el número en Contactos.\nSi no existe: clic en 'Nuevo Contacto' → ingresar nombre, número y canal de origen."),
    ("Paso 2: Mover en el pipeline",
     "Arrastrar el lead a la etapa correcta:\n→ 'En conversación activa' si ya respondiste\n→ 'Propuesta enviada' si ya cotizaste\n→ 'Seguimiento programado' si espera respuesta"),
    ("Paso 3: Agregar nota",
     "En la ficha del lead → 'Agregar nota' → escribir en 1-2 líneas qué pasó en la conversación.\nEjemplo: 'Interesado en masaje + facial. Tiene presupuesto. Espera confirmación de pareja.'"),
    ("Paso 4: Programar seguimiento",
     "Si el lead no cerró: hacer clic en 'Tarea' → escribir qué vas a hacer → fecha y hora.\nEl sistema te recordará automáticamente."),
    ("Paso 5: Cerrar el lead",
     "Cuando cierre (compró o perdió):\n→ Ganado: mover a 'Cerrado Ganado' + ingresar valor de la venta\n→ Perdido: mover a 'Cerrado Perdido' + OBLIGATORIO seleccionar la razón del dropdown"),
]

for i, (titulo, desc) in enumerate(pasos_crm):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Table Grid'
    c0 = tbl.cell(0, 0)
    c0.width = Inches(0.8)
    set_cell_bg(c0, "1A2B4A")
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run(str(i+1))
    r0.font.bold = True
    r0.font.size = Pt(20)
    r0.font.color.rgb = DORADO

    c1 = tbl.cell(0, 1)
    p1t = c1.paragraphs[0]
    p1t.paragraph_format.left_indent = Cm(0.3)
    rt = p1t.add_run(titulo)
    rt.font.bold = True
    rt.font.size = Pt(11)
    rt.font.color.rgb = AZUL

    p1d = c1.add_paragraph()
    p1d.paragraph_format.left_indent = Cm(0.5)
    rd = p1d.add_run(desc)
    rd.font.size = Pt(10)
    rd.font.color.rgb = NEGRO
    doc.add_paragraph()

h1(doc, "¿QUÉ HACER CON UN LEAD QUE NO RESPONDE?")
doc.add_paragraph()

section_box(doc, "Protocolo para lead sin respuesta",
    [
        "Después de 3 intentos (día 1, día 3, día 7) sin respuesta:",
        "",
        "1. Registrar en el CRM: 'Lead sin respuesta — 3 intentos realizados'",
        "2. Mover a 'Cerrado Perdido' con razón: 'Sin respuesta del lead'",
        "3. NO volver a contactar. Respetar la decisión implícita del prospecto.",
        "4. El sistema puede enviar 1 último mensaje automático al día 14 (configurable).",
        "",
        "NUNCA: insistir más de 3 veces antes del día 7. El sistema registra cada intento.",
    ],
    bg_hex="FDECEA", title_color=ROJO, icon="×")

h1(doc, "¿CÓMO ESCALAR UN LEAD CALIENTE?")
section_box(doc, "Si tienes un lead caliente que no puedes atender ahora mismo",
    [
        "1. Responde de inmediato con: 'Hola [Nombre], gracias por escribir. Estoy con un cliente en este momento. Te llamo / escribo en máximo 20 minutos. ¿Te parece bien?'",
        "2. Regístralo en el CRM como URGENTE (etiqueta 'Lead Caliente')",
        "3. Notifica al supervisor o al asesor de turno para que haga cobertura",
        "4. Retoma el contacto DENTRO de los 20 minutos prometidos — sin excepción",
        "",
        "Un lead caliente que espera más de 30 minutos se enfría. Las ventas se pierden ahí.",
    ],
    bg_hex="E8F8F0", title_color=VERDE, icon="→")

doc.save(OUTPUT)
print(f"✓ Creado: {OUTPUT}")
