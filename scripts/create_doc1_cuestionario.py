#!/usr/bin/env python3
"""Archivo 1: diagnostico_cuestionario.docx"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

OUTPUT = "/mnt/user-data/outputs/diagnostico_cuestionario.docx"

AZUL = RGBColor(0x1A, 0x2B, 0x4A)
DORADO = RGBColor(0xC9, 0x97, 0x3A)
GRIS = RGBColor(0xF5, 0xF5, 0xF5)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)
GRIS_TEXTO = RGBColor(0x55, 0x55, 0x55)


def set_cell_bg(cell, color_hex):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)


def add_header(doc):
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("SISTEMA COMERCIAL OPERATIVO | León López")
    run.font.size = Pt(9)
    run.font.color.rgb = AZUL
    run.font.bold = True
    # Right side
    run2 = p.add_run("                                                    [LOGO EMPRESA]")
    run2.font.size = Pt(9)
    run2.font.color.rgb = DORADO


def add_footer(doc):
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Diagnóstico Comercial | Versión 1.0 | Marzo 2026 | Confidencial")
    run.font.size = Pt(8)
    run.font.color.rgb = GRIS_TEXTO


def add_title_block(doc, title, subtitle=None):
    # Title table for portada
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_bg(cell, "1A2B4A")
    cell.width = Inches(6.5)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(title)
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = BLANCO

    if subtitle:
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(4)
        p2.paragraph_format.space_after = Pt(20)
        r2 = p2.add_run(subtitle)
        r2.font.size = Pt(12)
        r2.font.color.rgb = DORADO


def heading1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = AZUL
    # underline decoration via border bottom
    return p


def heading2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = DORADO
    return p


def body(doc, text, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    run.font.italic = italic
    return p


def add_question(doc, num, question, opciones=None, espacio_respuesta=True):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.5)
    run_num = p.add_run(f"{num}. ")
    run_num.font.bold = True
    run_num.font.color.rgb = AZUL
    run_num.font.size = Pt(11)
    run_q = p.add_run(question)
    run_q.font.size = Pt(11)
    run_q.font.color.rgb = RGBColor(0x22, 0x22, 0x22)

    if opciones:
        p2 = doc.add_paragraph()
        p2.paragraph_format.left_indent = Cm(1.2)
        p2.paragraph_format.space_before = Pt(2)
        p2.paragraph_format.space_after = Pt(2)
        run_op = p2.add_run("Opciones: " + " / ".join(opciones))
        run_op.font.size = Pt(10)
        run_op.font.italic = True
        run_op.font.color.rgb = GRIS_TEXTO

    if espacio_respuesta:
        p3 = doc.add_paragraph()
        p3.paragraph_format.left_indent = Cm(1.2)
        p3.paragraph_format.space_before = Pt(2)
        p3.paragraph_format.space_after = Pt(6)
        run_esp = p3.add_run("Respuesta: _______________________________________________________________________________")
        run_esp.font.size = Pt(10)
        run_esp.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)

    return p


def add_open_question(doc, num, question):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.5)
    run_num = p.add_run(f"{num}. ")
    run_num.font.bold = True
    run_num.font.color.rgb = AZUL
    run_num.font.size = Pt(11)
    run_q = p.add_run(question)
    run_q.font.size = Pt(11)
    run_q.font.color.rgb = RGBColor(0x22, 0x22, 0x22)

    for _ in range(3):
        p_line = doc.add_paragraph()
        p_line.paragraph_format.left_indent = Cm(1.2)
        p_line.paragraph_format.space_before = Pt(2)
        p_line.paragraph_format.space_after = Pt(2)
        r = p_line.add_run("_______________________________________________________________________________")
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)


def add_block_header(doc, num, title, description):
    doc.add_paragraph()
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_bg(cell, "1A2B4A")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.left_indent = Cm(0.3)
    run = p.add_run(f"BLOQUE {num} — {title.upper()}")
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = BLANCO

    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p2.paragraph_format.space_before = Pt(2)
    p2.paragraph_format.space_after = Pt(8)
    p2.paragraph_format.left_indent = Cm(0.3)
    r2 = p2.add_run(description)
    r2.font.size = Pt(10)
    r2.font.color.rgb = DORADO
    r2.font.italic = True
    doc.add_paragraph()


def add_separator(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run("─" * 80)
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)


# ─── BUILD DOC ────────────────────────────────────────────────────────────────

doc = Document()

# Page margins
for section in doc.sections:
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)

# Default paragraph font
doc.styles['Normal'].font.name = 'Calibri'
doc.styles['Normal'].font.size = Pt(11)

add_header(doc)
add_footer(doc)

# ── PORTADA ──────────────────────────────────────────────────────────────────
p_logo = doc.add_paragraph()
p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_logo.paragraph_format.space_before = Pt(40)
run_logo = p_logo.add_run("[LOGO DE LA EMPRESA AQUÍ]")
run_logo.font.size = Pt(12)
run_logo.font.color.rgb = GRIS_TEXTO
run_logo.font.italic = True

doc.add_paragraph()
add_title_block(doc,
    "DIAGNÓSTICO COMERCIAL",
    "Cuestionario de Radiografía Operativa | 60–90 minutos")

doc.add_paragraph()
p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_sub = p_sub.add_run("Sistema Comercial Operativo | León López Consultoría")
r_sub.font.size = Pt(11)
r_sub.font.color.rgb = GRIS_TEXTO

p_fecha = doc.add_paragraph()
p_fecha.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_fecha = p_fecha.add_run("Fecha de sesión: _________________________   Empresa: _________________________")
r_fecha.font.size = Pt(11)
r_fecha.font.color.rgb = GRIS_TEXTO

doc.add_page_break()

# ── INSTRUCCIONES PARA LEÓN ───────────────────────────────────────────────────
heading1(doc, "INSTRUCCIONES PARA EL FACILITADOR (León López)")

# Box with instructions
table = doc.add_table(rows=1, cols=1)
cell = table.cell(0, 0)
set_cell_bg(cell, "F5F5F5")
p_inst = cell.paragraphs[0]
p_inst.paragraph_format.space_before = Pt(8)
p_inst.paragraph_format.space_after = Pt(4)
p_inst.paragraph_format.left_indent = Cm(0.5)
run_t = p_inst.add_run("Cómo facilitar esta sesión:")
run_t.font.bold = True
run_t.font.size = Pt(11)
run_t.font.color.rgb = AZUL

instrucciones = [
    "ACTITUD: Eres un diagnosta, no un vendedor. Tu trabajo es entender, no convencer. Haz silencio cuando el cliente habla.",
    "TIEMPO: Asigna máximo 5 minutos por bloque para los bloques 0–5. El Bloque 6 (dolor) puede extenderse — ahí está el oro.",
    "ESCUCHA: Anota las palabras exactas que usa el cliente. No parafrasees. Esas palabras son su lenguaje, úsalo en el informe.",
    "NÚMEROS: Si el cliente dice 'no sé' a una cifra, ayúdalo a estimarla: '¿Cuántos mensajes llegan en una mañana típica?'",
    "SEÑALES: Cuando el cliente se ponga emocional (frustración, orgullo, miedo), profundiza con '¿Y eso qué le ha costado?'",
    "REGISTRO: Completa este documento durante la sesión o inmediatamente después. La calculadora se llena en paralelo.",
    "CIERRE: No propongas soluciones durante el diagnóstico. Solo di: 'Con esto tengo lo suficiente para preparar su informe.'",
]

for i, inst in enumerate(instrucciones):
    pi = cell.add_paragraph()
    pi.paragraph_format.left_indent = Cm(0.5)
    pi.paragraph_format.space_before = Pt(3)
    pi.paragraph_format.space_after = Pt(3)
    ri = pi.add_run(f"{'✓'} {inst}")
    ri.font.size = Pt(10)
    ri.font.color.rgb = RGBColor(0x22, 0x22, 0x22)

p_end = cell.add_paragraph()
p_end.paragraph_format.space_after = Pt(8)

doc.add_paragraph()

# ── BLOQUE 0 ──────────────────────────────────────────────────────────────────
add_block_header(doc, "0", "Datos de la Empresa",
    "5 preguntas | Objetivo: tener el contexto base antes de profundizar")

add_question(doc, 1, "¿Cuál es el nombre completo de la empresa y a qué sector pertenece?")
add_question(doc, 2, "¿En qué ciudad o ciudades opera actualmente?")
add_question(doc, 3, "¿Cuántos colaboradores tienen en total? ¿Cómo se distribuyen por área?",
    espacio_respuesta=False)
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(1.2)
r = p.add_run("Total: _______   Comercial/Ventas: _______   Operativo/Servicio: _______   Admin: _______")
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
doc.add_paragraph()

add_question(doc, 4, "¿Cuál es la facturación mensual aproximada de la empresa?",
    opciones=["Menos de $100M COP", "$100M – $300M COP", "$300M – $1.000M COP", "Más de $1.000M COP"])
add_question(doc, 5, "¿Cuántos años llevan operando? ¿Actualmente tienen algún CRM activo? ¿Cuál?")

# ── BLOQUE 1 ──────────────────────────────────────────────────────────────────
add_block_header(doc, "1", "Flujo de Leads",
    "8 preguntas | Objetivo: entender por dónde llegan los clientes y qué les pasa")

add_question(doc, 1, "¿Por qué canales llegan sus prospectos hoy? Indique el porcentaje aproximado de cada uno.",
    espacio_respuesta=False)
table_canales = doc.add_table(rows=2, cols=6)
table_canales.alignment = WD_TABLE_ALIGNMENT.LEFT
headers = ["WhatsApp", "Instagram", "Referidos", "Presencial", "Llamadas", "Web/Forms"]
for i, h in enumerate(headers):
    cell = table_canales.cell(0, i)
    set_cell_bg(cell, "1A2B4A")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(h)
    r.font.size = Pt(9)
    r.font.bold = True
    r.font.color.rgb = BLANCO
    cell2 = table_canales.cell(1, i)
    p2 = cell2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("_____%")
    r2.font.size = Pt(10)
    r2.font.color.rgb = GRIS_TEXTO
doc.add_paragraph()

add_question(doc, 2, "¿Cuántas conversaciones nuevas reciben por día o por semana en cada canal?")
add_question(doc, 3, "¿Quién atiende cada canal? ¿Es la misma persona para todos o hay diferentes responsables?")
add_question(doc, 4, "¿Cuánto tiempo tarda en promedio el equipo en responder un mensaje nuevo?",
    opciones=["Menos de 5 min", "5–30 min", "30 min–2 horas", "Más de 2 horas", "No sabemos"])
add_question(doc, 5, "¿Tienen un protocolo escrito de respuesta, o cada persona responde como puede?")
add_question(doc, 6, "¿Usan chatbots o respuestas automáticas? ¿Están funcionando bien?")
add_question(doc, 7, "¿Saben cuántas conversaciones se quedan sin respuesta cada día? ¿Hay algún registro?")
add_question(doc, 8, "¿Tienen forma de saber de dónde viene cada cliente que termina comprando?")

# ── BLOQUE 2 ──────────────────────────────────────────────────────────────────
add_block_header(doc, "2", "Proceso de Cierre",
    "8 preguntas | Objetivo: mapear qué tan estructurado (o no) es el proceso de conversión")

add_question(doc, 1, "¿Cuántos pasos tiene el proceso desde que llega un lead hasta que paga? ¿Están definidos o cada persona lo hace a su manera?")
add_question(doc, 2, "¿Quién tiene autorización para dar un precio? ¿Existe un protocolo de cotización definido?")
add_question(doc, 3, "¿Qué porcentaje aproximado de las personas que consultan terminan agendando o comprando?",
    opciones=["Menos del 10%", "10%–20%", "20%–35%", "Más del 35%", "No tenemos el dato"])
add_question(doc, 4, "¿Qué pasa con los que no compran en el primer contacto? ¿Alguien les hace seguimiento? ¿Cuántas veces?")
add_question(doc, 5, "¿Tienen una plantilla o guión para el seguimiento, o se improvisa en cada caso?")
add_question(doc, 6, "¿Cuántos días pasan en promedio entre el primer contacto y el cierre de una venta?")
add_question(doc, 7, "¿Tienen definido cuándo un lead se considera 'perdido'? ¿Hay criterios claros?")
add_question(doc, 8, "¿Registran en algún lado las razones por las que los clientes no compran?")

# ── BLOQUE 3 ──────────────────────────────────────────────────────────────────
add_block_header(doc, "3", "Estructura del Equipo Comercial",
    "8 preguntas | Objetivo: entender quién vende, cómo y con qué visibilidad")

add_question(doc, 1, "¿Quién vende actualmente en la empresa?",
    opciones=["Solo el dueño", "Equipo dedicado a ventas", "Todos venden un poco", "Nadie tiene ese rol formal"])
add_question(doc, 2, "¿Cada persona del equipo sabe exactamente qué meta debe cumplir este mes, en pesos o en unidades?")
add_question(doc, 3, "¿Cómo saben si van bien o mal durante el mes? ¿En qué momento se enteran?")
add_question(doc, 4, "¿El dueño tiene visibilidad directa del estado de la operación, o debe preguntar para saber cómo va?")
add_question(doc, 5, "¿Tienen reuniones de seguimiento comercial? ¿Con qué frecuencia se hacen?",
    opciones=["Diarias", "Semanales", "Mensuales", "No hay reuniones formales"])
add_question(doc, 6, "¿El equipo sabe qué hacer si un asesor se enferma o renuncia? ¿Los procesos están documentados?")
add_question(doc, 7, "¿Han perdido clientes o cerrado mal negocios por falta de comunicación interna?")
add_question(doc, 8, "¿Qué tan dependiente es la operación de que el dueño esté presente?",
    opciones=["1 – Opera solo", "2", "3 – Necesita supervisión", "4", "5 – Sin el dueño todo para"])

# ── BLOQUE 4 ──────────────────────────────────────────────────────────────────
add_block_header(doc, "4", "Herramientas y Tecnología",
    "6 preguntas | Objetivo: inventariar el stack tecnológico real vs. el usado")

add_question(doc, 1, "¿Qué herramientas usan hoy para gestionar clientes y prospectos?",
    opciones=["WhatsApp Business", "Excel / Sheets", "CRM (¿cuál?)", "Cuaderno / notas", "Ninguna formal"])
add_question(doc, 2, "Si tienen CRM, ¿qué porcentaje del equipo lo usa de verdad en su día a día?",
    opciones=["Menos del 20%", "20%–50%", "50%–80%", "Más del 80%", "No aplica"])
add_question(doc, 3, "¿Dónde está guardada la información de clientes y prospectos? ¿Está centralizada o dispersa entre personas?")
add_question(doc, 4, "¿Tienen automatizaciones activas? ¿Cuáles funcionan bien y cuáles no?")
add_question(doc, 5, "¿Integran WhatsApp con alguna herramienta de gestión o CRM?")
add_question(doc, 6, "¿Quién en el equipo tiene acceso a qué información de clientes? ¿Hay niveles de acceso definidos?")

# ── BLOQUE 5 ──────────────────────────────────────────────────────────────────
add_block_header(doc, "5", "Visibilidad Gerencial",
    "6 preguntas | Objetivo: entender qué ve (y qué no ve) el dueño hoy")

add_question(doc, 1, "¿Qué métricas revisa el dueño de forma regular hoy?")
add_question(doc, 2, "¿Con qué frecuencia revisa el estado del negocio?",
    opciones=["Diariamente", "Semanalmente", "Mensualmente", "Cuando algo falla"])
add_question(doc, 3, "¿Qué información NO tiene hoy que le gustaría tener para tomar mejores decisiones?")
add_question(doc, 4, "¿Cuánto tiempo le toma entender cómo va el negocio cuando necesita saberlo?")
add_question(doc, 5, "¿Han perdido oportunidades importantes porque nadie las vio a tiempo? ¿Puede dar un ejemplo?")
add_question(doc, 6, "Si tuviera un tablero con 5 indicadores visibles cada mañana, ¿cuáles elegiría?")

# ── BLOQUE 6 ──────────────────────────────────────────────────────────────────
add_block_header(doc, "6", "Dolor y Motivación",
    "5 preguntas abiertas | Objetivo: encontrar el dolor real y la urgencia verdadera — tómese su tiempo aquí")

# Note for León
p_nota = doc.add_paragraph()
p_nota.paragraph_format.left_indent = Cm(0.5)
p_nota.paragraph_format.space_before = Pt(4)
p_nota.paragraph_format.space_after = Pt(8)
r_nota = p_nota.add_run("NOTA PARA LEÓN: Estas preguntas son conversacionales. No las lea como formulario. Úselas como guía para una conversación abierta. El silencio después de la pregunta es parte de la técnica.")
r_nota.font.size = Pt(10)
r_nota.font.italic = True
r_nota.font.color.rgb = DORADO

add_open_question(doc, 1, '"¿Cuál es la situación que más le quita el sueño de su operación comercial hoy?"')
doc.add_paragraph()
add_open_question(doc, 2, '"Si pudiera cambiar UNA SOLA COSA de cómo funciona su equipo comercial hoy, ¿qué sería?"')
doc.add_paragraph()
add_open_question(doc, 3, '"¿Qué ha intentado antes para resolver esto? ¿Qué pasó con eso?"')
doc.add_paragraph()
add_open_question(doc, 4, '"¿Qué significaría para usted tener visibilidad completa de su operación sin tener que preguntar a nadie?"')
doc.add_paragraph()

p_urg = doc.add_paragraph()
p_urg.paragraph_format.left_indent = Cm(0.5)
p_urg.paragraph_format.space_before = Pt(8)
r_urg = p_urg.add_run('5. "En una escala de 1 a 10, ¿qué tan urgente es resolver esto para usted ahora mismo?"')
r_urg.font.size = Pt(11)
r_urg.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
r_urg.font.bold = True

p_scale = doc.add_paragraph()
p_scale.paragraph_format.left_indent = Cm(1.2)
p_scale.paragraph_format.space_before = Pt(6)
r_scale = p_scale.add_run("1 ○   2 ○   3 ○   4 ○   5 ○   6 ○   7 ○   8 ○   9 ○   10 ○")
r_scale.font.size = Pt(14)
r_scale.font.color.rgb = AZUL

p_comment = doc.add_paragraph()
p_comment.paragraph_format.left_indent = Cm(1.2)
p_comment.paragraph_format.space_before = Pt(4)
p_comment.paragraph_format.space_after = Pt(4)
r_comment = p_comment.add_run("¿Por qué ese número y no uno más alto? ____________________________________________")
r_comment.font.size = Pt(10)
r_comment.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)

doc.add_paragraph()
add_separator(doc)

# ── NOTAS FINALES ─────────────────────────────────────────────────────────────
heading1(doc, "NOTAS DEL FACILITADOR")

for i in range(8):
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_before = Pt(10)
    r_l = p_line.add_run("_______________________________________________________________________________")
    r_l.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    r_l.font.size = Pt(11)

doc.add_paragraph()

p_cierre = doc.add_paragraph()
p_cierre.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_cierre = p_cierre.add_run("Con este diagnóstico León prepara el informe en 48–72 horas. El cliente recibe su radiografía completa con el costo de oportunidad calculado en pesos COP.")
r_cierre.font.size = Pt(10)
r_cierre.font.italic = True
r_cierre.font.color.rgb = GRIS_TEXTO

doc.save(OUTPUT)
print(f"✓ Creado: {OUTPUT}")
