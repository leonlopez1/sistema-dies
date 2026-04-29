#!/usr/bin/env python3
"""Archivo 9: socializacion_manual_dueno.docx"""
import sys; sys.path.insert(0, '/home/user/sistema-dies/scripts')
from docx_helpers import *

OUTPUT = "/mnt/user-data/outputs/socializacion_manual_dueno.docx"

doc = new_doc()
add_header_footer(doc, "Manual del Dueño")

portada(doc,
    "MANUAL DEL DUEÑO",
    "Cómo Tener Visibilidad Total Sin Preguntar a Nadie",
    "Pilar 4: Socialización | Guía de 10 minutos para el CEO")

doc.add_page_break()

h1(doc, "ANTES DE EMPEZAR — LO MÁS IMPORTANTE")
section_box(doc, "Su nuevo trabajo como dueño del sistema",
    [
        "Antes del sistema: usted PREGUNTABA para saber cómo iba el negocio.",
        "Con el sistema: usted OBSERVA. Los datos llegan a usted, no al revés.",
        "",
        "Su rol ahora tiene 3 responsabilidades concretas:",
        "  1. Revisar el dashboard 5 minutos cada mañana",
        "  2. Asistir a la reunión semanal de 30 minutos",
        "  3. Exigir al equipo que use el sistema — sin excepciones",
        "",
        "Si hace estas 3 cosas, el sistema funciona. Si las delega, el sistema muere.",
    ],
    bg_hex="FFF9E6", title_color=DORADO, icon="!")

doc.add_page_break()

h1(doc, "1. LOS 5 INDICADORES QUE DEBE VER CADA MAÑANA")
body(doc, "Abra el dashboard en 5 minutos. Revise estos 5 números en este orden exacto:")
doc.add_paragraph()

indicadores = [
    ("1", "Tiempo promedio de respuesta (hoy)", ROJO,
     "¿Está por debajo de 15 minutos? Verde. ¿Entre 15 y 30? Amarillo. ¿Más de 30? Acción inmediata.",
     "Si está en rojo: llame al líder del equipo. No al asesor. Al líder."),
    ("2", "Leads recibidos vs. leads atendidos (ayer)", AZUL,
     "¿Respondieron el 90% o más de los leads? Van bien. ¿Menos del 80%? Hay un problema operativo.",
     "Si hay leads sin atender de ayer: preguntar por qué y exigir protocolo de cobertura."),
    ("3", "Posición actual en el embudo", DORADO,
     "¿Cuántos leads están en propuesta? ¿Cuántos en seguimiento? ¿Cuántos llevan más de 7 días sin movimiento?",
     "Los leads estancados son ventas que se están enfriando. Ese número debe bajar cada semana."),
    ("4", "Meta del mes — % de avance", VERDE,
     "¿Qué porcentaje de la meta mensual se ha cumplido hasta hoy? ¿Va al ritmo correcto?",
     "Ritmo correcto: si estamos al 40% del mes, deberíamos tener al menos el 40% de la meta."),
    ("5", "Alertas activas", ROJO,
     "¿Hay alertas sin atender en el sistema? ¿Cuántas llevan más de 24 horas?",
     "Una alerta no atendida es una oportunidad perdida con fecha y valor calculado."),
]

for num, titulo, color, interpretacion, accion in indicadores:
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Table Grid'
    c0 = tbl.cell(0, 0)
    c0.width = Inches(0.7)
    set_cell_bg(c0, "1A2B4A")
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run(num)
    r0.font.bold = True
    r0.font.size = Pt(20)
    r0.font.color.rgb = DORADO

    c1 = tbl.cell(0, 1)
    pt = c1.paragraphs[0]
    pt.paragraph_format.left_indent = Cm(0.3)
    rt = pt.add_run(titulo)
    rt.font.bold = True
    rt.font.size = Pt(12)
    rt.font.color.rgb = AZUL

    pi = c1.add_paragraph()
    pi.paragraph_format.left_indent = Cm(0.5)
    ri = pi.add_run(interpretacion)
    ri.font.size = Pt(10)
    ri.font.color.rgb = NEGRO

    pa = c1.add_paragraph()
    pa.paragraph_format.left_indent = Cm(0.5)
    pa.paragraph_format.space_after = Pt(4)
    ra = pa.add_run(f"Acción si hay problema: {accion}")
    ra.font.size = Pt(10)
    ra.font.italic = True
    ra.font.color.rgb = GRIS_TEXTO

    doc.add_paragraph()

doc.add_page_break()

h1(doc, "2. SEÑALES DE ALERTA — CUÁNDO INTERVENIR")
body(doc, "La mayoría del tiempo, usted observa. Pero hay señales concretas que requieren su intervención directa:")
doc.add_paragraph()

semaforo_table(doc, [
    ("Tiempo de respuesta", "verde", "Promedio semanal por debajo de 10 minutos → No intervenir"),
    ("Tiempo de respuesta", "amarillo", "Entre 10 y 25 minutos → Preguntar en la reunión semanal"),
    ("Tiempo de respuesta", "rojo", "Más de 25 minutos → Intervención directa HOY"),
    ("Avance de meta", "verde", "Por encima del 85% del ritmo esperado → No intervenir"),
    ("Avance de meta", "amarillo", "Entre 70% y 85% → Revisar en reunión semanal + plan de choque"),
    ("Avance de meta", "rojo", "Por debajo del 70% → Reunión de emergencia esta semana"),
    ("Adopción del CRM", "verde", "Más del 85% de leads registrados → Sistema funcionando"),
    ("Adopción del CRM", "amarillo", "Entre 65% y 85% → Recordatorio en reunión"),
    ("Adopción del CRM", "rojo", "Menos del 65% → El equipo no está adoptando. Necesita su peso."),
])

doc.add_page_break()

h1(doc, "3. CÓMO LEER EL EMBUDO DE CONVERSIÓN")
body(doc, "El embudo le muestra dónde están las oportunidades y dónde se están perdiendo.")
doc.add_paragraph()

body(doc, "Un embudo sano se ve así (proporciones aproximadas):")
tbl_embudo = doc.add_table(rows=6, cols=3)
tbl_embudo.style = 'Table Grid'
etapas_embudo = [
    ("Lead Nuevo", "100%", "Todos los que llegan"),
    ("En conversación activa", "70-80%", "Los que responden"),
    ("Propuesta enviada", "40-50%", "Los que muestran interés real"),
    ("Seguimiento programado", "20-30%", "Los que no cierran inmediato"),
    ("Cerrado ganado", "15-20%", "Los que compran"),
    ("Cerrado perdido", "80-85%", "Los que no compran — registrar SIEMPRE la razón"),
]
h_eb = tbl_embudo.rows[0].cells
for i, h in enumerate(["ETAPA", "% ESPERADO", "QUÉ SIGNIFICA"]):
    set_cell_bg(h_eb[i], "1A2B4A")
    h_eb[i].paragraphs[0].add_run(h).font.color.rgb = DORADO
    h_eb[i].paragraphs[0].runs[0].font.bold = True
    h_eb[i].paragraphs[0].runs[0].font.size = Pt(10)

for i, (etapa, pct, significado) in enumerate(etapas_embudo):
    row = tbl_embudo.rows[i+1] if i < 5 else tbl_embudo.add_row()
    row.cells[0].paragraphs[0].add_run(etapa).font.size = Pt(10)
    row.cells[0].paragraphs[0].runs[0].font.bold = True
    row.cells[1].paragraphs[0].add_run(pct).font.size = Pt(10)
    row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    row.cells[2].paragraphs[0].add_run(significado).font.size = Pt(10)

doc.add_paragraph()
body(doc, "¿Dónde buscar la fuga? Si un porcentaje está muy por debajo de lo esperado, AHÍ está el problema de la semana. No hay que revisar todo — solo la caída más grande entre una etapa y la siguiente.", italic=True)

doc.add_page_break()

h1(doc, "4. CÓMO DAR RETROALIMENTACIÓN AL EQUIPO CON DATOS")
body(doc, "El sistema le da los datos. Usted decide cómo usarlos con el equipo. Estas son las reglas:")
doc.add_paragraph()

for item in [
    "Nunca hable de intuición cuando tiene datos. 'Siento que están respondiendo tarde' vs. 'Esta semana el tiempo promedio fue de 38 minutos — eso es el doble del estándar.'",
    "El dato es neutro. La interpretación es suya. El dato no acusa — informa.",
    "Cuando un asesor va bien, dígalo en público con el número. 'Juan cerró el 22% de sus leads esta semana — eso es lo que buscamos.'",
    "Cuando un asesor va mal, dígalo en privado con el número. 'Veo que esta semana registraste 8 de los 23 leads que te llegaron. ¿Qué pasó?'",
    "Nunca use los datos como amenaza. Úselos como conversación.",
]:
    bullet(doc, item)

doc.add_paragraph()

h1(doc, "5. CÓMO USAR LA REUNIÓN SEMANAL DE 30 MINUTOS")
body(doc, "Su rol en la reunión semanal:")
doc.add_paragraph()

roles_dueno = [
    ("Qué debe llevar", "El dashboard revisado. Los 2 números que más le preocupan de la semana."),
    ("Qué pregunta", "'¿Qué los frenó esta semana?' y '¿Qué necesitan de mí para que no pase de nuevo?'"),
    ("Qué NO hace", "Resolver problemas operativos en la reunión. Eso es trabajo de León y del equipo."),
    ("Cómo cierra", "Confirmando los compromisos y exigiendo que se cumplan ANTES de la próxima reunión."),
    ("Cuándo asiste", "Siempre que la proyección esté por debajo del 80% de la meta. En otros casos: opcional."),
]

for titulo, desc in roles_dueno:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.left_indent = Cm(0.5)
    rb = p.add_run(f"• {titulo}: ")
    rb.font.bold = True
    rb.font.color.rgb = AZUL
    rd = p.add_run(desc)
    rd.font.size = Pt(11)
    rd.font.color.rgb = NEGRO

doc.add_page_break()

h1(doc, "6. QUÉ HACER SI EL EQUIPO DEJA DE USAR EL SISTEMA")
body(doc, "Es normal que después de las primeras 2-3 semanas de emoción, el equipo empiece a volver a viejos hábitos. Esto es lo que hace el dueño cuando eso pasa:")
doc.add_paragraph()

acciones_resistencia = [
    ("Semana 1–2 de caída", "Recordatorio directo del dueño al líder del equipo. No a todos. Solo al líder. 'Vi que el registro cayó esta semana. Necesito que esto se corrija antes del viernes.'"),
    ("Semana 3 de caída", "Reunión 1-a-1 con la persona que más está fallando. Sin acusaciones. 'Qué está pasando con el registro. Qué necesito hacer yo para que esto funcione para ti.'"),
    ("Mes de caída sin mejora", "León convoca una reunión de relanzamiento del sistema. Se revisan las reglas y se redefinen los compromisos. Si alguien no puede comprometerse, es una conversación de RRHH."),
    ("Si el dueño mismo no usa el dashboard", "El sistema se cae. El equipo lo sabe. Si el dueño no revisa los datos, el equipo entiende que no importa. La adopción empieza por arriba."),
]

for situacion, accion in acciones_resistencia:
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Table Grid'
    c0 = tbl.cell(0, 0)
    c0.width = Inches(2.0)
    set_cell_bg(c0, "1A2B4A")
    c0.paragraphs[0].add_run(situacion).font.color.rgb = DORADO
    c0.paragraphs[0].runs[0].font.bold = True
    c0.paragraphs[0].runs[0].font.size = Pt(10)

    c1 = tbl.cell(0, 1)
    c1.paragraphs[0].add_run(accion).font.size = Pt(10)
    c1.paragraphs[0].paragraph_format.left_indent = Cm(0.3)
    doc.add_paragraph()

doc.add_page_break()

h1(doc, "7. GLOSARIO — TÉRMINOS DEL SISTEMA EN LENGUAJE SIMPLE")
body(doc, "Todos los términos que va a ver en el dashboard, explicados sin tecnicismos:")
doc.add_paragraph()

glosario = [
    ("Lead", "Una persona que mostró interés en su producto o servicio. No es un cliente todavía — es una oportunidad."),
    ("Pipeline", "El recorrido que hace un lead desde que llega hasta que compra (o no compra). Visualizado como un embudo."),
    ("Tasa de conversión", "De cada 100 personas que consultan, ¿cuántas terminan comprando? Si son 15, la tasa es 15%."),
    ("CRM", "La herramienta donde se registra cada lead, cada conversación y cada venta. En su caso: GoHighLevel."),
    ("Tiempo de respuesta", "Cuánto tarda el equipo en responder a un mensaje nuevo. El objetivo es menos de 5 minutos."),
    ("Lead caliente", "Un lead con alta probabilidad de comprar — mostró interés claro, tiene presupuesto y necesita pronto."),
    ("Lead en riesgo", "Un lead que lleva varios días sin actividad y podría perderse si no hay seguimiento inmediato."),
    ("Dashboard", "El tablero visual donde usted ve todos los indicadores del negocio en tiempo real."),
    ("Retainer", "El servicio mensual de León para mantener y optimizar el sistema después de la implementación."),
    ("KPI", "Indicador clave de desempeño. Es el número concreto que le dice si algo está bien o mal."),
    ("Automatización", "Una acción que el sistema hace sola, sin que nadie tenga que ejecutarla manualmente."),
]

tbl_glosario = doc.add_table(rows=len(glosario)+1, cols=2)
tbl_glosario.style = 'Table Grid'
for j, h in enumerate(["TÉRMINO", "SIGNIFICADO SIMPLE"]):
    c = tbl_glosario.cell(0, j)
    set_cell_bg(c, "1A2B4A")
    c.paragraphs[0].add_run(h).font.color.rgb = DORADO
    c.paragraphs[0].runs[0].font.bold = True
    c.paragraphs[0].runs[0].font.size = Pt(10)

for i, (termino, significado) in enumerate(glosario):
    row = tbl_glosario.rows[i+1]
    c0 = row.cells[0]
    c1 = row.cells[1]
    if i % 2 == 0:
        set_cell_bg(c0, "F0F4F8")
        set_cell_bg(c1, "F0F4F8")
    c0.paragraphs[0].add_run(termino).font.size = Pt(10)
    c0.paragraphs[0].runs[0].font.bold = True
    c0.paragraphs[0].runs[0].font.color.rgb = AZUL
    c1.paragraphs[0].add_run(significado).font.size = Pt(10)

doc.save(OUTPUT)
print(f"✓ Creado: {OUTPUT}")
