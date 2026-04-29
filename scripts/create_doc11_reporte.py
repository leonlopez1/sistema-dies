#!/usr/bin/env python3
"""Archivo 11: socializacion_reporte_mensual_template.docx"""
import sys; sys.path.insert(0, '/home/user/sistema-dies/scripts')
from docx_helpers import *

OUTPUT = "/mnt/user-data/outputs/socializacion_reporte_mensual_template.docx"

doc = new_doc()
add_header_footer(doc, "Reporte Mensual")

portada(doc,
    "REPORTE MENSUAL",
    "Estado del Sistema Comercial Operativo",
    "Entregable mensual para el Dueño | Retainer León López")

doc.add_paragraph()
p_datos = doc.add_paragraph()
p_datos.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Info block
tbl_info = doc.add_table(rows=1, cols=3)
tbl_info.alignment = WD_TABLE_ALIGNMENT.CENTER
celdas = [("EMPRESA", "[Nombre]"), ("MES REPORTADO", "[Mes / Año]"), ("PREPARADO POR", "León López")]
for i, (k, v) in enumerate(celdas):
    cell = tbl_info.cell(0, i)
    set_cell_bg(cell, "1A2B4A")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    rk = p.add_run(k + "\n")
    rk.font.size = Pt(9)
    rk.font.color.rgb = DORADO
    rk.font.bold = True
    rv = p.add_run(v)
    rv.font.size = Pt(11)
    rv.font.color.rgb = BLANCO
    rv.font.bold = False
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(8)

doc.add_page_break()

# 1. RESUMEN EJECUTIVO
h1(doc, "1. RESUMEN EJECUTIVO DEL MES")
body(doc, "[Párrafo de 4-6 líneas que describe los resultados del mes en términos de negocio. Incluir: cómo fue el mes en general, el indicador más positivo, el indicador que más requiere atención y la perspectiva para el mes siguiente. Escribir en primera persona: 'Este mes cerramos con...']")
doc.add_paragraph()

section_box(doc, "Titular del mes",
    ["[Una sola frase que resume el mes. Ejemplo: 'Mes de consolidación: el tiempo de respuesta cayó un 40% y la tasa de conversión subió 3 puntos.' O: 'Mes de ajuste: adoptamos el sistema pero el equipo necesita refuerzo en el registro de leads perdidos.']"],
    bg_hex="FFF9E6", title_color=DORADO, icon="→")

doc.add_page_break()

# 2. LOS 5 KPIs
h1(doc, "2. LOS 5 KPIs DEL MES VS. MES ANTERIOR")
body(doc, "Compare el mes reportado contra el mes anterior. Si es el primer mes: compare contra la línea base del diagnóstico.")
doc.add_paragraph()

tbl_kpis = doc.add_table(rows=7, cols=5)
tbl_kpis.style = 'Table Grid'

headers_k = ["KPI", "MES ANTERIOR", "MES ACTUAL", "VARIACIÓN", "ESTADO"]
for j, h in enumerate(headers_k):
    c = tbl_kpis.cell(0, j)
    set_cell_bg(c, "1A2B4A")
    r = c.paragraphs[0].add_run(h)
    r.font.bold = True
    r.font.color.rgb = DORADO
    r.font.size = Pt(10)
    r.alignment = WD_ALIGN_PARAGRAPH.CENTER

kpis_data = [
    ("Tiempo promedio de respuesta", "[X] minutos", "[X] minutos", "[+/- X%]", "[Verde/Amarillo/Rojo]"),
    ("% leads respondidos (mismo día)", "[X%]", "[X%]", "[+/- X pp]", "[Verde/Amarillo/Rojo]"),
    ("Tasa de conversión (lead → venta)", "[X%]", "[X%]", "[+/- X pp]", "[Verde/Amarillo/Rojo]"),
    ("Ventas realizadas (COP)", "$[X]", "$[X]", "[+/- X%]", "[Verde/Amarillo/Rojo]"),
    ("% adopción del CRM por el equipo", "[X%]", "[X%]", "[+/- X pp]", "[Verde/Amarillo/Rojo]"),
    ("Meta del mes vs. resultado real", "$[X] meta", "$[X] real", "[% cumplimiento]", "[Verde/Amarillo/Rojo]"),
]

for i, (kpi, ant, act, var, estado) in enumerate(kpis_data):
    row = tbl_kpis.rows[i+1]
    colores_estado = {"Verde": "E8F8F0", "Amarillo": "FFF9E6", "Rojo": "FDECEA"}

    c0 = row.cells[0]
    c0.paragraphs[0].add_run(kpi).font.size = Pt(10)
    c0.paragraphs[0].runs[0].font.bold = True
    c0.paragraphs[0].runs[0].font.color.rgb = AZUL

    for j, val in enumerate([ant, act, var, estado], 1):
        c = row.cells[j]
        c.paragraphs[0].add_run(val).font.size = Pt(10)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()
body(doc, "Nota: Los valores en verde representan mejora vs. mes anterior. Amarillo = estable. Rojo = requiere atención. Las acciones correctivas están en la sección 5.", italic=True)

doc.add_page_break()

# 3. OPORTUNIDADES RECUPERADAS
h1(doc, "3. OPORTUNIDADES RECUPERADAS ESTE MES")
body(doc, "Estas son las ventas que probablemente NO se habrían cerrado sin el sistema. Se calculan comparando el ritmo actual contra el ritmo del diagnóstico inicial.")
doc.add_paragraph()

tbl_oport = doc.add_table(rows=5, cols=3)
tbl_oport.style = 'Table Grid'
for j, h in enumerate(["SITUACIÓN", "ANTES DEL SISTEMA", "CON EL SISTEMA"]):
    c = tbl_oport.cell(0, j)
    set_cell_bg(c, "1A2B4A")
    c.paragraphs[0].add_run(h).font.color.rgb = DORADO
    c.paragraphs[0].runs[0].font.bold = True
    c.paragraphs[0].runs[0].font.size = Pt(10)

oport_data = [
    ("Leads respondidos en <5 minutos", "[X%]", "[X%] (+[X] pp)"),
    ("Leads con seguimiento programado", "[X]", "[X] (+[X] leads)"),
    ("Valor estimado recuperado en el mes", "$[X] COP", "$[X] COP"),
    ("Leads perdidos por falta de protocolo", "[X]", "[X] (-[X] leads)"),
]
for i, (sit, ant, act) in enumerate(oport_data):
    row = tbl_oport.rows[i+1]
    row.cells[0].paragraphs[0].add_run(sit).font.size = Pt(10)
    row.cells[0].paragraphs[0].runs[0].font.bold = True
    row.cells[1].paragraphs[0].add_run(ant).font.size = Pt(10)
    row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c_act = row.cells[2]
    set_cell_bg(c_act, "E8F8F0")
    c_act.paragraphs[0].add_run(act).font.size = Pt(10)
    c_act.paragraphs[0].runs[0].font.bold = True
    c_act.paragraphs[0].runs[0].font.color.rgb = VERDE
    c_act.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()

# Big highlight
section_box(doc, "VALOR TOTAL RECUPERADO ESTE MES",
    [
        "Estimado de oportunidades recuperadas gracias al sistema: $[XXX.XXX.XXX] COP",
        "",
        "Cómo se calcula: comparamos la tasa de conversión actual ([X%]) vs. la del diagnóstico ([X%]) aplicada al volumen de leads del mes ([X leads]). La diferencia × ticket promedio ($[X] COP) = valor recuperado.",
        "",
        "Acumulado desde el inicio del sistema: $[XXX.XXX.XXX] COP",
    ],
    bg_hex="E8F8F0", title_color=VERDE, icon="$")

doc.add_page_break()

# 4. CUELLOS RESUELTOS
h1(doc, "4. CUELLOS DE BOTELLA RESUELTOS")
body(doc, "Problemas identificados y resueltos durante el mes:")
doc.add_paragraph()

for i in range(3):
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = 'Table Grid'
    c0 = tbl.cell(0, 0)
    c0.width = Inches(0.5)
    set_cell_bg(c0, "27AE60")
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run("✓")
    r0.font.bold = True
    r0.font.size = Pt(14)
    r0.font.color.rgb = BLANCO

    c1 = tbl.cell(0, 1)
    c1.width = Inches(2.5)
    set_cell_bg(c1, "F5F5F5")
    p1 = c1.paragraphs[0]
    p1.paragraph_format.left_indent = Cm(0.3)
    r1 = p1.add_run(f"[Cuello #{i+1} resuelto]")
    r1.font.bold = True
    r1.font.size = Pt(11)
    r1.font.color.rgb = AZUL

    c2 = tbl.cell(0, 2)
    c2.width = Inches(3.5)
    p2 = c2.paragraphs[0]
    p2.paragraph_format.left_indent = Cm(0.3)
    r2 = p2.add_run("[Descripción de cómo se resolvió y qué impacto tuvo en la operación]")
    r2.font.size = Pt(10)
    r2.font.color.rgb = NEGRO
    doc.add_paragraph()

doc.add_paragraph()
h1(doc, "5. ESTADO DE ADOPCIÓN DEL EQUIPO")
body(doc, "Evaluación mensual del uso del sistema por cada miembro del equipo:")
doc.add_paragraph()

semaforo_table(doc, [
    ("[Asesor 1]", "verde", "[Adopción completa. Registra todos sus leads. Tiempo de respuesta promedio: X minutos.]"),
    ("[Asesor 2]", "amarillo", "[Adopción parcial. Registra el 70% de leads. Necesita refuerzo en el seguimiento.]"),
    ("[Asesor 3]", "rojo", "[Adopción baja. Solo registra el 40% de leads. Requiere sesión 1-a-1 en el mes siguiente.]"),
])

section_box(doc, "Índice de adopción del equipo este mes",
    [
        "Adopción promedio del equipo: [X%]",
        "Meta de adopción: 85%",
        "Tendencia: [Subiendo / Estable / Bajando]",
        "Acción del próximo mes: [Qué va a hacer León para mejorar la adopción]",
    ],
    bg_hex="F5F5F5", title_color=AZUL)

doc.add_page_break()

# 6. RECOMENDACIONES
h1(doc, "6. TRES RECOMENDACIONES PARA EL PRÓXIMO MES")
body(doc, "Estas recomendaciones son específicas, accionables y basadas en los datos del mes.")
doc.add_paragraph()

for i in range(3):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Table Grid'
    c0 = tbl.cell(0, 0)
    c0.width = Inches(0.7)
    set_cell_bg(c0, "1A2B4A")
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p0.add_run(f"0{i+1}")
    r0.font.bold = True
    r0.font.size = Pt(20)
    r0.font.color.rgb = DORADO

    c1 = tbl.cell(0, 1)
    pt = c1.paragraphs[0]
    pt.paragraph_format.left_indent = Cm(0.3)
    rt = pt.add_run(f"[Título de la recomendación {i+1}]")
    rt.font.bold = True
    rt.font.size = Pt(12)
    rt.font.color.rgb = AZUL

    pd = c1.add_paragraph()
    pd.paragraph_format.left_indent = Cm(0.5)
    pd.paragraph_format.space_after = Pt(4)
    rd = pd.add_run("[Descripción de la recomendación: qué hacer, por qué hacerlo y qué resultado esperar. Máximo 3 líneas.]")
    rd.font.size = Pt(10)
    rd.font.color.rgb = NEGRO
    doc.add_paragraph()

# 7. ESTADO DEL PLAN
h1(doc, "7. ESTADO DEL PLAN DE TRABAJO")
body(doc, "¿Dónde estamos en el plan de 8 semanas?")
doc.add_paragraph()

tbl_plan = doc.add_table(rows=9, cols=4)
tbl_plan.style = 'Table Grid'
for j, h in enumerate(["SEMANA", "HITO", "ESTADO", "OBSERVACIONES"]):
    c = tbl_plan.cell(0, j)
    set_cell_bg(c, "1A2B4A")
    c.paragraphs[0].add_run(h).font.color.rgb = DORADO
    c.paragraphs[0].runs[0].font.bold = True

semanas_plan = [
    ("1-2", "Diagnóstico y aprobación del informe", "[Completado / En proceso / Pendiente]", ""),
    ("3", "Proceso comercial rediseñado", "[Estado]", ""),
    ("4", "CRM configurado y activo", "[Estado]", ""),
    ("5", "Metas definidas y firmadas", "[Estado]", ""),
    ("6", "Dashboard activo", "[Estado]", ""),
    ("7", "Capacitación del equipo", "[Estado]", ""),
    ("8", "Pruebas y entrega formal", "[Estado]", ""),
    ("4+", "Retainer mensual activo", "[Estado]", ""),
]

for i, (sem, hito, estado, obs) in enumerate(semanas_plan):
    row = tbl_plan.rows[i+1]
    row.cells[0].paragraphs[0].add_run(f"Semana {sem}").font.size = Pt(10)
    row.cells[0].paragraphs[0].runs[0].font.bold = True
    row.cells[1].paragraphs[0].add_run(hito).font.size = Pt(10)
    row.cells[2].paragraphs[0].add_run(estado).font.size = Pt(10)
    row.cells[3].paragraphs[0].add_run(obs if obs else "—").font.size = Pt(10)

doc.add_paragraph()
body(doc, "Próxima entrega: [Fecha del próximo reporte mensual]", italic=True)
doc.add_paragraph()
firma_block(doc, "León López — Sistema Comercial Operativo")

doc.save(OUTPUT)
print(f"✓ Creado: {OUTPUT}")
