#!/usr/bin/env python3
"""Archivo 5: implementacion_acuerdo_cliente.docx"""
import sys; sys.path.insert(0, '/home/user/sistema-dies/scripts')
from docx_helpers import *

OUTPUT = "/mnt/user-data/outputs/implementacion_acuerdo_cliente.docx"

doc = new_doc()
add_header_footer(doc, "Acuerdo de Servicio")

portada(doc,
    "ACUERDO DE SERVICIO PROFESIONAL",
    "Sistema Comercial Operativo — Implementación",
    "Documento confidencial | Válido por 8 semanas")

doc.add_page_break()

# Partes
h1(doc, "IDENTIFICACIÓN DE LAS PARTES")
tbl = doc.add_table(rows=6, cols=2)
tbl.style = 'Table Grid'
partes = [
    ("PROVEEDOR DEL SERVICIO", ""),
    ("Nombre:", "León López"),
    ("Servicio:", "Sistema Comercial Operativo"),
    ("CLIENTE", ""),
    ("Empresa:", "[Nombre de la empresa]"),
    ("Representante legal:", "[Nombre del dueño o apoderado]"),
]
for i, (k, v) in enumerate(partes):
    c0 = tbl.cell(i, 0)
    c1 = tbl.cell(i, 1)
    if v == "":
        tbl.cell(i, 0).merge(tbl.cell(i, 1))
        set_cell_bg(c0, "1A2B4A")
        r = c0.paragraphs[0].add_run(k)
        r.font.bold = True
        r.font.color.rgb = DORADO
        r.font.size = Pt(11)
    else:
        set_cell_bg(c0, "F5F5F5")
        c0.paragraphs[0].add_run(k).font.size = Pt(10)
        c0.paragraphs[0].runs[0].font.bold = True
        c1.paragraphs[0].add_run(v).font.size = Pt(10)

doc.add_paragraph()
p_fecha = doc.add_paragraph()
r_f = p_fecha.add_run("Fecha de inicio del acuerdo: ___________________    Ciudad: ___________________")
r_f.font.size = Pt(11)
r_f.font.color.rgb = GRIS_TEXTO

# ALCANCE
doc.add_page_break()
h1(doc, "1. ALCANCE DEL SERVICIO")
body(doc, "León López (en adelante 'el proveedor') se compromete a diseñar, instalar y poner en operación el Sistema Comercial Operativo en la empresa del cliente durante un período máximo de 8 semanas, según la siguiente estructura:")
doc.add_paragraph()

semanas_scope = [
    ("Semanas 1–2", "Diagnóstico Comercial",
     "Sesión de radiografía (60–90 min), análisis de datos, calculadora de oportunidad, informe escrito con hallazgos y costo de oportunidad en COP."),
    ("Semana 3", "Rediseño del Proceso Comercial",
     "Mapeo AS-IS y diseño TO-BE del proceso comercial. Roles, tiempos y protocolos definidos y documentados."),
    ("Semana 4", "Configuración del CRM",
     "Activación y configuración de GoHighLevel: pipeline, automatizaciones, integración de WhatsApp y campos personalizados."),
    ("Semana 5", "Estructura de Metas",
     "Distribución de metas por canal y por asesor. Acuerdos de Desempeño individuales para el equipo."),
    ("Semana 6", "Dashboard de Visibilidad",
     "Implementación del tablero de indicadores en Lovable con las 4 pantallas: tiempos, embudo, meta vs. real y oportunidades en riesgo."),
    ("Semana 7", "Capacitación del Equipo",
     "Sesión presencial o virtual de 90 minutos con el equipo completo. Material de referencia entregado."),
    ("Semana 8", "Pruebas, Ajuste y Entrega",
     "Verificación del sistema, correcciones finales y declaración formal de implementación exitosa."),
]

for periodo, titulo, desc in semanas_scope:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    rp = p.add_run(f"{periodo} — ")
    rp.font.bold = True
    rp.font.color.rgb = DORADO
    rt = p.add_run(titulo + ": ")
    rt.font.bold = True
    rt.font.color.rgb = AZUL
    rd = p.add_run(desc)
    rd.font.size = Pt(11)
    rd.font.color.rgb = NEGRO

# ENTREGABLES
doc.add_paragraph()
h1(doc, "2. ENTREGABLES")
body(doc, "Al finalizar el proceso, el cliente recibirá los siguientes entregables en formato digital (Google Drive):")
doc.add_paragraph()

entregables = [
    "Diagnóstico Comercial completo (cuestionario + calculadora + informe)",
    "Proceso comercial documentado (AS-IS y TO-BE en formato visual)",
    "CRM configurado y operativo con automatizaciones activas",
    "Acuerdos de Desempeño individuales del equipo",
    "Dashboard de visibilidad en Lovable con acceso del dueño",
    "Playbook operativo (guía para el equipo y para futuros asesores)",
    "Material de capacitación (tarjeta de proceso + FAQ del equipo)",
    "Reporte de implementación (estado del sistema al finalizar semana 8)",
]
for e in entregables:
    bullet(doc, e)

# PAGOS
doc.add_paragraph()
h1(doc, "3. CONDICIONES DE PAGO")
doc.add_paragraph()

tbl_pagos = doc.add_table(rows=5, cols=4)
tbl_pagos.style = 'Table Grid'
headers_pagos = ["PAGO", "MOMENTO", "VALOR", "CONCEPTO"]
for i, h in enumerate(headers_pagos):
    c = tbl_pagos.cell(0, i)
    set_cell_bg(c, "1A2B4A")
    r = c.paragraphs[0].add_run(h)
    r.font.bold = True
    r.font.color.rgb = DORADO
    r.font.size = Pt(10)

pagos = [
    ("Pago 1", "Al firmar el acuerdo", "$5.000.000 COP", "Diagnóstico + diseño del sistema"),
    ("Pago 2", "Al inicio de la semana 4", "$3.500.000 COP", "Configuración CRM + dashboard"),
    ("Pago 3", "Al inicio de la semana 7", "$3.500.000 COP", "Capacitación + entrega del sistema"),
    ("TOTAL", "—", "$12.000.000 COP", "Implementación completa (8 semanas)"),
]
for i, (p1, p2, p3, p4) in enumerate(pagos):
    row = tbl_pagos.rows[i+1]
    if p1 == "TOTAL":
        for j, val in enumerate([p1, p2, p3, p4]):
            set_cell_bg(row.cells[j], "C9973A")
            r = row.cells[j].paragraphs[0].add_run(val)
            r.font.bold = True
            r.font.color.rgb = BLANCO
            r.font.size = Pt(10)
    else:
        for j, val in enumerate([p1, p2, p3, p4]):
            row.cells[j].paragraphs[0].add_run(val).font.size = Pt(10)

doc.add_paragraph()
body(doc, "Retainer mensual (mes 4 en adelante): $3.000.000 COP/mes. Incluye: 4 revisiones semanales de métricas, reporte mensual al dueño y ajustes al sistema.")
body(doc, "El retainer se activa automáticamente al mes 4 salvo que el cliente notifique su intención de no continuar con 15 días de anticipación.")

# COMPROMISOS DEL CLIENTE
doc.add_paragraph()
h1(doc, "4. COMPROMISOS DEL CLIENTE")
body(doc, "Para que el sistema funcione, el cliente debe garantizar:")
doc.add_paragraph()

compromisos = [
    ("Disponibilidad del dueño", "Mínimo 2 horas por semana durante las 8 semanas de implementación para sesiones de trabajo con León."),
    ("Disponibilidad del equipo", "El equipo comercial debe participar en las sesiones de capacitación sin excepción. Si alguien no puede asistir, se agenda sesión de recuperación."),
    ("Acceso a herramientas", "El cliente provee acceso a su cuenta de WhatsApp Business, correo corporativo y cualquier herramienta existente que deba integrarse."),
    ("Pago puntual de GoHighLevel", "La suscripción a GoHighLevel corre por cuenta del cliente (USD $97/mes aproximadamente). León no asume este costo."),
    ("Un contacto interno", "El cliente designa una persona de su empresa que será el punto de contacto durante la implementación. Esta persona tiene disponibilidad durante horario laboral."),
    ("Adopción real del sistema", "El dueño se compromete a exigir el uso del sistema a su equipo. Si el equipo no adopta el sistema, los resultados no pueden garantizarse."),
]

for titulo, desc in compromisos:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.left_indent = Cm(0.5)
    rb = p.add_run(f"• {titulo}: ")
    rb.font.bold = True
    rb.font.color.rgb = AZUL
    rb.font.size = Pt(11)
    rd = p.add_run(desc)
    rd.font.size = Pt(11)
    rd.font.color.rgb = NEGRO

# NO INCLUYE
doc.add_paragraph()
h1(doc, "5. QUÉ NO INCLUYE ESTE SERVICIO")
no_incluye = [
    "El costo mensual de GoHighLevel (lo paga directamente el cliente)",
    "Desarrollo de software personalizado fuera del alcance definido",
    "Gestión de campañas de marketing pagadas (pauta en Meta, Google, etc.)",
    "Contratación o despido de personal del área comercial",
    "Redacción de contratos legales con clientes de la empresa",
    "Soporte técnico directo a los asesores después de finalizada la semana 8 (eso corre por el retainer)",
    "Resultados garantizados en ventas (los resultados dependen de la adopción del equipo)",
]
for item in no_incluye:
    bullet(doc, item)

# CLÁUSULA RETAINER
doc.add_paragraph()
h1(doc, "6. CLÁUSULA DE CONTINUIDAD — RETAINER MENSUAL")
body(doc, "A partir del mes 4, el servicio continúa bajo modalidad de retainer mensual de $3.000.000 COP. El retainer incluye:")
for item in [
    "4 reuniones semanales de revisión de métricas (30 minutos cada una)",
    "1 reporte mensual al dueño con los KPIs del mes y recomendaciones",
    "Ajustes y optimización del sistema según los resultados observados",
    "Acceso directo a León por WhatsApp en horario laboral para consultas urgentes",
]:
    bullet(doc, item)

doc.add_paragraph()
body(doc, "Para cancelar el retainer, el cliente debe notificar con 15 días de anticipación. No hay penalidades por cancelación.")

# FIRMAS
doc.add_page_break()
h1(doc, "FIRMAS Y ACEPTACIÓN")
body(doc, "Ambas partes declaran haber leído, entendido y aceptado los términos de este acuerdo.")
doc.add_paragraph()
doc.add_paragraph()

for titulo in ["León López — Proveedor del Servicio", "Representante del cliente — [Nombre]"]:
    tbl_firma = doc.add_table(rows=3, cols=2)
    tbl_firma.style = 'Table Grid'
    tbl_firma.cell(0, 0).merge(tbl_firma.cell(0, 1))
    set_cell_bg(tbl_firma.cell(0, 0), "1A2B4A")
    tbl_firma.cell(0, 0).paragraphs[0].add_run(titulo).font.color.rgb = DORADO
    tbl_firma.cell(0, 0).paragraphs[0].runs[0].font.bold = True

    tbl_firma.cell(1, 0).paragraphs[0].add_run("Firma: ____________________________").font.size = Pt(11)
    tbl_firma.cell(1, 1).paragraphs[0].add_run("Fecha: ____________________________").font.size = Pt(11)
    tbl_firma.cell(2, 0).paragraphs[0].add_run("Cédula / NIT: _____________________").font.size = Pt(11)
    tbl_firma.cell(2, 1).paragraphs[0].add_run("Cargo: ____________________________").font.size = Pt(11)
    doc.add_paragraph()

doc.save(OUTPUT)
print(f"✓ Creado: {OUTPUT}")
