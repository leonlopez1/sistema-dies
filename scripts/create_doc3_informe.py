#!/usr/bin/env python3
"""Archivo 3: diagnostico_informe_cliente.docx"""
import sys; sys.path.insert(0, '/home/user/sistema-dies/scripts')
from docx_helpers import *

OUTPUT = "/mnt/user-data/outputs/diagnostico_informe_cliente.docx"

doc = new_doc()
add_header_footer(doc, "Informe de Diagnóstico Comercial")

# PORTADA
portada(doc,
    "INFORME DE DIAGNÓSTICO COMERCIAL",
    "Radiografía del Estado Actual de su Operación",
    "Entregable Semanas 1–2 | Confidencial para la empresa")

doc.add_paragraph()
# Datos empresa box
tbl = doc.add_table(rows=1, cols=2)
tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
datos = [
    ("Empresa:", "[Nombre de la empresa]"),
    ("Sector:", "[Sector]"),
    ("Ciudad:", "[Ciudad]"),
    ("Facturación mensual:", "[Rango COP]"),
    ("Colaboradores:", "[Número total]"),
    ("Fecha del diagnóstico:", "[DD/MM/AAAA]"),
    ("Facilitador:", "León López"),
]
for i, (k, v) in enumerate(datos):
    if i < len(datos):
        row_cells = tbl.rows[0].cells if i == 0 else tbl.add_row().cells
        if i == 0:
            row_cells = tbl.rows[0].cells
        c0 = row_cells[0] if i == 0 else tbl.rows[i].cells[0]
        c1 = row_cells[1] if i == 0 else tbl.rows[i].cells[1]
        # Rewrite using direct cell access
        pass

# Better approach: simple table
doc.tables[-1]._element.getparent().remove(doc.tables[-1]._element)
tbl2 = doc.add_table(rows=len(datos), cols=2)
tbl2.alignment = WD_TABLE_ALIGNMENT.LEFT
tbl2.style = 'Table Grid'
for i, (k, v) in enumerate(datos):
    c0 = tbl2.cell(i, 0)
    c1 = tbl2.cell(i, 1)
    set_cell_bg(c0, "1A2B4A")
    p0 = c0.paragraphs[0]
    r0 = p0.add_run(k)
    r0.font.bold = True
    r0.font.size = Pt(10)
    r0.font.color.rgb = DORADO
    p1 = c1.paragraphs[0]
    r1 = p1.add_run(v)
    r1.font.size = Pt(10)
    r1.font.color.rgb = NEGRO

doc.add_page_break()

# RESUMEN EJECUTIVO
h1(doc, "1. RESUMEN EJECUTIVO")
body(doc, "[Párrafo de 4-6 líneas que resume el estado actual de la empresa. Incluir: qué encontramos, cuál es el principal cuello de botella, cuál es la oportunidad de mejora y qué se recomienda hacer primero. Escribir en primera persona plural: 'Encontramos que...']")
doc.add_paragraph()

section_box(doc, "Hallazgo principal del diagnóstico",
    ["[Describa en 2-3 líneas el problema más crítico identificado. Ejemplo: 'La empresa recibe 35 conversaciones diarias por WhatsApp pero solo tiene protocolo de respuesta para el 40% de ellas. El 60% restante se gestiona sin criterio y sin registro.'"],
    bg_hex="FFF3CD", title_color=DORADO, icon="!")

# ESTADO ACTUAL POR DIMENSIÓN
h1(doc, "2. ESTADO ACTUAL DE LA OPERACIÓN")
body(doc, "A continuación se presenta el diagnóstico por dimensión evaluada. Cada una fue calificada durante la sesión con el equipo directivo.", italic=True)
doc.add_paragraph()

semaforo_table(doc, [
    ("Flujo de Leads", "amarillo", "[Descripción del estado actual del flujo de leads. Ej: Los leads llegan pero no hay protocolo claro de atención]"),
    ("Proceso de Cierre", "rojo", "[Descripción. Ej: No existe un proceso estandarizado. Cada asesor cierra como puede]"),
    ("Estructura del Equipo", "amarillo", "[Descripción. Ej: El equipo existe pero las metas no están definidas individualmente]"),
    ("Herramientas y Tecnología", "rojo", "[Descripción. Ej: Solo se usa WhatsApp. No hay CRM ni registro centralizado]"),
    ("Visibilidad Gerencial", "rojo", "[Descripción. Ej: El dueño solo sabe cómo va el negocio cuando pregunta directamente]"),
])

# COSTO DE OPORTUNIDAD
h1(doc, "3. COSTO DE OPORTUNIDAD")

section_box(doc, "Lo que está costando no tener el sistema",
    [
        "Conversaciones perdidas por mes (estimado): [X] conversaciones",
        "Valor perdido por tiempos de respuesta lentos: $[XXX.XXX.XXX] COP/mes",
        "Valor perdido por falta de seguimiento: $[XXX.XXX.XXX] COP/mes",
        "Valor perdido por ceguera gerencial: $[XXX.XXX.XXX] COP/mes",
        "",
        "TOTAL PÉRDIDA MENSUAL ESTIMADA: $[XXX.XXX.XXX] COP",
        "POTENCIAL DE RECUPERACIÓN (primeros 3 meses): $[XXX.XXX.XXX] COP",
    ],
    bg_hex="FDECEA", title_color=ROJO, icon="$")

body(doc, "Estos valores fueron calculados con base en los datos proporcionados durante el diagnóstico y proyecciones conservadoras del mercado. El detalle del cálculo está disponible en la Calculadora de Oportunidad entregada junto con este informe.", italic=True)

# CUELLOS DE BOTELLA
h1(doc, "4. LOS 3 CUELLOS DE BOTELLA CRÍTICOS")
body(doc, "Estos son los tres puntos donde la operación pierde más dinero y energía:")
doc.add_paragraph()

for i, (titulo, desc) in enumerate([
    ("CUELLO #1: [Nombre del cuello de botella]",
     "[Descripción detallada. Qué está pasando, por qué está pasando y qué consecuencia tiene. Ej: 'Los leads que llegan por WhatsApp no tienen un protocolo de respuesta. El tiempo promedio de atención es de 47 minutos. Esto reduce la tasa de conversión en aproximadamente un 40%.']"),
    ("CUELLO #2: [Nombre del cuello de botella]",
     "[Descripción detallada del segundo problema crítico.]"),
    ("CUELLO #3: [Nombre del cuello de botella]",
     "[Descripción detallada del tercer problema crítico.]"),
], 1):
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, "1A2B4A")
    pt = cell.paragraphs[0]
    pt.paragraph_format.space_before = Pt(6)
    pt.paragraph_format.space_after = Pt(2)
    pt.paragraph_format.left_indent = Cm(0.4)
    rt = pt.add_run(f"#{i} " + titulo.split(": ")[1] if ": " in titulo else titulo)
    rt.font.bold = True
    rt.font.size = Pt(12)
    rt.font.color.rgb = DORADO

    pd = cell.add_paragraph()
    pd.paragraph_format.left_indent = Cm(0.4)
    pd.paragraph_format.space_before = Pt(4)
    pd.paragraph_format.space_after = Pt(8)
    rd = pd.add_run(desc)
    rd.font.size = Pt(10)
    rd.font.color.rgb = BLANCO
    rd.font.italic = True
    doc.add_paragraph()

# RECOMENDACIÓN
h1(doc, "5. RECOMENDACIÓN: QUÉ HACER Y EN QUÉ ORDEN")
body(doc, "Con base en el diagnóstico, recomendamos implementar el Sistema Comercial Operativo en el siguiente orden:")
doc.add_paragraph()

pasos = [
    ("Semana 3", "Rediseñar el proceso comercial", "Mapear el proceso actual y diseñar el proceso nuevo con roles, tiempos y protocolos claros."),
    ("Semana 4", "Configurar el CRM (GoHighLevel)", "Activar la herramienta de gestión con el pipeline definido e integrar WhatsApp."),
    ("Semana 5", "Definir metas por persona", "Bajar la meta global a metas individuales claras con criterios de seguimiento."),
    ("Semana 6", "Activar el dashboard de visibilidad", "El dueño tendrá acceso en tiempo real a los 5 indicadores clave."),
    ("Semana 7", "Capacitar al equipo", "Sesión de 90 minutos. El equipo aprende el sistema y firma su compromiso de adopción."),
    ("Semana 8", "Pruebas, ajuste y entrega", "Verificar que el sistema funciona correctamente y hacer ajustes finales."),
]

for semana, titulo, desc in pasos:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.left_indent = Cm(0.5)
    rs = p.add_run(f"[{semana}] ")
    rs.font.bold = True
    rs.font.color.rgb = DORADO
    rs.font.size = Pt(11)
    rt = p.add_run(titulo + ": ")
    rt.font.bold = True
    rt.font.color.rgb = AZUL
    rt.font.size = Pt(11)
    rd = p.add_run(desc)
    rd.font.size = Pt(11)
    rd.font.color.rgb = NEGRO

# PRÓXIMO PASO
doc.add_paragraph()
h1(doc, "6. PROPUESTA DE PRÓXIMO PASO")

section_box(doc, "Lo que sucede si seguimos adelante",
    [
        "Inversión total del sistema: $12.000.000 COP",
        "  — $5.000.000 COP al inicio (diagnóstico + diseño)",
        "  — $3.500.000 COP semana 4 (configuración e implementación)",
        "  — $3.500.000 COP semana 7 (capacitación y entrega)",
        "",
        "Retainer mensual (mes 4 en adelante): $3.000.000 COP/mes",
        "  — Revisión semanal de métricas",
        "  — Reporte mensual al dueño",
        "  — Ajustes y optimización continua del sistema",
        "",
        "POTENCIAL DE RETORNO: $[XXX.XXX.XXX] COP en los primeros 3 meses",
    ],
    bg_hex="E8F8F0", title_color=VERDE, icon="→")

body(doc, "La decisión de avanzar está en sus manos. Lo que sí puedo garantizarle es que con este diagnóstico ya tiene claridad de dónde está perdiendo dinero. El siguiente paso es decidir si quiere seguir perdiéndolo o empezar a recuperarlo.", italic=True)

doc.add_paragraph()
firma_block(doc, "Firma del cliente — autorización para continuar")
firma_block(doc, "León López — Sistema Comercial Operativo")

doc.save(OUTPUT)
print(f"✓ Creado: {OUTPUT}")
