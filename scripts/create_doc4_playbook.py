#!/usr/bin/env python3
"""Archivo 4: implementacion_playbook_operativo.docx"""
import sys; sys.path.insert(0, '/home/user/sistema-dies/scripts')
from docx_helpers import *

OUTPUT = "/mnt/user-data/outputs/implementacion_playbook_operativo.docx"

doc = new_doc()
add_header_footer(doc, "Playbook de Implementación")

portada(doc,
    "PLAYBOOK DE IMPLEMENTACIÓN",
    "Guía Operativa — Semanas 3 a 8",
    "Pilar 2: Implementación | Sistema Comercial Operativo")

doc.add_page_break()

# Intro
h1(doc, "¿PARA QUÉ SIRVE ESTE DOCUMENTO?")
body(doc, "Este playbook es la guía de campo para implementar el Sistema Comercial Operativo. Está diseñado para que cualquier implementador entrenado pueda ejecutarlo con el mismo estándar, en el mismo orden y con los mismos resultados.")
body(doc, "Cada semana tiene objetivos claros, entregables concretos y criterios para saber si se logró lo que debía lograrse. No improvise. El orden importa.")
doc.add_paragraph()

section_box(doc, "Antes de empezar — verificar",
    [
        "☑  El diagnóstico fue completado y el informe fue entregado al cliente",
        "☑  El cliente firmó la autorización para continuar",
        "☑  Se recibió el primer pago ($5.000.000 COP)",
        "☑  El cliente tiene acceso a GoHighLevel o está dispuesto a suscribirse",
        "☑  El cliente asignó un contacto interno de coordinación",
    ],
    bg_hex="FFF9E6", title_color=DORADO, icon="✓")

# ─── SEMANA 3 ─────────────────────────────────────────────────────────────────
doc.add_page_break()
h1(doc, "SEMANA 3 — REDISEÑO DEL PROCESO COMERCIAL")
body(doc, "Objetivo: Que exista UN proceso comercial documentado, entendido y aceptado por el dueño antes de terminar la semana.")
body(doc, "Duración estimada de las sesiones: 2 sesiones de 90 minutos cada una.")
doc.add_paragraph()

h2(doc, "3.1 Mapeo del proceso actual (AS-IS)")
body(doc, "En la primera sesión con el cliente, documente el proceso actual exactamente como ocurre hoy — no como debería ocurrir.")
doc.add_paragraph()

add_checklist(doc, [
    "Preguntar: '¿Qué pasa desde que llega un lead hasta que paga?' — Dejar que el cliente lo cuente",
    "Identificar cuántos pasos tiene el proceso (máximo 8 etapas)",
    "Anotar quién es el responsable de cada etapa",
    "Identificar dónde hay brechas: etapas sin dueño, sin tiempo definido o sin protocolo",
    "Registrar cuánto tiempo toma cada etapa en promedio",
], title="Checklist AS-IS:")
doc.add_paragraph()

h3(doc, "Etapas estándar del proceso AS-IS (sector servicios):")
etapas_as_is = [
    ("Llegada del lead", "WhatsApp / Instagram / referido / presencial"),
    ("Primera respuesta", "¿Quién? ¿En cuánto tiempo? ¿Con qué mensaje?"),
    ("Calificación", "¿Alguien pregunta qué necesita el cliente?"),
    ("Propuesta / cotización", "¿Quién la hace? ¿Cuánto tarda?"),
    ("Seguimiento", "¿Alguien vuelve a contactar? ¿Cuántas veces?"),
    ("Cierre", "¿Cómo se formaliza la venta? ¿Hay contrato / recibo?"),
    ("Entrega / servicio", "¿Cuándo se considera cerrado el ciclo?"),
]
for etapa, preguntas in etapas_as_is:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(4)
    rb = p.add_run(f"• {etapa}: ")
    rb.font.bold = True
    rb.font.color.rgb = AZUL
    rb.font.size = Pt(11)
    r = p.add_run(preguntas)
    r.font.size = Pt(11)
    r.font.color.rgb = NEGRO

doc.add_paragraph()
h2(doc, "3.2 Diseño del proceso nuevo (TO-BE)")
body(doc, "En la segunda sesión, diseñe el proceso optimizado. El proceso nuevo debe resolver específicamente los cuellos de botella identificados en el diagnóstico.")
doc.add_paragraph()

h3(doc, "Etapas estándar del proceso TO-BE:")
etapas_to_be = [
    ("Lead nuevo", "Entra al CRM automáticamente. Asignado a un asesor en menos de 2 minutos."),
    ("Primer contacto", "El asesor responde en máximo 5 minutos con mensaje de bienvenida estandarizado."),
    ("Calificación activa", "El asesor hace las 3 preguntas de calificación en los primeros 10 minutos."),
    ("Propuesta", "Cotización enviada en menos de 24 horas con formato estándar aprobado."),
    ("Seguimiento D1/D3/D7", "Sistema recuerda al asesor. Mensajes pre-escritos por etapa."),
    ("Cierre", "El asesor registra en el CRM: ganado o perdido (con razón obligatoria)."),
    ("Post-venta", "Mensaje de bienvenida automático al cliente que compra."),
]
for etapa, protocolo in etapas_to_be:
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = 'Table Grid'
    c0 = tbl.cell(0, 0)
    c0.width = Inches(1.8)
    set_cell_bg(c0, "1A2B4A")
    p0 = c0.paragraphs[0]
    p0.paragraph_format.left_indent = Cm(0.2)
    r0 = p0.add_run(etapa)
    r0.font.bold = True
    r0.font.size = Pt(10)
    r0.font.color.rgb = DORADO

    c1 = tbl.cell(0, 1)
    c1.width = Inches(4.7)
    p1 = c1.paragraphs[0]
    p1.paragraph_format.left_indent = Cm(0.2)
    r1 = p1.add_run(protocolo)
    r1.font.size = Pt(10)
    r1.font.color.rgb = NEGRO

doc.add_paragraph()
h3(doc, "Roles y responsabilidades por etapa:")
body(doc, "Para cada etapa del proceso TO-BE, defina:")
for item in ["Quién es el responsable primario", "Quién es el backup si el primario no está", "Tiempo máximo para ejecutar la etapa", "Qué registrar en el CRM al completarla"]:
    bullet(doc, item)

doc.add_paragraph()
add_checklist(doc, [
    "El proceso nuevo tiene máximo 7 etapas",
    "Cada etapa tiene un responsable nombrado",
    "Cada etapa tiene un tiempo máximo definido",
    "El dueño revisó y aprobó el proceso nuevo",
    "El proceso fue documentado en 1 hoja visual (para el playbook del equipo)",
], title="Checklist de validación del proceso TO-BE:")

# ─── SEMANA 4 ─────────────────────────────────────────────────────────────────
doc.add_page_break()
h1(doc, "SEMANA 4 — CONFIGURACIÓN DEL CRM (GoHighLevel)")
body(doc, "Objetivo: El CRM está activo, el equipo tiene acceso y los primeros leads están siendo registrados antes de que termine la semana.")
doc.add_paragraph()

h2(doc, "4.1 Checklist de configuración de la subcuenta")
add_checklist(doc, [
    "Crear subcuenta en GoHighLevel con nombre y datos de la empresa",
    "Configurar los datos del negocio: logo, colores, dirección, zona horaria",
    "Crear usuarios para cada miembro del equipo comercial con permisos correctos",
    "Asignar roles: Admin (León) / Usuario (asesores) / Solo lectura (dueño)",
    "Configurar la integración de WhatsApp Business (número verificado)",
    "Activar la integración de correo si aplica",
    "Configurar las notificaciones del sistema",
])

doc.add_paragraph()
h2(doc, "4.2 Configuración del Pipeline")
body(doc, "Etapas estándar recomendadas del pipeline (en orden):")
doc.add_paragraph()

pipeline_stages = [
    ("1", "Lead Nuevo Sin Contacto", "Llega automáticamente. Sin acción del asesor aún.", "Rojo"),
    ("2", "En Conversación Activa", "El asesor hizo el primer contacto y hay respuesta del lead.", "Naranja"),
    ("3", "Propuesta / Cotización Enviada", "Se envió precio o propuesta formal al lead.", "Amarillo"),
    ("4", "Seguimiento Programado", "El lead no cerró pero hay interés. Se programó seguimiento.", "Azul"),
    ("5", "Cerrado — Ganado", "El lead compró. Se registra el valor de la venta.", "Verde"),
    ("6", "Cerrado — Perdido", "El lead no compró. OBLIGATORIO registrar la razón.", "Gris"),
]

for num, nombre, desc, color in pipeline_stages:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(6)
    rn = p.add_run(f"Etapa {num}: ")
    rn.font.bold = True
    rn.font.color.rgb = DORADO
    rn.font.size = Pt(11)
    rn2 = p.add_run(nombre + " — ")
    rn2.font.bold = True
    rn2.font.color.rgb = AZUL
    rn2.font.size = Pt(11)
    rd = p.add_run(desc)
    rd.font.size = Pt(11)
    rd.font.color.rgb = NEGRO

doc.add_paragraph()
h2(doc, "4.3 Campos personalizados obligatorios")
campos = [
    "Canal de origen (WhatsApp / Instagram / Referido / Presencial / Web)",
    "Producto o servicio de interés",
    "Presupuesto estimado del cliente",
    "Razón de pérdida (solo para etapa Cerrado-Perdido)",
    "Asesor responsable",
    "Fecha de primer contacto",
    "Fecha de cierre (real o estimada)",
]
for c in campos:
    bullet(doc, c)

doc.add_paragraph()
h2(doc, "4.4 Automatizaciones básicas a configurar")
autos = [
    ("Bienvenida automática", "Cuando un lead llega a 'Lead Nuevo': envía mensaje de bienvenida en máximo 1 minuto."),
    ("Alerta de tiempo de respuesta", "Si un lead lleva más de 10 minutos en etapa 1 sin acción: notificación al asesor."),
    ("Recordatorio de seguimiento", "Lead en etapa 4 por más de 2 días: recordatorio automático al asesor."),
    ("Alerta al dueño", "Lead en etapa 4 por más de 7 días sin movimiento: alerta al gerente."),
    ("Mensaje post-venta", "Cuando un lead pasa a Cerrado-Ganado: mensaje de bienvenida automático al nuevo cliente."),
]
for titulo, desc in autos:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(5)
    rb = p.add_run(f"• {titulo}: ")
    rb.font.bold = True
    rb.font.color.rgb = AZUL
    rb.font.size = Pt(11)
    rd = p.add_run(desc)
    rd.font.size = Pt(11)
    rd.font.color.rgb = NEGRO

doc.add_paragraph()
section_box(doc, "Integración con Dashboard Lovable — Endpoints API",
    [
        "Endpoint de leads activos: GET /api/v1/contacts?pipeline_stage=active",
        "Endpoint de conversiones: GET /api/v1/opportunities?status=won",
        "Endpoint de tiempo de respuesta: GET /api/v1/conversations/response_time",
        "Endpoint de pipeline completo: GET /api/v1/opportunities",
        "Autenticación: Bearer Token (configurar en el dashboard de Lovable)",
        "Frecuencia de actualización recomendada: cada 15 minutos",
    ],
    bg_hex="E8EAF6", title_color=AZUL, icon="⚡")

# ─── SEMANA 5 ─────────────────────────────────────────────────────────────────
doc.add_page_break()
h1(doc, "SEMANA 5 — ESTRUCTURA DE METAS")
body(doc, "Objetivo: Cada miembro del equipo comercial sabe exactamente qué debe vender este mes, en pesos, y cómo se medirá.")
doc.add_paragraph()

h2(doc, "5.1 Fórmula para bajar metas")
body(doc, "Use esta fórmula para distribuir la meta mensual global de la empresa:")
doc.add_paragraph()

section_box(doc, "Modelo de distribución de metas",
    [
        "Meta global del mes: $[X] COP",
        "",
        "División por canal:",
        "  → Canal marketing (leads inbound): 60% de la meta = $[X * 0.60] COP",
        "  → Equipo presencial / outbound: 40% de la meta = $[X * 0.40] COP",
        "",
        "División por asesor (equipo presencial):",
        "  → Meta individual = (Meta total equipo) ÷ (Número de asesores activos)",
        "  → Ejemplo: $30M ÷ 3 asesores = $10M COP por asesor",
        "",
        "Nota: Ajuste según historial individual si hay diferencias significativas de rendimiento.",
    ],
    bg_hex="F5F5F5", title_color=AZUL)

h2(doc, "5.2 Cómo comunicar las metas al equipo")
body(doc, "Reglas para la comunicación de metas:")
for item in [
    "Nunca anuncie una meta sin explicar el razonamiento detrás del número",
    "Use lenguaje de logro, no de presión: 'Si llegamos a X, todos ganamos Y'",
    "Muestre cómo la meta individual contribuye a la meta de la empresa",
    "Defina qué pasa si se supera la meta (incentivo) y qué pasa si no se llega (conversación, no castigo)",
    "Documente todo en un Acuerdo de Desempeño firmado por el asesor",
]:
    bullet(doc, item)

doc.add_paragraph()
h2(doc, "5.3 Template — Acuerdo de Desempeño")
body(doc, "[Este template se diligencia para cada asesor comercial al inicio del mes o al inicio del sistema]")
doc.add_paragraph()

tbl_acuerdo = doc.add_table(rows=9, cols=2)
tbl_acuerdo.style = 'Table Grid'
acuerdo_campos = [
    ("Nombre del asesor:", "[Nombre completo]"),
    ("Cargo:", "[Cargo en la empresa]"),
    ("Período:", "[Mes y año]"),
    ("Meta individual del mes:", "$[XXX.XXX.XXX] COP"),
    ("Meta en unidades (si aplica):", "[X] ventas / citas / servicios"),
    ("Indicadores de seguimiento:", "Tiempo de respuesta / Leads registrados / Seguimientos realizados"),
    ("Incentivo por cumplimiento al 100%:", "$[XXX.XXX] COP / [Beneficio]"),
    ("Incentivo por superar meta (>120%):", "$[XXX.XXX] COP / [Beneficio adicional]"),
    ("Firma:", "________________________________    Fecha: __________"),
]
for i, (k, v) in enumerate(acuerdo_campos):
    c0 = tbl_acuerdo.cell(i, 0)
    c1 = tbl_acuerdo.cell(i, 1)
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

# ─── SEMANA 6 ─────────────────────────────────────────────────────────────────
doc.add_page_break()
h1(doc, "SEMANA 6 — DASHBOARD LOVABLE")
body(doc, "Objetivo: El dueño puede ver el estado de su operación en tiempo real desde su celular o computador, sin necesitar preguntar a nadie.")
doc.add_paragraph()

h2(doc, "6.1 Las 4 Pantallas del Dashboard")
pantallas = [
    ("PANTALLA 1: Tiempos de Respuesta",
     ["Tiempo promedio de respuesta hoy (en minutos)", "Tiempo promedio de respuesta esta semana", "% de leads respondidos en menos de 5 minutos", "Alertas: leads esperando respuesta hace más de 15 minutos", "Semáforo: Verde (<5 min) / Amarillo (5-15 min) / Rojo (>15 min)"]),
    ("PANTALLA 2: Embudo de Conversión",
     ["Total de leads recibidos esta semana", "Leads en cada etapa del pipeline (barras visuales)", "Tasa de conversión actual vs. meta", "Leads en riesgo (más de 5 días sin movimiento)", "Comparativo semana actual vs. semana anterior"]),
    ("PANTALLA 3: Meta vs. Real",
     ["Meta del mes (en COP y en unidades)", "Ventas realizadas al día de hoy (COP y unidades)", "% de avance hacia la meta", "Proyección: ¿van a llegar al cierre del mes?", "Desglose por asesor: quién va bien y quién necesita apoyo"]),
    ("PANTALLA 4: Oportunidades en Riesgo",
     ["Leads en etapa 'Seguimiento' por más de 7 días", "Leads en etapa 1 sin primer contacto por más de 30 minutos", "Leads marcados como 'calientes' sin actividad reciente", "Valor total en riesgo (suma de tickets de leads en riesgo)", "Botón de acción: asignar / escalar / archivar"]),
]

for pantalla, variables in pantallas:
    h3(doc, pantalla)
    for v in variables:
        bullet(doc, v)
    doc.add_paragraph()

h2(doc, "6.2 Reglas de Alertas")
alertas = [
    ("Alerta inmediata (al asesor)", "Lead sin respuesta por más de 10 minutos"),
    ("Alerta urgente (al asesor + supervisor)", "Lead sin respuesta por más de 30 minutos"),
    ("Alerta gerencial (al dueño)", "Lead caliente sin actividad por más de 24 horas"),
    ("Alerta de meta (al dueño)", "Proyección de fin de mes por debajo del 80% de la meta"),
    ("Reporte automático", "Resumen diario a las 6:00 PM para el dueño (WhatsApp o email)"),
]
tbl_alertas = doc.add_table(rows=len(alertas)+1, cols=2)
tbl_alertas.style = 'Table Grid'
h_cells = tbl_alertas.rows[0].cells
set_cell_bg(h_cells[0], "1A2B4A"); h_cells[0].paragraphs[0].add_run("TIPO DE ALERTA").font.bold = True; h_cells[0].paragraphs[0].runs[0].font.color.rgb = DORADO
set_cell_bg(h_cells[1], "1A2B4A"); h_cells[1].paragraphs[0].add_run("CONDICIÓN").font.bold = True; h_cells[1].paragraphs[0].runs[0].font.color.rgb = DORADO
for i, (tipo, condicion) in enumerate(alertas):
    row = tbl_alertas.rows[i+1]
    row.cells[0].paragraphs[0].add_run(tipo).font.size = Pt(10)
    row.cells[1].paragraphs[0].add_run(condicion).font.size = Pt(10)

# ─── SEMANA 7 ─────────────────────────────────────────────────────────────────
doc.add_page_break()
h1(doc, "SEMANA 7 — CAPACITACIÓN DEL EQUIPO")
body(doc, "Objetivo: El equipo completo sabe usar el sistema, entiende las reglas y firma su compromiso de adopción.")
doc.add_paragraph()

h2(doc, "7.1 Agenda de Capacitación (90 minutos)")
agenda = [
    ("0–10 min", "Bienvenida y contexto", "El dueño explica por qué se implementa el sistema. León facilita."),
    ("10–25 min", "El proceso nuevo", "Recorrido por el proceso TO-BE. Cada paso, cada rol, cada protocolo."),
    ("25–45 min", "Demostración del CRM", "León muestra cómo registrar un lead, moverlo en el pipeline y hacer seguimiento."),
    ("45–60 min", "Práctica guiada", "Cada asesor practica con un lead de prueba. León corrige en tiempo real."),
    ("60–70 min", "El dashboard del dueño", "El equipo ve lo que el dueño verá. Transparencia total del sistema."),
    ("70–80 min", "Preguntas y respuestas", "Espacio abierto. León responde con honestidad."),
    ("80–90 min", "Firma de compromisos", "Cada asesor firma su Acuerdo de Desempeño y su hoja de capacitación."),
]

for tiempo, actividad, desc in agenda:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(5)
    rt = p.add_run(f"[{tiempo}] ")
    rt.font.bold = True
    rt.font.color.rgb = DORADO
    ra = p.add_run(actividad + ": ")
    ra.font.bold = True
    ra.font.color.rgb = AZUL
    rd = p.add_run(desc)
    rd.font.size = Pt(11)
    rd.font.color.rgb = NEGRO

doc.add_paragraph()
h2(doc, "7.2 Preguntas frecuentes del equipo")
faqs = [
    ("¿Me van a vigilar más?", "Sí, habrá más visibilidad — pero para toda la empresa, incluyendo el dueño. No es control, es claridad. El sistema también protege al asesor: si hubo un problema con un cliente, los registros muestran qué pasó realmente."),
    ("¿Es más trabajo para mí?", "Las primeras dos semanas sí requieren adaptación. Después, registrar en el CRM tarda menos de 2 minutos por lead. Lo que sí cambia es que ya no pierdes tiempo buscando dónde dejaste la información de ese cliente."),
    ("¿Van a cambiar mis metas?", "Las metas se definen con criterio, no se improvisan. Vas a tener una meta clara para el mes — lo que cambia es que ahora van a medirse de verdad."),
    ("¿Qué pasa si no cumplo?", "El sistema no es un mecanismo de castigo. Es una herramienta para que todos tengan la información que necesitan. Si no llegas a la meta, la conversación es sobre qué falló y cómo mejorar — no sobre quién tiene la culpa."),
]
for pregunta, respuesta in faqs:
    h3(doc, f"P: {pregunta}")
    body(doc, f"R: {respuesta}", indent=0.5)
    doc.add_paragraph()

# ─── SEMANA 8 ─────────────────────────────────────────────────────────────────
doc.add_page_break()
h1(doc, "SEMANA 8 — PRUEBAS Y ENTREGA")
body(doc, "Objetivo: Verificar que el sistema funciona como debe, hacer ajustes finales y declarar la implementación como exitosa.")
doc.add_paragraph()

add_checklist(doc, [
    "Todos los asesores tienen acceso al CRM y saben cómo usarlo",
    "Los primeros leads reales han sido registrados correctamente en el pipeline",
    "Las automatizaciones están funcionando (verificar con lead de prueba)",
    "El dashboard del dueño muestra datos en tiempo real",
    "Las alertas están configuradas y se están recibiendo correctamente",
    "El equipo tiene la tarjeta de bolsillo con el proceso en 1 hoja",
    "Los Acuerdos de Desempeño fueron firmados por todos los asesores",
    "El dueño sabe cómo leer el dashboard sin ayuda de León",
    "Se definió la fecha de la primera reunión semanal de métricas",
    "El retainer mensual fue acordado para el mes 4 en adelante",
], title="Checklist de entrega — La implementación es exitosa cuando:")

doc.add_paragraph()
section_box(doc, "Criterio de éxito de la implementación",
    [
        "El sistema se considera exitosamente implementado cuando:",
        "",
        "1. El 80% o más de los leads son registrados en el CRM el mismo día que llegan",
        "2. El tiempo promedio de primera respuesta es de 10 minutos o menos",
        "3. El dueño puede ver el estado del negocio sin preguntar a nadie",
        "4. El equipo tiene metas claras y sabe en tiempo real si va bien o mal",
        "5. Existe un proceso documentado que cualquier persona nueva podría aprender en 1 día",
    ],
    bg_hex="E8F8F0", title_color=VERDE, icon="✓")

doc.save(OUTPUT)
print(f"✓ Creado: {OUTPUT}")
