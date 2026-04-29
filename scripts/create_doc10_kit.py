#!/usr/bin/env python3
"""Archivo 10: socializacion_kit_lanzamiento_equipo.docx"""
import sys; sys.path.insert(0, '/home/user/sistema-dies/scripts')
from docx_helpers import *

OUTPUT = "/mnt/user-data/outputs/socializacion_kit_lanzamiento_equipo.docx"

doc = new_doc()
add_header_footer(doc, "Kit de Lanzamiento — Equipo")

portada(doc,
    "KIT DE LANZAMIENTO",
    "Para el Equipo Comercial",
    "Pilar 4: Socialización | Comunicación interna del nuevo sistema")

doc.add_page_break()

# CARTA DEL DUEÑO
h1(doc, "PARTE 1 — CARTA DEL DUEÑO AL EQUIPO")
body(doc, "[Template editable. El dueño personaliza con su nombre y detalles de la empresa antes de enviarla.]")
doc.add_paragraph()

tbl = doc.add_table(rows=1, cols=1)
cell = tbl.cell(0, 0)
set_cell_bg(cell, "F5F5F5")

carta_lineas = [
    ("[Ciudad], [Fecha]",),
    ("",),
    ("Estimado equipo de [Nombre de la empresa]:",),
    ("",),
    ("Quiero contarles algo importante sobre el camino que estamos tomando como empresa.",),
    ("",),
    ("Durante los últimos meses, he sentido que la operación nos ha exigido más de lo que hemos podido dar. Hemos perdido oportunidades, hemos tenido clientes insatisfechos y yo mismo me he quedado sin información para tomar buenas decisiones. Eso no es culpa de ninguno de ustedes — es una señal de que necesitamos un sistema mejor.",),
    ("",),
    ("Por eso decidí contratar a León López para implementar un Sistema Comercial Operativo. No es un sistema de control — es una herramienta para que todos, incluido yo, tengamos mayor claridad de lo que está pasando en el negocio.",),
    ("",),
    ("Lo que VA a cambiar:",),
    ("→ Vamos a tener un proceso claro de principio a fin. Cada uno sabrá exactamente qué hacer y cuándo.",),
    ("→ Vamos a registrar nuestras conversaciones con clientes en un sistema. No para vigilarlos — para no perder información valiosa.",),
    ("→ Vamos a tener metas claras y medibles. No voy a improvizar al final del mes para saber cómo va el negocio.",),
    ("",),
    ("Lo que NO va a cambiar:",),
    ("→ Su autonomía para atender a los clientes. El sistema los apoya, no los reemplaza.",),
    ("→ La confianza que tengo en cada uno de ustedes.",),
    ("→ El ambiente de trabajo que hemos construido juntos.",),
    ("",),
    ("Les pido que le den una oportunidad real a este proceso. Que hagan las preguntas que necesiten. Que usen el sistema aunque al principio se sienta diferente.",),
    ("",),
    ("Los resultados del negocio son también los resultados de ustedes. Cuando la empresa crece con orden, todos ganamos.",),
    ("",),
    ("Gracias por su compromiso.",),
    ("",),
    ("[Nombre del dueño]",),
    ("[Cargo]",),
    ("[Empresa]",),
]

for linea_tuple in carta_lineas:
    linea = linea_tuple[0]
    pl = cell.add_paragraph() if carta_lineas.index(linea_tuple) > 0 else cell.paragraphs[0]
    pl.paragraph_format.left_indent = Cm(0.6)
    pl.paragraph_format.space_before = Pt(2)
    pl.paragraph_format.space_after = Pt(2)
    is_arrow = linea.startswith("→")
    is_titulo = linea.endswith(":") and not linea.startswith("→")
    rl = pl.add_run(linea)
    rl.font.size = Pt(11)
    rl.font.italic = is_arrow
    rl.font.bold = is_titulo
    rl.font.color.rgb = AZUL if is_titulo else (GRIS_TEXTO if linea.startswith("[") else NEGRO)

p_end = cell.add_paragraph()
p_end.paragraph_format.space_after = Pt(10)

doc.add_page_break()

# FAQ DEL EQUIPO
h1(doc, "PARTE 2 — PREGUNTAS FRECUENTES DEL EQUIPO")
body(doc, "Las 10 preguntas que el equipo siempre hace — con respuestas honestas.")
doc.add_paragraph()

faqs = [
    ("1. ¿Me van a vigilar más con este sistema?",
     "Sí, va a haber más visibilidad — para todos, incluyendo el dueño y León. Pero 'vigilar' no es la palabra correcta. El sistema registra lo que pasa con cada lead para no perder información. Si un cliente llama 3 veces y nadie lo atiende, el sistema lo muestra. Eso protege al cliente y también al asesor: si hay un malentendido con un cliente, el registro muestra exactamente qué pasó."),
    ("2. ¿Es más trabajo para mí?",
     "Las primeras dos semanas sí requieren adaptación. Después, registrar un lead en el CRM tarda menos de 2 minutos. Lo que sí cambia es que ya no pierdes tiempo buscando información ('¿dónde dejé el número de ese cliente?'). El sistema centraliza todo. Neto: ahorra tiempo."),
    ("3. ¿Van a cambiar mis metas?",
     "Sí. Pero para bien. Antes, la meta era vaga o llegaba improvisada al final del mes. Con el sistema, cada persona tendrá una meta clara desde el primer día del mes. Sabrás exactamente cuánto necesitas vender y cómo vas en tiempo real."),
    ("4. ¿Qué pasa si no cumplo la meta?",
     "La primera conversación es sobre qué falló y cómo mejorar — no sobre quién tiene la culpa. El sistema nos da datos para entender por qué no se llegó. ¿Fueron pocos leads? ¿Mala tasa de cierre? ¿Falta de seguimiento? Dependiendo de la respuesta, la solución es diferente."),
    ("5. ¿Para qué sirve el CRM si ya tenemos WhatsApp?",
     "WhatsApp es para hablar con el cliente. El CRM es para no perder al cliente. En WhatsApp, si el asesor se va de vacaciones o renuncia, la información se va con él. En el CRM, queda para toda la empresa. Además, el CRM recuerda a quién hay que llamar y cuándo — WhatsApp no hace eso."),
    ("6. ¿León va a estar aquí todo el tiempo?",
     "Durante las 8 semanas de implementación, sí. Después, hay una reunión semanal de 30 minutos y un reporte mensual. León es el que mantiene el sistema funcionando — no es parte del equipo operativo."),
    ("7. ¿Qué pasa con los clientes que ya tengo en mis contactos de WhatsApp personal?",
     "Esos se migran al CRM durante la configuración. León te guía en ese proceso. No perderás ningún cliente. Lo que sí cambia es que esa información quedará en el sistema de la empresa, no en tu teléfono personal."),
    ("8. ¿Si tengo una duda del sistema, a quién le pregunto?",
     "A León directamente. Durante la implementación hay un canal de WhatsApp de soporte. Tu primera persona de contacto es siempre León, no el dueño."),
    ("9. ¿Qué pasa si el sistema falla o se cae?",
     "León es el responsable técnico del sistema. Si algo falla, notifícalo de inmediato por WhatsApp. Mientras tanto, el protocolo es: registrar la información en una nota de texto y cargarla al CRM cuando el sistema vuelva. Nunca dejar un lead sin registro."),
    ("10. ¿Este sistema va a reemplazar mi trabajo?",
     "No. El sistema automatiza las partes repetitivas y administrativas del trabajo (recordatorios, registros, alertas). El trabajo de construir relaciones con los clientes, entender sus necesidades y cerrar una venta — eso sigue siendo tuyo. Un sistema no puede reemplazar a una persona que sabe vender bien."),
]

for pregunta, respuesta in faqs:
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, "1A2B4A")
    pp = cell.paragraphs[0]
    pp.paragraph_format.space_before = Pt(8)
    pp.paragraph_format.space_after = Pt(4)
    pp.paragraph_format.left_indent = Cm(0.4)
    rp = pp.add_run(pregunta)
    rp.font.bold = True
    rp.font.size = Pt(11)
    rp.font.color.rgb = DORADO

    pr = cell.add_paragraph()
    pr.paragraph_format.left_indent = Cm(0.6)
    pr.paragraph_format.space_before = Pt(2)
    pr.paragraph_format.space_after = Pt(8)
    rr = pr.add_run(respuesta)
    rr.font.size = Pt(10)
    rr.font.color.rgb = BLANCO
    doc.add_paragraph()

doc.add_page_break()

# REGLAS DE JUEGO
h1(doc, "PARTE 3 — REGLAS DE JUEGO DEL SISTEMA")
doc.add_paragraph()

h2(doc, "No negociables — estas reglas aplican para todos sin excepción:")
no_negociables = [
    "Todo lead que llegue al negocio debe ser registrado en el CRM el mismo día que llega.",
    "El tiempo máximo de primera respuesta es de 5 minutos en WhatsApp.",
    "Ningún lead se cierra en el sistema sin una razón de cierre registrada (ganado con valor / perdido con motivo).",
    "La asistencia a la reunión semanal de métricas es obligatoria.",
    "Ningún asesor puede tener leads activos en su WhatsApp personal que no estén en el CRM.",
]
for item in no_negociables:
    bullet(doc, item, bold_prefix="✗ Sin excepción:")

doc.add_paragraph()
h2(doc, "Flexibles — el equipo puede proponer ajustes:")
flexibles = [
    "El formato del mensaje de primer contacto puede adaptarse según el estilo del asesor, siempre que incluya los elementos clave.",
    "La frecuencia de registro puede ser durante la conversación o máximo 2 horas después.",
    "Los horarios de las reuniones de seguimiento pueden rotarse según la disponibilidad del equipo.",
    "El nombre de las etapas del pipeline puede ajustarse según la terminología del sector.",
]
for item in flexibles:
    bullet(doc, item, bold_prefix="→ Adaptable:")

doc.add_page_break()

# PLAN DE INCENTIVOS
h1(doc, "PARTE 4 — PLAN DE INCENTIVOS DE ADOPCIÓN")
body(doc, "[OPCIONAL] El cliente puede activar o desactivar este plan de incentivos. Si decide activarlo, complete los campos en amarillo.")
doc.add_paragraph()

section_box(doc, "Incentivos por adopción del sistema — primeras 4 semanas",
    [
        "Premio al asesor con MAYOR porcentaje de leads registrados (semana 1-4):",
        "  Incentivo: [Definir por el cliente — ejemplo: bono de $200.000 COP, día libre, etc.]",
        "",
        "Premio al equipo si el tiempo promedio de respuesta está por debajo de 8 minutos en la semana 4:",
        "  Incentivo: [Definir por el cliente — ejemplo: almuerzo de equipo, bono grupal, etc.]",
        "",
        "Reconocimiento especial al asesor con MAYOR tasa de conversión del mes:",
        "  Incentivo: [Definir por el cliente]",
        "",
        "IMPORTANTE: Los incentivos deben ser anunciados AL INICIO del período, no al final.",
        "Un incentivo sorpresa no motiva. Un incentivo conocido de antemano sí.",
    ],
    bg_hex="E8F8F0", title_color=VERDE, icon="★")

doc.add_page_break()

# HOJA DE FIRMAS
h1(doc, "PARTE 5 — HOJA DE FIRMAS")
body(doc, "El equipo confirma que recibió la capacitación, leyó las reglas y entiende el sistema.")
doc.add_paragraph()

tbl_firmas = doc.add_table(rows=8, cols=4)
tbl_firmas.style = 'Table Grid'
headers_f = ["NOMBRE COMPLETO", "CARGO", "FIRMA", "FECHA"]
for i, h in enumerate(headers_f):
    c = tbl_firmas.cell(0, i)
    set_cell_bg(c, "1A2B4A")
    r = c.paragraphs[0].add_run(h)
    r.font.bold = True
    r.font.color.rgb = DORADO
    r.font.size = Pt(10)

for i in range(7):
    row = tbl_firmas.rows[i+1]
    for j in range(4):
        row.cells[j].paragraphs[0].add_run(" " * 20)
        row.cells[j].paragraphs[0].runs[0].font.size = Pt(12)

doc.add_paragraph()
body(doc, "Capacitación facilitada por:", italic=True)
firma_block(doc, "León López — Sistema Comercial Operativo")

doc.save(OUTPUT)
print(f"✓ Creado: {OUTPUT}")
