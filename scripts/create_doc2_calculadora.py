#!/usr/bin/env python3
"""Archivo 2: diagnostico_calculadora_oportunidad.xlsx"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint

OUTPUT = "/mnt/user-data/outputs/diagnostico_calculadora_oportunidad.xlsx"

AZUL = "1A2B4A"
DORADO = "C9973A"
BLANCO = "FFFFFF"
GRIS = "F5F5F5"
GRIS_BORDE = "CCCCCC"
ROJO = "C0392B"
VERDE = "27AE60"
AMARILLO = "F39C12"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="222222", size=11, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def border_all(color=GRIS_BORDE):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def border_bottom(color=GRIS_BORDE):
    s = Side(style="thin", color=color)
    return Border(bottom=s)

def apply_header_row(ws, row, cols, bg=AZUL, fg=BLANCO, size=11):
    for col in cols:
        c = ws.cell(row=row, column=col)
        c.fill = fill(bg)
        c.font = font(bold=True, color=fg, size=size)
        c.alignment = align("center")
        c.border = border_all()

def format_cop(ws, row, col):
    ws.cell(row=row, column=col).number_format = '#,##0" COP"'

def format_pct(ws, row, col):
    ws.cell(row=row, column=col).number_format = '0.0%'

wb = openpyxl.Workbook()

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 1 — DATOS DE ENTRADA
# ═══════════════════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "1 - Datos de Entrada"
ws1.sheet_view.showGridLines = False

# Column widths
ws1.column_dimensions['A'].width = 4
ws1.column_dimensions['B'].width = 48
ws1.column_dimensions['C'].width = 22
ws1.column_dimensions['D'].width = 18
ws1.column_dimensions['E'].width = 4

# Title
ws1.merge_cells('B1:D1')
c = ws1['B1']
c.value = "CALCULADORA DE OPORTUNIDAD COMERCIAL"
c.fill = fill(AZUL)
c.font = font(bold=True, color=BLANCO, size=16)
c.alignment = align("center", "center")
ws1.row_dimensions[1].height = 40

ws1.merge_cells('B2:D2')
c = ws1['B2']
c.value = "Sistema Comercial Operativo | León López | Diagnóstico Pilar 1"
c.fill = fill(DORADO)
c.font = font(bold=False, color=BLANCO, size=11, italic=True)
c.alignment = align("center", "center")
ws1.row_dimensions[2].height = 22

ws1.row_dimensions[3].height = 10

# Section header
ws1.merge_cells('B4:D4')
c = ws1['B4']
c.value = "DATOS DE ENTRADA — Complete los campos amarillos"
c.fill = fill("F0F0F0")
c.font = font(bold=True, color=AZUL, size=12)
c.alignment = align("left", "center")
ws1.row_dimensions[4].height = 28

# Input rows
inputs = [
    ("VOLUMEN DE LEADS", None, None),
    ("Conversaciones nuevas por día (WhatsApp + todos los canales)", 25, "número entero"),
    ("Días hábiles por mes", 22, "no modificar"),
    ("", None, None),
    ("TASA DE CONVERSIÓN", None, None),
    ("Tasa de cierre actual estimada (%)", 0.12, "ej: 0.12 = 12%"),
    ("Tasa de cierre potencial con sistema (%)", 0.20, "default 20%, editable"),
    ("", None, None),
    ("VALOR DEL NEGOCIO", None, None),
    ("Ticket promedio por venta (COP)", 280000, "valor en pesos"),
    ("Número de asesores comerciales", 3, "número entero"),
    ("", None, None),
    ("EFICIENCIA OPERATIVA", None, None),
    ("Tiempo promedio de respuesta actual (minutos)", 45, "ej: 45 = 45 minutos"),
    ("% de leads sin seguimiento después del primer contacto", 0.30, "default 30%, editable"),
    ("% de pérdida por falta de métricas individuales", 0.08, "default 8%, editable"),
]

row_map = {}  # label -> row number
row = 5
for label, value, note in inputs:
    if label in ("VOLUMEN DE LEADS", "TASA DE CONVERSIÓN", "VALOR DEL NEGOCIO", "EFICIENCIA OPERATIVA"):
        ws1.merge_cells(f'B{row}:D{row}')
        c = ws1[f'B{row}']
        c.value = label
        c.fill = fill(AZUL)
        c.font = font(bold=True, color=DORADO, size=11)
        c.alignment = align("left", "center")
        ws1.row_dimensions[row].height = 24
    elif label == "":
        ws1.row_dimensions[row].height = 8
    else:
        row_map[label] = row
        c_label = ws1[f'B{row}']
        c_label.value = label
        c_label.font = font(size=11)
        c_label.alignment = align("left", "center", wrap=True)
        c_label.border = border_all(GRIS_BORDE)

        c_val = ws1[f'C{row}']
        c_val.value = value
        c_val.fill = fill("FFF9E6")
        c_val.font = font(bold=True, color=AZUL, size=12)
        c_val.alignment = align("center", "center")
        c_val.border = border_all(DORADO)
        if isinstance(value, float) and value < 1:
            c_val.number_format = '0.0%'
        elif isinstance(value, int) and value > 1000:
            c_val.number_format = '#,##0" COP"'

        c_note = ws1[f'D{row}']
        c_note.value = note
        c_note.font = font(size=9, italic=True, color="888888")
        c_note.alignment = align("left", "center")

        ws1.row_dimensions[row].height = 30
    row += 1

# Save row numbers for formulas in sheet 2
# B6=conversaciones/dia, B7=dias habiles, B11=cierre actual, B12=cierre potencial
# B15=ticket, B16=asesores, B19=tiempo respuesta, B20=% sin seguimiento, B21=% perdida metricas

row += 1
ws1.merge_cells(f'B{row}:D{row}')
c = ws1[f'B{row}']
c.value = "NOTA: Los campos en amarillo son editables. El resto se calcula automáticamente en la Hoja 2."
c.font = font(size=9, italic=True, color="888888")
c.alignment = align("center")

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 2 — CÁLCULO DE PÉRDIDA
# ═══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("2 - Calculo de Perdida")
ws2.sheet_view.showGridLines = False

ws2.column_dimensions['A'].width = 4
ws2.column_dimensions['B'].width = 50
ws2.column_dimensions['C'].width = 24
ws2.column_dimensions['D'].width = 4

# Title
ws2.merge_cells('B1:C1')
c = ws2['B1']
c.value = "CÁLCULO DE PÉRDIDA MENSUAL Y OPORTUNIDAD"
c.fill = fill(AZUL)
c.font = font(bold=True, color=BLANCO, size=16)
c.alignment = align("center", "center")
ws2.row_dimensions[1].height = 40

ws2.merge_cells('B2:C2')
c = ws2['B2']
c.value = "Los valores se calculan automáticamente desde los datos de entrada"
c.fill = fill(DORADO)
c.font = font(italic=True, color=BLANCO, size=10)
c.alignment = align("center", "center")
ws2.row_dimensions[2].height = 20

ws2.row_dimensions[3].height = 10

# Data references from sheet 1
# Find actual row numbers
conv_dia_row = None
dias_row = None
cierre_actual_row = None
cierre_potencial_row = None
ticket_row = None
asesores_row = None
tiempo_resp_row = None
sin_seguimiento_row = None
sin_metricas_row = None

for label, r in row_map.items():
    if "Conversaciones nuevas" in label:
        conv_dia_row = r
    elif "Días hábiles" in label:
        dias_row = r
    elif "cierre actual" in label:
        cierre_actual_row = r
    elif "cierre potencial" in label:
        cierre_potencial_row = r
    elif "Ticket promedio" in label:
        ticket_row = r
    elif "asesores" in label:
        asesores_row = r
    elif "Tiempo promedio" in label:
        tiempo_resp_row = r
    elif "sin seguimiento" in label:
        sin_seguimiento_row = r
    elif "métricas individuales" in label:
        sin_metricas_row = r

# Hardcode defaults if mapping fails
if not conv_dia_row: conv_dia_row = 6
if not dias_row: dias_row = 7
if not cierre_actual_row: cierre_actual_row = 11
if not cierre_potencial_row: cierre_potencial_row = 12
if not ticket_row: ticket_row = 15
if not asesores_row: asesores_row = 16
if not tiempo_resp_row: tiempo_resp_row = 19
if not sin_seguimiento_row: sin_seguimiento_row = 20
if not sin_metricas_row: sin_metricas_row = 21

S1 = "'1 - Datos de Entrada'"

calc_items = [
    # (label, formula or value, is_section_header, bg_override, fg_override)
    ("PÉRDIDA POR TIEMPOS DE RESPUESTA LENTOS", None, True, AZUL, DORADO),
    ("Conversaciones totales al mes",
     f"={S1}!C{conv_dia_row}*{S1}!C{dias_row}", False, None, None),
    ("% leads perdidos por respuesta lenta (>30 min = -40% conversión)",
     0.40, False, None, None),
    ("Conversaciones perdidas por tiempo lento (estimado)",
     None, False, None, None),  # formula below
    ("Valor perdido por tiempos lentos (COP/mes)",
     None, False, ROJO, BLANCO),

    ("PÉRDIDA POR FALTA DE SEGUIMIENTO", None, True, AZUL, DORADO),
    ("% leads sin seguimiento (del diagnóstico)",
     f"={S1}!C{sin_seguimiento_row}", False, None, None),
    ("Conversaciones sin seguimiento al mes",
     None, False, None, None),
    ("Valor perdido por falta de seguimiento (COP/mes)",
     None, False, ROJO, BLANCO),

    ("PÉRDIDA POR FALTA DE MÉTRICAS Y VISIBILIDAD", None, True, AZUL, DORADO),
    ("% pérdida estimada por falta de métricas",
     f"={S1}!C{sin_metricas_row}", False, None, None),
    ("Valor perdido por ceguera gerencial (COP/mes)",
     None, False, ROJO, BLANCO),

    ("RESUMEN DE OPORTUNIDAD", None, True, AZUL, BLANCO),
    ("TOTAL PÉRDIDA MENSUAL ESTIMADA (COP)",
     None, False, ROJO, BLANCO),
    ("POTENCIAL DE RECUPERACIÓN CON SISTEMA (COP)",
     None, False, VERDE, BLANCO),
    ("Inversión total del sistema (COP)",
     12000000, False, None, None),
    ("ROI: Meses para recuperar la inversión",
     None, False, DORADO, BLANCO),
    ("Facturación potencial anual adicional (COP)",
     None, False, VERDE, BLANCO),
]

row = 4
calc_rows = {}  # index -> row
for i, item in enumerate(calc_items):
    label, val, is_section, bg, fg = item
    calc_rows[i] = row
    if is_section:
        ws2.merge_cells(f'B{row}:C{row}')
        c = ws2[f'B{row}']
        c.value = label
        c.fill = fill(bg)
        c.font = font(bold=True, color=fg, size=11)
        c.alignment = align("left", "center")
        ws2.row_dimensions[row].height = 26
    else:
        c_label = ws2[f'B{row}']
        c_label.value = label
        c_label.font = font(size=10, color="222222")
        c_label.alignment = align("left", "center", wrap=True)
        c_label.border = border_all()

        c_val = ws2[f'C{row}']
        if bg:
            c_val.fill = fill(bg)
            c_val.font = font(bold=True, color=fg, size=13)
            ws2.row_dimensions[row].height = 32
        else:
            c_val.font = font(bold=False, size=11, color="222222")
            ws2.row_dimensions[row].height = 26
        c_val.alignment = align("center", "center")
        c_val.border = border_all()
        if val is not None:
            c_val.value = val

    row += 1

# Now fill in formulas
# Row indices (0-based in calc_items):
# 0=section, 1=conv_mes, 2=pct_lento, 3=conv_perdidas, 4=valor_tiempos
# 5=section, 6=pct_seguimiento, 7=conv_sin_seg, 8=valor_seguimiento
# 9=section, 10=pct_metricas, 11=valor_metricas
# 12=section, 13=total_perdida, 14=potencial, 15=inversion, 16=roi, 17=facturacion_anual

r_conv_mes = calc_rows[1]
r_pct_lento = calc_rows[2]
r_conv_perdidas = calc_rows[3]
r_val_tiempos = calc_rows[4]
r_pct_seg = calc_rows[6]
r_conv_sin_seg = calc_rows[7]
r_val_seg = calc_rows[8]
r_pct_metricas = calc_rows[10]
r_val_metricas = calc_rows[11]
r_total_perdida = calc_rows[13]
r_potencial = calc_rows[14]
r_inversion = calc_rows[15]
r_roi = calc_rows[16]
r_facturacion_anual = calc_rows[17]

# Formula: conversaciones perdidas por tiempo lento
ws2[f'C{r_conv_perdidas}'].value = f"=C{r_conv_mes}*C{r_pct_lento}*{S1}!C{cierre_actual_row}"
ws2[f'C{r_conv_perdidas}'].number_format = '#,##0'

# Formula: valor perdido tiempos
ws2[f'C{r_val_tiempos}'].value = f"=C{r_conv_perdidas}*{S1}!C{ticket_row}"
ws2[f'C{r_val_tiempos}'].number_format = '$#,##0" COP"'

# Formula: conv sin seguimiento
ws2[f'C{r_conv_sin_seg}'].value = f"=C{r_conv_mes}*C{r_pct_seg}*{S1}!C{cierre_actual_row}"
ws2[f'C{r_conv_sin_seg}'].number_format = '#,##0'

# Formula: valor perdido seguimiento
ws2[f'C{r_val_seg}'].value = f"=C{r_conv_sin_seg}*{S1}!C{ticket_row}"
ws2[f'C{r_val_seg}'].number_format = '$#,##0" COP"'

# Formula: valor perdido metricas
ws2[f'C{r_val_metricas}'].value = f"=C{r_conv_mes}*{S1}!C{cierre_actual_row}*C{r_pct_metricas}*{S1}!C{ticket_row}"
ws2[f'C{r_val_metricas}'].number_format = '$#,##0" COP"'

# Total pérdida
ws2[f'C{r_total_perdida}'].value = f"=C{r_val_tiempos}+C{r_val_seg}+C{r_val_metricas}"
ws2[f'C{r_total_perdida}'].number_format = '$#,##0" COP"'

# Potencial recuperación (65% de la pérdida total con el sistema)
ws2[f'C{r_potencial}'].value = f"=C{r_total_perdida}*0.65"
ws2[f'C{r_potencial}'].number_format = '$#,##0" COP"'

# ROI en meses
ws2[f'C{r_roi}'].value = f"=IFERROR(C{r_inversion}/C{r_potencial},\"N/A\")"
ws2[f'C{r_roi}'].number_format = '0.0" meses"'

# Facturación anual adicional
ws2[f'C{r_facturacion_anual}'].value = f"=C{r_potencial}*12"
ws2[f'C{r_facturacion_anual}'].number_format = '$#,##0" COP"'

# Format percentage cells
ws2[f'C{r_pct_lento}'].number_format = '0%'

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 3 — RESUMEN EJECUTIVO
# ═══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("3 - Resumen Ejecutivo")
ws3.sheet_view.showGridLines = False

ws3.column_dimensions['A'].width = 3
ws3.column_dimensions['B'].width = 35
ws3.column_dimensions['C'].width = 28
ws3.column_dimensions['D'].width = 3

# Big title
ws3.merge_cells('B1:C1')
c = ws3['B1']
c.value = "RESUMEN EJECUTIVO"
c.fill = fill(AZUL)
c.font = font(bold=True, color=BLANCO, size=20)
c.alignment = align("center", "center")
ws3.row_dimensions[1].height = 50

ws3.merge_cells('B2:C2')
c = ws3['B2']
c.value = "Diagnóstico Comercial | Costo de Oportunidad"
c.fill = fill(DORADO)
c.font = font(italic=True, color=BLANCO, size=12)
c.alignment = align("center", "center")
ws3.row_dimensions[2].height = 24

ws3.row_dimensions[3].height = 16

# KPI cards
kpis = [
    ("PÉRDIDA MENSUAL\nESTIMADA", f"='2 - Calculo de Perdida'!C{r_total_perdida}", ROJO, '$#,##0" COP"', "El dinero que se pierde cada mes por ineficiencias en la operación comercial"),
    ("POTENCIAL DE\nRECUPERACIÓN", f"='2 - Calculo de Perdida'!C{r_potencial}", VERDE, '$#,##0" COP"', "Lo que el sistema puede recuperar en los primeros meses de operación"),
    ("INVERSIÓN\nDEL SISTEMA", 12000000, "2C3E50", '$#,##0" COP"', "Inversión total en la implementación completa del sistema"),
    ("ROI\n(MESES PARA RECUPERAR)", f"='2 - Calculo de Perdida'!C{r_roi}", DORADO, '0.0" meses"', "En cuántos meses la inversión se paga sola"),
    ("FACTURACIÓN\nADICIONAL ANUAL", f"='2 - Calculo de Perdida'!C{r_facturacion_anual}", "1A6B3A", '$#,##0" COP"', "Proyección de ingresos adicionales en 12 meses con el sistema activo"),
]

row = 4
for kpi_title, kpi_val, kpi_color, kpi_fmt, kpi_desc in kpis:
    ws3.row_dimensions[row].height = 28
    ws3.row_dimensions[row+1].height = 42
    ws3.row_dimensions[row+2].height = 22
    ws3.row_dimensions[row+3].height = 12

    # Label
    ws3.merge_cells(f'B{row}:C{row}')
    c = ws3[f'B{row}']
    c.value = kpi_title
    c.fill = fill(kpi_color)
    c.font = font(bold=True, color=BLANCO, size=11)
    c.alignment = align("center", "center")

    # Value
    ws3.merge_cells(f'B{row+1}:C{row+1}')
    c = ws3[f'B{row+1}']
    c.value = kpi_val
    c.fill = fill("F8F8F8")
    c.font = Font(bold=True, color=kpi_color, size=24, name="Calibri")
    c.alignment = align("center", "center")
    c.number_format = kpi_fmt
    c.border = border_all(kpi_color)

    # Description
    ws3.merge_cells(f'B{row+2}:C{row+2}')
    c = ws3[f'B{row+2}']
    c.value = kpi_desc
    c.font = font(size=9, italic=True, color="666666")
    c.alignment = align("center", "center", wrap=True)

    row += 4

# Footer
ws3.merge_cells(f'B{row}:C{row}')
c = ws3[f'B{row}']
c.value = "Proyección basada en datos del diagnóstico. Los resultados reales pueden variar según la adopción del equipo."
c.font = font(size=9, italic=True, color="888888")
c.alignment = align("center", "center", wrap=True)
ws3.row_dimensions[row].height = 28

row += 2
ws3.merge_cells(f'B{row}:C{row}')
c = ws3[f'B{row}']
c.value = "Sistema Comercial Operativo | León López Consultoría | Marzo 2026"
c.fill = fill(AZUL)
c.font = font(size=9, color=BLANCO)
c.alignment = align("center", "center")
ws3.row_dimensions[row].height = 20

wb.save(OUTPUT)
print(f"✓ Creado: {OUTPUT}")
