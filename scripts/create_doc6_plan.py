#!/usr/bin/env python3
"""Archivo 6: ejecucion_plan_trabajo_8semanas.xlsx"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT = "/mnt/user-data/outputs/ejecucion_plan_trabajo_8semanas.xlsx"

AZUL = "1A2B4A"; DORADO = "C9973A"; BLANCO = "FFFFFF"; GRIS = "F5F5F5"
ROJO = "C0392B"; VERDE = "27AE60"; AMARILLO = "F39C12"; AZUL_CLARO = "D6E4F0"

def fill(h): return PatternFill("solid", fgColor=h)
def font(bold=False, color="222222", size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")
def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def brd(color="CCCCCC"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

wb = openpyxl.Workbook()

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 1 — GANTT VISUAL
# ═══════════════════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "1 - Gantt Visual"
ws1.sheet_view.showGridLines = False

# Column setup
ws1.column_dimensions['A'].width = 20  # Responsable
ws1.column_dimensions['B'].width = 35  # Actividad
ws1.column_dimensions['C'].width = 8   # %
for i in range(8):
    ws1.column_dimensions[get_column_letter(4+i)].width = 10  # Semanas
ws1.column_dimensions[get_column_letter(12)].width = 22  # Entregable

# Title
ws1.merge_cells('A1:L1')
c = ws1['A1']
c.value = "PLAN DE TRABAJO — 8 SEMANAS | SISTEMA COMERCIAL OPERATIVO"
c.fill = fill(AZUL); c.font = font(True, BLANCO, 14); c.alignment = align("center")
ws1.row_dimensions[1].height = 36

ws1.merge_cells('A2:L2')
c = ws1['A2']
c.value = "León López Consultoría | Empresa: _____________________ | Inicio: _____________"
c.fill = fill(DORADO); c.font = font(False, BLANCO, 10, True); c.alignment = align("center")
ws1.row_dimensions[2].height = 20

ws1.row_dimensions[3].height = 8

# Header row
headers = ["RESPONSABLE", "ACTIVIDAD", "%", "SEM 1", "SEM 2", "SEM 3", "SEM 4", "SEM 5", "SEM 6", "SEM 7", "SEM 8", "ENTREGABLE"]
for j, h in enumerate(headers):
    c = ws1.cell(row=4, column=j+1)
    c.value = h
    c.fill = fill(AZUL)
    c.font = font(True, BLANCO, 10)
    c.alignment = align("center")
    c.border = brd()
ws1.row_dimensions[4].height = 24

# Data: (responsable, actividad, sem_inicio, sem_fin, semanas_active, color, entregable)
LEON = "León López"
CLIENTE = "Equipo Cliente"
DUENO = "Dueño / CEO"

actividades = [
    # Pilar 1
    ("PILAR 1 — DIAGNÓSTICO", None, None, None, [], AZUL, ""),
    (LEON, "Sesión de diagnóstico (60–90 min)", 1, 1, [1], DORADO, "Cuestionario diligenciado"),
    (LEON, "Análisis de datos + calculadora", 1, 2, [1,2], DORADO, "Calculadora de oportunidad"),
    (LEON, "Elaboración del informe de diagnóstico", 2, 2, [2], DORADO, "Informe diagnóstico .docx"),
    (DUENO, "Revisión y firma del informe", 2, 2, [2], AZUL_CLARO, "Firma de autorización"),

    # Pilar 2
    ("PILAR 2 — IMPLEMENTACIÓN", None, None, None, [], AZUL, ""),
    (LEON, "Mapeo proceso actual (AS-IS)", 3, 3, [3], DORADO, "Mapa AS-IS"),
    (LEON, "Diseño proceso nuevo (TO-BE)", 3, 3, [3], DORADO, "Proceso TO-BE documentado"),
    (DUENO, "Revisión y aprobación del proceso", 3, 3, [3], AZUL_CLARO, "Aprobación firmada"),
    (LEON, "Configuración CRM GoHighLevel", 4, 4, [4], DORADO, "CRM activo"),
    (LEON, "Integración WhatsApp + automatizaciones", 4, 4, [4], DORADO, "Automatizaciones activas"),
    (LEON, "Diseño y activación del dashboard", 6, 6, [6], DORADO, "Dashboard Lovable"),
    (CLIENTE, "Definición de metas con equipo", 5, 5, [5], "4A90D9", "Acuerdos de desempeño"),
    (LEON, "Capacitación del equipo (90 min)", 7, 7, [7], DORADO, "Firmas de capacitación"),
    (LEON, "Pruebas finales y ajustes", 8, 8, [8], DORADO, ""),
    (DUENO, "Revisión final + firma de entrega", 8, 8, [8], AZUL_CLARO, "Acta de entrega"),

    # Pilar 3
    ("PILAR 3 — EJECUCIÓN", None, None, None, [], AZUL, ""),
    (LEON, "Reuniones semanales de métricas", 3, 8, [3,4,5,6,7,8], DORADO, "Notas de reunión semanal"),
    (CLIENTE, "Registro diario en CRM", 4, 8, [4,5,6,7,8], "4A90D9", "Datos en el sistema"),
    (DUENO, "Revisión del dashboard (5 min/día)", 6, 8, [6,7,8], AZUL_CLARO, ""),

    # Pilar 4
    ("PILAR 4 — SOCIALIZACIÓN", None, None, None, [], AZUL, ""),
    (LEON, "Kit de lanzamiento al equipo", 7, 7, [7], DORADO, "Kit de comunicación interna"),
    (DUENO, "Carta al equipo + reunión de lanzamiento", 7, 7, [7], AZUL_CLARO, ""),
    (LEON, "Entrega manual del dueño", 8, 8, [8], DORADO, "Manual del dueño"),
    (LEON, "Primer reporte mensual", 8, 8, [8], DORADO, "Reporte mensual template"),
]

row = 5
for act in actividades:
    responsable, actividad, s_ini, s_fin, sems, color, entregable = act
    ws1.row_dimensions[row].height = 22

    if actividad is None:
        # Section header
        ws1.merge_cells(f'A{row}:L{row}')
        c = ws1[f'A{row}']
        c.value = responsable
        c.fill = fill(AZUL)
        c.font = font(True, DORADO, 11)
        c.alignment = align("left")
        row += 1
        continue

    c_resp = ws1.cell(row=row, column=1)
    c_resp.value = responsable
    c_resp.font = font(True if responsable == LEON else False, "222222", 10)
    c_resp.alignment = align("left", "center")
    c_resp.border = brd()

    c_act = ws1.cell(row=row, column=2)
    c_act.value = actividad
    c_act.font = font(size=10)
    c_act.alignment = align("left", "center", wrap=True)
    c_act.border = brd()

    c_pct = ws1.cell(row=row, column=3)
    c_pct.value = 0
    c_pct.number_format = '0%'
    c_pct.alignment = align("center")
    c_pct.border = brd()

    # Gantt bars
    for s in range(8):
        sem_num = s + 1
        c_sem = ws1.cell(row=row, column=4+s)
        c_sem.border = brd("DDDDDD")
        if sem_num in sems:
            c_sem.fill = fill(color)
            c_sem.value = "■"
            c_sem.font = Font(color=color, size=10, name="Calibri")
            c_sem.alignment = align("center")
        else:
            c_sem.fill = fill("F8F8F8")

    c_ent = ws1.cell(row=row, column=12)
    c_ent.value = entregable
    c_ent.font = font(size=9, italic=True, color="444444")
    c_ent.alignment = align("left", "center", wrap=True)
    c_ent.border = brd()

    row += 1

# Legend
row += 1
ws1.cell(row=row, column=1).value = "LEYENDA:"
ws1.cell(row=row, column=1).font = font(True, AZUL, 10)
for col, (label, color) in enumerate([
    ("León López", DORADO),
    ("Equipo Cliente", "4A90D9"),
    ("Dueño/CEO", AZUL_CLARO),
], 2):
    c = ws1.cell(row=row, column=col)
    c.value = f"■ {label}"
    c.fill = fill(color)
    c.font = font(True, BLANCO if color != AZUL_CLARO else AZUL, 10)
    c.alignment = align("center")

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 2 — DASHBOARD DE INDICADORES SEMANA A SEMANA
# ═══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("2 - Indicadores Semana")
ws2.sheet_view.showGridLines = False

ws2.merge_cells('A1:H1')
c = ws2['A1']
c.value = "DASHBOARD DE INDICADORES — SEMANA A SEMANA"
c.fill = fill(AZUL); c.font = font(True, BLANCO, 14); c.alignment = align("center")
ws2.row_dimensions[1].height = 36

ws2.merge_cells('A2:H2')
c = ws2['A2']
c.value = "Complete los datos al final de cada semana. Los campos en amarillo son editables."
c.fill = fill(DORADO); c.font = font(False, BLANCO, 10, True); c.alignment = align("center")
ws2.row_dimensions[2].height = 20

# Column widths
col_widths = [12, 22, 22, 22, 22, 18, 18, 30]
for i, w in enumerate(col_widths):
    ws2.column_dimensions[get_column_letter(i+1)].width = w

# Header
headers2 = ["SEMANA", "Conv. Recibidas\n(total)", "Conv. Atendidas\n(resp. <30min)", "% Respuesta",
            "Leads → Propuesta", "Ventas\nRealizadas", "Meta\nSemana", "Observaciones León"]
for j, h in enumerate(headers2):
    c = ws2.cell(row=3, column=j+1)
    c.value = h
    c.fill = fill(AZUL)
    c.font = font(True, BLANCO, 10)
    c.alignment = align("center", "center", wrap=True)
    c.border = brd()
ws2.row_dimensions[3].height = 35

# Data rows for 8 weeks
for s in range(8):
    row = 4 + s
    ws2.row_dimensions[row].height = 28
    c_sem = ws2.cell(row=row, column=1)
    c_sem.value = f"Semana {s+1}"
    c_sem.fill = fill(AZUL)
    c_sem.font = font(True, DORADO, 10)
    c_sem.alignment = align("center")
    c_sem.border = brd()

    for col in range(2, 8):
        c = ws2.cell(row=row, column=col)
        c.fill = fill("FFF9E6")  # Editable
        c.font = font(size=11)
        c.alignment = align("center")
        c.border = brd()
        if col == 4:  # % respuesta - formula
            c.value = f"=IFERROR(C{row}/B{row},0)"
            c.number_format = '0.0%'
            c.fill = fill(GRIS)

    # Observaciones
    c_obs = ws2.cell(row=row, column=8)
    c_obs.fill = fill("F5F5F5")
    c_obs.border = brd()
    c_obs.alignment = align("left", "center", wrap=True)

# Totals row
row_total = 12
ws2.row_dimensions[row_total].height = 28
ws2.cell(row=row_total, column=1).value = "TOTALES"
ws2.cell(row=row_total, column=1).fill = fill(AZUL)
ws2.cell(row=row_total, column=1).font = font(True, DORADO, 10)
ws2.cell(row=row_total, column=1).alignment = align("center")

for col in [2, 3, 5, 6, 7]:
    c = ws2.cell(row=row_total, column=col)
    c.value = f"=SUM({get_column_letter(col)}4:{get_column_letter(col)}11)"
    c.fill = fill(AZUL)
    c.font = font(True, BLANCO, 11)
    c.alignment = align("center")
    c.border = brd()

c_pct_total = ws2.cell(row=row_total, column=4)
c_pct_total.value = "=IFERROR(C12/B12,0)"
c_pct_total.number_format = '0.0%'
c_pct_total.fill = fill(AZUL)
c_pct_total.font = font(True, BLANCO, 11)
c_pct_total.alignment = align("center")
c_pct_total.border = brd()

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 3 — TRACKER DE ADOPCIÓN DEL EQUIPO
# ═══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("3 - Adopcion Equipo")
ws3.sheet_view.showGridLines = False

ws3.merge_cells('A1:I1')
c = ws3['A1']
c.value = "TRACKER DE ADOPCIÓN DEL EQUIPO"
c.fill = fill(AZUL); c.font = font(True, BLANCO, 14); c.alignment = align("center")
ws3.row_dimensions[1].height = 36

ws3.merge_cells('A2:I2')
c = ws3['A2']
c.value = "Evalúe cada asesor semanalmente. Verde = adopción completa | Amarillo = parcial | Rojo = no usa el sistema"
c.fill = fill(DORADO); c.font = font(False, BLANCO, 10); c.alignment = align("center")
ws3.row_dimensions[2].height = 20

# Columns
col_widths3 = [22, 12, 12, 16, 16, 16, 12, 12, 28]
for i, w in enumerate(col_widths3):
    ws3.column_dimensions[get_column_letter(i+1)].width = w

headers3 = ["ASESOR", "SEM 3", "SEM 4", "SEM 5", "SEM 6", "SEM 7", "SEM 8", "% ADOPCIÓN", "OBSERVACIONES"]
for j, h in enumerate(headers3):
    c = ws3.cell(row=3, column=j+1)
    c.value = h
    c.fill = fill(AZUL)
    c.font = font(True, BLANCO, 10)
    c.alignment = align("center", "center", wrap=True)
    c.border = brd()
ws3.row_dimensions[3].height = 28

for i in range(5):  # 5 asesores max
    row = 4 + i
    ws3.row_dimensions[row].height = 26
    c_asesor = ws3.cell(row=row, column=1)
    c_asesor.value = f"[Asesor {i+1}]"
    c_asesor.fill = fill(GRIS)
    c_asesor.font = font(True, AZUL, 10)
    c_asesor.alignment = align("left")
    c_asesor.border = brd()

    for col in range(2, 8):  # Sem 3-8
        c = ws3.cell(row=row, column=col)
        c.value = "—"
        c.fill = fill("FFF9E6")
        c.font = font(size=10)
        c.alignment = align("center")
        c.border = brd()
        # Validation note
        c.comment = None

    c_pct = ws3.cell(row=row, column=8)
    c_pct.value = f'=COUNTIF(B{row}:G{row},"VERDE")/6'
    c_pct.number_format = '0%'
    c_pct.fill = fill(GRIS)
    c_pct.font = font(size=10)
    c_pct.alignment = align("center")
    c_pct.border = brd()

    c_obs = ws3.cell(row=row, column=9)
    c_obs.fill = fill("F5F5F5")
    c_obs.border = brd()

# ═══════════════════════════════════════════════════════════════════════════════
# HOJA 4 — PROYECCIÓN DE CIERRE DE MES
# ═══════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("4 - Proyeccion Mes")
ws4.sheet_view.showGridLines = False

ws4.column_dimensions['A'].width = 4
ws4.column_dimensions['B'].width = 38
ws4.column_dimensions['C'].width = 22
ws4.column_dimensions['D'].width = 4

ws4.merge_cells('B1:C1')
c = ws4['B1']
c.value = "PROYECCIÓN DE CIERRE DE MES"
c.fill = fill(AZUL); c.font = font(True, BLANCO, 16); c.alignment = align("center")
ws4.row_dimensions[1].height = 40

ws4.merge_cells('B2:C2')
c = ws4['B2']
c.value = "Calcula automáticamente si van a llegar a la meta del mes"
c.fill = fill(DORADO); c.font = font(False, BLANCO, 10, True); c.alignment = align("center")
ws4.row_dimensions[2].height = 20

ws4.row_dimensions[3].height = 10

proj_items = [
    ("INPUTS", None, True, AZUL, DORADO),
    ("Meta total del mes (COP)", None, False, None, None),
    ("Días hábiles del mes", 22, False, None, None),
    ("Día hábil actual (del 1 al 22)", None, False, "FFF9E6", None),
    ("Ventas realizadas hasta hoy (COP)", None, False, "FFF9E6", None),
    ("", None, False, None, None),
    ("PROYECCIÓN AUTOMÁTICA", None, True, AZUL, DORADO),
    ("Ritmo diario actual (COP/día hábil)", None, False, GRIS, None),
    ("Proyección de cierre del mes (COP)", None, False, GRIS, None),
    ("% de avance sobre la meta", None, False, GRIS, None),
    ("¿Van a llegar a la meta?", None, False, GRIS, None),
    ("Días restantes para recuperar si van atrasados", None, False, GRIS, None),
    ("", None, False, None, None),
    ("ALERTA", None, True, ROJO, BLANCO),
    ("Si la proyección es < 80% de la meta:", None, False, "FDECEA", None),
]

row = 4
proj_rows = {}
for i, (label, val, is_sec, bg, fg) in enumerate(proj_items):
    proj_rows[i] = row
    ws4.row_dimensions[row].height = 28

    if is_sec:
        ws4.merge_cells(f'B{row}:C{row}')
        c = ws4[f'B{row}']
        c.value = label
        c.fill = fill(bg)
        c.font = font(True, fg, 11)
        c.alignment = align("left")
    elif label == "":
        ws4.row_dimensions[row].height = 10
    else:
        c_l = ws4[f'B{row}']
        c_l.value = label
        c_l.font = font(size=10)
        c_l.alignment = align("left", "center", wrap=True)
        c_l.border = brd()

        c_v = ws4[f'C{row}']
        if bg == "FFF9E6":
            c_v.fill = fill("FFF9E6")
            c_v.font = font(True, AZUL, 12)
        elif bg == GRIS:
            c_v.fill = fill("F5F5F5")
            c_v.font = font(True, AZUL, 12)
        elif bg == "FDECEA":
            ws4.merge_cells(f'B{row}:C{row}')
            c_v = ws4[f'B{row}']
            c_v.fill = fill("FDECEA")
            c_v.value = "¡ALERTA! Notifique al dueño de inmediato. Revisar plan de choque con el equipo."
            c_v.font = font(True, ROJO, 10)
            row += 1
            continue
        else:
            c_v.font = font(size=11)

        if val is not None:
            c_v.value = val
        c_v.alignment = align("center")
        c_v.border = brd()

        if val is not None and isinstance(val, int) and val > 1000:
            c_v.number_format = '$#,##0" COP"'

    row += 1

# Fill formulas (rows relative to proj_rows)
r_meta = proj_rows[1]; r_dias = proj_rows[2]; r_dia_actual = proj_rows[3]
r_ventas = proj_rows[4]; r_ritmo = proj_rows[7]; r_proyeccion = proj_rows[8]
r_pct = proj_rows[9]; r_llegan = proj_rows[10]; r_dias_rest = proj_rows[11]

ws4[f'C{r_ritmo}'].value = f"=IFERROR(C{r_ventas}/C{r_dia_actual},0)"
ws4[f'C{r_ritmo}'].number_format = '$#,##0" COP"'
ws4[f'C{r_ritmo}'].fill = fill(GRIS)
ws4[f'C{r_ritmo}'].font = font(True, AZUL, 12)
ws4[f'C{r_ritmo}'].alignment = align("center")

ws4[f'C{r_proyeccion}'].value = f"=C{r_ritmo}*C{r_dias}"
ws4[f'C{r_proyeccion}'].number_format = '$#,##0" COP"'
ws4[f'C{r_proyeccion}'].fill = fill(VERDE)
ws4[f'C{r_proyeccion}'].font = font(True, BLANCO, 14)
ws4[f'C{r_proyeccion}'].alignment = align("center")

ws4[f'C{r_pct}'].value = f"=IFERROR(C{r_proyeccion}/C{r_meta},0)"
ws4[f'C{r_pct}'].number_format = '0.0%'
ws4[f'C{r_pct}'].fill = fill(GRIS)
ws4[f'C{r_pct}'].font = font(True, AZUL, 12)
ws4[f'C{r_pct}'].alignment = align("center")

ws4[f'C{r_llegan}'].value = f'=IF(C{r_pct}>=0.8,"SÍ — VAN BIEN","NO — REVISAR PLAN")'
ws4[f'C{r_llegan}'].fill = fill(GRIS)
ws4[f'C{r_llegan}'].font = font(True, VERDE, 12)
ws4[f'C{r_llegan}'].alignment = align("center")

ws4[f'C{r_dias_rest}'].value = f"=MAX(0,C{r_dias}-C{r_dia_actual})"
ws4[f'C{r_dias_rest}'].fill = fill(GRIS)
ws4[f'C{r_dias_rest}'].font = font(True, AZUL, 12)
ws4[f'C{r_dias_rest}'].alignment = align("center")
ws4[f'C{r_dias_rest}'].number_format = '0" días"'

wb.save(OUTPUT)
print(f"✓ Creado: {OUTPUT}")
