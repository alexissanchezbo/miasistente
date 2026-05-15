"""
Genera el archivo Excel de salida con 4 pestañas y formato profesional.
"""

import io
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ── Paleta de colores ────────────────────────────────────────────────────────
C_TITULO_BG   = "1A1A2E"   # Azul marino muy oscuro
C_TITULO_FG   = "FFFFFF"
C_HEADER_BG   = "16213E"   # Azul oscuro
C_HEADER_FG   = "FFFFFF"
C_SEC1_BG     = "0F3460"   # Azul medio (sección raíz 4, 5.1, 5.2)
C_SEC1_FG     = "FFFFFF"
C_SEC2_BG     = "1B4F72"   # Azul más claro (subsección)
C_SEC2_FG     = "FFFFFF"
C_SUBTOT_BG   = "154360"   # Azul enfatizado para subtotales
C_SUBTOT_FG   = "F0F3FF"
C_DETAIL_BG   = "FFFFFF"
C_DETAIL_FG   = "1A1A2E"
C_DETAIL_ALT  = "EBF5FB"   # Filas alternas
C_NEG_FG      = "C0392B"   # Rojo para negativos
C_OBS_ALERTA  = "FDEDEC"   # Fondo alerta
C_OBS_AVISO   = "FEFBD8"   # Fondo aviso
C_OBS_INFO    = "EAF4FB"   # Fondo info

FMT_MONEY   = '#,##0.00'
FMT_MONEY_N = '#,##0.00;[Red]-#,##0.00'
FMT_PCT     = '0.0%'


def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=10, italic=False):
    return Font(bold=bold, color=color, size=size, name="Calibri", italic=italic)

def _align(horizontal="right", vertical="center", wrap=False):
    return Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap)

def _thin_border():
    thin = Side(style="thin", color="CCCCCC")
    return Border(bottom=thin)


def _nivel_cuenta(cod):
    return str(cod).count(".")


def _write_pyg_sheet(ws, empresa, titulo, filas, value_cols, show_pct=True):
    """
    Escribe una hoja del PyG con las filas y columnas dadas.
    filas: lista de dicts {'tipo', 'cod', 'concepto', 'values', 'pct'}
    value_cols: lista de nombres de columnas de valor
    """
    # ── Encabezados ──────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 14

    total_cols = 2 + (len(value_cols) * (2 if show_pct else 1))  # cod+concepto + valores
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    c = ws.cell(1, 1, empresa)
    c.font = _font(bold=True, color=C_TITULO_FG, size=13)
    c.fill = _fill(C_TITULO_BG)
    c.alignment = _align("center")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_cols)
    c = ws.cell(2, 1, titulo)
    c.font = _font(bold=False, color=C_TITULO_FG, size=11)
    c.fill = _fill(C_HEADER_BG)
    c.alignment = _align("center")

    # Fila en blanco
    for col in range(1, total_cols + 1):
        ws.cell(3, col).fill = _fill(C_HEADER_BG)

    # ── Fila de columnas ────────────────────────────────────────────────────
    ws.row_dimensions[4].height = 28
    ws.cell(4, 1, "Cód.").font = _font(bold=True, color=C_HEADER_FG, size=9)
    ws.cell(4, 1).fill = _fill(C_HEADER_BG)
    ws.cell(4, 1).alignment = _align("center")

    ws.cell(4, 2, "Concepto").font = _font(bold=True, color=C_HEADER_FG, size=9)
    ws.cell(4, 2).fill = _fill(C_HEADER_BG)
    ws.cell(4, 2).alignment = _align("left")

    col_idx = 3
    for vc in value_cols:
        ws.merge_cells(start_row=4, start_column=col_idx, end_row=4,
                       end_column=col_idx + (1 if show_pct else 0))
        c = ws.cell(4, col_idx, str(vc))
        c.font = _font(bold=True, color=C_HEADER_FG, size=9)
        c.fill = _fill(C_HEADER_BG)
        c.alignment = _align("center")
        col_idx += 2 if show_pct else 1

    # Subencabezados $ / %
    if show_pct:
        ws.row_dimensions[5].height = 14
        ws.cell(5, 1).fill = _fill(C_HEADER_BG)
        ws.cell(5, 2).fill = _fill(C_HEADER_BG)
        col_idx = 3
        for vc in value_cols:
            c1 = ws.cell(5, col_idx, "$")
            c1.font = _font(bold=True, color=C_HEADER_FG, size=8, italic=True)
            c1.fill = _fill(C_HEADER_BG)
            c1.alignment = _align("center")
            c2 = ws.cell(5, col_idx + 1, "%")
            c2.font = _font(bold=True, color=C_HEADER_FG, size=8, italic=True)
            c2.fill = _fill(C_HEADER_BG)
            c2.alignment = _align("center")
            col_idx += 2
        data_start_row = 6
    else:
        data_start_row = 5

    # ── Datos ────────────────────────────────────────────────────────────────
    alt = False
    for fila in filas:
        tipo    = fila.get("tipo", "data")
        cod     = str(fila.get("cod", ""))
        concepto = str(fila.get("concepto", ""))
        values  = fila.get("values", {})
        pcts    = fila.get("pct", {})
        nivel   = _nivel_cuenta(cod)

        row = ws.max_row + 1
        ws.row_dimensions[row].height = 14 if tipo == "data" else 16

        # ── Colores según tipo y nivel ───────────────────────────────────────
        if tipo == "subtotal":
            bg = C_SUBTOT_BG
            fg = C_SUBTOT_FG
            bold = True
            size = 10
        elif nivel == 0:  # Raíz: cuentas 4, 5
            bg = C_SEC1_BG
            fg = C_SEC1_FG
            bold = True
            size = 10
        elif nivel == 1:  # 5.1, 5.2
            bg = C_SEC2_BG
            fg = C_SEC2_FG
            bold = True
            size = 9
        elif nivel == 2:  # 5.1.1, 5.1.2...
            bg = "2E4057" if not alt else "3D566E"
            fg = "DDEEFF"
            bold = True
            size = 9
        else:
            bg = C_DETAIL_ALT if alt else C_DETAIL_BG
            fg = C_DETAIL_FG
            bold = False
            size = 9

        if tipo == "data" and nivel >= 3:
            alt = not alt

        # ── Escribir celdas ─────────────────────────────────────────────────
        c_cod = ws.cell(row, 1, cod if tipo == "data" else "")
        c_cod.font = _font(bold=bold, color=fg, size=size - 1)
        c_cod.fill = _fill(bg)
        c_cod.alignment = _align("left")

        indent = "  " * max(0, nivel - 2) if tipo == "data" else ""
        c_con = ws.cell(row, 2, indent + concepto)
        c_con.font = _font(bold=bold, color=fg, size=size)
        c_con.fill = _fill(bg)
        c_con.alignment = _align("left", wrap=True)

        col_idx = 3
        for vc in value_cols:
            val = values.get(vc, 0.0) or 0.0
            pct = pcts.get(vc, 0.0) or 0.0

            # Color rojo si negativo (solo en filas de detalle o subtotal)
            val_fg = fg
            if val < 0 and nivel >= 2:
                val_fg = C_NEG_FG if bg in (C_DETAIL_BG, C_DETAIL_ALT) else "FFAAAA"

            c_val = ws.cell(row, col_idx, val)
            c_val.number_format = FMT_MONEY
            c_val.font = _font(bold=bold, color=val_fg, size=size)
            c_val.fill = _fill(bg)
            c_val.alignment = _align("right")

            if show_pct:
                c_pct = ws.cell(row, col_idx + 1, pct / 100)
                c_pct.number_format = FMT_PCT
                c_pct.font = _font(bold=False, color=fg, size=size - 1, italic=True)
                c_pct.fill = _fill(bg)
                c_pct.alignment = _align("right")
                col_idx += 2
            else:
                col_idx += 1

    # ── Anchos de columna ────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 46
    col_idx = 3
    for _ in value_cols:
        ws.column_dimensions[get_column_letter(col_idx)].width = 16
        if show_pct:
            ws.column_dimensions[get_column_letter(col_idx + 1)].width = 8
            col_idx += 2
        else:
            col_idx += 1

    # Fija las dos primeras columnas y las primeras 5 filas
    ws.freeze_panes = ws.cell(data_start_row, 3)


def _write_observaciones_sheet(ws, observaciones, empresa):
    """Escribe la pestaña de observaciones."""
    ws.row_dimensions[1].height = 22
    ws.merge_cells("A1:F1")
    c = ws.cell(1, 1, empresa)
    c.font = _font(bold=True, color=C_TITULO_FG, size=13)
    c.fill = _fill(C_TITULO_BG)
    c.alignment = _align("center")

    ws.merge_cells("A2:F2")
    c = ws.cell(2, 1, "Observaciones y Análisis de Estados Financieros")
    c.font = _font(bold=False, color=C_TITULO_FG, size=11)
    c.fill = _fill(C_HEADER_BG)
    c.alignment = _align("center")

    # Encabezados
    ws.row_dimensions[4].height = 20
    headers = ["#", "Categoría", "Nivel", "Observación"]
    col_widths = [5, 32, 12, 80]
    for i, (h, w) in enumerate(zip(headers, col_widths), 1):
        c = ws.cell(4, i, h)
        c.font = _font(bold=True, color=C_HEADER_FG, size=10)
        c.fill = _fill(C_HEADER_BG)
        c.alignment = _align("center" if i != 4 else "left")
        ws.column_dimensions[get_column_letter(i)].width = w

    nivel_colors = {
        "ALERTA": C_OBS_ALERTA,
        "AVISO":  C_OBS_AVISO,
        "INFO":   C_OBS_INFO,
    }
    nivel_fg = {
        "ALERTA": "922B21",
        "AVISO":  "7D6608",
        "INFO":   "1A5276",
    }

    for idx, obs in enumerate(observaciones, 1):
        row = ws.max_row + 1
        ws.row_dimensions[row].height = 30

        nivel = obs.get("nivel", "INFO")
        bg = nivel_colors.get(nivel, C_OBS_INFO)
        fg_n = nivel_fg.get(nivel, "000000")

        ws.cell(row, 1, idx).font = _font(size=9, color="555555")
        ws.cell(row, 1).fill = _fill(bg)
        ws.cell(row, 1).alignment = _align("center")

        ws.cell(row, 2, obs.get("categoria", "")).font = _font(bold=True, size=9, color=fg_n)
        ws.cell(row, 2).fill = _fill(bg)
        ws.cell(row, 2).alignment = _align("left", wrap=True)

        ws.cell(row, 3, nivel).font = _font(bold=True, size=9, color=fg_n)
        ws.cell(row, 3).fill = _fill(bg)
        ws.cell(row, 3).alignment = _align("center")

        ws.cell(row, 4, obs.get("descripcion", "")).font = _font(size=9)
        ws.cell(row, 4).fill = _fill(bg)
        ws.cell(row, 4).alignment = _align("left", wrap=True)

    ws.freeze_panes = "A5"


def exportar_excel(
    empresa,
    filas_mes, value_cols_mes,
    filas_proyecto, value_cols_proyecto,
    filas_cc, value_cols_cc,
    observaciones,
    titulo_mes="Estado de Resultados Comparativo Mensual",
    titulo_proyecto="Estado de Resultados por Proyecto (MOD y CIF prorrateados por ingresos)",
    titulo_cc="Estado de Resultados por Centro de Costo",
):
    wb = Workbook()

    # ── Pestaña 1: Por Mes ───────────────────────────────────────────────────
    ws_mes = wb.active
    ws_mes.title = "P&G por Mes"
    _write_pyg_sheet(ws_mes, empresa, titulo_mes, filas_mes, value_cols_mes)

    # ── Pestaña 2: Por Proyecto ──────────────────────────────────────────────
    ws_proy = wb.create_sheet("P&G por Proyecto")
    _write_pyg_sheet(ws_proy, empresa, titulo_proyecto, filas_proyecto, value_cols_proyecto)

    # ── Pestaña 3: Por Centro de Costo ───────────────────────────────────────
    ws_cc = wb.create_sheet("P&G por Centro de Costo")
    _write_pyg_sheet(ws_cc, empresa, titulo_cc, filas_cc, value_cols_cc)

    # ── Pestaña 4: Observaciones ─────────────────────────────────────────────
    ws_obs = wb.create_sheet("Observaciones")
    _write_observaciones_sheet(ws_obs, observaciones, empresa)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
