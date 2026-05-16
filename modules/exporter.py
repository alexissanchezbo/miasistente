"""
Genera el archivo Excel de salida con 4 pestañas y formato profesional.
"""

import io
import pandas as pd
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
# ▲ rojo = sube · ▼ verde = baja · — = sin cambio
FMT_VAR_PCT = '[Red]"▲ "0.0%;[Green]"▼ "0.0%;"  —"'


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


def _write_pyg_sheet(ws, empresa, titulo, filas, value_cols, show_pct=True, pct_mode="income"):
    """
    Escribe una hoja del PyG con las filas y columnas dadas.
    filas      : lista de dicts {'tipo', 'cod', 'concepto', 'values', 'pct'}
    value_cols : lista de nombres de columnas de valor
    pct_mode   : "income"    → % sobre ingresos (default)
                 "variation" → variación % vs período inmediato anterior
                               (▲ rojo si sube · ▼ verde si baja)
    """
    # ── Detectar cuentas madre (tienen hijos en la lista) ───────────────────
    all_cods = {str(f.get("cod", "")) for f in filas if f.get("tipo") == "data"}
    parent_cods = {
        cod for cod in all_cods
        if any(c.startswith(cod + ".") for c in all_cods if c != cod)
    }

    # ── Encabezados ──────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 14

    total_cols = 2 + (len(value_cols) * (2 if show_pct else 1))
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

    for col in range(1, total_cols + 1):
        ws.cell(3, col).fill = _fill(C_HEADER_BG)

    # ── Fila de nombres de columnas ─────────────────────────────────────────
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

    # ── Subencabezados $ / % (o Var.%) ──────────────────────────────────────
    if show_pct:
        ws.row_dimensions[5].height = 14
        ws.cell(5, 1).fill = _fill(C_HEADER_BG)
        ws.cell(5, 2).fill = _fill(C_HEADER_BG)
        col_idx = 3
        for vi, vc in enumerate(value_cols):
            c1 = ws.cell(5, col_idx, "$")
            c1.font = _font(bold=True, color=C_HEADER_FG, size=8, italic=True)
            c1.fill = _fill(C_HEADER_BG)
            c1.alignment = _align("center")

            # Primer período en modo variación: sin subencabezado de %
            if pct_mode == "variation" and vi == 0:
                lbl_pct = ""
            elif pct_mode == "variation":
                lbl_pct = "Var.%"
            else:
                lbl_pct = "%"

            c2 = ws.cell(5, col_idx + 1, lbl_pct)
            c2.font = _font(bold=True, color=C_HEADER_FG, size=8, italic=True)
            c2.fill = _fill(C_HEADER_BG)
            c2.alignment = _align("center")
            col_idx += 2
        data_start_row = 6
    else:
        data_start_row = 5

    # ── Datos ────────────────────────────────────────────────────────────────
    # Registrar columnas de variación para aplicar formato condicional al final
    var_pct_col_letters = []   # letras de columna Excel de cada Var.%

    alt = False
    for fila in filas:
        tipo     = fila.get("tipo", "data")
        cod      = str(fila.get("cod", ""))
        concepto = str(fila.get("concepto", ""))
        values   = fila.get("values", {})
        pcts     = fila.get("pct", {})
        nivel    = _nivel_cuenta(cod)

        row = ws.max_row + 1
        ws.row_dimensions[row].height = 14 if tipo == "data" else 16

        # ── Colores según tipo y nivel ───────────────────────────────────────
        if tipo == "subtotal":
            bg, fg, bold, size = C_SUBTOT_BG, C_SUBTOT_FG, True, 10
        elif nivel == 0:
            bg, fg, bold, size = C_SEC1_BG, C_SEC1_FG, True, 10
        elif nivel == 1:
            bg, fg, bold, size = C_SEC2_BG, C_SEC2_FG, True, 9
        elif nivel == 2:
            bg = "2E4057" if not alt else "3D566E"
            fg, bold, size = "DDEEFF", True, 9
        else:
            bg   = C_DETAIL_ALT if alt else C_DETAIL_BG
            fg   = C_DETAIL_FG
            # Cuenta madre en nivel de detalle → negrita
            bold = (cod in parent_cods)
            size = 9

        if tipo == "data" and nivel >= 3:
            alt = not alt

        # ── Código y concepto ────────────────────────────────────────────────
        c_cod = ws.cell(row, 1, cod if tipo == "data" else "")
        c_cod.font = _font(bold=bold, color=fg, size=size - 1)
        c_cod.fill = _fill(bg)
        c_cod.alignment = _align("left")

        indent = "  " * max(0, nivel - 2) if tipo == "data" else ""
        c_con = ws.cell(row, 2, indent + concepto)
        c_con.font = _font(bold=bold, color=fg, size=size)
        c_con.fill = _fill(bg)
        c_con.alignment = _align("left", wrap=True)

        # ── Valores ──────────────────────────────────────────────────────────
        col_idx = 3
        for vi, vc in enumerate(value_cols):
            val      = values.get(vc, 0.0) or 0.0
            pct      = pcts.get(vc, 0.0)   or 0.0

            val_fg = fg
            if val < 0 and nivel >= 2:
                val_fg = C_NEG_FG if bg in (C_DETAIL_BG, C_DETAIL_ALT) else "FFAAAA"

            c_val = ws.cell(row, col_idx, val)
            c_val.number_format = FMT_MONEY
            c_val.font  = _font(bold=bold, color=val_fg, size=size)
            c_val.fill  = _fill(bg)
            c_val.alignment = _align("right")

            if show_pct:
                pct_col_letter = get_column_letter(col_idx + 1)

                if pct_mode == "variation":
                    if vi == 0:
                        # Primer período: celda vacía
                        c_pct = ws.cell(row, col_idx + 1, None)
                        c_pct.number_format = "@"
                    else:
                        prev_vc  = value_cols[vi - 1]
                        prev_val = values.get(prev_vc, 0.0) or 0.0
                        if abs(prev_val) > 0.0001:
                            variation = (val - prev_val) / abs(prev_val)
                        else:
                            variation = None   # indeterminado

                        c_pct = ws.cell(row, col_idx + 1, variation)
                        c_pct.number_format = FMT_VAR_PCT if variation is not None else "@"

                        # Registrar la columna para el formato condicional de iconos
                        if pct_col_letter not in var_pct_col_letters:
                            var_pct_col_letters.append(pct_col_letter)

                    c_pct.font      = _font(bold=False, color=fg, size=size - 1)
                    c_pct.fill      = _fill(bg)
                    c_pct.alignment = _align("center")

                else:
                    # Modo normal: % sobre ingresos
                    c_pct = ws.cell(row, col_idx + 1, pct / 100)
                    c_pct.number_format = FMT_PCT
                    c_pct.font      = _font(bold=False, color=fg, size=size - 1, italic=True)
                    c_pct.fill      = _fill(bg)
                    c_pct.alignment = _align("right")

                col_idx += 2
            else:
                col_idx += 1

    # (formato condicional de íconos manejado por el número de formato FMT_VAR_PCT)

    # ── Anchos de columna ────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 46
    col_idx = 3
    for _ in value_cols:
        ws.column_dimensions[get_column_letter(col_idx)].width = 16
        if show_pct:
            ws.column_dimensions[get_column_letter(col_idx + 1)].width = 10
            col_idx += 2
        else:
            col_idx += 1

    ws.freeze_panes = ws.cell(data_start_row, 3)


def _write_balance_sheet(ws, empresa, titulo, df_bg):
    """
    Escribe el Estado de Situación Financiera (Balance General).
    df_bg: DataFrame con columnas Cod, Concepto, Total
    Estructura: cuentas 1=Activos, 2=Pasivos, 3=Patrimonio
    """
    # Colores por grupo raíz
    COLORES = {
        "1": ("0B3D91", "FFFFFF"),   # Activos → azul oscuro
        "2": ("6B1A1A", "FFFFFF"),   # Pasivos → burdeos
        "3": ("145A32", "FFFFFF"),   # Patrimonio → verde oscuro
    }
    C_SUB   = ("1C2833", "F0F3FF")  # Subtotales
    C_D_BG  = "FFFFFF"
    C_D_ALT = "F2F3F4"
    C_D_FG  = "1A1A2E"

    # ── Encabezados ──────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 14

    for r in (1, 2, 3):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(1, 1, empresa)
    c.font = _font(bold=True, color=C_TITULO_FG, size=13)
    c.fill = _fill(C_TITULO_BG); c.alignment = _align("center")

    c = ws.cell(2, 1, titulo)
    c.font = _font(bold=False, color=C_TITULO_FG, size=11)
    c.fill = _fill(C_HEADER_BG); c.alignment = _align("center")

    for col in range(1, 4):
        ws.cell(3, col).fill = _fill(C_HEADER_BG)

    ws.row_dimensions[4].height = 24
    for col, (lbl, w, al) in enumerate([
        ("Cód.", 12, "center"), ("Concepto", 52, "left"), ("Saldo", 18, "right")
    ], 1):
        c = ws.cell(4, col, lbl)
        c.font = _font(bold=True, color=C_HEADER_FG, size=10)
        c.fill = _fill(C_HEADER_BG)
        c.alignment = _align(al)
        ws.column_dimensions[get_column_letter(col)].width = w

    data_start = 5

    # ── Detectar cuentas madre ───────────────────────────────────────────────
    all_cods = set(df_bg["Cod"].astype(str).str.strip())
    parent_cods = {
        cod for cod in all_cods
        if any(c.startswith(cod + ".") for c in all_cods if c != cod)
    }

    # ── Acumuladores para subtotales ─────────────────────────────────────────
    totales = {"1": 0.0, "2": 0.0, "3": 0.0}

    alt = False
    grupo_actual = None

    for _, rec in df_bg.iterrows():
        cod      = str(rec["Cod"]).strip()
        concepto = str(rec["Concepto"]).strip()
        valor    = float(rec["Total"]) if pd.notna(rec["Total"]) else 0.0
        nivel    = cod.count(".")
        raiz     = cod[0] if cod else ""

        if raiz not in ("1", "2", "3"):
            continue

        # Separador de sección al cambiar de raíz
        if raiz != grupo_actual:
            if grupo_actual is not None:
                # Línea subtotal del grupo anterior
                lbl_sub = {"1": "TOTAL ACTIVOS", "2": "TOTAL PASIVOS", "3": "TOTAL PATRIMONIO"}.get(grupo_actual, "TOTAL")
                row_s = ws.max_row + 1
                ws.row_dimensions[row_s].height = 16
                bg_s, fg_s = C_SUB
                for col in range(1, 4):
                    ws.cell(row_s, col).fill = _fill(bg_s)
                ws.cell(row_s, 2, lbl_sub).font = _font(bold=True, color=fg_s, size=10)
                ws.cell(row_s, 2).fill = _fill(bg_s)
                ws.cell(row_s, 2).alignment = _align("left")
                c_tot = ws.cell(row_s, 3, totales[grupo_actual])
                c_tot.number_format = FMT_MONEY
                c_tot.font = _font(bold=True, color=fg_s, size=10)
                c_tot.fill = _fill(bg_s)
                c_tot.alignment = _align("right")

            grupo_actual = raiz

        # Acumular solo hojas
        if cod not in parent_cods:
            totales[raiz] = totales.get(raiz, 0.0) + valor

        # ── Estilo por nivel ─────────────────────────────────────────────────
        if nivel == 0:
            bg_r, fg_r = COLORES.get(raiz, ("333333", "FFFFFF"))
            bg, fg, bold, size = bg_r, fg_r, True, 11
        elif nivel == 1:
            bg_r, fg_r = COLORES.get(raiz, ("333333", "FFFFFF"))
            # Versión más clara del color raíz
            bg = {"1": "154360", "2": "922B21", "3": "1E8449"}.get(raiz, "333333")
            fg, bold, size = "FFFFFF", True, 10
        elif nivel == 2:
            bg = {"1": "1A5276", "2": "A93226", "3": "27AE60"}.get(raiz, "444444")
            fg, bold, size = "EAFAF1" if raiz == "3" else "DDEEFF" if raiz == "1" else "FDEDEC", True, 9
        else:
            bg   = C_D_ALT if alt else C_D_BG
            fg   = C_D_FG
            bold = cod in parent_cods
            size = 9
            alt  = not alt

        row = ws.max_row + 1
        ws.row_dimensions[row].height = 14 if nivel >= 2 else 16

        indent = "  " * max(0, nivel - 1) if nivel >= 2 else ""

        c_cod = ws.cell(row, 1, cod if nivel >= 2 else "")
        c_cod.font = _font(bold=bold, color=fg, size=size - 1)
        c_cod.fill = _fill(bg); c_cod.alignment = _align("left")

        c_con = ws.cell(row, 2, indent + concepto)
        c_con.font = _font(bold=bold, color=fg, size=size)
        c_con.fill = _fill(bg); c_con.alignment = _align("left", wrap=True)

        c_val = ws.cell(row, 3, valor)
        c_val.number_format = FMT_MONEY
        val_color = "FFAAAA" if valor < 0 else fg
        c_val.font = _font(bold=bold, color=val_color, size=size)
        c_val.fill = _fill(bg); c_val.alignment = _align("right")

    # Subtotal del último grupo
    if grupo_actual:
        lbl_sub = {"1": "TOTAL ACTIVOS", "2": "TOTAL PASIVOS", "3": "TOTAL PATRIMONIO"}.get(grupo_actual, "TOTAL")
        row_s = ws.max_row + 1
        ws.row_dimensions[row_s].height = 16
        bg_s, fg_s = C_SUB
        for col in range(1, 4):
            ws.cell(row_s, col).fill = _fill(bg_s)
        ws.cell(row_s, 2, lbl_sub).font = _font(bold=True, color=fg_s, size=10)
        ws.cell(row_s, 2).fill = _fill(bg_s)
        ws.cell(row_s, 2).alignment = _align("left")
        c_tot = ws.cell(row_s, 3, totales.get(grupo_actual, 0.0))
        c_tot.number_format = FMT_MONEY
        c_tot.font = _font(bold=True, color=fg_s, size=10)
        c_tot.fill = _fill(bg_s)
        c_tot.alignment = _align("right")

    # ── Línea final: TOTAL PASIVO + PATRIMONIO ───────────────────────────────
    total_pp = totales.get("2", 0.0) + totales.get("3", 0.0)
    row_f = ws.max_row + 1
    ws.row_dimensions[row_f].height = 18
    for col in range(1, 4):
        ws.cell(row_f, col).fill = _fill(C_TITULO_BG)
    ws.cell(row_f, 2, "TOTAL PASIVO + PATRIMONIO").font = _font(bold=True, color="FFFFFF", size=11)
    ws.cell(row_f, 2).fill = _fill(C_TITULO_BG)
    ws.cell(row_f, 2).alignment = _align("left")
    c_fp = ws.cell(row_f, 3, total_pp)
    c_fp.number_format = FMT_MONEY
    c_fp.font = _font(bold=True, color="FFD700", size=11)
    c_fp.fill = _fill(C_TITULO_BG)
    c_fp.alignment = _align("right")

    ws.freeze_panes = ws.cell(data_start, 1)


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

    ws.row_dimensions[4].height = 20
    headers = ["#", "Categoría", "Nivel", "Observación"]
    col_widths = [5, 32, 12, 80]
    for i, (h, w) in enumerate(zip(headers, col_widths), 1):
        c = ws.cell(4, i, h)
        c.font = _font(bold=True, color=C_HEADER_FG, size=10)
        c.fill = _fill(C_HEADER_BG)
        c.alignment = _align("center" if i != 4 else "left")
        ws.column_dimensions[get_column_letter(i)].width = w

    nivel_colors = {"ALERTA": C_OBS_ALERTA, "AVISO": C_OBS_AVISO, "INFO": C_OBS_INFO}
    nivel_fg     = {"ALERTA": "922B21",      "AVISO": "7D6608",    "INFO": "1A5276"}

    for idx, obs in enumerate(observaciones, 1):
        row  = ws.max_row + 1
        ws.row_dimensions[row].height = 30
        nivel = obs.get("nivel", "INFO")
        bg    = nivel_colors.get(nivel, C_OBS_INFO)
        fg_n  = nivel_fg.get(nivel, "000000")

        ws.cell(row, 1, idx).font          = _font(size=9, color="555555")
        ws.cell(row, 1).fill               = _fill(bg)
        ws.cell(row, 1).alignment          = _align("center")

        ws.cell(row, 2, obs.get("categoria", "")).font = _font(bold=True, size=9, color=fg_n)
        ws.cell(row, 2).fill               = _fill(bg)
        ws.cell(row, 2).alignment          = _align("left", wrap=True)

        ws.cell(row, 3, nivel).font        = _font(bold=True, size=9, color=fg_n)
        ws.cell(row, 3).fill               = _fill(bg)
        ws.cell(row, 3).alignment          = _align("center")

        ws.cell(row, 4, obs.get("descripcion", "")).font = _font(size=9)
        ws.cell(row, 4).fill               = _fill(bg)
        ws.cell(row, 4).alignment          = _align("left", wrap=True)

    ws.freeze_panes = "A5"


def exportar_excel(
    empresa,
    filas_mes, value_cols_mes,
    filas_proyecto, value_cols_proyecto,
    filas_cc, value_cols_cc,
    observaciones,
    df_balance=None,
    titulo_mes="Estado de Resultados Comparativo Mensual",
    titulo_proyecto="Estado de Resultados por Proyecto (MOD y CIF prorrateados por ingresos)",
    titulo_cc="Estado de Resultados por Centro de Costo",
    titulo_balance="Estado de Situación Financiera",
):
    wb = Workbook()

    # ── Pestaña 1: Por Mes — variación % vs período anterior ────────────────
    ws_mes = wb.active
    ws_mes.title = "P&G por Mes"
    _write_pyg_sheet(
        ws_mes, empresa, titulo_mes,
        filas_mes, value_cols_mes,
        pct_mode="variation",
    )

    # ── Pestaña 2: Por Proyecto ──────────────────────────────────────────────
    ws_proy = wb.create_sheet("P&G por Proyecto")
    _write_pyg_sheet(ws_proy, empresa, titulo_proyecto, filas_proyecto, value_cols_proyecto)

    # ── Pestaña 3: Por Centro de Costo ───────────────────────────────────────
    ws_cc = wb.create_sheet("P&G por Centro de Costo")
    _write_pyg_sheet(ws_cc, empresa, titulo_cc, filas_cc, value_cols_cc)

    # ── Pestaña 4: Estado de Situación Financiera (Balance) ──────────────────
    if df_balance is not None and not df_balance.empty:
        ws_bg = wb.create_sheet("Situación Financiera")
        _write_balance_sheet(ws_bg, empresa, titulo_balance, df_balance)

    # ── Pestaña 5: Observaciones ─────────────────────────────────────────────
    ws_obs = wb.create_sheet("Observaciones")
    _write_observaciones_sheet(ws_obs, observaciones, empresa)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
