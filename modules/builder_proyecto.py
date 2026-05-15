"""
Construye la data estructurada para la pestaña P&G por Proyecto
usando los mayores de ingresos y costos/gastos filtrados por rango de fechas.

Los gastos sin proyecto asignado se prorratean proporcionalmente
al volumen de ingresos de cada proyecto en el período.
"""

import pandas as pd
from collections import defaultdict
from .pyg_structure import calcular_subtotales, insertar_subtotales, filtrar_ceros, es_ingreso, es_cuenta_hoja


def build_pyg_proyecto(
    df_ingresos: pd.DataFrame,
    df_costos: pd.DataFrame,
    df_mes: pd.DataFrame,
    fecha_inicio=None,
    fecha_fin=None,
):
    """
    Parámetros:
      df_ingresos:  DataFrame de loader.load_mayor() — mayor de cuentas 4.x
      df_costos:    DataFrame de loader.load_mayor() — mayor de cuentas 5.x
      df_mes:       DataFrame de loader.load_pyg_mes() — define orden de filas
      fecha_inicio: datetime o None
      fecha_fin:    datetime o None

    Retorna (filas, value_cols) donde filas incluye subtotales.
    """
    # ── 1. Unir y filtrar por período ────────────────────────────────────────
    mayor = pd.concat([df_ingresos, df_costos], ignore_index=True)

    if not mayor.empty and "Fecha" in mayor.columns:
        mayor["Fecha"] = pd.to_datetime(mayor["Fecha"], errors="coerce")
        if fecha_inicio:
            mayor = mayor[mayor["Fecha"] >= pd.Timestamp(fecha_inicio)]
        if fecha_fin:
            mayor = mayor[mayor["Fecha"] <= pd.Timestamp(fecha_fin)]

    if mayor.empty:
        return [], []

    # ── 2. Acumular montos por (proyecto, codigo) ────────────────────────────
    por_proyecto  = defaultdict(lambda: defaultdict(float))  # {proy: {cod: monto}}
    sin_proyecto  = defaultdict(float)                        # {cod: monto}

    for _, row in mayor.iterrows():
        cod   = str(row.get("Codigo", "")).strip()
        proy  = str(row.get("Proyecto", "")).strip()
        monto = float(row.get("Monto", 0) or 0)

        if not cod or cod == "nan":
            continue

        if proy and proy != "nan":
            por_proyecto[proy][cod] += monto
        else:
            sin_proyecto[cod] += monto

    proyectos = sorted(por_proyecto.keys())
    if not proyectos:
        return [], []

    # ── 3. Calcular prorrateo por ingresos de cada proyecto ─────────────────
    ventas_por_proy = {
        proy: sum(v for cod, v in por_proyecto[proy].items() if es_ingreso(cod))
        for proy in proyectos
    }
    total_ventas = sum(ventas_por_proy.values()) or 1

    factor = {proy: ventas_por_proy[proy] / total_ventas for proy in proyectos}

    # ── 4. Distribuir gastos/costos sin proyecto ─────────────────────────────
    for cod, monto in sin_proyecto.items():
        if not es_ingreso(cod):
            # Los gastos sin proyecto se prorratean
            for proy in proyectos:
                por_proyecto[proy][cod] += monto * factor[proy]
        else:
            # Ingresos sin proyecto van a columna GENERAL
            por_proyecto["__SIN_PROYECTO__"][cod] += monto

    # Si hay ingresos sin proyecto, agregar columna especial al inicio
    tiene_sin_proy = bool(por_proyecto["__SIN_PROYECTO__"])
    all_proyectos = (["Sin Proyecto"] if tiene_sin_proy else []) + proyectos
    proy_key_map = {p: p for p in proyectos}
    if tiene_sin_proy:
        proy_key_map["Sin Proyecto"] = "__SIN_PROYECTO__"

    # ── 5. Construir data_rows con el orden del PyG por Mes ─────────────────
    data_rows = []
    for _, row in df_mes.iterrows():
        cod      = str(row["Cod"]).strip()
        concepto = str(row["Concepto"]).strip()
        values   = {}
        for proy_label in all_proyectos:
            key = proy_key_map.get(proy_label, proy_label)
            values[proy_label] = por_proyecto[key].get(cod, 0.0)
        values["TOTAL"] = sum(values[p] for p in all_proyectos)
        data_rows.append({"cod": cod, "concepto": concepto, "values": values})

    all_cols = all_proyectos + ["TOTAL"]

    # ── 5.5 Recalcular cuentas padre como suma de sus hijas hoja ─────────────
    # El mayor solo tiene movimientos en cuentas hoja; las cuentas de grupo
    # quedan en 0. Aquí se propaga el valor hacia arriba para que cada cuenta
    # padre muestre la suma de sus hijos.
    all_cods_set = {r["cod"] for r in data_rows}
    cod_to_row   = {r["cod"]: r for r in data_rows}

    for row in data_rows:
        cod = row["cod"]
        if not es_cuenta_hoja(cod, all_cods_set):
            prefix = cod + "."
            hojas = [c for c in all_cods_set if c.startswith(prefix) and es_cuenta_hoja(c, all_cods_set)]
            for col in all_cols:
                row["values"][col] = sum(cod_to_row[c]["values"].get(col, 0.0) for c in hojas)
            # Recalcular TOTAL de este padre
            row["values"]["TOTAL"] = sum(row["values"][p] for p in all_proyectos)

    # ── 6. Subtotales ────────────────────────────────────────────────────────
    subtotales = calcular_subtotales(data_rows, all_cols)
    filas = insertar_subtotales(data_rows, subtotales, all_cols)
    filas = filtrar_ceros(filas, all_cols)

    # ── 7. Porcentajes ───────────────────────────────────────────────────────
    ingresos_netos = {col: 0.0 for col in all_cols}
    for row in data_rows:
        cod = str(row["cod"])
        if es_ingreso(cod) and es_cuenta_hoja(cod, all_cods_set):
            for col in all_cols:
                ingresos_netos[col] = ingresos_netos.get(col, 0) + row["values"].get(col, 0)

    for fila in filas:
        fila["pct"] = {}
        for col in all_cols:
            base = ingresos_netos.get(col, 0) or ingresos_netos.get("TOTAL", 0)
            val  = fila["values"].get(col, 0)
            fila["pct"][col] = (val / base * 100) if base else 0.0

    return filas, all_cols
