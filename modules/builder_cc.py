"""
Construye la data estructurada para la pestaña P&G por Centro de Costo.
"""

import pandas as pd
from .pyg_structure import calcular_subtotales, insertar_subtotales, filtrar_ceros, es_ingreso, es_cuenta_hoja


def build_pyg_cc(df_cc: pd.DataFrame):
    """
    Recibe el DataFrame de loader.load_pyg_cc().
    Retorna (filas, cc_cols).
    """
    df = df_cc.copy()

    # Columnas de centros de costo (todo excepto Cod y Concepto)
    cc_cols = [c for c in df.columns if c not in ("Cod", "Concepto")]

    # Agrega columna TOTAL si no existe
    if "TOTAL" not in cc_cols:
        df["TOTAL"] = df[cc_cols].sum(axis=1)
        cc_cols = cc_cols + ["TOTAL"]

    all_cols = cc_cols

    # Construye data_rows
    data_rows = []
    for _, row in df.iterrows():
        values = {col: float(row[col]) for col in all_cols}
        data_rows.append({
            "cod": str(row["Cod"]).strip(),
            "concepto": str(row["Concepto"]).strip(),
            "values": values,
        })

    # Calcula subtotales
    subtotales = calcular_subtotales(data_rows, all_cols)

    # Inserta subtotales
    filas = insertar_subtotales(data_rows, subtotales, all_cols)
    filas = filtrar_ceros(filas, all_cols)

    # Porcentajes respecto al TOTAL de ingresos (solo cuentas hoja)
    all_cods_set = {str(r["cod"]) for r in data_rows}
    ingreso_total = {col: 0.0 for col in all_cols}
    for row in data_rows:
        cod = str(row["cod"])
        if es_ingreso(cod) and es_cuenta_hoja(cod, all_cods_set):
            for col in all_cols:
                ingreso_total[col] = ingreso_total.get(col, 0) + row["values"].get(col, 0)

    for fila in filas:
        fila["pct"] = {}
        for col in all_cols:
            base = ingreso_total.get(col, 0) or ingreso_total.get("TOTAL", 0)
            val = fila["values"].get(col, 0)
            fila["pct"][col] = (val / base * 100) if base else 0.0

    return filas, all_cols
