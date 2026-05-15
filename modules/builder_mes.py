"""
Construye la data estructurada para la pestaña P&G por Mes.
"""

import pandas as pd
from .pyg_structure import calcular_subtotales, insertar_subtotales, filtrar_ceros, es_ingreso, es_cuenta_hoja


def build_pyg_mes(df_mes: pd.DataFrame):
    """
    Recibe el DataFrame cargado por loader.load_pyg_mes().
    Retorna (filas, value_cols, mes_cols) donde filas es la lista de dicts
    con subtotales insertados.
    """
    df = df_mes.copy()

    # Columnas de valor (todo excepto Cod y Concepto)
    value_cols = [c for c in df.columns if c not in ("Cod", "Concepto")]

    # Construye data_rows
    data_rows = []
    for _, row in df.iterrows():
        values = {col: float(row[col]) for col in value_cols}
        data_rows.append({
            "cod": str(row["Cod"]).strip(),
            "concepto": str(row["Concepto"]).strip(),
            "values": values,
        })

    # Calcula subtotales
    subtotales = calcular_subtotales(data_rows, value_cols)

    # Inserta subtotales en la lista
    filas = insertar_subtotales(data_rows, subtotales, value_cols)
    filas = filtrar_ceros(filas, value_cols)

    # Calcula ingresos netos sumando solo cuentas hoja de ingreso
    all_cods_set = {str(r["cod"]) for r in data_rows}
    ingresos_netos = {col: 0.0 for col in value_cols}
    for row in data_rows:
        cod = str(row["cod"])
        if es_ingreso(cod) and es_cuenta_hoja(cod, all_cods_set):
            for col in value_cols:
                ingresos_netos[col] = ingresos_netos.get(col, 0) + row["values"].get(col, 0)

    # Agrega porcentajes a cada fila
    for fila in filas:
        fila["pct"] = {}
        for col in value_cols:
            base = ingresos_netos.get(col, 0)
            val = fila["values"].get(col, 0)
            fila["pct"][col] = (val / base * 100) if base else 0.0

    return filas, value_cols, ingresos_netos
