"""
Módulo de depreciación mensual devengada.
Lee el registro de Activos Fijos exportado de Contifico y calcula
la depreciación mensual lineal por cuenta contable, para inyectarla
en el P&G por Mes (gestión / reportes gerenciales).

Fórmula:
    Dep. mensual = (Valor Inicial × (1 − % Residual) × % Anual) / 12
    Solo si el activo aún no está totalmente depreciado.
"""

import pandas as pd

# ── Mapeo Categoría → configuración (basado en Contifico BAEC) ───────────────
# Las claves deben coincidir con el campo "Categoria" del archivo Activos.xls
# (normalizado: sin tilde, strip).
CATEGORIA_DEP_MAP = {
    "Edificios": {
        "pct_anual":    0.05,
        "pct_residual": 0.01,
        "cuenta_gasto": "5.2.1.1.46.1",
        "no_depreciar": False,
    },
    "Equipos de Computacion": {
        "pct_anual":    0.33,
        "pct_residual": 0.01,
        "cuenta_gasto": "5.2.1.1.46.4",
        "no_depreciar": False,
    },
    "MAQUINARIA COMPRESOR": {
        "pct_anual":    0.10,
        "pct_residual": 0.10,
        "cuenta_gasto": "5.1.4.1.2",
        "no_depreciar": False,
    },
    "Maquinarias y Equipos": {
        "pct_anual":    0.10,
        "pct_residual": 0.01,
        "cuenta_gasto": "5.1.4.1.2",
        "no_depreciar": False,
    },
    "Muebles y Enseres": {
        "pct_anual":    0.10,
        "pct_residual": 0.01,
        "cuenta_gasto": "5.2.1.1.46.3",
        "no_depreciar": False,
    },
    "Terreno": {
        "pct_anual":    0.00,
        "pct_residual": 0.01,
        "cuenta_gasto": None,
        "no_depreciar": True,   # ✅ marcado en Contifico
    },
    "Vehiculos": {              # normalizado sin tilde
        "pct_anual":    0.20,
        "pct_residual": 0.01,
        "cuenta_gasto": "5.2.1.1.46.5",
        "no_depreciar": False,
    },
}

def _normalizar(texto: str) -> str:
    """Quita tildes, strip, lower para comparación robusta."""
    import unicodedata
    nfkd = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in nfkd if not unicodedata.combining(c)).strip()


def calcular_dep_mensual_por_cuenta(df_activos: pd.DataFrame) -> dict:
    """
    Calcula la depreciación mensual total agrupada por cuenta gasto.

    Parámetros
    ----------
    df_activos : DataFrame de load_activos()

    Retorna
    -------
    dict  {cuenta_codigo: dep_mensual_total}
    """
    resultado: dict[str, float] = {}

    for _, row in df_activos.iterrows():
        # Solo activos vigentes
        estado = _normalizar(str(row.get("Estado", ""))).lower()
        if estado != "activo":
            continue

        categoria_raw = str(row.get("Categoria", "")).strip()
        # Intentar match directo, luego normalizado
        config = CATEGORIA_DEP_MAP.get(categoria_raw)
        if config is None:
            cat_norm = _normalizar(categoria_raw)
            for key, val in CATEGORIA_DEP_MAP.items():
                if _normalizar(key) == cat_norm:
                    config = val
                    break

        if config is None or config["no_depreciar"] or config["pct_anual"] == 0:
            continue

        valor_inicial = float(row.get("Valor Inicial", 0) or 0)
        valor_actual  = float(row.get("Valor Actual",  0) or 0)
        valor_residual_min = valor_inicial * config["pct_residual"]

        # Si ya está totalmente depreciado → omitir
        if valor_actual <= valor_residual_min + 0.01:
            continue

        # Depreciación mensual lineal
        base        = valor_inicial * (1.0 - config["pct_residual"])
        dep_mensual = (base * config["pct_anual"]) / 12.0

        cuenta = config["cuenta_gasto"]
        resultado[cuenta] = resultado.get(cuenta, 0.0) + dep_mensual

    return resultado


def inyectar_en_df_mes(df_mes: pd.DataFrame, dep_map: dict) -> tuple[pd.DataFrame, list[str], list[str]]:
    """
    Inyecta los montos de depreciación mensual en df_mes.
    Reemplaza los valores actuales (normalmente 0) de cada cuenta de
    depreciación en TODAS las columnas de mes con el monto mensual calculado.

    Parámetros
    ----------
    df_mes  : DataFrame de load_pyg_mes()
              Columnas: Cod, Concepto + columnas de mes (Ene, Feb, …)
    dep_map : {cuenta_codigo: dep_mensual}

    Retorna
    -------
    (df_ajustado, cuentas_inyectadas, cuentas_no_encontradas)
    """
    df = df_mes.copy()

    skip_cols = {"Cod", "Concepto", "cod", "concepto"}
    mes_cols  = [c for c in df.columns if c not in skip_cols]
    cod_col   = "Cod" if "Cod" in df.columns else "cod"

    inyectadas      = []
    no_encontradas  = []

    for cuenta, dep_mensual in dep_map.items():
        mask = df[cod_col].astype(str).str.strip() == cuenta
        if mask.any():
            for mc in mes_cols:
                df.loc[mask, mc] = dep_mensual
            inyectadas.append(cuenta)
        else:
            no_encontradas.append(cuenta)

    return df, inyectadas, no_encontradas


def resumen_depreciacion(df_activos: pd.DataFrame, dep_map: dict, n_meses: int = 1) -> list[dict]:
    """
    Genera un resumen por categoría para mostrar en la app.

    Retorna lista de dicts:
        {categoria, cuenta, dep_mensual, dep_ytd, activos_n}
    """
    resumen: dict[str, dict] = {}

    for _, row in df_activos.iterrows():
        estado = _normalizar(str(row.get("Estado", ""))).lower()
        if estado != "activo":
            continue

        categoria_raw = str(row.get("Categoria", "")).strip()
        config = CATEGORIA_DEP_MAP.get(categoria_raw)
        if config is None:
            cat_norm = _normalizar(categoria_raw)
            for key, val in CATEGORIA_DEP_MAP.items():
                if _normalizar(key) == cat_norm:
                    config = val
                    break

        if config is None or config["no_depreciar"] or config["pct_anual"] == 0:
            continue

        valor_inicial  = float(row.get("Valor Inicial", 0) or 0)
        valor_actual   = float(row.get("Valor Actual",  0) or 0)
        valor_residual_min = valor_inicial * config["pct_residual"]
        if valor_actual <= valor_residual_min + 0.01:
            continue

        base        = valor_inicial * (1.0 - config["pct_residual"])
        dep_mensual = (base * config["pct_anual"]) / 12.0
        cuenta      = config["cuenta_gasto"]

        if categoria_raw not in resumen:
            resumen[categoria_raw] = {
                "categoria":    categoria_raw,
                "cuenta":       cuenta,
                "dep_mensual":  0.0,
                "activos_n":    0,
            }
        resumen[categoria_raw]["dep_mensual"] += dep_mensual
        resumen[categoria_raw]["activos_n"]   += 1

    filas = []
    for cat, d in resumen.items():
        filas.append({
            **d,
            "dep_ytd": d["dep_mensual"] * n_meses,
        })
    filas.sort(key=lambda x: x["dep_mensual"], reverse=True)
    return filas
