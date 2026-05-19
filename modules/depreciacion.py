"""
Depreciación mensual devengada — BAEC BALDOSAS DEL ECUADOR S.A.S.
Valores fijos calculados con base en el registro de Activos Fijos de Contifico.
Fórmula: (Valor Inicial × (1 − % Residual) × % Anual) ÷ 12
Solo activos con Estado = Activo y que aún no estén totalmente depreciados.

── CUANDO CAMBIES ALGO ──────────────────────────────────────────────────────
Actualiza el bloque DEP_MENSUAL_FIJA y el detalle de activos de abajo.
El resto de la app lo toma automáticamente sin tocar ningún otro archivo.
─────────────────────────────────────────────────────────────────────────────

Última actualización: 16/05/2026
Activos procesados: 27 activos Activos (1 Inactivo excluido)
Activos omitidos (100% depreciados): 6
  · RELOJ BIOMETRICO
  · LAPTOP DELL INSPIRION #1  ($445.01)
  · IMPRESORA MATRICIAL EPSON LX350
  · COPIADORA E IMPRESORA RICOH
  · LAPTOP DELL INSPIRION #2  ($500.00)
  · LAPTOP DELL INSPIRION #3  ($559.01)
Categorías sin activos: Edificios · Vehículos · Terreno (no deprecia)

── DETALLE POR CUENTA ───────────────────────────────────────────────────────

5.2.1.1.46.4 — Depreciación Equipos de Computación  ($49.22 / mes)
  Activo                              Val.Inicial  %Res  %Anual  Dep/mes
  LENOVO IDEAPAD SLIM 3 15IAN8          328.00     1%    33%      8.93
  NOTEBOOK DELL INSPIRON 3530 i5        620.00     1%    33%     16.88
  NOTEBOOK DELL INSPIRON 3520 i7        625.00     1%    33%     17.02
  IMPRESORA EPSON L3250                 235.00     1%    33%      6.40
                                                              ─────────
                                                               TOTAL  49.22

5.1.4.1.2 — Depreciación Maquinarias y Equipos  ($4,542.57 / mes)
  Activo                              Val.Inicial  %Res  %Anual  Dep/mes
  MONTACARGAS HELI CPCD35-Q22K2       17,400.00    1%   10%    143.55
  BOMBA SUMERGIBLE AANN                1,390.70    1%   10%     11.47
  PULIDORA OMC MOD OM800/6           127,055.59    1%   10%  1,048.21
  PRENSA OMC PRESS MOD. OM600-300T   358,351.14    1%   10%  2,956.40
  MONTACARGAS HELI CPCD30-4.5 TND     28,560.00    1%   10%    235.62
  BOMBA SUMERGIBLE PEDROLLO 7.5HP      1,204.34    1%   10%      9.94
  COMPRESOR CPBG30 30HP               18,317.54   10%   10%    137.38
                                                              ─────────
                                                              TOTAL 4,542.57

5.2.1.1.46.3 — Depreciación Muebles y Enseres  ($325.56 / mes)
  Activo                              Val.Inicial  %Res  %Anual  Dep/mes
  SPLIT CONFORTSTART R410 #1             561.37    1%   10%      4.63
  SPLIT CONFORTSTART R410 #2             615.08    1%   10%      5.07
  CONTENEDOR METALICO ASTM A-36 #1     3,650.00    1%   10%     30.11
  CONTENEDOR METALICO ASTM A-36 #2     3,650.00    1%   10%     30.11
  CONTENEDOR METALICO ASTM A-36 #3     7,300.00    1%   10%     60.23
  CONTENEDOR METALICO ASTM A-36 #4     1,460.00    1%   10%     12.05
  CONTENEDOR METALICO ASTM A-36 #5     7,300.00    1%   10%     60.23
  CONTENEDOR METALICO ASTM A-36 #6     7,300.00    1%   10%     60.23
  CONTENEDOR METALICO ASTM A-36 #7     7,300.00    1%   10%     60.23
  SPLIT 12000 BTU TCL                    325.00    1%   10%      2.68
                                                              ─────────
                                                              TOTAL  325.56
"""

import pandas as pd

# ── Mapeo cuentas P&G → cuentas Balance (dep. acumulada) ────────────────────
DEP_BALANCE_MAP: dict[str, str] = {
    "5.1.4.1.2":    "1.2.1.11.2",  # Dep. Acum. Maquinaria y Equipos
    "5.2.1.1.46.3": "1.2.1.11.3",  # Dep. Acum. Muebles y Enseres
    "5.2.1.1.46.4": "1.2.1.11.4",  # Dep. Acum. Equipos de Computación
}

# ── MONTOS FIJOS (actualizar aquí cuando cambien los activos) ─────────────────
# {cuenta_contable: depreciación_mensual_$}
DEP_MENSUAL_FIJA: dict[str, float] = {
    "5.2.1.1.46.4": 49.22,     # Equipos de Computación  (4 activos)
    "5.1.4.1.2":    4_542.57,  # Maquinarias y Equipos + Compresor  (7 activos)
    "5.2.1.1.46.3": 325.56,    # Muebles y Enseres  (10 activos)
}

# Etiquetas legibles para la UI
DEP_ETIQUETAS: dict[str, str] = {
    "5.2.1.1.46.4": "Depreciación Equipos de Computación",
    "5.1.4.1.2":    "Depreciación Maquinarias y Equipos",
    "5.2.1.1.46.3": "Depreciación Muebles y Enseres",
}


def inyectar_en_df_mes(df_mes: pd.DataFrame) -> tuple[pd.DataFrame, list[str], list[str]]:
    """
    Inyecta los montos fijos de depreciación mensual en df_mes.
    Reemplaza el valor de cada cuenta de depreciación en TODAS las
    columnas de mes con el monto fijo (línea recta constante).

    Parámetros
    ----------
    df_mes : DataFrame de load_pyg_mes()
             Columnas: Cod, Concepto + columnas de mes

    Retorna
    -------
    (df_ajustado, cuentas_inyectadas, cuentas_no_encontradas)
    """
    df = df_mes.copy()

    skip_cols = {"Cod", "Concepto", "cod", "concepto"}
    mes_cols  = [c for c in df.columns if c not in skip_cols]
    cod_col   = "Cod" if "Cod" in df.columns else "cod"

    inyectadas     = []
    no_encontradas = []

    for cuenta, dep_mensual in DEP_MENSUAL_FIJA.items():
        mask = df[cod_col].astype(str).str.strip() == cuenta
        if mask.any():
            for mc in mes_cols:
                df.loc[mask, mc] = dep_mensual
            inyectadas.append(cuenta)
        else:
            no_encontradas.append(cuenta)

    return df, inyectadas, no_encontradas


def inyectar_en_df_cos(
    df_cos: pd.DataFrame,
    fecha_ini,
    n_meses: int,
) -> pd.DataFrame:
    """
    Agrega filas sintéticas de depreciación a df_cos SIN proyecto ni CC.
    El builder_proyecto las prorateará automáticamente por ingresos de
    cada proyecto, igual que cualquier otro gasto sin asignación.

    Parámetros
    ----------
    df_cos    : DataFrame de load_mayor() — mayor de costos
    fecha_ini : fecha de inicio del período (para que pase el filtro de fechas)
    n_meses   : número de meses del período reportado

    Retorna df_cos enriquecido con las filas de depreciación.
    """
    import pandas as _pd
    rows = []
    for cuenta, dep_mensual in DEP_MENSUAL_FIJA.items():
        dep_total = dep_mensual * n_meses
        rows.append({
            "Fecha":       _pd.Timestamp(fecha_ini),
            "Codigo":      cuenta,
            "Cuenta":      DEP_ETIQUETAS.get(cuenta, f"Dep. {cuenta}"),
            "CentroCosto": "",
            "Proyecto":    "",           # sin proyecto → prorrateo automático
            "TipoDoc":     "DEP",
            "NumDoc":      "DEVENGADO",
            "Tercero":     "",
            "Debe":        dep_total,
            "Haber":       0.0,
            "Monto":       dep_total,    # Monto = Debe - Haber para cuentas 5.x
        })
    df_sintetico = _pd.DataFrame(rows)
    return _pd.concat([df_cos, df_sintetico], ignore_index=True)


def inyectar_en_df_cc(df_cc: pd.DataFrame, n_meses: int) -> pd.DataFrame:
    """
    Distribuye la depreciación del período en df_cc proporcionalmente
    a los ingresos (cuentas 4.x) de cada Centro de Costo.

    Parámetros
    ----------
    df_cc   : DataFrame de load_pyg_cc() — Cod, Concepto + columnas CC
    n_meses : número de meses del período

    Retorna df_cc con los valores de depreciación inyectados.
    """
    df = df_cc.copy()
    cc_cols = [c for c in df.columns if c not in ("Cod", "Concepto", "TOTAL")]

    # Calcular ingresos por CC para obtener el ratio de distribución
    mask_ing = df["Cod"].astype(str).str.startswith("4")
    if mask_ing.any() and cc_cols:
        ing_por_cc = df.loc[mask_ing, cc_cols].sum()
        total_ing  = ing_por_cc.sum()
        ratio_cc   = (ing_por_cc / total_ing) if total_ing > 0 else {c: 1/len(cc_cols) for c in cc_cols}
    else:
        ratio_cc = {c: 1 / len(cc_cols) for c in cc_cols} if cc_cols else {}

    cod_col = "Cod"
    for cuenta, dep_mensual in DEP_MENSUAL_FIJA.items():
        dep_total = dep_mensual * n_meses
        mask = df[cod_col].astype(str).str.strip() == cuenta
        if mask.any():
            for cc in cc_cols:
                df.loc[mask, cc] = dep_total * float(ratio_cc.get(cc, 0))
            if "TOTAL" in df.columns:
                df.loc[mask, "TOTAL"] = dep_total
    return df


def ajustar_balance(df_bg: pd.DataFrame, n_meses: int) -> pd.DataFrame:
    """
    Ajusta el balance general para reflejar la depreciación devengada
    del período que aún no ha sido contabilizada:

    1. Aumenta las cuentas de depreciación acumulada (1.2.1.11.x):
       las hace más negativas (mayor deducción al activo bruto).
    2. Reduce el resultado del período en patrimonio (cuenta 3.x)
       por el mismo importe total.

    Parámetros
    ----------
    df_bg   : DataFrame de load_balance() — Cod, Concepto, Total
    n_meses : meses del período (para calcular dep total)

    Retorna df_bg ajustado.
    """
    df = df_bg.copy()
    cod_col = "Cod"
    dep_total_global = 0.0

    # ── 1. Ajustar cuentas de depreciación acumulada ─────────────────────
    for cuenta_pyg, cuenta_bg in DEP_BALANCE_MAP.items():
        dep_mensual = DEP_MENSUAL_FIJA.get(cuenta_pyg, 0.0)
        dep_periodo = dep_mensual * n_meses
        dep_total_global += dep_periodo

        mask = df[cod_col].astype(str).str.strip() == cuenta_bg
        if mask.any():
            # La dep acumulada es contra-activo: restar hace el activo neto menor
            df.loc[mask, "Total"] = df.loc[mask, "Total"] - dep_periodo
        else:
            # La cuenta no está en el balance → agregar como fila nueva
            nueva = {
                "Cod":      cuenta_bg,
                "Concepto": DEP_ETIQUETAS.get(cuenta_pyg, f"Dep. Acum. {cuenta_bg}"),
                "Total":    -dep_periodo,   # negativo = contra-activo
            }
            df = pd.concat([df, pd.DataFrame([nueva])], ignore_index=True)

    # ── 2. Ajustar resultado del período en Patrimonio ────────────────────
    # Buscar cuenta de resultado del ejercicio: primera cuenta 3.x que
    # contenga "resultado", "utilidad" o "ejercicio" en el Concepto.
    keywords = ("resultado", "utilidad", "ejercicio", "periodo", "período")
    mask_res = (
        df[cod_col].astype(str).str.startswith("3") &
        df["Concepto"].astype(str).str.lower().str.contains(
            "|".join(keywords), regex=True
        )
    )
    if mask_res.any():
        df.loc[mask_res, "Total"] = df.loc[mask_res, "Total"] - dep_total_global
    else:
        # Crear fila de ajuste si no se encuentra la cuenta
        df = pd.concat([df, pd.DataFrame([{
            "Cod":      "3.ADJ.DEP",
            "Concepto": "Ajuste depreciación devengada (gestión)",
            "Total":    -dep_total_global,
        }])], ignore_index=True)

    return df


def resumen_depreciacion(n_meses: int = 1) -> list[dict]:
    """
    Retorna resumen de depreciación fija para mostrar en la UI.

    Parámetros
    ----------
    n_meses : número de meses del período reportado (para YTD)

    Retorna
    -------
    Lista de dicts: {cuenta, etiqueta, dep_mensual, dep_ytd}
    """
    return [
        {
            "cuenta":      cuenta,
            "etiqueta":    DEP_ETIQUETAS.get(cuenta, cuenta),
            "dep_mensual": dep_mensual,
            "dep_ytd":     round(dep_mensual * n_meses, 2),
        }
        for cuenta, dep_mensual in DEP_MENSUAL_FIJA.items()
    ]
