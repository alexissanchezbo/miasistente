"""
Genera observaciones automáticas sobre los estados financieros:
- Cuadre entre P&G y Balance General
- Variaciones mensuales significativas
- Costos/gastos inflados
- Alertas del Balance
- Concentración de cuentas
"""

import pandas as pd
from collections import defaultdict


def _find(df, cod, col="Total"):
    """Retorna el valor exacto de una cuenta por código."""
    row = df[df["Cod"].astype(str).str.strip() == str(cod).strip()]
    if row.empty:
        return 0.0
    return float(row[col].iloc[0])


def _val(df, cod_prefix, col="Total"):
    """
    Retorna el valor de la cuenta RAÍZ más cercana al prefijo dado
    (evita doble conteo de subcuentas jerárquicas).
    Primero intenta código exacto, luego el primer código que empiece con el prefijo
    con el menor número de niveles.
    """
    # Intenta exacto primero
    v = _find(df, cod_prefix, col)
    if v != 0.0:
        return v
    # Si no existe exacto, suma solo los hijos directos (un nivel más)
    mask = df["Cod"].astype(str).str.match(
        r"^" + str(cod_prefix).replace(".", r"\.") + r"\.\d+$"
    )
    subset = df.loc[mask] if col in df.columns else df.loc[mask]
    return subset[col].sum() if col in df.columns and not subset.empty else 0.0


def generar_observaciones(df_mes, df_cc, df_transacciones, df_balance):
    """
    Retorna lista de dicts: [{categoria, nivel, descripcion}, ...]
    nivel: 'ALERTA' | 'AVISO' | 'INFO'
    """
    obs = []

    # ── 1. CUADRE ENTRE P&G Y BALANCE ──────────────────────────────────────
    try:
        resultado_balance = _find(df_balance, "3.1.7.1")
        if resultado_balance == 0:
            resultado_balance = _find(df_balance, "3.1.7")

        # Ingresos y gastos totales del PyG
        mes_cols = [c for c in df_mes.columns if c not in ("Cod", "Concepto")]
        if "Total" in mes_cols:
            total_col = "Total"
        else:
            total_col = mes_cols[-1] if mes_cols else None

        if total_col:
            ingresos_pyg = _val(df_mes, "4", total_col)
            costos_pyg   = _val(df_mes, "5", total_col)
            utilidad_pyg = ingresos_pyg - costos_pyg
            diferencia   = abs(utilidad_pyg - resultado_balance)

            if diferencia < 1:
                obs.append({
                    "categoria": "Cuadre de Estados Financieros",
                    "nivel": "INFO",
                    "descripcion": (
                        f"✔ El P&G cuadra con el Balance General. "
                        f"Utilidad Neta P&G: ${utilidad_pyg:,.2f} | "
                        f"Resultado Ejercicio Balance: ${resultado_balance:,.2f}"
                    ),
                })
            else:
                obs.append({
                    "categoria": "Cuadre de Estados Financieros",
                    "nivel": "ALERTA",
                    "descripcion": (
                        f"✘ DESCUADRE entre P&G y Balance. "
                        f"Utilidad P&G: ${utilidad_pyg:,.2f} | "
                        f"Resultado Balance: ${resultado_balance:,.2f} | "
                        f"Diferencia: ${diferencia:,.2f}"
                    ),
                })
    except Exception as e:
        obs.append({"categoria": "Cuadre", "nivel": "AVISO", "descripcion": f"No se pudo verificar cuadre: {e}"})

    # ── 2. VARIACIONES MENSUALES SIGNIFICATIVAS ─────────────────────────────
    # Solo cuentas de 4to nivel o más con variación >80% y monto >500
    try:
        mes_cols = [c for c in df_mes.columns if c not in ("Cod", "Concepto") and "total" not in c.lower()]
        variaciones = []
        if len(mes_cols) >= 2:
            for _, row in df_mes.iterrows():
                cod = str(row["Cod"])
                concepto = str(row["Concepto"])
                if cod.count(".") < 3:  # Solo hojas (nivel 4+)
                    continue
                vals = [float(row[c]) for c in mes_cols if pd.notna(row[c])]
                for i in range(1, len(vals)):
                    prev, curr = vals[i - 1], vals[i]
                    if abs(prev) < 500:
                        continue
                    cambio = (curr - prev) / abs(prev) * 100
                    if abs(cambio) >= 80:
                        variaciones.append({
                            "cambio_abs": abs(cambio),
                            "categoria": "Variación Mensual Significativa",
                            "nivel": "ALERTA" if abs(cambio) >= 150 else "AVISO",
                            "descripcion": (
                                f"Cuenta {cod} ({concepto}): "
                                f"{'aumento' if cambio > 0 else 'reducción'} del {abs(cambio):.1f}% "
                                f"entre {mes_cols[i-1]} y {mes_cols[i]} "
                                f"(${prev:,.2f} → ${curr:,.2f})"
                            ),
                        })
        # Solo las 12 variaciones más grandes
        for v in sorted(variaciones, key=lambda x: -x["cambio_abs"])[:12]:
            obs.append({k: v[k] for k in ("categoria", "nivel", "descripcion")})
    except Exception as e:
        obs.append({"categoria": "Variación Mensual", "nivel": "AVISO", "descripcion": f"Error en análisis mensual: {e}"})

    # ── 3. COSTOS/GASTOS INFLADOS (>3x el promedio mensual, monto >1000) ───
    try:
        mes_cols_num = [c for c in df_mes.columns if c not in ("Cod", "Concepto") and "total" not in c.lower()]
        inflados = []
        if mes_cols_num:
            for _, row in df_mes.iterrows():
                cod = str(row["Cod"])
                concepto = str(row["Concepto"])
                if not cod.startswith("5") or cod.count(".") < 3:
                    continue
                vals_raw = [(mes_cols_num[i], float(row[c]))
                            for i, c in enumerate(mes_cols_num)
                            if pd.notna(row[c]) and float(row[c]) > 0]
                if len(vals_raw) < 2:
                    continue
                nums = [v for _, v in vals_raw]
                promedio = sum(nums) / len(nums)
                if promedio < 1000:
                    continue
                for mes_n, v in vals_raw:
                    ratio = v / promedio
                    if ratio > 3.0:
                        inflados.append({
                            "ratio": ratio,
                            "categoria": "Costo/Gasto Inflado",
                            "nivel": "ALERTA",
                            "descripcion": (
                                f"Cuenta {cod} ({concepto}): ${v:,.2f} en {mes_n} "
                                f"es {ratio:.1f}x el promedio mensual (${promedio:,.2f}). "
                                "Verificar si corresponde a un gasto extraordinario o error de registro."
                            ),
                        })
        for v in sorted(inflados, key=lambda x: -x["ratio"])[:8]:
            obs.append({k: v[k] for k in ("categoria", "nivel", "descripcion")})
    except Exception as e:
        obs.append({"categoria": "Costos Inflados", "nivel": "AVISO", "descripcion": f"Error en análisis: {e}"})

    # ── 4. ALERTAS DEL BALANCE GENERAL ─────────────────────────────────────
    try:
        # Bancos en negativo (sobregiro)
        bancos = df_balance[
            df_balance["Cod"].astype(str).str.startswith("1.1.1.3.") &
            (df_balance["Total"] < 0)
        ]
        for _, b in bancos.iterrows():
            obs.append({
                "categoria": "Alerta Balance — Bancos",
                "nivel": "ALERTA",
                "descripcion": (
                    f"Cuenta bancaria en SOBREGIRO: {b['Cod']} {b['Concepto']} = ${b['Total']:,.2f}. "
                    "Verificar disponibilidad de fondos y créditos bancarios."
                ),
            })

        # Cuentas con signo inusual (activos negativos, pasivos positivos en contexto inverso)
        caja_neg = df_balance[
            df_balance["Cod"].astype(str).str.startswith("1.1.1.2.") &
            (df_balance["Total"] < 0)
        ]
        for _, c in caja_neg.iterrows():
            obs.append({
                "categoria": "Alerta Balance — Caja",
                "nivel": "AVISO",
                "descripcion": f"Caja Chica con saldo negativo: {c['Cod']} {c['Concepto']} = ${c['Total']:,.2f}.",
            })

        # Deuda bancaria total
        deuda_cp = _val(df_balance, "2.1.4")
        deuda_lp = _val(df_balance, "2.2.3")
        deuda_total = deuda_cp + deuda_lp
        if deuda_total > 0:
            obs.append({
                "categoria": "Estructura de Deuda",
                "nivel": "INFO",
                "descripcion": (
                    f"Deuda bancaria total: ${deuda_total:,.2f} "
                    f"(corriente: ${deuda_cp:,.2f} | largo plazo: ${deuda_lp:,.2f})."
                ),
            })

        # Relacionadas con saldos altos
        rel = _val(df_balance, "1.1.2.5.5")
        if rel > 50000:
            obs.append({
                "categoria": "Cuentas por Cobrar Relacionadas",
                "nivel": "AVISO",
                "descripcion": (
                    f"Cuentas por cobrar a compañías relacionadas: ${rel:,.2f}. "
                    "Verificar condiciones y antigüedad de estos saldos."
                ),
            })

        # Otros pasivos de costo de producción (cuenta 2.1.13.21)
        otros_pas = _find(df_balance, "2.1.13.21")
        if otros_pas > 50000:
            obs.append({
                "categoria": "Pasivos Pendientes de Registro",
                "nivel": "ALERTA",
                "descripcion": (
                    f"'Otros Pasivos costo de producción' (2.1.13.21): ${otros_pas:,.2f}. "
                    "Monto elevado — verificar si están todos los costos correctamente registrados en el P&G."
                ),
            })

        # Patrimonio vs deuda
        patrimonio = _val(df_balance, "3")
        if deuda_total > 0 and patrimonio > 0:
            ratio = deuda_total / patrimonio
            nivel = "ALERTA" if ratio > 2 else "AVISO" if ratio > 1 else "INFO"
            obs.append({
                "categoria": "Indicadores Financieros",
                "nivel": nivel,
                "descripcion": f"Ratio Deuda/Patrimonio: {ratio:.2f}x (Deuda: ${deuda_total:,.2f} | Patrimonio: ${patrimonio:,.2f}).",
            })

        # Inventarios altos
        inventario = _val(df_balance, "1.1.3")
        ingresos_total = 0
        mes_cols2 = [c for c in df_mes.columns if c not in ("Cod", "Concepto") and "total" not in c.lower()]
        if mes_cols2:
            last_col = mes_cols2[-1]
            ingresos_total = _val(df_mes, "4", last_col)
        if ingresos_total > 0 and inventario > ingresos_total * 2:
            obs.append({
                "categoria": "Inventarios",
                "nivel": "AVISO",
                "descripcion": (
                    f"Inventario total (${inventario:,.2f}) supera 2x los ingresos del último mes (${ingresos_total:,.2f}). "
                    "Evaluar rotación de inventario."
                ),
            })

    except Exception as e:
        obs.append({"categoria": "Balance", "nivel": "AVISO", "descripcion": f"Error en análisis de balance: {e}"})

    # ── 5. CONCENTRACIÓN DE COSTOS (solo secciones nivel 2, > 10%) ──────────
    try:
        mes_cols_t = [c for c in df_mes.columns if "total" in c.lower()]
        total_col2 = mes_cols_t[0] if mes_cols_t else ([c for c in df_mes.columns if c not in ("Cod","Concepto")][-1])

        total_costos = _find(df_mes, "5", total_col2)
        if total_costos == 0:
            total_costos = _val(df_mes, "5", total_col2)
        for _, row in df_mes.iterrows():
            cod = str(row["Cod"])
            concepto = str(row["Concepto"])
            # Solo secciones de nivel 2 (ej: 5.1.1, 5.1.2, 5.2.1.1)
            if not cod.startswith("5") or cod.count(".") != 2:
                continue
            val = float(row[total_col2])
            if total_costos > 0 and val / total_costos > 0.10:
                obs.append({
                    "categoria": "Concentración de Costos",
                    "nivel": "INFO",
                    "descripcion": (
                        f"Cuenta {cod} ({concepto}): ${val:,.2f} representa el "
                        f"{val/total_costos*100:.1f}% del total de costos y gastos."
                    ),
                })
    except Exception as e:
        obs.append({"categoria": "Concentración", "nivel": "AVISO", "descripcion": f"Error: {e}"})

    # ── 6. ANÁLISIS POR CENTRO DE COSTO ────────────────────────────────────
    try:
        cc_cols_all = [c for c in df_cc.columns if c not in ("Cod", "Concepto")]
        ingreso_row = df_cc[df_cc["Cod"].astype(str) == "4"]
        if not ingreso_row.empty and cc_cols_all:
            for cc in cc_cols_all:
                ing_cc = float(ingreso_row[cc].iloc[0])
                if ing_cc != 0:
                    # Costo total de ese CC
                    cos_cc = df_cc[df_cc["Cod"].astype(str).str.startswith("5")][cc].sum()
                    margen = (ing_cc - cos_cc) / ing_cc * 100
                    if margen < 0:
                        obs.append({
                            "categoria": "Rentabilidad por Centro de Costo",
                            "nivel": "ALERTA",
                            "descripcion": (
                                f"Centro de Costo '{cc}' tiene margen NEGATIVO: {margen:.1f}% "
                                f"(Ingresos: ${ing_cc:,.2f} | Costos: ${cos_cc:,.2f})."
                            ),
                        })
    except Exception as e:
        obs.append({"categoria": "Centros de Costo", "nivel": "AVISO", "descripcion": f"Error: {e}"})

    return obs
