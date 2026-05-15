"""
Carga y une los archivos CarteraPorPagar y Personas para el módulo de pagos.
"""

import io
import pandas as pd


def _to_bytes(file_input):
    if hasattr(file_input, "read"):
        return file_input.read()
    return file_input


def _engine_for(raw_bytes):
    """Detecta el motor pandas según la firma del archivo."""
    return "xlrd" if raw_bytes[:4] == b"\xd0\xcf\x11\xe0" else "openpyxl"


def _find_col(columns, *keywords):
    """Devuelve el primer nombre de columna que contenga TODOS los keywords dados."""
    kw_lower = [k.lower() for k in keywords]
    for col in columns:
        col_lower = col.lower().replace(" ", "")
        if all(k.replace(" ", "") in col_lower for k in kw_lower):
            return col
    return None


def _clean_num(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", ".").str.replace(" ", ""),
        errors="coerce",
    ).fillna(0.0)


# ─── Cartera por Pagar ────────────────────────────────────────────────────────

def cargar_cartera(file_input):
    """
    Lee CarteraPorPagar (.xls/.xlsx).
    Estructura: 4 filas de metadata, encabezados en fila 4, datos desde fila 5.
    Devuelve DataFrame filtrado solo a filas FAC con columna 'Bucket' añadida.
    """
    raw = _to_bytes(file_input)
    engine = _engine_for(raw)
    df_raw = pd.read_excel(io.BytesIO(raw), header=None, engine=engine)

    # Encabezados en fila 4 (índice 4)
    headers = [str(h).strip() for h in df_raw.iloc[4].tolist()]
    df = df_raw.iloc[5:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    # Normalizar columnas numéricas
    buckets = ["Por vencer", "30 días", "60 días", "90 días", "120 días", "> 120 días"]
    extra   = ["Total", "Valor documento", "Retenciones", "Pagos"]
    for nombre in buckets + extra:
        col = _find_col(df.columns, nombre) or nombre
        if col in df.columns:
            df[col] = _clean_num(df[col])

    # Determinar bucket de aging (mutually exclusive: solo uno tiene valor por fila)
    df["Bucket"] = "Por vencer"
    for b in ["30 días", "60 días", "90 días", "120 días", "> 120 días"]:
        col = _find_col(df.columns, b) or b
        if col in df.columns:
            df.loc[df[col] > 0, "Bucket"] = b

    # Columna auxiliar: ¿está vencida?
    vencido_buckets = ["30 días", "60 días", "90 días", "120 días", "> 120 días"]
    df["Vencida"] = False
    for b in vencido_buckets:
        col = _find_col(df.columns, b) or b
        if col in df.columns:
            df.loc[df[col] > 0, "Vencida"] = True

    df = df.reset_index(drop=True)
    return df


# ─── Personas ─────────────────────────────────────────────────────────────────

def cargar_personas(file_input):
    """
    Lee Personas (.xls/.xlsx).
    Estructura: 3 filas de metadata, encabezados en fila 3, datos desde fila 4.
    Devuelve DataFrame con columnas bancarias relevantes.
    """
    raw = _to_bytes(file_input)
    engine = _engine_for(raw)
    df_raw = pd.read_excel(io.BytesIO(raw), header=None, engine=engine)

    # Encabezados en fila 3 (índice 3)
    headers = [str(h).strip() for h in df_raw.iloc[3].tolist()]
    df = df_raw.iloc[4:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    # Mantener solo columnas útiles para pagos
    needed = ["Razón Social", "Cédula", "RUC", "Banco", "N° Cta. Bco.", "Tipo Cta. Bco."]
    present = [c for c in needed if c in df.columns]
    df = df[present].copy()

    # Limpiar strings
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip().replace({"nan": "", "None": ""})

    return df


# ─── Join ─────────────────────────────────────────────────────────────────────

def merge_cartera_personas(df_cartera, df_personas):
    """
    LEFT JOIN cartera ↔ personas por Razón Social (normalizado a mayúsculas).
    Añade columnas calculadas: ID_PAGO, TIPO_CTA, NRO_CUENTA.
    """
    # Columna de join en cartera
    rs_cartera = _find_col(df_cartera.columns, "raz", "social") or "Razón Social"
    rs_personas = _find_col(df_personas.columns, "raz", "social") or "Razón Social"

    df_cartera = df_cartera.copy()
    df_personas = df_personas.copy()

    df_cartera["_jk"] = df_cartera[rs_cartera].astype(str).str.strip().str.upper()
    df_personas["_jk"] = df_personas[rs_personas].astype(str).str.strip().str.upper()

    # Desduplicar personas por join key (queda el primer registro con datos bancarios)
    df_personas = (
        df_personas
        .sort_values("N° Cta. Bco." if "N° Cta. Bco." in df_personas.columns else df_personas.columns[0],
                     ascending=False, na_position="last")
        .drop_duplicates(subset=["_jk"], keep="first")
    )

    # Para el merge renombramos la col de personas para evitar colisión
    df_personas = df_personas.rename(columns={rs_personas: "_pers_rs"})

    df = df_cartera.merge(df_personas, on="_jk", how="left")
    df = df.drop(columns=["_jk"], errors="ignore")

    # ID de pago: cédula si existe, sino RUC
    ced_col = _find_col(df.columns, "cédula") or _find_col(df.columns, "cedula") or "Cédula"
    ruc_col = _find_col(df.columns, "ruc") or "RUC"

    def get_id(row):
        ced = str(row.get(ced_col, "")).strip()
        ruc = str(row.get(ruc_col, "")).strip()
        if ced and ced not in ("", "nan", "None"):
            return ced
        if ruc and ruc not in ("", "nan", "None"):
            return ruc
        return ""

    df["ID_PAGO"] = df.apply(get_id, axis=1)

    # Tipo de cuenta: Ahorro → aho, Corriente → cte
    tipo_col = _find_col(df.columns, "tipo", "cta") or "Tipo Cta. Bco."
    if tipo_col in df.columns:
        def map_tipo(v):
            v = str(v).strip().lower()
            if "ahorro" in v:
                return "aho"
            if "corriente" in v:
                return "cte"
            return ""
        df["TIPO_CTA"] = df[tipo_col].apply(map_tipo)
    else:
        df["TIPO_CTA"] = ""

    # Número de cuenta
    nro_col = _find_col(df.columns, "n°", "cta") or _find_col(df.columns, "cta", "bco") or "N° Cta. Bco."
    if nro_col in df.columns:
        df["NRO_CUENTA"] = df[nro_col].astype(str).str.strip().replace({"nan": "", "None": ""})
    else:
        df["NRO_CUENTA"] = ""

    df = df.reset_index(drop=True)
    return df
