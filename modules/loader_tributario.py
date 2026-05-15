"""
Carga archivos SRI (TXT tab-separados) y el Excel ReporteComprasVentas.
"""

import io
import re
import pandas as pd


# ── Tipos de comprobante reconocidos ──────────────────────────────────────────
TIPO_FACTURA      = "Factura"
TIPO_RETENCION    = "Comprobante de Retención"
TIPO_NOTA_CREDITO = "Nota de Crédito"
TIPO_NOTA_DEBITO  = "Nota de Débito"
TIPO_LIQUIDACION  = "Liquidación de Compra de Bienes y Prestación de Servicios"

# Prefijos en el Excel No. Comprobante que hay que eliminar para normalizar
_PREFIJOS = re.compile(
    r"^(FAC|LIQ|LC|NC|ND|RET|REINT|COMP|LIQ\.?|LQC|ATS)\s*",
    re.IGNORECASE,
)


def _to_bytes(f):
    if hasattr(f, "read"):
        return f.read()
    return f


def normalizar_serie(serie: str) -> str:
    """Elimina prefijos del Excel y deja solo 'XXX-XXX-XXXXXXXXX'."""
    s = str(serie).strip()
    s = _PREFIJOS.sub("", s).strip()
    # Asegurar que los segmentos tengan el formato correcto (rellenar con ceros)
    partes = re.split(r"[-–]", s)
    if len(partes) == 3:
        try:
            return f"{int(partes[0]):03d}-{int(partes[1]):03d}-{int(partes[2]):09d}"
        except ValueError:
            pass
    return s


# ── SRI TXT ───────────────────────────────────────────────────────────────────

def cargar_sri_txt(file_input) -> pd.DataFrame:
    """
    Lee un TXT del SRI (tab-separado, encoding latin-1).
    Devuelve DataFrame con columnas normalizadas y columna _SERIE_NORM.
    """
    raw = _to_bytes(file_input)
    df  = pd.read_csv(
        io.BytesIO(raw),
        sep="\t",
        encoding="latin-1",
        dtype=str,
    ).fillna("")

    df.columns = [c.strip() for c in df.columns]

    # Normalizar serie del comprobante
    if "SERIE_COMPROBANTE" in df.columns:
        df["_SERIE_NORM"] = df["SERIE_COMPROBANTE"].apply(normalizar_serie)

    # Normalizar RUC emisor (quitar espacios)
    if "RUC_EMISOR" in df.columns:
        df["RUC_EMISOR"] = df["RUC_EMISOR"].str.strip()

    # Convertir importes
    for col in ("VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL"):
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].str.replace(",", ".").str.strip(), errors="coerce"
            ).fillna(0.0)

    # Tipo limpio
    if "TIPO_COMPROBANTE" in df.columns:
        df["TIPO_COMPROBANTE"] = df["TIPO_COMPROBANTE"].str.strip()

    return df


def detectar_tipo(df_sri: pd.DataFrame) -> str:
    """Devuelve el tipo predominante del DataFrame SRI."""
    if "TIPO_COMPROBANTE" in df_sri.columns:
        tipos = df_sri["TIPO_COMPROBANTE"].value_counts()
        if not tipos.empty:
            return tipos.index[0]
    return "Desconocido"


# ── Excel ReporteComprasVentas (2 pestañas: COMPRAS y VENTAS) ─────────────────

def _leer_hoja(xl: pd.ExcelFile, sheet: str) -> pd.DataFrame:
    """
    Lee una hoja del Excel buscando la fila de encabezado (contiene 'emisi' o 'ruc').
    Deduplica nombres de columna repetidos. Descarta filas sin RUC válido.
    """
    df_raw = xl.parse(sheet, header=None, dtype=str).fillna("")

    # Encontrar fila de encabezados: primera que tenga 'emisi' o 'ruc' en alguna celda
    hdr_row = 0
    for i, row in df_raw.iterrows():
        vals = [str(v).strip().lower() for v in row.tolist()]
        if any("emisi" in v or v == "ruc" for v in vals):
            hdr_row = i
            break

    raw_headers = [str(h).strip() for h in df_raw.iloc[hdr_row].tolist()]
    seen: dict[str, int] = {}
    headers: list[str] = []
    for h in raw_headers:
        if h in seen:
            seen[h] += 1
            headers.append(f"{h}.{seen[h]}")
        else:
            seen[h] = 0
            headers.append(h)

    df = df_raw.iloc[hdr_row + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    # Descartar filas sin RUC válido (totales, vacías, separadores)
    col_ruc = _col_like(df, "ruc")
    if col_ruc:
        df = df[df[col_ruc].astype(str).str.strip().str.len() >= 10].copy()

    return df.reset_index(drop=True)


def cargar_excel_reporte(file_input) -> dict:
    """
    Lee el Excel de compras/ventas esperando dos pestañas:
      - Una cuyo nombre contenga 'compras' (mayúsculas o minúsculas)
      - Una cuyo nombre contenga 'ventas'

    Devuelve dict con claves:
      'compras'  → DataFrame hoja COMPRAS
      'ventas'   → DataFrame hoja VENTAS
      '_debug'   → info de diagnóstico
    """
    raw = _to_bytes(file_input)

    with pd.ExcelFile(io.BytesIO(raw)) as xl:
        sheets = xl.sheet_names

        # Localizar pestañas por nombre
        sheet_compras = next(
            (s for s in sheets if "compra" in s.lower()), None)
        sheet_ventas  = next(
            (s for s in sheets if "venta"  in s.lower()), None)

        if sheet_compras is None:
            raise ValueError(
                f"No se encontró una pestaña con 'COMPRAS' en el nombre. "
                f"Pestañas disponibles: {sheets}"
            )
        if sheet_ventas is None:
            raise ValueError(
                f"No se encontró una pestaña con 'VENTAS' en el nombre. "
                f"Pestañas disponibles: {sheets}"
            )

        df_compras = _leer_hoja(xl, sheet_compras)
        df_ventas  = _leer_hoja(xl, sheet_ventas)

    # ── Normalizar COMPRAS ────────────────────────────────────────────────────
    col_doc = _col_like(df_compras, "comprobante")
    if col_doc:
        df_compras["_SERIE_NORM"] = df_compras[col_doc].apply(normalizar_serie)
    col_ruc = _col_like(df_compras, "ruc")
    if col_ruc:
        df_compras["_RUC"] = df_compras[col_ruc].astype(str).str.strip()
    col_ret = _col_like(df_compras, "retenci")
    if col_ret:
        df_compras["_SERIE_RET_NORM"] = df_compras[col_ret].apply(normalizar_serie)
    for c in df_compras.columns:
        if c in ("Subtotal", "IVA", "Total", "IVA Gasto"):
            df_compras[c] = pd.to_numeric(
                df_compras[c].astype(str).str.replace(",", "."), errors="coerce"
            ).fillna(0.0)

    # ── Normalizar VENTAS ─────────────────────────────────────────────────────
    col_doc_v = _col_like(df_ventas, "comprobante")
    if col_doc_v:
        df_ventas["_SERIE_NORM"] = df_ventas[col_doc_v].apply(normalizar_serie)
    col_ruc_v = _col_like(df_ventas, "ruc")
    if col_ruc_v:
        df_ventas["_RUC"] = df_ventas[col_ruc_v].astype(str).str.strip()

    # Detectar columnas de retención en VENTAS.
    # Prioridad: columnas con "ret" que NO sean de fecha (F., Fecha...).
    # Separar las de número de serie de las de autorización.
    def _es_ret_num(c: str) -> bool:
        nl = c.strip().lower()
        return ("ret" in nl
                and not nl.startswith(("f.", "fec", "fecha"))
                and "aut" not in nl and "autoriza" not in nl and "clave" not in nl
                and c.strip() not in ("", "nan", "None"))

    def _es_ret_auth(c: str) -> bool:
        nl = c.strip().lower()
        return ("ret" in nl
                and ("aut" in nl or "autoriza" in nl or "clave" in nl)
                and c.strip() not in ("", "nan", "None"))

    cols_ret_num  = [c for c in df_ventas.columns if _es_ret_num(c)]
    cols_ret_auth = [c for c in df_ventas.columns if _es_ret_auth(c)]

    if cols_ret_num:
        df_ventas["_COL_RET_SERIE"]  = cols_ret_num[0]
        df_ventas["_SERIE_RET_NORM"] = df_ventas[cols_ret_num[0]].apply(normalizar_serie)
    if cols_ret_auth:
        df_ventas["_COL_RET_AUTH"] = cols_ret_auth[0]

    _debug = {
        "sheet_compras":   sheet_compras,
        "sheet_ventas":    sheet_ventas,
        "compras_filas":   len(df_compras),
        "ventas_filas":    len(df_ventas),
        "ventas_cols":     list(df_ventas.columns),
        "ventas_cols_ret": cols_ret_num + cols_ret_auth,
    }

    return {"compras": df_compras, "ventas": df_ventas, "_debug": _debug}


def _col_like(df: pd.DataFrame, keyword: str) -> str | None:
    kw = keyword.lower()
    for c in df.columns:
        if kw in c.lower():
            return c
    return None
