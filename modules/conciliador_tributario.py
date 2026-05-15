"""
Lógica de conciliación entre documentos SRI y Excel contable.

Resultado por cada tipo de documento:
  - conciliados      : en SRI Y en Excel (con diferencia de valor si aplica)
  - solo_sri         : en SRI pero NO en Excel  (no registrados)
  - solo_excel       : en Excel pero NO en SRI  (registrados pero no constan en SRI)
"""

import re
import pandas as pd
from modules.loader_tributario import (
    TIPO_FACTURA, TIPO_RETENCION, TIPO_NOTA_CREDITO,
    TIPO_NOTA_DEBITO, TIPO_LIQUIDACION, _col_like,
)


def _col_exact(df: pd.DataFrame, name: str) -> str | None:
    """Retorna la columna cuyo nombre es exactamente `name` (case-insensitive)."""
    nl = name.strip().lower()
    for c in df.columns:
        if c.strip().lower() == nl:
            return c
    return None


def _key_sri(row) -> tuple:
    """Clave de conciliación SRI: (RUC_EMISOR, _SERIE_NORM)."""
    return (str(row.get("RUC_EMISOR", "")).strip(),
            str(row.get("_SERIE_NORM", "")).strip())


def _key_excel(row, ruc_col="_RUC", serie_col="_SERIE_NORM") -> tuple:
    return (str(row.get(ruc_col, "")).strip(),
            str(row.get(serie_col, "")).strip())


def _auth_norm(s: str) -> str:
    """
    Normaliza el número de autorización/clave de acceso para comparación cruzada.
    - Toma los primeros 35 chars (lo que el Excel almacena)
    - Fuerza el dígito de ambiente (posición 24, 0-indexed 23) a '2' (producción)
      porque el Excel siempre tiene producción y el SRI puede devolver '1' en test.
    - Resultado: clave comparable de 35 chars única por documento.
    """
    s = str(s).strip()[:35]
    if len(s) >= 24:
        s = s[:23] + "2" + s[24:]
    return s


# ── Conciliación facturas / liquidaciones / notas ────────────────────────────

def conciliar_facturas(df_sri: pd.DataFrame, df_excel: pd.DataFrame) -> dict:
    """
    Concilia facturas (o liquidaciones / notas) SRI vs sección COMPRAS del Excel.
    Compara importes: IMPORTE_TOTAL (SRI) vs Total (Excel).
    """
    col_total_xl  = _col_exact(df_excel, "Total")   or _col_like(df_excel, "total")   or "Total"
    col_subtot_xl = _col_exact(df_excel, "Subtotal") or _col_like(df_excel, "subtotal") or "Subtotal"
    col_auth_xl     = _col_like(df_excel, "autoriza") or ""

    # ── Índice 1: (RUC, SERIE_NORM) ───────────────────────────────────────────
    idx_serie: dict[tuple, int] = {}
    for i, row in df_excel.iterrows():
        k = _key_excel(row)
        if k[0] and k[1] and k not in idx_serie:
            idx_serie[k] = i

    # ── Índice 2: triple key (RUC, auth_norm_35, SERIE_NORM) ──────────────────
    # Más preciso que solo auth prefix; elimina falsa coincidencia cuando el mismo
    # RUC aparece como proveedor y cliente en el mismo período.
    idx_triple: dict[tuple, int] = {}
    if col_auth_xl:
        for i, row in df_excel.iterrows():
            ruc   = str(row.get("_RUC",        "")).strip()
            auth  = _auth_norm(str(row.get(col_auth_xl, "")))
            serie = str(row.get("_SERIE_NORM",  "")).strip()
            k3    = (ruc, auth, serie)
            if ruc and auth and k3 not in idx_triple:
                idx_triple[k3] = i

    conciliados  = []
    solo_sri     = []
    usados_excel = set()

    for _, row_sri in df_sri.iterrows():
        k     = _key_sri(row_sri)
        clave = str(row_sri.get("CLAVE_ACCESO", "")).strip()
        serie = str(row_sri.get("_SERIE_NORM",  "")).strip()

        # Intento 1: (RUC, SERIE_NORM) — fast and exact
        xl_idx = idx_serie.get(k)

        # Intento 2: (RUC, auth_norm_35, SERIE_NORM) — más preciso para edge cases
        if xl_idx is None:
            k3     = (k[0], _auth_norm(clave), serie)
            xl_idx = idx_triple.get(k3)

        if xl_idx is not None:
            usados_excel.add(xl_idx)
            row_xl = df_excel.loc[xl_idx]

            # Valores SRI
            base_sri  = float(row_sri.get("VALOR_SIN_IMPUESTOS", 0) or 0)
            iva_sri   = float(row_sri.get("IVA",                 0) or 0)
            total_sri = float(row_sri.get("IMPORTE_TOTAL",        0) or 0)

            # Valores Excel
            base_xl   = float(row_xl.get(col_subtot_xl, 0) or 0)
            total_xl  = float(row_xl.get(col_total_xl,  0) or 0)

            dif_base  = round(base_sri  - base_xl,  2)
            dif_total = round(total_sri - total_xl, 2)

            if dif_total == 0 and dif_base == 0:
                estado = "✅ OK"
            elif dif_total == 0:
                estado = f"⚠️ Base dif. ${dif_base:,.2f}"
            else:
                estado = f"⚠️ Total dif. ${dif_total:,.2f}"

            conciliados.append({
                "RUC Emisor":         row_sri.get("RUC_EMISOR", ""),
                "Razón Social":       row_sri.get("RAZON_SOCIAL_EMISOR", ""),
                "Tipo":               row_sri.get("TIPO_COMPROBANTE", ""),
                "Serie":              row_sri.get("SERIE_COMPROBANTE", ""),
                "Clave Acceso":       row_sri.get("CLAVE_ACCESO", ""),
                "Fecha Emisión SRI":  row_sri.get("FECHA_EMISION", ""),
                "Base SRI":           base_sri,
                "IVA SRI":            iva_sri,
                "Total SRI":          total_sri,
                "Base Excel":         base_xl,
                "Total Excel":        total_xl,
                "Dif. Base":          dif_base,
                "Dif. Total":         dif_total,
                "No. Comprobante XL": row_xl.get(_col_like(df_excel, "comprobante") or "", ""),
                "Estado":             estado,
            })
        else:
            solo_sri.append({
                "RUC Emisor":        row_sri.get("RUC_EMISOR", ""),
                "Razón Social":      row_sri.get("RAZON_SOCIAL_EMISOR", ""),
                "Tipo":              row_sri.get("TIPO_COMPROBANTE", ""),
                "Serie":             row_sri.get("SERIE_COMPROBANTE", ""),
                "Clave Acceso":      row_sri.get("CLAVE_ACCESO", ""),
                "Fecha Emisión SRI": row_sri.get("FECHA_EMISION", ""),
                "Total SRI":         float(row_sri.get("IMPORTE_TOTAL", 0) or 0),
            })

    solo_excel = []
    col_doc_xl = _col_like(df_excel, "comprobante") or ""
    col_ruc_xl = _col_like(df_excel, "ruc") or "RUC"
    col_rs_xl  = _col_like(df_excel, "social") or "Razón Social"
    col_em_xl  = _col_like(df_excel, "emisi") or "F. Emisión"

    for i, row_xl in df_excel.iterrows():
        if i not in usados_excel:
            auth_raw = str(row_xl.get(col_auth_xl, "") if col_auth_xl else "").strip()
            auth_digits = len(re.sub(r"\D", "", auth_raw))
            if auth_digits == 49:
                origen = "⚡ Electrónica (revisar)"
            elif 8 <= auth_digits <= 10:
                origen = "📄 Manual / Física"
            elif auth_digits == 0:
                origen = "❓ Sin autorización"
            else:
                origen = f"❓ Auth {auth_digits} dígitos"

            solo_excel.append({
                "RUC Emisor":        str(row_xl.get(col_ruc_xl, "")).strip(),
                "Razón Social":      str(row_xl.get(col_rs_xl,  "")).strip(),
                "No. Comprobante":   str(row_xl.get(col_doc_xl, "")).strip(),
                "Fecha Emisión":     str(row_xl.get(col_em_xl,  "")).strip(),
                "Total Excel":       float(row_xl.get(col_total_xl, 0) or 0),
                "No. Autorización":  auth_raw,
                "Origen":            origen,
            })

    return {
        "conciliados": pd.DataFrame(conciliados),
        "solo_sri":    pd.DataFrame(solo_sri),
        "solo_excel":  pd.DataFrame(solo_excel),
    }


# ── Conciliación retenciones ──────────────────────────────────────────────────

def conciliar_retenciones(df_sri: pd.DataFrame, df_ventas: pd.DataFrame) -> dict:
    """
    Concilia retenciones recibidas (SRI) vs sección VENTAS del Excel.

    El TXT del SRI tiene: RUC_EMISOR, SERIE_COMPROBANTE, CLAVE_ACCESO, FECHA_EMISION.
    El Excel VENTAS tiene: columna de serie de retención y columna de autorización de retención
    (detectadas dinámicamente por el loader según el encabezado).

    Match:
      Primario  → (RUC_EMISOR, SERIE_RET_NORM)
      Secundario → CLAVE_ACCESO normalizada a 35 chars (ambiente forzado a '2')
    """
    col_ruc_v = _col_like(df_ventas, "ruc")        or "RUC"
    col_rs_v  = _col_like(df_ventas, "social")      or "Razón Social"
    col_em_v  = _col_like(df_ventas, "emisi")       or "F. Emisión"
    col_doc_v = _col_like(df_ventas, "comprobante") or ""

    # Nombre real de la columna de serie de retención (guardado por loader)
    col_ret_serie = None
    col_ret_auth  = None
    if not df_ventas.empty:
        if "_COL_RET_SERIE" in df_ventas.columns:
            _v = df_ventas["_COL_RET_SERIE"].iloc[0]
            col_ret_serie = str(_v).strip() if str(_v) not in ("", "nan") else None
        if "_COL_RET_AUTH" in df_ventas.columns:
            _v = df_ventas["_COL_RET_AUTH"].iloc[0]
            col_ret_auth = str(_v).strip() if str(_v) not in ("", "nan") else None

    # Valores que indican "sin retención" — se excluyen del índice y de Solo-Excel
    _SIN_RET = {"", "nan", "0", "999-999-999999999"}

    # ── Índice 1: (RUC_CLIENTE, SERIE_RET_NORM) → lista de filas ─────────────
    # Una misma retención puede afectar varias facturas (mismo No. retención,
    # dos filas en Excel). Guardamos TODAS las filas para marcarlas conciliadas.
    idx_serie: dict[tuple, list[int]] = {}
    for i, row in df_ventas.iterrows():
        ruc  = str(row.get(col_ruc_v,       "")).strip()
        sret = str(row.get("_SERIE_RET_NORM","")).strip()
        if ruc and sret and sret not in _SIN_RET:
            idx_serie.setdefault((ruc, sret), []).append(i)

    # ── Índice 2: autorización de retención normalizada → lista de filas ─────
    idx_auth: dict[str, list[int]] = {}
    if col_ret_auth:
        for i, row in df_ventas.iterrows():
            auth = _auth_norm(str(row.get(col_ret_auth, "")).strip())
            if auth and len(re.sub(r"\D", "", auth)) >= 10:
                idx_auth.setdefault(auth, []).append(i)

    conciliados   = []
    solo_sri      = []
    usados_ventas = set()

    for _, row_sri in df_sri.iterrows():
        ruc_sri  = str(row_sri.get("RUC_EMISOR",  "")).strip()
        serie    = str(row_sri.get("_SERIE_NORM",  "")).strip()
        clave    = str(row_sri.get("CLAVE_ACCESO", "")).strip()
        fecha    = str(row_sri.get("FECHA_EMISION","")).strip()

        # Intento 1: (RUC, serie retención)
        xl_indices = idx_serie.get((ruc_sri, serie))

        # Intento 2: clave de acceso normalizada
        if not xl_indices:
            xl_indices = idx_auth.get(_auth_norm(clave))

        if xl_indices:
            # Marcar TODAS las filas Excel con esa retención (puede afectar varias facturas)
            usados_ventas.update(xl_indices)
            facturas_xl = ", ".join(
                str(df_ventas.loc[idx].get(col_doc_v, "")).strip()
                for idx in xl_indices
            )
            no_ret_xl = str(df_ventas.loc[xl_indices[0]].get(col_ret_serie, "")).strip() \
                        if col_ret_serie else ""
            n_fac = len(xl_indices)
            conciliados.append({
                "RUC Emisor (cliente)": ruc_sri,
                "Razón Social":         row_sri.get("RAZON_SOCIAL_EMISOR", ""),
                "Serie Ret. SRI":       row_sri.get("SERIE_COMPROBANTE", ""),
                "Fecha Emisión SRI":    fecha,
                "No. Ret. Excel":       no_ret_xl,
                "Fac(s). Venta Excel":  facturas_xl,
                "Estado":               "✅ OK" if n_fac == 1 else f"✅ OK ({n_fac} facturas)",
            })
        else:
            solo_sri.append({
                "RUC Emisor (cliente)": ruc_sri,
                "Razón Social":         row_sri.get("RAZON_SOCIAL_EMISOR", ""),
                "Serie Ret.":           row_sri.get("SERIE_COMPROBANTE", ""),
                "Clave Acceso":         clave,
                "Fecha Emisión SRI":    fecha,
                "Nota":                 "No encontrado en Excel (VENTAS)",
            })

    # Solo Excel: retenciones en VENTAS sin par en SRI
    solo_excel = []
    for i, row_xl in df_ventas.iterrows():
        if i in usados_ventas:
            continue
        ret_val = str(row_xl.get("_SERIE_RET_NORM", "")).strip()
        if ret_val in _SIN_RET:
            continue
        no_ret_xl = str(row_xl.get(col_ret_serie, "")).strip() if col_ret_serie else ret_val
        solo_excel.append({
            "RUC Emisor (cliente)": str(row_xl.get(col_ruc_v, "")).strip(),
            "Razón Social":         str(row_xl.get(col_rs_v,  "")).strip(),
            "No. Retención":        no_ret_xl,
            "No. Factura":          str(row_xl.get(col_doc_v, "")).strip(),
            "Fecha":                str(row_xl.get(col_em_v,  "")).strip(),
        })

    return {
        "conciliados": pd.DataFrame(conciliados),
        "solo_sri":    pd.DataFrame(solo_sri),
        "solo_excel":  pd.DataFrame(solo_excel),
    }


# ── Dispatcher ───────────────────────────────────────────────────────────────

def conciliar(df_sri: pd.DataFrame, df_excel_compras: pd.DataFrame,
              tipo: str, df_excel_ventas: pd.DataFrame | None = None) -> dict:
    """
    Enruta al conciliador correcto según el tipo de comprobante.
      - Retenciones recibidas → contra VENTAS (emitidas por clientes sobre facturas de venta)
      - Facturas / liquidaciones / notas → contra COMPRAS
    """
    if tipo == TIPO_RETENCION:
        df_contra = df_excel_ventas if df_excel_ventas is not None else df_excel_compras
        return conciliar_retenciones(df_sri, df_contra)
    else:
        return conciliar_facturas(df_sri, df_excel_compras)
