"""
Mi Asistente de Pagos
Carga CarteraPorPagar (por sesión) + Maestro de Personas (persistente).
Permite seleccionar facturas y exportar un Excel con Resumen y Detalle.
"""

import json
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from modules.theme import render_toggle, apply_theme
from modules.loader_pagos import cargar_cartera, cargar_personas, merge_cartera_personas
from modules.exporter_pagos import exportar_pagos, exportar_txt_pichincha

render_toggle()
apply_theme()

# ── Rutas de almacenamiento persistente ───────────────────────────────────────
_APP_DIR        = Path(__file__).parent.parent
_DATA_DIR       = _APP_DIR / "data"
_PERSONAS_PKL   = _DATA_DIR / "personas_maestro.pkl"
_PERSONAS_META  = _DATA_DIR / "personas_maestro_meta.json"
_HISTORIAL_JSON = _DATA_DIR / "historial_pagos.json"
_DATA_DIR.mkdir(exist_ok=True)

# ── Carpeta donde se guardan los archivos de pago generados ───────────────────
_PAGOS_DIR = Path(r"C:\Users\CONTABILIDAD\OneDrive\Escritorio\ALEXIS\PAGOS\PAGOS REALIZADOS")
_PAGOS_DIR.mkdir(parents=True, exist_ok=True)


# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* .mod-eyebrow / .main-title / .sub-title / .gold-line
       vienen de CARDS_DARK / CARDS_LIGHT en theme.py — no redefinir aquí */

    /* Badges de estado del maestro de personas */
    .badge-ok   { background:#1a3a1a; color:#4CAF50; padding:4px 10px;
                  border-radius:6px; font-size:12px; font-weight:600; }
    .badge-warn { background:#3a2a0a; color:#FFC107; padding:4px 10px;
                  border-radius:6px; font-size:12px; font-weight:600; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="mod-eyebrow">Mi Asistente · Suite Financiera</div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">💳 Mi Asistente de Pagos</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Selección y exportación de pagos a proveedores</div>',
            unsafe_allow_html=True)
st.markdown('<div class="gold-line"></div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS — Maestro de Personas persistente
# ══════════════════════════════════════════════════════════════════════════════
def _personas_guardadas() -> bool:
    return _PERSONAS_PKL.exists()


def _load_personas_disco() -> pd.DataFrame | None:
    if _PERSONAS_PKL.exists():
        try:
            return pd.read_pickle(str(_PERSONAS_PKL))
        except Exception:
            return None
    return None


def _save_personas_disco(df: pd.DataFrame, filename: str):
    _DATA_DIR.mkdir(exist_ok=True)
    df.to_pickle(str(_PERSONAS_PKL))
    meta = {
        "filename": filename,
        "rows": len(df),
        "saved_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    _PERSONAS_META.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")


def _delete_personas_disco():
    if _PERSONAS_PKL.exists():
        _PERSONAS_PKL.unlink()
    if _PERSONAS_META.exists():
        _PERSONAS_META.unlink()


def _get_meta() -> dict | None:
    if _PERSONAS_META.exists():
        try:
            return json.loads(_PERSONAS_META.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None


# ── Historial de pagos ────────────────────────────────────────────────────────
def _cargar_historial() -> list:
    if _HISTORIAL_JSON.exists():
        try:
            return json.loads(_HISTORIAL_JSON.read_text(encoding="utf-8"))
        except Exception:
            return []
    return []


def _guardar_en_historial(df_export: "pd.DataFrame", col_prov: str, col_doc: str):
    """Añade un registro de pago al historial JSON persistente."""
    historial = _cargar_historial()
    now = datetime.now()

    detalle = []
    for _, row in df_export.iterrows():
        detalle.append({
            "proveedor": str(row.get(col_prov, "")).strip(),
            "documento":  str(row.get(col_doc,  "")).strip(),
            "valor":      float(row.get("_VALOR_PAGO", 0) or 0),
        })

    total     = sum(d["valor"] for d in detalle)
    n_provs   = len({d["proveedor"] for d in detalle})

    _DIAS_ES  = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    dia_str   = _DIAS_ES[now.weekday()]
    semana    = now.isocalendar()[1]

    _slug = now.strftime("%Y%m%d_%H%M")
    registro = {
        "id":            now.strftime("%Y%m%d_%H%M%S"),
        "fecha_str":     f"{dia_str} {now.strftime('%d/%m/%Y')}",
        "semana":        f"Semana {semana}",
        "timestamp":     now.isoformat(),
        "total":         round(total, 2),
        "n_registros":   len(detalle),
        "n_proveedores": n_provs,
        "detalle":       detalle,
        "ruta":          str(_PAGOS_DIR / now.strftime("%Y-%m-%d")),
        "archivos":      [f"pago_{_slug}.xlsx", f"pichincha_{_slug}.txt"],
    }
    historial.append(registro)
    _HISTORIAL_JSON.write_text(
        json.dumps(historial, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ── Helpers generales ─────────────────────────────────────────────────────────
def _col(df, *keywords):
    kw = [k.lower().replace(" ", "") for k in keywords]
    for c in df.columns:
        cl = c.lower().replace(" ", "")
        if all(k in cl for k in kw):
            return c
    return None


def _safe_key(text):
    return re.sub(r"[^a-zA-Z0-9]", "_", str(text))[:50]


def _toggle_prov(prov_key, idxs):
    val = st.session_state[prov_key]
    for i in idxs:
        st.session_state[f"chk_pago_{i}"] = val


def _clear_checks():
    for k in list(st.session_state.keys()):
        if k.startswith("chk_pago_") or k.startswith("chk_prov_") or k.startswith("amt_pago_"):
            del st.session_state[k]
    st.session_state.pagos_manuales = []


# ── Estado de sesión ──────────────────────────────────────────────────────────
if "pagos_df"         not in st.session_state: st.session_state.pagos_df         = None
if "pagos_cargado"    not in st.session_state: st.session_state.pagos_cargado    = False
if "pers_replacing"   not in st.session_state: st.session_state.pers_replacing   = False
if "pagos_manuales"   not in st.session_state: st.session_state.pagos_manuales   = []
if "pagos_archivos"   not in st.session_state: st.session_state.pagos_archivos   = None


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN A — CONFIGURACIÓN DEL MAESTRO DE PERSONAS (persistente)
# ══════════════════════════════════════════════════════════════════════════════
personas_ok = _personas_guardadas()
config_label = (
    "⚙️ Maestro de Personas · ✅ Activo"
    if personas_ok and not st.session_state.pers_replacing
    else "⚙️ Maestro de Personas · ⚠️ Sin configurar"
)

with st.expander(config_label, expanded=not personas_ok or st.session_state.pers_replacing):

    if personas_ok and not st.session_state.pers_replacing:
        # ── Maestro YA cargado ────────────────────────────────────────────
        meta = _get_meta()
        if meta:
            c1, c2, c3 = st.columns([2, 1.5, 1.5])
            c1.markdown(f"**📄 Archivo:** `{meta.get('filename','—')}`")
            c2.markdown(f"**👥 Registros:** {meta.get('rows', '—'):,}")
            c3.markdown(f"**🕐 Guardado:** {meta.get('saved_at','—')}")
        else:
            st.markdown("✅ Maestro de Personas cargado en memoria.")

        st.markdown("")
        if st.button("🔄 Reemplazar maestro de Personas", type="secondary"):
            st.session_state.pers_replacing = True
            # Invalidar datos cargados para forzar recarga con el nuevo maestro
            st.session_state.pagos_df      = None
            st.session_state.pagos_cargado = False
            _clear_checks()
            st.rerun()

    else:
        # ── Sin maestro / modo reemplazo ─────────────────────────────────
        if st.session_state.pers_replacing:
            st.info("🔄 Sube el nuevo Maestro de Personas para reemplazar el actual.")
        else:
            st.warning(
                "⚠️ **Maestro de Personas no configurado.** "
                "Súbelo una sola vez y quedará guardado para todas las sesiones."
            )

        f_personas = st.file_uploader(
            "Maestro de Personas (.xls / .xlsx)",
            type=["xls", "xlsx"],
            key="up_personas_config",
            help="Exportado del sistema — contiene datos bancarios de proveedores",
        )

        col_btn1, col_btn2 = st.columns([1, 3])
        with col_btn1:
            guardar = st.button(
                "💾 Guardar maestro",
                type="primary",
                disabled=(f_personas is None),
            )
        with col_btn2:
            if st.session_state.pers_replacing:
                if st.button("✖ Cancelar", type="secondary"):
                    st.session_state.pers_replacing = False
                    st.rerun()

        if guardar and f_personas:
            with st.spinner("Procesando y guardando maestro…"):
                try:
                    df_pers = cargar_personas(f_personas)
                    _save_personas_disco(df_pers, f_personas.name)
                    st.session_state.pers_replacing = False
                    st.success(
                        f"✅ Maestro guardado: {len(df_pers):,} proveedores "
                        f"desde `{f_personas.name}`"
                    )
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error al procesar el archivo: {e}")
                    import traceback; st.code(traceback.format_exc())

# Si no hay maestro configurado, detener aquí
if not _personas_guardadas():
    st.info("👆 Configura primero el Maestro de Personas para continuar.")
    st.caption("© 2026 Alexis Sánchez · Mi Asistente de Pagos")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN B — CARGA DE CARTERA POR PAGAR (por sesión)
# ══════════════════════════════════════════════════════════════════════════════
with st.expander("📁 Cargar Cartera por Pagar", expanded=not st.session_state.pagos_cargado):
    st.caption("La Cartera se carga por sesión. El Maestro de Personas ya está configurado arriba.")

    f_cartera = st.file_uploader(
        "Cartera por Pagar (.xls / .xlsx)",
        type=["xls", "xlsx"],
        key="up_cartera",
        help="Exportado del sistema contable — facturas pendientes de pago",
    )

    if st.button("🔄 Cargar y cruzar datos", type="primary", disabled=(f_cartera is None)):
        with st.spinner("Procesando…"):
            try:
                df_cartera  = cargar_cartera(f_cartera)
                df_personas = _load_personas_disco()
                if df_personas is None:
                    st.error("❌ No se pudo leer el Maestro de Personas guardado.")
                    st.stop()
                df_merged = merge_cartera_personas(df_cartera, df_personas)
                st.session_state.pagos_df      = df_merged
                st.session_state.pagos_cargado = True
                _clear_checks()
                col_rs = _col(df_merged, "raz", "social") or df_merged.columns[0]
                st.success(
                    f"✅ {len(df_merged)} facturas cargadas de "
                    f"{df_merged[col_rs].nunique()} proveedores."
                )
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error al procesar los archivos: {e}")
                import traceback; st.code(traceback.format_exc())

if not st.session_state.pagos_cargado or st.session_state.pagos_df is None:
    st.info("👆 Sube la Cartera por Pagar y haz clic en **Cargar y cruzar datos** para continuar.")
    st.caption("© 2026 Alexis Sánchez · Mi Asistente de Pagos")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# PREPARAR DATOS
# ══════════════════════════════════════════════════════════════════════════════
df = st.session_state.pagos_df.copy()

COL_PROV  = _col(df, "raz", "social")  or df.columns[0]
COL_DOC   = _col(df, "#", "doc")       or _col(df, "documento") or df.columns[1]
COL_EMIS  = _col(df, "emis")           or df.columns[2]
COL_VENC  = _col(df, "venc")           or df.columns[3]
COL_TOTAL = _col(df, "total")          or df.columns[4]
COL_DESC  = (
    _col(df, "glosa")
    or _col(df, "descrip")
    or _col(df, "concepto")
    or _col(df, "detalle")
)

df["_venc_dt"] = pd.to_datetime(df[COL_VENC], errors="coerce")


# ══════════════════════════════════════════════════════════════════════════════
# FILTROS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("### 🔍 Filtros")
col_f1, col_f2, col_f3, col_f4 = st.columns([1.1, 1.3, 2, 1])

with col_f1:
    solo_vencidas = st.toggle("⚠️ Solo facturas vencidas", value=True)

with col_f2:
    fecha_corte = st.date_input(
        "📅 Fecha de corte",
        value=date.today(),
        help=(
            "Facturas con F. Vencimiento ≤ esta fecha se consideran vencidas. "
            "También define los días mostrados en el estado de cada factura."
        ),
    )

with col_f3:
    texto_buscar = st.text_input("Buscar proveedor", placeholder="Nombre o parte del nombre…")

with col_f4:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🔄 Nueva carga", help="Vuelve a la carga de la cartera"):
        st.session_state.pagos_cargado = False
        st.session_state.pagos_df      = None
        _clear_checks()
        st.rerun()

fecha_corte_ts = pd.Timestamp(fecha_corte)

df_filtrado = df.copy()
if solo_vencidas:
    df_filtrado = df_filtrado[df_filtrado["_venc_dt"] <= fecha_corte_ts]
if texto_buscar.strip():
    mask = df_filtrado[COL_PROV].astype(str).str.upper().str.contains(
        texto_buscar.strip().upper(), na=False
    )
    df_filtrado = df_filtrado[mask]


# ══════════════════════════════════════════════════════════════════════════════
# MÉTRICAS EN TIEMPO REAL
# ══════════════════════════════════════════════════════════════════════════════
def get_seleccionados():
    return [i for i in df_filtrado.index if st.session_state.get(f"chk_pago_{i}", False)]


st.markdown("---")
st.markdown("### 📋 Selección de facturas")

m1, m2, m3 = st.columns(3)
sel_indices = get_seleccionados()
total_sel   = sum(
    st.session_state.get(f"amt_pago_{i}", float(df_filtrado.loc[i, COL_TOTAL] or 0))
    for i in sel_indices
)
n_prov_sel  = len(set(str(df_filtrado.loc[i, COL_PROV]) for i in sel_indices))

m1.metric("💰 Total a pagar",          f"${total_sel:,.2f}")
m2.metric("📄 Facturas seleccionadas", str(len(sel_indices)))
m3.metric("🏢 Proveedores",            str(n_prov_sel))
st.markdown("---")

if df_filtrado.empty:
    st.warning("No se encontraron facturas con los filtros aplicados.")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# TABLA CON CHECKBOXES POR PROVEEDOR
# ══════════════════════════════════════════════════════════════════════════════
for proveedor in sorted(df_filtrado[COL_PROV].unique()):
    grupo         = df_filtrado[df_filtrado[COL_PROV] == proveedor]
    indices_grupo = list(grupo.index)
    n_sel_grupo   = sum(1 for i in indices_grupo if st.session_state.get(f"chk_pago_{i}", False))
    total_grupo   = sum(
        st.session_state.get(f"amt_pago_{i}", float(grupo.loc[i, COL_TOTAL] or 0))
        for i in indices_grupo
    )
    all_sel       = bool(indices_grupo) and n_sel_grupo == len(indices_grupo)

    prov_label = (
        f"**{proveedor}**  "
        f"· {len(indices_grupo)} factura{'s' if len(indices_grupo) != 1 else ''}  "
        f"· ${total_grupo:,.2f}  "
        f"· ✅ {n_sel_grupo}/{len(indices_grupo)} seleccionadas"
    )

    # ── Checkbox de proveedor (fuera del expander, sin necesidad de expandir) ─
    prov_key = f"chk_prov_{_safe_key(proveedor)}"
    st.session_state[prov_key] = all_sel  # sincronizar con estado individual

    col_chk, col_exp = st.columns([0.04, 0.96])

    with col_chk:
        st.checkbox(
            "",
            key=prov_key,
            on_change=_toggle_prov,
            args=(prov_key, indices_grupo),
            label_visibility="collapsed",
            help=f"Seleccionar / deseleccionar todas las facturas de {proveedor}",
        )

    with col_exp:
        with st.expander(prov_label, expanded=(n_sel_grupo > 0)):

            # Cabecera — 7 columnas
            _W = [0.45, 3.1, 1.5, 1.5, 1.9, 1.15, 1.7]
            hdr = st.columns(_W)
            hdr[0].markdown("**✓**")
            hdr[1].markdown("**# Factura · Descripción**")
            hdr[2].markdown("**F. Emisión**")
            hdr[3].markdown("**F. Vencimiento**")
            hdr[4].markdown("**Estado**")
            hdr[5].markdown("**Saldo doc.**")
            hdr[6].markdown("**✏️ Valor a pagar**")
            st.markdown('<hr style="margin:4px 0; opacity:0.2">', unsafe_allow_html=True)

            for i in indices_grupo:
                row_data  = df_filtrado.loc[i]
                total_val = float(row_data.get(COL_TOTAL, 0) or 0)
                cols = st.columns(_W)

                with cols[0]:
                    st.checkbox("", key=f"chk_pago_{i}", label_visibility="collapsed")

                # Nro. factura + descripción con tooltip
                doc_val   = str(row_data.get(COL_DOC, "")).strip()
                desc_raw  = str(row_data.get(COL_DESC, "")).strip() if COL_DESC else ""
                desc_text = "" if desc_raw in ("nan", "None", "") else desc_raw

                if desc_text:
                    short = desc_text[:48] + "…" if len(desc_text) > 48 else desc_text
                    desc_safe = (
                        desc_text
                        .replace("&", "&amp;")
                        .replace('"', "&quot;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;")
                    )
                    cols[1].markdown(
                        f'<code>{doc_val}</code>'
                        f'<br><span style="font-size:10px;color:#999;cursor:help" '
                        f'title="{desc_safe}">{short}</span>',
                        unsafe_allow_html=True,
                    )
                else:
                    cols[1].markdown(f"`{doc_val}`")

                cols[2].markdown(str(row_data.get(COL_EMIS, "")).split(" ")[0])
                cols[3].markdown(str(row_data.get(COL_VENC, "")).split(" ")[0])

                # Estado basado en fecha de corte vs fecha de vencimiento
                venc_dt = row_data.get("_venc_dt")
                if pd.notna(venc_dt):
                    dias = (fecha_corte_ts - venc_dt).days
                    if dias < 0:
                        color, estado = "#27AE60", f"Vence en {-dias}d"
                    elif dias == 0:
                        color, estado = "#F39C12", "Vence hoy"
                    elif dias <= 30:
                        color, estado = "#F39C12", f"Vencida {dias}d"
                    elif dias <= 90:
                        color, estado = "#E67E22", f"Vencida {dias}d"
                    else:
                        color, estado = "#C0392B", f"Vencida {dias}d"
                else:
                    color, estado = "#888", str(row_data.get("Bucket", "—"))

                cols[4].markdown(
                    f'<span style="color:{color};font-size:12px">⬤</span> '
                    f'<span style="font-size:11px">{estado}</span>',
                    unsafe_allow_html=True,
                )

                # Saldo original (solo lectura, referencia)
                cols[5].markdown(
                    f'<span style="font-size:12px;color:#888">${total_val:,.2f}</span>',
                    unsafe_allow_html=True,
                )

                # Valor a pagar (editable) — usa total_val como default si aún no se editó
                with cols[6]:
                    st.number_input(
                        "",
                        key=f"amt_pago_{i}",
                        min_value=0.0,
                        value=float(st.session_state.get(f"amt_pago_{i}", total_val)),
                        step=0.01,
                        format="%.2f",
                        label_visibility="collapsed",
                    )


# ══════════════════════════════════════════════════════════════════════════════
# PAGOS MANUALES ADICIONALES
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("### ✍️ Pagos manuales adicionales")
st.caption("Agrega transferencias que no están en la cartera (anticipos, pagos directos, etc.)")

# Maestro de Personas en memoria para lookup bancario silencioso
_df_pers_ref = _load_personas_disco()

# ── Construir lista de nombres del Maestro para el selectbox ──────────────────
_MANUAL_LIBRE = "✏️  Escribir nombre manualmente…"
_col_rs_pers  = None
_nombres_maestro: list[str] = []
if _df_pers_ref is not None:
    _col_rs_pers = next(
        (c for c in _df_pers_ref.columns if "social" in c.lower()),
        _df_pers_ref.columns[0],
    )
    _nombres_maestro = sorted(
        _df_pers_ref[_col_rs_pers]
        .astype(str).str.strip()
        .replace({"nan": "", "None": ""})
        .loc[lambda s: s != ""]
        .unique()
        .tolist()
    )

_opciones_sb = [_MANUAL_LIBRE] + _nombres_maestro

# ── Formulario ─────────────────────────────────────────────────────────────────
with st.form("form_pago_manual", clear_on_submit=True):
    fc1, fc2, fc3 = st.columns([2, 2.5, 1])

    with fc1:
        _nombre_input = st.text_input(
            "Nombre",
            placeholder="Escribe el nombre del beneficiario…",
            help="Se buscará automáticamente en el Maestro de Personas para completar los datos bancarios.",
        )
        # Sugerencias en tiempo real (solo informativas, dentro del form)
        if _nombre_input.strip() and _nombres_maestro:
            _sugs = [n for n in _nombres_maestro if _nombre_input.strip().upper() in n.upper()][:5]
            if _sugs:
                st.caption("💡 " + "  ·  ".join(_sugs))

    with fc2:
        _desc = st.text_input(
            "Descripción",
            max_chars=40,
            placeholder="Ej. ANTICIPO CONTRATO JUNIO",
        )

    with fc3:
        _valor_m = st.number_input("Valor ($)", min_value=0.01, value=1.00, step=0.01, format="%.2f")

    _agregar = st.form_submit_button("➕ Agregar", type="primary", use_container_width=True)

    if _agregar:
        _nombre = _nombre_input.strip()
        _err = []
        if not _nombre:
            _err.append("Debes indicar el nombre.")
        if not _desc.strip():
            _err.append("Debes ingresar una descripción.")
        if _err:
            for e in _err:
                st.error(e)
        else:
            # Buscar datos bancarios en el Maestro (silencioso)
            _id_p, _tcta, _nrocta, _banco_m = "", "", "", ""
            if _df_pers_ref is not None and _col_rs_pers is not None:
                _match = _df_pers_ref[
                    _df_pers_ref[_col_rs_pers].astype(str).str.strip().str.upper()
                    == _nombre.upper()
                ]
                if not _match.empty:
                    _rp = _match.iloc[0]
                    _ced = str(_rp.get("Cédula", "")).strip()
                    _ruc = str(_rp.get("RUC",    "")).strip()
                    _id_p = _ced if _ced not in ("", "nan", "None") else (
                             _ruc if _ruc not in ("", "nan", "None") else "")
                    _tv = str(_rp.get("Tipo Cta. Bco.", "")).lower()
                    _tcta = "AHO" if "ahorro" in _tv else "CTE" if "corriente" in _tv else ""
                    _nrocta = str(_rp.get("N° Cta. Bco.", "")).strip()
                    if _nrocta in ("nan", "None"): _nrocta = ""
                    _banco_m = str(_rp.get("Banco", "")).strip()
                    if _banco_m in ("nan", "None"): _banco_m = ""

            st.session_state.pagos_manuales.append({
                "prov":       _nombre,
                "desc":       _desc.strip()[:40],
                "valor":      _valor_m,
                "id_pago":    _id_p,
                "tipo_cta":   _tcta,
                "nro_cuenta": _nrocta,
                "banco":      _banco_m,
            })
            st.rerun()

# ── Lista de pagos manuales agregados ────────────────────────────────────────
if st.session_state.pagos_manuales:
    total_manual = sum(p["valor"] for p in st.session_state.pagos_manuales)

    hm = st.columns([2.8, 2.8, 1.2, 1.2, 0.6])
    hm[0].markdown("**Beneficiario**")
    hm[1].markdown("**Descripción / Referencia**")
    hm[2].markdown("**Datos bancarios**")
    hm[3].markdown("**Valor**")
    hm[4].markdown("")
    st.markdown('<hr style="margin:2px 0; opacity:0.2">', unsafe_allow_html=True)

    for _idx, _p in enumerate(st.session_state.pagos_manuales):
        _cm = st.columns([2.8, 2.8, 1.2, 1.2, 0.6])
        _cm[0].markdown(f"**{_p['prov']}**")
        _cm[1].markdown(f"_{_p['desc']}_")
        _banco_info = f"{_p['tipo_cta']} · {_p['nro_cuenta']}" if _p["nro_cuenta"] else "⚠️ Sin cuenta"
        _cm[2].markdown(f'<span style="font-size:11px;color:#aaa">{_banco_info}</span>',
                        unsafe_allow_html=True)
        _cm[3].markdown(f"**${_p['valor']:,.2f}**")
        if _cm[4].button("🗑️", key=f"del_man_{_idx}", help="Eliminar este pago manual"):
            st.session_state.pagos_manuales.pop(_idx)
            st.rerun()

    st.markdown(
        f'<div style="text-align:right;padding:4px 8px;background:#1a2a1a;border-radius:4px;'
        f'font-size:13px;color:#4CAF50">➕ Subtotal pagos manuales: <strong>${total_manual:,.2f}</strong></div>',
        unsafe_allow_html=True,
    )
    st.markdown("")
else:
    st.caption("Aún no has agregado pagos manuales.")


# ══════════════════════════════════════════════════════════════════════════════
# EXPORTAR
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
sel_indices_final = get_seleccionados()

_manuales     = st.session_state.pagos_manuales
_hay_cartera  = bool(sel_indices_final)
_hay_manuales = bool(_manuales)

if not _hay_cartera and not _hay_manuales:
    st.info("Selecciona al menos una factura o agrega un pago manual para habilitar la exportación.")
    st.session_state.pagos_archivos = None  # invalidar archivos previos
else:
    total_cartera    = sum(
        st.session_state.get(f"amt_pago_{i}", float(df_filtrado.loc[i, COL_TOTAL] or 0))
        for i in sel_indices_final
    )
    total_manual_exp = sum(p["valor"] for p in _manuales)
    total_exp        = total_cartera + total_manual_exp

    resumen_partes = []
    if _hay_cartera:
        resumen_partes.append(f"{len(sel_indices_final)} factura(s) de cartera")
    if _hay_manuales:
        resumen_partes.append(f"{len(_manuales)} pago(s) manual(es)")
    st.markdown(f"**{' + '.join(resumen_partes)} — Total combinado: ${total_exp:,.2f}**")

    # ── Construir df_export (no genera archivos aún) ───────────────────────
    if _hay_cartera:
        df_export = df.loc[sel_indices_final].copy()
        df_export["_VALOR_PAGO"] = [
            float(st.session_state.get(f"amt_pago_{i}", float(df_export.loc[i, COL_TOTAL] or 0)))
            for i in sel_indices_final
        ]
    else:
        df_export = pd.DataFrame()

    if _hay_manuales:
        filas_m = []
        for _pm in _manuales:
            filas_m.append({
                COL_PROV:      _pm["prov"],
                COL_DOC:       _pm["desc"],
                "_VALOR_PAGO": _pm["valor"],
                "ID_PAGO":     _pm["id_pago"],
                "TIPO_CTA":    _pm["tipo_cta"],
                "NRO_CUENTA":  _pm["nro_cuenta"],
                "Banco":       _pm["banco"],
                "_REFERENCIA": _pm["desc"],
            })
        df_export = pd.concat([df_export, pd.DataFrame(filas_m)], ignore_index=True)

    # ── Botón "Generar" — guarda en historial y prepara los archivos ───────
    if st.button("📤 Generar archivos de pago", type="primary"):
        with st.spinner("Generando archivos y registrando en historial…"):
            _excel_buf  = exportar_pagos(df_export)
            _txt_bytes  = exportar_txt_pichincha(df_export)
            _guardar_en_historial(df_export, COL_PROV, COL_DOC)

            _n_con = sum(
                1 for _, row in df_export.iterrows()
                if str(row.get("NRO_CUENTA", "")).strip() not in ("", "nan", "None")
            )
            _fecha_slug = datetime.now().strftime("%Y%m%d_%H%M")

            # ── Guardar físicamente en la carpeta de pagos realizados ──────
            try:
                _pagos_dir_fecha = _PAGOS_DIR / datetime.now().strftime("%Y-%m-%d")
                _pagos_dir_fecha.mkdir(parents=True, exist_ok=True)

                _excel_path = _pagos_dir_fecha / f"pago_{_fecha_slug}.xlsx"
                _txt_path   = _pagos_dir_fecha / f"pichincha_{_fecha_slug}.txt"

                _excel_path.write_bytes(_excel_buf.getvalue())
                _txt_path.write_bytes(_txt_bytes)
                _ruta_guardado = str(_pagos_dir_fecha)
            except Exception as _e:
                _ruta_guardado = None
                st.warning(f"⚠️ No se pudo guardar en disco: {_e}")

            st.session_state.pagos_archivos = {
                "excel":        _excel_buf.getvalue(),
                "txt":          _txt_bytes,
                "n_con_banco":  _n_con,
                "n_sin_banco":  len(df_export) - _n_con,
                "fecha_slug":   _fecha_slug,
                "ruta":         _ruta_guardado,
            }
        if st.session_state.pagos_archivos.get("ruta"):
            st.success(
                f"✅ Archivos guardados en:\n`{st.session_state.pagos_archivos['ruta']}`"
            )
        else:
            st.success("✅ Archivos generados y registrados en historial")

    # ── Botones de descarga (desde session_state) ──────────────────────────
    if st.session_state.pagos_archivos:
        _arch      = st.session_state.pagos_archivos
        _slug      = _arch["fecha_slug"]
        _n_con     = _arch["n_con_banco"]
        _n_sin     = _arch["n_sin_banco"]

        btn1, btn2 = st.columns(2)
        with btn1:
            st.download_button(
                label="📥 Descargar Excel de pago",
                data=_arch["excel"],
                file_name=f"pago_{_slug}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
            st.caption("Resumen de pago + Detalle de transferencias")

        with btn2:
            st.download_button(
                label="🏦 Descargar TXT Pichincha",
                data=_arch["txt"],
                file_name=f"pichincha_{_slug}.txt",
                mime="text/plain",
                type="primary",
                use_container_width=True,
                disabled=(_n_con == 0),
            )
            if _n_con > 0:
                st.caption(
                    f"Formato Cash Management · {_n_con} transferencia{'s' if _n_con != 1 else ''}"
                    + (f" · ⚠️ {_n_sin} sin cuenta bancaria" if _n_sin > 0 else "")
                )
            else:
                st.caption("⚠️ Ningún proveedor seleccionado tiene cuenta bancaria registrada")


# ══════════════════════════════════════════════════════════════════════════════
# HISTORIAL DE PAGOS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("### 📊 Historial de Pagos")
st.caption("Se guarda automáticamente cada vez que generas los archivos de pago.")

_historial = _cargar_historial()

if not _historial:
    st.info("Aún no hay pagos registrados. Aparecerán aquí cada vez que generes un pago.")
else:
    # ── Métricas globales ──────────────────────────────────────────────────
    _total_hist  = sum(p["total"]       for p in _historial)
    _total_regs  = sum(p["n_registros"] for p in _historial)
    hm1, hm2, hm3 = st.columns(3)
    hm1.metric("📋 Pagos registrados",  str(len(_historial)))
    hm2.metric("💰 Total acumulado",    f"${_total_hist:,.2f}")
    hm3.metric("📄 Registros totales",  str(_total_regs))
    st.markdown("")

    # ── Detalle por cada pago (más reciente primero) ───────────────────────
    for _pr in reversed(_historial):
        _lbl = (
            f"📅 {_pr['fecha_str']}  ·  {_pr.get('semana', '')}  "
            f"·  {_pr['n_registros']} registro(s)  "
            f"·  {_pr['n_proveedores']} proveedor(es)  "
            f"·  **${_pr['total']:,.2f}**"
        )
        with st.expander(_lbl, expanded=False):
            # Ruta en disco
            if _pr.get("ruta"):
                _archs = "  ·  ".join(_pr.get("archivos", []))
                st.caption(f"📁 `{_pr['ruta']}`  —  {_archs}")

            _det = _pr.get("detalle", [])
            if _det:
                _df_det = pd.DataFrame(_det)
                _df_det.columns = ["Proveedor", "Documento / Descripción", "Valor ($)"]
                _df_det["Valor ($)"] = _df_det["Valor ($)"].map(lambda v: f"${v:,.2f}")
                st.dataframe(_df_det, use_container_width=True, hide_index=True)

            if st.button(
                "🗑️ Eliminar este registro",
                key=f"del_hist_{_pr['id']}",
                type="secondary",
            ):
                _h2 = [p for p in _cargar_historial() if p["id"] != _pr["id"]]
                _HISTORIAL_JSON.write_text(
                    json.dumps(_h2, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                st.rerun()

    # ── Borrar todo ────────────────────────────────────────────────────────
    st.markdown("")
    if st.button("🗑️ Limpiar todo el historial", type="secondary"):
        if _HISTORIAL_JSON.exists():
            _HISTORIAL_JSON.unlink()
        st.rerun()

st.markdown("---")
st.caption("© 2026 Alexis Sánchez · Mi Asistente de Pagos")
