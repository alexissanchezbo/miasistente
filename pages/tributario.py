"""
Mi Asistente Tributario — Cierre Mensual de Impuestos
Paso 1: Barrido  |  Paso 2: Formulario 103  |  Paso 3: Formulario 104
"""

import shutil
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from modules.theme import render_toggle, apply_theme
from modules.loader_tributario import (
    cargar_sri_txt, cargar_excel_reporte, detectar_tipo,
    TIPO_FACTURA, TIPO_RETENCION, TIPO_NOTA_CREDITO,
    TIPO_NOTA_DEBITO, TIPO_LIQUIDACION,
)
from modules.conciliador_tributario import conciliar
from modules.exporter_tributario import exportar_conciliacion

render_toggle()
apply_theme()

# ── Estilos del stepper ───────────────────────────────────────────────────────
st.markdown("""
<style>
.stepper-wrap { display:flex; align-items:center; gap:0; margin:18px 0 24px; }
.step-box {
    display:flex; flex-direction:column; align-items:center;
    min-width:160px; padding:10px 8px; border-radius:10px;
    border: 2px solid transparent; flex:1;
}
.step-box.done   { border-color:#2ECC71; background:rgba(46,204,113,.08); }
.step-box.active { border-color:#F0A500; background:rgba(240,165,0,.10); }
.step-box.locked { border-color:#555; background:rgba(100,100,100,.06); opacity:.55; }
.step-num {
    width:32px; height:32px; border-radius:50%; display:flex;
    align-items:center; justify-content:center;
    font-weight:700; font-size:15px; margin-bottom:4px;
}
.done   .step-num { background:#2ECC71; color:#fff; }
.active .step-num { background:#F0A500; color:#fff; }
.locked .step-num { background:#555;    color:#ccc; }
.step-label { font-size:12px; font-weight:600; text-align:center; }
.done   .step-label { color:#2ECC71; }
.active .step-label { color:#F0A500; }
.locked .step-label { color:#888; }
.step-sub { font-size:10px; color:#888; text-align:center; margin-top:2px; }
.step-connector { height:2px; flex:0 0 24px; background:#444; margin:0 4px; margin-bottom:14px; }
.done + .step-connector, .step-connector.done { background:#2ECC71; }
</style>
""", unsafe_allow_html=True)

# ── Encabezado ────────────────────────────────────────────────────────────────
st.markdown('<div class="mod-eyebrow">Mi Asistente · Suite Financiera</div>',
            unsafe_allow_html=True)
st.markdown('<div class="main-title">📋 Mi Asistente Tributario</div>',
            unsafe_allow_html=True)
st.markdown('<div class="sub-title">Cierre Mensual de Impuestos</div>',
            unsafe_allow_html=True)
st.markdown('<div class="gold-line"></div>', unsafe_allow_html=True)

_ETIQUETAS = {
    TIPO_FACTURA:      "🧾 Facturas",
    TIPO_RETENCION:    "📌 Retenciones",
    TIPO_NOTA_CREDITO: "🔵 Notas de Crédito",
    TIPO_NOTA_DEBITO:  "🔴 Notas de Débito",
    TIPO_LIQUIDACION:  "📄 Liquidaciones de Compra",
}

_MODULES_DIR = Path(__file__).parent.parent / "modules"


def _limpiar_todo():
    for k in ("trib_resultado", "trib_tipo", "trib_paso",
              "trib_barrido_ok", "trib_f103_ok"):
        st.session_state.pop(k, None)
    st.session_state.trib_uploader_key = st.session_state.get("trib_uploader_key", 0) + 1
    pycache = _MODULES_DIR / "__pycache__"
    if pycache.exists():
        shutil.rmtree(pycache, ignore_errors=True)


# ── Session state ─────────────────────────────────────────────────────────────
_ss = st.session_state
if "trib_resultado"    not in _ss: _ss.trib_resultado    = None
if "trib_tipo"         not in _ss: _ss.trib_tipo         = {}
if "trib_uploader_key" not in _ss: _ss.trib_uploader_key = 0
if "trib_paso"         not in _ss: _ss.trib_paso         = 1
if "trib_barrido_ok"   not in _ss: _ss.trib_barrido_ok   = False
if "trib_f103_ok"      not in _ss: _ss.trib_f103_ok      = False


# ── Stepper visual ────────────────────────────────────────────────────────────
def _step_class(n):
    if n < _ss.trib_paso:   return "done"
    if n == _ss.trib_paso:  return "active"
    return "locked"

def _step_icon(n):
    if n < _ss.trib_paso:  return "✓"
    return str(n)

_pasos = [
    (1, "Paso 1", "Barrido SRI"),
    (2, "Paso 2", "Formulario 103"),
    (3, "Paso 3", "Formulario 104"),
]

_html_steps = '<div class="stepper-wrap">'
for idx, (n, label, sub) in enumerate(_pasos):
    cls = _step_class(n)
    icon = _step_icon(n)
    _html_steps += f"""
    <div class="step-box {cls}">
      <div class="step-num">{icon}</div>
      <div class="step-label">{label}: {sub}</div>
    </div>"""
    if idx < len(_pasos) - 1:
        _html_steps += '<div class="step-connector"></div>'
_html_steps += "</div>"
st.markdown(_html_steps, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PASO 1 — BARRIDO
# ══════════════════════════════════════════════════════════════════════════════
if _ss.trib_paso == 1:

    with st.expander("📁 Archivos fuente", expanded=_ss.trib_resultado is None):
        st.caption(
            "Sube uno o varios archivos TXT del SRI y el Excel de compras/ventas. "
            "El sistema detecta automáticamente el tipo de cada TXT."
        )
        _uk = _ss.trib_uploader_key
        col_l, col_r = st.columns(2)
        with col_l:
            st.markdown("**Archivos TXT del SRI** *(uno o varios)*")
            f_sri_list = st.file_uploader(
                "TXT del SRI", type=["txt"], accept_multiple_files=True,
                key=f"up_sri_{_uk}",
                help="Facturas, Retenciones, Notas de Crédito/Débito, Liquidaciones",
                label_visibility="collapsed",
            )
        with col_r:
            st.markdown("**Excel ReporteComprasVentas**")
            f_excel = st.file_uploader(
                "ReporteComprasVentas.xlsx", type=["xls", "xlsx"],
                key=f"up_excel_trib_{_uk}", label_visibility="collapsed",
            )

        archivos_ok = bool(f_sri_list) and f_excel is not None
        if f_sri_list:
            st.caption(f"✅ {len(f_sri_list)} archivo(s) TXT cargados")
        if not archivos_ok:
            falt = []
            if not f_sri_list: falt.append("TXT del SRI")
            if not f_excel:    falt.append("Excel de compras/ventas")
            st.info(f"⬆️ Falta: {' y '.join(falt)}")

        btn_c1, btn_c2 = st.columns([3, 1])
        with btn_c2:
            if st.button("🗑️ Limpiar todo", use_container_width=True,
                         help="Reinicia todos los pasos"):
                _limpiar_todo(); st.rerun()
        with btn_c1:
            _procesar = st.button("🔄 Ejecutar barrido", type="primary",
                                  disabled=not archivos_ok, use_container_width=True)

        if _procesar:
            with st.spinner("Procesando…"):
                try:
                    excel_data = cargar_excel_reporte(f_excel)
                    df_compras = excel_data["compras"]
                    df_ventas  = excel_data["ventas"]
                    _dbg       = excel_data.get("_debug", {})

                    with st.expander("🔬 Diagnóstico carga Excel", expanded=False):
                        st.caption(
                            f"Pestaña COMPRAS: **{_dbg.get('sheet_compras','?')}** "
                            f"({_dbg.get('compras_filas','?')} filas)  |  "
                            f"Pestaña VENTAS: **{_dbg.get('sheet_ventas','?')}** "
                            f"({_dbg.get('ventas_filas','?')} filas)"
                        )
                        cols_ret_found = _dbg.get("ventas_cols_ret", [])
                        if cols_ret_found:
                            st.success(f"✅ Columnas retención VENTAS: **{cols_ret_found}**")
                        else:
                            st.error("❌ No se detectaron columnas de retención en VENTAS")
                            st.code("\n".join(
                                f"  [{i:02d}] {c}"
                                for i, c in enumerate(_dbg.get("ventas_cols", []))
                            ))
                        if cols_ret_found and not df_ventas.empty:
                            _muestra = df_ventas[cols_ret_found].head(10).fillna("").astype(str)
                            _muestra = _muestra[_muestra.apply(
                                lambda r: any(v.strip() not in ("", "nan", "0") for v in r),
                                axis=1)].head(5)
                            if not _muestra.empty:
                                st.dataframe(_muestra, use_container_width=True, hide_index=True)

                    resultados_por_tipo: dict[str, dict] = {}
                    tipo_label_map: dict[str, str] = {}
                    for f_sri in f_sri_list:
                        f_sri.seek(0)
                        df_sri = cargar_sri_txt(f_sri)
                        tipo   = detectar_tipo(df_sri)
                        label  = _ETIQUETAS.get(tipo, tipo)
                        if tipo in resultados_por_tipo:
                            df_prev = resultados_por_tipo[tipo]["_df_sri"]
                            df_sri  = pd.concat([df_prev, df_sri], ignore_index=True)
                        resultado = conciliar(df_sri, df_compras, tipo,
                                             df_excel_ventas=df_ventas)
                        resultado["_df_sri"] = df_sri
                        resultados_por_tipo[tipo] = resultado
                        tipo_label_map[tipo]       = label

                    _ss.trib_resultado = resultados_por_tipo
                    _ss.trib_tipo      = tipo_label_map
                    st.rerun()

                except Exception:
                    st.error("❌ Error al procesar los archivos")
                    with st.expander("🔍 Detalle del error"):
                        st.code(traceback.format_exc())

    # ── Resultados del barrido ────────────────────────────────────────────────
    if _ss.trib_resultado is not None:
        resultados_por_tipo: dict = _ss.trib_resultado
        tipo_label_map:      dict = _ss.trib_tipo

        tipos_lista = list(resultados_por_tipo.keys())
        tab_labels  = [tipo_label_map.get(t, t) for t in tipos_lista]
        tabs        = st.tabs(tab_labels)

        _total_advertencias = 0

        for tab, tipo in zip(tabs, tipos_lista):
            res    = resultados_por_tipo[tipo]
            label  = tipo_label_map.get(tipo, tipo)
            es_ret = tipo == TIPO_RETENCION

            df_conc = res.get("conciliados", pd.DataFrame())
            df_sri  = res.get("solo_sri",    pd.DataFrame())
            df_xl   = res.get("solo_excel",  pd.DataFrame())

            n_conc  = len(df_conc)
            n_sri   = len(df_sri)
            n_xl    = len(df_xl)
            n_total = n_conc + n_sri + n_xl
            _total_advertencias += n_sri + n_xl

            with tab:
                st.markdown(f"### {label}")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("📊 Total", str(n_total))
                m2.metric("✅ Conciliados", str(n_conc))
                m3.metric("⚠️ Solo SRI", str(n_sri),
                          delta=f"-{n_sri}" if n_sri else None, delta_color="inverse")
                m4.metric("🔍 Solo Excel", str(n_xl),
                          delta=f"-{n_xl}" if n_xl else None, delta_color="inverse")

                if not es_ret:
                    t_sri  = pd.to_numeric(df_sri.get("Total SRI",  pd.Series(dtype=float)), errors="coerce").sum()
                    t_xl_s = pd.to_numeric(df_xl.get("Total Excel", pd.Series(dtype=float)), errors="coerce").sum()
                    t_con  = pd.to_numeric(df_conc.get("Total SRI", pd.Series(dtype=float)), errors="coerce").sum()
                    v1, v2, v3 = st.columns(3)
                    v1.metric("💰 Total conciliado",    f"${t_con:,.2f}")
                    v2.metric("🔴 Sin registrar (SRI)", f"${t_sri:,.2f}")
                    v3.metric("🟡 No en SRI (Excel)",   f"${t_xl_s:,.2f}")

                st.markdown("---")
                sub1, sub2, sub3 = st.tabs([
                    f"⚠️ Solo en SRI ({n_sri})",
                    f"🔍 Solo en Excel ({n_xl})",
                    f"✅ Conciliados ({n_conc})",
                ])

                with sub1:
                    if df_sri.empty:
                        st.success("✅ Todos los documentos del SRI están registrados.")
                    else:
                        st.warning(f"**{n_sri} documento(s) no registrados en contabilidad.**"
                                   + ("" if es_ret else f"  Total: **${t_sri:,.2f}**"))
                        st.dataframe(df_sri.drop(columns=["Clave Acceso"], errors="ignore"),
                                     use_container_width=True, hide_index=True)

                with sub2:
                    if df_xl.empty:
                        st.success("✅ Todos los documentos del Excel constan en el SRI.")
                    else:
                        _lbl = f"  Total: **${t_xl_s:,.2f}**" if not es_ret else ""
                        st.warning(f"**{n_xl} documento(s) en Excel no constan en el SRI.**" + _lbl)
                        if not es_ret and "Origen" in df_xl.columns:
                            n_man = df_xl["Origen"].str.contains("Manual|Física", na=False).sum()
                            n_ele = df_xl["Origen"].str.contains("Electrónica",   na=False).sum()
                            c1, c2 = st.columns(2)
                            c1.caption(f"📄 **{n_man}** manual(es)/física(s)")
                            c2.caption(f"⚡ **{n_ele}** electrónica(s) — verificar")
                        st.dataframe(df_xl, use_container_width=True, hide_index=True)

                with sub3:
                    if df_conc.empty:
                        st.info("No hay documentos conciliados.")
                    else:
                        st.dataframe(df_conc.drop(columns=["Clave Acceso"], errors="ignore"),
                                     use_container_width=True, hide_index=True)

                st.markdown("---")
                fecha_slug = datetime.now().strftime("%Y%m%d")
                tipo_slug  = tipo.replace(" ", "_").replace("/", "_")[:20]
                buf = exportar_conciliacion(res, tipo)
                st.download_button(
                    label=f"📥 Descargar Excel — {label}",
                    data=buf,
                    file_name=f"conciliacion_{tipo_slug}_{fecha_slug}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
                st.caption("3 pestañas: Solo en SRI · Solo en Excel · Conciliados")

        # ── Panel de cierre del barrido ───────────────────────────────────────
        st.markdown("---")
        if _total_advertencias > 0:
            st.warning(
                f"⚠️ El barrido encontró **{_total_advertencias} partida(s) con diferencias**. "
                "Puedes avanzar al siguiente paso igualmente — las advertencias quedarán "
                "registradas como contexto del cierre."
            )
        else:
            st.success("✅ Barrido sin diferencias. Todo cuadra con el SRI.")

        col_av, col_nu = st.columns([2, 1])
        with col_av:
            if st.button("➡️ Confirmar barrido y avanzar al Formulario 103",
                         type="primary", use_container_width=True):
                _ss.trib_barrido_ok = True
                _ss.trib_paso = 2
                st.rerun()
        with col_nu:
            if st.button("🔄 Nuevo barrido", use_container_width=True):
                _limpiar_todo(); st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PASO 2 — FORMULARIO 103
# ══════════════════════════════════════════════════════════════════════════════
elif _ss.trib_paso == 2:

    st.markdown("### 📄 Formulario 103 — Retenciones en la Fuente")

    st.info(
        "Este paso utilizará los datos del barrido para asistir en la elaboración "
        "del Formulario 103. **Próximamente disponible.**",
        icon="🔧",
    )

    # Datos disponibles del barrido (para desarrollo futuro)
    _barrido = _ss.trib_resultado or {}
    _resumen_barrido = {
        tipo: {
            "conciliados": len(res.get("conciliados", [])),
            "solo_sri":    len(res.get("solo_sri",    [])),
            "solo_excel":  len(res.get("solo_excel",  [])),
        }
        for tipo, res in _barrido.items()
    }

    with st.expander("📊 Resumen del barrido disponible para F103", expanded=True):
        for tipo, cnts in _resumen_barrido.items():
            lbl = _ss.trib_tipo.get(tipo, tipo)
            st.caption(
                f"**{lbl}** — "
                f"Conciliados: {cnts['conciliados']} · "
                f"Solo SRI: {cnts['solo_sri']} · "
                f"Solo Excel: {cnts['solo_excel']}"
            )

    st.markdown("---")
    col_b, col_s = st.columns([1, 2])
    with col_b:
        if st.button("⬅️ Volver al barrido", use_container_width=True):
            _ss.trib_paso = 1; st.rerun()
    with col_s:
        if st.button("➡️ Avanzar al Formulario 104",
                     type="primary", use_container_width=True,
                     help="El F103 aún no está implementado — puedes avanzar igualmente"):
            _ss.trib_f103_ok = True
            _ss.trib_paso = 3
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# PASO 3 — FORMULARIO 104
# ══════════════════════════════════════════════════════════════════════════════
elif _ss.trib_paso == 3:

    st.markdown("### 📄 Formulario 104 — Declaración de IVA")

    st.info(
        "Este paso utilizará los datos del barrido y el F103 para asistir en la "
        "elaboración del Formulario 104. **Próximamente disponible.**",
        icon="🔧",
    )

    _barrido = _ss.trib_resultado or {}
    _resumen_barrido = {
        tipo: {
            "conciliados": len(res.get("conciliados", [])),
            "solo_sri":    len(res.get("solo_sri",    [])),
            "solo_excel":  len(res.get("solo_excel",  [])),
        }
        for tipo, res in _barrido.items()
    }

    with st.expander("📊 Resumen del barrido disponible para F104", expanded=True):
        for tipo, cnts in _resumen_barrido.items():
            lbl = _ss.trib_tipo.get(tipo, tipo)
            st.caption(
                f"**{lbl}** — "
                f"Conciliados: {cnts['conciliados']} · "
                f"Solo SRI: {cnts['solo_sri']} · "
                f"Solo Excel: {cnts['solo_excel']}"
            )

    st.markdown("---")
    col_b, col_r = st.columns([1, 2])
    with col_b:
        if st.button("⬅️ Volver al F103", use_container_width=True):
            _ss.trib_paso = 2; st.rerun()
    with col_r:
        if st.button("🔄 Iniciar nuevo cierre mensual",
                     use_container_width=True):
            _limpiar_todo(); st.rerun()


st.markdown("---")
st.caption("© 2026 Alexis Sánchez · Mi Asistente Tributario")
