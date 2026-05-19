import streamlit as st
import traceback
from datetime import date, datetime
from modules.loader import load_pyg_mes, load_pyg_cc, load_mayor, load_balance, load_cartera, load_activos
from modules.depreciacion import calcular_dep_mensual_por_cuenta, inyectar_en_df_mes, resumen_depreciacion
from modules.builder_mes import build_pyg_mes
from modules.builder_proyecto import build_pyg_proyecto
from modules.builder_cc import build_pyg_cc
from modules.observaciones import generar_observaciones
from modules.exporter import exportar_excel
from modules.theme import render_toggle, apply_theme

render_toggle()
apply_theme()

st.markdown("""
<style>
    /* .stApp / .mod-eyebrow / .main-title / .sub-title / .gold-line
       vienen de apply_theme() en theme.py — no redefinir aquí */
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="mod-eyebrow">Mi Asistente · Suite Financiera</div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">📊 Mi Asistente EEFF</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Consolidación P&G con estructura NIIF/NIC — BAEC BALDOSAS DEL ECUADOR S.A.S.</div>',
            unsafe_allow_html=True)
st.markdown('<div class="gold-line"></div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PANEL DE CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════════
with st.expander("⚙️ Configuración del reporte", expanded=True):

    col_emp, col_per = st.columns([2, 1])
    with col_emp:
        empresa = st.text_input(
            "Nombre de la empresa",
            value="BAEC BALDOSAS DEL ECUADOR S.A.S.",
        )
        periodo_desc = st.text_input(
            "Descripción del período",
            value="Enero – Abril 2026",
            help="Ej: Q1 2026 / Enero–Marzo 2026 / Semestre I 2026",
        )

    with col_per:
        st.markdown("**Rango de fechas para proyectos**")
        st.caption("Filtra los mayores para que coincidan con el P&G por Mes")
        fecha_ini = st.date_input(
            "Desde",
            value=date(2026, 1, 1),
            min_value=date(2020, 1, 1),
            max_value=date(2030, 12, 31),
        )
        fecha_fin = st.date_input(
            "Hasta",
            value=date(2026, 4, 30),
            min_value=date(2020, 1, 1),
            max_value=date(2030, 12, 31),
        )
        if fecha_fin < fecha_ini:
            st.error("⚠️ La fecha 'Hasta' debe ser posterior a 'Desde'.")

# ══════════════════════════════════════════════════════════════════════════════
# PANEL DE ARCHIVOS
# ══════════════════════════════════════════════════════════════════════════════
with st.expander("📁 Archivos fuente", expanded=True):
    st.caption("Sube los 5 archivos requeridos del sistema contable · el 6to es opcional para el análisis de cartera")

    col1, col2 = st.columns(2)
    with col1:
        f_mes = st.file_uploader(
            "1. P&G por Mes · Estado de Resultados mensual (.xls)",
            type=["xls", "xlsx"],
            key="f_mes",
        )
        f_ing = st.file_uploader(
            "2. Mayor de Ingresos · Reporte 4 (.xlsx / .xls)",
            type=["xls", "xlsx"],
            key="f_ing",
            help="Mayor de cuentas 4.x con columnas Fecha, Código, Proyecto, Debe, Haber",
        )
        f_cos = st.file_uploader(
            "3. Mayor de Costos y Gastos · Reporte 5 (.xlsx / .xls)",
            type=["xls", "xlsx"],
            key="f_cos",
            help="Mayor de cuentas 5.x con columnas Fecha, Código, Proyecto, Debe, Haber",
        )
    with col2:
        f_cc = st.file_uploader(
            "4. P&G por Centro de Costo (.xlsx)",
            type=["xls", "xlsx"],
            key="f_cc",
        )
        f_bg = st.file_uploader(
            "5. Balance General · Estado de Situación Financiera (.xls)",
            type=["xls", "xlsx"],
            key="f_bg",
        )
        f_trx = st.file_uploader(
            "6. Cartera por Cobrar · Facturas Pendientes (.xls)  ⬅ opcional",
            type=["xls", "xlsx"],
            key="f_trx",
            help="Reporte 'Cartera por Cobrar' del sistema contable. Habilita la pestaña Recuperación de Cartera con aging por cliente y por proyecto.",
        )
        f_activos = st.file_uploader(
            "7. Activos Fijos · Registro de Depreciación (.xls)  ⬅ opcional",
            type=["xls", "xlsx"],
            key="f_activos",
            help="Exporta desde Contifico → Activos Fijos → Listado. "
                 "Permite calcular la depreciación mensual devengada y distribuirla en el P&G por mes.",
        )

# ══════════════════════════════════════════════════════════════════════════════
# BOTÓN DE GENERACIÓN
# ══════════════════════════════════════════════════════════════════════════════
archivos_requeridos = [f_mes, f_ing, f_cos, f_cc, f_bg]
todos_listos = all(archivos_requeridos)

if not todos_listos:
    faltantes = sum(1 for f in archivos_requeridos if not f)
    st.info(f"⬆️ Faltan {faltantes} archivo(s) requerido(s) para habilitar la generación.")
if f_trx:
    st.success("✅ Cartera por Cobrar cargada — se generará la pestaña **Recuperación de Cartera**")
if f_activos:
    st.success("✅ Activos Fijos cargados — la depreciación mensual se distribuirá en el P&G por Mes")

generar = st.button(
    "⚡ Generar Estados Financieros",
    disabled=not todos_listos or fecha_fin < fecha_ini,
    type="primary",
    use_container_width=False,
)

# ══════════════════════════════════════════════════════════════════════════════
# PROCESO
# ══════════════════════════════════════════════════════════════════════════════
if generar and todos_listos:
    prog = st.progress(0, "Iniciando…")
    status = st.empty()

    try:
        prog.progress(8, "Cargando P&G por Mes…")
        df_mes = load_pyg_mes(f_mes)

        prog.progress(20, "Cargando Mayor de Ingresos…")
        df_ing = load_mayor(f_ing)

        prog.progress(35, "Cargando Mayor de Costos y Gastos…")
        df_cos = load_mayor(f_cos)

        prog.progress(48, "Cargando P&G por Centro de Costo…")
        df_cc = load_pyg_cc(f_cc)

        prog.progress(55, "Cargando Balance General…")
        df_bg = load_balance(f_bg)

        df_trx_data = None
        if f_trx:
            prog.progress(58, "Cargando Cartera por Cobrar…")
            df_trx_data = load_cartera(f_trx)

        # ── Depreciación mensual devengada (archivo opcional) ─────────────
        dep_resumen = []
        dep_no_encontradas = []
        if f_activos:
            prog.progress(60, "Calculando depreciación mensual devengada…")
            df_activos_data = load_activos(f_activos)
            dep_map = calcular_dep_mensual_por_cuenta(df_activos_data)
            n_meses_rep = len([c for c in df_mes.columns
                               if c not in ("Cod", "Concepto", "Total", "TOTAL")])
            dep_resumen = resumen_depreciacion(df_activos_data, dep_map, n_meses=n_meses_rep)
            df_mes, _inyectadas, dep_no_encontradas = inyectar_en_df_mes(df_mes, dep_map)

        prog.progress(62, "Construyendo P&G por Mes…")
        filas_mes, vcols_mes, _ = build_pyg_mes(df_mes)

        prog.progress(72, f"Construyendo P&G por Proyecto ({fecha_ini} → {fecha_fin})…")
        filas_proy, vcols_proy = build_pyg_proyecto(
            df_ing, df_cos, df_mes,
            fecha_inicio=datetime.combine(fecha_ini, datetime.min.time()),
            fecha_fin=datetime.combine(fecha_fin, datetime.max.time()),
        )

        prog.progress(82, "Construyendo P&G por Centro de Costo…")
        filas_cc, vcols_cc = build_pyg_cc(df_cc)

        prog.progress(90, "Generando observaciones…")
        obs = generar_observaciones(df_mes, df_cc, df_ing, df_bg)

        prog.progress(96, "Generando Excel…")
        import pandas as pd
        df_mayor_all = pd.concat([df_ing, df_cos], ignore_index=True)

        buf = exportar_excel(
            empresa             = empresa,
            filas_mes           = filas_mes,
            value_cols_mes      = vcols_mes,
            filas_proyecto      = filas_proy,
            value_cols_proyecto = vcols_proy,
            filas_cc            = filas_cc,
            value_cols_cc       = vcols_cc,
            observaciones       = obs,
            df_balance          = df_bg,
            df_mayor_completo   = df_mayor_all,
            df_cartera          = df_trx_data,
            periodo_desc        = periodo_desc,
            titulo_mes          = f"Estado de Resultados Comparativo Mensual  |  {periodo_desc}",
            titulo_proyecto     = (
                f"Estado de Resultados por Proyecto  |  {periodo_desc}  "
                f"({fecha_ini.strftime('%d/%m/%Y')} – {fecha_fin.strftime('%d/%m/%Y')})"
            ),
            titulo_cc           = f"Estado de Resultados por Centro de Costo  |  {periodo_desc}",
            titulo_balance      = f"Estado de Situación Financiera  |  {periodo_desc}",
        )

        prog.progress(100, "✅ Listo")
        status.empty()

        st.success("✅ Estados Financieros generados correctamente")

        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("Cuentas P&G Mes",        len(df_mes))
        m2.metric("Líneas Mayor Ingresos",   len(df_ing))
        m3.metric("Líneas Mayor C&G",        len(df_cos))
        m4.metric("Proyectos detectados",    len([c for c in vcols_proy if c != "TOTAL"]))
        m5.metric("Observaciones",           len(obs))
        if df_trx_data is not None and not df_trx_data.empty:
            m6.metric("Facturas en cartera", len(df_trx_data))

        nombre_archivo = f"Estados_Financieros_{periodo_desc.replace(' ', '_').replace('–', '-')}.xlsx"

        col_dl, col_info = st.columns([1, 2])
        with col_dl:
            st.download_button(
                label="📥 Descargar Excel",
                data=buf.getvalue(),
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        with col_info:
            alertas = [o for o in obs if o["nivel"] == "ALERTA"]
            avisos  = [o for o in obs if o["nivel"] == "AVISO"]
            infos   = [o for o in obs if o["nivel"] == "INFO"]
            if alertas:
                st.error(f"🔴 {len(alertas)} alerta(s) crítica(s) — ver pestaña Observaciones")
            if avisos:
                st.warning(f"🟡 {len(avisos)} aviso(s) a revisar")
            if infos:
                st.info(f"🔵 {len(infos)} nota(s) informativa(s)")

        # ── Resumen de depreciación devengada ─────────────────────────────
        if dep_resumen:
            dep_total_mensual = sum(d["dep_mensual"] for d in dep_resumen)
            dep_total_ytd     = sum(d["dep_ytd"]     for d in dep_resumen)
            with st.expander(
                f"📉 Depreciación mensual devengada distribuida — "
                f"${dep_total_mensual:,.2f}/mes · YTD ${dep_total_ytd:,.2f}",
                expanded=False,
            ):
                df_dep_show = pd.DataFrame([{
                    "Categoría":    d["categoria"],
                    "Cuenta":       d["cuenta"],
                    "Activos":      d["activos_n"],
                    "Dep. Mensual": f"${d['dep_mensual']:,.2f}",
                    "YTD":          f"${d['dep_ytd']:,.2f}",
                } for d in dep_resumen])
                st.dataframe(df_dep_show, use_container_width=True, hide_index=True)
                if dep_no_encontradas:
                    st.warning(
                        f"⚠️ Las siguientes cuentas no se encontraron en el P&G por Mes "
                        f"y no pudieron inyectarse: {', '.join(dep_no_encontradas)}"
                    )

        proyectos_lista = [c for c in vcols_proy if c != "TOTAL"]
        if proyectos_lista:
            with st.expander(f"📋 {len(proyectos_lista)} proyectos detectados en el período"):
                for p in proyectos_lista:
                    st.markdown(f"• {p}")

        with st.expander("📋 Resumen de observaciones", expanded=bool(alertas)):
            for o in obs:
                icon = {"ALERTA": "🔴", "AVISO": "🟡", "INFO": "🔵"}.get(o["nivel"], "•")
                st.markdown(f"{icon} **[{o['categoria']}]** {o['descripcion']}")

    except Exception:
        prog.empty()
        status.empty()
        st.error("❌ Error al procesar los archivos")
        with st.expander("🔍 Detalle del error (para soporte técnico)"):
            st.code(traceback.format_exc())

st.markdown("---")
st.caption("© 2026 Alexis Sánchez · Mi Asistente EEFF")
