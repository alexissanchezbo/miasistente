import streamlit as st
from modules.theme import render_toggle, apply_theme

render_toggle()
apply_theme()

# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="padding:48px 0 28px 0; text-align:center;">
    <div class="hero-eyebrow">Suite Financiera &amp; Contable</div>
    <div class="hero-name">Mi<span>Asistente</span></div>
    <div class="hero-product">by Sánchez</div>
    <div class="hero-tagline">Analiza. Decide. Controla.</div>
</div>
<div class="gold-line-center"></div>
<div class="intro-text">
    Herramientas de análisis financiero y contable diseñadas para tomar
    decisiones con claridad y confianza.
</div>
<div class="section-label">Módulos</div>
""", unsafe_allow_html=True)

# ── Tarjetas ──────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2, gap="medium")

with col1:
    st.markdown("""
    <div class="card card-active">
        <div class="card-icon">📊</div>
        <div class="card-title">Mi Asistente EEFF</div>
        <div class="card-desc">
            Estado de Resultados (P&G) con estructura NIIF/NIC.<br>
            Análisis por mes, proyecto y centro de costo.<br>
            Prorrateo automático · Observaciones · Excel descargable.
        </div>
        <span class="badge-ready">Disponible</span>
    </div>
    """, unsafe_allow_html=True)
    if st.button("→ Abrir módulo", key="go_eeff", type="primary", use_container_width=True):
        st.switch_page("pages/eeff.py")

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    st.markdown("""
    <div class="card card-coming">
        <div class="card-icon">📋</div>
        <div class="card-title">Mi Asistente Tributario</div>
        <div class="card-desc">
            IVA, retenciones y declaración de Impuesto a la Renta.<br>
            Gastos no deducibles · Calendario tributario con alertas.
        </div>
        <span class="badge-wip">Próximamente</span>
    </div>
    """, unsafe_allow_html=True)
    if st.button("→ Ver módulo", key="go_trib", use_container_width=True):
        st.switch_page("pages/tributario.py")

with col2:
    st.markdown("""
    <div class="card card-coming">
        <div class="card-icon">💳</div>
        <div class="card-title">Mi Asistente de Pagos</div>
        <div class="card-desc">
            Programación y seguimiento de pagos a proveedores.<br>
            Flujo de caja por vencimientos · Alertas de pagos críticos.
        </div>
        <span class="badge-wip">Próximamente</span>
    </div>
    """, unsafe_allow_html=True)
    if st.button("→ Ver módulo", key="go_pagos", use_container_width=True):
        st.switch_page("pages/pagos.py")

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    st.markdown("""
    <div class="card card-coming">
        <div class="card-icon">🏦</div>
        <div class="card-title">Mi Asistente de Bancos</div>
        <div class="card-desc">
            Conciliación bancaria automática contra el mayor contable.<br>
            Seguimiento de saldos · Partidas en tránsito · Diferencias.
        </div>
        <span class="badge-wip">Próximamente</span>
    </div>
    """, unsafe_allow_html=True)
    if st.button("→ Ver módulo", key="go_bancos", use_container_width=True):
        st.switch_page("pages/bancos.py")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
    MiAsistente v1.0 &nbsp;·&nbsp; <span>© 2026 Alexis Sánchez</span>
    &nbsp;·&nbsp; Todos los derechos reservados
</div>
""", unsafe_allow_html=True)
