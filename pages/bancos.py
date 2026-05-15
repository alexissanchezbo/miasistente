import streamlit as st
from modules.theme import render_toggle, apply_theme

render_toggle()
apply_theme()

st.markdown('<div class="mod-eyebrow">Mi Asistente · Suite Financiera</div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">🏦 Mi Asistente de Bancos</div>', unsafe_allow_html=True)
st.markdown('<div class="gold-line"></div>', unsafe_allow_html=True)

tema = st.session_state.get("tema", "oscuro")
bg  = "#161616" if tema == "oscuro" else "#FFFFFF"
brd = "#2A2A2A" if tema == "oscuro" else "#E0D9CC"
txt = "#999"    if tema == "oscuro" else "#666"
h2  = "#F0F0F0" if tema == "oscuro" else "#1A1A2E"

st.markdown(f"""
<div style="background:{bg}; border:1px solid {brd}; border-left:4px solid #C9A84C;
            border-radius:8px; padding:48px 40px; color:{txt};">
    <h2 style="color:{h2}; margin-bottom:16px;">🚧 En desarrollo</h2>
    <p>Conciliación bancaria y gestión de cuentas. Estará disponible próximamente.</p>
    <br>
    <b style="color:#C9A84C">Funcionalidades planificadas:</b>
    <ul style="margin-top:12px;">
        <li style="margin-bottom:8px; font-size:14px;">Conciliación bancaria automática (extracto vs. mayor contable)</li>
        <li style="margin-bottom:8px; font-size:14px;">Identificación de partidas en tránsito y notas de débito/crédito</li>
        <li style="margin-bottom:8px; font-size:14px;">Seguimiento de saldos bancarios por cuenta</li>
        <li style="margin-bottom:8px; font-size:14px;">Reporte de diferencias clasificadas por tipo</li>
        <li style="margin-bottom:8px; font-size:14px;">Histórico de conciliaciones por período y banco</li>
    </ul>
</div>
""", unsafe_allow_html=True)

st.markdown("---")
st.caption("© 2026 Alexis Sánchez · Mi Asistente de Bancos")
