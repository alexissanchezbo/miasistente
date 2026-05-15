import streamlit as st

st.set_page_config(
    page_title="MiAsistente · Sánchez",
    page_icon="⚡",
    layout="wide",
)

pages = st.navigation(
    {
        "": [
            st.Page("home.py", title="Inicio", icon="⚡", default=True),
        ],
        "Módulos": [
            st.Page("pages/eeff.py",       title="Mi Asistente EEFF",        icon="📊"),
            st.Page("pages/pagos.py",      title="Mi Asistente de Pagos",    icon="💳"),
            st.Page("pages/tributario.py", title="Mi Asistente Tributario",  icon="📋"),
            st.Page("pages/bancos.py",     title="Mi Asistente de Bancos",   icon="🏦"),
        ],
    }
)

pages.run()
