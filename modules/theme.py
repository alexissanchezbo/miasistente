import streamlit as st

DARK = """
<style>
/* ── Fondo y estructura ── */
.stApp { background: linear-gradient(160deg, #0D0D0D 0%, #1A1A1A 60%, #0F1923 100%) !important; }
section[data-testid="stSidebar"] { background: #0F0F0F !important; border-right: 1px solid #2A2A2A; }
section[data-testid="stSidebar"] * { color: #BBBBBB !important; }
section[data-testid="stSidebar"] a:hover { color: #C9A84C !important; }
section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"] { border-radius:6px; }
section[data-testid="stSidebar"] [aria-current="page"] { background:#1E1E1E !important; color:#C9A84C !important; }

/* ── Texto general ── */
p, li, span, label, div { color: #CCCCCC; }
h1, h2, h3, h4 { color: #F0F0F0 !important; }

/* ── Inputs / widgets ── */
input, textarea, [data-baseweb="input"] input, [data-baseweb="textarea"] textarea {
    background: #1E1E1E !important;
    color: #E8E8E8 !important;
    border: 1px solid #3A3A3A !important;
    border-radius: 6px !important;
}
[data-baseweb="input"]:focus-within, [data-baseweb="textarea"]:focus-within {
    border-color: #C9A84C !important;
    box-shadow: 0 0 0 2px rgba(201,168,76,0.2) !important;
}
[data-testid="stFileUploader"] {
    background: #1A1A1A !important;
    border: 1px dashed #3A3A3A !important;
    border-radius: 8px !important;
}
[data-testid="stFileUploader"] * { color: #AAAAAA !important; }
[data-testid="stFileUploader"]:hover { border-color: #C9A84C !important; }
[data-testid="stDateInput"] input {
    background: #1E1E1E !important;
    color: #E8E8E8 !important;
    border: 1px solid #3A3A3A !important;
}
[data-baseweb="select"] > div {
    background: #1E1E1E !important;
    border: 1px solid #3A3A3A !important;
    color: #E8E8E8 !important;
}
/* Texto visible dentro del select cerrado */
[data-baseweb="select"] span,
[data-baseweb="select"] div { color: #E8E8E8 !important; }

/* ── Dropdown / popover del selectbox (portal fuera del árbol) ── */
[data-baseweb="popover"],
[data-baseweb="popover"] > div,
ul[data-baseweb="menu"] {
    background: #1E1E1E !important;
    border: 1px solid #3A3A3A !important;
    border-radius: 6px !important;
}
ul[data-baseweb="menu"] li,
[data-baseweb="option"] {
    background: #1E1E1E !important;
    color: #E8E8E8 !important;
}
ul[data-baseweb="menu"] li:hover,
ul[data-baseweb="menu"] li[aria-selected="true"],
[data-baseweb="option"]:hover {
    background: #2D2D2D !important;
    color: #FFFFFF !important;
}

/* ── Labels de widgets ── */
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] label { color: #AAAAAA !important; font-size:13px !important; }

/* ── Expanders ── */
[data-testid="stExpander"] {
    background: #161616 !important;
    border: 1px solid #2A2A2A !important;
    border-radius: 8px !important;
}
/* Header: fondo oscuro → texto claro */
[data-testid="stExpander"] summary,
[data-testid="stExpander"] [data-testid="stExpanderHeader"],
[data-testid="stExpander"] details > summary {
    background: #1E1E1E !important;
    border-radius: 8px !important;
}
[data-testid="stExpander"] summary p,
[data-testid="stExpander"] summary span,
[data-testid="stExpander"] summary div,
[data-testid="stExpander"] summary label,
[data-testid="stExpander"] summary strong,
[data-testid="stExpander"] summary em,
[data-testid="stExpander"] [data-testid="stExpanderHeader"] p,
[data-testid="stExpander"] [data-testid="stExpanderHeader"] span,
[data-testid="stExpander"] [data-testid="stExpanderHeader"] div,
[data-testid="stExpander"] [data-testid="stExpanderHeader"] label {
    color: #CCCCCC !important;
}
[data-testid="stExpander"] summary svg,
[data-testid="stExpander"] [data-testid="stExpanderHeader"] svg {
    fill: #CCCCCC !important;
    stroke: #CCCCCC !important;
}
/* Cuerpo del expander */
[data-testid="stExpander"] [data-testid="stExpanderDetails"],
[data-testid="stExpander"] details > div {
    background: #161616 !important;
}

/* ── Botón primario ── */
button[kind="primary"] { background: #C9A84C !important; color: #0D0D0D !important; border:none !important; font-weight:700 !important; }
button[kind="primary"]:hover { background: #B8973B !important; }
button[kind="secondary"] { background: transparent !important; border: 1px solid #3A3A3A !important; color:#CCCCCC !important; }
button[kind="secondary"]:hover { border-color: #C9A84C !important; color:#C9A84C !important; }

/* ── Info / warning / error boxes ── */
[data-testid="stAlert"] { border-radius:6px !important; }

/* ── Métricas ── */
[data-testid="stMetricValue"] { color: #C9A84C !important; font-size:22px !important; }
[data-testid="stMetricLabel"] { color: #888 !important; }

/* ── Progress ── */
[data-testid="stProgress"] > div > div { background: #C9A84C !important; }
</style>
"""

LIGHT = """
<style>
/* ── Fondo y estructura ── */
.stApp { background: #F5F2EC !important; }
section[data-testid="stSidebar"] { background: #FFFFFF !important; border-right: 1px solid #E0D9CC; }
section[data-testid="stSidebar"] * { color: #333333 !important; }
section[data-testid="stSidebar"] a:hover { color: #8B6914 !important; }
section[data-testid="stSidebar"] [aria-current="page"] { background:#FFF8E7 !important; color:#8B6914 !important; }

/* ── Texto general ── */
p, li, span, label, div { color: #333333; }
h1, h2, h3, h4 { color: #1A1A2E !important; }

/* ── Inputs / widgets ── */
input, textarea, [data-baseweb="input"] input, [data-baseweb="textarea"] textarea {
    background: #FFFFFF !important;
    color: #1A1A2E !important;
    border: 1px solid #D4C9B0 !important;
    border-radius: 6px !important;
}
[data-baseweb="input"]:focus-within, [data-baseweb="textarea"]:focus-within {
    border-color: #C9A84C !important;
    box-shadow: 0 0 0 2px rgba(201,168,76,0.15) !important;
}
[data-testid="stFileUploader"] {
    background: #FFFFFF !important;
    border: 1px dashed #C9A84C !important;
    border-radius: 8px !important;
}
[data-testid="stFileUploader"] * { color: #666666 !important; }
[data-testid="stDateInput"] input {
    background: #FFFFFF !important;
    color: #1A1A2E !important;
    border: 1px solid #D4C9B0 !important;
}
[data-baseweb="select"] > div {
    background: #FFFFFF !important;
    border: 1px solid #D4C9B0 !important;
    color: #1A1A2E !important;
}
/* Texto visible dentro del select cerrado */
[data-baseweb="select"] span,
[data-baseweb="select"] div { color: #1A1A2E !important; }

/* ── Dropdown / popover del selectbox (portal fuera del árbol) ── */
[data-baseweb="popover"],
[data-baseweb="popover"] > div,
ul[data-baseweb="menu"] {
    background: #FFFFFF !important;
    border: 1px solid #D4C9B0 !important;
    border-radius: 6px !important;
}
ul[data-baseweb="menu"] li,
[data-baseweb="option"] {
    background: #FFFFFF !important;
    color: #1A1A2E !important;
}
ul[data-baseweb="menu"] li:hover,
ul[data-baseweb="menu"] li[aria-selected="true"],
[data-baseweb="option"]:hover {
    background: #F0EDE6 !important;
    color: #1A1A2E !important;
}

/* ── Labels de widgets ── */
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] label { color: #555555 !important; font-size:13px !important; }

/* ── Expanders ── */
[data-testid="stExpander"] {
    background: #FFFFFF !important;
    border: 1px solid #E0D9CC !important;
    border-radius: 8px !important;
}
/* Header: fondo claro → texto oscuro */
[data-testid="stExpander"] summary,
[data-testid="stExpander"] [data-testid="stExpanderHeader"],
[data-testid="stExpander"] details > summary {
    background: #F0EDE6 !important;
    border-radius: 8px !important;
}
[data-testid="stExpander"] summary p,
[data-testid="stExpander"] summary span,
[data-testid="stExpander"] summary div,
[data-testid="stExpander"] summary label,
[data-testid="stExpander"] summary strong,
[data-testid="stExpander"] summary em,
[data-testid="stExpander"] [data-testid="stExpanderHeader"] p,
[data-testid="stExpander"] [data-testid="stExpanderHeader"] span,
[data-testid="stExpander"] [data-testid="stExpanderHeader"] div,
[data-testid="stExpander"] [data-testid="stExpanderHeader"] label {
    color: #1A1A2E !important;
}
[data-testid="stExpander"] summary svg,
[data-testid="stExpander"] [data-testid="stExpanderHeader"] svg {
    fill: #1A1A2E !important;
    stroke: #1A1A2E !important;
}
/* Cuerpo del expander */
[data-testid="stExpander"] [data-testid="stExpanderDetails"],
[data-testid="stExpander"] details > div {
    background: #FFFFFF !important;
}

/* ── Botón primario ── */
button[kind="primary"] { background: #C9A84C !important; color: #FFFFFF !important; border:none !important; font-weight:700 !important; }
button[kind="primary"]:hover { background: #8B6914 !important; }
button[kind="secondary"] { background: transparent !important; border: 1px solid #C9A84C !important; color:#8B6914 !important; }
button[kind="secondary"]:hover { background: #FFF8E7 !important; }

/* ── Métricas ── */
[data-testid="stMetricValue"] { color: #8B6914 !important; font-size:22px !important; }
[data-testid="stMetricLabel"] { color: #888 !important; }

/* ── Progress ── */
[data-testid="stProgress"] > div > div { background: #C9A84C !important; }
</style>
"""

# CSS de tarjetas para cada tema
CARDS_DARK = """
<style>
    .mod-eyebrow { font-size:10px; font-weight:700; letter-spacing:4px;
                   text-transform:uppercase; color:#C9A84C; margin-bottom:10px; }
    .main-title  { font-size:28px; font-weight:800; color:#F0F0F0; margin-bottom:4px; }
    .sub-title   { font-size:13px; color:#888; margin-bottom:4px; }
    .gold-line   { width:60px; height:2px;
                   background:linear-gradient(90deg,transparent,#C9A84C,transparent);
                   margin:14px 0 28px 0; }
    .card { border-radius:8px; padding:24px 22px; margin-bottom:8px;
            background:#161616; border:1px solid #2A2A2A; position:relative; }
    .card::before { content:''; position:absolute; top:0; left:0;
                    width:4px; height:100%; border-radius:8px 0 0 8px; }
    .card-active::before { background:linear-gradient(180deg,#C9A84C,#8B6914); }
    .card-coming::before { background:#2A2A2A; }
    .card-icon  { font-size:26px; margin-bottom:8px; }
    .card-title { font-size:16px; font-weight:700; color:#F0F0F0; margin-bottom:5px; }
    .card-desc  { font-size:12px; color:#777; line-height:1.6; margin-bottom:12px; }
    .badge-ready { display:inline-block; border:1px solid #C9A84C; color:#C9A84C;
                   font-size:9px; font-weight:700; letter-spacing:2px;
                   text-transform:uppercase; border-radius:2px; padding:2px 8px; }
    .badge-wip   { display:inline-block; border:1px solid #333; color:#555;
                   font-size:9px; font-weight:700; letter-spacing:2px;
                   text-transform:uppercase; border-radius:2px; padding:2px 8px; }
    .hero-eyebrow { font-size:11px; font-weight:700; letter-spacing:5px;
                    text-transform:uppercase; color:#C9A84C; margin-bottom:8px; }
    .hero-name    { font-size:64px; font-weight:900; color:#F5F5F5;
                    letter-spacing:-2px; line-height:1; margin-bottom:4px; }
    .hero-name span { color:#C9A84C; }
    .hero-product { font-size:18px; font-weight:300; color:#999;
                    letter-spacing:4px; text-transform:uppercase; margin-bottom:12px; }
    .hero-tagline { font-size:14px; color:#666; font-style:italic; }
    .gold-line-center { width:80px; height:3px;
                        background:linear-gradient(90deg,transparent,#C9A84C,transparent);
                        margin:20px auto 36px auto; }
    .intro-text { text-align:center; color:#888; font-size:14px;
                  max-width:600px; margin:0 auto 36px auto; line-height:1.7; }
    .section-label { font-size:10px; font-weight:700; letter-spacing:4px;
                     text-transform:uppercase; color:#C9A84C; margin-bottom:16px; }
    .footer { text-align:center; color:#333; font-size:11px; letter-spacing:1px;
              margin-top:48px; padding-top:20px; border-top:1px solid #1E1E1E; }
    .footer span { color:#C9A84C; }
</style>
"""

CARDS_LIGHT = """
<style>
    .mod-eyebrow { font-size:10px; font-weight:700; letter-spacing:4px;
                   text-transform:uppercase; color:#8B6914; margin-bottom:10px; }
    .main-title  { font-size:28px; font-weight:800; color:#1A1A2E; margin-bottom:4px; }
    .sub-title   { font-size:13px; color:#666; margin-bottom:4px; }
    .gold-line   { width:60px; height:2px;
                   background:linear-gradient(90deg,transparent,#C9A84C,transparent);
                   margin:14px 0 28px 0; }
    .card { border-radius:8px; padding:24px 22px; margin-bottom:8px;
            background:#FFFFFF; border:1px solid #E0D9CC; position:relative;
            box-shadow:0 2px 8px rgba(0,0,0,0.06); }
    .card::before { content:''; position:absolute; top:0; left:0;
                    width:4px; height:100%; border-radius:8px 0 0 8px; }
    .card-active::before { background:linear-gradient(180deg,#C9A84C,#8B6914); }
    .card-coming::before { background:#E0D9CC; }
    .card-icon  { font-size:26px; margin-bottom:8px; }
    .card-title { font-size:16px; font-weight:700; color:#1A1A2E; margin-bottom:5px; }
    .card-desc  { font-size:12px; color:#666; line-height:1.6; margin-bottom:12px; }
    .badge-ready { display:inline-block; border:1px solid #C9A84C; color:#8B6914;
                   font-size:9px; font-weight:700; letter-spacing:2px;
                   text-transform:uppercase; border-radius:2px; padding:2px 8px; }
    .badge-wip   { display:inline-block; border:1px solid #D4C9B0; color:#AAA;
                   font-size:9px; font-weight:700; letter-spacing:2px;
                   text-transform:uppercase; border-radius:2px; padding:2px 8px; }
    .hero-eyebrow { font-size:11px; font-weight:700; letter-spacing:5px;
                    text-transform:uppercase; color:#8B6914; margin-bottom:8px; }
    .hero-name    { font-size:64px; font-weight:900; color:#1A1A2E;
                    letter-spacing:-2px; line-height:1; margin-bottom:4px; }
    .hero-name span { color:#C9A84C; }
    .hero-product { font-size:18px; font-weight:300; color:#888;
                    letter-spacing:4px; text-transform:uppercase; margin-bottom:12px; }
    .hero-tagline { font-size:14px; color:#999; font-style:italic; }
    .gold-line-center { width:80px; height:3px;
                        background:linear-gradient(90deg,transparent,#C9A84C,transparent);
                        margin:20px auto 36px auto; }
    .intro-text { text-align:center; color:#666; font-size:14px;
                  max-width:600px; margin:0 auto 36px auto; line-height:1.7; }
    .section-label { font-size:10px; font-weight:700; letter-spacing:4px;
                     text-transform:uppercase; color:#8B6914; margin-bottom:16px; }
    .footer { text-align:center; color:#AAA; font-size:11px; letter-spacing:1px;
              margin-top:48px; padding-top:20px; border-top:1px solid #E0D9CC; }
    .footer span { color:#8B6914; }
</style>
"""


def render_toggle():
    """Muestra el toggle de tema en el sidebar y aplica el CSS correspondiente."""
    if "tema" not in st.session_state:
        st.session_state.tema = "claro"

    with st.sidebar:
        st.markdown("---")
        col_l, col_r = st.columns([1, 2])
        with col_l:
            st.markdown(
                "🌙" if st.session_state.tema == "oscuro" else "☀️",
                unsafe_allow_html=True,
            )
        with col_r:
            etiqueta = "Claro" if st.session_state.tema == "oscuro" else "Oscuro"
            if st.button(etiqueta, key="toggle_tema", use_container_width=True):
                st.session_state.tema = "claro" if st.session_state.tema == "oscuro" else "oscuro"
                st.rerun()


def apply_theme():
    """Inyecta el CSS del tema activo. Llamar después de render_toggle()."""
    if st.session_state.get("tema", "oscuro") == "oscuro":
        st.markdown(DARK, unsafe_allow_html=True)
        st.markdown(CARDS_DARK, unsafe_allow_html=True)
    else:
        st.markdown(LIGHT, unsafe_allow_html=True)
        st.markdown(CARDS_LIGHT, unsafe_allow_html=True)
