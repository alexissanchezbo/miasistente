import streamlit as st

LIGHT = """
<style>
/* ── Fondo y estructura ── */
.stApp { background: #FFFFFF !important; }
section[data-testid="stSidebar"] { background: #F5F5F5 !important; border-right: 1px solid #DDDDDD; }
section[data-testid="stSidebar"] * { color: #333333 !important; }
section[data-testid="stSidebar"] a:hover { color: #111111 !important; }
section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"] { border-radius:6px; }
section[data-testid="stSidebar"] [aria-current="page"] { background:#E2E2E2 !important; color:#111111 !important; font-weight:700; }

/* ── Texto general ── */
p, li, span, label, div { color: #222222; }
h1, h2, h3, h4 { color: #111111 !important; }

/* ── Inputs / widgets ── */
input, textarea, [data-baseweb="input"] input, [data-baseweb="textarea"] textarea {
    background: #FFFFFF !important;
    color: #111111 !important;
    border: 1px solid #CCCCCC !important;
    border-radius: 6px !important;
}
[data-baseweb="input"]:focus-within, [data-baseweb="textarea"]:focus-within {
    border-color: #555555 !important;
    box-shadow: 0 0 0 2px rgba(0,0,0,0.08) !important;
}
[data-testid="stFileUploader"] {
    background: #FAFAFA !important;
    border: 1px dashed #AAAAAA !important;
    border-radius: 8px !important;
}
[data-testid="stFileUploader"] * { color: #555555 !important; }
[data-testid="stFileUploader"]:hover { border-color: #444444 !important; }
[data-testid="stDateInput"] input {
    background: #FFFFFF !important;
    color: #111111 !important;
    border: 1px solid #CCCCCC !important;
}
[data-baseweb="select"] > div {
    background: #FFFFFF !important;
    border: 1px solid #CCCCCC !important;
    color: #111111 !important;
}
[data-baseweb="select"] span,
[data-baseweb="select"] div { color: #111111 !important; }

/* ── Dropdown / popover del selectbox ── */
[data-baseweb="popover"],
[data-baseweb="popover"] > div,
ul[data-baseweb="menu"] {
    background: #FFFFFF !important;
    border: 1px solid #CCCCCC !important;
    border-radius: 6px !important;
}
ul[data-baseweb="menu"] li,
[data-baseweb="option"] {
    background: #FFFFFF !important;
    color: #111111 !important;
}
ul[data-baseweb="menu"] li:hover,
ul[data-baseweb="menu"] li[aria-selected="true"],
[data-baseweb="option"]:hover {
    background: #EEEEEE !important;
    color: #000000 !important;
}

/* ── Labels de widgets ── */
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] label { color: #555555 !important; font-size:13px !important; }

/* ── Expanders ── */
[data-testid="stExpander"] {
    background: #FFFFFF !important;
    border: 1px solid #DDDDDD !important;
    border-radius: 8px !important;
}
[data-testid="stExpander"] summary,
[data-testid="stExpander"] [data-testid="stExpanderHeader"],
[data-testid="stExpander"] details > summary {
    background: #F0F0F0 !important;
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
    color: #222222 !important;
}
[data-testid="stExpander"] summary svg,
[data-testid="stExpander"] [data-testid="stExpanderHeader"] svg {
    fill: #444444 !important;
    stroke: #444444 !important;
}
[data-testid="stExpander"] [data-testid="stExpanderDetails"],
[data-testid="stExpander"] details > div {
    background: #FFFFFF !important;
}

/* ── Botones ── */
button[kind="primary"] { background: #333333 !important; color: #FFFFFF !important; border:none !important; font-weight:700 !important; }
button[kind="primary"]:hover { background: #111111 !important; }
button[kind="secondary"] { background: transparent !important; border: 1px solid #AAAAAA !important; color:#333333 !important; }
button[kind="secondary"]:hover { border-color: #333333 !important; color:#000000 !important; }

/* ── Info / warning / error boxes ── */
[data-testid="stAlert"] { border-radius:6px !important; }

/* ── Métricas ── */
[data-testid="stMetricValue"] { color: #111111 !important; font-size:22px !important; font-weight:700 !important; }
[data-testid="stMetricLabel"] { color: #777777 !important; }

/* ── Progress ── */
[data-testid="stProgress"] > div > div { background: #444444 !important; }
</style>
"""

CARDS = """
<style>
    .mod-eyebrow { font-size:10px; font-weight:700; letter-spacing:4px;
                   text-transform:uppercase; color:#555555; margin-bottom:10px; }
    .main-title  { font-size:28px; font-weight:800; color:#111111; margin-bottom:4px; }
    .sub-title   { font-size:13px; color:#666666; margin-bottom:4px; }
    .gold-line   { width:60px; height:2px;
                   background:linear-gradient(90deg,transparent,#888888,transparent);
                   margin:14px 0 28px 0; }
    .card { border-radius:8px; padding:24px 22px; margin-bottom:8px;
            background:#FFFFFF; border:1px solid #DDDDDD; position:relative;
            box-shadow:0 2px 8px rgba(0,0,0,0.06); }
    .card::before { content:''; position:absolute; top:0; left:0;
                    width:4px; height:100%; border-radius:8px 0 0 8px; }
    .card-active::before { background:linear-gradient(180deg,#444444,#888888); }
    .card-coming::before { background:#DDDDDD; }
    .card-icon  { font-size:26px; margin-bottom:8px; }
    .card-title { font-size:16px; font-weight:700; color:#111111; margin-bottom:5px; }
    .card-desc  { font-size:12px; color:#666666; line-height:1.6; margin-bottom:12px; }
    .badge-ready { display:inline-block; border:1px solid #444444; color:#333333;
                   font-size:9px; font-weight:700; letter-spacing:2px;
                   text-transform:uppercase; border-radius:2px; padding:2px 8px; }
    .badge-wip   { display:inline-block; border:1px solid #CCCCCC; color:#AAAAAA;
                   font-size:9px; font-weight:700; letter-spacing:2px;
                   text-transform:uppercase; border-radius:2px; padding:2px 8px; }
    .hero-eyebrow { font-size:11px; font-weight:700; letter-spacing:5px;
                    text-transform:uppercase; color:#555555; margin-bottom:8px; }
    .hero-name    { font-size:64px; font-weight:900; color:#111111;
                    letter-spacing:-2px; line-height:1; margin-bottom:4px; }
    .hero-name span { color:#444444; }
    .hero-product { font-size:18px; font-weight:300; color:#888888;
                    letter-spacing:4px; text-transform:uppercase; margin-bottom:12px; }
    .hero-tagline { font-size:14px; color:#777777; font-style:italic; }
    .gold-line-center { width:80px; height:3px;
                        background:linear-gradient(90deg,transparent,#888888,transparent);
                        margin:20px auto 36px auto; }
    .intro-text { text-align:center; color:#666666; font-size:14px;
                  max-width:600px; margin:0 auto 36px auto; line-height:1.7; }
    .section-label { font-size:10px; font-weight:700; letter-spacing:4px;
                     text-transform:uppercase; color:#555555; margin-bottom:16px; }
    .footer { text-align:center; color:#AAAAAA; font-size:11px; letter-spacing:1px;
              margin-top:48px; padding-top:20px; border-top:1px solid #EEEEEE; }
    .footer span { color:#444444; }
</style>
"""


def render_toggle():
    """Sin toggle — tema fijo en claro."""
    pass


def apply_theme():
    """Inyecta el CSS del tema claro en escala de grises."""
    st.markdown(LIGHT, unsafe_allow_html=True)
    st.markdown(CARDS, unsafe_allow_html=True)
