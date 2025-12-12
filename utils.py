import os
import re
import time as pytime

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials


# =========================================================
# CONSTANTES (com aliases p/ compatibilidade com suas páginas)
# =========================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

CREDS_FILE = "google_credentials.json"

SPREADSHEET_URL = (
    "https://docs.google.com/spreadsheets/d/1KeJjbsLVP9DkxPCmNSN4VzbSBeG3SFSCAdPhir39iqg/edit?usp=sharing"
)
SPREADSHEETURL = SPREADSHEET_URL  # compat

SHEET_DESLIGAMENTOS = "DESLIGAMENTOS"
SHEET_EQUIPAMENTOS = "EQUIPAMENTOS"
SHEET_DADOS = "DADOS"
SHEET_DETALHADA = "Usinas_Detalhado"

SHEETDESLIGAMENTOS = SHEET_DESLIGAMENTOS  # compat
SHEETEQUIPAMENTOS = SHEET_EQUIPAMENTOS  # compat
SHEETDADOS = SHEET_DADOS  # compat
SHEETDETALHADA = SHEET_DETALHADA  # compat


# =========================================================
# GOOGLE SHEETS
# =========================================================
@st.cache_resource(ttl=600)
def connect_to_google_sheets():
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    elif os.path.exists(CREDS_FILE):
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    else:
        st.error("Credenciais não encontradas.")
        return None

    return gspread.authorize(creds)


def fetch_sheet_as_df(worksheet):
    data = worksheet.get_all_values()
    if not data:
        return pd.DataFrame()
    headers = [h.replace("\xa0", "").strip() for h in data.pop(0)]
    return pd.DataFrame(data, columns=headers)


# Aliases compat (nomes usados nas páginas)
def connecttogooglesheets():
    return connect_to_google_sheets()


def fetchsheetasdf(worksheet):
    return fetch_sheet_as_df(worksheet)


# =========================================================
# HELPERS
# =========================================================
def sanitize_key(text):
    return re.sub(r"[^A-Za-z0-9_]", "_", str(text))


def sanitizekey(text):
    return sanitize_key(text)


# =========================================================
# CSS / TEMA
# =========================================================
def _inject_common_css():
    """
    CSS que deve valer no tema claro e escuro (grid dias + toggle sidebar).
    Observação: o toggle usa variáveis CSS:
      --sb_toggle_bg
      --sb_toggle_border
      --sb_toggle_fg
    """
    st.markdown(
        """
<style>
/* =========================================================
   0) NÚMERO DO DIA (quando checkbox está sem label)
   ========================================================= */
.day-num{
  width: 100%;
  text-align: center !important;
  font-size: 12px !important;
  line-height: 12px !important;
  margin-top: -6px !important;
  white-space: nowrap !important;
  font-variant-numeric: tabular-nums !important;
  user-select: none !important;
}

/* =========================================================
   1) GRID DE DIAS (checkboxes) - (não atrapalha com checkbox sem label)
   ========================================================= */
div[data-testid="stExpander"] label[data-baseweb="checkbox"]{
  display: flex !important;
  align-items: center !important;
  justify-content: flex-start !important;
  gap: 6px !important;
}
div[data-testid="stExpander"] label[data-baseweb="checkbox"] > div:last-child,
div[data-testid="stExpander"] label[data-baseweb="checkbox"] > span:last-child{
  min-width: 2.6ch !important;
  text-align: center !important;
  white-space: nowrap !important;
  font-variant-numeric: tabular-nums !important;
}
div[data-testid="stExpander"] div[data-testid="stCheckbox"]{
  margin: 0 !important;
  padding: 0 !important;
}

/* =========================================================
   2) SIDEBAR TOGGLE - SEMPRE VISÍVEL (sem depender de hover)
   ========================================================= */

/* Wrapper do botão recolhido (algumas versões) */
div[data-testid="stSidebarCollapsedControl"],
div[data-testid="stSidebarCollapsedControl"] *{
  opacity: 1 !important;
  visibility: visible !important;
  pointer-events: auto !important;
}

/* Lista de seletores do botão (varia por versão/ambiente) */
button[data-testid="collapsedControl"],
button[data-testid="stSidebarCollapseButton"],
section[data-testid="stSidebar"] button[kind="header"],
header button[kind="header"],
/* fallback por aria-label (bem mais estável) */
button[aria-label="Collapse sidebar"],
button[aria-label="Expand sidebar"],
button[aria-label="Open sidebar"],
button[aria-label="Close sidebar"],
button[aria-label="Toggle sidebar"]{
  /* Força aparecer mesmo fora do hover */
  display: inline-flex !important;
  opacity: 1 !important;
  visibility: visible !important;
  pointer-events: auto !important;

  /* Aparência */
  background: var(--sb_toggle_bg, rgba(255,255,255,.95)) !important;
  border: 1px solid var(--sb_toggle_border, rgba(0,0,0,.25)) !important;
  border-radius: 10px !important;
  box-shadow: 0 2px 10px rgba(0,0,0,.25) !important;

  /* Cor do ícone via currentColor */
  color: var(--sb_toggle_fg, #0E1117) !important;
}

/* SVG do ícone: força fill/stroke em qualquer implementação */
button[data-testid="collapsedControl"] svg,
button[data-testid="stSidebarCollapseButton"] svg,
section[data-testid="stSidebar"] button[kind="header"] svg,
header button[kind="header"] svg,
button[aria-label="Collapse sidebar"] svg,
button[aria-label="Expand sidebar"] svg,
button[aria-label="Open sidebar"] svg,
button[aria-label="Close sidebar"] svg,
button[aria-label="Toggle sidebar"] svg{
  fill: var(--sb_toggle_fg, #0E1117) !important;
  stroke: var(--sb_toggle_fg, #0E1117) !important;
  color: var(--sb_toggle_fg, #0E1117) !important;
}

/* Alguns temas setam fill/stroke nos paths */
button[data-testid="collapsedControl"] svg *,
button[data-testid="stSidebarCollapseButton"] svg *,
section[data-testid="stSidebar"] button[kind="header"] svg *,
header button[kind="header"] svg *,
button[aria-label="Collapse sidebar"] svg *,
button[aria-label="Expand sidebar"] svg *,
button[aria-label="Open sidebar"] svg *,
button[aria-label="Close sidebar"] svg *,
button[aria-label="Toggle sidebar"] svg *{
  fill: var(--sb_toggle_fg, #0E1117) !important;
  stroke: var(--sb_toggle_fg, #0E1117) !important;
}
</style>
""",
        unsafe_allow_html=True,
    )


def render_page_config_and_css():
    """
    Injeta CSS e controla tema (Automático / Sempre Escuro) via sidebar.
    """
    st.sidebar.markdown("### Visual")
    opcoes = ["Automático (Claro/Escuro)", "Sempre Escuro"]

    # Compat de estado (algumas páginas usam 'temaescolhido')
    if "tema_escolhido" not in st.session_state:
        st.session_state.tema_escolhido = st.session_state.get("temaescolhido", opcoes[0])
    st.session_state.temaescolhido = st.session_state.tema_escolhido

    try:
        index_atual = opcoes.index(st.session_state.tema_escolhido)
    except ValueError:
        index_atual = 0

    def atualizar_tema():
        st.session_state.tema_escolhido = st.session_state.key_radio_tema
        st.session_state.temaescolhido = st.session_state.tema_escolhido

    tema_cards = st.sidebar.radio(
        "Fundo dos Indicadores:",
        options=opcoes,
        index=index_atual,
        key="key_radio_tema",
        on_change=atualizar_tema,
    )

    # ---------------------------------------------------------
    # Variáveis do botão de sidebar:
    # - Em "Sempre Escuro": branco no ícone + fundo escuro
    # - Em "Automático": usa media query do sistema (prefers-color-scheme)
    # ---------------------------------------------------------
    if tema_cards == "Sempre Escuro":
        toggle_vars_css = """
        <style>
          :root{
            --sb_toggle_bg: rgba(38,39,48,.96);
            --sb_toggle_border: rgba(255,255,255,.55);
            --sb_toggle_fg: #FFFFFF;
          }
        </style>
        """
    else:
        # Base: claro
        toggle_vars_css = """
        <style>
          :root{
            --sb_toggle_bg: rgba(255,255,255,.96);
            --sb_toggle_border: rgba(0,0,0,.25);
            --sb_toggle_fg: #0E1117;
          }
          @media (prefers-color-scheme: dark){
            :root{
              --sb_toggle_bg: rgba(38,39,48,.96);
              --sb_toggle_border: rgba(255,255,255,.55);
              --sb_toggle_fg: #FFFFFF;
            }
          }
        </style>
        """

    st.markdown(toggle_vars_css, unsafe_allow_html=True)

    # CSS comum (claro + escuro) - usa as vars acima
    _inject_common_css()

    global_dark_override = ""

    if tema_cards == "Sempre Escuro":
        kpi_bg = "#333333"
        kpi_text = "#FFFFFF"
        kpi_border = "none"

        global_dark_override = """
        <style>
        /* =========================================================
           Base (fundo/texto)
           ========================================================= */
        .stApp {
          background-color: #0E1117 !important;
          color: #FAFAFA !important;
        }
        section[data-testid="stSidebar"], header[data-testid="stHeader"] {
          background-color: #262730 !important;
        }
        h1, h2, h3, h4, h5, h6, p, li, label,
        .stMarkdown, .stRadio label, .stCheckbox label,
        section[data-testid="stSidebar"] *,
        [data-testid="stSidebarNav"] a, [data-testid="stSidebarNav"] span {
          color: #FAFAFA !important;
        }

        /* =========================================================
           Inputs
           ========================================================= */
        .stTextInput input,
        .stNumberInput input,
        .stDateInput input,
        .stTimeInput input,
        .stTextArea textarea {
          background-color: #262730 !important;
          color: #FAFAFA !important;
          border: 1px solid #4A4A4A !important;
        }
        input::placeholder,
        textarea::placeholder {
          color: #A0A0A0 !important;
          opacity: 1 !important;
        }

        /* =========================================================
           DROPDOWNS (Selectbox / Multiselect) - BaseWeb
           ========================================================= */

        /* Caixa do select (fechado) */
        div[data-baseweb="select"] > div{
          background-color: #262730 !important;
          border-color: #4A4A4A !important;
        }
        div[data-baseweb="select"] *{
          color: #FAFAFA !important;
        }

        /* Popover/lista (aberto) */
        div[data-baseweb="popover"],
        div[data-baseweb="popover"] *,
        ul[data-baseweb="menu"],
        div[data-baseweb="menu"],
        div[role="listbox"],
        div[role="listbox"] * {
          background-color: #262730 !important;
          color: #FAFAFA !important;
          border-color: #4A4A4A !important;
        }

        /* Opções */
        li[data-baseweb="option"],
        div[data-baseweb="option"],
        div[role="option"]{
          background-color: #262730 !important;
          color: #FAFAFA !important;
        }
        li[data-baseweb="option"] *,
        div[data-baseweb="option"] *,
        div[role="option"] *{
          color: #FAFAFA !important;
        }

        /* Hover/selecionado */
        li[data-baseweb="option"]:hover,
        div[data-baseweb="option"]:hover,
        div[role="option"]:hover,
        li[data-baseweb="option"][aria-selected="true"],
        div[data-baseweb="option"][aria-selected="true"],
        div[role="option"][aria-selected="true"]{
          background-color: #FF4B4B !important;
          color: #FFFFFF !important;
        }
        li[data-baseweb="option"]:hover *,
        div[data-baseweb="option"]:hover *,
        div[role="option"]:hover *,
        li[data-baseweb="option"][aria-selected="true"] *,
        div[data-baseweb="option"][aria-selected="true"] *,
        div[role="option"][aria-selected="true"] *{
          color: #FFFFFF !important;
        }

        /* Ícone/seta do select */
        div[data-baseweb="select"] svg{
          fill: #FAFAFA !important;
        }

        /* Tags do Multiselect */
        .stMultiSelect [data-baseweb="tag"]{
          background-color: #FF4B4B !important;
          color: #FFFFFF !important;
        }
        .stMultiSelect [data-baseweb="tag"] *{
          color: #FFFFFF !important;
        }

        /* =========================================================
           EXPANDERS
           ========================================================= */
        div[data-testid="stExpander"] details > summary {
          background-color: #262730 !important;
          border: 1px solid #4A4A4A !important;
          border-radius: 8px !important;
        }
        div[data-testid="stExpander"] details > summary,
        div[data-testid="stExpander"] details > summary * {
          color: #FAFAFA !important;
        }
        div[data-testid="stExpander"] div[role="region"] {
          background-color: #0E1117 !important;
          border: 1px solid #4A4A4A !important;
          border-radius: 8px !important;
          padding: 10px 12px !important;
        }
        </style>
        """
    else:
        # Modo Automático (claro/escuro do Streamlit)
        kpi_bg = "var(--secondary-background-color)"
        kpi_text = "var(--text-color)"
        kpi_border = "1px solid rgba(128, 128, 128, 0.2)"

    if global_dark_override:
        st.markdown(global_dark_override, unsafe_allow_html=True)

    # CSS dos Cards KPI
    st.markdown(
        f"""
    <style>
        .kpi-card {{
            background-color: {kpi_bg};
            color: {kpi_text};
            padding: clamp(12px, 3vw, 20px);
            border-radius: 10px;
            text-align: center;
            border: {kpi_border};
            transition: all 0.3s ease;
        }}
        .kpi-label {{
            font-size: clamp(.85rem, 2.5vw, 1rem);
            color: {kpi_text};
            opacity: 0.9;
        }}
        .kpi-value {{
            font-size: clamp(1.6rem, 6vw, 3rem);
            font-weight: 700;
            margin-top: 5px;
        }}
    </style>
    """,
        unsafe_allow_html=True,
    )


# Alias compat
def renderpageconfigandcss():
    return render_page_config_and_css()


# =========================================================
# LOADING OVERLAY
# =========================================================
def render_loading_overlay(ui_phase: str | None = None):
    # compat: aceita uiphase/ui_phase
    phase = (
        ui_phase
        or st.session_state.get("uiphase")
        or st.session_state.get("ui_phase")
        or "ready"
    )
    display = "flex" if phase == "loading" else "none"

    st.markdown(
        f"""
    <style>
      #__overlay__ {{
        position: fixed; inset: 0;
        display: {display};
        align-items: center; justify-content: center;
        z-index: 10000;
        background: rgba(0,0,0,0);
        animation: bgIn 0s linear 0.25s forwards;
      }}
      .loader {{
        width: 64px; height: 64px; border-radius: 50%;
        border: 6px solid rgba(255,255,255,.25);
        border-top-color: #FF4B4B;
        animation: spin 1s linear infinite, appear 0s linear 0.25s forwards;
        opacity: 0;
      }}
      @keyframes appear {{ to {{ opacity: 1; }} }}
      @keyframes bgIn {{ to {{ background: rgba(0,0,0,.55); }} }}
      @keyframes spin {{ to {{ transform: rotate(360deg); }} }}
    </style>
    <div id="__overlay__"><div class="loader"></div></div>
    """,
        unsafe_allow_html=True,
    )


def overlay_on():
    # compat: escreve nos dois formatos usados no projeto
    st.session_state.uiphase = "loading"
    st.session_state.loadingts = pytime.time()
    st.session_state.ui_phase = "loading"
    st.session_state.loading_ts = st.session_state.loadingts
    render_loading_overlay("loading")


def overlay_off():
    st.session_state.uiphase = "ready"
    st.session_state.loadingts = 0
    st.session_state.ui_phase = "ready"
    st.session_state.loading_ts = 0
    render_loading_overlay("ready")


# Aliases compat
def renderloadingoverlay(ui_phase: str | None = None):
    return render_loading_overlay(ui_phase)


def overlayon():
    return overlay_on()


def overlayoff():
    return overlay_off()