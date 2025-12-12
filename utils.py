import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import time as pytime
import os
import re

# --- CONSTANTES GLOBAIS ---
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
CREDS_FILE = "google_credentials.json"
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1KeJjbsLVP9DkxPCmNSN4VzbSBeG3SFSCAdPhir39iqg/edit?usp=sharing"

SHEET_DESLIGAMENTOS = "DESLIGAMENTOS"
SHEET_EQUIPAMENTOS = "EQUIPAMENTOS"
SHEET_DADOS = "DADOS"
SHEET_DETALHADA = "Usinas_Detalhado"


# --- CONEXÃO GOOGLE SHEETS ---
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
    client = gspread.authorize(creds)
    return client


def fetch_sheet_as_df(worksheet):
    data = worksheet.get_all_values()
    if not data:
        return pd.DataFrame()
    headers = [h.replace("\xa0", "").strip() for h in data.pop(0)]
    return pd.DataFrame(data, columns=headers)


# --- HELPERS ---
def sanitize_key(text):
    return re.sub(r"[^A-Za-z0-9_]", "_", str(text))


# --- CSS E TEMA (CORRIGIDO: Contraste Dropdown) ---
def render_page_config_and_css():
    """
    Injeta o CSS dinâmico e controla o tema visual com persistência manual.
    """
    st.sidebar.markdown("### 🎨 Visual")

    opcoes = ["Automático (Claro/Escuro)", "Sempre Escuro"]

    # 1. Inicializa a memória do tema se ela não existir
    if "tema_escolhido" not in st.session_state:
        st.session_state.tema_escolhido = opcoes[0]

    # 2. Descobre qual o índice da opção salva na memória
    try:
        index_atual = opcoes.index(st.session_state.tema_escolhido)
    except ValueError:
        index_atual = 0

    # 3. Função para atualizar a memória quando o usuário clicar
    def atualizar_tema():
        st.session_state.tema_escolhido = st.session_state.key_radio_tema

    # 4. Renderiza o botão usando o índice da memória
    tema_cards = st.sidebar.radio(
        "Fundo dos Indicadores:",
        options=opcoes,
        index=index_atual,
        key="key_radio_tema",
        on_change=atualizar_tema,
    )

    global_dark_override = ""

    # Lógica de cores baseada na escolha
    if tema_cards == "Sempre Escuro":
        kpi_bg = "#333333"
        kpi_text = "#FFFFFF"
        kpi_border = "none"

        # CSS EXTENDIDO: Força modo escuro em inputs, dropdowns e menus
        # (cobre variações do BaseWeb: popover/listbox/option via data-baseweb e via role)
        global_dark_override = """
        <style>
        /* 1. Fundo Global e Texto Base */
        .stApp {
          background-color: #0E1117 !important;
          color: #FAFAFA !important;
        }

        /* 2. Barra Lateral e Header */
        section[data-testid="stSidebar"], header[data-testid="stHeader"] {
          background-color: #262730 !important;
        }

        /* 3. Textos Gerais e Links */
        h1, h2, h3, h4, h5, h6, p, li, label,
        .stMarkdown, .stRadio label, .stCheckbox label,
        section[data-testid="stSidebar"] *,
        [data-testid="stSidebarNav"] a, [data-testid="stSidebarNav"] span {
          color: #FAFAFA !important;
        }

        /* 4. Inputs básicos */
        .stTextInput input,
        .stNumberInput input,
        .stDateInput input,
        .stTimeInput input,
        .stTextArea textarea {
          background-color: #262730 !important;
          color: #FAFAFA !important;
          border: 1px solid #4A4A4A !important;
        }

        /* Placeholder (inputs e selects) */
        input::placeholder,
        textarea::placeholder,
        div[data-baseweb="select"] [data-testid="stMarkdownContainer"] p {
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

        /* Popover/lista (aberto): cobre variações do BaseWeb */
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

        /* Opções: cobre li e div + role option */
        li[data-baseweb="option"],
        div[data-baseweb="option"],
        div[role="option"]{
          background-color: #262730 !important;
          color: #FAFAFA !important;
        }

        /* Texto interno das opções */
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

        /* Ícone/seta */
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
        </style>
        """
    else:
        # Modo Automático
        kpi_bg = "var(--secondary-background-color)"
        kpi_text = "var(--text-color)"
        kpi_border = "1px solid rgba(128, 128, 128, 0.2)"

    # Injeta o CSS Global
    if global_dark_override:
        st.markdown(global_dark_override, unsafe_allow_html=True)

    # Injeta o CSS dos Cards KPI
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


# --- LOADING OVERLAY ---
def render_loading_overlay(ui_phase: str | None = None):
    # compat: aceita "uiphase/ui_phase"
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
    # compat: escreve nos dois formatos
    st.session_state.uiphase = "loading"
    st.session_state.loadingts = pytime.time()
    st.session_state.ui_phase = "loading"
    st.session_state.loading_ts = st.session_state.loadingts
    render_loading_overlay("loading")


def overlay_off():
    # compat: escreve nos dois formatos
    st.session_state.uiphase = "ready"
    st.session_state.loadingts = 0
    st.session_state.ui_phase = "ready"
    st.session_state.loading_ts = 0
    render_loading_overlay("ready")


# =========================================================
# ALIASES PARA COMPATIBILIDADE COM O CÓDIGO DAS PÁGINAS
# (mantém funcionando quem chama os nomes antigos)
# =========================================================
def renderpageconfigandcss():
    return render_page_config_and_css()


def renderloadingoverlay(ui_phase: str | None = None):
    return render_loading_overlay(ui_phase)


def overlayon():
    return overlay_on()


def overlayoff():
    return overlay_off()
