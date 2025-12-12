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
    headers = [h.replace('\xa0', '').strip() for h in data.pop(0)]
    return pd.DataFrame(data, columns=headers)

# --- HELPERS ---
def sanitize_key(text):
    return re.sub(r'[^A-Za-z0-9_]', '_', str(text))

# --- CSS E TEMA (NOVO) ---
def render_page_config_and_css(page_title="Monitoramento"):
    """
    Injeta o CSS global e controla o tema dos Cards (KPIs).
    """
    
    # Controle de Tema na Sidebar
    st.sidebar.markdown("### 🎨 Configuração Visual")
    tema_cards = st.sidebar.radio(
        "Estilo dos Cartões (KPIs):",
        options=["Automático (Adaptável)", "Sempre Escuro"],
        index=0,
        help="Automático: Fundo claro no modo claro, escuro no modo escuro.\nSempre Escuro: Mantém o visual 'Dark' original."
    )

    # Definição das variáveis CSS baseadas na escolha
    if tema_cards == "Sempre Escuro":
        # Força as cores escuras originais
        kpi_bg = "#333333"
        kpi_text = "#FFFFFF"
        kpi_border = "1px solid #444"
    else:
        # Usa variáveis nativas do Streamlit para adaptar ao navegador
        kpi_bg = "var(--secondary-background-color)"
        kpi_text = "var(--text-color)"
        kpi_border = "1px solid rgba(128, 128, 128, 0.2)"

    # CSS Global
    st.markdown(f"""
    <style>
        /* Ajuste global de largura e padding */
        .main .block-container{{
            max-width: 100% !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
        }}

        /* Estilo dos KPI Cards (Os retângulos superiores) */
        .kpi-card {{
            background-color: {kpi_bg};
            color: {kpi_text};
            padding: clamp(12px, 3vw, 20px);
            border-radius: 10px;
            text-align: center;
            border: {kpi_border};
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            transition: all 0.3s ease;
        }}
        
        /* Rótulo do KPI */
        .kpi-label {{
            font-size: clamp(.85rem, 2.5vw, 1rem);
            color: {kpi_text};
            opacity: 0.9;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}

        /* Valores dos KPIs - A cor do número é definida inline no HTML (Red/Blue) 
           mas definimos o tamanho aqui */
        .kpi-value {{
            font-size: clamp(1.6rem, 6vw, 3rem);
            font-weight: 700;
            margin-top: 5px;
        }}

        /* Melhoria nos botões */
        .stButton button {{
            border-radius: 8px;
            font-weight: 600;
            transition: transform 0.1s;
        }}
        .stButton button:active {{
            transform: scale(0.98);
        }}
    </style>
    """, unsafe_allow_html=True)

# --- LOADING OVERLAY ---
def render_loading_overlay(ui_phase: str | None = None):
    phase = ui_phase or st.session_state.get('ui_phase', 'ready')
    display = 'flex' if phase == 'loading' else 'none'
    st.markdown(f"""
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
    """, unsafe_allow_html=True)

def overlay_on():
    st.session_state.ui_phase = 'loading'
    st.session_state.loading_ts = pytime.time()
    render_loading_overlay('loading')

def overlay_off():
    st.session_state.ui_phase = 'ready'
    st.session_state.loading_ts = 0
    render_loading_overlay('ready')