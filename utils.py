import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import time as pytime
import re
from collections import Counter, defaultdict

# --- CONFIGURAÇÕES GERAIS ---
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ID da Planilha Principal
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1KeJjbsLVP9DkxPCmNSN4VzbSBeG3SFSCAdPhir39iqg/edit?usp=sharing"

# Nomes das Abas
SHEET_DESLIGAMENTOS = "DESLIGAMENTOS"
SHEET_EQUIPAMENTOS = "EQUIPAMENTOS"
SHEET_DADOS = "DADOS"
SHEET_DETALHADA = "Usinas_Detalhado"

# Mapeamento de Meses
MESES_TRADUCAO = {
    'January': 'Janeiro', 'February': 'Fevereiro', 'March': 'Março',
    'April': 'Abril', 'May': 'Maio', 'June': 'Junho',
    'July': 'Julho', 'August': 'Agosto', 'September': 'Setembro',
    'October': 'Outubro', 'November': 'Novembro', 'December': 'Dezembro'
}
MESES_CRONOLOGICOS = list(MESES_TRADUCAO.values())

# Mapeamento de Colunas
MAPA_RENOMEAR = {
    'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência',
    'ATIVO': 'Ativo', 'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência',
    'QUANTIDADE': 'Quantidade', 'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização',
    'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 'DESCRIÇÃO': 'Descrição',
    'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop',
    'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo', 'CLIENTE AVISADO': 'Cliente Avisado'
}

# --- FUNÇÕES DE CONEXÃO ---

@st.cache_resource(ttl=600)
def connect_to_google_sheets():
    """Conecta ao Google Sheets usando secrets.toml ou fallback local."""
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        else:
            # Fallback para desenvolvimento local se o secrets.toml não existir
            creds = Credentials.from_service_account_file("google_credentials.json", scopes=SCOPES)
        
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Erro na conexão com Google Sheets: {e}")
        return None

def fetch_sheet_as_df(worksheet):
    """Lê uma aba e retorna como DataFrame Pandas limpo."""
    data = worksheet.get_all_values()
    if not data:
        return pd.DataFrame()
    headers = [h.replace('\xa0', '').strip() for h in data.pop(0)]
    return pd.DataFrame(data, columns=headers)

@st.cache_data(ttl=600)
def carregar_dados_completos(cache_buster: int = 0):
    """Carrega dados das duas abas principais e processa colunas."""
    try:
        client = connect_to_google_sheets()
        if not client: return pd.DataFrame()
        
        workbook = client.open_by_url(SPREADSHEET_URL)
        
        df_desligamentos = fetch_sheet_as_df(workbook.worksheet(SHEET_DESLIGAMENTOS))
        df_equipamentos = fetch_sheet_as_df(workbook.worksheet(SHEET_EQUIPAMENTOS))

        df_desligamentos['Categoria'] = SHEET_DESLIGAMENTOS
        df_equipamentos['Categoria']  = SHEET_EQUIPAMENTOS
        
        df_final = pd.concat([df_desligamentos, df_equipamentos], ignore_index=True)

        # Renomear colunas
        cols_atuais = df_final.columns
        renomear = {}
        for col in cols_atuais:
            c_upper = col.strip().upper()
            if c_upper in MAPA_RENOMEAR:
                renomear[col] = MAPA_RENOMEAR[c_upper]
        df_final.rename(columns=renomear, inplace=True)
        df_final.fillna('', inplace=True)

        # Tratar Datas
        cols_data = ['Normalização', 'Desligamento', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']
        for c in cols_data:
            if c in df_final.columns:
                df_final[c] = pd.to_datetime(df_final[c], errors='coerce', dayfirst=True)

        # Enriquecimento de Dados
        if 'Desligamento' in df_final.columns:
            df_final['Data'] = df_final['Desligamento'].dt.strftime('%Y-%m-%d')
            df_final['Hora'] = df_final['Desligamento'].dt.strftime('%H:%M:%S')
            df_final['Mês']  = df_final['Desligamento'].dt.strftime('%B').map(MESES_TRADUCAO)
            df_final['Ano']  = df_final['Desligamento'].dt.year.fillna(0).astype(int)
            df_final['Dia']  = df_final['Desligamento'].dt.day.fillna(0).astype(int)
            
            # ID Único Crítico para Edição
            df_final['ID_Unico'] = (
                df_final['UG'].astype(str).str.upper() + "|" +
                df_final['Ativo'].astype(str).str.upper() + "|" +
                df_final['Ocorrência'].astype(str).str.upper() + "|" +
                df_final['Desligamento'].astype(str)
            )

        return df_final
    except Exception as e:
        st.error(f"Erro ao carregar dados completos: {e}")
        return pd.DataFrame()

# --- HELPERS DE TEXTO E FILTROS ---

def collapse_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()

def canon(s) -> str:
    if not s: return ""
    return collapse_spaces(str(s)).casefold()

def options_from(series: pd.Series) -> list:
    ser = series.astype(str).map(collapse_spaces)
    ser = ser[(ser != "") & (ser != "-") & (ser != "0")]
    return sorted(ser.unique().tolist())

def matches_any_canon(series: pd.Series, selected: list) -> pd.Series:
    if not selected:
        return pd.Series([True] * len(series), index=series.index)
    selected_canon = {canon(s) for s in selected}
    return series.astype(str).map(canon).isin(selected_canon)

def sanitize_key(text):
    return re.sub(r'[^A-Za-z0-9_]', '_', str(text))

# --- UI / OVERLAY ---

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

def init_overlay():
    if 'ui_phase' not in st.session_state:
        st.session_state.ui_phase = 'ready'
    if 'loading_ts' not in st.session_state:
        st.session_state.loading_ts = 0
    render_loading_overlay()