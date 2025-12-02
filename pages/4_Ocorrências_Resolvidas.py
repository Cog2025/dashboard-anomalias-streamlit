import os
import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import html
import re
import utils

# --- Configuração ---
st.set_page_config(layout="wide", page_title="Histórico Resolvidas")

# Overlay
utils.render_loading_overlay(st.session_state.get('ui_phase', 'ready'))

def start_loading():
    st.session_state.ui_phase = 'loading'
    st.session_state.loading_ts = pytime.time()

if st.session_state.ui_phase == 'init': 
    start_loading(); st.rerun()

# --- Helpers ---
def _collapse_spaces(s: str) -> str:
    return " ".join(str(s).split())

def canon(s) -> str:
    if s is None: return ""
    return _collapse_spaces(str(s)).casefold()

def options_from(series: pd.Series) -> list:
    series = series.astype(str).map(_collapse_spaces)
    unique_vals = sorted([x for x in series.unique() if x and x.lower() != "nan" and x != "" and x != "-"])
    return ["-"] + unique_vals

def matches_any_canon(series: pd.Series, selected: list[str]) -> pd.Series:
    if not selected:
        return pd.Series([True]*len(series), index=series.index)
    sel_c = {canon(s) for s in selected if s and s != "-"}
    return series.astype(str).map(canon).isin(sel_c)

# --- Dados ---
@st.cache_data(ttl=600)
def carregar_dados(cb):
    try:
        client = utils.connect_to_google_sheets()
        if not client: return pd.DataFrame()
        
        wb = client.open_by_url(utils.SPREADSHEET_URL)
        df1 = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_DESLIGAMENTOS))
        df2 = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_EQUIPAMENTOS))
        
        df1['Categoria'] = 'DESLIGAMENTOS'
        df2['Categoria'] = 'EQUIPAMENTOS'
        df = pd.concat([df1, df2], ignore_index=True)
        
        # Mapa de Renomeação
        mapa = {
            'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência',
            'ATIVO': 'Ativo', 'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência',
            'QUANTIDADE': 'Quantidade', 'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização',
            'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 'DESCRIÇÃO': 'Descrição',
            'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop',
            'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo', 'CLIENTE AVISADO': 'Cliente Avisado'
        }
        
        renomear = {}
        for col in df.columns:
            c_upper = col.strip().upper()
            if c_upper in mapa: renomear[col] = mapa[c_upper]
        df.rename(columns=renomear, inplace=True)
        
        # Datas
        cols_dt = ['Normalização', 'Desligamento', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']
        for c in cols_dt:
            if c in df.columns: df[c] = pd.to_datetime(df[c], errors='coerce', dayfirst=True)
            
        if 'Desligamento' in df.columns:
            df['Ano'] = df['Desligamento'].dt.year.fillna(0).astype(int)
            meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
            m_map = {i+1: m for i, m in enumerate(meses)}
            df['Mês'] = df['Desligamento'].dt.month.map(m_map)
            
        return df
    except Exception as e:
        st.error(f"Erro: {e}")
        return pd.DataFrame()

if 'cache_buster' not in st.session_state: st.session_state.cache_buster = int(pytime.time())
df_todos = carregar_dados(st.session_state.cache_buster)

if st.session_state.ui_phase == 'loading': 
    st.session_state.ui_phase = 'ready'
    st.session_state.loading_ts = 0
    utils.render_loading_overlay('ready')

# Filtra resolvidas
if 'Normalização' not in df_todos.columns:
    st.error("Coluna 'Normalização' não encontrada.")
    st.stop()

df_resolvidas_base = df_todos[df_todos['Normalização'].notna()].copy() if not df_todos.empty else pd.DataFrame()

# ==============================================================================
# --- SIDEBAR ---
# ==============================================================================
with st.sidebar:
    st.header("Filtros Histórico")
    if st.button("🔄 Atualizar"):
        st.cache_data.clear()
        start_loading()
        st.rerun()
    
    st.markdown("---")
    
    if 'hist_anos' not in st.session_state: st.session_state.hist_anos = []
    if 'hist_meses' not in st.session_state: st.session_state.hist_meses = []
    for k in ['hist_cli', 'hist_ug', 'hist_tipo', 'hist_ativo', 'hist_ocr']:
        if k not in st.session_state: st.session_state[k] = []

    def set_filtro(key, values):
        st.session_state[key] = values
        start_loading()

    if not df_resolvidas_base.empty:
        opts_ano = sorted([a for a in df_resolvidas_base['Ano'].unique() if a != 0])
        meses_cron = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
        
        with st.expander("📅 Período", expanded=True):
            c1, c2 = st.columns(2)
            c1.button("Todos", key="all_anos_h", on_click=set_filtro, args=('hist_anos', opts_ano))
            c2.button("Nada", key="none_anos_h", on_click=set_filtro, args=('hist_anos', []))
            st.session_state.hist_anos = st.multiselect("Anos", opts_ano, default=st.session_state.hist_anos, label_visibility="collapsed")
            
            c1, c2 = st.columns(2)
            c1.button("Todos", key="all_mes_h", on_click=set_filtro, args=('hist_meses', meses_cron))
            c2.button("Nada", key="none_mes_h", on_click=set_filtro, args=('hist_meses', []))
            st.session_state.hist_meses = st.multiselect("Meses", meses_cron, default=st.session_state.hist_meses, label_visibility="collapsed")

        def render_sidebar_filter(label, key, col_name):
            if col_name not in df_resolvidas_base.columns: return
            with st.expander(label):
                opts = options_from(df_resolvidas_base[col_name])
                opts_clean = [o for o in opts if o != "-"]
                
                c1, c2 = st.columns(2)
                c1.button("Todos", key=f'all_{key}', on_click=set_filtro, args=(key, opts_clean))
                c2.button("Nada", key=f'none_{key}', on_click=set_filtro, args=(key, []))
                
                st.session_state[key] = st.multiselect("Selecione", opts_clean, default=[x for x in st.session_state[key] if x in opts_clean], key=f'ms_{key}', label_visibility="collapsed")

        render_sidebar_filter("Clientes", 'hist_cli', 'Cliente')
        render_sidebar_filter("UGs", 'hist_ug', 'UG')
        render_sidebar_filter("Tipos", 'hist_tipo', 'Tipo de ocorrência')
        render_sidebar_filter("Ativos", 'hist_ativo', 'Ativo')
        render_sidebar_filter("Ocorrências", 'hist_ocr', 'Ocorrência')

# ==============================================================================
# --- MAIN AREA ---
# ==============================================================================
st.title("Histórico de Ocorrências Resolvidas")

# KPIs
c1, c2 = st.columns(2)
# CSS KPI
st.markdown("""
<style>
    .kpi-card { background-color: #333333; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px; }
    .kpi-value { font-size: 2.5em; font-weight: bold; color: #4b4eff; }
    .kpi-label { font-size: 1.1em; color: #FFFFFF; }
    .card-container { background-color: #089641; color: white; padding: 15px; border-radius: 8px; margin-bottom: 15px; }
    .card-title { font-size: 1.4em; font-weight: bold; border-bottom: 1px solid rgba(255,255,255,0.5); padding-bottom: 5px; margin-bottom: 10px; }
    .card-item { margin-bottom: 4px; font-size: 0.95em; }
    .card-label { font-weight: bold; }
    .stButton > button { background-color: #28a745; color: white; font-weight: bold; width: 100%; }
</style>
""", unsafe_allow_html=True)

v1 = df_resolvidas_base[df_resolvidas_base['Categoria'] == 'DESLIGAMENTOS'].shape[0] if not df_resolvidas_base.empty else 0
v2 = df_resolvidas_base[df_resolvidas_base['Categoria'] == 'EQUIPAMENTOS'].shape[0] if not df_resolvidas_base.empty else 0

c1.markdown(f"<div class='kpi-card'><div class='kpi-label'>DESLIGAMENTOS RESOLVIDOS</div><div class='kpi-value'>{v1}</div></div>", unsafe_allow_html=True)
c2.markdown(f"<div class='kpi-card'><div class='kpi-label'>EQUIPAMENTOS RESOLVIDOS</div><div class='kpi-value'>{v2}</div></div>", unsafe_allow_html=True)

# Aplica Filtros
df_filt = df_resolvidas_base.copy()
if not df_filt.empty:
    if st.session_state.hist_anos: df_filt = df_filt[df_filt['Ano'].isin(st.session_state.hist_anos)]
    if st.session_state.hist_meses: df_filt = df_filt[df_filt['Mês'].isin(st.session_state.hist_meses)]
    
    df_filt = df_filt[matches_any_canon(df_filt['Cliente'], st.session_state.hist_cli)]
    df_filt = df_filt[matches_any_canon(df_filt['UG'], st.session_state.hist_ug)]
    df_filt = df_filt[matches_any_canon(df_filt['Tipo de ocorrência'], st.session_state.hist_tipo)]
    df_filt = df_filt[matches_any_canon(df_filt['Ativo'], st.session_state.hist_ativo)]
    df_filt = df_filt[matches_any_canon(df_filt['Ocorrência'], st.session_state.hist_ocr)]

st.markdown("---")
st.markdown(f"### 🔍 Visualizando: **{len(df_filt)}** registros")

if not df_filt.empty:
    # --- CORREÇÃO DO KEYERROR: Verifica colunas disponíveis antes de exibir ---
    colunas_desejadas = ['UG', 'Categoria', 'Data', 'Hora', 'Normalização', 'Ocorrência', 'Descrição']
    colunas_existentes = [c for c in colunas_desejadas if c in df_filt.columns]
    
    st.dataframe(df_filt[colunas_existentes], use_container_width=True)

    # Cards
    st.markdown("### Detalhes (Cards)")
    num_cols = 4
    rows = list(df_filt.iterrows())
    
    def fmt_dt(dt):
        if pd.notna(dt): return dt.strftime('%d/%m/%Y'), dt.strftime('%H:%M')
        return '', ''

    for i in range(0, len(rows), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j < len(rows):
                _, row = rows[i+j]
                with cols[j]:
                    ug = html.escape(str(row.get("UG", "N/A")))
                    cli = html.escape(str(row.get("Cliente", "")))
                    ocr = html.escape(str(row.get("Ocorrência", "")))
                    d_norm, h_norm = fmt_dt(row.get('Normalização'))
                    desc = html.escape(str(row.get("Descrição", "")))[:100] + "..." # Resumo

                    st.markdown(f"""
                    <div class="card-container">
                      <div class="card-title">{ug}</div>
                      <div class="card-item"><span class="card-label">Cliente:</span> {cli}</div>
                      <div class="card-item"><span class="card-label">Ocorrência:</span> {ocr}</div>
                      <div class="card-item"><span class="card-label">Normalização:</span> {d_norm} {h_norm}</div>
                      <br>
                      <div class="card-item"><span class="card-label">Desc:</span> {desc}</div>
                    </div>
                    """, unsafe_allow_html=True)
else:
    st.info("Nenhum histórico.")