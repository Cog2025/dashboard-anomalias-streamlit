import os
import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import html
import re
from collections import Counter, defaultdict
import utils

# --- Configuração Inicial ---
st.set_page_config(layout="wide", page_title="Histórico Resolvidas")

# Renderiza Overlay
utils.render_loading_overlay(st.session_state.get('ui_phase', 'ready'))

if 'cache_buster' not in st.session_state:
    st.session_state.cache_buster = int(pytime.time())

# --- Helpers de Texto (Lógica idêntica à Pg 1 para garantir o filtro correto) ---
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

# Helper de Loading
def start_loading():
    st.session_state.ui_phase = 'loading'
    st.session_state.loading_ts = pytime.time()

if st.session_state.ui_phase == 'init': 
    start_loading()
    st.rerun()

# --- CSS Personalizado ---
st.markdown("""
<style>
    .stButton > button { background-color: #28a745; color: white; font-weight: bold; width: 100%; }
    .stButton > button:hover { background-color: #218838; }
    
    .kpi-card { background-color: #333333; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px; }
    .kpi-value { font-size: 2.5em; font-weight: bold; color: #4b4eff; }
    .kpi-label { font-size: 1.1em; color: #FFFFFF; }
    
    .card-container { background-color: #089641; color: white; padding: 15px; border-radius: 8px; margin-bottom: 15px; height: 100%; box-shadow: 0 4px 8px rgba(0,0,0,0.2); }
    .card-title { font-size: 1.4em; font-weight: bold; border-bottom: 1px solid rgba(255,255,255,0.5); padding-bottom: 5px; margin-bottom: 10px; }
    .card-item { margin-bottom: 4px; font-size: 0.95em; }
    .card-label { font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- Carregamento de Dados ---
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
        
        # Limpeza básica de colunas vazias
        if 'Cliente' in df.columns:
            df = df[(df['Cliente'] != '') & (df['UG'] != '')]

        # Tratamento de Datas
        for c in ['Normalização', 'Desligamento', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']:
            if c in df.columns: df[c] = pd.to_datetime(df[c], errors='coerce')
        
        # Colunas Auxiliares
        if 'Desligamento' in df.columns:
            df['Ano'] = df['Desligamento'].dt.year.fillna(0).astype(int)
            meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
            m_map = {i+1: m for i, m in enumerate(meses)}
            df['Mês'] = df['Desligamento'].dt.month.map(m_map)
            
        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

df_todos = carregar_dados(st.session_state.cache_buster)

if st.session_state.ui_phase == 'loading': 
    st.session_state.ui_phase = 'ready'
    st.session_state.loading_ts = 0
    utils.render_loading_overlay('ready')

# Filtra apenas resolvidas (Base para tudo)
df_resolvidas_base = df_todos[df_todos['Normalização'].notna()].copy() if not df_todos.empty else pd.DataFrame()

# ==============================================================================
# --- BARRA LATERAL (FILTROS) ---
# ==============================================================================
with st.sidebar:
    st.title("Filtros Histórico")
    
    if st.button("🔄 Atualizar Dados"):
        st.cache_data.clear()
        start_loading()
        st.rerun()
    
    st.markdown("---")
    
    # Inicializa estado
    if 'hist_anos' not in st.session_state: st.session_state.hist_anos = []
    if 'hist_meses' not in st.session_state: st.session_state.hist_meses = []
    for k in ['hist_cli', 'hist_ug', 'hist_tipo', 'hist_ativo', 'hist_ocr']:
        if k not in st.session_state: st.session_state[k] = []

    # Opções Disponíveis (Baseado nas resolvidas)
    if not df_resolvidas_base.empty:
        opts_ano = sorted([a for a in df_resolvidas_base['Ano'].unique() if a != 0])
        meses_cron = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
        
        # 1. Filtros de Período
        with st.expander("📅 Período", expanded=True):
            st.session_state.hist_anos = st.multiselect("Anos", opts_ano, default=st.session_state.hist_anos)
            st.session_state.hist_meses = st.multiselect("Meses", meses_cron, default=st.session_state.hist_meses)
        
        # 2. Filtros Adicionais
        st.markdown("**Filtros Adicionais**")
        
        def render_sidebar_filter(label, key, col_name):
            with st.expander(label):
                opts = options_from(df_resolvidas_base[col_name])
                opts_clean = [o for o in opts if o != "-"]
                
                # Botões de ação rápida
                b1, b2 = st.columns(2)
                if b1.button('Todos', key=f'all_{key}'):
                    st.session_state[key] = opts_clean
                    st.rerun()
                if b2.button('Nada', key=f'none_{key}'):
                    st.session_state[key] = []
                    st.rerun()
                
                st.session_state[key] = st.multiselect(
                    "Selecione", 
                    opts_clean, 
                    default=[x for x in st.session_state[key] if x in opts_clean],
                    key=f'ms_{key}'
                )

        render_sidebar_filter("Clientes", 'hist_cli', 'Cliente')
        render_sidebar_filter("UGs", 'hist_ug', 'UG')
        render_sidebar_filter("Tipos", 'hist_tipo', 'Tipo de ocorrência')
        render_sidebar_filter("Ativos", 'hist_ativo', 'Ativo')
        render_sidebar_filter("Ocorrências", 'hist_ocr', 'Ocorrência')

# ==============================================================================
# --- ÁREA PRINCIPAL ---
# ==============================================================================
st.title("Histórico de Ocorrências Resolvidas")

# KPIs Topo
c1, c2 = st.columns(2)
v1 = df_resolvidas_base[df_resolvidas_base['Categoria'] == 'DESLIGAMENTOS'].shape[0] if not df_resolvidas_base.empty else 0
v2 = df_resolvidas_base[df_resolvidas_base['Categoria'] == 'EQUIPAMENTOS'].shape[0] if not df_resolvidas_base.empty else 0

c1.markdown(f"<div class='kpi-card'><div class='kpi-label'>DESLIGAMENTOS RESOLVIDOS</div><div class='kpi-value'>{v1}</div></div>", unsafe_allow_html=True)
c2.markdown(f"<div class='kpi-card'><div class='kpi-label'>EQUIPAMENTOS RESOLVIDOS</div><div class='kpi-value'>{v2}</div></div>", unsafe_allow_html=True)

# Aplicação Filtros
df_filt = df_resolvidas_base.copy()
if not df_filt.empty:
    if st.session_state.hist_anos:
        df_filt = df_filt[df_filt['Ano'].isin(st.session_state.hist_anos)]
    if st.session_state.hist_meses:
        df_filt = df_filt[df_filt['Mês'].isin(st.session_state.hist_meses)]
    
    # Aplica filtros de texto usando a lógica robusta "matches_any_canon"
    df_filt = df_filt[matches_any_canon(df_filt['Cliente'], st.session_state.hist_cli)]
    df_filt = df_filt[matches_any_canon(df_filt['UG'], st.session_state.hist_ug)]
    df_filt = df_filt[matches_any_canon(df_filt['Tipo de ocorrência'], st.session_state.hist_tipo)]
    df_filt = df_filt[matches_any_canon(df_filt['Ativo'], st.session_state.hist_ativo)]
    df_filt = df_filt[matches_any_canon(df_filt['Ocorrência'], st.session_state.hist_ocr)]

st.markdown("---")
st.markdown(f"### 🔍 Visualizando: **{len(df_filt)}** registros filtrados")

if not df_filt.empty:
    # Ordenação
    col_sort1, col_sort2 = st.columns(2)
    with col_sort1:
        sort_col = st.selectbox("Ordenar por:", ["Desligamento", "Normalização", "UG"])
    with col_sort2:
        sort_ord = st.radio("Ordem:", ["Descendente", "Ascendente"], horizontal=True)
    
    df_filt = df_filt.sort_values(by=sort_col, ascending=(sort_ord == "Ascendente"))

    # Tabela
    st.dataframe(df_filt[['UG', 'Categoria', 'Data', 'Hora', 'Normalização', 'Ocorrência', 'Descrição']], use_container_width=True)

    # Cards Detalhados
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
                    # Helpers de texto seguro
                    cli = html.escape(str(row.get("Cliente", "")))
                    cat = html.escape(str(row.get("Categoria", "")))
                    ug = html.escape(str(row.get("UG", "N/A")))
                    tipo = html.escape(str(row.get("Tipo de ocorrência", "")))
                    ativo = html.escape(str(row.get("Ativo", "")))
                    nome = html.escape(str(row.get("Nome Ativo", "")))
                    ocr = html.escape(str(row.get("Ocorrência", "")))
                    oper = html.escape(str(row.get("Operador", "")))
                    desc = html.escape(str(row.get("Descrição", ""))).replace('\n', '<br>')
                    prot = html.escape(str(row.get("Protocolo", "")))
                    osv = html.escape(str(row.get("OS", "")))
                    
                    d_des, h_des = fmt_dt(row.get('Desligamento'))
                    d_norm, h_norm = fmt_dt(row.get('Normalização'))
                    d_ca, h_ca = fmt_dt(row.get('Cliente Avisado'))
                    d_loop, h_loop = fmt_dt(row.get('Atendimento Loop'))
                    d_terc, h_terc = fmt_dt(row.get('Atendimento Terceiros'))
                    
                    qtd_html = ''
                    if row.get('Categoria') == 'EQUIPAMENTOS':
                        try:
                            qv = float(row.get('Quantidade', 0))
                            if qv > 0: qtd_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {int(qv)}</div>'
                        except: pass

                    st.markdown(f"""
                    <div class="card-container">
                      <div class="card-title">{ug}</div>
                      <div class="card-item"><span class="card-label">Cliente:</span> {cli}</div>
                      <div class="card-item"><span class="card-label">Categoria:</span> {cat}</div>
                      <div class="card-item"><span class="card-label">Tipo:</span> {tipo}</div>
                      <div class="card-item"><span class="card-label">Ativo:</span> {ativo}</div>
                      <div class="card-item"><span class="card-label">Nome:</span> {nome}</div>
                      <div class="card-item"><span class="card-label">Ocorrência:</span> {ocr}</div>
                      <div class="card-item"><span class="card-label">Operador:</span> {oper}</div>
                      {qtd_html}
                      <br>
                      <div class="card-item"><span class="card-label">Desligamento:</span> {d_des} {h_des}</div>
                      <div class="card-item"><span class="card-label">Normalização:</span> {d_norm} {h_norm}</div>
                      <div class="card-item"><span class="card-label">Aviso:</span> {d_ca} {h_ca}</div>
                      <div class="card-item"><span class="card-label">Loop:</span> {d_loop} {h_loop}</div>
                      <div class="card-item"><span class="card-label">Terc.:</span> {d_terc} {h_terc}</div>
                      <br>
                      <div class="card-item"><span class="card-label">Desc:</span> {desc}</div>
                      <div class="card-item"><span class="card-label">Prot:</span> {prot}</div>
                      <div class="card-item"><span class="card-label">OS:</span> {osv}</div>
                    </div>
                    """, unsafe_allow_html=True)
else:
    st.info("Nenhuma ocorrência resolvida encontrada para os filtros selecionados.")