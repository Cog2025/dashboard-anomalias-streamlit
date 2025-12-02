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

if st.session_state.get('ui_phase') == 'init': 
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

# --- Dados (Lógica Original para evitar perda de KPIs) ---
@st.cache_data(ttl=600)
def carregar_dados(cb):
    try:
        client = utils.connect_to_google_sheets()
        if not client: return pd.DataFrame()
        
        wb = client.open_by_url(utils.SPREADSHEET_URL)
        df1 = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_DESLIGAMENTOS))
        df2 = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_EQUIPAMENTOS))
        
        if 'IDENTIFICADOR' in df1.columns: df1 = df1[df1['IDENTIFICADOR'] != '']
        if 'IDENTIFICADOR' in df2.columns: df2 = df2[df2['IDENTIFICADOR'] != '']

        df1['Categoria'] = 'DESLIGAMENTOS'
        df2['Categoria'] = 'EQUIPAMENTOS'
        df = pd.concat([df1, df2], ignore_index=True)
        
        # MAPA COMPLETO PARA EVITAR KEYERROR
        mapa_renomear = {
            'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência',
            'ATIVO': 'Ativo', 'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência',
            'QUANTIDADE': 'Quantidade', 'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização',
            'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 'DESCRIÇÃO': 'Descrição',
            'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop',
            'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo', 'CLIENTE AVISADO': 'Cliente Avisado'
        }
        
        renomear_final = {}
        for col in df.columns:
            c_upper = col.strip().upper()
            if c_upper in mapa_renomear:
                renomear_final[col] = mapa_renomear[c_upper]
        
        df.rename(columns=renomear_final, inplace=True)
        df.fillna('', inplace=True)
        
        # Datas
        cols_dt = ['Normalização', 'Desligamento', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']
        for c in cols_dt:
            if c in df.columns: df[c] = pd.to_datetime(df[c], errors='coerce')
            
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

# Verifica coluna crítica
if 'Normalização' not in df_todos.columns:
    st.error("Erro Crítico: Coluna 'Normalização' não encontrada. Verifique a planilha.")
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
            c1.button("Sel. Todos", key="all_anos_h", on_click=set_filtro, args=('hist_anos', opts_ano))
            c2.button("Desmarcar", key="none_anos_h", on_click=set_filtro, args=('hist_anos', []))
            st.session_state.hist_anos = st.multiselect("Selecione", opts_ano, default=st.session_state.hist_anos, label_visibility="collapsed")
            
            c1, c2 = st.columns(2)
            c1.button("Sel. Todos", key="all_mes_h", on_click=set_filtro, args=('hist_meses', meses_cron))
            c2.button("Desmarcar", key="none_mes_h", on_click=set_filtro, args=('hist_meses', []))
            st.session_state.hist_meses = st.multiselect("Selecione", meses_cron, default=st.session_state.hist_meses, label_visibility="collapsed")

        st.markdown("**Filtros Adicionais**")

        def render_sidebar_filter(label, key, col_name):
            if col_name not in df_resolvidas_base.columns: return
            with st.expander(label):
                opts = options_from(df_resolvidas_base[col_name])
                opts_clean = [o for o in opts if o != "-"]
                
                c1, c2 = st.columns(2)
                c1.button("Sel. Todos", key=f'all_{key}', on_click=set_filtro, args=(key, opts_clean))
                c2.button("Desmarcar", key=f'none_{key}', on_click=set_filtro, args=(key, []))
                
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
# CSS KPI (Mantido original)
st.markdown("""
<style>
    .kpi-card { background-color: #333333; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px; }
    .kpi-value { font-size: 3em; font-weight: bold; color: #4b4eff; }
    .kpi-label { font-size: 1.2em; color: #FFFFFF; }
    .card-container { background-color: #089641; color: white; padding: 15px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0 4px 8px rgba(0,0,0,0.2); }
    .card-title { font-size: 1.5em; font-weight: bold; border-bottom: 1px solid rgba(255,255,255,0.5); padding-bottom: 5px; margin-bottom: 10px; }
    .card-item { margin-bottom: 5px; font-size: 1em; }
    .card-label { font-weight: bold; }
    /* Botões Verdes Sidebar */
    div[data-testid="stExpander"] .stButton > button {
        background-color: #28a745; color: white; font-weight: bold; border-radius: 4px; border: none; height: auto; padding: 4px 10px; width: 100%;
    }
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

    # Cards (RESTAURADO HTML COMPLETO)
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
                _, r = rows[i + j]
                with cols[j]:
                    cliente   = html.escape(str(r.get("Cliente", "")))
                    categoria = html.escape(str(r.get("Categoria", "")))
                    ug        = html.escape(str(r.get("UG", "N/A")))
                    tipo      = html.escape(str(r.get("Tipo de ocorrência", "")))
                    ativo     = html.escape(str(r.get("Ativo", "")))
                    nome_ativo= html.escape(str(r.get("Nome Ativo", "")))
                    ocr       = html.escape(str(r.get("Ocorrência", "")))
                    oper      = html.escape(str(r.get("Operador", "")))
                    desc      = html.escape(str(r.get("Descrição", ""))).replace('\n', '<br>')
                    prot      = html.escape(str(r.get("Protocolo", "")))
                    osv       = html.escape(str(r.get("OS", "")))

                    d_des, h_des = fmt_dt(r.get('Desligamento'))
                    d_norm, h_norm = fmt_dt(r.get('Normalização'))
                    d_ca, h_ca = fmt_dt(r.get('Cliente Avisado'))
                    d_loop, h_loop = fmt_dt(r.get('Atendimento Loop'))
                    d_terc, h_terc = fmt_dt(r.get('Atendimento Terceiros'))

                    qtd_html = ''
                    if r.get('Categoria') == 'EQUIPAMENTOS':
                        try:
                            qv = float(r.get('Quantidade', 0))
                            if qv > 0: qtd_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {int(qv)}</div>'
                        except: pass

                    st.markdown(f"""
                    <div class="card-container">
                      <div class="card-title">{ug}</div>
                      <div class="card-item"><span class="card-label">Cliente:</span> {cliente}</div>
                      <div class="card-item"><span class="card-label">Categoria:</span> {categoria}</div>
                      <div class="card-item"><span class="card-label">Tipo:</span> {tipo}</div>
                      <div class="card-item"><span class="card-label">Ativo:</span> {ativo}</div>
                      <div class="card-item"><span class="card-label">Nome:</span> {nome_ativo}</div>
                      <div class="card-item"><span class="card-label">Ocorrência:</span> {ocr}</div>
                      <div class="card-item"><span class="card-label">Operador:</span> {oper}</div>
                      {qtd_html}
                      <br>
                      <div class="card-item"><span class="card-label">Data do desligamento:</span> {d_des}</div>
                      <div class="card-item"><span class="card-label">Hora do desligamento:</span> {h_des}</div>
                      <div class="card-item"><span class="card-label">Data da normalização:</span> {d_norm}</div>
                      <div class="card-item"><span class="card-label">Hora da normalização:</span> {h_norm}</div>
                      <div class="card-item"><span class="card-label">Data cliente avisado:</span> {d_ca}</div>
                      <div class="card-item"><span class="card-label">Hora cliente avisado:</span> {h_ca}</div>
                      <div class="card-item"><span class="card-label">Data atendimento LOOP:</span> {d_loop}</div>
                      <div class="card-item"><span class="card-label">Hora atendimento LOOP:</span> {h_loop}</div>
                      <div class="card-item"><span class="card-label">Data atendimento terceiros:</span> {d_terc}</div>
                      <div class="card-item"><span class="card-label">Hora atendimento terceiros:</span> {h_terc}</div>
                      <br>
                      <div class="card-item"><span class="card-label">Descrição:</span> {desc}</div>
                      <div class="card-item"><span class="card-label">Protocolo:</span> {prot}</div>
                      <div class="card-item"><span class="card-label">OS:</span> {osv}</div>
                    </div>
                    """, unsafe_allow_html=True)
else:
    st.info("Nenhum histórico.")