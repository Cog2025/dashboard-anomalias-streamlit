import os
import io
import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import html
import gspread
from google.oauth2.service_account import Credentials
import re
from collections import Counter, defaultdict
import utils

# --- 1. Configuração da Página e Layout ---
st.set_page_config(layout="wide", page_title="Dashboard Ocorrências")

# Estado do overlay (Inicialização Segura)
if 'ui_phase' not in st.session_state:
    st.session_state['ui_phase'] = 'init'
if 'loading_ts' not in st.session_state:
    st.session_state['loading_ts'] = 0
if 'categoria_top' not in st.session_state:
    st.session_state['categoria_top'] = 'Ambas'

# Renderiza Overlay
utils.render_loading_overlay(st.session_state['ui_phase'])

def start_loading():
    st.session_state['ui_phase'] = 'loading'
    st.session_state['loading_ts'] = pytime.time()

if st.session_state['ui_phase'] == 'init':
    start_loading()

# Failsafe para destravar loading
if st.session_state['ui_phase'] == 'loading' and (pytime.time() - st.session_state['loading_ts']) > 20:
    st.session_state['ui_phase'] = 'ready'
    st.session_state['loading_ts'] = 0

# --- Helpers de Texto ---
def _collapse_spaces(s: str) -> str:
    return " ".join(str(s).split())

def canon(s) -> str:
    if s is None: return ""
    return _collapse_spaces(str(s)).casefold()

def build_display_map(series: pd.Series) -> dict:
    buckets = defaultdict(Counter)
    for v in series.dropna():
        v_str = _collapse_spaces(str(v))
        if not v_str: continue
        buckets[canon(v_str)][v_str] += 1
    display_map = {}
    for ckey, counter in buckets.items():
        best, _ = counter.most_common(1)[0]
        display_map[ckey] = best
    return display_map

def options_from(series: pd.Series) -> list:
    series = series.astype(str).map(_collapse_spaces)
    series = series[series != ""]
    dmap = build_display_map(series)
    labels = {dmap[canon(v)] for v in series}
    return ["-"] + sorted(labels)

def matches_any_canon(series: pd.Series, selected: list[str]) -> pd.Series:
    if not selected:
        return pd.Series([True]*len(series), index=series.index)
    sel_c = {canon(s) for s in selected if s and s != "-"}
    return series.astype(str).map(canon).isin(sel_c)

# Helper de checkbox/loading
def marcar_e_loading(prefixo_key, itens, filtro_key, marcar_todos):
    st.session_state[filtro_key] = list(itens) if marcar_todos else []
    for x in itens:
        st.session_state[f"{prefixo_key}{x}"] = marcar_todos
    start_loading()

meses_traducao = {
    'January': 'Janeiro', 'February': 'Fevereiro', 'March': 'Março',
    'April': 'Abril', 'May': 'Maio', 'June': 'Junho',
    'July': 'Julho', 'August': 'Agosto', 'September': 'Setembro',
    'October': 'Outubro', 'November': 'Novembro', 'December': 'Dezembro'
}
meses_cronologicos = list(meses_traducao.values())

# --- CSS Personalizado ---
st.markdown("""
<style>
    /* Botões responsivos */
    div[data-testid="stExpander"] .stButton > button {
        background-color: #28a745; 
        color: white; 
        font-weight: bold;
        border-radius: 5px;
        width: 100%;
        border: none;
        white-space: nowrap; 
        padding: 0.5rem 0.5rem;
    }
    .stButton > button:hover { background-color: #218838; }
    
    /* Cards de KPI */
    .kpi-card {
        background-color: #333333; 
        padding: 15px; 
        border-radius: 10px;
        text-align: center; 
        margin-bottom: 10px;
        min-height: 120px;
    }
    .kpi-value { font-size: 2.5em; font-weight: bold; color: #FF4B4B; }
    .kpi-label { font-size: 1.0em; color: #FFFFFF; }
    
    /* Cards de Ocorrência */
    .card-container {
        background-color: #FF4B4B;
        color: white;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 15px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        height: 100%;
    }
    .card-title {
        font-size: 1.4em; font-weight: bold; color: white;
        border-bottom: 1px solid rgba(255,255,255,0.5);
        padding-bottom: 5px; margin-bottom: 10px;
    }
    .card-item { margin-bottom: 4px; font-size: 0.95em; }
    .card-label { font-weight: bold; }
    
    /* Ajuste para evitar colunas muito finas */
    div[data-testid="column"] { min-width: 100px; }
</style>
""", unsafe_allow_html=True)

# --- Carregamento de Dados ---
@st.cache_data(ttl=600)
def carregar_dados_google_sheets(cache_buster: int = 0):
    try:
        client = utils.connect_to_google_sheets()
        if not client: return pd.DataFrame()

        workbook = client.open_by_url(utils.SPREADSHEET_URL)
        df1 = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_DESLIGAMENTOS))
        df2 = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_EQUIPAMENTOS))

        if 'IDENTIFICADOR' in df1.columns:
            df1['IDENTIFICADOR'] = df1['IDENTIFICADOR'].astype(str)
        if 'IDENTIFICADOR' in df2.columns:
            df2['IDENTIFICADOR'] = df2['IDENTIFICADOR'].astype(str)

        df1.dropna(how='all', inplace=True)
        df2.dropna(how='all', inplace=True)
        
        df1['Categoria'] = 'DESLIGAMENTOS'
        df2['Categoria'] = 'EQUIPAMENTOS'
        df_todos = pd.concat([df1, df2], ignore_index=True)

        mapa = {
            'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência',
            'ATIVO': 'Ativo', 'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência', 'QUANTIDADE': 'Quantidade', 
            'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização', 'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 
            'DESCRIÇÃO': 'Descrição', 'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop',
            'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo', 'CLIENTE AVISADO': 'Cliente Avisado'
        }
        
        renomear = {}
        for col in df_todos.columns:
            c_upper = col.strip().upper()
            if c_upper in mapa: renomear[col] = mapa[c_upper]
        df_todos.rename(columns=renomear, inplace=True)
        df_todos.fillna('', inplace=True)
        
        if 'Cliente' in df_todos.columns:
            df_todos = df_todos[(df_todos['Cliente'] != '') & (df_todos['UG'] != '')]

        cols_dt = ['Normalização', 'Desligamento', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']
        for c in cols_dt:
            if c in df_todos.columns: df_todos[c] = pd.to_datetime(df_todos[c], errors='coerce')

        if 'Desligamento' in df_todos.columns:
            df_todos['Data'] = df_todos['Desligamento'].dt.strftime('%Y-%m-%d')
            df_todos['Hora'] = df_todos['Desligamento'].dt.strftime('%H:%M:%S')
            df_todos['Mês']  = df_todos['Desligamento'].dt.strftime('%B').map(meses_traducao)
            df_todos['Ano']  = df_todos['Desligamento'].dt.year.fillna(0).astype(int)
            df_todos['Dia']  = df_todos['Desligamento'].dt.day.fillna(0).astype(int)
            
            df_todos['ID_Unico'] = (
                df_todos['UG'].astype(str).str.upper() + "|" +
                df_todos['Ativo'].astype(str).str.upper() + "|" +
                df_todos['Ocorrência'].astype(str).str.upper() + "|" +
                df_todos['Desligamento'].astype(str)
            )
        return df_todos

    except Exception as e:
        st.error(f"Erro: {e}")
        return pd.DataFrame()

if 'cache_buster' not in st.session_state:
    st.session_state.cache_buster = int(pytime.time())

df_todos_dados = carregar_dados_google_sheets(st.session_state.cache_buster)

if st.session_state['ui_phase'] == 'loading':
    st.session_state['ui_phase'] = 'ready'
    st.session_state['loading_ts'] = 0
    utils.render_loading_overlay('ready')

df_todos_dados['Desligamento'] = pd.to_datetime(df_todos_dados['Desligamento'], errors='coerce')

# --- KPIs do Topo (Main Area) ---
count_deslig = df_todos_dados[(df_todos_dados['Categoria'] == 'DESLIGAMENTOS') & (pd.isna(df_todos_dados['Normalização']) | (df_todos_dados['Normalização'] == ''))].shape[0] if not df_todos_dados.empty else 0
count_equip = df_todos_dados[(df_todos_dados['Categoria'] == 'EQUIPAMENTOS') & (pd.isna(df_todos_dados['Normalização']) | (df_todos_dados['Normalização'] == ''))].shape[0] if not df_todos_dados.empty else 0

with st.container(border=True):
    st.markdown("<h1 style='margin:0; text-align:center;'>OCORRÊNCIAS ATIVAS</h1>", unsafe_allow_html=True)
    st.write("")
    c1, c2 = st.columns(2)
    c1.markdown(f"<div class='kpi-card'><div class='kpi-label'>USINAS DESLIGADAS</div><div class='kpi-value'>{count_deslig}</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='kpi-card'><div class='kpi-label'>EQUIPAMENTOS PARADOS</div><div class='kpi-value'>{count_equip}</div></div>", unsafe_allow_html=True)

# ==============================================================================
# --- BARRA LATERAL (SIDEBAR) ---
# ==============================================================================
with st.sidebar:
    st.title("Filtros")
    
    # Botão de Atualizar
    if st.button("🔄 Atualizar Dados"):
        st.cache_data.clear()
        st.session_state['ui_phase'] = 'init'
        st.rerun()
    
    st.markdown("---")

    # --- Filtro Categoria ---
    st.markdown("**Categoria**")
    st.radio("Categoria:", ["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"], horizontal=True, key="categoria_top", on_change=start_loading, label_visibility="collapsed")
    
    st.markdown("---")
    st.markdown("**Período**")

    # Inicialização Filtros
    if 'filtros_meses' not in st.session_state: st.session_state.filtros_meses = [meses_cronologicos[datetime.now().month - 1]]
    if 'filtros_anos' not in st.session_state:
        anos = sorted(df_todos_dados['Ano'].unique().tolist()) if not df_todos_dados.empty else []
        st.session_state.filtros_anos = [a for a in anos if a != 0]

    anos_disp = sorted([a for a in df_todos_dados['Ano'].unique() if a != 0]) if not df_todos_dados.empty else []

    # 1. ANOS (Sidebar Expander)
    with st.expander("📅 Anos", expanded=True):
        if anos_disp:
            for ano in anos_disp:
                st.checkbox(str(ano), key=f'cb_ano_{ano}', value=(ano in st.session_state.filtros_anos))
            
            b1, b2 = st.columns(2)
            b1.button('Todos', key='sel_ano', on_click=marcar_e_loading, args=('cb_ano_', anos_disp, 'filtros_anos', True))
            b2.button('Limpar', key='des_ano', on_click=marcar_e_loading, args=('cb_ano_', anos_disp, 'filtros_anos', False))
            
            st.session_state.filtros_anos = [a for a in anos_disp if st.session_state.get(f'cb_ano_{a}', False)]
        else:
            st.write("Sem dados.")

    # 2. MESES (Sidebar Expander)
    with st.expander("📆 Meses", expanded=False):
        for mes in meses_cronologicos:
            st.checkbox(mes, key=f'cb_mes_{mes}', value=(mes in st.session_state.filtros_meses))
        
        b1, b2 = st.columns(2)
        b1.button('Todos', key='sel_mes', on_click=marcar_e_loading, args=('cb_mes_', meses_cronologicos, 'filtros_meses', True))
        b2.button('Limpar', key='des_mes', on_click=marcar_e_loading, args=('cb_mes_', meses_cronologicos, 'filtros_meses', False))
        
        st.session_state.filtros_meses = [m for m in meses_cronologicos if st.session_state.get(f'cb_mes_{m}', False)]

    # 3. DIAS (Sidebar Expander)
    with st.expander("numeric Dias", expanded=False):
        if 'filtros_dias' not in st.session_state: st.session_state.filtros_dias = []
        
        dias_disp = []
        if not df_todos_dados.empty:
            df_dates = df_todos_dados[(df_todos_dados['Ano'].isin(st.session_state.filtros_anos)) & (df_todos_dados['Mês'].isin(st.session_state.filtros_meses))]
            dias_disp = sorted(df_dates['Dia'].unique().astype(int).tolist())
            dias_disp = [d for d in dias_disp if d != 0]
        
        cols_d = st.columns(4) 
        for i, dia in enumerate(range(1, 32)):
            with cols_d[i % 4]:
                if dia in dias_disp:
                    st.checkbox(str(dia), key=f'cb_dia_{dia}', value=(dia in st.session_state.filtros_dias))
                else:
                    st.checkbox(str(dia), key=f'cb_dia_{dia}', disabled=True)
        
        b1, b2 = st.columns(2)
        b1.button('Todos', key='sel_dia', on_click=marcar_e_loading, args=('cb_dia_', list(range(1,32)), 'filtros_dias', True))
        b2.button('Limpar', key='des_dia', on_click=marcar_e_loading, args=('cb_dia_', list(range(1,32)), 'filtros_dias', False))
        
        st.session_state.filtros_dias = [d for d in dias_disp if st.session_state.get(f'cb_dia_{d}', False)]

    st.markdown("---")
    st.markdown("**Filtros Adicionais**")

    # Inicializa
    for k in ['filtros_clientes', 'filtros_ugs', 'filtros_tipos', 'filtros_ativos', 'filtros_ocorrencias']:
        if k not in st.session_state: st.session_state[k] = []

    def render_sidebar_multiselect(label, key, options, sel_key, des_key):
        with st.expander(label):
            opts_clean = [o for o in options if o != "-"]
            
            b1, b2 = st.columns(2)
            if b1.button('Todos', key=sel_key):
                 st.session_state[key] = opts_clean
                 st.rerun()
            if b2.button('Limpar', key=des_key):
                 st.session_state[key] = []
                 st.rerun()
            
            # Filtra defaults validos
            valid_defaults = [x for x in st.session_state[key] if x in opts_clean]
            st.session_state[key] = st.multiselect("Selecione", opts_clean, default=valid_defaults, key=f"ms_{key}", label_visibility="collapsed")

    if not df_todos_dados.empty:
        cli_opts = options_from(df_todos_dados['Cliente'])
        ug_opts = sorted([u for u in df_todos_dados['UG'].unique().tolist() if u]) 
        
        render_sidebar_multiselect("Clientes", 'filtros_clientes', cli_opts, 'sel_cli', 'des_cli')
        render_sidebar_multiselect("UGs", 'filtros_ugs', ug_opts, 'sel_ug', 'des_ug')
        render_sidebar_multiselect("Tipo", 'filtros_tipos', options_from(df_todos_dados['Tipo de ocorrência']), 'sel_tip', 'des_tip')
        render_sidebar_multiselect("Ativo", 'filtros_ativos', options_from(df_todos_dados['Ativo']), 'sel_atv', 'des_atv')
        render_sidebar_multiselect("Ocorrência", 'filtros_ocorrencias', options_from(df_todos_dados['Ocorrência']), 'sel_ocr', 'des_ocr')

        # ==============================================================================
# --- ÁREA PRINCIPAL (Main Area) ---
# ==============================================================================

# Aplicação dos Filtros (LÓGICA ORIGINAL)
m_ano = df_todos_dados['Ano'].isin(st.session_state.filtros_anos)
m_mes = df_todos_dados['Mês'].isin(st.session_state.filtros_meses)
m_dia = df_todos_dados['Dia'].isin(st.session_state.filtros_dias) if st.session_state.filtros_dias else pd.Series([True]*len(df_todos_dados))

m_cat = pd.Series([True]*len(df_todos_dados))
if st.session_state.categoria_top != "Ambas":
    m_cat = df_todos_dados['Categoria'] == st.session_state.categoria_top

m_cli = matches_any_canon(df_todos_dados['Cliente'], st.session_state.filtros_clientes)
m_ug = df_todos_dados['UG'].isin(st.session_state.filtros_ugs) if st.session_state.filtros_ugs else pd.Series([True]*len(df_todos_dados))
m_tip = matches_any_canon(df_todos_dados['Tipo de ocorrência'], st.session_state.filtros_tipos)
m_atv = matches_any_canon(df_todos_dados['Ativo'], st.session_state.filtros_ativos)
m_ocr = matches_any_canon(df_todos_dados['Ocorrência'], st.session_state.filtros_ocorrencias)

df_filt = df_todos_dados[m_cat & m_ano & m_mes & m_dia & m_cli & m_ug & m_tip & m_atv & m_ocr].copy()
df_abertas = df_filt[pd.isna(df_filt['Normalização']) | (df_filt['Normalização'] == '')].copy()

# KPI Filtrado
st.markdown("---")
st.markdown("### Resultados Filtrados")
total_db = df_todos_dados[(pd.isna(df_todos_dados['Normalização']) | (df_todos_dados['Normalização'] == ''))].shape[0] if not df_todos_dados.empty else 0

c_kpi1, c_kpi2 = st.columns(2)
c_kpi1.markdown(f"<div class='kpi-card'><div class='kpi-label'>Total no Banco (Abertas)</div><div class='kpi-value'>{total_db}</div></div>", unsafe_allow_html=True)
c_kpi2.markdown(f"<div class='kpi-card'><div class='kpi-label'>Total com Filtro Selecionado</div><div class='kpi-value'>{len(df_abertas)}</div></div>", unsafe_allow_html=True)

if not df_abertas.empty:
    mask_valid = df_abertas['Desligamento'].notna()
    df_abertas.loc[mask_valid, 'Tempo em Segundos'] = (
        (datetime.now() - df_abertas.loc[mask_valid, 'Desligamento']).dt.total_seconds().astype(int)
    )
    df_abertas.loc[~mask_valid, 'Tempo em Segundos'] = 0

    st.markdown("---")
    
    # Ordenação
    col_ctrl1, col_ctrl2 = st.columns([2, 2])
    with col_ctrl1:
        sort_by = st.selectbox("Ordenar por:", ["Data do Desligamento", "Tempo de Desligamento", "UG", "Ativo"])
    with col_ctrl2:
        sort_order = st.radio("Ordem:", ["Descendente", "Ascendente"], horizontal=True)
    
    asc = sort_order == "Ascendente"
    col_map = {'Data do Desligamento': 'Desligamento', 'Tempo de Desligamento': 'Tempo em Segundos', 'UG': 'UG', 'Ativo': 'Ativo'}
    df_sorted = df_abertas.sort_values(by=col_map[sort_by], ascending=asc)

    # Seletor Edição
    df_sorted['Display'] = (
        df_sorted['UG'].astype(str) + " | " + df_sorted['Ativo'].astype(str) + " | " +
        df_sorted['Nome Ativo'].astype(str) + " | " + df_sorted['Ocorrência'].astype(str) + " | " +
        df_sorted['Desligamento'].dt.strftime('%d/%m/%Y %H:%M').fillna('') + 
        "  ·  " + df_sorted['ID_Unico'].astype(str).str[-6:]
    )
    
    st.session_state['df_lista_para_editar'] = df_sorted.copy()
    
    st.markdown("### Editar Ocorrência")
    opts = df_sorted['Display'].tolist()
    sel = st.selectbox("Selecione na lista abaixo:", options=opts, index=None, placeholder="Escolha uma ocorrência para editar...")
    
    if sel:
            id_unico = df_sorted.loc[df_sorted['Display'] == sel, 'ID_Unico'].values[0]
            st.session_state['id_unico_para_editar'] = id_unico
    elif 'id_unico_para_editar' in st.session_state:
            st.session_state.pop('id_unico_para_editar')
    
    if st.button("📝 Ir para Edição", disabled=not bool(sel)):
            st.switch_page("pages/3_Editar_Ocorrência.py")

    # Tabela
    st.markdown("### Lista de Ocorrências")
    def fmt_tempo(row):
        s = row['Tempo em Segundos']
        d, r = divmod(s, 86400); h, r = divmod(r, 3600); m, s = divmod(r, 60)
        return f"{int(d)}d {int(h)}h {int(m)}m"
    
    df_tab = df_sorted.copy()
    df_tab['Tempo de Desligamento'] = df_tab.apply(fmt_tempo, axis=1)
    st.dataframe(df_tab[['Categoria', 'Tempo de Desligamento', 'UG', 'Data', 'Hora', 'Tipo de ocorrência', 'Ativo', 'Ocorrência', 'Operador', 'Descrição', 'OS']], use_container_width=True)

    # Cards (Mesmo código visual detalhado)
    st.markdown("### Detalhes (Cards)")
    num_cols = 4
    rows = list(df_sorted.iterrows())
    
    def fmt_dt(dt):
        if pd.notna(dt): return dt.strftime('%d/%m/%Y'), dt.strftime('%H:%M')
        return '', ''

    for i in range(0, len(rows), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j < len(rows):
                _, r = rows[i + j]
                with cols[j]:
                    cli = html.escape(str(r.get("Cliente", "")))
                    cat = html.escape(str(r.get("Categoria", "")))
                    ug = html.escape(str(r.get("UG", "N/A")))
                    tipo = html.escape(str(r.get("Tipo de ocorrência", "")))
                    ativo = html.escape(str(r.get("Ativo", "")))
                    nome = html.escape(str(r.get("Nome Ativo", "")))
                    ocr = html.escape(str(r.get("Ocorrência", "")))
                    oper = html.escape(str(r.get("Operador", "")))
                    desc = html.escape(str(r.get("Descrição", ""))).replace('\n', '<br>')
                    prot = html.escape(str(r.get("Protocolo", "")))
                    osv = html.escape(str(r.get("OS", "")))
                    
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
                        <div class="card-item"><span class="card-label">Descrição:</span> {desc}</div>
                        <div class="card-item"><span class="card-label">Protocolo:</span> {prot}</div>
                        <div class="card-item"><span class="card-label">OS:</span> {osv}</div>
                    </div>
                    """, unsafe_allow_html=True)
else:
    st.info("Nenhuma usina encontrada com os filtros atuais.")