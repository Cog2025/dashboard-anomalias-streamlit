import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import html
import re
from collections import Counter, defaultdict
import utils

# --- Configuração ---
st.set_page_config(layout="wide", page_title="Dashboard Ocorrências")

if 'ui_phase' not in st.session_state: st.session_state.ui_phase = 'init'
if 'loading_ts' not in st.session_state: st.session_state.loading_ts = 0
if 'categoria_top' not in st.session_state: st.session_state.categoria_top = 'Ambas'

utils.render_loading_overlay(st.session_state.ui_phase)

def start_loading():
    st.session_state.ui_phase = 'loading'
    st.session_state.loading_ts = pytime.time()

if st.session_state.ui_phase == 'init': start_loading()
if st.session_state.ui_phase == 'loading' and (pytime.time() - st.session_state.loading_ts) > 5:
    st.session_state.ui_phase = 'ready'; st.session_state.loading_ts = 0

# --- Helpers de Texto ---
def _collapse_spaces(s: str) -> str: return " ".join(str(s).split())
def canon(s) -> str: return _collapse_spaces(str(s)).casefold() if s is not None else ""

def options_from(series: pd.Series) -> list:
    # Garante conversão para string para evitar erros de comparação
    series = series.astype(str).map(_collapse_spaces)
    unique_vals = sorted([x for x in series.unique() if x and x.lower() != "nan" and x != "" and x != "-"])
    return unique_vals # Removemos o traço daqui para o multiselect ficar limpo

def matches_any_canon(series: pd.Series, selected: list[str]) -> pd.Series:
    if not selected: return pd.Series([True]*len(series), index=series.index)
    # Normaliza tudo para string e canonico para comparação
    sel_c = {canon(s) for s in selected}
    return series.astype(str).map(canon).isin(sel_c)

meses_traducao = {'January': 'Janeiro', 'February': 'Fevereiro', 'March': 'Março', 'April': 'Abril', 'May': 'Maio', 'June': 'Junho', 'July': 'Julho', 'August': 'Agosto', 'September': 'Setembro', 'October': 'Outubro', 'November': 'Novembro', 'December': 'Dezembro'}
meses_cronologicos = list(meses_traducao.values())

# --- CSS Responsivo ---
st.markdown("""
<style>
    /* Botões: Impedir quebra de linha e ajustar tamanho */
    div[data-testid="stExpander"] .stButton > button, 
    div[data-testid="column"] .stButton > button {
        background-color: #28a745; color: white; font-weight: bold;
        border-radius: 5px; width: 100%; border: none;
        white-space: nowrap; padding: 0.5rem 0.5rem;
    }
    .stButton > button:hover { background-color: #218838; }
    
    .kpi-card { background-color: #333333; padding: 15px; border-radius: 10px; text-align: center; margin-bottom: 10px; min-height: 120px; }
    .kpi-value { font-size: 2.5em; font-weight: bold; color: #FF4B4B; }
    .kpi-label { font-size: 1.0em; color: #FFFFFF; }
    
    .card-container { background-color: #FF4B4B; color: white; padding: 15px; border-radius: 8px; margin-bottom: 15px; height: 100%; }
    .card-title { font-size: 1.4em; font-weight: bold; color: white; border-bottom: 1px solid rgba(255,255,255,0.5); padding-bottom: 5px; margin-bottom: 10px; }
    .card-item { margin-bottom: 4px; font-size: 0.95em; }
    .card-label { font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- Dados ---
@st.cache_data(ttl=600)
def carregar_dados(cb):
    try:
        client = utils.connect_to_google_sheets()
        if not client: return pd.DataFrame()
        wb = client.open_by_url(utils.SPREADSHEET_URL)
        df1 = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_DESLIGAMENTOS))
        df2 = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_EQUIPAMENTOS))
        
        # Garante que identificador é string
        if 'IDENTIFICADOR' in df1.columns: df1['IDENTIFICADOR'] = df1['IDENTIFICADOR'].astype(str)
        if 'IDENTIFICADOR' in df2.columns: df2['IDENTIFICADOR'] = df2['IDENTIFICADOR'].astype(str)
        
        df1.dropna(how='all', inplace=True); df2.dropna(how='all', inplace=True)
        df1['Categoria'] = 'DESLIGAMENTOS'; df2['Categoria'] = 'EQUIPAMENTOS'
        df = pd.concat([df1, df2], ignore_index=True)
        
        mapa = {'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência', 'ATIVO': 'Ativo', 'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência', 'QUANTIDADE': 'Quantidade', 'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização', 'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 'DESCRIÇÃO': 'Descrição', 'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop', 'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo', 'CLIENTE AVISADO': 'Cliente Avisado'}
        renomear = {}
        for col in df.columns:
            if col.strip().upper() in mapa: renomear[col] = mapa[col.strip().upper()]
        df.rename(columns=renomear, inplace=True)
        df.fillna('', inplace=True)
        
        # Filtra linhas vazias
        if 'Cliente' in df.columns: df = df[(df['Cliente'] != '')]
        
        for c in ['Normalização', 'Desligamento', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']:
            if c in df.columns: df[c] = pd.to_datetime(df[c], errors='coerce')
            
        if 'Desligamento' in df.columns:
            df['Data'] = df['Desligamento'].dt.strftime('%Y-%m-%d')
            df['Hora'] = df['Desligamento'].dt.strftime('%H:%M:%S')
            df['Mês'] = df['Desligamento'].dt.strftime('%B').map(meses_traducao)
            df['Ano'] = df['Desligamento'].dt.year.fillna(0).astype(int)
            df['Dia'] = df['Desligamento'].dt.day.fillna(0).astype(int)
            df['ID_Unico'] = df['UG'].astype(str).str.upper() + "|" + df['Ativo'].astype(str).str.upper() + "|" + df['Ocorrência'].astype(str).str.upper() + "|" + df['Desligamento'].astype(str)
        return df
    except Exception as e:
        st.error(f"Erro: {e}"); return pd.DataFrame()

if 'cache_buster' not in st.session_state: st.session_state.cache_buster = int(pytime.time())
df_todos = carregar_dados(st.session_state.cache_buster)

if st.session_state.ui_phase == 'loading': 
    st.session_state.ui_phase = 'ready'; st.session_state.loading_ts = 0; utils.render_loading_overlay('ready')

df_todos['Desligamento'] = pd.to_datetime(df_todos['Desligamento'], errors='coerce')

# KPIs
c_deslig = df_todos[(df_todos['Categoria'] == 'DESLIGAMENTOS') & (pd.isna(df_todos['Normalização']) | (df_todos['Normalização'] == ''))].shape[0] if not df_todos.empty else 0
c_equip = df_todos[(df_todos['Categoria'] == 'EQUIPAMENTOS') & (pd.isna(df_todos['Normalização']) | (df_todos['Normalização'] == ''))].shape[0] if not df_todos.empty else 0

with st.container(border=True):
    st.markdown("<h1 style='margin:0; text-align:center;'>OCORRÊNCIAS ATIVAS</h1>", unsafe_allow_html=True)
    st.write("")
    c1, c2 = st.columns(2)
    c1.markdown(f"<div class='kpi-card'><div class='kpi-label'>USINAS DESLIGADAS</div><div class='kpi-value'>{c_deslig}</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='kpi-card'><div class='kpi-label'>EQUIPAMENTOS PARADOS</div><div class='kpi-value'>{c_equip}</div></div>", unsafe_allow_html=True)

# Filtros Globais
if 'filtros_meses' not in st.session_state: st.session_state.filtros_meses = [meses_cronologicos[datetime.now().month - 1]]
if 'filtros_anos' not in st.session_state:
    anos = sorted(df_todos['Ano'].unique().tolist()) if not df_todos.empty else []
    st.session_state.filtros_anos = [a for a in anos if a != 0]

st.header('OCORRÊNCIAS FILTRADAS')
total_db = df_todos[(pd.isna(df_todos['Normalização']) | (df_todos['Normalização'] == ''))].shape[0] if not df_todos.empty else 0
c1, c2 = st.columns([1, 1])
c1.markdown(f"<div class='kpi-card'><div class='kpi-label'>Total no Banco (Abertas)</div><div class='kpi-value'>{total_db}</div></div>", unsafe_allow_html=True)
with c2:
    st.write(""); 
    if st.button("Atualizar Dados"): st.cache_data.clear(); st.session_state.ui_phase = 'init'; st.rerun()

# --- ÁREA DE FILTROS ---
def set_filtro(key, values):
    st.session_state[key] = values
    start_loading()

def marcar_e_loading(prefixo_key, itens, filtro_key, marcar_todos):
    st.session_state[filtro_key] = list(itens) if marcar_todos else []
    for x in itens: st.session_state[f"{prefixo_key}{x}"] = marcar_todos
    start_loading()

if not df_todos.empty:
    st.markdown("#### Filtrar por categoria")
    st.radio("Categoria:", ["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"], horizontal=True, key="categoria_top", on_change=start_loading, label_visibility="collapsed")

    st.subheader("Selecione o período")
    anos_disp = sorted([a for a in df_todos['Ano'].unique() if a != 0])
    
    # Colunas de Data
    c_ano, c_mes, c_dia = st.columns([1, 1, 2])
    
    with c_ano:
        with st.container(border=True):
            st.write("**Ano(s):**")
            with st.expander("Expandir"):
                for ano in anos_disp: st.checkbox(str(ano), key=f'cb_ano_{ano}', value=(ano in st.session_state.filtros_anos))
            b1, b2 = st.columns(2)
            b1.button('Todos', key='sel_ano', on_click=marcar_e_loading, args=('cb_ano_', anos_disp, 'filtros_anos', True))
            b2.button('Limpar', key='des_ano', on_click=marcar_e_loading, args=('cb_ano_', anos_disp, 'filtros_anos', False))
            st.session_state.filtros_anos = [a for a in anos_disp if st.session_state.get(f'cb_ano_{a}', False)]

    with c_mes:
        with st.container(border=True):
            st.write("**Mês(es):**")
            with st.expander("Expandir"):
                for mes in meses_cronologicos: st.checkbox(mes, key=f'cb_mes_{mes}', value=(mes in st.session_state.filtros_meses))
            b1, b2 = st.columns(2)
            b1.button('Todos', key='sel_mes', on_click=marcar_e_loading, args=('cb_mes_', meses_cronologicos, 'filtros_meses', True))
            b2.button('Limpar', key='des_mes', on_click=marcar_e_loading, args=('cb_mes_', meses_cronologicos, 'filtros_meses', False))
            st.session_state.filtros_meses = [m for m in meses_cronologicos if st.session_state.get(f'cb_mes_{m}', False)]

    with c_dia:
        with st.container(border=True):
            st.write("**Dia(s):**") # CORREÇÃO DO NOME (Era numeric Dias)
            if 'filtros_dias' not in st.session_state: st.session_state.filtros_dias = []
            df_d = df_todos[(df_todos['Ano'].isin(st.session_state.filtros_anos)) & (df_todos['Mês'].isin(st.session_state.filtros_meses))]
            dias_disp = sorted(df_d['Dia'].unique().astype(int).tolist()) if not df_d.empty else []
            dias_disp = [d for d in dias_disp if d != 0]
            
            with st.expander("Expandir"):
                cd = st.columns(7)
                for i, dia in enumerate(range(1, 32)):
                    with cd[i%7]:
                        dis = dia not in dias_disp
                        st.checkbox(str(dia), key=f'cb_dia_{dia}', value=(dia in st.session_state.filtros_dias) if not dis else False, disabled=dis)
            
            b1, b2 = st.columns(2)
            b1.button('Todos', key='sel_dia', on_click=marcar_e_loading, args=('cb_dia_', list(range(1,32)), 'filtros_dias', True))
            b2.button('Limpar', key='des_dia', on_click=marcar_e_loading, args=('cb_dia_', list(range(1,32)), 'filtros_dias', False))
            st.session_state.filtros_dias = [d for d in dias_disp if st.session_state.get(f'cb_dia_{d}', False)]

    # Filtros Adicionais - Layout em 2 Linhas para Responsividade
    st.subheader("Filtros Adicionais")
    for k in ['filtros_clientes', 'filtros_ugs', 'filtros_tipos', 'filtros_ativos', 'filtros_ocorrencias']:
        if k not in st.session_state: st.session_state[k] = []

    # Função Helper para renderizar bloco de filtro
    def render_filter_block(col, label, key, options, k_s, k_d):
        with col:
            with st.container(border=True):
                st.write(f"**{label}**")
                b1, b2 = st.columns(2)
                # Selecionar Todos
                if b1.button('Todos', key=k_s): 
                    st.session_state[key] = options; st.rerun()
                # Limpar
                if b2.button('Limpar', key=k_d): 
                    st.session_state[key] = []; st.rerun()
                
                # Multiselect
                st.session_state[key] = st.multiselect(" ", options, default=[x for x in st.session_state[key] if x in options], key=f"ms_{key}", label_visibility="collapsed")

    # Linha 1 (3 colunas)
    row1 = st.columns(3)
    render_filter_block(row1[0], "Cliente", 'filtros_clientes', options_from(df_todos['Cliente']), 's_cli', 'd_cli')
    render_filter_block(row1[1], "UG", 'filtros_ugs', sorted([u for u in df_todos['UG'].unique().tolist() if u]), 's_ug', 'd_ug')
    render_filter_block(row1[2], "Tipo", 'filtros_tipos', options_from(df_todos['Tipo de ocorrência']), 's_tip', 'd_tip')
    
    # Linha 2 (2 colunas) - Para dar espaço e não espremer
    row2 = st.columns(2)
    render_filter_block(row2[0], "Ativo", 'filtros_ativos', options_from(df_todos['Ativo']), 's_atv', 'd_atv')
    render_filter_block(row2[1], "Ocorrência", 'filtros_ocorrencias', options_from(df_todos['Ocorrência']), 's_ocr', 'd_ocr')

    # Aplicação dos Filtros
    m_cat = pd.Series([True]*len(df_todos))
    if st.session_state.categoria_top != "Ambas": m_cat = df_todos['Categoria'] == st.session_state.categoria_top
    
    m_ano = df_todos['Ano'].isin(st.session_state.filtros_anos)
    m_mes = df_todos['Mês'].isin(st.session_state.filtros_meses)
    m_dia = df_todos['Dia'].isin(st.session_state.filtros_dias) if st.session_state.filtros_dias else pd.Series([True]*len(df_todos))
    
    m_cli = matches_any_canon(df_todos['Cliente'], st.session_state.filtros_clientes)
    m_ug = df_todos['UG'].isin(st.session_state.filtros_ugs) if st.session_state.filtros_ugs else pd.Series([True]*len(df_todos))
    m_tip = matches_any_canon(df_todos['Tipo de ocorrência'], st.session_state.filtros_tipos)
    m_atv = matches_any_canon(df_todos['Ativo'], st.session_state.filtros_ativos)
    m_ocr = matches_any_canon(df_todos['Ocorrência'], st.session_state.filtros_ocorrencias)

    df_filt = df_todos[m_cat & m_ano & m_mes & m_dia & m_cli & m_ug & m_tip & m_atv & m_ocr].copy()
    df_abertas = df_filt[pd.isna(df_filt['Normalização']) | (df_filt['Normalização'] == '')].copy()

    with c2: st.markdown(f"<div class='kpi-card'><div class='kpi-label'>Total com Filtro Selecionado</div><div class='kpi-value'>{len(df_abertas)}</div></div>", unsafe_allow_html=True)

    if not df_abertas.empty:
        mask = df_abertas['Desligamento'].notna()
        df_abertas.loc[mask, 'Tempo em Segundos'] = (datetime.now() - df_abertas.loc[mask, 'Desligamento']).dt.total_seconds().astype(int)
        df_abertas.loc[~mask, 'Tempo em Segundos'] = 0

        st.markdown("---")
        c1, c2 = st.columns(2)
        sort_by = c1.selectbox("Ordenar por:", ["Data do Desligamento", "Tempo de Desligamento", "UG", "Ativo"])
        order = c2.radio("Ordem:", ["Descendente", "Ascendente"], horizontal=True)
        
        df_sorted = df_abertas.sort_values(by={'Data do Desligamento': 'Desligamento', 'Tempo de Desligamento': 'Tempo em Segundos', 'UG': 'UG', 'Ativo': 'Ativo'}[sort_by], ascending=(order == "Ascendente"))
        
        # Edição
        df_sorted['Display'] = df_sorted['UG'].astype(str) + " | " + df_sorted['Ocorrência'].astype(str) + " | " + df_sorted['Desligamento'].dt.strftime('%d/%m %H:%M').fillna('')
        st.session_state['df_lista_para_editar'] = df_sorted.copy()
        
        sel = st.selectbox("Selecione para editar:", df_sorted['Display'].tolist(), index=None)
        if sel: st.session_state['id_unico_para_editar'] = df_sorted.loc[df_sorted['Display'] == sel, 'ID_Unico'].values[0]
        if st.button("📝 Ir para Edição", disabled=not bool(sel)): st.switch_page("pages/3_Editar_Ocorrência.py")

        # Tabela
        def fmt_t(r):
            s = r['Tempo em Segundos']; d, r = divmod(s, 86400); h, r = divmod(r, 3600); m, s = divmod(r, 60)
            return f"{int(d)}d {int(h)}h {int(m)}m"
        df_sorted['Tempo de Desligamento'] = df_sorted.apply(fmt_t, axis=1)
        st.dataframe(df_sorted[['Categoria', 'Tempo de Desligamento', 'UG', 'Data', 'Hora', 'Tipo de ocorrência', 'Ativo', 'Ocorrência', 'Descrição']], use_container_width=True)

        # Cards
        st.markdown("### Detalhes (Cards)")
        num_cols = 4
        rows = list(df_sorted.iterrows())
        def fmt_dt(dt): return dt.strftime('%d/%m/%Y'), dt.strftime('%H:%M') if pd.notna(dt) else ('', '')

        for i in range(0, len(rows), num_cols):
            cols = st.columns(num_cols)
            for j in range(num_cols):
                if i + j < len(rows):
                    _, r = rows[i+j]
                    with cols[j]:
                        cli = html.escape(str(r.get("Cliente", "")))
                        ug = html.escape(str(r.get("UG", "N/A")))
                        ocr = html.escape(str(r.get("Ocorrência", "")))
                        desc = html.escape(str(r.get("Descrição", ""))).replace('\n', '<br>')
                        dd, hd = fmt_dt(r.get('Desligamento'))
                        
                        st.markdown(f"""
                        <div class="card-container">
                            <div class="card-title">{ug}</div>
                            <div class="card-item"><span class="card-label">Cliente:</span> {cli}</div>
                            <div class="card-item"><span class="card-label">Categoria:</span> {html.escape(str(r.get('Categoria', '')))}</div>
                            <div class="card-item"><span class="card-label">Ocorrência:</span> {ocr}</div>
                            <div class="card-item"><span class="card-label">Data:</span> {dd} {hd}</div>
                            <br>
                            <div class="card-item"><span class="card-label">Descrição:</span> {desc}</div>
                        </div>
                        """, unsafe_allow_html=True)
    else:
        st.info("Nenhuma ocorrência encontrada.")