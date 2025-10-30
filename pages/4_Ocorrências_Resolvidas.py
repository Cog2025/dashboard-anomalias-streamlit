import os
from webdav3.client import Client  # opcional se não usa aqui
import io                          # opcional se não usa aqui
import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import html
import gspread
from google.oauth2.service_account import Credentials
import re
import unicodedata
from collections import Counter, defaultdict

# --- 1. Configuração da Página e Layout ---
st.set_page_config(layout="wide")

# Estado do overlay
if 'ui_phase' not in st.session_state:
    st.session_state.ui_phase = 'init'
if 'loading_ts' not in st.session_state:
    st.session_state.loading_ts = 0

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
      @keyframes bgIn   {{ to {{ background: rgba(0,0,0,.55); }} }}
      @keyframes spin   {{ to {{ transform: rotate(360deg); }} }}
    </style>
    <div id="__overlay__"><div class="loader"></div></div>
    """, unsafe_allow_html=True)

# Injetar overlay no DOM o quanto antes
render_loading_overlay(st.session_state.ui_phase)

# Promover init -> loading na 1ª passada e rerodar
if st.session_state.ui_phase == 'init':
    st.session_state.ui_phase = 'loading'
    st.session_state.loading_ts = pytime.time()
    st.rerun()

# Helpers (chame-os nos gatilhos pesados)
def start_loading():
    st.session_state.ui_phase = 'loading'
    st.session_state.loading_ts = pytime.time()
    render_loading_overlay('loading')

def stop_loading():
    st.session_state.ui_phase = 'ready'
    st.session_state.loading_ts = 0
    render_loading_overlay('ready')

def marcar_loading(prefixo_key, itens, filtro_key, marcar_todos, validos=None):
    start_loading()
    _marcar(prefixo_key, itens, filtro_key, marcar_todos, validos)

# Failsafe opcional (20s)
if st.session_state.ui_phase == 'loading' and (pytime.time() - st.session_state.loading_ts) > 20:
    stop_loading()

def _collapse_spaces(s: str) -> str:
    return re.sub(r'\s+', ' ', str(s)).strip()

def canon(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = _collapse_spaces(s)
    # case-insensitive robusto
    s = s.casefold()
    return s

def build_display_map(series: pd.Series) -> dict:
    # mapeia forma canônica -> rótulo preferido (mais frequente)
    buckets = defaultdict(Counter)
    for v in series.dropna():
        v_str = _collapse_spaces(str(v))
        if not v_str:
            continue
        buckets[canon(v_str)][v_str] += 1
    display_map = {}
    for ckey, counter in buckets.items():
        # rótulo preferido = o mais frequente
        best, _ = counter.most_common(1)[0]
        display_map[ckey] = best
    return display_map

def options_from(series: pd.Series) -> list[str]:
    ser = series.astype(str).map(_collapse_spaces)
    ser = ser[(ser != "") & (ser != "-") & (ser != "0")]
    return sorted(ser.unique().tolist())

def matches_any_canon(series: pd.Series, selected: list[str]) -> pd.Series:
    if not selected:
        return pd.Series([True] * len(series), index=series.index)
    ser = series.astype(str).map(_collapse_spaces)
    return ser.isin(set(selected))

# --- 2. Dicionário para tradução dos meses (mesmo da página principal) ---
meses_traducao = {
    'January': 'Janeiro', 'February': 'Fevereiro', 'March': 'Março',
    'April': 'Abril', 'May': 'Maio', 'June': 'Junho',
    'July': 'Julho', 'August': 'Agosto', 'September': 'Setembro',
    'October': 'Outubro', 'November': 'Novembro', 'December': 'Dezembro'
}
meses_cronologicos = list(meses_traducao.values())

# --- 3. CSS essencial (reaproveita classes usadas na principal) ---
st.markdown("""
<style>
.kpi-card {
  background-color: #333333;
  padding: 20px;
  border-radius: 10px;
  text-align: center;
  margin-bottom: 20px;
}
.kpi-value { font-size: 3em; font-weight: bold; color: #FF4B4B; }
.kpi-label { font-size: 1.2em; color: #FFFFFF; }
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
  font-size: 1.5em; font-weight: bold; color: white;
  border-bottom: 1px solid rgba(255,255,255,0.5);
  padding-bottom: 5px; margin-bottom: 10px;
}
.card-item { margin-bottom: 5px; font-size: 1em; }
.card-label { font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- 4. Carregar e Tratar os Dados (mesmo pipeline da principal) ---
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
CREDS_FILE = "google_credentials.json"
PLANILHA_NOME_1 = "DESLIGAMENTOS"
PLANILHA_NOME_2 = "EQUIPAMENTOS"
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1KeJjbsLVP9DkxPCmNSN4VzbSBeG3SFSCAdPhir39iqg/edit?usp=sharing"

@st.cache_resource(ttl=600)
def connect_to_google_sheets():
    if os.path.exists(CREDS_FILE):
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def fetch_sheet_as_df(worksheet):
    data = worksheet.get_all_values()
    if not data:
        return pd.DataFrame()
    headers = [header.strip() for header in data.pop(0)]
    return pd.DataFrame(data, columns=headers)

@st.cache_data(ttl=600)
def carregar_dados_google_sheets(cache_buster: int = 0):
    try:
        client = connect_to_google_sheets()
        workbook = client.open_by_url(SPREADSHEET_URL)

        df_desligamentos = fetch_sheet_as_df(workbook.worksheet(PLANILHA_NOME_1))
        df_equipamentos = fetch_sheet_as_df(workbook.worksheet(PLANILHA_NOME_2))

        # marca categorias
        df_desligamentos['Categoria'] = 'DESLIGAMENTOS'
        df_equipamentos['Categoria']  = 'EQUIPAMENTOS'

        # concat e renomeação igual à principal
        df_todos = pd.concat([df_desligamentos, df_equipamentos], ignore_index=True)

        mapa_renomear = {
            'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência',
            'ATIVO': 'Ativo', 'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência',
            'QUANTIDADE': 'Quantidade', 'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização',
            'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 'DESCRIÇÃO': 'Descrição',
            'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop',
            'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo', 'CLIENTE AVISADO': 'Cliente Avisado'
        }
        renomear_final = {}
        for col in df_todos.columns:
            cu = col.strip().upper()
            if cu in mapa_renomear:
                renomear_final[col] = mapa_renomear[cu]
        df_todos.rename(columns=renomear_final, inplace=True)
        df_todos.fillna('', inplace=True)

        # coerção de datas
        for c in ['Normalização','Desligamento','Atendimento Loop','Atendimento Terceiros','Cliente Avisado']:
            if c in df_todos.columns:
                df_todos[c] = pd.to_datetime(df_todos[c], errors='coerce')

        # enriquecimento igual à principal
        if 'Desligamento' in df_todos.columns and not df_todos['Desligamento'].isnull().all():
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
        else:
            for c in ['Data','Hora','Mês','Ano','Dia','ID_Unico']:
                df_todos[c] = None

        # garantir tipos de texto
        for c in ['Operador','Descrição','OS','Protocolo']:
            if c in df_todos.columns:
                df_todos[c] = df_todos[c].astype(str).fillna('')

        return df_todos
    except Exception as e:
        st.error(f"Erro ao carregar ou processar os dados do Google Sheets: {e}")
        return pd.DataFrame()

# cache bust
if 'cache_buster' not in st.session_state:
    st.session_state.cache_buster = int(pytime.time())

df = carregar_dados_google_sheets(st.session_state.cache_buster)


df['Desligamento'] = pd.to_datetime(df['Desligamento'], errors='coerce')

# --- 5. KPIs do topo (RESOLVIDAS) ---
count_resolvidos_deslig = df[(df['Categoria']=='DESLIGAMENTOS') & (~df['Normalização'].isna())].shape[0]
count_resolvidos_equip  = df[(df['Categoria']=='EQUIPAMENTOS')  & (~df['Normalização'].isna())].shape[0]

with st.container(border=True):
    st.markdown("<h1 style='margin:0'>OCORRÊNCIAS RESOLVIDAS</h1>", unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown(f"<div class='kpi-card'><div class='kpi-label'>DESLIGAMENTOS</div><div class='kpi-value'>{count_resolvidos_deslig}</div></div>", unsafe_allow_html=True)
    with col_b:
        st.markdown(f"<div class='kpi-card'><div class='kpi-label'>EQUIPAMENTOS</div><div class='kpi-value'>{count_resolvidos_equip}</div></div>", unsafe_allow_html=True)


# --- 6. Filtros (idênticos à principal) ---
st.header('OCORRÊNCIAS FILTRADAS')

# Inicialização de estados (mesmo padrão da principal)
if 'filtros_meses' not in st.session_state:
    st.session_state.filtros_meses = [meses_traducao[datetime.now().strftime('%B')]]
if 'filtros_anos' not in st.session_state:
    if not df.empty and 'Ano' in df.columns:
        anos_atuais = sorted(df['Ano'].unique().tolist())
        st.session_state.filtros_anos = [a for a in anos_atuais if a != 0]
    else:
        st.session_state.filtros_anos = []
if 'filtros_dias' not in st.session_state:
    if not df.empty and {'Mês','Ano'}.issubset(df.columns):
        dias_atuais = sorted(df[(df['Mês'].isin(st.session_state.filtros_meses)) &
                                 (df['Ano'].isin(st.session_state.filtros_anos))]['Dia'].unique().tolist())
        st.session_state.filtros_dias = [d for d in dias_atuais if d != 0]
    else:
        st.session_state.filtros_dias = []
if 'filtros_categorias' not in st.session_state:
    st.session_state.filtros_categorias = sorted(df['Categoria'].unique().tolist()) if not df.empty else []
if 'filtros_clientes' not in st.session_state:
    st.session_state.filtros_clientes = sorted(df['Cliente'].unique().tolist()) if not df.empty else []
if 'filtros_ugs' not in st.session_state:
    st.session_state.filtros_ugs = sorted(df['UG'].unique().tolist()) if not df.empty else []
if 'filtros_tipos' not in st.session_state:
    st.session_state.filtros_tipos = sorted(df['Tipo de ocorrência'].unique().tolist()) if not df.empty else []
if 'filtros_ativos' not in st.session_state:
    st.session_state.filtros_ativos = sorted(df['Ativo'].unique().tolist()) if not df.empty else []
if 'filtros_ocorrencias' not in st.session_state:
    st.session_state.filtros_ocorrencias = sorted(df['Ocorrência'].unique().tolist()) if not df.empty else []

# KPI esquerdo: total resolvidas no banco completo
col_kpi1, col_kpi2 = st.columns(2)
with col_kpi1:
    total_resolvidas_banco = df[~df['Normalização'].isna()].shape[0]
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total Resolvidas no Banco Completo</div>
        <div class="kpi-value">{total_resolvidas_banco}</div>
    </div>
    """, unsafe_allow_html=True)

# Botão de atualizar (mesmo padrão)
col_left, _ = st.columns([0.2, 0.8])
with col_left:
    if st.button("Atualizar Dados"):
        st.cache_data.clear()
        start_loading()
        st.session_state.cache_buster = int(pytime.time())
        st.rerun()


# Helper para marcar/desmarcar checkboxes em massa
def _marcar(prefixo_key: str, itens: list, filtro_key: str, marcar_todos: bool, validos: set | None = None):
    validos = set(itens) if validos is None else set(validos)
    st.session_state[filtro_key] = [x for x in itens if (x in validos) and marcar_todos]
    for x in itens:
        st.session_state[f"{prefixo_key}{x}"] = marcar_todos and (x in validos)

if not df.empty:
    if 'categoria_top' not in st.session_state:
        st.session_state['categoria_top'] = 'Ambas'

    st.markdown("#### Filtrar por categoria (planilha)")
    st.radio(
        "Categoria:",
        options=["Ambas","DESLIGAMENTOS","EQUIPAMENTOS"],
        horizontal=True,
        label_visibility="collapsed",
        key="categoria_top",
        on_change=start_loading,   # <- novo
    )


    st.subheader("Selecione o período desejado")
    anos_disponiveis = sorted([a for a in df['Ano'].unique() if a != 0])
    meses_disponiveis = meses_cronologicos[:]

    col_ano, col_mes, col_dia = st.columns(3)

    # Ano(s)
    with col_ano:
        with st.container(border=True):
            st.write("### Ano(s):")
            with st.expander("Expandir anos"):
                for ano in anos_disponiveis:
                    st.checkbox(str(ano), key=f'cb_ano_{ano}', value=(ano in st.session_state.filtros_anos))
            c = st.columns(2)
            clicked_sel_ano = c[0].button(
                'Sel. Todos', key='sel_ano', use_container_width=True,
                on_click=marcar_loading, args=('cb_ano_', anos_disponiveis, 'filtros_anos', True)
            )
            clicked_des_ano = c[1].button(
                'Desmarcar', key='des_ano', use_container_width=True,
                on_click=marcar_loading, args=('cb_ano_', anos_disponiveis, 'filtros_anos', False)
            )
            if not (clicked_sel_ano or clicked_des_ano):
                st.session_state.filtros_anos = [a for a in anos_disponiveis if st.session_state.get(f'cb_ano_{a}', False)]

    # Mês(es)
    with col_mes:
        with st.container(border=True):
            st.write("### Mês(es):")
            with st.expander("Expandir meses"):
                for mes in meses_disponiveis:
                    st.checkbox(mes, key=f'cb_mes_{mes}', value=(mes in st.session_state.filtros_meses))
            c = st.columns(2)
            clicked_sel_mes = c[0].button(
                'Sel. Todos', key='sel_mes', use_container_width=True,
                on_click=marcar_loading, args=('cb_mes_', meses_disponiveis, 'filtros_meses', True)
            )
            clicked_des_mes = c[1].button(
                'Desmarcar', key='des_mes', use_container_width=True,
                on_click=marcar_loading, args=('cb_mes_', meses_disponiveis, 'filtros_meses', False)
            )
            if not (clicked_sel_mes or clicked_des_mes):
                st.session_state.filtros_meses = [m for m in meses_disponiveis if st.session_state.get(f'cb_mes_{m}', False)]

    # Seleções atuais de Ano/Mês (usadas para dias dinâmicos)
    anos_sel  = [a for a in anos_disponiveis  if st.session_state.get(f'cb_ano_{a}', False)]
    meses_sel = [m for m in meses_disponiveis if st.session_state.get(f'cb_mes_{m}', False)]

    # Dias disponíveis dinâmicos, coerentes com Ano/Mês selecionados
    dias_disponiveis = sorted(
        df[(df['Ano'].isin(anos_sel)) & (df['Mês'].isin(meses_sel))]['Dia']
        .dropna().astype(int).loc[lambda s: s.gt(0)].unique().tolist()
    ) or list(range(1, 32))

    # Dia(s)
    with col_dia:
        with st.container(border=True):
            st.write("### Dia(s):")
            with st.expander("Expandir dias"):
                dias_cols = st.columns(7)
                for i, dia in enumerate(range(1, 32)):
                    with dias_cols[i % 7]:
                        if dia in dias_disponiveis:
                            st.checkbox(str(dia), key=f'cb_dia_{dia}', value=(dia in st.session_state.filtros_dias))
                        else:
                            st.checkbox(str(dia), key=f'cb_dia_{dia}', disabled=True)
            c = st.columns(2)
            clicked_sel_dia = c[0].button(
                'Sel. Todos', key='sel_dia', use_container_width=True,
                on_click=marcar_loading, args=('cb_dia_', list(range(1, 32)), 'filtros_dias', True, set(dias_disponiveis))
            )
            clicked_des_dia = c[1].button(
                'Desmarcar', key='des_dia', use_container_width=True,
                on_click=marcar_loading, args=('cb_dia_', list(range(1, 32)), 'filtros_dias', False, set(dias_disponiveis))
            )
            if not (clicked_sel_dia or clicked_des_dia):
                st.session_state.filtros_dias = [d for d in dias_disponiveis if st.session_state.get(f'cb_dia_{d}', False)]

    # ------ opções + defaults canônicos (substitui "Filtros Adicionais") ------
    df_ref = df[df['Normalização'].notna()].copy()
    # Cliente
    cli_opts = options_from(df_ref['Cliente'])
    st.session_state.filtros_clientes = [c for c in st.session_state.get('filtros_clientes', []) if c in cli_opts]
    st.session_state.filtros_clientes = st.multiselect('Cliente', options=cli_opts, default=st.session_state.filtros_clientes)

    # UG normalizada
    ugs_series = df_ref['UG'].astype(str).map(_collapse_spaces)
    ug_opts = sorted([u for u in ugs_series.unique().tolist() if u and u != "-" and u != "0"])
    st.session_state.filtros_ugs = [u for u in st.session_state.get('filtros_ugs', []) if u in ug_opts]
    st.session_state.filtros_ugs = st.multiselect('UG', options=ug_opts, default=st.session_state.filtros_ugs)

    # Tipo / Ocorrência / Ativo
    tip_opts = options_from(df_ref['Tipo de ocorrência']) if 'Tipo de ocorrência' in df_ref.columns else []
    ocr_opts = options_from(df_ref['Ocorrência']) if 'Ocorrência' in df_ref.columns else []
    atv_opts = options_from(df_ref['Ativo']) if 'Ativo' in df_ref.columns else []
    for key, opts in [('filtros_tipos', tip_opts), ('filtros_ocorrencias', ocr_opts), ('filtros_ativos', atv_opts)]:
        st.session_state[key] = [x for x in st.session_state.get(key, []) if x in opts]

    # --- Aplicação dos filtros + RESOLVIDAS ---
    meses_sel = [m for m in meses_cronologicos if st.session_state.get(f'cb_mes_{m}', False)]
    anos_sel  = [a for a in anos_disponiveis      if st.session_state.get(f'cb_ano_{a}', False)]
    dias_sel = [d for d in dias_disponiveis if st.session_state.get(f'cb_dia_{d}', False)]

    set_anos  = set(anos_disponiveis)
    set_meses = set(meses_cronologicos)
    set_dias  = set(dias_disponiveis)

    all_anos  = set(anos_sel)  == set_anos  and len(set_anos)  > 0
    all_meses = set(meses_sel) == set_meses and len(set_meses) > 0
    all_dias  = set(dias_sel)  == set_dias  and len(set_dias)  > 0

    s_ano, s_mes, s_dia = df['Ano'], df['Mês'], df['Dia']
    m_ano = s_ano.isin(anos_sel)  if not all_anos  else (s_ano.isin(anos_sel)  | s_ano.isna() | (s_ano == 0))
    m_mes = s_mes.isin(meses_sel) if not all_meses else (s_mes.isin(meses_sel) | s_mes.isna() | (s_mes.astype(str) == ''))
    m_dia = s_dia.isin(dias_sel)  if not all_dias  else (s_dia.isin(dias_sel)  | s_dia.isna() | (s_dia == 0))

    m_cat = df['Categoria'].isin(st.session_state.filtros_categorias)
    if st.session_state.get("categoria_top") in ("DESLIGAMENTOS", "EQUIPAMENTOS"):
        m_cat = m_cat & (df['Categoria'] == st.session_state["categoria_top"])

    m_cli = matches_any_canon(df['Cliente'], st.session_state.get('filtros_clientes', []))
    m_tip = matches_any_canon(df['Tipo de ocorrência'], st.session_state.get('filtros_tipos', [])) if 'Tipo de ocorrência' in df.columns else True
    m_ocr = matches_any_canon(df['Ocorrência'], st.session_state.get('filtros_ocorrencias', [])) if 'Ocorrência' in df.columns else True
    m_atv = matches_any_canon(df['Ativo'], st.session_state.get('filtros_ativos', [])) if 'Ativo' in df.columns else True
    m_ug  = df['UG'].astype(str).map(_collapse_spaces).isin(set(st.session_state.get('filtros_ugs', [])))

    m_final = m_cli & m_tip & m_ocr & m_atv & m_ug & m_ano & m_mes & m_dia & m_cat

    df_filtrado   = df[m_final].copy()
    df_resolvidas = df_filtrado[~df_filtrado['Normalização'].isna()].copy()


    if st.session_state.ui_phase == 'loading':
        st.session_state.ui_phase = 'ready'
        st.session_state.loading_ts = 0
        render_loading_overlay('ready')  # opcional, para garantir CSS em estado 'ready'
        st.rerun()


    # KPI direito: total resolvidas com filtro
    with col_kpi2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Resolvidas com Filtro</div>
            <div class="kpi-value">{len(df_resolvidas)}</div>
        </div>
        """, unsafe_allow_html=True)

    # Lista e cards
    if not df_resolvidas.empty:
        # Ordenação
        st.markdown("---")
        st.write("### Ordenar e Exibir")
        sort_cols = st.columns(2)
        with sort_cols[0]:
            sort_options_display = {
                'Data do Desligamento': 'Desligamento',
                'Data da Normalização': 'Normalização',
                'UG': 'UG',
                'Ativo': 'Ativo'
            }
            sort_by_display = st.selectbox("Ordenar por:", options=sort_options_display.keys(), index=1)
            sort_by_column = sort_options_display[sort_by_display]
        with sort_cols[1]:
            sort_order = st.radio("Ordem:", options=['Descendente', 'Ascendente'], index=0, horizontal=True)
            is_ascending = (sort_order == 'Ascendente')

        df_sorted = df_resolvidas.sort_values(by=sort_by_column, ascending=is_ascending, na_position='last')

        # Tabela resumida
        st.header("Lista de Ocorrências (Tabela)")
        df_tab = df_sorted.copy()
        df_tab.reset_index(inplace=True, drop=True)
        df_tab['Linha'] = df_tab.index + 1
        st.dataframe(df_tab[[
            'Linha','Categoria','UG','Data','Hora','Tipo de ocorrência',
            'Ativo','Ocorrência','Operador','Descrição','OS'
        ]], use_container_width=True)

        # Cards de detalhes
        st.header("Detalhes por Ocorrência (Cards)")

        num_cols = 4  # igual à página principal
        rows = list(df_sorted.iterrows())

        def fmt_dt(dt):
            if pd.notna(dt):
                return dt.strftime('%d/%m/%Y'), dt.strftime('%H:%M')
            return '', ''

        for i in range(0, len(rows), num_cols):
            try:
                cols = st.columns(num_cols, gap="small")  # grid por linha
            except TypeError:
                cols = st.columns(num_cols)               # fallback

            for j in range(num_cols):
                if i + j >= len(rows):
                    break
                _, r = rows[i + j]
                with cols[j]:
                    cliente   = html.escape(str(r.get("Cliente", "")))
                    categoria = html.escape(str(r.get("Categoria", "")))
                    ug        = html.escape(str(r.get("UG", "N/A")))
                    tipo      = html.escape(str(r.get("Tipo de ocorrência", "")))
                    ativo     = html.escape(str(r.get("Ativo", "")))
                    nome_ativo= html.escape(str(r.get("Nome Ativo", "")))
                    ocorr     = html.escape(str(r.get("Ocorrência", "")))
                    operador  = html.escape(str(r.get("Operador", "")))
                    descricao = html.escape(str(r.get("Descrição", ""))).replace('\n','<br>')
                    protocolo = html.escape(str(r.get("Protocolo", "")))
                    osv       = html.escape(str(r.get("OS", "")))

                    d_des, h_des = fmt_dt(r.get('Desligamento'))
                    d_norm, h_norm = fmt_dt(r.get('Normalização'))
                    d_ca, h_ca = fmt_dt(r.get('Cliente Avisado'))
                    d_loop, h_loop = fmt_dt(r.get('Atendimento Loop'))
                    d_terc, h_terc = fmt_dt(r.get('Atendimento Terceiros'))

                    qtd_html = ''
                    if r.get('Categoria') == 'EQUIPAMENTOS':
                        qv = r.get('Quantidade', 0)
                        try:
                            if pd.notna(qv) and float(qv) > 0:
                                qtd_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {int(float(qv))}</div>'
                        except (ValueError, TypeError):
                            qtd_html = ''

                    card_html = f"""
                    <div class="card-container">
                      <div class="card-title">{ug}</div>
                      <div class="card-item"><span class="card-label">Cliente:</span> {cliente}</div>
                      <div class="card-item"><span class="card-label">Categoria:</span> {categoria}</div>
                      <div class="card-item"><span class="card-label">Tipo de Ocorrência:</span> {tipo}</div>
                      <div class="card-item"><span class="card-label">Ativo:</span> {ativo}</div>
                      <div class="card-item"><span class="card-label">Nome do ativo:</span> {nome_ativo}</div>
                      <div class="card-item"><span class="card-label">Ocorrência:</span> {ocorr}</div>
                      <div class="card-item"><span class="card-label">Operador:</span> {operador}</div>
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
                      <div class="card-item"><span class="card-label">Descrição:</span> {descricao}</div>
                      <div class="card-item"><span class="card-label">Protocolo:</span> {protocolo}</div>
                      <div class="card-item"><span class="card-label">OS:</span> {osv}</div>
                    </div>
                    """
                    st.html(card_html)
    else:
        st.info("Nenhuma ocorrência resolvida encontrada para os filtros selecionados.")
else:
    st.warning("Não foi possível carregar os dados. Verifique as credenciais ou a conexão.")