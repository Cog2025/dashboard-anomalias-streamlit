# 1_Pagina_Principal.py
# === Imports ===
import os
from webdav3.client import Client  # opcional se não usa aqui
import io                          # opcional se não usa aqui
import streamlit as st
import pandas as pd
from datetime import datetime
from time import time
import html
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe

# === Identificação da página e helper de nomes de estado ===
PAGE_ID = "p1"  # na 4_Ocorrencias_Resolvidas.py; use "p4" lá
def K(name: str) -> str:
    return f"{PAGE_ID}:{name}"

# --- 1. Configuração da Página e Overlay ---
st.set_page_config(layout="wide")

# === Overlay de carregamento (logo após st.set_page_config) ===
from time import time

# 1) Estado inicial dos controladores de overlay
if 'ui_phase' not in st.session_state:           # 'init' -> 1a render só para desenhar overlay
    st.session_state.ui_phase = 'init'
if 'loading_ts' not in st.session_state:         # timestamp do overlay (failsafe)
    st.session_state.loading_ts = 0

def render_loading_overlay():
    # Mostra overlay somente na fase 'loading'
    display = 'flex' if st.session_state.get('ui_phase') == 'loading' else 'none'
    st.markdown(f"""
    <style>
      #__overlay__ {{
        position: fixed; inset: 0; background: rgba(0,0,0,.55);
        display: {display}; align-items: center; justify-content: center;
        z-index: 10000;
      }}
      .loader {{
        width: 64px; height: 64px; border-radius: 50%;
        border: 6px solid rgba(255,255,255,.25);
        border-top-color: #FF4B4B;
        animation: spin 1s linear infinite;
      }}
      @keyframes spin {{ to {{ transform: rotate(360deg); }} }}
    </style>
    <div id="__overlay__"><div class="loader"></div></div>
    """, unsafe_allow_html=True)

def start_loading():
    st.session_state.ui_phase = 'loading'
    st.session_state.loading_ts = time()
    st.experimental_rerun()

def stop_loading():
    st.session_state.ui_phase = 'ready'
    st.session_state.loading_ts = 0
    st.experimental_rerun()

# 2) Desenha overlay (com base na fase)
render_loading_overlay()

# 3) Gate de fase para garantir que o overlay apareça ANTES do trabalho pesado
if st.session_state.ui_phase == 'init':
    # Primeira passada: só liga overlay e rerender
    start_loading()

# 4) Failsafe: se passar de 20s carregando, força fechar overlay para não travar a UI
if st.session_state.ui_phase == 'loading' and (time() - st.session_state.loading_ts) > 20:
    stop_loading()


# --- 2. Dicionário de meses (en-US -> pt-BR) ---
meses_traducao = {
    'January': 'Janeiro', 'February': 'Fevereiro', 'March': 'Março',
    'April': 'Abril', 'May': 'Maio', 'June': 'Junho',
    'July': 'Julho', 'August': 'Agosto', 'September': 'Setembro',
    'October': 'Outubro', 'November': 'Novembro', 'December': 'Dezembro'
}
meses_cronologicos = list(meses_traducao.values())

# --- 3. CSS utilitário (cards e KPIs) ---
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

# --- 4. Google Sheets: credenciais e leitura ---
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

        # Lê as duas planilhas
        df_desligamentos = fetch_sheet_as_df(workbook.worksheet(PLANILHA_NOME_1))
        df_equipamentos  = fetch_sheet_as_df(workbook.worksheet(PLANILHA_NOME_2))

        # Marca categorias
        df_desligamentos['Categoria'] = 'DESLIGAMENTOS'
        df_equipamentos['Categoria']  = 'EQUIPAMENTOS'

        # Concat
        df_todos = pd.concat([df_desligamentos, df_equipamentos], ignore_index=True)

        # Renomeia (mapa robusto por UPPER)
        mapa_renomear = {
            'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG',
            'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência', 'ATIVO': 'Ativo',
            'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência',
            'QUANTIDADE': 'Quantidade', 'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização',
            'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 'DESCRIÇÃO': 'Descrição',
            'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop',
            'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo',
            'CLIENTE AVISADO': 'Cliente Avisado'
        }
        renomear_final = {}
        for col in df_todos.columns:
            cu = col.strip().upper()
            if cu in mapa_renomear:
                renomear_final[col] = mapa_renomear[cu]
        df_todos.rename(columns=renomear_final, inplace=True)
        df_todos.fillna('', inplace=True)

        # Datetimes
        for c in ['Normalização','Desligamento','Atendimento Loop','Atendimento Terceiros','Cliente Avisado']:
            if c in df_todos.columns:
                df_todos[c] = pd.to_datetime(df_todos[c], errors='coerce')

        # Enriquecimento a partir de Desligamento
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

        # Textos obrigatórios como str
        for c in ['Operador','Descrição','OS','Protocolo']:
            if c in df_todos.columns:
                df_todos[c] = df_todos[c].astype(str).fillna('')

        return df_todos
    except Exception as e:
        st.error(f"Erro ao carregar ou processar os dados do Google Sheets: {e}")
        return pd.DataFrame()

# --- 5. Cache-buster + Carregamento ---
if 'cache_buster' not in st.session_state:
    st.session_state.cache_buster = int(time())

df_todos_dados = carregar_dados_google_sheets(st.session_state.cache_buster)

# Guarda de erro: sem dados
if df_todos_dados is None or df_todos_dados.empty:
    st.warning("Não foi possível carregar os dados. Verifique as credenciais ou a conexão.")
    if st.session_state.get('loading', False):
        st.session_state.loading = False
        st.session_state.overlay_rendered = False
    st.stop()

# Fecha overlay e rerender quando dados prontos
if st.session_state.get('loading', False):
    st.session_state.loading = False
    st.session_state.overlay_rendered = False
    st.rerun()

# --- 6. Helpers (sanitização, sincronização, marcação) ---
def _sanitize_default(default_list, valid_options):
    if not isinstance(default_list, (list, tuple)):
        return []
    seen, out, valid = set(), [], set(valid_options)
    for x in default_list:
        if x in valid and x not in seen:
            seen.add(x)
            out.append(x)
    return out

# Sincroniza estado do widget e do modelo (multiselects)
def _sync_ms(key_ms: str, key_modelo: str, valores: list):
    st.session_state[key_modelo] = list(valores)
    st.session_state[key_ms] = list(valores)

# Marca/desmarca checkboxes em massa (Ano/Mês/Dia)
def _marcar(prefixo_key: str, itens: list, filtro_key: str, marcar_todos: bool, validos: set | None = None):
    validos = set(itens) if validos is None else set(validos)
    st.session_state[filtro_key] = [x for x in itens if (x in validos) and marcar_todos]
    for x in itens:
        st.session_state[f"{prefixo_key}{x}"] = marcar_todos and (x in validos)

# Deduplicação case-insensitive e remoção de ''/'0'
def dedup_case_insensitive(seq):
    seen, out = set(), []
    for x in seq:
        s = str(x).strip()
        k = s.upper()
        if k and k != '0' and k not in seen:
            seen.add(k)
            out.append(s)
    return sorted(out)

# Normaliza colunas para comparação case-insensitive
for col, key in [('Cliente','CLI'),('UG','UG'),('Tipo de ocorrência','TIPO'),('Ativo','ATV'),('Ocorrência','OCR')]:
    if col in df_todos_dados.columns:
        df_todos_dados[f'__{key}_NORM__'] = df_todos_dados[col].astype(str).str.strip().str.upper()

# --- 7. Cabeçalho, KPI e botão atualizar ---
with st.container(border=True):
    st.markdown("<h1 style='margin:0'>PAINEL GERAL</h1>", unsafe_allow_html=True)

col_kpi1, col_kpi2 = st.columns(2)
with col_kpi1:
    total_banco = len(df_todos_dados)
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total de Registros (Banco Completo)</div>
        <div class="kpi-value">{total_banco}</div>
    </div>
    """, unsafe_allow_html=True)

col_left, _ = st.columns([0.2, 0.8])
with col_left:
    if st.button('Atualizar Dados'):
        st.cache_data.clear()
        st.session_state.ui_phase = 'init'   # volta para fase inicial
        st.session_state.loading_ts = 0
        st.experimental_rerun()


# --- 8. Categoria (uma única UI; remova duplicações desta página) ---
CAT_KEY = 'categoria_top_global'
if CAT_KEY not in st.session_state:
    st.session_state[CAT_KEY] = 'Ambas'

st.markdown("#### Filtrar por categoria (planilha)")
st.radio(
    "Categoria:",
    options=["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"],
    horizontal=True,
    label_visibility="collapsed",
    key=CAT_KEY
)
cat_sel = st.session_state[CAT_KEY]
m_cat = (df_todos_dados['Categoria'] == cat_sel) if cat_sel in ("DESLIGAMENTOS", "EQUIPAMENTOS") else df_todos_dados['Categoria'].notna()

# --- 9. Seletor de período (Ano/Mês/Dia) com máscaras que incluem NaN/0 quando “tudo” ---
st.subheader("Selecione o período desejado")
anos_disponiveis  = sorted([a for a in df_todos_dados['Ano'].unique() if a != 0]) if 'Ano' in df_todos_dados.columns else []
meses_disponiveis = meses_cronologicos[:]  # 12 meses PT-BR
col_ano, col_mes, col_dia = st.columns(3)

# Estado inicial
if K('filtros_meses') not in st.session_state:
    st.session_state[K('filtros_meses')] = [meses_traducao[datetime.now().strftime('%B')]]
if K('filtros_anos') not in st.session_state:
    st.session_state[K('filtros_anos')] = [a for a in anos_disponiveis if a != 0]
if K('filtros_dias') not in st.session_state:
    st.session_state[K('filtros_dias')] = list(range(1, 32))

# Ano(s)
with col_ano:
    with st.container(border=True):
        st.write("### Ano(s):")
        with st.expander("Expandir anos"):
            for ano in anos_disponiveis:
                st.checkbox(
                    str(ano),
                    key=f'cb_ano_{PAGE_ID}_{ano}',
                    value=(ano in st.session_state.get(K('filtros_anos'), []))
                )
        c = st.columns(2)
        clicked_sel_ano = c[0].button('Sel. Todos', key=f'sel_ano_{PAGE_ID}', use_container_width=True,
                                      on_click=_marcar, args=(f'cb_ano_{PAGE_ID}_', anos_disponiveis, K('filtros_anos'), True))
        clicked_des_ano = c[1].button('Desmarcar', key=f'des_ano_{PAGE_ID}', use_container_width=True,
                                      on_click=_marcar, args=(f'cb_ano_{PAGE_ID}_', anos_disponiveis, K('filtros_anos'), False))
        if not (clicked_sel_ano or clicked_des_ano):
            st.session_state[K('filtros_anos')] = [
                a for a in anos_disponiveis if st.session_state.get(f'cb_ano_{PAGE_ID}_{a}', False)
            ]

# Mês(es)
with col_mes:
    with st.container(border=True):
        st.write("### Mês(es):")
        with st.expander("Expandir meses"):
            for mes in meses_disponiveis:
                st.checkbox(
                    mes,
                    key=f'cb_mes_{PAGE_ID}_{mes}',
                    value=(mes in st.session_state.get(K('filtros_meses'), []))
                )
        c = st.columns(2)
        clicked_sel_mes = c[0].button('Sel. Todos', key=f'sel_mes_{PAGE_ID}', use_container_width=True,
                                      on_click=_marcar, args=(f'cb_mes_{PAGE_ID}_', meses_disponiveis, K('filtros_meses'), True))
        clicked_des_mes = c[1].button('Desmarcar', key=f'des_mes_{PAGE_ID}', use_container_width=True,
                                      on_click=_marcar, args=(f'cb_mes_{PAGE_ID}_', meses_disponiveis, K('filtros_meses'), False))
        if not (clicked_sel_mes or clicked_des_mes):
            st.session_state[K('filtros_meses')] = [
                m for m in meses_disponiveis if st.session_state.get(f'cb_mes_{PAGE_ID}_{m}', False)
            ]

# Dia(s)
with col_dia:
    with st.container(border=True):
        st.write("### Dia(s):")
        dias_disponiveis = list(range(1, 32))
        with st.expander("Expandir dias"):
            dias_cols = st.columns(7)
            for i, dia in enumerate(range(1, 32)):
                with dias_cols[i % 7]:
                    st.checkbox(
                        str(dia),
                        key=f'cb_dia_{PAGE_ID}_{dia}',
                        value=(dia in st.session_state.get(K('filtros_dias'), []))
                    )
        c = st.columns(2)
        clicked_sel_dia = c[0].button('Sel. Todos', key=f'sel_dia_{PAGE_ID}', use_container_width=True,
                                      on_click=_marcar, args=(f'cb_dia_{PAGE_ID}_', list(range(1, 32)), K('filtros_dias'), True, set(dias_disponiveis)))
        clicked_des_dia = c[1].button('Desmarcar', key=f'des_dia_{PAGE_ID}', use_container_width=True,
                                      on_click=_marcar, args=(f'cb_dia_{PAGE_ID}_', list(range(1, 32)), K('filtros_dias'), False, set(dias_disponiveis)))
        if not (clicked_sel_dia or clicked_des_dia):
            st.session_state[K('filtros_dias')] = [
                d for d in dias_disponiveis if st.session_state.get(f'cb_dia_{PAGE_ID}_{d}', False)
            ]

# Conjuntos e flags de "tudo selecionado" + máscaras de período com inclusão de ausentes/0
anos_sel  = st.session_state.get(K('filtros_anos'), [])
meses_sel = st.session_state.get(K('filtros_meses'), [])
dias_sel  = st.session_state.get(K('filtros_dias'), [])

set_anos  = set(anos_disponiveis)
set_meses = set(meses_cronologicos)
set_dias  = set(dias_disponiveis)

all_anos  = set(anos_sel)  == set_anos  and len(set_anos)  > 0
all_meses = set(meses_sel) == set_meses and len(set_meses) > 0
all_dias  = set(dias_sel)  == set_dias  and len(set_dias)  > 0

s_ano, s_mes, s_dia = df_todos_dados['Ano'], df_todos_dados['Mês'], df_todos_dados['Dia']
m_ano = s_ano.isin(anos_sel)  if not all_anos  else (s_ano.isin(anos_sel)  | s_ano.isna() | (s_ano == 0))
m_mes = s_mes.isin(meses_sel) if not all_meses else (s_mes.isin(meses_sel) | s_mes.isna() | (s_mes.astype(str) == ''))
m_dia = s_dia.isin(dias_sel)  if not all_dias  else (s_dia.isin(dias_sel)  | s_dia.isna() | (s_dia == 0))

# --- 10. Filtros Adicionais (multiselects sincronizados e case-insensitive) ---
st.subheader("Filtros Adicionais")
col_cliente, col_ug, col_tipo, col_ativo, col_ocorrencia = st.columns(5)

def _sel_norm(key_modelo):
    return [str(v).strip().upper() for v in st.session_state.get(key_modelo, []) if str(v).strip() and str(v).strip() != '0']

# Cliente
with col_cliente:
    with st.container(border=True):
        st.write("Cliente:")
        options_clientes = dedup_case_insensitive(df_todos_dados['Cliente'].tolist())
        if K('filtros_clientes') not in st.session_state:
            st.session_state[K('filtros_clientes')] = options_clientes[:]
        key_ms_cli = f'ms_clientes_{PAGE_ID}'
        if key_ms_cli not in st.session_state:
            st.session_state[key_ms_cli] = st.session_state[K('filtros_clientes')]

        c = st.columns(2)
        if c[0].button('Sel. Todos', key=f'sel_cli_{PAGE_ID}', use_container_width=True):
            _sync_ms(key_ms_cli, K('filtros_clientes'), options_clientes)
            st.rerun()
        if c[1].button('Desmarcar', key=f'des_cli_{PAGE_ID}', use_container_width=True):
            _sync_ms(key_ms_cli, K('filtros_clientes'), [])
            st.rerun()

        st.session_state[K('filtros_clientes')] = st.multiselect(
            ' ', options=options_clientes,
            default=st.session_state.get(key_ms_cli, options_clientes),
            label_visibility='hidden',
            key=key_ms_cli
        )

# UG (dependente de Cliente)
with col_ug:
    with st.container(border=True):
        st.write("UG:")
        sel_cli_norm = _sel_norm(K('filtros_clientes'))
        df_temp = df_todos_dados[df_todos_dados['__CLI_NORM__'].isin(sel_cli_norm)] if sel_cli_norm else df_todos_dados
        ugs_disponiveis = dedup_case_insensitive(df_temp['UG'].tolist())

        if K('filtros_ugs') not in st.session_state:
            st.session_state[K('filtros_ugs')] = ugs_disponiveis[:]
        key_ms_ug = f'ms_ugs_{PAGE_ID}'
        if key_ms_ug not in st.session_state:
            st.session_state[key_ms_ug] = st.session_state[K('filtros_ugs')]

        c = st.columns(2)
        if c[0].button('Sel. Todos', key=f'sel_ug_{PAGE_ID}', use_container_width=True):
            _sync_ms(key_ms_ug, K('filtros_ugs'), ugs_disponiveis)
            st.rerun()
        if c[1].button('Desmarcar', key=f'des_ug_{PAGE_ID}', use_container_width=True):
            _sync_ms(key_ms_ug, K('filtros_ugs'), [])
            st.rerun()

        st.session_state[K('filtros_ugs')] = st.multiselect(
            ' ', options=ugs_disponiveis,
            default=st.session_state.get(key_ms_ug, ugs_disponiveis),
            label_visibility='hidden',
            key=key_ms_ug
        )

# Tipo de Ocorrência
with col_tipo:
    with st.container(border=True):
        st.write("Tipo de Ocorrência:")
        options_tipos = dedup_case_insensitive(df_todos_dados['Tipo de ocorrência'].tolist())

        if K('filtros_tipos') not in st.session_state:
            st.session_state[K('filtros_tipos')] = options_tipos[:]
        key_ms_tipos = f'ms_tipos_{PAGE_ID}'
        if key_ms_tipos not in st.session_state:
            st.session_state[key_ms_tipos] = st.session_state[K('filtros_tipos')]

        c = st.columns(2)
        if c[0].button('Sel. Todos', key=f'sel_tipo_{PAGE_ID}', use_container_width=True):
            _sync_ms(key_ms_tipos, K('filtros_tipos'), options_tipos)
            st.rerun()
        if c[1].button('Desmarcar', key=f'des_tipo_{PAGE_ID}', use_container_width=True):
            _sync_ms(key_ms_tipos, K('filtros_tipos'), [])
            st.rerun()

        st.session_state[K('filtros_tipos')] = st.multiselect(
            ' ', options=options_tipos,
            default=st.session_state.get(key_ms_tipos, options_tipos),
            label_visibility='hidden',
            key=key_ms_tipos
        )

# Ativo
with col_ativo:
    with st.container(border=True):
        st.write("Ativo:")
        options_ativos = dedup_case_insensitive(df_todos_dados['Ativo'].tolist())

        if K('filtros_ativos') not in st.session_state:
            st.session_state[K('filtros_ativos')] = options_ativos[:]
        key_ms_ativos = f'ms_ativos_{PAGE_ID}'
        if key_ms_ativos not in st.session_state:
            st.session_state[key_ms_ativos] = st.session_state[K('filtros_ativos')]

        c = st.columns(2)
        if c[0].button('Sel. Todos', key=f'sel_ativo_{PAGE_ID}', use_container_width=True):
            _sync_ms(key_ms_ativos, K('filtros_ativos'), options_ativos)
            st.rerun()
        if c[1].button('Desmarcar', key=f'des_ativo_{PAGE_ID}', use_container_width=True):
            _sync_ms(key_ms_ativos, K('filtros_ativos'), [])
            st.rerun()

        st.session_state[K('filtros_ativos')] = st.multiselect(
            ' ', options=options_ativos,
            default=st.session_state.get(key_ms_ativos, options_ativos),
            label_visibility='hidden',
            key=key_ms_ativos
        )

# Ocorrência
with col_ocorrencia:
    with st.container(border=True):
        st.write("Ocorrência:")
        options_ocorr = dedup_case_insensitive(df_todos_dados['Ocorrência'].tolist())

        if K('filtros_ocorrencias') not in st.session_state:
            st.session_state[K('filtros_ocorrencias')] = options_ocorr[:]
        key_ms_ocr = f'ms_ocorrencias_{PAGE_ID}'
        if key_ms_ocr not in st.session_state:
            st.session_state[key_ms_ocr] = st.session_state[K('filtros_ocorrencias')]

        c = st.columns(2)
        if c[0].button('Sel. Todos', key=f'sel_ocr_{PAGE_ID}', use_container_width=True):
            _sync_ms(key_ms_ocr, K('filtros_ocorrencias'), options_ocorr)
            st.rerun()
        if c[1].button('Desmarcar', key=f'des_ocr_{PAGE_ID}', use_container_width=True):
            _sync_ms(key_ms_ocr, K('filtros_ocorrencias'), [])
            st.rerun()

        st.session_state[K('filtros_ocorrencias')] = st.multiselect(
            ' ', options=options_ocorr,
            default=st.session_state.get(key_ms_ocr, options_ocorr),
            label_visibility='hidden',
            key=key_ms_ocr
        )

# Máscaras case-insensitive finais (para uso a seguir)
sel_cli_norm = _sel_norm(K('filtros_clientes'))
sel_ug_norm  = _sel_norm(K('filtros_ugs'))
sel_tip_norm = _sel_norm(K('filtros_tipos'))
sel_atv_norm = _sel_norm(K('filtros_ativos'))
sel_ocr_norm = _sel_norm(K('filtros_ocorrencias'))

m_cli = df_todos_dados['__CLI_NORM__'].isin(sel_cli_norm) if sel_cli_norm else df_todos_dados['__CLI_NORM__'].notna()
m_ug  = df_todos_dados['__UG_NORM__'].isin(sel_ug_norm)   if sel_ug_norm  else df_todos_dados['__UG_NORM__'].notna()
m_tip = df_todos_dados['__TIPO_NORM__'].isin(sel_tip_norm)if sel_tip_norm else df_todos_dados['__TIPO_NORM__'].notna()
m_atv = df_todos_dados['__ATV_NORM__'].isin(sel_atv_norm) if sel_atv_norm else df_todos_dados['__ATV_NORM__'].notna()
m_ocr = df_todos_dados['__OCR_NORM__'].isin(sel_ocr_norm) if sel_ocr_norm else df_todos_dados['__OCR_NORM__'].notna()

# --- 11. Aplicação dos filtros combinados ---
df_filtrado = df_todos_dados[m_ano & m_mes & m_dia & m_cat & m_cli & m_ug & m_tip & m_atv & m_ocr].copy()

# KPIs: banco completo vs. com filtro
with col_kpi2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total (com Filtro)</div>
        <div class="kpi-value">{len(df_filtrado)}</div>
    </div>
    """, unsafe_allow_html=True)

# --- 12. Ordenação e exibição (tabela resumida) ---
if not df_filtrado.empty:
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
        sort_by_display = st.selectbox("Ordenar por:", options=sort_options_display.keys(), index=0)
        sort_by_column = sort_options_display[sort_by_display]
    with sort_cols[1]:
        sort_order = st.radio("Ordem:", options=['Descendente', 'Ascendente'], index=0, horizontal=True)
        is_ascending = (sort_order == 'Ascendente')

    df_sorted = df_filtrado.sort_values(by=sort_by_column, ascending=is_ascending, na_position='last')

    # Tabela resumida
    st.header("Lista (Tabela)")
    df_tab = df_sorted.copy()
    df_tab['Desligamento'] = pd.to_datetime(df_tab['Desligamento'], errors='coerce')
    df_tab['Normalização'] = pd.to_datetime(df_tab['Normalização'], errors='coerce')
    df_tab.reset_index(inplace=True, drop=True)
    df_tab['Linha'] = df_tab.index + 1
    st.dataframe(df_tab[[
        'Linha','Categoria','Cliente','UG','Data','Hora','Tipo de ocorrência',
        'Ativo','Ocorrência','Operador','Descrição','OS'
    ]], use_container_width=True)

    # --- 13. Cards (detalhes por ocorrência) ---
    st.header("Detalhes (Cards)")

    num_cols = 4
    rows = list(df_sorted.iterrows())

    def fmt_dt(dt):
        if pd.notna(dt):
            return dt.strftime('%d/%m/%Y'), dt.strftime('%H:%M')
        return '', ''

    for i in range(0, len(rows), num_cols):
        try:
            cols = st.columns(num_cols, gap="small")
        except TypeError:
            cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j >= len(rows):
                break
            _, r = rows[i + j]
            with cols[j]:
                cliente    = html.escape(str(r.get("Cliente", "")))
                categoria  = html.escape(str(r.get("Categoria", "")))
                ug         = html.escape(str(r.get("UG", "N/A")))
                tipo       = html.escape(str(r.get("Tipo de ocorrência", "")))
                ativo      = html.escape(str(r.get("Ativo", "")))
                nome_ativo = html.escape(str(r.get("Nome Ativo", "")))
                ocorr      = html.escape(str(r.get("Ocorrência", "")))
                operador   = html.escape(str(r.get("Operador", "")))
                descricao  = html.escape(str(r.get("Descrição", ""))).replace('\n','<br>')
                protocolo  = html.escape(str(r.get("Protocolo", "")))
                osv        = html.escape(str(r.get("OS", "")))

                d_des, h_des   = fmt_dt(pd.to_datetime(r.get('Desligamento'), errors='coerce'))
                d_norm, h_norm = fmt_dt(pd.to_datetime(r.get('Normalização'), errors='coerce'))
                d_ca, h_ca     = fmt_dt(pd.to_datetime(r.get('Cliente Avisado'), errors='coerce'))
                d_loop, h_loop = fmt_dt(pd.to_datetime(r.get('Atendimento Loop'), errors='coerce'))
                d_terc, h_terc = fmt_dt(pd.to_datetime(r.get('Atendimento Terceiros'), errors='coerce'))

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
                  <div class="card-item"><span class="card-label">Data atendimento de terceiros:</span> {d_terc}</div>
                  <div class="card-item"><span class="card-label">Hora atendimento de terceiros:</span> {h_terc}</div>
                  <br>
                  <div class="card-item"><span class="card-label">Descrição:</span> {descricao}</div>
                  <div class="card-item"><span class="card-label">Protocolo:</span> {protocolo}</div>
                  <div class="card-item"><span class="card-label">OS:</span> {osv}</div>
                </div>
                """
                st.html(card_html)
else:
    st.info("Não há registros para os filtros selecionados.")