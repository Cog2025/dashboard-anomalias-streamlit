import os
from webdav3.client import Client
import io
import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
from time import timeimport html
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe


# --- 1. Configuração da Página e Layout ---
st.set_page_config(layout="wide")

if 'ui_phase' not in st.session_state:
    st.session_state.ui_phase = 'init'
if 'loading_ts' not in st.session_state:
    st.session_state.loading_ts = 0

def render_loading_overlay(ui_phase: str):
    display = 'flex' if ui_phase == 'loading' else 'none'
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
    st.session_state.loading_ts = pytime.time()
    st.rerun()

def stop_loading():
    st.session_state.ui_phase = 'ready'
    st.session_state.loading_ts = 0
    st.rerun()

render_loading_overlay()
if st.session_state.ui_phase == 'init':
    start_loading()

# Failsafe opcional (20s)
if st.session_state.ui_phase == 'loading' and (pytime.time() - st.session_state.loading_ts) > 20:
    stop_loading()


# --- 2. Dicionário para tradução dos meses ---
meses_traducao = {
    'January': 'Janeiro', 'February': 'Fevereiro', 'March': 'Março',
    'April': 'Abril', 'May': 'Maio', 'June': 'Junho',
    'July': 'Julho', 'August': 'Agosto', 'September': 'Setembro',
    'October': 'Outubro', 'November': 'Novembro', 'December': 'Dezembro'
}
meses_cronologicos = list(meses_traducao.values())


# --- 3. CSS ---
st.markdown("""
<style>
    .stButton > button {
        background-color: #28a745;
        color: white;
        font-weight: bold;
        border-radius: 5px;
        padding: 10px 20px;
        width: 100%;
        border: none;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: background-color 0.3s;
    }
    .stButton > button:hover { background-color: #218838; }
    .kpi-card {
        background-color: #333333;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 20px;
    }
    .kpi-value { font-size: 3em; font-weight: bold; color: #FF4B4B; }
    .kpi-label { font-size: 1.2em; color: #FFFFFF; }
    .stMultiSelect { max-height: 200px; overflow-y: auto; }
    .column-header { font-weight: bold; font-size: 1.2em; }
    
    /* Estilo para os cards */
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
    
    .streamlit-dataframe table td {
        word-break: break-word;
        white-space: normal;
    }

    /* Retângulo do bloco superior (OCORRÊNCIAS ATIVAS) */
    .top-block {
    border: 2px solid rgba(255,255,255,0.15);
    border-radius: 12px;
    padding: 16px 16px 8px 16px;
    margin-bottom: 28px;
    background: #111418;
    box-shadow: 0 8px 18px rgba(0,0,0,0.25);
    }
    .top-block h1 { margin-top: 0; }
    @media (max-width: 768px) {
    .top-block { padding: 12px; }
    }
    </style>
    """, unsafe_allow_html=True)

# Pequeno ajuste opcional para padding/borda dos quadros
st.markdown("""
<style>
  .boxed { padding: 12px; border-radius: 8px; }
  .boxed h4 { margin-top: 0; }
</style>
""", unsafe_allow_html=True)

# --- 4. Carregar e Tratar os Dados ---


# Define os "escopos" - as permissões que nosso script solicitará.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# Nome das planilhas que vamos ler
CREDS_FILE = "google_credentials.json"
PLANILHA_NOME_1 = "DESLIGAMENTOS"
PLANILHA_NOME_2 = "EQUIPAMENTOS"


@st.cache_resource(ttl=600)
def connect_to_google_sheets():
    # Verifica se está rodando localmente (o arquivo existe) ou na nuvem (usa st.secrets)
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
    df = pd.DataFrame(data, columns=headers)
    return df


@st.cache_data(ttl=600)
def carregar_dados_google_sheets(cache_buster: int = 0):
    try:
        client = connect_to_google_sheets()
        spreadsheet_url = "https://docs.google.com/spreadsheets/d/1KeJjbsLVP9DkxPCmNSN4VzbSBeG3SFSCAdPhir39iqg/edit?usp=sharing"
        workbook = client.open_by_url(spreadsheet_url)


        df_desligamentos = fetch_sheet_as_df(workbook.worksheet(PLANILHA_NOME_1))
        df_equipamentos = fetch_sheet_as_df(workbook.worksheet(PLANILHA_NOME_2))


        if 'IDENTIFICADOR' in df_desligamentos.columns:
            df_desligamentos['IDENTIFICADOR'] = df_desligamentos['IDENTIFICADOR'].astype(str)
        if 'IDENTIFICADOR' in df_equipamentos.columns:
            df_equipamentos['IDENTIFICADOR'] = df_equipamentos['IDENTIFICADOR'].astype(str)


        df_desligamentos.dropna(how='all', inplace=True)
        df_equipamentos.dropna(how='all', inplace=True)
        
        df_desligamentos['Categoria'] = 'DESLIGAMENTOS'
        df_equipamentos['Categoria']  = 'EQUIPAMENTOS'
        df_todos_dados = pd.concat([df_desligamentos, df_equipamentos], ignore_index=True)


        mapa_renomear = {
            'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência',
            'ATIVO': 'Ativo', 'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência',
            'QUANTIDADE': 'Quantidade', 'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização',
            'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 'DESCRIÇÃO': 'Descrição',
            'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop',
            'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo', 'CLIENTE AVISADO': 'Cliente Avisado'
        }
        colunas_atuais = df_todos_dados.columns
        renomear_final = {}
        for col in colunas_atuais:
            col_strip_upper = col.strip().upper()
            if col_strip_upper in mapa_renomear:
                renomear_final[col] = mapa_renomear[col_strip_upper]
        df_todos_dados.rename(columns=renomear_final, inplace=True)


        df_todos_dados.fillna('', inplace=True)


        # Garante que a coluna 'Cliente' existe antes de filtrar
        if 'Cliente' in df_todos_dados.columns:
            df_todos_dados = df_todos_dados[
                (df_todos_dados['Cliente'] != '') &
                (df_todos_dados['UG'] != '') &
                (df_todos_dados['Sigla'] != '')
            ].copy()
        
        colunas_datetime = ['Normalização', 'Desligamento', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']
        for col in colunas_datetime:
            if col in df_todos_dados.columns:
                df_todos_dados[col] = pd.to_datetime(df_todos_dados[col], errors='coerce')


        colunas_texto = ['Operador', 'Descrição', 'OS', 'Protocolo']
        for col in colunas_texto:
            if col in df_todos_dados.columns:
                 df_todos_dados[col] = df_todos_dados[col].astype(str).fillna('')


        # Verifica se a coluna 'Desligamento' existe e não está vazia antes de processar
        if 'Desligamento' in df_todos_dados.columns and not df_todos_dados['Desligamento'].isnull().all():
            df_todos_dados['Data'] = df_todos_dados['Desligamento'].dt.strftime('%Y-%m-%d')
            df_todos_dados['Hora'] = df_todos_dados['Desligamento'].dt.strftime('%H:%M:%S')
            df_todos_dados['Mês']  = df_todos_dados['Desligamento'].dt.strftime('%B').map(meses_traducao)
            df_todos_dados['Ano']  = df_todos_dados['Desligamento'].dt.year.fillna(0).astype(int)
            df_todos_dados['Dia']  = df_todos_dados['Desligamento'].dt.day.fillna(0).astype(int)


            df_todos_dados['ID_Unico'] = df_todos_dados['UG'].astype(str).str.upper() + "|" + \
                                    df_todos_dados['Ativo'].astype(str).str.upper() + "|" + \
                                    df_todos_dados['Ocorrência'].astype(str).str.upper() + "|" + \
                                    df_todos_dados['Desligamento'].astype(str)
        else:
            # Cria colunas vazias se 'Desligamento' não existir, para evitar erros posteriores
            for col in ['Data', 'Hora', 'Mês', 'Ano', 'Dia', 'ID_Unico']:
                df_todos_dados[col] = None



        return df_todos_dados


    except FileNotFoundError:
        st.error(f"Erro: O arquivo de credenciais '{CREDS_FILE}' não foi encontrado. Verifique se ele está na mesma pasta do seu script principal (app.py).")
        return pd.DataFrame()
    except gspread.exceptions.SpreadsheetNotFound:
        st.error("Erro: Planilha não encontrada. Verifique o link e se você compartilhou a planilha com o email da conta de serviço.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Ocorreu um erro ao carregar ou processar os dados do Google Sheets: {e}")
        return pd.DataFrame()



if 'cache_buster' not in st.session_state:
    st.session_state.cache_buster = int(pytime.time())  # novo valor a cada reload da página

df_todos_dados = carregar_dados_google_sheets(st.session_state.cache_buster)



# Garante que a coluna de data/hora está no formato correto
df_todos_dados['Desligamento'] = pd.to_datetime(df_todos_dados['Desligamento'], errors='coerce')

count_deslig = df_todos_dados[
    (df_todos_dados['Categoria'] == 'DESLIGAMENTOS') &
    (pd.isna(df_todos_dados['Normalização']) | (df_todos_dados['Normalização'] == ''))
].shape[0]

count_equip = df_todos_dados[
    (df_todos_dados['Categoria'] == 'EQUIPAMENTOS') &
    (pd.isna(df_todos_dados['Normalização']) | (df_todos_dados['Normalização'] == ''))
].shape[0]


# --- Cards gerais por categoria (fixos, sem filtros) ---
# --- Seção destacada: OCORRÊNCIAS ATIVAS ---
with st.container(border=True):
    st.markdown("<h1 style='margin:0'>OCORRÊNCIAS ATIVAS</h1>", unsafe_allow_html=True)

    col_top1, col_top2 = st.columns(2)
    with col_top1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">USINAS DESLIGADAS NO MOMENTO</div>
            <div class="kpi-value">{count_deslig}</div>
        </div>
        """, unsafe_allow_html=True)

    with col_top2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">EQUIPAMENTOS PARADOS NO MOMENTO</div>
            <div class="kpi-value">{count_equip}</div>
        </div>
        """, unsafe_allow_html=True)




# --- 5. Inicialização dos Filtros ---
if 'filtros_meses' not in st.session_state:
    st.session_state.filtros_meses = [meses_traducao[datetime.now().strftime('%B')]]
if 'filtros_anos' not in st.session_state:
    if not df_todos_dados.empty and 'Ano' in df_todos_dados.columns:
        anos_atuais = sorted(df_todos_dados['Ano'].unique().tolist())
        st.session_state.filtros_anos = [a for a in anos_atuais if a != 0]
    else:
        st.session_state.filtros_anos = []
if 'filtros_dias' not in st.session_state:
    if not df_todos_dados.empty and {'Mês','Ano'}.issubset(df_todos_dados.columns):
        dias_atuais = sorted(df_todos_dados[(df_todos_dados['Mês'].isin(st.session_state.filtros_meses)) & (df_todos_dados['Ano'].isin(st.session_state.filtros_anos))]['Dia'].unique().tolist())
        st.session_state.filtros_dias = [d for d in dias_atuais if d != 0]
    else:
        st.session_state.filtros_dias = []
if 'filtros_categorias' not in st.session_state:
    st.session_state.filtros_categorias = sorted(df_todos_dados['Categoria'].unique().tolist()) if not df_todos_dados.empty else []
if 'filtros_clientes' not in st.session_state:
    st.session_state.filtros_clientes = sorted(df_todos_dados['Cliente'].unique().tolist()) if not df_todos_dados.empty else []
if 'filtros_ugs' not in st.session_state:
    st.session_state.filtros_ugs = sorted(df_todos_dados['UG'].unique().tolist()) if not df_todos_dados.empty else []
if 'filtros_tipos' not in st.session_state:
    st.session_state.filtros_tipos = sorted(df_todos_dados['Tipo de ocorrência'].unique().tolist()) if not df_todos_dados.empty else []
if 'filtros_ativos' not in st.session_state:
    st.session_state.filtros_ativos = sorted(df_todos_dados['Ativo'].unique().tolist()) if not df_todos_dados.empty else []
if 'filtros_ocorrencias' not in st.session_state:
    st.session_state.filtros_ocorrencias = sorted(df_todos_dados['Ocorrência'].unique().tolist()) if not df_todos_dados.empty else []


# --- 6. Título e KPIs ---
st.header('OCORRÊNCIAS FILTRADAS')
col_kpi1, col_kpi2 = st.columns(2)
with col_kpi1:
    if not df_todos_dados.empty and 'Normalização' in df_todos_dados.columns:
        df_desligadas_geral = df_todos_dados[pd.isna(df_todos_dados['Normalização']) | (df_todos_dados['Normalização'] == '')].copy()
        total_kpi_value = df_desligadas_geral.shape[0]
    else:
        total_kpi_value = 0
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total no Banco de Dados Completo</div>
        <div class="kpi-value">{total_kpi_value}</div>
    </div>
    """, unsafe_allow_html=True)


# --- 7. Botão de Atualização ---
col_top_left, col_top_right = st.columns([0.2, 0.8])
with col_top_left:
    if st.button('Atualizar Dados'):
        st.cache_data.clear()
        st.session_state.ui_phase = 'init'
        st.session_state.loading_ts = 0
        st.rerun()


# --- 8. Interface de Filtros ---
# Utilitário: marca/desmarca e mantém consistência entre filtros e checkboxes
def _marcar(prefixo_key: str, itens: list, filtro_key: str, marcar_todos: bool, validos: set | None = None):
    validos = set(itens) if validos is None else set(validos)
    # Atualiza lista
    st.session_state[filtro_key] = [x for x in itens if (x in validos) and marcar_todos]
    # Atualiza checkboxes correspondentes
    for x in itens:
        st.session_state[f"{prefixo_key}{x}"] = marcar_todos and (x in validos)


if not df_todos_dados.empty:
    if 'categoria_top' not in st.session_state:
        st.session_state['categoria_top'] = 'Ambas'   # inicializa só uma vez

    st.markdown("#### Filtrar por categoria (planilha)")
    st.radio(
        "Categoria:",
        options=["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"],
        horizontal=True,
        label_visibility="collapsed",
        key="categoria_top"   # sem index
    )

    st.subheader("Selecione o período desejado")
    # imediatamente após st.subheader("Selecione o período desejado")
    anos_disponiveis = sorted([a for a in df_todos_dados['Ano'].unique() if a != 0])
    meses_disponiveis = meses_cronologicos[:]  # cria alias local que o bloco usa

    col_ano, col_mes, col_dia = st.columns(3)

    # --- Ano(s) ---
    with col_ano:
        with st.container(border=True):
            st.write("### Ano(s):")
            with st.expander("Expandir anos"):
                for ano in anos_disponiveis:
                    st.checkbox(str(ano), key=f'cb_ano_{ano}', value=(ano in st.session_state.filtros_anos))
            col_botoes = st.columns(2)
            clicked_sel_ano = col_botoes[0].button('Sel. Todos', key='sel_ano', use_container_width=True,
                                                on_click=_marcar, args=('cb_ano_', anos_disponiveis, 'filtros_anos', True))
            clicked_des_ano = col_botoes[1].button('Desmarcar', key='des_ano', use_container_width=True,
                                                on_click=_marcar, args=('cb_ano_', anos_disponiveis, 'filtros_anos', False))
            if not (clicked_sel_ano or clicked_des_ano):
                st.session_state.filtros_anos = [a for a in anos_disponiveis if st.session_state.get(f'cb_ano_{a}', False)]

    # --- Mês(es) ---
    with col_mes:
        with st.container(border=True):
            st.write("### Mês(es):")
            with st.expander("Expandir meses"):
                for mes in meses_disponiveis:
                    st.checkbox(mes, key=f'cb_mes_{mes}', value=(mes in st.session_state.filtros_meses))
            col_botoes = st.columns(2)
            clicked_sel_mes = col_botoes[0].button('Sel. Todos', key='sel_mes', use_container_width=True,
                                                on_click=_marcar, args=('cb_mes_', meses_disponiveis, 'filtros_meses', True))
            clicked_des_mes = col_botoes[1].button('Desmarcar', key='des_mes', use_container_width=True,
                                                on_click=_marcar, args=('cb_mes_', meses_disponiveis, 'filtros_meses', False))
            if not (clicked_sel_mes or clicked_des_mes):
                st.session_state.filtros_meses = [m for m in meses_disponiveis if st.session_state.get(f'cb_mes_{m}', False)]

    # --- Dia(s) ---
    if 'dias_disponiveis' not in locals():
        # usa todos os dias 1..31 na primeira carga
        dias_disponiveis = list(range(1, 32))
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
            col_botoes = st.columns(2)
            clicked_sel_dia = col_botoes[0].button('Sel. Todos', key='sel_dia', use_container_width=True,
                                                on_click=_marcar, args=('cb_dia_', list(range(1, 32)), 'filtros_dias', True, set(dias_disponiveis)))
            clicked_des_dia = col_botoes[1].button('Desmarcar', key='des_dia', use_container_width=True,
                                                on_click=_marcar, args=('cb_dia_', list(range(1, 32)), 'filtros_dias', False, set(dias_disponiveis)))
            if not (clicked_sel_dia or clicked_des_dia):
                st.session_state.filtros_dias = [d for d in dias_disponiveis if st.session_state.get(f'cb_dia_{d}', False)]



    st.subheader("Filtros Adicionais")
    col_cliente, col_ug, col_tipo, col_ativo, col_ocorrencia = st.columns(5)

    with col_cliente:
        with st.container(border=True):
            st.write("Cliente:")
            col_b = st.columns(2)
            with col_b[0]:
                if st.button('Sel. Todos', key='sel_cli', use_container_width=True):
                    st.session_state.filtros_clientes = sorted(df_todos_dados['Cliente'].unique().tolist()); st.rerun()
            with col_b[1]:
                if st.button('Desmarcar', key='des_cli', use_container_width=True):
                    st.session_state.filtros_clientes = []; st.rerun()
            st.session_state.filtros_clientes = st.multiselect(' ', options=sorted(df_todos_dados['Cliente'].unique().tolist()),
                                                            default=st.session_state.filtros_clientes, label_visibility='hidden')

    with col_ug:
        with st.container(border=True):
            st.write("UG:")
            df_temp = df_todos_dados[df_todos_dados['Cliente'].isin(st.session_state.filtros_clientes)]
            ugs_disponiveis = sorted(df_temp['UG'].unique().tolist())
            st.session_state.filtros_ugs = [ug for ug in st.session_state.filtros_ugs if ug in ugs_disponiveis]
            col_b = st.columns(2)
            with col_b[0]:
                if st.button('Sel. Todos', key='sel_ug', use_container_width=True):
                    st.session_state.filtros_ugs = ugs_disponiveis; st.rerun()
            with col_b[1]:
                if st.button('Desmarcar', key='des_ug', use_container_width=True):
                    st.session_state.filtros_ugs = []; st.rerun()
            st.session_state.filtros_ugs = st.multiselect(' ', options=ugs_disponiveis,
                                                        default=st.session_state.filtros_ugs, label_visibility='hidden')

    with col_tipo:
        with st.container(border=True):
            st.write("Tipo de Ocorrência:")
            col_b = st.columns(2)
            with col_b[0]:
                if st.button('Sel. Todos', key='sel_tipo', use_container_width=True):
                    st.session_state.filtros_tipos = sorted(df_todos_dados['Tipo de ocorrência'].unique().tolist()); st.rerun()
            with col_b[1]:
                if st.button('Desmarcar', key='des_tipo', use_container_width=True):
                    st.session_state.filtros_tipos = []; st.rerun()
            st.session_state.filtros_tipos = st.multiselect(' ', options=sorted(df_todos_dados['Tipo de ocorrência'].unique().tolist()),
                                                            default=st.session_state.filtros_tipos, label_visibility='hidden')

    with col_ativo:
        with st.container(border=True):
            st.write("Ativo:")
            col_b = st.columns(2)
            with col_b[0]:
                if st.button('Sel. Todos', key='sel_ativo', use_container_width=True):
                    st.session_state.filtros_ativos = sorted(df_todos_dados['Ativo'].unique().tolist()); st.rerun()
            with col_b[1]:
                if st.button('Desmarcar', key='des_ativo', use_container_width=True):
                    st.session_state.filtros_ativos = []; st.rerun()
            st.session_state.filtros_ativos = st.multiselect(' ', options=sorted(df_todos_dados['Ativo'].unique().tolist()),
                                                            default=st.session_state.filtros_ativos, label_visibility='hidden')

    with col_ocorrencia:
        with st.container(border=True):
            st.write("Ocorrência:")
            col_b = st.columns(2)
            with col_b[0]:
                if st.button('Sel. Todos', key='sel_ocorr', use_container_width=True):
                    st.session_state.filtros_ocorrencias = sorted(df_todos_dados['Ocorrência'].unique().tolist()); st.rerun()
            with col_b[1]:
                if st.button('Desmarcar', key='des_ocorr', use_container_width=True):
                    st.session_state.filtros_ocorrencias = []; st.rerun()
            st.session_state.filtros_ocorrencias = st.multiselect(' ', options=sorted(df_todos_dados['Ocorrência'].unique().tolist()),
                                                                default=st.session_state.filtros_ocorrencias, label_visibility='hidden')



    # --- Aplicação dos Filtros ---
    meses_selecionados = [mes for mes in meses_cronologicos if st.session_state.get(f'cb_mes_{mes}', False)]
    anos_selecionados  = [ano for ano in anos_disponiveis     if st.session_state.get(f'cb_ano_{ano}', False)]
    dias_selecionados  = [dia for dia in range(1, 32)         if st.session_state.get(f'cb_dia_{dia}', False)]

    # Conjuntos disponíveis no período atual
    set_anos_disp  = set(anos_disponiveis)
    set_meses_disp = set(meses_cronologicos)
    # Dias disponíveis dependem do período atual; use os que você já calculou
    # dias_disponiveis já existe acima no seu código
    set_dias_disp  = set(dias_disponiveis)

    # Detecta se o usuário realmente selecionou tudo em cada dimensão
    all_anos  = set(anos_selecionados)  == set_anos_disp and len(set_anos_disp) > 0
    all_meses = set(meses_selecionados) == set_meses_disp and len(set_meses_disp) > 0
    all_dias  = set(dias_selecionados)  == set_dias_disp and len(set_dias_disp) > 0

    s_ano = df_todos_dados['Ano']
    s_mes = df_todos_dados['Mês']
    s_dia = df_todos_dados['Dia']

    # Máscaras de período (sempre ativas); quando "tudo" for selecionado, inclui também ausentes/0
    m_ano = s_ano.isin(anos_selecionados)  if not all_anos  else (s_ano.isin(anos_selecionados)  | s_ano.isna() | (s_ano == 0))
    m_mes = s_mes.isin(meses_selecionados) if not all_meses else (s_mes.isin(meses_selecionados) | s_mes.isna() | (s_mes.astype(str) == ''))
    m_dia = s_dia.isin(dias_selecionados)  if not all_dias  else (s_dia.isin(dias_selecionados)  | s_dia.isna() | (s_dia == 0))

    m_cat = df_todos_dados['Categoria'].isin(st.session_state.filtros_categorias)
    # Respeita seletor de categoria do topo
    if st.session_state.get("categoria_top") in ("DESLIGAMENTOS", "EQUIPAMENTOS"):
        m_cat = m_cat & (df_todos_dados['Categoria'] == st.session_state["categoria_top"])

    m_cli = df_todos_dados['Cliente'].isin(st.session_state.filtros_clientes)
    m_ug  = df_todos_dados['UG'].isin(st.session_state.filtros_ugs)
    m_tip = df_todos_dados['Tipo de ocorrência'].isin(st.session_state.filtros_tipos)
    m_atv = df_todos_dados['Ativo'].isin(st.session_state.filtros_ativos)
    m_ocr = df_todos_dados['Ocorrência'].isin(st.session_state.filtros_ocorrencias)

    df_filtrado = df_todos_dados[m_ano & m_mes & m_dia & m_cat & m_cli & m_ug & m_tip & m_atv & m_ocr].copy()
    df_desligadas = df_filtrado[pd.isna(df_filtrado['Normalização']) | (df_filtrado['Normalização'] == '')].copy()
    
    with col_kpi2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total com Filtro Selecionado</div>
            <div class="kpi-value">{len(df_desligadas)}</div>
        </div>
        """, unsafe_allow_html=True)
    
    if not df_desligadas.empty:
        mask_valid = df_desligadas['Desligamento'].notna()
        df_desligadas.loc[mask_valid, 'Tempo em Segundos'] = (
            (datetime.now() - df_desligadas.loc[mask_valid, 'Desligamento']).dt.total_seconds().astype(int)
        )
        df_desligadas.loc[~mask_valid, 'Tempo em Segundos'] = 0



        # --- CONTROLES DE ORDENAÇÃO ---
        st.markdown("---")
        st.write("### Ordenar e Editar")
        sort_cols = st.columns(2)
        
        with sort_cols[0]:
            sort_options_display = {
                'Data do Desligamento': 'Desligamento',
                'Tempo de Desligamento': 'Tempo em Segundos',
                'UG': 'UG',
                'Ativo': 'Ativo'
            }
            sort_by_display = st.selectbox(
                "Ordenar por:",
                options=sort_options_display.keys(), index=0)
            sort_by_column = sort_options_display[sort_by_display]


        with sort_cols[1]:
            sort_order = st.radio(
                "Ordem:",
                options=['Descendente', 'Ascendente'], index=0, horizontal=True)
            is_ascending = (sort_order == 'Ascendente')


        df_sorted = df_desligadas.sort_values(by=sort_by_column, ascending=is_ascending, na_position='last')

        # Criamos uma coluna 'Display' para facilitar a seleção no selectbox
        df_sorted['Display'] = (
            df_sorted['UG'].astype(str) + " | " + df_sorted['Ativo'].astype(str) + " | " +
            df_sorted['Nome Ativo'].astype(str) + " | " + df_sorted['Ocorrência'].astype(str) + " | " +
            df_sorted['Desligamento'].dt.strftime('%d/%m/%Y %H:%M').fillna('') + 
            "  ·  " + df_sorted['ID_Unico'].astype(str).str[-6:]
        )
        
        # Disponibiliza lista filtrada/ordenada para a página de edição
        cols_minimos = ['ID_Unico','UG','Ativo','Nome Ativo','Ocorrência','Desligamento','Categoria',
                        'Tipo de ocorrência','Operador','Descrição','OS','Protocolo',
                        'Normalização','Atendimento Loop','Atendimento Terceiros','Cliente Avisado']
        cols_salvar = [c for c in cols_minimos if c in df_sorted.columns] + ['Display']
        st.session_state['df_lista_para_editar'] = df_sorted[cols_salvar].copy()



        # ***** NOVO: SELEÇÃO PARA EDIÇÃO *****
        st.markdown("---")
        st.write("### Editar uma Ocorrência")


        opts = df_sorted['Display'].dropna().astype(str).tolist()
        ocorrencia_selecionada_display = st.selectbox(
            "Selecione a ocorrência para editar:",
            options=opts,
            index=None,
            placeholder="Escolha uma ocorrência..."
        )

        # Prepara o ID quando houver seleção
        if ocorrencia_selecionada_display:
            id_unico_para_editar = df_sorted.loc[
                df_sorted['Display'] == ocorrencia_selecionada_display, 'ID_Unico'
            ].head(1).item()
            st.session_state['id_unico_para_editar'] = id_unico_para_editar

        # Limpa o ID se a seleção for removida
        if not ocorrencia_selecionada_display and 'id_unico_para_editar' in st.session_state:
            st.session_state.pop('id_unico_para_editar')

        # Botão sempre visível, habilita só com seleção
        btn_disabled = not bool(ocorrencia_selecionada_display)
        if st.button("📝 Editar Ocorrência Selecionada", disabled=btn_disabled):
            st.switch_page("pages/3_Editar_Ocorrência.py")


        
        # --- LISTA DE OCORRÊNCIAS (TABELA) ---
        st.header("Lista de Ocorrências (Tabela)")
        df_para_tabela = df_sorted.copy()
        
        def formatar_tempo_estatico(row):
            dias = row['Tempo em Segundos'] // 86400
            horas = (row['Tempo em Segundos'] % 86400) // 3600
            minutos = (row['Tempo em Segundos'] % 3600) // 60
            return f"{dias}d {horas}h {minutos}m"
        
        df_para_tabela['Tempo de Desligamento'] = df_para_tabela.apply(formatar_tempo_estatico, axis=1)
        df_para_tabela.reset_index(inplace=True, drop=True)
        df_para_tabela['Linha'] = df_para_tabela.index + 1
        
        st.dataframe(df_para_tabela[[
            'Linha', 'Categoria', 'Tempo de Desligamento', 'UG', 'Data', 'Hora', 'Tipo de ocorrência', 
            'Ativo', 'Ocorrência', 'Operador', 'Descrição', 'OS'
        ]], use_container_width=True)


        # --- DETALHES POR OCORRÊNCIA (CARDS) ---
        st.header("Detalhes por Ocorrência (Cards)")
        
        num_cols = 4
        rows = list(df_sorted.iterrows())
        
        def format_datetime_card(dt_obj):
            if pd.notna(dt_obj):
                return dt_obj.strftime('%d/%m/%Y'), dt_obj.strftime('%H:%M')
            return '', ''


        for i in range(0, len(rows), num_cols):
            cols = st.columns(num_cols)
            for j in range(num_cols):
                if i + j < len(rows):
                    index, row = rows[i + j]
                    with cols[j]:
                        cliente   = html.escape(str(row.get("Cliente", "")))
                        categoria = html.escape(str(row.get("Categoria", "")))
                        ug = html.escape(str(row.get("UG", "N/A")))
                        tipo_ocorrencia = html.escape(str(row.get("Tipo de ocorrência", "")))
                        ativo = html.escape(str(row.get("Ativo", "")))
                        nome_ativo = html.escape(str(row.get("Nome Ativo", "")))
                        ocorrencia = html.escape(str(row.get("Ocorrência", "")))
                        operador = html.escape(str(row.get("Operador", "")))
                        descricao = html.escape(str(row.get("Descrição", ""))).replace('\n', '<br>')
                        protocolo = html.escape(str(row.get("Protocolo", "")))
                        os = html.escape(str(row.get("OS", "")))


                        data_ocor, hora_ocor = format_datetime_card(row.get('Desligamento'))
                        data_ca, hora_ca = format_datetime_card(row.get('Cliente Avisado'))
                        data_loop, hora_loop = format_datetime_card(row.get('Atendimento Loop'))
                        data_terc, hora_terc = format_datetime_card(row.get('Atendimento Terceiros'))
                        data_norm, hora_norm = format_datetime_card(row.get('Normalização'))


                        quantidade_html = ''
                        if row.get('Categoria') == 'EQUIPAMENTOS':
                            quantidade_val = row.get('Quantidade', 0)
                            try:
                                if pd.notna(quantidade_val) and float(quantidade_val) > 0:
                                    quantidade_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {int(float(quantidade_val))}</div>'
                            except (ValueError, TypeError):
                                quantidade_html = ''


                        card_html = f"""
                        <div class="card-container">
                            <div class="card-title">{ug}</div>
                            <div class="card-item"><span class="card-label">Cliente:</span> {cliente}</div>
                            <div class="card-item"><span class="card-label">Categoria:</span> {categoria}</div>
                            <div class="card-item"><span class="card-label">Tipo de Ocorrência:</span> {tipo_ocorrencia}</div>
                            <div class="card-item"><span class="card-label">Ativo:</span> {ativo}</div>
                            <div class="card-item"><span class="card-label">Nome do ativo:</span> {nome_ativo}</div>
                            <div class="card-item"><span class="card-label">Ocorrência:</span> {ocorrencia}</div>
                            <div class="card-item"><span class="card-label">Operador:</span> {operador}</div>
                            {quantidade_html}
                            <br>
                            <div class="card-item"><span class="card-label">Data da ocorrência:</span> {data_ocor}</div>
                            <div class="card-item"><span class="card-label">Hora da ocorrência:</span> {hora_ocor}</div>
                            <div class="card-item"><span class="card-label">Data cliente avisado:</span> {data_ca}</div>
                            <div class="card-item"><span class="card-label">Hora cliente avisado:</span> {hora_ca}</div>
                            <div class="card-item"><span class="card-label">Data do atendimento LOOP:</span> {data_loop}</div>
                            <div class="card-item"><span class="card-label">Hora do atendimento LOOP:</span> {hora_loop}</div>
                            <div class="card-item"><span class="card-label">Data do atendimento de terceiros:</span> {data_terc}</div>
                            <div class="card-item"><span class="card-label">Hora do atendimento de terceiros:</span> {hora_terc}</div>
                            <div class="card-item"><span class="card-label">Data de normalização:</span> {data_norm}</div>
                            <div class="card-item"><span class="card-label">Hora de normalização:</span> {hora_norm}</div>
                            <br>
                            <div class="card-item"><span class="card-label">Descrição:</span> {descricao}</div>
                            <div class="card-item"><span class="card-label">Protocolo:</span> {protocolo}</div>
                            <div class="card-item"><span class="card-label">OS:</span> {os}</div>
                        </div>
                        """
                        st.html(card_html)
                
    else:
        st.info("Nenhuma usina encontrada com o campo 'Normalização' em branco para os filtros selecionados.")
else:
    st.warning("Não foi possível carregar os dados. Verifique o arquivo local ou os filtros aplicados.")
