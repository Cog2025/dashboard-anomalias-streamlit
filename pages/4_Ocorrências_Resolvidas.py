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

PAGE_ID = "p4"  # na 1_Pagina_Principal.py; use "p4" na 4_Ocorrencias_Resolvidas.py
def K(name: str) -> str:
    return f"{PAGE_ID}:{name}"

# --- 1. Configuração da Página e Layout ---
st.set_page_config(layout="wide")

# === Overlay de carregamento (logo após st.set_page_config) ===
if 'loading' not in st.session_state:
    st.session_state.loading = True

def render_loading_overlay():
    display = 'flex' if st.session_state.get('loading', False) else 'none'
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

render_loading_overlay()


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
    st.session_state.cache_buster = int(time())

# 1) Carrega dados
df = carregar_dados_google_sheets(st.session_state.cache_buster)

# 2) Valida e para cedo se vazio
if df is None or df.empty:
    st.warning("Não foi possível carregar os dados. Verifique as credenciais ou a conexão.")
    # Desliga overlay antes de parar, para não deixar a tela escura em navegação
    if st.session_state.get('loading', False):
        st.session_state.loading = False
    st.stop()

# 3) Desliga overlay e rerun (apenas uma vez após termos df válido)
if st.session_state.get('loading', False):
    st.session_state.loading = False
    st.rerun()

# 4) A partir daqui, estados e UI (somente quando df é válido)
# Inicialização de estados (namespaced por página)
if K('filtros_meses') not in st.session_state:
    st.session_state[K('filtros_meses')] = [meses_traducao[datetime.now().strftime('%B')]]

if K('filtros_anos') not in st.session_state:
    anos_atuais = sorted(df['Ano'].unique().tolist()) if 'Ano' in df.columns else []
    st.session_state[K('filtros_anos')] = [a for a in anos_atuais if a != 0]

if K('filtros_dias') not in st.session_state:
    if {'Mês','Ano'}.issubset(df.columns):
        dias_atuais = sorted(
            df[
                (df['Mês'].isin(st.session_state[K('filtros_meses')])) &
                (df['Ano'].isin(st.session_state[K('filtros_anos')]))
            ]['Dia'].unique().tolist()
        )
        st.session_state[K('filtros_dias')] = [d for d in dias_atuais if d != 0]
    else:
        st.session_state[K('filtros_dias')] = []




df['Desligamento'] = pd.to_datetime(df['Desligamento'], errors='coerce')

# --- 5. KPIs do topo (RESOLVIDAS) ---
count_resolvidos_deslig = df[
    (df['Categoria'] == 'DESLIGAMENTOS') &
    (~df['Normalização'].isna())
].shape[0]

count_resolvidos_equip = df[
    (df['Categoria'] == 'EQUIPAMENTOS') &
    (~df['Normalização'].isna())
].shape[0]

with st.container(border=True):
    st.markdown("<h1 style='margin:0'>OCORRÊNCIAS RESOLVIDAS</h1>", unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">DESLIGAMENTOS NORMALIZADOS (Banco Completo)</div>
            <div class="kpi-value">{count_resolvidos_deslig}</div>
        </div>
        """, unsafe_allow_html=True)
    with col_b:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">EQUIPAMENTOS NORMALIZADOS (Banco Completo)</div>
            <div class="kpi-value">{count_resolvidos_equip}</div>
        </div>
        """, unsafe_allow_html=True)

# --- 6. Filtros (idênticos à principal) ---
st.header('OCORRÊNCIAS FILTRADAS')

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
    if st.button('Atualizar Dados'):
        st.session_state.loading = True
        st.session_state.cache_buster = int(time())
        st.cache_data.clear()
        st.rerun()


def _sanitize_default(default_list, valid_options):
    # Garante que o default só contenha itens válidos e remove duplicatas preservando ordem
    if not isinstance(default_list, (list, tuple)):
        return []
    seen = set()
    result = []
    valid_set = set(valid_options)
    for x in default_list:
        if x in valid_set and x not in seen:
            seen.add(x)
            result.append(x)
    return result

# === Helpers (sanitização e sincronização) ===
# Sincroniza o estado do widget (key_ms) e o estado de modelo (key_modelo)
def _sync_ms(key_ms: str, key_modelo: str, valores: list):
    st.session_state[key_modelo] = list(valores)
    st.session_state[key_ms] = list(valores)

# Marcação em massa de checkboxes (já existente abaixo): _marcar(...)


# Helper para marcar/desmarcar checkboxes em massa
def _marcar(prefixo_key: str, itens: list, filtro_key: str, marcar_todos: bool, validos: set | None = None):
    validos = set(itens) if validos is None else set(validos)
    st.session_state[filtro_key] = [x for x in itens if (x in validos) and marcar_todos]
    for x in itens:
        st.session_state[f"{prefixo_key}{x}"] = marcar_todos and (x in validos)

if not df.empty:

    # Categoria global (persistente entre páginas)
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
    m_cat_base = df['Categoria'].isin(
        st.session_state.get(K('filtros_categorias'), df['Categoria'].unique().tolist())
    )
    m_cat = (df['Categoria'] == cat_sel) if cat_sel in ("DESLIGAMENTOS", "EQUIPAMENTOS") else m_cat_base

    # Máscaras adicionais (namespaced)
    m_cli = df['Cliente'].isin(st.session_state.get(K('filtros_clientes'), []))
    m_ug  = df['UG'].isin(st.session_state.get(K('filtros_ugs'), []))
    m_tip = df['Tipo de ocorrência'].isin(st.session_state.get(K('filtros_tipos'), []))
    m_atv = df['Ativo'].isin(st.session_state.get(K('filtros_ativos'), []))
    m_ocr = df['Ocorrência'].isin(st.session_state.get(K('filtros_ocorrencias'), []))

    st.subheader("Selecione o período desejado")
    anos_disponiveis = sorted([a for a in df['Ano'].unique() if a != 0])
    meses_disponiveis = meses_cronologicos[:]

    col_ano, col_mes, col_dia = st.columns(3)

    # Ano(s)
    with col_ano:
        with st.container(border=True):
            st.write("### Ano(s):")

            # checkboxes com key única por página
            with st.expander("Expandir anos"):
                for ano in anos_disponiveis:
                    st.checkbox(
                        str(ano),
                        key=f'cb_ano_{PAGE_ID}_{ano}',
                        value=(ano in st.session_state.get(K('filtros_anos'), []))
                    )

            # botões com key única por página
            c = st.columns(2)
            clicked_sel_ano = c[0].button('Sel. Todos', key=f'sel_ano_{PAGE_ID}', use_container_width=True,
                                        on_click=_marcar, args=(f'cb_ano_{PAGE_ID}_', anos_disponiveis, K('filtros_anos'), True))
            clicked_des_ano = c[1].button('Desmarcar', key=f'des_ano_{PAGE_ID}', use_container_width=True,
                                        on_click=_marcar, args=(f'cb_ano_{PAGE_ID}_', anos_disponiveis, K('filtros_anos'), False))

            # se não clicou nos botões, lê os checks deste PAGE_ID
            if not (clicked_sel_ano or clicked_des_ano):
                st.session_state[K('filtros_anos')] = [
                    a for a in anos_disponiveis if st.session_state.get(f'cb_ano_{PAGE_ID}_{a}', False)
                ]


    # Mês(es)
    with col_mes:
        with st.container(border=True):
            st.write("### Mês(es):")

            # Checkboxes com keys únicas por página
            with st.expander("Expandir meses"):
                for mes in meses_disponiveis:
                    st.checkbox(
                        mes,
                        key=f'cb_mes_{PAGE_ID}_{mes}',
                        value=(mes in st.session_state.get(K('filtros_meses'), []))
                    )

            # Botões com keys únicas por página
            c = st.columns(2)
            clicked_sel_mes = c[0].button(
                'Sel. Todos', key=f'sel_mes_{PAGE_ID}', use_container_width=True,
                on_click=_marcar, args=(f'cb_mes_{PAGE_ID}_', meses_disponiveis, K('filtros_meses'), True)
            )
            clicked_des_mes = c[1].button(
                'Desmarcar', key=f'des_mes_{PAGE_ID}', use_container_width=True,
                on_click=_marcar, args=(f'cb_mes_{PAGE_ID}_', meses_disponiveis, K('filtros_meses'), False)
            )

            # Se não clicou nos botões, lê os checks deste PAGE_ID
            if not (clicked_sel_mes or clicked_des_mes):
                st.session_state[K('filtros_meses')] = [
                    m for m in meses_disponiveis if st.session_state.get(f'cb_mes_{PAGE_ID}_{m}', False)
                ]


    # Dia(s)
    if 'dias_disponiveis' not in locals():
        dias_disponiveis = list(range(1, 32))

    with col_dia:
        with st.container(border=True):
            st.write("### Dia(s):")

            # Checkboxes com grid e keys únicas por página
            with st.expander("Expandir dias"):
                dias_cols = st.columns(7)
                for i, dia in enumerate(range(1, 32)):
                    with dias_cols[i % 7]:
                        if dia in dias_disponiveis:
                            st.checkbox(
                                str(dia),
                                key=f'cb_dia_{PAGE_ID}_{dia}',
                                value=(dia in st.session_state.get(K('filtros_dias'), []))
                            )
                        else:
                            st.checkbox(str(dia), key=f'cb_dia_{PAGE_ID}_{dia}', disabled=True)

            # Botões com keys únicas por página
            c = st.columns(2)
            clicked_sel_dia = c[0].button(
                'Sel. Todos', key=f'sel_dia_{PAGE_ID}', use_container_width=True,
                on_click=_marcar, args=(f'cb_dia_{PAGE_ID}_', list(range(1, 32)), K('filtros_dias'), True, set(dias_disponiveis))
            )
            clicked_des_dia = c[1].button(
                'Desmarcar', key=f'des_dia_{PAGE_ID}', use_container_width=True,
                on_click=_marcar, args=(f'cb_dia_{PAGE_ID}_', list(range(1, 32)), K('filtros_dias'), False, set(dias_disponiveis))
            )

            # Se não clicou nos botões, lê os checks deste PAGE_ID e respeita disponibilidade
            if not (clicked_sel_dia or clicked_des_dia):
                st.session_state[K('filtros_dias')] = [
                    d for d in dias_disponiveis if st.session_state.get(f'cb_dia_{PAGE_ID}_{d}', False)
                ]

    # Filtros adicionais
    st.subheader("Filtros Adicionais")
    col_cliente, col_ug, col_tipo, col_ativo, col_ocr = st.columns(5)

    with col_cliente:
        with st.container(border=True):
            st.write("Cliente:")
            options_clientes = sorted(df['Cliente'].unique().tolist())
            default_clientes = _sanitize_default(st.session_state.get(K('filtros_clientes'), []), options_clientes)
            st.session_state[K('filtros_clientes')] = default_clientes

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
                default=st.session_state.get(key_ms_cli, default_clientes),
                label_visibility='hidden',
                key=key_ms_cli
            )



    with col_ug:
        with st.container(border=True):
            st.write("UG:")
            df_temp = df[df['Cliente'].isin(st.session_state.get(K('filtros_clientes'), []))]
            ugs_disp = sorted(df_temp['UG'].unique().tolist())

            default_ugs = _sanitize_default(st.session_state.get(K('filtros_ugs'), []), ugs_disp)
            st.session_state[K('filtros_ugs')] = default_ugs

            key_ms_ug = f'ms_ugs_{PAGE_ID}'
            if key_ms_ug not in st.session_state:
                st.session_state[key_ms_ug] = st.session_state[K('filtros_ugs')]

            c = st.columns(2)
            if c[0].button('Sel. Todos', key=f'sel_ug_{PAGE_ID}', use_container_width=True):
                _sync_ms(key_ms_ug, K('filtros_ugs'), ugs_disp)
                st.rerun()
            if c[1].button('Desmarcar', key=f'des_ug_{PAGE_ID}', use_container_width=True):
                _sync_ms(key_ms_ug, K('filtros_ugs'), [])
                st.rerun()

            st.session_state[K('filtros_ugs')] = st.multiselect(
                ' ', options=ugs_disp,
                default=st.session_state.get(key_ms_ug, default_ugs),
                label_visibility='hidden',
                key=key_ms_ug
            )



    with col_tipo:
        with st.container(border=True):
            st.write("Tipo de Ocorrência:")

            # Opções e default sanitizado
            opts_tipos = sorted(df['Tipo de ocorrência'].unique().tolist())
            default_tipos = _sanitize_default(st.session_state.get(K('filtros_tipos'), []), opts_tipos)
            st.session_state[K('filtros_tipos')] = default_tipos

            # Key do widget do multiselect (estado do componente)
            key_ms_tipos = f'ms_tipos_{PAGE_ID}'
            if key_ms_tipos not in st.session_state:
                st.session_state[key_ms_tipos] = st.session_state[K('filtros_tipos')]

            # Botões de seleção/limpeza sincronizando UI + modelo
            c = st.columns(2)
            if c[0].button('Sel. Todos', key=f'sel_tipo_{PAGE_ID}', use_container_width=True):
                _sync_ms(key_ms_tipos, K('filtros_tipos'), opts_tipos)
                st.rerun()
            if c[1].button('Desmarcar', key=f'des_tipo_{PAGE_ID}', use_container_width=True):
                _sync_ms(key_ms_tipos, K('filtros_tipos'), [])
                st.rerun()

            # Multiselect lê o estado do widget (key_ms_tipos)
            st.session_state[K('filtros_tipos')] = st.multiselect(
                ' ', options=opts_tipos,
                default=st.session_state.get(key_ms_tipos, default_tipos),
                label_visibility='hidden',
                key=key_ms_tipos
            )


    with col_ativo:
        with st.container(border=True):
            st.write("Ativo:")

            opts_ativos = sorted(df['Ativo'].unique().tolist())
            default_ativos = _sanitize_default(st.session_state.get(K('filtros_ativos'), []), opts_ativos)
            st.session_state[K('filtros_ativos')] = default_ativos

            key_ms_ativos = f'ms_ativos_{PAGE_ID}'
            if key_ms_ativos not in st.session_state:
                st.session_state[key_ms_ativos] = st.session_state[K('filtros_ativos')]

            c = st.columns(2)
            if c[0].button('Sel. Todos', key=f'sel_ativo_{PAGE_ID}', use_container_width=True):
                _sync_ms(key_ms_ativos, K('filtros_ativos'), opts_ativos)
                st.rerun()
            if c[1].button('Desmarcar', key=f'des_ativo_{PAGE_ID}', use_container_width=True):
                _sync_ms(key_ms_ativos, K('filtros_ativos'), [])
                st.rerun()

            st.session_state[K('filtros_ativos')] = st.multiselect(
                ' ', options=opts_ativos,
                default=st.session_state.get(key_ms_ativos, default_ativos),
                label_visibility='hidden',
                key=key_ms_ativos
            )


    with col_ocr:
        with st.container(border=True):
            st.write("Ocorrência:")

            opts_ocr = sorted(df['Ocorrência'].unique().tolist())
            default_ocr = _sanitize_default(st.session_state.get(K('filtros_ocorrencias'), []), opts_ocr)
            st.session_state[K('filtros_ocorrencias')] = default_ocr

            key_ms_ocr = f'ms_ocorrencias_{PAGE_ID}'
            if key_ms_ocr not in st.session_state:
                st.session_state[key_ms_ocr] = st.session_state[K('filtros_ocorrencias')]

            c = st.columns(2)
            if c[0].button('Sel. Todos', key=f'sel_ocr_{PAGE_ID}', use_container_width=True):
                _sync_ms(key_ms_ocr, K('filtros_ocorrencias'), opts_ocr)
                st.rerun()
            if c[1].button('Desmarcar', key=f'des_ocr_{PAGE_ID}', use_container_width=True):
                _sync_ms(key_ms_ocr, K('filtros_ocorrencias'), [])
                st.rerun()

            st.session_state[K('filtros_ocorrencias')] = st.multiselect(
                ' ', options=opts_ocr,
                default=st.session_state.get(key_ms_ocr, default_ocr),
                label_visibility='hidden',
                key=key_ms_ocr
            )


    # --- Aplicação dos filtros + RESOLVIDAS ---
    anos_sel  = st.session_state.get(K('filtros_anos'), [])
    meses_sel = st.session_state.get(K('filtros_meses'), [])
    if anos_sel or meses_sel:
        df_cal = df.copy()
        if anos_sel:
            df_cal = df_cal[df_cal['Ano'].isin(anos_sel)]
        if meses_sel:
            df_cal = df_cal[df_cal['Mês'].isin(meses_sel)]
        dias_disponiveis = sorted([d for d in df_cal['Dia'].unique().tolist() if d and d > 0])
    else:
        dias_disponiveis = list(range(1, 32))

    dias_sel  = st.session_state.get(K('filtros_dias'), [])

    s_ano, s_mes, s_dia = df['Ano'], df['Mês'], df['Dia']
    m_ano = s_ano.isin(anos_sel)  if anos_sel  else s_ano.notna()
    m_mes = s_mes.isin(meses_sel) if meses_sel else s_mes.notna()
    m_dia = s_dia.isin(dias_sel)  if dias_sel  else s_dia.notna()

    m_cli = df['Cliente'].isin(st.session_state.get(K('filtros_clientes'), []))
    m_ug  = df['UG'].isin(st.session_state.get(K('filtros_ugs'), []))
    m_tip = df['Tipo de ocorrência'].isin(st.session_state.get(K('filtros_tipos'), []))
    m_atv = df['Ativo'].isin(st.session_state.get(K('filtros_ativos'), []))
    m_ocr = df['Ocorrência'].isin(st.session_state.get(K('filtros_ocorrencias'), []))

    # Filtrado + somente resolvidas (Normalização não nula)
    df_filtrado   = df[m_ano & m_mes & m_dia & m_cat & m_cli & m_ug & m_tip & m_atv & m_ocr].copy()
    df_resolvidas = df_filtrado[~df_filtrado['Normalização'].isna()].copy()

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