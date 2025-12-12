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
import utils  # Importando o utils modificado

# --- 1. Configuração da Página e Layout ---
st.set_page_config(layout="wide", page_title="Ocorrências Ativas")

# [NOVO] Injeta CSS Inteligente e Toggle de Tema (Gerenciado pelo utils.py)
# Isso corrige as cores dos retângulos de KPI no modo claro/escuro
utils.render_page_config_and_css()

# CSS Específico desta página (Apenas para os Cards de Detalhes em Vermelho)
# O CSS dos KPIs agora vem do utils, então não precisamos repeti-lo aqui.
st.markdown("""
<style>
    /* Card Container (Vermelho nesta página) */
    .card-container {
        background-color: #FF4B4B; 
        color: white; 
        padding: clamp(12px, 3vw, 16px); 
        border-radius: 8px;
        box-shadow: 0 4px 8px rgba(0,0,0,.2); 
        word-wrap: break-word;
        height: 100%;
    }
    .card-title {
        font-size: clamp(1rem, 3vw, 1.2rem); font-weight: 700;
        border-bottom: 1px solid rgba(255,255,255,.45); 
        padding-bottom: 6px; margin-bottom: 10px;
    }
    .card-item {
        font-size: clamp(.8rem, 2.1vw, .95rem); 
        line-height: 1.35; margin-bottom: 6px;
    }
    .card-label { font-weight: 700; }
</style>
""", unsafe_allow_html=True)

if "categoria_top" not in st.session_state:
    st.session_state["categoria_top"] = "Ambas"

if st.session_state["categoria_top"] not in ["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"]:
    st.session_state["categoria_top"] = "Ambas"

# Estado do overlay
if "ui_phase" not in st.session_state:
    st.session_state.ui_phase = "init"
if "loading_ts" not in st.session_state:
    st.session_state.loading_ts = 0

utils.render_loading_overlay(st.session_state.ui_phase)

def start_loading():
    st.session_state.ui_phase = "loading"
    st.session_state.loading_ts = pytime.time()

def stop_loading():
    st.session_state.ui_phase = "ready"
    st.session_state.loading_ts = 0

if st.session_state.ui_phase == "init":
    start_loading()

if st.session_state.ui_phase == "loading" and (pytime.time() - st.session_state.loading_ts) > 20:
    stop_loading()

# --- Funções Auxiliares ---
def _collapse_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()

def canon(s) -> str:
    if s is None: return ""
    s = str(s)
    s = _collapse_spaces(s)
    s = s.casefold()
    return s

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
        return pd.Series([True] * len(series), index=series.index)
    sel_c = {canon(s) for s in selected if s and s != "-"}
    return series.astype(str).map(canon).isin(sel_c)

meses_traducao = {
    "January": "Janeiro", "February": "Fevereiro", "March": "Março", "April": "Abril",
    "May": "Maio", "June": "Junho", "July": "Julho", "August": "Agosto",
    "September": "Setembro", "October": "Outubro", "November": "Novembro", "December": "Dezembro",
}
meses_cronologicos = list(meses_traducao.values())

# --- Carregar Dados ---
@st.cache_data(ttl=600)
def carregar_dados_google_sheets(cache_buster: int = 0):
    try:
        client = utils.connect_to_google_sheets()
        if not client: return pd.DataFrame()

        workbook = client.open_by_url(utils.SPREADSHEET_URL)
        df_desligamentos = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_DESLIGAMENTOS))
        df_equipamentos = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_EQUIPAMENTOS))

        if "IDENTIFICADOR" in df_desligamentos.columns:
            df_desligamentos["IDENTIFICADOR"] = df_desligamentos["IDENTIFICADOR"].astype(str)
        if "IDENTIFICADOR" in df_equipamentos.columns:
            df_equipamentos["IDENTIFICADOR"] = df_equipamentos["IDENTIFICADOR"].astype(str)

        df_desligamentos.dropna(how="all", inplace=True)
        df_equipamentos.dropna(how="all", inplace=True)

        df_desligamentos["Categoria"] = "DESLIGAMENTOS"
        df_equipamentos["Categoria"] = "EQUIPAMENTOS"
        df_todos_dados = pd.concat([df_desligamentos, df_equipamentos], ignore_index=True)

        mapa_renomear = {
            "IDENTIFICADOR": "Identificador", "CLIENTE": "Cliente", "UG": "UG",
            "TIPO DE OCORRÊNCIA": "Tipo de ocorrência", "ATIVO": "Ativo", "NOME ATIVO": "Nome Ativo",
            "OCORRÊNCIA": "Ocorrência", "QUANTIDADE": "Quantidade", "SIGLA": "Sigla",
            "NORMALIZAÇÃO": "Normalização", "DESLIGAMENTO": "Desligamento", "OPERADOR": "Operador",
            "DESCRIÇÃO": "Descrição", "OS": "OS", "ATENDIMENTO LOOP": "Atendimento Loop",
            "ATENDIMENTO TERCEIROS": "Atendimento Terceiros", "PROTOCOLO": "Protocolo",
            "CLIENTE AVISADO": "Cliente Avisado",
        }
        
        renomear_final = {}
        for col in df_todos_dados.columns:
            col_strip_upper = col.strip().upper()
            if col_strip_upper in mapa_renomear:
                renomear_final[col] = mapa_renomear[col_strip_upper]
        df_todos_dados.rename(columns=renomear_final, inplace=True)
        df_todos_dados.fillna("", inplace=True)

        if "Cliente" in df_todos_dados.columns:
            df_todos_dados = df_todos_dados[(df_todos_dados["Cliente"] != "") & (df_todos_dados["UG"] != "")].copy()

        colunas_datetime = ["Normalização", "Desligamento", "Atendimento Loop", "Atendimento Terceiros", "Cliente Avisado"]
        for col in colunas_datetime:
            if col in df_todos_dados.columns:
                df_todos_dados[col] = pd.to_datetime(df_todos_dados[col], errors="coerce")

        if "Desligamento" in df_todos_dados.columns:
            df_todos_dados["Data"] = df_todos_dados["Desligamento"].dt.strftime("%Y-%m-%d")
            df_todos_dados["Hora"] = df_todos_dados["Desligamento"].dt.strftime("%H:%M:%S")
            df_todos_dados["Mês"] = df_todos_dados["Desligamento"].dt.strftime("%B").map(meses_traducao)
            df_todos_dados["Ano"] = df_todos_dados["Desligamento"].dt.year.fillna(0).astype(int)
            df_todos_dados["Dia"] = df_todos_dados["Desligamento"].dt.day.fillna(0).astype(int)
            
            df_todos_dados["ID_Unico"] = (
                df_todos_dados["UG"].astype(str).str.upper() + "|" +
                df_todos_dados["Ativo"].astype(str).str.upper() + "|" +
                df_todos_dados["Ocorrência"].astype(str).str.upper() + "|" +
                df_todos_dados["Desligamento"].astype(str)
            )

        return df_todos_dados

    except Exception as e:
        st.error(f"Ocorreu um erro ao carregar os dados: {e}")
        return pd.DataFrame()

if "cache_buster" not in st.session_state:
    st.session_state.cache_buster = int(pytime.time())

df_todos_dados = carregar_dados_google_sheets(st.session_state.cache_buster)

if st.session_state.ui_phase != "ready":
    stop_loading()

# --- KPIs (Retângulos Superiores) ---
count_deslig = 0
count_equip = 0

if not df_todos_dados.empty:
    df_todos_dados["Desligamento"] = pd.to_datetime(df_todos_dados["Desligamento"], errors="coerce")
    count_deslig = df_todos_dados[(df_todos_dados["Categoria"] == "DESLIGAMENTOS") & (pd.isna(df_todos_dados["Normalização"]) | (df_todos_dados["Normalização"] == ""))].shape[0]
    count_equip = df_todos_dados[(df_todos_dados["Categoria"] == "EQUIPAMENTOS") & (pd.isna(df_todos_dados["Normalização"]) | (df_todos_dados["Normalização"] == ""))].shape[0]

with st.container(border=True):
    st.markdown("<h1 style='margin:0; text-align:center;'>OCORRÊNCIAS ATIVAS</h1>", unsafe_allow_html=True)
    col_top1, col_top2 = st.columns(2)
    
    # [CORRIGIDO] Usando a classe kpi-card gerada pelo utils
    with col_top1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">USINAS DESLIGADAS NO MOMENTO</div>
            <div class="kpi-value" style="color: #FF4B4B;">{count_deslig}</div>
        </div>
        """, unsafe_allow_html=True)

    with col_top2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">EQUIPAMENTOS PARADOS NO MOMENTO</div>
            <div class="kpi-value" style="color: #FF4B4B;">{count_equip}</div>
        </div>
        """, unsafe_allow_html=True)

# --- 5. Inicialização dos Filtros ---
if "filtros_meses" not in st.session_state:
    st.session_state.filtros_meses = [meses_traducao[datetime.now().strftime("%B")]]
if "filtros_anos" not in st.session_state:
    if not df_todos_dados.empty and "Ano" in df_todos_dados.columns:
        anos_atuais = sorted(df_todos_dados["Ano"].unique().tolist())
        st.session_state.filtros_anos = [a for a in anos_atuais if a != 0]
    else:
        st.session_state.filtros_anos = []
if "filtros_dias" not in st.session_state:
    if not df_todos_dados.empty and {"Mês", "Ano"}.issubset(df_todos_dados.columns):
        dias_atuais = sorted(
            df_todos_dados[
                (df_todos_dados["Mês"].isin(st.session_state.filtros_meses))
                & (df_todos_dados["Ano"].isin(st.session_state.filtros_anos))
            ]["Dia"].unique().tolist()
        )
        st.session_state.filtros_dias = [d for d in dias_atuais if d != 0]
    else:
        st.session_state.filtros_dias = []
if "filtros_categorias" not in st.session_state:
    st.session_state.filtros_categorias = sorted(df_todos_dados["Categoria"].unique().tolist()) if not df_todos_dados.empty else []
if "filtros_clientes" not in st.session_state:
    cli_series = df_todos_dados["Cliente"].astype(str).map(_collapse_spaces)
    st.session_state.filtros_clientes = sorted([v for v in cli_series.unique().tolist() if v and v != "-" and v != "0"])
if "filtros_ugs" not in st.session_state:
    st.session_state.filtros_ugs = sorted(df_todos_dados["UG"].unique().tolist()) if not df_todos_dados.empty else []
if "filtros_tipos" not in st.session_state:
    st.session_state.filtros_tipos = sorted([x for x in options_from(df_todos_dados["Tipo de ocorrência"]) if x != "-"])
if "filtros_ativos" not in st.session_state:
    st.session_state.filtros_ativos = sorted(df_todos_dados["Ativo"].unique().tolist()) if not df_todos_dados.empty else []
if "filtros_ocorrencias" not in st.session_state:
    st.session_state.filtros_ocorrencias = sorted([x for x in options_from(df_todos_dados["Ocorrência"]) if x != "-"])

# --- 6. KPIs Filtrados ---
st.markdown("---")
st.header("OCORRÊNCIAS FILTRADAS")

# Botão de Atualização
col_top_left, col_top_right = st.columns([0.2, 0.8])
with col_top_left:
    if st.button("Atualizar Dados"):
        st.cache_data.clear()
        start_loading()
        st.rerun()

# Aplicação de filtros para contagem
def _marcar(prefixo_key: str, itens: list, filtro_key: str, marcar_todos: bool, validos: set | None = None):
    validos = set(itens) if validos is None else set(validos)
    st.session_state[filtro_key] = [x for x in itens if (x in validos) and marcar_todos]
    for x in itens:
        st.session_state[f"{prefixo_key}{x}"] = marcar_todos and (x in validos)

def marcar_e_loading(prefixo_key, itens, filtro_key, marcar_todos, validos=None):
    _marcar(prefixo_key, itens, filtro_key, marcar_todos, validos)
    start_loading()

# Lógica de Filtragem (Recriando o dataframe filtrado)
meses_selecionados = [m for m in meses_cronologicos if st.session_state.get(f"cb_mes_{m}", False)]
anos_selecionados = [a for a in st.session_state.filtros_anos if st.session_state.get(f"cb_ano_{a}", False)] # Simplificado base na logica anterior
# Nota: A lógica exata de checkbox vs session_state pode variar, mantendo a consistência do original:
# Reconstruindo a máscara de filtro baseada no estado atual
s_ano = df_todos_dados["Ano"]
s_mes = df_todos_dados["Mês"]
s_dia = df_todos_dados["Dia"]

# Checkboxes de Meses/Anos controlam a lista st.session_state.filtros_X no original?
# O código original usava checkboxes para definir listas locais. Vamos usar as listas do session_state diretamente para simplificar a visualização dos KPIs filtrados.
# Mas para manter a funcionalidade completa da UI de filtros original, precisamos renderizá-la.

# --- Renderização da Interface de Filtros (Categorias, Anos, Meses, Dias) ---
if not df_todos_dados.empty:
    st.markdown("#### Filtrar por categoria (planilha)")
    st.radio(
        "Categoria:",
        options=["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"],
        horizontal=True,
        label_visibility="collapsed",
        key="categoria_top",
        on_change=start_loading,
    )

    st.subheader("Selecione o período desejado")
    anos_disponiveis = sorted([a for a in df_todos_dados["Ano"].unique() if a != 0])
    meses_disponiveis = meses_cronologicos[:]
    
    col_ano, col_mes, col_dia = st.columns(3)
    
    # Ano
    with col_ano:
        with st.container(border=True):
            st.write("### Ano(s):")
            with st.expander("Expandir anos"):
                for ano in anos_disponiveis:
                    st.checkbox(str(ano), key=f"cb_ano_{ano}", value=(ano in st.session_state.filtros_anos))
            if st.button("Sel. Todos", key="sel_ano"):
                 st.session_state.filtros_anos = anos_disponiveis
                 st.rerun()
            if st.button("Desmarcar", key="des_ano"):
                 st.session_state.filtros_anos = []
                 st.rerun()
            # Atualiza lista baseada nos checkboxes
            st.session_state.filtros_anos = [a for a in anos_disponiveis if st.session_state.get(f"cb_ano_{a}", False)]

    # Mês
    with col_mes:
        with st.container(border=True):
            st.write("### Mês(es):")
            with st.expander("Expandir meses"):
                for mes in meses_disponiveis:
                    st.checkbox(mes, key=f"cb_mes_{mes}", value=(mes in st.session_state.filtros_meses))
            if st.button("Sel. Todos", key="sel_mes"):
                 st.session_state.filtros_meses = meses_disponiveis
                 st.rerun()
            if st.button("Desmarcar", key="des_mes"):
                 st.session_state.filtros_meses = []
                 st.rerun()
            st.session_state.filtros_meses = [m for m in meses_disponiveis if st.session_state.get(f"cb_mes_{m}", False)]

    # Dia
    dias_disponiveis = list(range(1, 32))
    with col_dia:
        with st.container(border=True):
            st.write("### Dia(s):")
            with st.expander("Expandir dias"):
                dias_cols = st.columns(7)
                for i, dia in enumerate(range(1, 32)):
                    with dias_cols[i % 7]:
                        st.checkbox(str(dia), key=f"cb_dia_{dia}", value=(dia in st.session_state.filtros_dias))
            if st.button("Sel. Todos", key="sel_dia"):
                 st.session_state.filtros_dias = dias_disponiveis
                 st.rerun()
            if st.button("Desmarcar", key="des_dia"):
                 st.session_state.filtros_dias = []
                 st.rerun()
            st.session_state.filtros_dias = [d for d in dias_disponiveis if st.session_state.get(f"cb_dia_{d}", False)]

# Filtros Adicionais (Clientes, UGs, etc)
st.subheader("Filtros Adicionais")
row1_c1, row1_c2, row1_c3 = st.columns(3)
row2_c1, row2_c2 = st.columns(2)
col_cliente, col_ug, col_tipo, col_ativo, col_ocorrencia = row1_c1, row1_c2, row1_c3, row2_c1, row2_c2

# Cliente
with col_cliente:
    with st.container(border=True):
        st.write("Cliente")
        cli_series = df_todos_dados["Cliente"].astype(str).map(_collapse_spaces)
        cli_opts = sorted([v for v in cli_series.unique().tolist() if v and v != "-" and v != "0"])
        if st.button("Limpar", key="limpar_cli"): st.session_state.filtros_clientes = []
        st.session_state.filtros_clientes = st.multiselect("", options=cli_opts, default=[x for x in st.session_state.filtros_clientes if x in cli_opts])

# UG
with col_ug:
    with st.container(border=True):
        st.write("UG")
        if st.session_state.filtros_clientes:
            df_temp = df_todos_dados[df_todos_dados["Cliente"].isin(st.session_state.filtros_clientes)]
        else:
            df_temp = df_todos_dados
        ugs_series = df_temp["UG"].astype(str).map(_collapse_spaces) if "UG" in df_temp.columns else pd.Series([], dtype=str)
        ugs_disponiveis = sorted([u for u in ugs_series.unique().tolist() if u and u != "-"])
        if st.button("Limpar", key="limpar_ug"): st.session_state.filtros_ugs = []
        st.session_state.filtros_ugs = st.multiselect("", options=ugs_disponiveis, default=[x for x in st.session_state.filtros_ugs if x in ugs_disponiveis])

# Tipo
with col_tipo:
    with st.container(border=True):
        st.write("Tipo de Ocorrência")
        tip_opts = sorted([x for x in options_from(df_todos_dados["Tipo de ocorrência"]) if x != "-"])
        if st.button("Limpar", key="limpar_tipo"): st.session_state.filtros_tipos = []
        st.session_state.filtros_tipos = st.multiselect("", options=tip_opts, default=[x for x in st.session_state.filtros_tipos if x in tip_opts])

# Ativo
with col_ativo:
    with st.container(border=True):
        st.write("Ativo")
        atv_opts = sorted([x for x in options_from(df_todos_dados["Ativo"]) if x != "-"])
        if st.button("Limpar", key="limpar_atv"): st.session_state.filtros_ativos = []
        st.session_state.filtros_ativos = st.multiselect("", options=atv_opts, default=[x for x in st.session_state.filtros_ativos if x in atv_opts])

# Ocorrência
with col_ocorrencia:
    with st.container(border=True):
        st.write("Ocorrência")
        ocr_opts = sorted([x for x in options_from(df_todos_dados["Ocorrência"]) if x != "-"])
        if st.button("Limpar", key="limpar_ocr"): st.session_state.filtros_ocorrencias = []
        st.session_state.filtros_ocorrencias = st.multiselect("", options=ocr_opts, default=[x for x in st.session_state.filtros_ocorrencias if x in ocr_opts])


# Construção do DataFrame Filtrado Final
m_ano = s_ano.isin(st.session_state.filtros_anos) if st.session_state.filtros_anos else (s_ano.isin(st.session_state.filtros_anos) | s_ano.isna() | (s_ano == 0))
m_mes = s_mes.isin(st.session_state.filtros_meses) if st.session_state.filtros_meses else (s_mes.isin(st.session_state.filtros_meses) | s_mes.isna() | (s_mes == ""))
m_dia = s_dia.isin(st.session_state.filtros_dias) if st.session_state.filtros_dias else (s_dia.isin(st.session_state.filtros_dias) | s_dia.isna() | (s_dia == 0))

m_cat = df_todos_dados["Categoria"].isin(st.session_state.filtros_categorias)
if st.session_state.get("categoria_top") in ("DESLIGAMENTOS", "EQUIPAMENTOS"):
    m_cat = m_cat & (df_todos_dados["Categoria"] == st.session_state["categoria_top"])

m_cli = matches_any_canon(df_todos_dados["Cliente"], st.session_state.filtros_clientes)
m_ug = df_todos_dados["UG"].astype(str).map(_collapse_spaces).isin(st.session_state.filtros_ugs) if st.session_state.filtros_ugs else pd.Series(True, index=df_todos_dados.index)
m_tip = matches_any_canon(df_todos_dados["Tipo de ocorrência"], st.session_state.filtros_tipos)
m_atv = matches_any_canon(df_todos_dados["Ativo"], st.session_state.filtros_ativos)
m_ocr = matches_any_canon(df_todos_dados["Ocorrência"], st.session_state.filtros_ocorrencias)

df_filtrado = df_todos_dados[m_ano & m_mes & m_dia & m_cat & m_cli & m_ug & m_tip & m_atv & m_ocr].copy()

df_desligadas = df_filtrado[pd.isna(df_filtrado["Normalização"]) | (df_filtrado["Normalização"] == "")].copy()

total_filtrado = len(df_desligadas)
total_banco = df_todos_dados[pd.isna(df_todos_dados["Normalização"]) | (df_todos_dados["Normalização"] == "")].shape[0]

# KPIs Filtrados (Correção das Cores)
col_kpi1, col_kpi2 = st.columns(2)
with col_kpi1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total no Banco de Dados Completo</div>
        <div class="kpi-value" style="color: #FF4B4B;">{total_banco}</div>
    </div>
    """, unsafe_allow_html=True)

with col_kpi2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total com Filtro Selecionado</div>
        <div class="kpi-value" style="color: #FF4B4B;">{total_filtrado}</div>
    </div>
    """, unsafe_allow_html=True)

if not df_desligadas.empty:
    maskvalid = df_desligadas["Desligamento"].notna()
    df_desligadas.loc[maskvalid, "Tempo em Segundos"] = (datetime.now() - df_desligadas.loc[maskvalid, "Desligamento"]).dt.total_seconds().astype(int)
    df_desligadas.loc[~maskvalid, "Tempo em Segundos"] = 0

    st.markdown("---")
    st.write("Ordenar e Editar")

    sortcols = st.columns(2)
    with sortcols[0]:
        sortoptionsdisplay = {
            "Data do Desligamento": "Desligamento",
            "Tempo de Desligamento": "Tempo em Segundos",
            "UG": "UG",
            "Ativo": "Ativo",
        }
        sortbydisplay = st.selectbox("Ordenar por", options=list(sortoptionsdisplay.keys()), index=0)
        sortbycolumn = sortoptionsdisplay[sortbydisplay]

    with sortcols[1]:
        sortorder = st.radio("Ordem", options=["Descendente", "Ascendente"], index=0, horizontal=True)
        isascending = sortorder == "Ascendente"

    dfsorted = df_desligadas.sort_values(by=sortbycolumn, ascending=isascending, na_position="last")
    
    # Cria coluna Display para edição
    dfsorted["Display"] = (
        dfsorted["UG"].astype(str) + " | " + dfsorted["Ativo"].astype(str) + " | " + 
        dfsorted["Nome Ativo"].astype(str) + " | " + dfsorted["Ocorrência"].astype(str) + " | " + 
        dfsorted["Desligamento"].dt.strftime("%d/%m/%Y %H:%M").fillna("") + " | " + 
        dfsorted["ID_Unico"].astype(str).str[-6:]
    )
    
    # Botão de Editar
    st.markdown("---")
    st.write("Editar uma Ocorrência")
    opts = dfsorted["Display"].dropna().astype(str).tolist()
    ocorrenciaselecionadadisplay = st.selectbox("Selecione a ocorrência para editar", options=opts, index=None, placeholder="Escolha uma ocorrência...")

    if ocorrenciaselecionadadisplay:
        idunicoparaeditar = dfsorted.loc[dfsorted["Display"] == ocorrenciaselecionadadisplay, "ID_Unico"].head(1).item()
        st.session_state["id_unico_para_editar"] = idunicoparaeditar
        st.session_state["df_lista_para_editar"] = dfsorted.copy() # Salva lista para pagina de edição

    if (not ocorrenciaselecionadadisplay) and "id_unico_para_editar" in st.session_state:
        st.session_state.pop("id_unico_para_editar")

    btndisabled = not bool(ocorrenciaselecionadadisplay)
    if st.button("Editar Ocorrência Selecionada", disabled=btndisabled, use_container_width=True):
        st.switch_page("pages/3_Editar_Ocorrência.py")

    # Tabela
    st.header("Lista de Ocorrências (Tabela)")
    dfparatabela = dfsorted.copy()
    def formatartempoestatico(row):
        dias = row["Tempo em Segundos"] // 86400
        horas = (row["Tempo em Segundos"] % 86400) // 3600
        minutos = (row["Tempo em Segundos"] % 3600) // 60
        return f"{dias}d {horas}h {minutos}m"

    dfparatabela["Tempo de Desligamento"] = dfparatabela.apply(formatartempoestatico, axis=1)
    dfparatabela.reset_index(inplace=True, drop=True)
    dfparatabela["Linha"] = dfparatabela.index + 1

    st.dataframe(dfparatabela[["Linha", "Categoria", "Tempo de Desligamento", "UG", "Data", "Hora", "Tipo de ocorrência", "Ativo", "Ocorrência", "Operador", "Descrição", "OS"]], use_container_width=True)

    # Cards Detalhados
    st.header("Detalhes por Ocorrência (Cards)")
    num_cols = 3
    rows = list(dfsorted.iterrows())
    
    def formatdatetimecard(dtobj):
        if pd.notna(dtobj): return dtobj.strftime("%d/%m/%Y"), dtobj.strftime("%H:%M")
        return "", ""

    for i in range(0, len(rows), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j >= len(rows): continue
            index, row = rows[i + j]
            with cols[j]:
                # Extração de dados
                cliente = html.escape(str(row.get("Cliente", "")))
                categoria = html.escape(str(row.get("Categoria", "")))
                ug = html.escape(str(row.get("UG", "N/A")))
                tipoocorrencia = html.escape(str(row.get("Tipo de ocorrência", "")))
                ativo = html.escape(str(row.get("Ativo", "")))
                nomeativo = html.escape(str(row.get("Nome Ativo", "")))
                ocorrencia = html.escape(str(row.get("Ocorrência", "")))
                operador = html.escape(str(row.get("Operador", "")))
                descricao = html.escape(str(row.get("Descrição", ""))).replace("\\n", "<br>")
                protocolo = html.escape(str(row.get("Protocolo", "")))
                os = html.escape(str(row.get("OS", "")))

                dataocor, horaocor = formatdatetimecard(row.get("Desligamento"))
                dataca, horaca = formatdatetimecard(row.get("Cliente Avisado"))
                dataloop, horaloop = formatdatetimecard(row.get("Atendimento Loop"))
                dataterc, horaterc = formatdatetimecard(row.get("Atendimento Terceiros"))
                datanorm, horanorm = formatdatetimecard(row.get("Normalização"))

                quantidadehtml = ""
                if row.get("Categoria") == "EQUIPAMENTOS":
                    qv = row.get("Quantidade", 0)
                    try:
                        if pd.notna(qv) and float(qv) > 0:
                            quantidadehtml = f'<div class="card-item"><span class="card-label">Quantidade:</span> {int(float(qv))}</div>'
                    except: pass

                # Card HTML (Vermelho, estilo definido no inicio do arquivo)
                cardhtml = f"""
                <div class="card-container">
                    <div class="card-title">{ug}</div>
                    <div class="card-item"><span class="card-label">Cliente:</span> {cliente}</div>
                    <div class="card-item"><span class="card-label">Categoria:</span> {categoria}</div>
                    <div class="card-item"><span class="card-label">Tipo:</span> {tipoocorrencia}</div>
                    <div class="card-item"><span class="card-label">Ativo:</span> {ativo} | {nomeativo}</div>
                    <div class="card-item"><span class="card-label">Ocorrência:</span> {ocorrencia}</div>
                    <div class="card-item"><span class="card-label">Operador:</span> {operador}</div>
                    {quantidadehtml}
                    <br>
                    <div class="card-item"><span class="card-label">Data Ocorrência:</span> {dataocor} {horaocor}</div>
                    <div class="card-item"><span class="card-label">Cliente Avisado:</span> {dataca} {horaca}</div>
                    <div class="card-item"><span class="card-label">Loop:</span> {dataloop} {horaloop}</div>
                    <div class="card-item"><span class="card-label">Terceiros:</span> {dataterc} {horaterc}</div>
                    <br>
                    <div class="card-item"><span class="card-label">Descrição:</span> {descricao}</div>
                    <div class="card-item"><span class="card-label">Protocolo:</span> {protocolo}</div>
                    <div class="card-item"><span class="card-label">OS:</span> {os}</div>
                </div>
                """
                st.html(cardhtml)
else:
    st.info("Nenhuma usina encontrada com o campo Normalização em branco para os filtros selecionados.")