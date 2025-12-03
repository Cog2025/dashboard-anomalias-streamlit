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
import utils  # [MODIFICADO] Importando utils

# --- 1. Configuração da Página e Layout ---
st.set_page_config(layout="wide")

if "categoria_top" not in st.session_state:
    st.session_state["categoria_top"] = "Ambas"

if st.session_state["categoria_top"] not in ["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"]:
    st.session_state["categoria_top"] = "Ambas"

# Estado do overlay
if "ui_phase" not in st.session_state:
    st.session_state.ui_phase = "init"
if "loading_ts" not in st.session_state:
    st.session_state.loading_ts = 0

# [MODIFICADO] Usando utils para renderizar
utils.render_loading_overlay(st.session_state.ui_phase)


def start_loading():
    st.session_state.ui_phase = "loading"
    st.session_state.loading_ts = pytime.time()
    # st.rerun()


def stop_loading():
    st.session_state.ui_phase = "ready"
    st.session_state.loading_ts = 0
    # st.rerun()


# [MODIFICADO] Usando utils
utils.render_loading_overlay()
if st.session_state.ui_phase == "init":
    start_loading()

# Failsafe opcional (20s)
if st.session_state.ui_phase == "loading" and (
    pytime.time() - st.session_state.loading_ts
) > 20:
    stop_loading()


def _collapse_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def canon(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = _collapse_spaces(s)
    s = s.casefold()
    return s


def build_display_map(series: pd.Series) -> dict:
    buckets = defaultdict(Counter)
    for v in series.dropna():
        v_str = _collapse_spaces(str(v))
        if not v_str:
            continue
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


def matches_canon(series: pd.Series, selected: str) -> pd.Series:
    if not selected or selected == "-":
        return pd.Series([True] * len(series))
    sel_c = canon(selected)
    return series.astype(str).map(canon).eq(sel_c)


meses_traducao = {
    "January": "Janeiro",
    "February": "Fevereiro",
    "March": "Março",
    "April": "Abril",
    "May": "Maio",
    "June": "Junho",
    "July": "Julho",
    "August": "Agosto",
    "September": "Setembro",
    "October": "Outubro",
    "November": "Novembro",
    "December": "Dezembro",
}
meses_cronologicos = list(meses_traducao.values())

# --- 3. CSS ---
st.markdown(
    """
<style>
/* Usar toda a largura disponível do app */
.main .block-container{
  max-width: 100% !important;
  padding-left: 1rem !important;
  padding-right: 1rem !important;
}

/* Layout: linhas horizontais de colunas flexíveis */
[data-testid="stHorizontalBlock"]{
  display:flex !important;
  flex-wrap:wrap !important;
  gap:12px !important;
}

/* Colunas base */
[data-testid="column"]{
  flex:1 1 320px !important;
  min-width:280px !important;
}

/* Wrapper dos botões */
.stButton{
  width:100%;
}

/* Botões padrão */
.stButton button{
  background-color:#28a745; color:#fff; border:none; border-radius:6px;
  box-shadow:0 2px 4px rgba(0,0,0,.2); transition:.2s ease;
  width:100%;
  min-height:42px; padding:8px 12px;
  font-size:clamp(12px, 2.3vw, 14px); font-weight:500;
  white-space:normal !important; word-break:normal !important; overflow-wrap:anywhere !important;
  margin-bottom:6px !important;
}

/* Select/Multiselect */
.stSelectbox > div > div,
.stMultiSelect div[data-baseweb="select"]{
  max-height: 90px !important;
  overflow-y: auto !important;
  white-space: normal !important;
  line-height: 1.2 !important;
  font-size: clamp(12px, 1.4vw, 15px) !important;
}
.stMultiSelect [data-baseweb="tag"]{
  font-size: 11px !important;
  padding: 2px 6px !important;
}
div[data-baseweb="select"]{ min-height:42px !important; }
div[data-baseweb="select"] div{
  white-space:normal !important; overflow:visible !important; text-overflow:clip !important;
}

/* KPIs e cards */
.kpi-card{
  background:#333; padding:clamp(12px, 3vw, 20px); border-radius:10px; text-align:center;
}
.kpi-value{
  font-size:clamp(1.6rem, 6vw, 3rem); font-weight:700; color:#FF4B4B;
}
.kpi-label{
  font-size:clamp(.85rem, 2.5vw, 1rem); color:#fff;
}
.card-container{
  background:#FF4B4B; color:#fff; padding:clamp(12px, 3vw, 16px); border-radius:8px;
  box-shadow:0 4px 8px rgba(0,0,0,.2); word-wrap:break-word;
}
.card-title{
  font-size:clamp(1rem, 3vw, 1.2rem); font-weight:700;
  border-bottom:1px solid rgba(255,255,255,.45); padding-bottom:6px; margin-bottom:10px;
}
.card-item{
  font-size:clamp(.8rem, 2.1vw, .95rem); line-height:1.35; margin-bottom:6px;
}
.card-label{
  font-weight:700;
}

/* Títulos */
h1{
  font-size:clamp(1.6rem, 5vw, 2.6rem) !important; text-align:center;
}

/* Em telas médias, botões um pouco mais compactos */
@media (max-width:1200px){
  .stButton button{
    padding:6px 8px !important;
    font-size:12px !important;
  }
}

/* Em meia tela ou menor, empilhar colunas que contenham botões
   (Sel. Todos / Desmarcar ficam um embaixo do outro) */
@media (max-width:1050px){
  /* Cada coluna que tiver stButton passa a ocupar 100% da linha */
  [data-testid="stHorizontalBlock"] > div:has(.stButton){
    flex-basis:100% !important;
    min-width:100% !important;
  }
}

</style>
""",
    unsafe_allow_html=True,
)



st.markdown(
    """
<style>
  .boxed { padding: 12px; border-radius: 8px; }
  .boxed h4 { margin-top: 0; }
</style>
""",
    unsafe_allow_html=True,
)

# --- 4. Carregar e Tratar os Dados ---


@st.cache_data(ttl=600)
def carregar_dados_google_sheets(cache_buster: int = 0):
    try:
        client = utils.connect_to_google_sheets()
        if not client:
            return pd.DataFrame()

        spreadsheet_url = utils.SPREADSHEET_URL
        workbook = client.open_by_url(spreadsheet_url)

        df_desligamentos = utils.fetch_sheet_as_df(
            workbook.worksheet(utils.SHEET_DESLIGAMENTOS)
        )
        df_equipamentos = utils.fetch_sheet_as_df(
            workbook.worksheet(utils.SHEET_EQUIPAMENTOS)
        )

        if "IDENTIFICADOR" in df_desligamentos.columns:
            df_desligamentos["IDENTIFICADOR"] = df_desligamentos["IDENTIFICADOR"].astype(
                str
            )
        if "IDENTIFICADOR" in df_equipamentos.columns:
            df_equipamentos["IDENTIFICADOR"] = df_equipamentos["IDENTIFICADOR"].astype(
                str
            )

        df_desligamentos.dropna(how="all", inplace=True)
        df_equipamentos.dropna(how="all", inplace=True)

        df_desligamentos["Categoria"] = "DESLIGAMENTOS"
        df_equipamentos["Categoria"] = "EQUIPAMENTOS"
        df_todos_dados = pd.concat(
            [df_desligamentos, df_equipamentos], ignore_index=True
        )

        mapa_renomear = {
            "IDENTIFICADOR": "Identificador",
            "CLIENTE": "Cliente",
            "UG": "UG",
            "TIPO DE OCORRÊNCIA": "Tipo de ocorrência",
            "ATIVO": "Ativo",
            "NOME ATIVO": "Nome Ativo",
            "OCORRÊNCIA": "Ocorrência",
            "QUANTIDADE": "Quantidade",
            "SIGLA": "Sigla",
            "NORMALIZAÇÃO": "Normalização",
            "DESLIGAMENTO": "Desligamento",
            "OPERADOR": "Operador",
            "DESCRIÇÃO": "Descrição",
            "OS": "OS",
            "ATENDIMENTO LOOP": "Atendimento Loop",
            "ATENDIMENTO TERCEIROS": "Atendimento Terceiros",
            "PROTOCOLO": "Protocolo",
            "CLIENTE AVISADO": "Cliente Avisado",
        }
        colunas_atuais = df_todos_dados.columns
        renomear_final = {}
        for col in colunas_atuais:
            col_strip_upper = col.strip().upper()
            if col_strip_upper in mapa_renomear:
                renomear_final[col] = mapa_renomear[col_strip_upper]
        df_todos_dados.rename(columns=renomear_final, inplace=True)

        df_todos_dados.fillna("", inplace=True)

        if "Cliente" in df_todos_dados.columns:
            df_todos_dados = df_todos_dados[
                (df_todos_dados["Cliente"] != "")
                & (df_todos_dados["UG"] != "")
                & (df_todos_dados["Sigla"] != "")
            ].copy()

        colunas_datetime = [
            "Normalização",
            "Desligamento",
            "Atendimento Loop",
            "Atendimento Terceiros",
            "Cliente Avisado",
        ]
        for col in colunas_datetime:
            if col in df_todos_dados.columns:
                df_todos_dados[col] = pd.to_datetime(
                    df_todos_dados[col], errors="coerce"
                )

        colunas_texto = ["Operador", "Descrição", "OS", "Protocolo"]
        for col in colunas_texto:
            if col in df_todos_dados.columns:
                df_todos_dados[col] = (
                    df_todos_dados[col].astype(str).fillna("")
                )

        if "Desligamento" in df_todos_dados.columns and not df_todos_dados[
            "Desligamento"
        ].isnull().all():
            df_todos_dados["Data"] = df_todos_dados["Desligamento"].dt.strftime(
                "%Y-%m-%d"
            )
            df_todos_dados["Hora"] = df_todos_dados["Desligamento"].dt.strftime(
                "%H:%M:%S"
            )
            df_todos_dados["Mês"] = (
                df_todos_dados["Desligamento"]
                .dt.strftime("%B")
                .map(meses_traducao)
            )
            df_todos_dados["Ano"] = (
                df_todos_dados["Desligamento"].dt.year.fillna(0).astype(int)
            )
            df_todos_dados["Dia"] = (
                df_todos_dados["Desligamento"].dt.day.fillna(0).astype(int)
            )

            df_todos_dados["ID_Unico"] = (
                df_todos_dados["UG"].astype(str).str.upper()
                + "|"
                + df_todos_dados["Ativo"].astype(str).str.upper()
                + "|"
                + df_todos_dados["Ocorrência"].astype(str).str.upper()
                + "|"
                + df_todos_dados["Desligamento"].astype(str)
            )
        else:
            for col in ["Data", "Hora", "Mês", "Ano", "Dia", "ID_Unico"]:
                df_todos_dados[col] = None

        return df_todos_dados

    except FileNotFoundError:
        st.error("Erro: O arquivo de credenciais não foi encontrado.")
        return pd.DataFrame()
    except gspread.exceptions.SpreadsheetNotFound:
        st.error("Erro: Planilha não encontrada.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Ocorreu um erro ao carregar os dados: {e}")
        return pd.DataFrame()


if "cache_buster" not in st.session_state:
    st.session_state.cache_buster = int(pytime.time())

df_todos_dados = carregar_dados_google_sheets(st.session_state.cache_buster)

if st.session_state.ui_phase != "ready":
    st.session_state.ui_phase = "ready"
    st.session_state.loading_ts = 0
    utils.render_loading_overlay("ready")

df_todos_dados["Desligamento"] = pd.to_datetime(
    df_todos_dados["Desligamento"], errors="coerce"
)

count_deslig = df_todos_dados[
    (df_todos_dados["Categoria"] == "DESLIGAMENTOS")
    & (
        pd.isna(df_todos_dados["Normalização"])
        | (df_todos_dados["Normalização"] == "")
    )
].shape[0]

count_equip = df_todos_dados[
    (df_todos_dados["Categoria"] == "EQUIPAMENTOS")
    & (
        pd.isna(df_todos_dados["Normalização"])
        | (df_todos_dados["Normalização"] == "")
    )
].shape[0]

with st.container(border=True):
    st.markdown(
        "<h1 style='margin:0'>OCORRÊNCIAS ATIVAS</h1>", unsafe_allow_html=True
    )

    col_top1, col_top2 = st.columns(2)
    with col_top1:
        st.markdown(
            f"""
        <div class="kpi-card">
            <div class="kpi-label">USINAS DESLIGADAS NO MOMENTO</div>
            <div class="kpi-value">{count_deslig}</div>
        </div>
        """,
            unsafe_allow_html=True,
        )

    with col_top2:
        st.markdown(
            f"""
        <div class="kpi-card">
            <div class="kpi-label">EQUIPAMENTOS PARADOS NO MOMENTO</div>
            <div class="kpi-value">{count_equip}</div>
        </div>
        """,
            unsafe_allow_html=True,
        )

# --- 5. Inicialização dos Filtros ---
if "filtros_meses" not in st.session_state:
    st.session_state.filtros_meses = [
        meses_traducao[datetime.now().strftime("%B")]
    ]
if "filtros_anos" not in st.session_state:
    if not df_todos_dados.empty and "Ano" in df_todos_dados.columns:
        anos_atuais = sorted(df_todos_dados["Ano"].unique().tolist())
        st.session_state.filtros_anos = [a for a in anos_atuais if a != 0]
    else:
        st.session_state.filtros_anos = []
if "filtros_dias" not in st.session_state:
    if not df_todos_dados.empty and {"Mês", "Ano"}.issubset(
        df_todos_dados.columns
    ):
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
    st.session_state.filtros_categorias = (
        sorted(df_todos_dados["Categoria"].unique().tolist())
        if not df_todos_dados.empty
        else []
    )
if "filtros_clientes" not in st.session_state:
    cli_series = df_todos_dados["Cliente"].astype(str).map(_collapse_spaces)
    st.session_state.filtros_clientes = sorted(
        [v for v in cli_series.unique().tolist() if v and v != "-" and v != "0"]
    )
if "filtros_ugs" not in st.session_state:
    st.session_state.filtros_ugs = (
        sorted(df_todos_dados["UG"].unique().tolist())
        if not df_todos_dados.empty
        else []
    )
if "filtros_tipos" not in st.session_state:
    tip_opts_init = sorted(
        [x for x in options_from(df_todos_dados["Tipo de ocorrência"]) if x != "-"]
    )
    st.session_state.filtros_tipos = tip_opts_init[:]
if "filtros_ativos" not in st.session_state:
    st.session_state.filtros_ativos = (
        sorted(df_todos_dados["Ativo"].unique().tolist())
        if not df_todos_dados.empty
        else []
    )
if "filtros_ocorrencias" not in st.session_state:
    ocr_opts_init = sorted(
        [x for x in options_from(df_todos_dados["Ocorrência"]) if x != "-"]
    )
    st.session_state.filtros_ocorrencias = ocr_opts_init[:]

# --- 6. Título e KPIs ---
st.header("OCORRÊNCIAS FILTRADAS")
col_kpi1, col_kpi2 = st.columns(2)
with col_kpi1:
    if not df_todos_dados.empty and "Normalização" in df_todos_dados.columns:
        df_desligadas_geral = df_todos_dados[
            pd.isna(df_todos_dados["Normalização"])
            | (df_todos_dados["Normalização"] == "")
        ].copy()
        total_kpi_value = df_desligadas_geral.shape[0]
    else:
        total_kpi_value = 0
    st.markdown(
        f"""
    <div class="kpi-card">
        <div class="kpi-label">Total no Banco de Dados Completo</div>
        <div class="kpi-value">{total_kpi_value}</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

# --- 7. Botão de Atualização ---
col_top_left, col_top_right = st.columns([0.2, 0.8])
with col_top_left:
    if st.button("Atualizar Dados"):
        st.cache_data.clear()
        st.session_state.ui_phase = "init"
        st.session_state.loading_ts = 0
        st.rerun()

# --- 8. Interface de Filtros ---


def _marcar(
    prefixo_key: str,
    itens: list,
    filtro_key: str,
    marcar_todos: bool,
    validos: set | None = None,
):
    validos = set(itens) if validos is None else set(validos)
    st.session_state[filtro_key] = [
        x for x in itens if (x in validos) and marcar_todos
    ]
    for x in itens:
        st.session_state[f"{prefixo_key}{x}"] = marcar_todos and (x in validos)


def matches_any_canon(series: pd.Series, selected: list[str]) -> pd.Series:
    if not selected:
        return pd.Series([True] * len(series), index=series.index)
    sel_c = {canon(s) for s in selected if s and s != "-"}
    return series.astype(str).map(canon).isin(sel_c)


def marcar_e_loading(prefixo_key, itens, filtro_key, marcar_todos, validos=None):
    _marcar(prefixo_key, itens, filtro_key, marcar_todos, validos)
    start_loading()


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
    anos_disponiveis = sorted(
        [a for a in df_todos_dados["Ano"].unique() if a != 0]
    )
    meses_disponiveis = meses_cronologicos[:]

    col_ano, col_mes, col_dia = st.columns(3)

    # --- Ano(s) ---
    with col_ano:
        with st.container(border=True):
            st.write("### Ano(s):")
            with st.expander("Expandir anos"):
                for ano in anos_disponiveis:
                    st.checkbox(
                        str(ano),
                        key=f"cb_ano_{ano}",
                        value=(ano in st.session_state.filtros_anos),
                    )

            btn_ano_col1, btn_ano_col2 = st.columns(2)
            with btn_ano_col1:
                clicked_sel_ano = st.button(
                    "Sel. Todos",
                    key="sel_ano",
                    use_container_width=True,
                    on_click=_marcar,
                    args=("cb_ano_", anos_disponiveis, "filtros_anos", True),
                )
            with btn_ano_col2:
                clicked_des_ano = st.button(
                    "Desmarcar",
                    key="des_ano",
                    use_container_width=True,
                    on_click=_marcar,
                    args=("cb_ano_", anos_disponiveis, "filtros_anos", False),
                )
            if not (clicked_sel_ano or clicked_des_ano):
                st.session_state.filtros_anos = [
                    a
                    for a in anos_disponiveis
                    if st.session_state.get(f"cb_ano_{a}", False)
                ]

    # --- Mês(es) ---
    with col_mes:
        with st.container(border=True):
            st.write("### Mês(es):")
            with st.expander("Expandir meses"):
                for mes in meses_disponiveis:
                    st.checkbox(
                        mes,
                        key=f"cb_mes_{mes}",
                        value=(mes in st.session_state.filtros_meses),
                    )

            btn_mes_col1, btn_mes_col2 = st.columns(2)
            with btn_mes_col1:
                clicked_sel_mes = st.button(
                    "Sel. Todos",
                    key="sel_mes",
                    use_container_width=True,
                    on_click=marcar_e_loading,
                    args=("cb_mes_", meses_disponiveis, "filtros_meses", True),
                )
            with btn_mes_col2:
                clicked_des_mes = st.button(
                    "Desmarcar",
                    key="des_mes",
                    use_container_width=True,
                    on_click=marcar_e_loading,
                    args=("cb_mes_", meses_disponiveis, "filtros_meses", False),
                )
            if not (clicked_sel_mes or clicked_des_mes):
                st.session_state.filtros_meses = [
                    m
                    for m in meses_disponiveis
                    if st.session_state.get(f"cb_mes_{m}", False)
                ]

    # --- Dia(s) ---
    if "dias_disponiveis" not in locals():
        dias_disponiveis = list(range(1, 32))
    with col_dia:
        with st.container(border=True):
            st.write("### Dia(s):")
            with st.expander("Expandir dias"):
                dias_cols = st.columns(7)
                for i, dia in enumerate(range(1, 32)):
                    with dias_cols[i % 7]:
                        if dia in dias_disponiveis:
                            st.checkbox(
                                str(dia),
                                key=f"cb_dia_{dia}",
                                value=(dia in st.session_state.filtros_dias),
                            )
                        else:
                            st.checkbox(
                                str(dia),
                                key=f"cb_dia_{dia}",
                                disabled=True,
                            )

            btn_dia_col1, btn_dia_col2 = st.columns(2)
            with btn_dia_col1:
                clicked_sel_dia = st.button(
                    "Sel. Todos",
                    key="sel_dia",
                    use_container_width=True,
                    on_click=_marcar,
                    args=(
                        "cb_dia_",
                        list(range(1, 32)),
                        "filtros_dias",
                        True,
                        set(dias_disponiveis),
                    ),
                )
            with btn_dia_col2:
                clicked_des_dia = st.button(
                    "Desmarcar",
                    key="des_dia",
                    use_container_width=True,
                    on_click=_marcar,
                    args=(
                        "cb_dia_",
                        list(range(1, 32)),
                        "filtros_dias",
                        False,
                        set(dias_disponiveis),
                    ),
                )
            if not (clicked_sel_dia or clicked_des_dia):
                st.session_state.filtros_dias = [
                    d
                    for d in dias_disponiveis
                    if st.session_state.get(f"cb_dia_{d}", False)
                ]

# --- Filtros Adicionais ---
st.subheader("Filtros Adicionais")

row1_c1, row1_c2, row1_c3 = st.columns(3)
row2_c1, row2_c2 = st.columns(2)
col_cliente, col_ug, col_tipo, col_ativo, col_ocorrencia = (
    row1_c1,
    row1_c2,
    row1_c3,
    row2_c1,
    row2_c2,
)

# --- Cliente ---
with col_cliente:
    with st.container(border=True):
        st.write("Cliente")
        cli_series = df_todos_dados["Cliente"].astype(str).map(_collapse_spaces)
        cli_opts = sorted(
            [v for v in cli_series.unique().tolist() if v and v != "-" and v != "0"]
        )

        btn_cli1, btn_cli2 = st.columns(2)
        with btn_cli1:
            if st.button("Sel. Todos", key="sel_cli", use_container_width=True):
                st.session_state.filtros_clientes = cli_opts
                st.rerun()
        with btn_cli2:
            if st.button("Desmarcar", key="des_cli", use_container_width=True):
                st.session_state.filtros_clientes = []
                st.rerun()

        st.session_state.filtros_clientes = st.multiselect(
            "",
            options=cli_opts,
            default=[x for x in st.session_state.filtros_clientes if x in cli_opts],
            label_visibility="hidden",
        )

# --- UG ---
with col_ug:
    with st.container(border=True):
        st.write("UG")

        if "Cliente" in df_todos_dados.columns and st.session_state.filtros_clientes:
            df_temp = df_todos_dados[
                df_todos_dados["Cliente"].isin(st.session_state.filtros_clientes)
            ]
        else:
            df_temp = df_todos_dados

        ugs_series = (
            df_temp["UG"].astype(str).map(_collapse_spaces)
            if "UG" in df_temp.columns
            else pd.Series([], dtype=str)
        )
        ugs_disponiveis = sorted(
            [u for u in ugs_series.unique().tolist() if u and u != "-"]
        )

        st.session_state.filtros_ugs = [
            ug for ug in st.session_state.filtros_ugs if ug in ugs_disponiveis
        ]

        btn_ug1, btn_ug2 = st.columns(2)
        with btn_ug1:
            if st.button("Sel. Todos", key="sel_ug", use_container_width=True):
                st.session_state.filtros_ugs = ugs_disponiveis
                st.rerun()
        with btn_ug2:
            if st.button("Desmarcar", key="des_ug", use_container_width=True):
                st.session_state.filtros_ugs = []
                st.rerun()

        st.session_state.filtros_ugs = st.multiselect(
            "",
            options=ugs_disponiveis,
            default=st.session_state.filtros_ugs,
            label_visibility="hidden",
        )

# --- Tipo de Ocorrência ---
with col_tipo:
    with st.container(border=True):
        st.write("Tipo de Ocorrência")
        tip_opts = sorted(
            [x for x in options_from(df_todos_dados["Tipo de ocorrência"]) if x != "-"]
        )

        btn_tipo1, btn_tipo2 = st.columns(2)
        with btn_tipo1:
            if st.button("Sel. Todos", key="sel_tipo", use_container_width=True):
                st.session_state.filtros_tipos = tip_opts
                st.rerun()
        with btn_tipo2:
            if st.button("Desmarcar", key="des_tipo", use_container_width=True):
                st.session_state.filtros_tipos = []
                st.rerun()

        st.session_state.filtros_tipos = [
            x for x in st.session_state.filtros_tipos if x in tip_opts
        ]
        st.session_state.filtros_tipos = st.multiselect(
            "",
            options=tip_opts,
            default=st.session_state.filtros_tipos,
            label_visibility="hidden",
        )

# --- Ativo ---
with col_ativo:
    with st.container(border=True):
        st.write("Ativo")
        atv_opts = sorted(
            [x for x in options_from(df_todos_dados["Ativo"]) if x != "-"]
        )

        btn_atv1, btn_atv2 = st.columns(2)
        with btn_atv1:
            if st.button("Sel. Todos", key="sel_ativo", use_container_width=True):
                st.session_state.filtros_ativos = atv_opts
                st.rerun()
        with btn_atv2:
            if st.button("Desmarcar", key="des_ativo", use_container_width=True):
                st.session_state.filtros_ativos = []
                st.rerun()

        st.session_state.filtros_ativos = [
            x for x in st.session_state.filtros_ativos if x in atv_opts
        ]
        st.session_state.filtros_ativos = st.multiselect(
            "",
            options=atv_opts,
            default=st.session_state.filtros_ativos,
            label_visibility="hidden",
        )

# --- Ocorrência ---
with col_ocorrencia:
    with st.container(border=True):
        st.write("Ocorrência")
        ocr_opts = sorted(
            [x for x in options_from(df_todos_dados["Ocorrência"]) if x != "-"]
        )

        btn_ocr1, btn_ocr2 = st.columns(2)
        with btn_ocr1:
            if st.button("Sel. Todos", key="sel_ocorr", use_container_width=True):
                st.session_state.filtros_ocorrencias = ocr_opts
                st.rerun()
        with btn_ocr2:
            if st.button("Desmarcar", key="des_ocorr", use_container_width=True):
                st.session_state.filtros_ocorrencias = []
                st.rerun()

        st.session_state.filtros_ocorrencias = [
            x for x in st.session_state.filtros_ocorrencias if x in ocr_opts
        ]
        st.session_state.filtros_ocorrencias = st.multiselect(
            "",
            options=ocr_opts,
            default=st.session_state.filtros_ocorrencias,
            label_visibility="hidden",
        )

# --- Aplicação dos Filtros ---
meses_selecionados = [
    mes for mes in meses_cronologicos if st.session_state.get(f"cb_mes_{mes}", False)
]
anos_selecionados = [
    ano for ano in anos_disponiveis if st.session_state.get(f"cb_ano_{ano}", False)
]
dias_selecionados = [
    dia for dia in range(1, 32) if st.session_state.get(f"cb_dia_{dia}", False)
]

set_anos_disp = set(anos_disponiveis)
set_meses_disp = set(meses_cronologicos)
set_dias_disp = (
    set(dias_disponiveis) if "dias_disponiveis" in locals() else set(range(1, 32))
)

all_anos = set(anos_selecionados) == set_anos_disp and len(set_anos_disp) > 0
all_meses = set(meses_selecionados) == set_meses_disp and len(set_meses_disp) > 0
all_dias = set(dias_selecionados) == set_dias_disp and len(set_dias_disp) > 0

s_ano = df_todos_dados["Ano"]
s_mes = df_todos_dados["Mês"]
s_dia = df_todos_dados["Dia"]

m_ano = (
    s_ano.isin(anos_selecionados)
    if not all_anos
    else (s_ano.isin(anos_selecionados) | s_ano.isna() | (s_ano == 0))
)
m_mes = (
    s_mes.isin(meses_selecionados)
    if not all_meses
    else (s_mes.isin(meses_selecionados) | s_mes.isna() | (s_mes.astype(str) == ""))
)
m_dia = (
    s_dia.isin(dias_selecionados)
    if not all_dias
    else (s_dia.isin(dias_selecionados) | s_dia.isna() | (s_dia == 0))
)

m_cat = df_todos_dados["Categoria"].isin(st.session_state.filtros_categorias)
if st.session_state.get("categoria_top") in ("DESLIGAMENTOS", "EQUIPAMENTOS"):
    m_cat = m_cat & (
        df_todos_dados["Categoria"] == st.session_state["categoria_top"]
    )

m_cli = matches_any_canon(
    df_todos_dados["Cliente"], st.session_state.filtros_clientes
)
m_ug = df_todos_dados["UG"].astype(str).map(_collapse_spaces).isin(
    st.session_state.filtros_ugs
)
m_tip = matches_any_canon(
    df_todos_dados["Tipo de ocorrência"], st.session_state.filtros_tipos
)
m_atv = matches_any_canon(
    df_todos_dados["Ativo"], st.session_state.filtros_ativos
)
m_ocr = matches_any_canon(
    df_todos_dados["Ocorrência"], st.session_state.filtros_ocorrencias
)

df_filtrado = df_todos_dados[
    m_ano & m_mes & m_dia & m_cat & m_cli & m_ug & m_tip & m_atv & m_ocr
].copy()

df_desligadas = df_filtrado[
    pd.isna(df_filtrado["Normalização"])
    | (df_filtrado["Normalização"] == "")
].copy()

with col_kpi2:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Total com Filtro Selecionado</div>
            <div class="kpi-value">{len(df_desligadas)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

if not df_desligadas.empty:
    maskvalid = df_desligadas["Desligamento"].notna()
    df_desligadas.loc[maskvalid, "Tempo em Segundos"] = (
        datetime.now() - df_desligadas.loc[maskvalid, "Desligamento"]
    ).dt.total_seconds().astype(int)
    df_desligadas.loc[~maskvalid, "Tempo em Segundos"] = 0

    # --- CONTROLES DE ORDENAÇÃO ---
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
        sortbydisplay = st.selectbox(
            "Ordenar por",
            options=sortoptionsdisplay.keys(),
            index=0,
        )
        sortbycolumn = sortoptionsdisplay[sortbydisplay]

    with sortcols[1]:
        sortorder = st.radio(
            "Ordem",
            options=["Descendente", "Ascendente"],
            index=0,
            horizontal=True,
        )
        isascending = sortorder == "Ascendente"

    dfsorted = df_desligadas.sort_values(
        by=sortbycolumn, ascending=isascending, na_position="last"
    )

    dfsorted["Display"] = (
        dfsorted["UG"].astype(str)
        + " | "
        + dfsorted["Ativo"].astype(str)
        + " | "
        + dfsorted["Nome Ativo"].astype(str)
        + " | "
        + dfsorted["Ocorrência"].astype(str)
        + " | "
        + dfsorted["Desligamento"].dt.strftime("%d/%m/%Y %H:%M").fillna("")
        + " | "
        + dfsorted["ID_Unico"].astype(str).str[-6:]
    )

    colsminimos = [
        "ID_Unico",
        "UG",
        "Ativo",
        "Nome Ativo",
        "Ocorrência",
        "Desligamento",
        "Categoria",
        "Tipo de ocorrência",
        "Operador",
        "Descrição",
        "OS",
        "Protocolo",
        "Normalização",
        "Atendimento Loop",
        "Atendimento Terceiros",
        "Cliente Avisado",
    ]
    colssalvar = [c for c in colsminimos if c in dfsorted.columns]
    st.session_state["df_lista_para_editar"] = dfsorted[colssalvar].copy()

    # --- SELEÇÃO PARA EDIÇÃO ---
    st.markdown("---")
    st.write("Editar uma Ocorrência")

    opts = dfsorted["Display"].dropna().astype(str).tolist()
    ocorrenciaselecionadadisplay = st.selectbox(
        "Selecione a ocorrência para editar",
        options=opts,
        index=None,
        placeholder="Escolha uma ocorrência...",
    )

    if ocorrenciaselecionadadisplay:
        idunicoparaeditar = (
            dfsorted.loc[
                dfsorted["Display"] == ocorrenciaselecionadadisplay, "ID_Unico"
            ]
            .head(1)
            .item()
        )
        st.session_state["id_unico_para_editar"] = idunicoparaeditar

    if (not ocorrenciaselecionadadisplay) and "id_unico_para_editar" in st.session_state:
        st.session_state.pop("id_unico_para_editar")

    btndisabled = not bool(ocorrenciaselecionadadisplay)
    if st.button(
        "Editar Ocorrência Selecionada",
        disabled=btndisabled,
        use_container_width=True,
    ):
        st.switch_page("pages/3_Editar_Ocorrencia.py")

    # --- LISTA DE OCORRÊNCIAS (TABELA) ---
    st.header("Lista de Ocorrências (Tabela)")

    dfparatabela = dfsorted.copy()

    def formatartempoestatico(row):
        dias = row["Tempo em Segundos"] // 86400
        horas = (row["Tempo em Segundos"] % 86400) // 3600
        minutos = (row["Tempo em Segundos"] % 3600) // 60
        return f"{dias}d {horas}h {minutos}m"

    dfparatabela["Tempo de Desligamento"] = dfparatabela.apply(
        formatartempoestatico, axis=1
    )
    dfparatabela.reset_index(inplace=True, drop=True)
    dfparatabela["Linha"] = dfparatabela.index + 1

    st.dataframe(
        dfparatabela[
            [
                "Linha",
                "Categoria",
                "Tempo de Desligamento",
                "UG",
                "Data",
                "Hora",
                "Tipo de ocorrência",
                "Ativo",
                "Ocorrência",
                "Operador",
                "Descrição",
                "OS",
            ]
        ],
        use_container_width=True,
    )

    # --- DETALHES POR OCORRÊNCIA (CARDS) ---
    st.header("Detalhes por Ocorrência (Cards)")

    num_cols = 2
    rows = list(dfsorted.iterrows())

    def formatdatetimecard(dtobj):
        if pd.notna(dtobj):
            return dtobj.strftime("%d/%m/%Y"), dtobj.strftime("%H:%M")
        return "", ""

    for i in range(0, len(rows), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j >= len(rows):
                continue

            index, row = rows[i + j]
            with cols[j]:
                cliente = html.escape(str(row.get("Cliente", "")))
                categoria = html.escape(str(row.get("Categoria", "")))
                ug = html.escape(str(row.get("UG", "N/A")))
                tipoocorrencia = html.escape(str(row.get("Tipo de ocorrência", "")))
                ativo = html.escape(str(row.get("Ativo", "")))
                nomeativo = html.escape(str(row.get("Nome Ativo", "")))
                ocorrencia = html.escape(str(row.get("Ocorrência", "")))
                operador = html.escape(str(row.get("Operador", "")))
                descricao = html.escape(str(row.get("Descrição", ""))).replace(
                    "\\n", "<br>"
                )
                protocolo = html.escape(str(row.get("Protocolo", "")))
                os = html.escape(str(row.get("OS", "")))

                dataocor, horaocor = formatdatetimecard(row.get("Desligamento"))
                dataca, horaca = formatdatetimecard(row.get("Cliente Avisado"))
                dataloop, horaloop = formatdatetimecard(row.get("Atendimento Loop"))
                dataterc, horaterc = formatdatetimecard(
                    row.get("Atendimento Terceiros")
                )
                datanorm, horanorm = formatdatetimecard(row.get("Normalização"))

                quantidadehtml = ""
                if row.get("Categoria") == "EQUIPAMENTOS":
                    quantidadeval = row.get("Quantidade", 0)
                    try:
                        if pd.notna(quantidadeval) and float(quantidadeval) > 0:
                            quantidadehtml = (
                                '<div class="card-item"><span class="card-label">'
                                f"Quantidade:</span> {int(float(quantidadeval))}</div>"
                            )
                    except (ValueError, TypeError):
                        quantidadehtml = ""

                cardhtml = f"""
                <div class="card-container">
                    <div class="card-title">{ug}</div>
                    <div class="card-item"><span class="card-label">Cliente:</span> {cliente}</div>
                    <div class="card-item"><span class="card-label">Categoria:</span> {categoria}</div>
                    <div class="card-item"><span class="card-label">Tipo de Ocorrência:</span> {tipoocorrencia}</div>
                    <div class="card-item"><span class="card-label">Ativo:</span> {ativo}</div>
                    <div class="card-item"><span class="card-label">Nome do ativo:</span> {nomeativo}</div>
                    <div class="card-item"><span class="card-label">Ocorrência:</span> {ocorrencia}</div>
                    <div class="card-item"><span class="card-label">Operador:</span> {operador}</div>
                    {quantidadehtml}
                    <br>
                    <div class="card-item"><span class="card-label">Data da ocorrência:</span> {dataocor}</div>
                    <div class="card-item"><span class="card-label">Hora da ocorrência:</span> {horaocor}</div>
                    <div class="card-item"><span class="card-label">Data cliente avisado:</span> {dataca}</div>
                    <div class="card-item"><span class="card-label">Hora cliente avisado:</span> {horaca}</div>
                    <div class="card-item"><span class="card-label">Data do atendimento LOOP:</span> {dataloop}</div>
                    <div class="card-item"><span class="card-label">Hora do atendimento LOOP:</span> {horaloop}</div>
                    <div class="card-item"><span class="card-label">Data do atendimento de terceiros:</span> {dataterc}</div>
                    <div class="card-item"><span class="card-label">Hora do atendimento de terceiros:</span> {horaterc}</div>
                    <div class="card-item"><span class="card-label">Data de normalização:</span> {datanorm}</div>
                    <div class="card-item"><span class="card-label">Hora de normalização:</span> {horanorm}</div>
                    <br>
                    <div class="card-item"><span class="card-label">Descrição:</span> {descricao}</div>
                    <div class="card-item"><span class="card-label">Protocolo:</span> {protocolo}</div>
                    <div class="card-item"><span class="card-label">OS:</span> {os}</div>
                </div>
                """
                st.html(cardhtml)
else:
    st.info(
        "Nenhuma usina encontrada com o campo Normalização em branco para os filtros selecionados."
    )
