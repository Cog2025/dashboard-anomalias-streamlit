import os
import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import html
import gspread
from google.oauth2.service_account import Credentials
import re
from collections import Counter, defaultdict
import utils  # [MODIFICADO]

# --- 1. Configuração da Página e Layout ---
st.set_page_config(layout="wide")

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


utils.render_loading_overlay()
if st.session_state.ui_phase == "init":
    start_loading()

# Failsafe opcional (20s)
if st.session_state.ui_phase == "loading" and (
    pytime.time() - st.session_state.loading_ts
) > 20:
    stop_loading()


def _collapse_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()


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
    ser = series.astype(str).map(_collapse_spaces)
    ser = ser[(ser != "") & (ser != "-") & (ser != "0")]
    return sorted(ser.unique().tolist())


def matches_any_canon(series: pd.Series, selected: list) -> pd.Series:
    if not selected:
        return pd.Series([True] * len(series), index=series.index)
    ser = series.astype(str).map(_collapse_spaces)
    return ser.isin(set(selected))


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

/* KPIs e cards – cores específicas da página 4 */
.kpi-card{
  background:#333; padding:clamp(12px, 3vw, 20px); border-radius:10px; text-align:center;
}
.kpi-value{
  font-size:clamp(1.6rem, 6vw, 3rem); font-weight:700; color:#4B4EFF;
}
.kpi-label{
  font-size:clamp(.85rem, 2.5vw, 1rem); color:#fff;
}
.card-container{
  background:#089641; color:#fff; padding:clamp(12px, 3vw, 16px); border-radius:8px;
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
@media (max-width:1100px){
  [data-testid="stHorizontalBlock"] > div:has(.stButton){
    flex-basis:100% !important;
    min-width:100% !important;
  }
}
</style>
""",
    unsafe_allow_html=True,
)


# --- 4. Carregar e Tratar os Dados ---


@st.cache_data(ttl=600)
def carregar_dados_google_sheets(cache_buster: int = 0):
    try:
        client = utils.connect_to_google_sheets()
        workbook = client.open_by_url(utils.SPREADSHEET_URL)

        df_desligamentos = utils.fetch_sheet_as_df(
            workbook.worksheet(utils.SHEET_DESLIGAMENTOS)
        )
        df_equipamentos = utils.fetch_sheet_as_df(
            workbook.worksheet(utils.SHEET_EQUIPAMENTOS)
        )

        df_desligamentos["Categoria"] = "DESLIGAMENTOS"
        df_equipamentos["Categoria"] = "EQUIPAMENTOS"
        df_todos = pd.concat([df_desligamentos, df_equipamentos], ignore_index=True)

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
        renomear_final = {}
        for col in df_todos.columns:
            cu = col.strip().upper()
            if cu in mapa_renomear:
                renomear_final[col] = mapa_renomear[cu]
        df_todos.rename(columns=renomear_final, inplace=True)

        df_todos.fillna("", inplace=True)

        for c in [
            "Normalização",
            "Desligamento",
            "Atendimento Loop",
            "Atendimento Terceiros",
            "Cliente Avisado",
        ]:
            if c in df_todos.columns:
                df_todos[c] = pd.to_datetime(df_todos[c], errors="coerce")

        if "Desligamento" in df_todos.columns and not df_todos[
            "Desligamento"
        ].isnull().all():
            df_todos["Data"] = df_todos["Desligamento"].dt.strftime("%Y-%m-%d")
            df_todos["Hora"] = df_todos["Desligamento"].dt.strftime("%H:%M:%S")
            df_todos["Mês"] = (
                df_todos["Desligamento"].dt.strftime("%B").map(meses_traducao)
            )
            df_todos["Ano"] = (
                df_todos["Desligamento"].dt.year.fillna(0).astype(int)
            )
            df_todos["Dia"] = (
                df_todos["Desligamento"].dt.day.fillna(0).astype(int)
            )
            df_todos["ID_Unico"] = (
                df_todos["UG"].astype(str).str.upper()
                + "|"
                + df_todos["Ativo"].astype(str).str.upper()
                + "|"
                + df_todos["Ocorrência"].astype(str).str.upper()
                + "|"
                + df_todos["Desligamento"].astype(str)
            )
        else:
            for c in ["Data", "Hora", "Mês", "Ano", "Dia", "ID_Unico"]:
                df_todos[c] = None

        for c in ["Operador", "Descrição", "OS", "Protocolo"]:
            if c in df_todos.columns:
                df_todos[c] = df_todos[c].astype(str).fillna("")

        return df_todos

    except Exception as e:
        st.error(f"Erro ao carregar ou processar os dados do Google Sheets: {e}")
        return pd.DataFrame()


if "cache_buster" not in st.session_state:
    st.session_state.cache_buster = int(pytime.time())

df = carregar_dados_google_sheets(st.session_state.cache_buster)
df["Desligamento"] = pd.to_datetime(df["Desligamento"], errors="coerce")

# --- 5. KPIs do topo (RESOLVIDAS) ---
count_resolvidos_deslig = df[
    (df["Categoria"] == "DESLIGAMENTOS") & (~df["Normalização"].isna())
].shape[0]
count_resolvidos_equip = df[
    (df["Categoria"] == "EQUIPAMENTOS") & (~df["Normalização"].isna())
].shape[0]

with st.container(border=True):
    st.markdown(
        "<h1 style='margin:0'>OCORRÊNCIAS RESOLVIDAS</h1>",
        unsafe_allow_html=True,
    )
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown(
            f"""
        <div class="kpi-card">
            <div class="kpi-label">DESLIGAMENTOS</div>
            <div class="kpi-value">{count_resolvidos_deslig}</div>
        </div>
        """,
            unsafe_allow_html=True,
        )
    with col_b:
        st.markdown(
            f"""
        <div class="kpi-card">
            <div class="kpi-label">EQUIPAMENTOS</div>
            <div class="kpi-value">{count_resolvidos_equip}</div>
        </div>
        """,
            unsafe_allow_html=True,
        )

# --- 6. Inicialização de filtros ---
st.header("OCORRÊNCIAS FILTRADAS")

if "filtros_meses" not in st.session_state:
    st.session_state.filtros_meses = [
        meses_traducao[datetime.now().strftime("%B")]
    ]
if "filtros_anos" not in st.session_state:
    if not df.empty and "Ano" in df.columns:
        anos_atuais = sorted(df["Ano"].unique().tolist())
        st.session_state.filtros_anos = [a for a in anos_atuais if a != 0]
    else:
        st.session_state.filtros_anos = []
if "filtros_dias" not in st.session_state:
    if not df.empty and {"Mês", "Ano"}.issubset(df.columns):
        dias_atuais = sorted(
            df[
                (df["Mês"].isin(st.session_state.filtros_meses))
                & (df["Ano"].isin(st.session_state.filtros_anos))
            ]["Dia"].unique().tolist()
        )
        st.session_state.filtros_dias = [d for d in dias_atuais if d != 0]
    else:
        st.session_state.filtros_dias = []
if "filtros_categorias" not in st.session_state:
    st.session_state.filtros_categorias = (
        sorted(df["Categoria"].unique().tolist()) if not df.empty else []
    )
if "filtros_clientes" not in st.session_state:
    st.session_state.filtros_clientes = (
        sorted(df["Cliente"].unique().tolist()) if not df.empty else []
    )
if "filtros_ugs" not in st.session_state:
    st.session_state.filtros_ugs = (
        sorted(df["UG"].unique().tolist()) if not df.empty else []
    )
if "filtros_tipos" not in st.session_state:
    st.session_state.filtros_tipos = (
        sorted(df["Tipo de ocorrência"].unique().tolist())
        if (not df.empty and "Tipo de ocorrência" in df.columns)
        else []
    )
if "filtros_ativos" not in st.session_state:
    st.session_state.filtros_ativos = (
        sorted(df["Ativo"].unique().tolist())
        if (not df.empty and "Ativo" in df.columns)
        else []
    )
if "filtros_ocorrencias" not in st.session_state:
    st.session_state.filtros_ocorrencias = (
        sorted(df["Ocorrência"].unique().tolist())
        if (not df.empty and "Ocorrência" in df.columns)
        else []
    )

col_kpi1, col_kpi2 = st.columns(2)
with col_kpi1:
    total_resolvidas_banco = df[~df["Normalização"].isna()].shape[0]
    st.markdown(
        f"""
    <div class="kpi-card">
        <div class="kpi-label">Total Resolvidas no Banco Completo</div>
        <div class="kpi-value">{total_resolvidas_banco}</div>
    </div>
    """,
        unsafe_allow_html=True,
    )

# Botão atualizar
col_left, _ = st.columns([0.2, 0.8])
with col_left:
    if st.button("Atualizar Dados"):
        st.cache_data.clear()
        start_loading()
        st.session_state.cache_buster = int(pytime.time())
        st.rerun()


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


def marcar_loading(prefixo_key, itens, filtro_key, marcar_todos, validos=None):
    start_loading()
    _marcar(prefixo_key, itens, filtro_key, marcar_todos, validos)


# --- Filtros por período ---
if not df.empty:
    if "categoria_top" not in st.session_state:
        st.session_state.categoria_top = "Ambas"

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
    anos_disponiveis = sorted([a for a in df["Ano"].unique() if a != 0])
    meses_disponiveis = meses_cronologicos[:]

    col_ano, col_mes, col_dia = st.columns(3)

    # Ano(s)
    with col_ano:
        with st.container(border=True):
            st.write("Ano(s)")
            with st.expander("Expandir anos"):
                for ano in anos_disponiveis:
                    st.checkbox(
                        str(ano),
                        key=f"cb_ano_{ano}",
                        value=(ano in st.session_state.filtros_anos),
                    )

            btn_ano1, btn_ano2 = st.columns(2)
            with btn_ano1:
                clicked_sel_ano = st.button(
                    "Sel. Todos",
                    key="sel_ano_res",
                    use_container_width=True,
                    on_click=marcar_loading,
                    args=("cb_ano_", anos_disponiveis, "filtros_anos", True),
                )
            with btn_ano2:
                clicked_des_ano = st.button(
                    "Desmarcar",
                    key="des_ano_res",
                    use_container_width=True,
                    on_click=marcar_loading,
                    args=("cb_ano_", anos_disponiveis, "filtros_anos", False),
                )
            if not (clicked_sel_ano or clicked_des_ano):
                st.session_state.filtros_anos = [
                    a
                    for a in anos_disponiveis
                    if st.session_state.get(f"cb_ano_{a}", False)
                ]


    # Mês(es)
    with col_mes:
        with st.container(border=True):
            st.write("Mês(es)")
            with st.expander("Expandir meses"):
                for mes in meses_disponiveis:
                    st.checkbox(
                        mes,
                        key=f"cb_mes_{mes}",
                        value=(mes in st.session_state.filtros_meses),
                    )

            btn_mes1, btn_mes2 = st.columns(2)
            with btn_mes1:
                clicked_sel_mes = st.button(
                    "Sel. Todos",
                    key="sel_mes_res",
                    use_container_width=True,
                    on_click=marcar_loading,
                    args=("cb_mes_", meses_disponiveis, "filtros_meses", True),
                )
            with btn_mes2:
                clicked_des_mes = st.button(
                    "Desmarcar",
                    key="des_mes_res",
                    use_container_width=True,
                    on_click=marcar_loading,
                    args=("cb_mes_", meses_disponiveis, "filtros_meses", False),
                )
            if not (clicked_sel_mes or clicked_des_mes):
                st.session_state.filtros_meses = [
                    m
                    for m in meses_disponiveis
                    if st.session_state.get(f"cb_mes_{m}", False)
                ]


    # Dia(s)
    anossel = [
        a for a in anos_disponiveis if st.session_state.get(f"cb_ano_{a}", False)
    ]
    mesessel = [
        m for m in meses_disponiveis if st.session_state.get(f"cb_mes_{m}", False)
    ]
    dias_disponiveis = sorted(
        df[
            df["Ano"].isin(anossel)
            & df["Mês"].isin(mesessel)
            & df["Dia"].notna()
            & (df["Dia"] > 0)
        ]["Dia"]
        .astype(int)
        .unique()
        .tolist()
        or list(range(1, 32))
    )

    with col_dia:
        with st.container(border=True):
            st.write("Dia(s)")
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

            btn_dia1, btn_dia2 = st.columns(2)
            with btn_dia1:
                clicked_sel_dia = st.button(
                    "Sel. Todos",
                    key="sel_dia_res",
                    use_container_width=True,
                    on_click=marcar_loading,
                    args=(
                        "cb_dia_",
                        list(range(1, 32)),
                        "filtros_dias",
                        True,
                        set(dias_disponiveis),
                    ),
                )
            with btn_dia2:
                clicked_des_dia = st.button(
                    "Desmarcar",
                    key="des_dia_res",
                    use_container_width=True,
                    on_click=marcar_loading,
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


# --- Filtros Adicionais (sobre resolvidas) ---
st.subheader("Filtros Adicionais")

df_ref = df[~df["Normalização"].isna()].copy()

cli_opts = options_from(df_ref["Cliente"]) if "Cliente" in df_ref.columns else []
ugs_series = (
    df_ref["UG"].astype(str).map(_collapse_spaces)
    if "UG" in df_ref.columns
    else pd.Series([], dtype=str)
)
ug_opts = sorted([u for u in ugs_series.unique().tolist() if u and u != "-"])

tip_opts = (
    options_from(df_ref["Tipo de ocorrência"])
    if "Tipo de ocorrência" in df_ref.columns
    else []
)
ocr_opts = (
    options_from(df_ref["Ocorrência"])
    if "Ocorrência" in df_ref.columns
    else []
)
atv_opts = (
    options_from(df_ref["Ativo"]) if "Ativo" in df_ref.columns else []
)


def set_filter_and_rerun(key, values):
    start_loading()
    st.session_state[key] = list(values)
    st.rerun()


st.session_state.filtros_clientes = [
    x for x in st.session_state.get("filtros_clientes", []) if x in cli_opts
]
st.session_state.filtros_ugs = [
    x for x in st.session_state.get("filtros_ugs", []) if x in ug_opts
]
st.session_state.filtros_tipos = [
    x for x in st.session_state.get("filtros_tipos", []) if x in tip_opts
]
st.session_state.filtros_ocorrencias = [
    x for x in st.session_state.get("filtros_ocorrencias", []) if x in ocr_opts
]
st.session_state.filtros_ativos = [
    x for x in st.session_state.get("filtros_ativos", []) if x in atv_opts
]

# 3 colunas em cima, 2 embaixo (como página 1)
row1_c1, row1_c2, row1_c3 = st.columns(3)
row2_c1, row2_c2 = st.columns(2)
col_cliente, col_ug, col_tipo, col_ativo, col_ocr = (
    row1_c1,
    row1_c2,
    row1_c3,
    row2_c1,
    row2_c2,
)

with col_cliente:
    with st.container(border=True):
        st.write("Cliente")

        btn_cli1, btn_cli2 = st.columns(2)
        with btn_cli1:
            if st.button(
                "Sel. Todos",
                key="cli_sel_all_res",
                use_container_width=True,
                on_click=set_filter_and_rerun,
                args=("filtros_clientes", cli_opts),
            ):
                pass
        with btn_cli2:
            if st.button(
                "Desmarcar",
                key="cli_clear_res",
                use_container_width=True,
                on_click=set_filter_and_rerun,
                args=("filtros_clientes", []),
            ):
                pass

        st.multiselect(
            "",
            options=cli_opts,
            default=st.session_state.filtros_clientes,
            label_visibility="hidden",
        )


with col_ug:
    with st.container(border=True):
        st.write("UG")

        btn_ug1, btn_ug2 = st.columns(2)
        with btn_ug1:
            if st.button(
                "Sel. Todos",
                key="ug_sel_all_res",
                use_container_width=True,
                on_click=set_filter_and_rerun,
                args=("filtros_ugs", ug_opts),
            ):
                pass
        with btn_ug2:
            if st.button(
                "Desmarcar",
                key="ug_clear_res",
                use_container_width=True,
                on_click=set_filter_and_rerun,
                args=("filtros_ugs", []),
            ):
                pass

        st.multiselect(
            "",
            options=ug_opts,
            default=st.session_state.filtros_ugs,
            label_visibility="hidden",
        )


with col_tipo:
    with st.container(border=True):
        st.write("Tipo de Ocorrência")

        btn_tipo1, btn_tipo2 = st.columns(2)
        with btn_tipo1:
            if st.button(
                "Sel. Todos",
                key="tipo_sel_all_res",
                use_container_width=True,
                on_click=set_filter_and_rerun,
                args=("filtros_tipos", tip_opts),
            ):
                pass
        with btn_tipo2:
            if st.button(
                "Desmarcar",
                key="tipo_clear_res",
                use_container_width=True,
                on_click=set_filter_and_rerun,
                args=("filtros_tipos", []),
            ):
                pass

        st.multiselect(
            "",
            options=tip_opts,
            default=st.session_state.filtros_tipos,
            label_visibility="hidden",
        )


with col_ativo:
    with st.container(border=True):
        st.write("Ativo")

        btn_atv1, btn_atv2 = st.columns(2)
        with btn_atv1:
            if st.button(
                "Sel. Todos",
                key="ativo_sel_all_res",
                use_container_width=True,
                on_click=set_filter_and_rerun,
                args=("filtros_ativos", atv_opts),
            ):
                pass
        with btn_atv2:
            if st.button(
                "Desmarcar",
                key="ativo_clear_res",
                use_container_width=True,
                on_click=set_filter_and_rerun,
                args=("filtros_ativos", []),
            ):
                pass

        st.multiselect(
            "",
            options=atv_opts,
            default=st.session_state.filtros_ativos,
            label_visibility="hidden",
        )


with col_ocr:
    with st.container(border=True):
        st.write("Ocorrência")

        btn_ocr1, btn_ocr2 = st.columns(2)
        with btn_ocr1:
            if st.button(
                "Sel. Todos",
                key="ocr_sel_all_res",
                use_container_width=True,
                on_click=set_filter_and_rerun,
                args=("filtros_ocorrencias", ocr_opts),
            ):
                pass
        with btn_ocr2:
            if st.button(
                "Desmarcar",
                key="ocr_clear_res",
                use_container_width=True,
                on_click=set_filter_and_rerun,
                args=("filtros_ocorrencias", []),
            ):
                pass

        st.multiselect(
            "",
            options=ocr_opts,
            default=st.session_state.filtros_ocorrencias,
            label_visibility="hidden",
        )


# --- Aplicação dos filtros ---
meses_selecionados = [
    mes for mes in meses_cronologicos if st.session_state.get(f"cb_mes_{mes}", False)
]
anos_selecionados = [
    ano for ano in anos_disponiveis if st.session_state.get(f"cb_ano_{ano}", False)
]
dias_selecionados = [
    dia for dia in dias_disponiveis if st.session_state.get(f"cb_dia_{dia}", False)
]

set_anos_disp = set(anos_disponiveis)
set_meses_disp = set(meses_cronologicos)
set_dias_disp = set(dias_disponiveis)

all_anos = set(anos_selecionados) == set_anos_disp and len(set_anos_disp) > 0
all_meses = set(meses_selecionados) == set_meses_disp and len(set_meses_disp) > 0
all_dias = set(dias_selecionados) == set_dias_disp and len(set_dias_disp) > 0

s_ano = df["Ano"]
s_mes = df["Mês"]
s_dia = df["Dia"]

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

m_cat = df["Categoria"].isin(st.session_state.filtros_categorias)
if st.session_state.get("categoria_top") in ("DESLIGAMENTOS", "EQUIPAMENTOS"):
    m_cat = m_cat & (df["Categoria"] == st.session_state["categoria_top"])

m_cli = matches_any_canon(df["Cliente"], st.session_state.filtros_clientes)
m_ug = df["UG"].astype(str).map(_collapse_spaces).isin(
    set(st.session_state.filtros_ugs)
)
m_tip = (
    matches_any_canon(df["Tipo de ocorrência"], st.session_state.filtros_tipos)
    if "Tipo de ocorrência" in df.columns
    else True
)
m_ocr = (
    matches_any_canon(df["Ocorrência"], st.session_state.filtros_ocorrencias)
    if "Ocorrência" in df.columns
    else True
)
m_atv = (
    matches_any_canon(df["Ativo"], st.session_state.filtros_ativos)
    if "Ativo" in df.columns
    else True
)

m_final = (
    m_cli & m_tip & m_ocr & m_atv & m_ug & m_ano & m_mes & m_dia & m_cat
)
df_filtrado = df[m_final].copy()

df_resolvidas = df_filtrado[~df_filtrado["Normalização"].isna()].copy()

if st.session_state.ui_phase == "loading":
    st.session_state.ui_phase = "ready"
    st.session_state.loading_ts = 0
    utils.render_loading_overlay("ready")

with col_kpi2:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Resolvidas com Filtro</div>
            <div class="kpi-value">{len(df_resolvidas)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# --- Exibição: tabela e cards ---
if not df_resolvidas.empty:
    st.markdown("---")
    st.write("Ordenar e Exibir")

    sortcols = st.columns(2)
    with sortcols[0]:
        sortoptionsdisplay = {
            "Data do Desligamento": "Desligamento",
            "Data da Normalização": "Normalização",
            "UG": "UG",
            "Ativo": "Ativo",
        }
        sortbydisplay = st.selectbox(
            "Ordenar por",
            options=list(sortoptionsdisplay.keys()),
            index=1,
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

    dfsorted = df_resolvidas.sort_values(
        by=sortbycolumn, ascending=isascending, na_position="last"
    )

    st.header("Lista de Ocorrências (Tabela)")
    dftab = dfsorted.copy()
    dftab.reset_index(inplace=True, drop=True)
    dftab["Linha"] = dftab.index + 1

    st.dataframe(
        dftab[
            [
                "Linha",
                "Categoria",
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

    # Cards
    st.header("Detalhes por Ocorrência (Cards)")
    num_cols = 3
    rows = list(dfsorted.iterrows())

    def fmt_dt(dtobj):
        if pd.notna(dtobj):
            return dtobj.strftime("%d/%m/%Y"), dtobj.strftime("%H:%M")
        return "", ""

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
                cliente = html.escape(str(r.get("Cliente", "")))
                categoria = html.escape(str(r.get("Categoria", "")))
                ug = html.escape(str(r.get("UG", "N/A")))
                tipo = html.escape(str(r.get("Tipo de ocorrência", "")))
                ativo = html.escape(str(r.get("Ativo", "")))
                nomeativo = html.escape(str(r.get("Nome Ativo", "")))
                ocorr = html.escape(str(r.get("Ocorrência", "")))
                operador = html.escape(str(r.get("Operador", "")))
                descricao = html.escape(str(r.get("Descrição", ""))).replace(
                    "\\n", "<br>"
                )
                protocolo = html.escape(str(r.get("Protocolo", "")))
                osv = html.escape(str(r.get("OS", "")))

                d_des, h_des = fmt_dt(r.get("Desligamento"))
                d_norm, h_norm = fmt_dt(r.get("Normalização"))
                d_ca, h_ca = fmt_dt(r.get("Cliente Avisado"))
                d_loop, h_loop = fmt_dt(r.get("Atendimento Loop"))
                d_terc, h_terc = fmt_dt(r.get("Atendimento Terceiros"))

                qtd_html = ""
                if r.get("Categoria") == "EQUIPAMENTOS":
                    qv = r.get("Quantidade", 0)
                    try:
                        if pd.notna(qv) and float(qv) > 0:
                            qtd_html = (
                                '<div class="card-item"><span class="card-label">'
                                f"Quantidade:</span> {int(float(qv))}</div>"
                            )
                    except (ValueError, TypeError):
                        qtd_html = ""

                cardhtml = f"""
                <div class="card-container">
                    <div class="card-title">{ug}</div>
                    <div class="card-item"><span class="card-label">Cliente:</span> {cliente}</div>
                    <div class="card-item"><span class="card-label">Categoria:</span> {categoria}</div>
                    <div class="card-item"><span class="card-label">Tipo de Ocorrência:</span> {tipo}</div>
                    <div class="card-item"><span class="card-label">Ativo:</span> {ativo}</div>
                    <div class="card-item"><span class="card-label">Nome do ativo:</span> {nomeativo}</div>
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
                st.html(cardhtml)
else:
    st.info(
        "Nenhuma ocorrência resolvida encontrada para os filtros selecionados."
    )
