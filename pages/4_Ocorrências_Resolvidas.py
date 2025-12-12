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
import utils # Importando utils

# Configuração da página
st.set_page_config(layout="wide", page_title="Ocorrências Resolvidas")

# [NOVO] Injeta CSS Inteligente para os KPIs
utils.render_page_config_and_css()

# CSS Específico desta página (Cards Verdes para os detalhes)
st.markdown("""
<style>
    /* Card Container (Verde nesta página) */
    .card-container {
        background-color: #089641; 
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
    return re.sub(r"\s+", " ", str(s)).strip()

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
        workbook = client.open_by_url(utils.SPREADSHEET_URL)
        df_desligamentos = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_DESLIGAMENTOS))
        df_equipamentos = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_EQUIPAMENTOS))

        df_desligamentos["Categoria"] = "DESLIGAMENTOS"
        df_equipamentos["Categoria"] = "EQUIPAMENTOS"
        df_todos = pd.concat([df_desligamentos, df_equipamentos], ignore_index=True)

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
        for col in df_todos.columns:
            cu = col.strip().upper()
            if cu in mapa_renomear:
                renomear_final[col] = mapa_renomear[cu]
        df_todos.rename(columns=renomear_final, inplace=True)
        df_todos.fillna("", inplace=True)

        for c in ["Normalização", "Desligamento", "Atendimento Loop", "Atendimento Terceiros", "Cliente Avisado"]:
            if c in df_todos.columns:
                df_todos[c] = pd.to_datetime(df_todos[c], errors="coerce")

        if "Desligamento" in df_todos.columns:
            df_todos["Data"] = df_todos["Desligamento"].dt.strftime("%Y-%m-%d")
            df_todos["Hora"] = df_todos["Desligamento"].dt.strftime("%H:%M:%S")
            df_todos["Mês"] = df_todos["Desligamento"].dt.strftime("%B").map(meses_traducao)
            df_todos["Ano"] = df_todos["Desligamento"].dt.year.fillna(0).astype(int)
            df_todos["Dia"] = df_todos["Desligamento"].dt.day.fillna(0).astype(int)
        
        return df_todos

    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

if "cache_buster" not in st.session_state:
    st.session_state.cache_buster = int(pytime.time())

df = carregar_dados_google_sheets(st.session_state.cache_buster)
if not df.empty:
    df["Desligamento"] = pd.to_datetime(df["Desligamento"], errors="coerce")

if st.session_state.ui_phase != "ready":
    stop_loading()


# --- KPIs do topo (RESOLVIDAS) ---
count_resolvidos_deslig = 0
count_resolvidos_equip = 0
if not df.empty:
    count_resolvidos_deslig = df[(df["Categoria"] == "DESLIGAMENTOS") & (~df["Normalização"].isna())].shape[0]
    count_resolvidos_equip = df[(df["Categoria"] == "EQUIPAMENTOS") & (~df["Normalização"].isna())].shape[0]

with st.container(border=True):
    st.markdown("<h1 style='margin:0; text-align:center;'>OCORRÊNCIAS RESOLVIDAS</h1>", unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    # [CORRIGIDO] KPIs com CSS do utils, mas cor do texto azul (estilo original desta página)
    with col_a:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">DESLIGAMENTOS</div>
            <div class="kpi-value" style="color: #4B4EFF;">{count_resolvidos_deslig}</div>
        </div>
        """, unsafe_allow_html=True)
    with col_b:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">EQUIPAMENTOS</div>
            <div class="kpi-value" style="color: #4B4EFF;">{count_resolvidos_equip}</div>
        </div>
        """, unsafe_allow_html=True)

# --- 6. Inicialização de filtros ---
st.header("OCORRÊNCIAS FILTRADAS")

# Inicialização de Session State para filtros (igual ao original, simplificado aqui para renderização)
if "filtros_meses" not in st.session_state: st.session_state.filtros_meses = [meses_traducao[datetime.now().strftime("%B")]]
if "filtros_anos" not in st.session_state: st.session_state.filtros_anos = [a for a in sorted(df["Ano"].unique().tolist()) if a != 0] if not df.empty else []
if "filtros_dias" not in st.session_state: st.session_state.filtros_dias = []
if "filtros_categorias" not in st.session_state: st.session_state.filtros_categorias = sorted(df["Categoria"].unique().tolist()) if not df.empty else []
if "filtros_clientes" not in st.session_state: st.session_state.filtros_clientes = []
if "filtros_ugs" not in st.session_state: st.session_state.filtros_ugs = []
if "filtros_tipos" not in st.session_state: st.session_state.filtros_tipos = []
if "filtros_ativos" not in st.session_state: st.session_state.filtros_ativos = []
if "filtros_ocorrencias" not in st.session_state: st.session_state.filtros_ocorrencias = []
if "categoria_top" not in st.session_state: st.session_state.categoria_top = "Ambas"

# Botão Atualizar
col_left, _ = st.columns([0.2, 0.8])
with col_left:
    if st.button("Atualizar Dados"):
        st.cache_data.clear()
        start_loading()
        st.rerun()

# --- Interface de Filtros ---
if not df.empty:
    st.markdown("#### Filtrar por categoria (planilha)")
    st.radio("Categoria:", options=["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"], horizontal=True, label_visibility="collapsed", key="categoria_top")

    st.subheader("Selecione o período desejado")
    anos_disponiveis = sorted([a for a in df["Ano"].unique() if a != 0])
    meses_disponiveis = meses_cronologicos[:]

    col_ano, col_mes, col_dia = st.columns(3)
    
    # Ano
    with col_ano:
        with st.container(border=True):
            st.write("Ano(s)")
            with st.expander("Expandir anos"):
                for ano in anos_disponiveis:
                    st.checkbox(str(ano), key=f"cb_ano_{ano}", value=(ano in st.session_state.filtros_anos))
            if st.button("Sel. Todos", key="sel_ano_res"):
                st.session_state.filtros_anos = anos_disponiveis
                st.rerun()
            if st.button("Desmarcar", key="des_ano_res"):
                st.session_state.filtros_anos = []
                st.rerun()
            st.session_state.filtros_anos = [a for a in anos_disponiveis if st.session_state.get(f"cb_ano_{a}", False)]

    # Mês
    with col_mes:
        with st.container(border=True):
            st.write("Mês(es)")
            with st.expander("Expandir meses"):
                for mes in meses_disponiveis:
                    st.checkbox(mes, key=f"cb_mes_{mes}", value=(mes in st.session_state.filtros_meses))
            if st.button("Sel. Todos", key="sel_mes_res"):
                st.session_state.filtros_meses = meses_disponiveis
                st.rerun()
            if st.button("Desmarcar", key="des_mes_res"):
                st.session_state.filtros_meses = []
                st.rerun()
            st.session_state.filtros_meses = [m for m in meses_disponiveis if st.session_state.get(f"cb_mes_{m}", False)]

    # Dia
    dias_disponiveis = list(range(1, 32))
    with col_dia:
        with st.container(border=True):
            st.write("Dia(s)")
            with st.expander("Expandir dias"):
                dias_cols = st.columns(7)
                for i, dia in enumerate(range(1, 32)):
                    with dias_cols[i % 7]:
                        st.checkbox(str(dia), key=f"cb_dia_{dia}", value=(dia in st.session_state.filtros_dias))
            if st.button("Sel. Todos", key="sel_dia_res"):
                st.session_state.filtros_dias = dias_disponiveis
                st.rerun()
            if st.button("Desmarcar", key="des_dia_res"):
                st.session_state.filtros_dias = []
                st.rerun()
            st.session_state.filtros_dias = [d for d in dias_disponiveis if st.session_state.get(f"cb_dia_{d}", False)]

    # Filtros Adicionais
    st.subheader("Filtros Adicionais")
    df_ref = df[~df["Normalização"].isna()].copy()

    row1_c1, row1_c2, row1_c3 = st.columns(3)
    row2_c1, row2_c2 = st.columns(2)
    col_cliente, col_ug, col_tipo, col_ativo, col_ocorrencia = row1_c1, row1_c2, row1_c3, row2_c1, row2_c2

    # Cliente
    with col_cliente:
        with st.container(border=True):
            st.write("Cliente")
            cli_opts = sorted([x for x in df_ref["Cliente"].unique() if x and x!="-"])
            if st.button("Limpar", key="limpar_cli_res"): st.session_state.filtros_clientes = []
            st.session_state.filtros_clientes = st.multiselect("", options=cli_opts, default=[x for x in st.session_state.filtros_clientes if x in cli_opts])

    # UG
    with col_ug:
        with st.container(border=True):
            st.write("UG")
            ug_opts = sorted([x for x in df_ref["UG"].unique() if x and x!="-"])
            if st.button("Limpar", key="limpar_ug_res"): st.session_state.filtros_ugs = []
            st.session_state.filtros_ugs = st.multiselect("", options=ug_opts, default=[x for x in st.session_state.filtros_ugs if x in ug_opts])
    
    # Tipo
    with col_tipo:
        with st.container(border=True):
            st.write("Tipo")
            tp_opts = sorted([x for x in df_ref["Tipo de ocorrência"].unique() if x and x!="-"])
            if st.button("Limpar", key="limpar_tp_res"): st.session_state.filtros_tipos = []
            st.session_state.filtros_tipos = st.multiselect("", options=tp_opts, default=[x for x in st.session_state.filtros_tipos if x in tp_opts])

    # Ativo
    with col_ativo:
        with st.container(border=True):
            st.write("Ativo")
            at_opts = sorted([x for x in df_ref["Ativo"].unique() if x and x!="-"])
            if st.button("Limpar", key="limpar_at_res"): st.session_state.filtros_ativos = []
            st.session_state.filtros_ativos = st.multiselect("", options=at_opts, default=[x for x in st.session_state.filtros_ativos if x in at_opts])

    # Ocorrência
    with col_ocorrencia:
        with st.container(border=True):
            st.write("Ocorrência")
            oc_opts = sorted([x for x in df_ref["Ocorrência"].unique() if x and x!="-"])
            if st.button("Limpar", key="limpar_oc_res"): st.session_state.filtros_ocorrencias = []
            st.session_state.filtros_ocorrencias = st.multiselect("", options=oc_opts, default=[x for x in st.session_state.filtros_ocorrencias if x in oc_opts])

    if st.button("FILTRAR AGORA", use_container_width=True):
        st.rerun()

# --- Aplicação dos filtros ---
m_ano = df["Ano"].isin(st.session_state.filtros_anos) if st.session_state.filtros_anos else pd.Series(True, index=df.index)
m_mes = df["Mês"].isin(st.session_state.filtros_meses) if st.session_state.filtros_meses else pd.Series(True, index=df.index)
m_dia = df["Dia"].isin(st.session_state.filtros_dias) if st.session_state.filtros_dias else pd.Series(True, index=df.index)

m_cat = df["Categoria"].isin(st.session_state.filtros_categorias)
if st.session_state.categoria_top in ("DESLIGAMENTOS", "EQUIPAMENTOS"):
    m_cat = m_cat & (df["Categoria"] == st.session_state.categoria_top)

m_cli = df["Cliente"].isin(st.session_state.filtros_clientes) if st.session_state.filtros_clientes else pd.Series(True, index=df.index)
m_ug = df["UG"].isin(st.session_state.filtros_ugs) if st.session_state.filtros_ugs else pd.Series(True, index=df.index)
m_tip = df["Tipo de ocorrência"].isin(st.session_state.filtros_tipos) if st.session_state.filtros_tipos else pd.Series(True, index=df.index)
m_atv = df["Ativo"].isin(st.session_state.filtros_ativos) if st.session_state.filtros_ativos else pd.Series(True, index=df.index)
m_ocr = df["Ocorrência"].isin(st.session_state.filtros_ocorrencias) if st.session_state.filtros_ocorrencias else pd.Series(True, index=df.index)

m_final = m_cli & m_tip & m_ocr & m_atv & m_ug & m_ano & m_mes & m_dia & m_cat
df_filtrado = df[m_final].copy()
df_resolvidas = df_filtrado[~df_filtrado["Normalização"].isna()].copy()

# KPIs Filtrados (Correção das Cores)
col_kpi1, col_kpi2 = st.columns(2)
with col_kpi1:
    total_resolvidas_banco = df[~df["Normalização"].isna()].shape[0]
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total Resolvidas no Banco Completo</div>
        <div class="kpi-value" style="color: #4B4EFF;">{total_resolvidas_banco}</div>
    </div>
    """, unsafe_allow_html=True)

with col_kpi2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total Resolvidas com Filtro</div>
        <div class="kpi-value" style="color: #4B4EFF;">{len(df_resolvidas)}</div>
    </div>
    """, unsafe_allow_html=True)


if not df_resolvidas.empty:
    st.markdown("---")
    st.write("Ordenar e Exibir")

    sortcols = st.columns(2)
    with sortcols[0]:
        sortoptionsdisplay = {"Data do Desligamento": "Desligamento", "Data da Normalização": "Normalização", "UG": "UG", "Ativo": "Ativo"}
        sortbydisplay = st.selectbox("Ordenar por", options=list(sortoptionsdisplay.keys()), index=1)
        sortbycolumn = sortoptionsdisplay[sortbydisplay]

    with sortcols[1]:
        sortorder = st.radio("Ordem", options=["Descendente", "Ascendente"], index=0, horizontal=True)
        isascending = sortorder == "Ascendente"

    dfsorted = df_resolvidas.sort_values(by=sortbycolumn, ascending=isascending, na_position="last")
    dfsorted.reset_index(inplace=True, drop=True)
    dfsorted["Linha"] = dfsorted.index + 1

    st.header("Lista de Ocorrências (Tabela)")
    st.dataframe(dfsorted[["Linha", "Categoria", "UG", "Data", "Hora", "Tipo de ocorrência", "Ativo", "Ocorrência", "Operador", "Descrição", "OS"]], use_container_width=True)

    st.header("Detalhes por Ocorrência (Cards)")
    num_cols = 3
    rows = list(dfsorted.iterrows())

    def fmt_dt(dtobj):
        if pd.notna(dtobj): return dtobj.strftime("%d/%m/%Y"), dtobj.strftime("%H:%M")
        return "", ""

    for i in range(0, len(rows), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j >= len(rows): break
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
                descricao = html.escape(str(r.get("Descrição", ""))).replace("\\n", "<br>")
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
                            qtd_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {int(float(qv))}</div>'
                    except: pass

                # Card HTML (Verde)
                cardhtml = f"""
                <div class="card-container">
                    <div class="card-title">{ug}</div>
                    <div class="card-item"><span class="card-label">Cliente:</span> {cliente}</div>
                    <div class="card-item"><span class="card-label">Categoria:</span> {categoria}</div>
                    <div class="card-item"><span class="card-label">Tipo:</span> {tipo}</div>
                    <div class="card-item"><span class="card-label">Ativo:</span> {ativo} | {nomeativo}</div>
                    <div class="card-item"><span class="card-label">Ocorrência:</span> {ocorr}</div>
                    <div class="card-item"><span class="card-label">Operador:</span> {operador}</div>
                    {qtd_html}
                    <br>
                    <div class="card-item"><span class="card-label">Desligamento:</span> {d_des} {h_des}</div>
                    <div class="card-item"><span class="card-label">Normalização:</span> {d_norm} {h_norm}</div>
                    <div class="card-item"><span class="card-label">Cliente Avisado:</span> {d_ca} {h_ca}</div>
                    <div class="card-item"><span class="card-label">Loop:</span> {d_loop} {h_loop}</div>
                    <div class="card-item"><span class="card-label">Terceiros:</span> {d_terc} {h_terc}</div>
                    <br>
                    <div class="card-item"><span class="card-label">Descrição:</span> {descricao}</div>
                    <div class="card-item"><span class="card-label">Protocolo:</span> {protocolo}</div>
                    <div class="card-item"><span class="card-label">OS:</span> {osv}</div>
                </div>
                """
                st.html(cardhtml)
else:
    st.info("Nenhuma ocorrência resolvida encontrada para os filtros selecionados.")