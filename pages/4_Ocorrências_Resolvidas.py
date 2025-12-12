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
import utils

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(layout="wide", page_title="Ocorrências Resolvidas")

# Injeta o CSS do tema (Gerencia a cor de fundo dos cards KPI)
utils.render_page_config_and_css()

# CSS Específico desta página (Botões Verdes e Cards de Detalhe Verdes)
st.markdown("""
<style>
/* Ajuste de layout */
.main .block-container{ max-width: 100% !important; padding-left: 1rem !important; padding-right: 1rem !important; }
[data-testid="stHorizontalBlock"]{ display:flex !important; flex-wrap:wrap !important; gap:12px !important; }
[data-testid="column"]{ flex:1 1 320px !important; min-width:280px !important; }

/* Botões Verdes (Original) */
.stButton button{
  background-color:#28a745; color:#fff; border:none; border-radius:6px;
  box-shadow:0 2px 4px rgba(0,0,0,.2); transition:.2s ease;
  width:100%; min-height:42px; padding:8px 12px;
  font-weight:500; margin-bottom:6px !important;
}

/* Selectbox adjustments */
.stSelectbox > div > div, .stMultiSelect div[data-baseweb="select"]{ max-height: 90px !important; overflow-y: auto !important; }

/* Cards Verdes de Detalhe (Inferior) */
.card-container{
  background:#089641; color:#fff; padding:clamp(12px, 3vw, 16px); border-radius:8px;
  box-shadow:0 4px 8px rgba(0,0,0,.2); word-wrap:break-word; height: 100%;
}
.card-title{ font-size:clamp(1rem, 3vw, 1.2rem); font-weight:700; border-bottom:1px solid rgba(255,255,255,.45); padding-bottom:6px; margin-bottom:10px; }
.card-item{ font-size:clamp(.8rem, 2.1vw, .95rem); line-height:1.35; margin-bottom:6px; }
.card-label{ font-weight:700; }
</style>
""", unsafe_allow_html=True)

# Estado do overlay
if "ui_phase" not in st.session_state: st.session_state.ui_phase = "init"
if "loading_ts" not in st.session_state: st.session_state.loading_ts = 0
utils.render_loading_overlay(st.session_state.ui_phase)

def start_loading():
    st.session_state.ui_phase = "loading"
    st.session_state.loading_ts = pytime.time()

def stop_loading():
    st.session_state.ui_phase = "ready"
    st.session_state.loading_ts = 0

if st.session_state.ui_phase == "init": start_loading()
if st.session_state.ui_phase == "loading" and (pytime.time() - st.session_state.loading_ts) > 20: stop_loading()

# --- Funções Auxiliares ---
def _collapse_spaces(s: str) -> str: return re.sub(r"\s+", " ", str(s)).strip()
def canon(s) -> str: return _collapse_spaces(str(s)).casefold() if s else ""
def build_display_map(series: pd.Series) -> dict:
    buckets = defaultdict(Counter)
    for v in series.dropna():
        v_str = _collapse_spaces(str(v))
        if v_str: buckets[canon(v_str)][v_str] += 1
    return {ckey: counter.most_common(1)[0][0] for ckey, counter in buckets.items()}
def options_from(series: pd.Series) -> list:
    series = series.astype(str).map(_collapse_spaces)
    series = series[series != ""]
    dmap = build_display_map(series)
    return ["-"] + sorted({dmap[canon(v)] for v in series})
def matches_any_canon(series: pd.Series, selected: list[str]) -> pd.Series:
    if not selected: return pd.Series([True] * len(series), index=series.index)
    sel_c = {canon(s) for s in selected if s and s != "-"}
    return series.astype(str).map(canon).isin(sel_c)

meses_traducao = {"January": "Janeiro", "February": "Fevereiro", "March": "Março", "April": "Abril", "May": "Maio", "June": "Junho", "July": "Julho", "August": "Agosto", "September": "Setembro", "October": "Outubro", "November": "Novembro", "December": "Dezembro"}
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
        
        mapa_renomear = {"IDENTIFICADOR": "Identificador", "CLIENTE": "Cliente", "UG": "UG", "TIPO DE OCORRÊNCIA": "Tipo de ocorrência", "ATIVO": "Ativo", "NOME ATIVO": "Nome Ativo", "OCORRÊNCIA": "Ocorrência", "QUANTIDADE": "Quantidade", "SIGLA": "Sigla", "NORMALIZAÇÃO": "Normalização", "DESLIGAMENTO": "Desligamento", "OPERADOR": "Operador", "DESCRIÇÃO": "Descrição", "OS": "OS", "ATENDIMENTO LOOP": "Atendimento Loop", "ATENDIMENTO TERCEIROS": "Atendimento Terceiros", "PROTOCOLO": "Protocolo", "CLIENTE AVISADO": "Cliente Avisado"}
        renomear_final = {col: mapa_renomear[col.strip().upper()] for col in df_todos.columns if col.strip().upper() in mapa_renomear}
        df_todos.rename(columns=renomear_final, inplace=True)
        df_todos.fillna("", inplace=True)
        
        for c in ["Normalização", "Desligamento", "Atendimento Loop", "Atendimento Terceiros", "Cliente Avisado"]:
            if c in df_todos.columns: df_todos[c] = pd.to_datetime(df_todos[c], errors="coerce")
            
        if "Desligamento" in df_todos.columns:
            df_todos["Data"] = df_todos["Desligamento"].dt.strftime("%Y-%m-%d")
            df_todos["Hora"] = df_todos["Desligamento"].dt.strftime("%H:%M:%S")
            df_todos["Mês"] = df_todos["Desligamento"].dt.strftime("%B").map(meses_traducao)
            df_todos["Ano"] = df_todos["Desligamento"].dt.year.fillna(0).astype(int)
            df_todos["Dia"] = df_todos["Desligamento"].dt.day.fillna(0).astype(int)
        return df_todos
    except Exception as e:
        st.error(f"Erro ao processar dados: {e}")
        return pd.DataFrame()

if "cache_buster" not in st.session_state: st.session_state.cache_buster = int(pytime.time())
df = carregar_dados_google_sheets(st.session_state.cache_buster)
if not df.empty and "Desligamento" in df.columns: df["Desligamento"] = pd.to_datetime(df["Desligamento"], errors="coerce")

# --- 2. KPIs SUPERIORES (RESOLVIDAS) ---
count_resolvidos_deslig = 0
count_resolvidos_equip = 0
if not df.empty:
    count_resolvidos_deslig = df[(df["Categoria"] == "DESLIGAMENTOS") & (~df["Normalização"].isna())].shape[0]
    count_resolvidos_equip = df[(df["Categoria"] == "EQUIPAMENTOS") & (~df["Normalização"].isna())].shape[0]

with st.container(border=True):
    st.markdown("<h1 style='margin:0'>OCORRÊNCIAS RESOLVIDAS</h1>", unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    # Usando class 'kpi-card' (CSS dinâmico) + cor azul original
    with col_a:
        st.markdown(f"""<div class="kpi-card"><div class="kpi-label">DESLIGAMENTOS</div><div class="kpi-value" style="color: #4B4EFF;">{count_resolvidos_deslig}</div></div>""", unsafe_allow_html=True)
    with col_b:
        st.markdown(f"""<div class="kpi-card"><div class="kpi-label">EQUIPAMENTOS</div><div class="kpi-value" style="color: #4B4EFF;">{count_resolvidos_equip}</div></div>""", unsafe_allow_html=True)

# --- 3. FILTROS E LÓGICA ---
st.header("OCORRÊNCIAS FILTRADAS")

# Inicializa Session State
if "filtros_meses" not in st.session_state: st.session_state.filtros_meses = [meses_traducao[datetime.now().strftime("%B")]]
if "filtros_anos" not in st.session_state: st.session_state.filtros_anos = sorted([a for a in df["Ano"].unique() if a!=0]) if not df.empty else []
if "filtros_dias" not in st.session_state: st.session_state.filtros_dias = []
if "filtros_categorias" not in st.session_state: st.session_state.filtros_categorias = sorted(df["Categoria"].unique().tolist()) if not df.empty else []
if "categoria_top" not in st.session_state: st.session_state.categoria_top = "Ambas"
if "categoria_top_aplicada" not in st.session_state: st.session_state.categoria_top_aplicada = "Ambas"

# Filtros Adicionais State
for k in ["filtros_clientes", "filtros_ugs", "filtros_tipos", "filtros_ativos", "filtros_ocorrencias"]:
    if k not in st.session_state: st.session_state[k] = []

# Estados UI (Cópia para interface)
for k in ["ui_filtros_anos", "ui_filtros_meses", "ui_filtros_dias", "ui_filtros_clientes", "ui_filtros_ugs", "ui_filtros_tipos", "ui_filtros_ativos", "ui_filtros_ocorrencias"]:
    if k not in st.session_state: 
        original_key = k.replace("ui_", "")
        st.session_state[k] = st.session_state[original_key].copy()
if "ui_categoria_top" not in st.session_state: st.session_state.ui_categoria_top = st.session_state.categoria_top

# Botão Atualizar
col_left, _ = st.columns([0.2, 0.8])
with col_left:
    if st.button("Atualizar Dados"):
        st.cache_data.clear()
        start_loading()
        # Reset total para forçar reinit
        for k in list(st.session_state.keys()):
             if any(x in k for x in ["filtros_", "ui_", "cb_", "categoria"]): del st.session_state[k]
        st.rerun()

def _marcar(prefixo_key, itens, filtro_key, marcar_todos, validos=None):
    validos = set(itens) if validos is None else set(validos)
    st.session_state[filtro_key] = [x for x in itens if (x in validos) and marcar_todos]
    for x in itens: st.session_state[f"{prefixo_key}{x}"] = marcar_todos and (x in validos)

def marcar_loading(prefixo_key, itens, filtro_key, marcar_todos, validos=None):
    start_loading()
    _marcar(prefixo_key, itens, filtro_key, marcar_todos, validos)

# --- Renderização dos Filtros UI ---
if not df.empty:
    st.markdown("#### Filtrar por categoria (planilha)")
    st.session_state.ui_categoria_top = st.radio("Categoria:", ["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"], horizontal=True, label_visibility="collapsed", index=["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"].index(st.session_state.ui_categoria_top))

    st.subheader("Selecione o período desejado")
    
    # [CORREÇÃO NAME ERROR] Define os sets de disponíveis ANTES de usar
    anos_disponiveis = sorted([a for a in df["Ano"].unique() if a != 0])
    set_anos_disp = set(anos_disponiveis)
    
    meses_disponiveis = meses_cronologicos[:]
    set_meses_disp = set(meses_disponiveis)

    col_ano, col_mes, col_dia = st.columns(3)
    
    with col_ano:
        with st.container(border=True):
            st.write("Ano(s)")
            with st.expander("Expandir anos"):
                for ano in anos_disponiveis:
                    st.checkbox(str(ano), key=f"cb_ano_{ano}", value=(ano in st.session_state.ui_filtros_anos))
            b1, b2 = st.columns(2)
            with b1: st.button("Sel. Todos", key="sa_ano", on_click=marcar_loading, args=("cb_ano_", anos_disponiveis, "ui_filtros_anos", True))
            with b2: st.button("Desmarcar", key="dm_ano", on_click=marcar_loading, args=("cb_ano_", anos_disponiveis, "ui_filtros_anos", False))
            # Sincroniza UI com Checkbox
            st.session_state.ui_filtros_anos = [a for a in anos_disponiveis if st.session_state.get(f"cb_ano_{a}", False)]

    with col_mes:
        with st.container(border=True):
            st.write("Mês(es)")
            with st.expander("Expandir meses"):
                for mes in meses_disponiveis:
                    st.checkbox(mes, key=f"cb_mes_{mes}", value=(mes in st.session_state.ui_filtros_meses))
            b1, b2 = st.columns(2)
            with b1: st.button("Sel. Todos", key="sa_mes", on_click=marcar_loading, args=("cb_mes_", meses_disponiveis, "ui_filtros_meses", True))
            with b2: st.button("Desmarcar", key="dm_mes", on_click=marcar_loading, args=("cb_mes_", meses_disponiveis, "ui_filtros_meses", False))
            st.session_state.ui_filtros_meses = [m for m in meses_disponiveis if st.session_state.get(f"cb_mes_{m}", False)]

    # Calculo dias
    anossel = st.session_state.ui_filtros_anos or anos_disponiveis
    mesessel = st.session_state.ui_filtros_meses or meses_disponiveis
    dias_disponiveis = sorted(df[(df["Ano"].isin(anossel)) & (df["Mês"].isin(mesessel)) & (df["Dia"] > 0)]["Dia"].astype(int).unique().tolist() or list(range(1, 32)))
    set_dias_disp = set(dias_disponiveis)

    with col_dia:
        with st.container(border=True):
            st.write("Dia(s)")
            with st.expander("Expandir dias"):
                dias_cols = st.columns(7)
                for i, dia in enumerate(range(1, 32)):
                    with dias_cols[i % 7]:
                        disabled = dia not in dias_disponiveis
                        st.checkbox(str(dia), key=f"cb_dia_{dia}", value=(dia in st.session_state.ui_filtros_dias), disabled=disabled)
            b1, b2 = st.columns(2)
            with b1: st.button("Sel. Todos", key="sa_dia", on_click=marcar_loading, args=("cb_dia_", list(range(1, 32)), "ui_filtros_dias", True, set_dias_disp))
            with b2: st.button("Desmarcar", key="dm_dia", on_click=marcar_loading, args=("cb_dia_", list(range(1, 32)), "ui_filtros_dias", False, set_dias_disp))
            st.session_state.ui_filtros_dias = [d for d in dias_disponiveis if st.session_state.get(f"cb_dia_{d}", False)]

    # Filtros Adicionais
    st.subheader("Filtros Adicionais")
    df_ref = df[~df["Normalização"].isna()].copy()
    
    # Helper para renderizar bloco de filtro adicional
    def render_add_filter(col, label, key_suffix, col_name, state_key):
        with col:
            with st.container(border=True):
                st.write(label)
                opts = sorted([x for x in options_from(df_ref[col_name]) if x != "-"]) if col_name != "UG" and col_name != "Cliente" else sorted([x for x in df_ref[col_name].unique() if x and x!="-"])
                
                # Sincroniza: remove inválidos
                st.session_state[state_key] = [x for x in st.session_state[state_key] if x in opts]
                
                b1, b2 = st.columns(2)
                with b1: 
                    if st.button("Sel. Todos", key=f"sa_{key_suffix}"): st.session_state[state_key] = opts[:]
                with b2:
                    if st.button("Desmarcar", key=f"dm_{key_suffix}"): st.session_state[state_key] = []
                
                st.session_state[state_key] = st.multiselect("", opts, default=st.session_state[state_key], key=f"ms_{key_suffix}", label_visibility="hidden")

    r1c1, r1c2, r1c3 = st.columns(3)
    r2c1, r2c2 = st.columns(2)
    
    render_add_filter(r1c1, "Cliente", "cli", "Cliente", "ui_filtros_clientes")
    render_add_filter(r1c2, "UG", "ug", "UG", "ui_filtros_ugs")
    render_add_filter(r1c3, "Tipo", "tipo", "Tipo de ocorrência", "ui_filtros_tipos")
    render_add_filter(r2c1, "Ativo", "ativo", "Ativo", "ui_filtros_ativos")
    render_add_filter(r2c2, "Ocorrência", "ocorr", "Ocorrência", "ui_filtros_ocorrencias")

    if st.button("FILTRAR AGORA", use_container_width=True):
        start_loading()
        # Copia UI para Filtros Reais
        st.session_state.filtros_anos = st.session_state.ui_filtros_anos[:]
        st.session_state.filtros_meses = st.session_state.ui_filtros_meses[:]
        st.session_state.filtros_dias = st.session_state.ui_filtros_dias[:]
        st.session_state.categoria_top_aplicada = st.session_state.ui_categoria_top
        for k in ["filtros_clientes", "filtros_ugs", "filtros_tipos", "filtros_ativos", "filtros_ocorrencias"]:
            st.session_state[k] = st.session_state["ui_" + k][:]
        st.rerun()

# --- APLICAÇÃO DOS FILTROS ---
# [CORREÇÃO] set_anos_disp agora existe aqui
all_anos = set(st.session_state.filtros_anos) == set_anos_disp and len(set_anos_disp) > 0
all_meses = set(st.session_state.filtros_meses) == set_meses_disp and len(set_meses_disp) > 0
# Para dias, assumimos todos se a lista estiver vazia na logica original ou cheia
all_dias = set(st.session_state.filtros_dias) == set_dias_disp and len(set_dias_disp) > 0

m_ano = df["Ano"].isin(st.session_state.filtros_anos) if not all_anos else pd.Series(True, index=df.index)
m_mes = df["Mês"].isin(st.session_state.filtros_meses) if not all_meses else pd.Series(True, index=df.index)
m_dia = df["Dia"].isin(st.session_state.filtros_dias) if not all_dias and st.session_state.filtros_dias else pd.Series(True, index=df.index)

m_cat = df["Categoria"].isin(st.session_state.filtros_categorias)
if st.session_state.categoria_top_aplicada in ["DESLIGAMENTOS", "EQUIPAMENTOS"]:
    m_cat = m_cat & (df["Categoria"] == st.session_state.categoria_top_aplicada)

def mk_mask(col, sel): return matches_any_canon(df[col], sel) if sel else pd.Series(True, index=df.index)
def mk_mask_exact(col, sel): return df[col].isin(sel) if sel else pd.Series(True, index=df.index)

m_final = m_ano & m_mes & m_dia & m_cat & mk_mask("Cliente", st.session_state.filtros_clientes) & mk_mask_exact("UG", st.session_state.filtros_ugs) & mk_mask("Tipo de ocorrência", st.session_state.filtros_tipos) & mk_mask("Ativo", st.session_state.filtros_ativos) & mk_mask("Ocorrência", st.session_state.filtros_ocorrencias)

df_filtrado = df[m_final].copy()
df_resolvidas = df_filtrado[~df_filtrado["Normalização"].isna()].copy()

if st.session_state.ui_phase == "loading": stop_loading()

# --- 4. KPIs INFERIORES (FILTRADAS) ---
c1, c2 = st.columns(2)
with c1:
    total_banco = df[~df["Normalização"].isna()].shape[0]
    st.markdown(f"""<div class="kpi-card"><div class="kpi-label">Total Resolvidas no Banco Completo</div><div class="kpi-value" style="color: #4B4EFF;">{total_banco}</div></div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""<div class="kpi-card"><div class="kpi-label">Total Resolvidas com Filtro</div><div class="kpi-value" style="color: #4B4EFF;">{len(df_resolvidas)}</div></div>""", unsafe_allow_html=True)

# --- 5. TABELA E CARDS ---
if not df_resolvidas.empty:
    st.markdown("---")
    st.write("Ordenar e Exibir")
    sc1, sc2 = st.columns(2)
    with sc1:
        sort_col = st.selectbox("Ordenar por", ["Desligamento", "Normalização", "UG", "Ativo"], index=1)
    with sc2:
        asc = st.radio("Ordem", ["Descendente", "Ascendente"], horizontal=True) == "Ascendente"
    
    df_sorted = df_resolvidas.sort_values(by=sort_col, ascending=asc, na_position="last").reset_index(drop=True)
    df_sorted["Linha"] = df_sorted.index + 1
    
    st.header("Lista de Ocorrências (Tabela)")
    st.dataframe(df_sorted[["Linha", "Categoria", "UG", "Data", "Hora", "Tipo de ocorrência", "Ativo", "Ocorrência", "Operador", "Descrição", "OS"]], use_container_width=True)

    st.header("Detalhes por Ocorrência (Cards)")
    num_cols = 3
    rows = list(df_sorted.iterrows())
    
    def fmt_dt(d): return (d.strftime('%d/%m/%Y'), d.strftime('%H:%M')) if pd.notna(d) else ('', '')

    for i in range(0, len(rows), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i+j >= len(rows): break
            _, r = rows[i+j]
            with cols[j]:
                d_des, h_des = fmt_dt(r.get("Desligamento"))
                d_norm, h_norm = fmt_dt(r.get("Normalização"))
                d_ca, h_ca = fmt_dt(r.get("Cliente Avisado"))
                d_loop, h_loop = fmt_dt(r.get("Atendimento Loop"))
                d_terc, h_terc = fmt_dt(r.get("Atendimento Terceiros"))
                
                qtd_html = ""
                try: 
                    if r.get("Categoria") == "EQUIPAMENTOS" and float(r.get("Quantidade",0)) > 0:
                        qtd_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {int(float(r.get("Quantidade")))}</div>'
                except: pass

                st.markdown(f"""
                <div class="card-container">
                    <div class="card-title">{html.escape(str(r.get("UG","")))}</div>
                    <div class="card-item"><span class="card-label">Cliente:</span> {html.escape(str(r.get("Cliente","")))}</div>
                    <div class="card-item"><span class="card-label">Categoria:</span> {html.escape(str(r.get("Categoria","")))}</div>
                    <div class="card-item"><span class="card-label">Tipo:</span> {html.escape(str(r.get("Tipo de ocorrência","")))}</div>
                    <div class="card-item"><span class="card-label">Ativo:</span> {html.escape(str(r.get("Ativo","")))} | {html.escape(str(r.get("Nome Ativo","")))}</div>
                    <div class="card-item"><span class="card-label">Ocorrência:</span> {html.escape(str(r.get("Ocorrência","")))}</div>
                    <div class="card-item"><span class="card-label">Operador:</span> {html.escape(str(r.get("Operador","")))}</div>
                    {qtd_html}
                    <br>
                    <div class="card-item"><span class="card-label">Desligamento:</span> {d_des} {h_des}</div>
                    <div class="card-item"><span class="card-label">Normalização:</span> {d_norm} {h_norm}</div>
                    <div class="card-item"><span class="card-label">Cliente Avisado:</span> {d_ca} {h_ca}</div>
                    <div class="card-item"><span class="card-label">Loop:</span> {d_loop} {h_loop}</div>
                    <div class="card-item"><span class="card-label">Terceiros:</span> {d_terc} {h_terc}</div>
                    <br>
                    <div class="card-item"><span class="card-label">Descrição:</span> {html.escape(str(r.get("Descrição",""))).replace("\\n", "<br>")}</div>
                    <div class="card-item"><span class="card-label">Protocolo:</span> {html.escape(str(r.get("Protocolo","")))}</div>
                    <div class="card-item"><span class="card-label">OS:</span> {html.escape(str(r.get("OS","")))}</div>
                </div>
                """, unsafe_allow_html=True)
else:
    st.info("Nenhuma ocorrência resolvida encontrada para os filtros selecionados.")