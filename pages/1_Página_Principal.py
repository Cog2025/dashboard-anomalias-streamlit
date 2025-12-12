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

# --- 1. CONFIGURAÇÃO ---
st.set_page_config(layout="wide", page_title="Ocorrências Ativas")
utils.render_page_config_and_css()

# CSS Específico desta página (Cards de Detalhe Vermelhos)
st.markdown("""
<style>
/* Layout e Botões Verdes */
.main .block-container{ max-width: 100% !important; padding-left: 1rem !important; padding-right: 1rem !important; }
[data-testid="stHorizontalBlock"]{ display:flex !important; flex-wrap:wrap !important; gap:12px !important; }
[data-testid="column"]{ flex:1 1 320px !important; min-width:280px !important; }
.stButton button{ background-color:#28a745; color:#fff; border:none; border-radius:6px; box-shadow:0 2px 4px rgba(0,0,0,.2); width:100%; min-height:42px; font-weight:500; margin-bottom:6px !important; }

/* Cards Vermelhos */
.card-container{ background:#FF4B4B; color:#fff; padding:clamp(12px, 3vw, 16px); border-radius:8px; box-shadow:0 4px 8px rgba(0,0,0,.2); word-wrap:break-word; height: 100%; }
.card-title{ font-size:clamp(1rem, 3vw, 1.2rem); font-weight:700; border-bottom:1px solid rgba(255,255,255,.45); padding-bottom:6px; margin-bottom:10px; }
.card-item{ font-size:clamp(.8rem, 2.1vw, .95rem); line-height:1.35; margin-bottom:6px; }
.card-label{ font-weight:700; }
</style>
""", unsafe_allow_html=True)

if "categoria_top" not in st.session_state: st.session_state["categoria_top"] = "Ambas"
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
def _collapse_spaces(s): return re.sub(r"\s+", " ", str(s)).strip()
def canon(s): return _collapse_spaces(s).casefold() if s else ""
def build_display_map(series):
    buckets = defaultdict(Counter)
    for v in series.dropna():
        vs = _collapse_spaces(str(v))
        if vs: buckets[canon(vs)][vs] += 1
    return {ckey: c.most_common(1)[0][0] for ckey, c in buckets.items()}
def options_from(series):
    return ["-"] + sorted({build_display_map(series.astype(str))[canon(str(v))] for v in series if str(v).strip()})
def matches_any_canon(series, selected):
    if not selected: return pd.Series([True]*len(series), index=series.index)
    sel_c = {canon(s) for s in selected if s and s != "-"}
    return series.astype(str).map(canon).isin(sel_c)

meses_traducao = {"January": "Janeiro", "February": "Fevereiro", "March": "Março", "April": "Abril", "May": "Maio", "June": "Junho", "July": "Julho", "August": "Agosto", "September": "Setembro", "October": "Outubro", "November": "Novembro", "December": "Dezembro"}
meses_cronologicos = list(meses_traducao.values())

@st.cache_data(ttl=600)
def carregar_dados_google_sheets(cb):
    try:
        client = utils.connect_to_google_sheets()
        if not client: return pd.DataFrame()
        workbook = client.open_by_url(utils.SPREADSHEET_URL)
        df_des = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_DESLIGAMENTOS))
        df_eq = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_EQUIPAMENTOS))
        df_des["Categoria"] = "DESLIGAMENTOS"
        df_eq["Categoria"] = "EQUIPAMENTOS"
        df = pd.concat([df_des, df_eq], ignore_index=True)
        
        rename = {"IDENTIFICADOR": "Identificador", "CLIENTE": "Cliente", "UG": "UG", "TIPO DE OCORRÊNCIA": "Tipo de ocorrência", "ATIVO": "Ativo", "NOME ATIVO": "Nome Ativo", "OCORRÊNCIA": "Ocorrência", "QUANTIDADE": "Quantidade", "SIGLA": "Sigla", "NORMALIZAÇÃO": "Normalização", "DESLIGAMENTO": "Desligamento", "OPERADOR": "Operador", "DESCRIÇÃO": "Descrição", "OS": "OS", "ATENDIMENTO LOOP": "Atendimento Loop", "ATENDIMENTO TERCEIROS": "Atendimento Terceiros", "PROTOCOLO": "Protocolo", "CLIENTE AVISADO": "Cliente Avisado"}
        df.rename(columns={c: rename[c.strip().upper()] for c in df.columns if c.strip().upper() in rename}, inplace=True)
        df.fillna("", inplace=True)
        
        for c in ["Normalização", "Desligamento", "Atendimento Loop", "Atendimento Terceiros", "Cliente Avisado"]:
            if c in df.columns: df[c] = pd.to_datetime(df[c], errors="coerce")
            
        if "Desligamento" in df.columns:
            df["Data"] = df["Desligamento"].dt.strftime("%Y-%m-%d")
            df["Hora"] = df["Desligamento"].dt.strftime("%H:%M:%S")
            df["Mês"] = df["Desligamento"].dt.strftime("%B").map(meses_traducao)
            df["Ano"] = df["Desligamento"].dt.year.fillna(0).astype(int)
            df["Dia"] = df["Desligamento"].dt.day.fillna(0).astype(int)
            df["ID_Unico"] = df["UG"].astype(str).str.upper() + "|" + df["Ativo"].astype(str).str.upper() + "|" + df["Ocorrência"].astype(str).str.upper() + "|" + df["Desligamento"].astype(str)
        return df
    except Exception as e:
        st.error(f"Erro: {e}")
        return pd.DataFrame()

if "cache_buster" not in st.session_state: st.session_state.cache_buster = int(pytime.time())
df_todos = carregar_dados_google_sheets(st.session_state.cache_buster)
if st.session_state.ui_phase != "ready": stop_loading()

# --- 2. KPIs SUPERIORES (ATIVAS) ---
c_deslig = 0
c_equip = 0
if not df_todos.empty:
    c_deslig = df_todos[(df_todos["Categoria"]=="DESLIGAMENTOS") & (pd.isna(df_todos["Normalização"]) | (df_todos["Normalização"]==""))].shape[0]
    c_equip = df_todos[(df_todos["Categoria"]=="EQUIPAMENTOS") & (pd.isna(df_todos["Normalização"]) | (df_todos["Normalização"]==""))].shape[0]

with st.container(border=True):
    st.markdown("<h1 style='margin:0; text-align:center;'>OCORRÊNCIAS ATIVAS</h1>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1: st.markdown(f"""<div class="kpi-card"><div class="kpi-label">USINAS DESLIGADAS NO MOMENTO</div><div class="kpi-value" style="color: #FF4B4B;">{c_deslig}</div></div>""", unsafe_allow_html=True)
    with c2: st.markdown(f"""<div class="kpi-card"><div class="kpi-label">EQUIPAMENTOS PARADOS NO MOMENTO</div><div class="kpi-value" style="color: #FF4B4B;">{c_equip}</div></div>""", unsafe_allow_html=True)

# --- 3. FILTROS ---
if "filtros_meses" not in st.session_state: st.session_state.filtros_meses = [meses_traducao[datetime.now().strftime("%B")]]
if "filtros_anos" not in st.session_state: st.session_state.filtros_anos = sorted([a for a in df_todos["Ano"].unique() if a!=0]) if not df_todos.empty else []
if "filtros_dias" not in st.session_state: st.session_state.filtros_dias = []
if "filtros_categorias" not in st.session_state: st.session_state.filtros_categorias = sorted(df_todos["Categoria"].unique()) if not df_todos.empty else []
for k in ["filtros_clientes", "filtros_ugs", "filtros_tipos", "filtros_ativos", "filtros_ocorrencias"]:
    if k not in st.session_state: st.session_state[k] = []

col_up_l, col_up_r = st.columns([0.2, 0.8])
with col_up_l:
    if st.button("Atualizar Dados"):
        st.cache_data.clear()
        start_loading()
        st.rerun()

st.markdown("---")
st.header("OCORRÊNCIAS FILTRADAS")

def _marcar(prefixo, itens, key, val, validos=None):
    validos = set(itens) if validos is None else set(validos)
    st.session_state[key] = [x for x in itens if (x in validos) and val]
    for x in itens: st.session_state[f"{prefixo}{x}"] = val and (x in validos)

if not df_todos.empty:
    st.markdown("#### Filtrar por categoria (planilha)")
    st.radio("Categoria:", ["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"], horizontal=True, label_visibility="collapsed", key="categoria_top", on_change=start_loading)
    
    st.subheader("Selecione o período desejado")
    anos_disp = sorted([a for a in df_todos["Ano"].unique() if a!=0])
    meses_disp = meses_cronologicos[:]
    
    ca, cm, cd = st.columns(3)
    with ca:
        with st.container(border=True):
            st.write("### Ano(s):")
            with st.expander("Expandir"):
                for a in anos_disp: st.checkbox(str(a), key=f"cb_ano_{a}", value=(a in st.session_state.filtros_anos))
            b1, b2 = st.columns(2)
            with b1: st.button("Sel. Todos", key="sa_ano", on_click=_marcar, args=("cb_ano_", anos_disp, "filtros_anos", True))
            with b2: st.button("Desmarcar", key="dm_ano", on_click=_marcar, args=("cb_ano_", anos_disp, "filtros_anos", False))
            st.session_state.filtros_anos = [a for a in anos_disp if st.session_state.get(f"cb_ano_{a}", False)]
            
    with cm:
        with st.container(border=True):
            st.write("### Mês(es):")
            with st.expander("Expandir"):
                for m in meses_disp: st.checkbox(m, key=f"cb_mes_{m}", value=(m in st.session_state.filtros_meses))
            b1, b2 = st.columns(2)
            with b1: st.button("Sel. Todos", key="sa_mes", on_click=_marcar, args=("cb_mes_", meses_disp, "filtros_meses", True))
            with b2: st.button("Desmarcar", key="dm_mes", on_click=_marcar, args=("cb_mes_", meses_disp, "filtros_meses", False))
            st.session_state.filtros_meses = [m for m in meses_disp if st.session_state.get(f"cb_mes_{m}", False)]

    dias_disp = list(range(1, 32))
    with cd:
        with st.container(border=True):
            st.write("### Dia(s):")
            with st.expander("Expandir"):
                dc = st.columns(7)
                for i, d in enumerate(dias_disp):
                    with dc[i%7]: st.checkbox(str(d), key=f"cb_dia_{d}", value=(d in st.session_state.filtros_dias))
            b1, b2 = st.columns(2)
            with b1: st.button("Sel. Todos", key="sa_dia", on_click=_marcar, args=("cb_dia_", dias_disp, "filtros_dias", True))
            with b2: st.button("Desmarcar", key="dm_dia", on_click=_marcar, args=("cb_dia_", dias_disp, "filtros_dias", False))
            st.session_state.filtros_dias = [d for d in dias_disp if st.session_state.get(f"cb_dia_{d}", False)]

# Filtros Adicionais
st.subheader("Filtros Adicionais")
r1c1, r1c2, r1c3 = st.columns(3)
r2c1, r2c2 = st.columns(2)

def render_filt(col, nome, col_df, key):
    with col:
        with st.container(border=True):
            st.write(nome)
            opts = sorted([x for x in options_from(df_todos[col_df]) if x!="-"]) if col_df!="Cliente" and col_df!="UG" else sorted([x for x in df_todos[col_df].unique() if x and x!="-"])
            st.session_state[key] = [x for x in st.session_state[key] if x in opts]
            b1, b2 = st.columns(2)
            with b1: 
                if st.button("Sel. Todos", key=f"sa_{key}"): st.session_state[key] = opts[:]
                # Bugfix for re-rendering:
                if st.session_state.get(f"sa_{key}"): st.rerun()
            with b2:
                if st.button("Limpar", key=f"dm_{key}"): st.session_state[key] = []
                if st.session_state.get(f"dm_{key}"): st.rerun()
            st.session_state[key] = st.multiselect("", opts, default=st.session_state[key], label_visibility="hidden", key=f"ms_{key}")

render_filt(r1c1, "Cliente", "Cliente", "filtros_clientes")
render_filt(r1c2, "UG", "UG", "filtros_ugs")
render_filt(r1c3, "Tipo", "Tipo de ocorrência", "filtros_tipos")
render_filt(r2c1, "Ativo", "Ativo", "filtros_ativos")
render_filt(r2c2, "Ocorrência", "Ocorrência", "filtros_ocorrencias")

# --- 4. APLICAÇÃO ---
m_ano = df_todos["Ano"].isin(st.session_state.filtros_anos) if st.session_state.filtros_anos else (df_todos["Ano"].notna() | True)
m_mes = df_todos["Mês"].isin(st.session_state.filtros_meses) if st.session_state.filtros_meses else (df_todos["Mês"].notna() | True)
m_dia = df_todos["Dia"].isin(st.session_state.filtros_dias) if st.session_state.filtros_dias else (df_todos["Dia"].notna() | True)
m_cat = df_todos["Categoria"].isin(st.session_state.filtros_categorias)
if st.session_state.categoria_top in ["DESLIGAMENTOS", "EQUIPAMENTOS"]: m_cat = m_cat & (df_todos["Categoria"] == st.session_state.categoria_top)

def mm(col, sel): return matches_any_canon(df_todos[col], sel)
m_final = m_ano & m_mes & m_dia & m_cat & mm("Cliente", st.session_state.filtros_clientes) & df_todos["UG"].isin(st.session_state.filtros_ugs) if st.session_state.filtros_ugs else True
if isinstance(m_final, bool): m_final = pd.Series(True, index=df_todos.index) # Fallback
else: m_final = m_final & mm("Tipo de ocorrência", st.session_state.filtros_tipos) & mm("Ativo", st.session_state.filtros_ativos) & mm("Ocorrência", st.session_state.filtros_ocorrencias)

df_filt = df_todos[m_final].copy()
df_desligadas = df_filt[pd.isna(df_filt["Normalização"]) | (df_filt["Normalização"]=="")].copy()

# KPIs Inferiores (Filtradas)
c_total = df_todos[pd.isna(df_todos["Normalização"]) | (df_todos["Normalização"]=="")].shape[0]
c_filt = len(df_desligadas)

k1, k2 = st.columns(2)
with k1: st.markdown(f"""<div class="kpi-card"><div class="kpi-label">Total no Banco de Dados Completo</div><div class="kpi-value" style="color: #FF4B4B;">{c_total}</div></div>""", unsafe_allow_html=True)
with k2: st.markdown(f"""<div class="kpi-card"><div class="kpi-label">Total com Filtro Selecionado</div><div class="kpi-value" style="color: #FF4B4B;">{c_filt}</div></div>""", unsafe_allow_html=True)

# --- 5. TABELA E CARDS (Vermelhos) ---
if not df_desligadas.empty:
    maskvalid = df_desligadas["Desligamento"].notna()
    df_desligadas.loc[maskvalid, "Tempo em Segundos"] = (datetime.now() - df_desligadas.loc[maskvalid, "Desligamento"]).dt.total_seconds().astype(int)
    df_desligadas.loc[~maskvalid, "Tempo em Segundos"] = 0
    
    st.markdown("---")
    st.write("Ordenar e Editar")
    sc1, sc2 = st.columns(2)
    with sc1: sort_col = st.selectbox("Ordenar por", ["Desligamento", "Tempo em Segundos", "UG", "Ativo"], index=0)
    with sc2: asc = st.radio("Ordem", ["Descendente", "Ascendente"]) == "Ascendente"
    
    map_sort = {"Desligamento": "Desligamento", "Tempo em Segundos": "Tempo em Segundos", "UG": "UG", "Ativo": "Ativo"}
    df_sorted = df_desligadas.sort_values(by=map_sort[sort_col], ascending=asc, na_position="last")
    
    df_sorted["Display"] = df_sorted["UG"].astype(str) + " | " + df_sorted["Ativo"].astype(str) + " | " + df_sorted["Nome Ativo"].astype(str) + " | " + df_sorted["Ocorrência"].astype(str) + " | " + df_sorted["Desligamento"].dt.strftime("%d/%m/%Y %H:%M").fillna("") + " | " + df_sorted["ID_Unico"].astype(str).str[-6:]
    st.session_state["df_lista_para_editar"] = df_sorted.copy()
    
    st.markdown("---")
    st.write("Editar uma Ocorrência")
    sel = st.selectbox("Selecione", df_sorted["Display"].tolist(), index=None, placeholder="Escolha...")
    if sel: st.session_state["id_unico_para_editar"] = df_sorted.loc[df_sorted["Display"]==sel, "ID_Unico"].iloc[0]
    if st.button("Editar Ocorrência Selecionada", disabled=not sel, use_container_width=True): st.switch_page("pages/3_Editar_Ocorrência.py")
    
    st.header("Lista (Tabela)")
    df_tab = df_sorted.copy().reset_index(drop=True)
    df_tab["Linha"] = df_tab.index + 1
    def fmt_tempo(row):
        t = row["Tempo em Segundos"]
        return f"{t//86400}d {(t%86400)//3600}h {(t%3600)//60}m"
    df_tab["Tempo de Desligamento"] = df_tab.apply(fmt_tempo, axis=1)
    st.dataframe(df_tab[["Linha", "Categoria", "Tempo de Desligamento", "UG", "Data", "Hora", "Tipo de ocorrência", "Ativo", "Ocorrência", "Operador", "Descrição", "OS"]], use_container_width=True)

    st.header("Detalhes (Cards)")
    rows = list(df_sorted.iterrows())
    ncols = 3
    for i in range(0, len(rows), ncols):
        cols = st.columns(ncols)
        for j in range(ncols):
            if i+j >= len(rows): break
            _, r = rows[i+j]
            with cols[j]:
                def fd(d): return d.strftime("%d/%m/%Y %H:%M") if pd.notna(d) else ""
                qtd_html = ""
                if r.get("Categoria") == "EQUIPAMENTOS" and float(r.get("Quantidade",0))>0:
                    qtd_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {int(float(r.get("Quantidade")))}</div>'
                
                st.markdown(f"""
                <div class="card-container">
                    <div class="card-title">{html.escape(str(r.get("UG","")))}</div>
                    <div class="card-item"><span class="card-label">Cliente:</span> {html.escape(str(r.get("Cliente","")))}</div>
                    <div class="card-item"><span class="card-label">Categoria:</span> {html.escape(str(r.get("Categoria","")))}</div>
                    <div class="card-item"><span class="card-label">Ativo:</span> {html.escape(str(r.get("Ativo","")))} | {html.escape(str(r.get("Nome Ativo","")))}</div>
                    <div class="card-item"><span class="card-label">Ocorrência:</span> {html.escape(str(r.get("Ocorrência","")))}</div>
                    <div class="card-item"><span class="card-label">Operador:</span> {html.escape(str(r.get("Operador","")))}</div>
                    {qtd_html}
                    <br>
                    <div class="card-item"><span class="card-label">Ocorrência:</span> {fd(r.get("Desligamento"))}</div>
                    <div class="card-item"><span class="card-label">Cliente Avisado:</span> {fd(r.get("Cliente Avisado"))}</div>
                    <div class="card-item"><span class="card-label">Loop:</span> {fd(r.get("Atendimento Loop"))}</div>
                    <div class="card-item"><span class="card-label">Terceiros:</span> {fd(r.get("Atendimento Terceiros"))}</div>
                    <br>
                    <div class="card-item"><span class="card-label">Descrição:</span> {html.escape(str(r.get("Descrição","")))}</div>
                    <div class="card-item"><span class="card-label">OS:</span> {html.escape(str(r.get("OS","")))}</div>
                </div>
                """, unsafe_allow_html=True)
else:
    st.info("Nenhuma ocorrência ativa encontrada.")