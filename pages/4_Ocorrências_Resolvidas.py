import streamlit as st
import pandas as pd
import html
import utils
from datetime import datetime
import time as pytime

st.set_page_config(layout="wide")
utils.init_overlay()

if 'cache_buster' not in st.session_state:
    st.session_state.cache_buster = int(pytime.time())

# CSS Cards
st.markdown("""
<style>
    .kpi-card { background-color: #333333; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px; }
    .kpi-value { font-size: 3em; font-weight: bold; color: #4b4eff; }
    .kpi-label { font-size: 1.2em; color: #FFFFFF; }
    .card-container { background-color: #089641; color: white; padding: 15px; border-radius: 8px; margin-bottom: 15px; height: 100%; box-shadow: 0 4px 8px rgba(0,0,0,0.2); }
    .card-title { font-size: 1.5em; font-weight: bold; border-bottom: 1px solid rgba(255,255,255,0.5); padding-bottom: 5px; margin-bottom: 10px; }
    .card-item { margin-bottom: 5px; }
</style>
""", unsafe_allow_html=True)

# Topo
st.title("Histórico de Ocorrências Resolvidas")

if st.button("Atualizar Dados"):
    st.cache_data.clear()
    utils.overlay_on()
    st.rerun()

# Carregar Dados
if st.session_state.ui_phase == 'init': utils.overlay_on(); st.rerun()
df_todos = utils.carregar_dados_completos(st.session_state.cache_buster)
if st.session_state.ui_phase != 'ready': utils.overlay_off()

# Filtro Base: Apenas Resolvidas
df_resolvidas_base = pd.DataFrame()
if not df_todos.empty:
    df_resolvidas_base = df_todos[df_todos['Normalização'].notna() & (df_todos['Normalização'] != '')].copy()

# KPIs Topo
count_res_deslig = df_resolvidas_base[df_resolvidas_base['Categoria'] == utils.SHEET_DESLIGAMENTOS].shape[0] if not df_resolvidas_base.empty else 0
count_res_equip = df_resolvidas_base[df_resolvidas_base['Categoria'] == utils.SHEET_EQUIPAMENTOS].shape[0] if not df_resolvidas_base.empty else 0

col_a, col_b = st.columns(2)
col_a.markdown(f"<div class='kpi-card'><div class='kpi-label'>DESLIGAMENTOS RESOLVIDOS</div><div class='kpi-value'>{count_res_deslig}</div></div>", unsafe_allow_html=True)
col_b.markdown(f"<div class='kpi-card'><div class='kpi-label'>EQUIPAMENTOS RESOLVIDOS</div><div class='kpi-value'>{count_res_equip}</div></div>", unsafe_allow_html=True)

# --- Filtros (Lógica similar à Home) ---
# Inicialização simples
if 'hist_anos' not in st.session_state: st.session_state.hist_anos = []
if 'hist_meses' not in st.session_state: st.session_state.hist_meses = []

st.subheader("Filtrar Histórico")
c1, c2 = st.columns(2)
with c1:
    opts_ano = sorted([a for a in df_resolvidas_base['Ano'].unique() if a != 0]) if not df_resolvidas_base.empty else []
    st.session_state.hist_anos = st.multiselect("Ano", opts_ano, default=st.session_state.hist_anos)
with c2:
    st.session_state.hist_meses = st.multiselect("Mês", utils.MESES_CRONOLOGICOS, default=st.session_state.hist_meses)

# Aplicação Filtro
df_filt = df_resolvidas_base
if not df_filt.empty:
    if st.session_state.hist_anos:
        df_filt = df_filt[df_filt['Ano'].isin(st.session_state.hist_anos)]
    if st.session_state.hist_meses:
        df_filt = df_filt[df_filt['Mês'].isin(st.session_state.hist_meses)]

st.metric("Total Visualizado", len(df_filt))

# --- Tabela ---
st.write("### Lista")
if not df_filt.empty:
    df_show = df_filt[['UG', 'Categoria', 'Data', 'Hora', 'Normalização', 'Ocorrência', 'Descrição']].copy()
    st.dataframe(df_show, use_container_width=True)

    # --- Cards ---
    st.write("### Detalhes")
    num_cols = 4
    rows = list(df_filt.iterrows())
    
    for i in range(0, len(rows), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j < len(rows):
                _, row = rows[i+j]
                with cols[j]:
                    norm_str = row['Normalização'].strftime('%d/%m/%Y %H:%M') if pd.notna(row['Normalização']) else ""
                    st.markdown(f"""
                    <div class="card-container">
                        <div class="card-title">{html.escape(str(row.get('UG')))}</div>
                        <div class="card-item"><b>Ocorrência:</b> {html.escape(str(row.get('Ocorrência')))}</div>
                        <div class="card-item"><b>Normalização:</b> {norm_str}</div>
                        <div class="card-item"><b>Desc:</b> {html.escape(str(row.get('Descrição')))}</div>
                    </div>
                    """, unsafe_allow_html=True)
else:
    st.info("Nenhum histórico para os filtros selecionados.")