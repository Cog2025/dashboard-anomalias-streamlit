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

# CSS Cards (Compatível com principal)
st.markdown("""
<style>
    .kpi-card { background-color: #333333; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px; }
    .kpi-value { font-size: 3em; font-weight: bold; color: #4b4eff; }
    .kpi-label { font-size: 1.2em; color: #FFFFFF; }
    
    .card-container { background-color: #089641; color: white; padding: 15px; border-radius: 8px; margin-bottom: 15px; height: 100%; box-shadow: 0 4px 8px rgba(0,0,0,0.2); }
    .card-title { font-size: 1.5em; font-weight: bold; border-bottom: 1px solid rgba(255,255,255,0.5); padding-bottom: 5px; margin-bottom: 10px; }
    .card-item { margin-bottom: 5px; font-size: 1em; }
    .card-label { font-weight: bold; }
</style>
""", unsafe_allow_html=True)

st.title("Histórico de Ocorrências Resolvidas")

if st.button("Atualizar Dados"):
    st.cache_data.clear()
    utils.overlay_on()
    st.rerun()

if st.session_state.ui_phase == 'init': utils.overlay_on(); st.rerun()
df_todos = utils.carregar_dados_completos(st.session_state.cache_buster)
if st.session_state.ui_phase != 'ready': utils.overlay_off()

# Filtra apenas resolvidas
df_resolvidas_base = pd.DataFrame()
if not df_todos.empty:
    df_resolvidas_base = df_todos[df_todos['Normalização'].notna() & (df_todos['Normalização'] != '')].copy()

# KPIs Topo
count_res_deslig = df_resolvidas_base[df_resolvidas_base['Categoria'] == utils.SHEET_DESLIGAMENTOS].shape[0] if not df_resolvidas_base.empty else 0
count_res_equip = df_resolvidas_base[df_resolvidas_base['Categoria'] == utils.SHEET_EQUIPAMENTOS].shape[0] if not df_resolvidas_base.empty else 0

col_a, col_b = st.columns(2)
col_a.markdown(f"<div class='kpi-card'><div class='kpi-label'>DESLIGAMENTOS RESOLVIDOS</div><div class='kpi-value'>{count_res_deslig}</div></div>", unsafe_allow_html=True)
col_b.markdown(f"<div class='kpi-card'><div class='kpi-label'>EQUIPAMENTOS RESOLVIDOS</div><div class='kpi-value'>{count_res_equip}</div></div>", unsafe_allow_html=True)

# Filtros
if 'hist_anos' not in st.session_state: st.session_state.hist_anos = []
if 'hist_meses' not in st.session_state: st.session_state.hist_meses = []

st.subheader("Filtrar Histórico")
c1, c2 = st.columns(2)
with c1:
    opts_ano = sorted([a for a in df_resolvidas_base['Ano'].unique() if a != 0]) if not df_resolvidas_base.empty else []
    st.session_state.hist_anos = st.multiselect("Ano", opts_ano, default=st.session_state.hist_anos)
with c2:
    st.session_state.hist_meses = st.multiselect("Mês", utils.MESES_CRONOLOGICOS, default=st.session_state.hist_meses)

# Aplicação
df_filt = df_resolvidas_base
if not df_filt.empty:
    if st.session_state.hist_anos:
        df_filt = df_filt[df_filt['Ano'].isin(st.session_state.hist_anos)]
    if st.session_state.hist_meses:
        df_filt = df_filt[df_filt['Mês'].isin(st.session_state.hist_meses)]

st.metric("Total Visualizado", len(df_filt))

# Lista e Cards Detalhados
st.write("### Lista")
if not df_filt.empty:
    st.dataframe(df_filt[['UG', 'Categoria', 'Data', 'Hora', 'Normalização', 'Ocorrência', 'Descrição']], use_container_width=True)

    st.write("### Detalhes por Ocorrência (Cards)")
    num_cols = 4
    rows = list(df_filt.iterrows())
    
    def fmt_dt(dt):
        if pd.notna(dt): return dt.strftime('%d/%m/%Y'), dt.strftime('%H:%M')
        return '', ''

    for i in range(0, len(rows), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j < len(rows):
                _, row = rows[i+j]
                with cols[j]:
                    # Recupera todos os dados com tratamento de erro
                    cli = html.escape(str(row.get("Cliente", "")))
                    cat = html.escape(str(row.get("Categoria", "")))
                    ug = html.escape(str(row.get("UG", "N/A")))
                    tipo = html.escape(str(row.get("Tipo de ocorrência", "")))
                    ativo = html.escape(str(row.get("Ativo", "")))
                    nome = html.escape(str(row.get("Nome Ativo", "")))
                    ocr = html.escape(str(row.get("Ocorrência", "")))
                    oper = html.escape(str(row.get("Operador", "")))
                    desc = html.escape(str(row.get("Descrição", ""))).replace('\n', '<br>')
                    prot = html.escape(str(row.get("Protocolo", "")))
                    osv = html.escape(str(row.get("OS", "")))
                    
                    d_des, h_des = fmt_dt(row.get('Desligamento'))
                    d_norm, h_norm = fmt_dt(row.get('Normalização'))
                    d_ca, h_ca = fmt_dt(row.get('Cliente Avisado'))
                    d_loop, h_loop = fmt_dt(row.get('Atendimento Loop'))
                    d_terc, h_terc = fmt_dt(row.get('Atendimento Terceiros'))

                    st.markdown(f"""
                    <div class="card-container">
                      <div class="card-title">{ug}</div>
                      <div class="card-item"><span class="card-label">Cliente:</span> {cli}</div>
                      <div class="card-item"><span class="card-label">Categoria:</span> {cat}</div>
                      <div class="card-item"><span class="card-label">Tipo:</span> {tipo}</div>
                      <div class="card-item"><span class="card-label">Ativo:</span> {ativo}</div>
                      <div class="card-item"><span class="card-label">Nome:</span> {nome}</div>
                      <div class="card-item"><span class="card-label">Ocorrência:</span> {ocr}</div>
                      <div class="card-item"><span class="card-label">Operador:</span> {oper}</div>
                      <br>
                      <div class="card-item"><span class="card-label">Desligamento:</span> {d_des} {h_des}</div>
                      <div class="card-item"><span class="card-label">Normalização:</span> {d_norm} {h_norm}</div>
                      <div class="card-item"><span class="card-label">Aviso:</span> {d_ca} {h_ca}</div>
                      <div class="card-item"><span class="card-label">Loop:</span> {d_loop} {h_loop}</div>
                      <div class="card-item"><span class="card-label">Terc.:</span> {d_terc} {h_terc}</div>
                      <br>
                      <div class="card-item"><span class="card-label">Desc:</span> {desc}</div>
                      <div class="card-item"><span class="card-label">Prot:</span> {prot}</div>
                      <div class="card-item"><span class="card-label">OS:</span> {osv}</div>
                    </div>
                    """, unsafe_allow_html=True)
else:
    st.info("Nenhum histórico para os filtros selecionados.")