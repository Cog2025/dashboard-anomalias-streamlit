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

# CSS Cards (Restaurado)
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

# --- Tabela e Cards ---
if not df_filt.empty:
    # Tabela Simples
    st.write("### Lista")
    st.dataframe(df_filt[['UG', 'Categoria', 'Data', 'Hora', 'Normalização', 'Ocorrência', 'Descrição']], use_container_width=True)

    # Cards Detalhados (Restaurado para coincidir com a página principal)
    st.write("### Detalhes por Ocorrência (Cards)")
    num_cols = 4
    rows = list(df_filt.iterrows())
    
    def format_dt(dt_obj):
        if pd.notna(dt_obj):
            return dt_obj.strftime('%d/%m/%Y'), dt_obj.strftime('%H:%M')
        return '', ''
    
    for i in range(0, len(rows), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j < len(rows):
                index, row = rows[i+j]
                with cols[j]:
                    cliente   = html.escape(str(row.get("Cliente", "")))
                    categoria = html.escape(str(row.get("Categoria", "")))
                    ug        = html.escape(str(row.get("UG", "N/A")))
                    tipo      = html.escape(str(row.get("Tipo de ocorrência", "")))
                    ativo     = html.escape(str(row.get("Ativo", "")))
                    nome_ativo= html.escape(str(row.get("Nome Ativo", "")))
                    ocorr     = html.escape(str(row.get("Ocorrência", "")))
                    operador  = html.escape(str(row.get("Operador", "")))
                    descricao = html.escape(str(row.get("Descrição", ""))).replace('\n','<br>')
                    protocolo = html.escape(str(row.get("Protocolo", "")))
                    osv       = html.escape(str(row.get("OS", "")))

                    d_des, h_des = format_dt(row.get('Desligamento'))
                    d_norm, h_norm = format_dt(row.get('Normalização'))
                    d_ca, h_ca = format_dt(row.get('Cliente Avisado'))
                    d_loop, h_loop = format_dt(row.get('Atendimento Loop'))
                    d_terc, h_terc = format_dt(row.get('Atendimento Terceiros'))

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
    st.info("Nenhum histórico para os filtros selecionados.")