import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import html
import utils

# --- Configuração Inicial ---
st.set_page_config(layout="wide")
utils.init_overlay()

if 'cache_buster' not in st.session_state:
    st.session_state.cache_buster = int(pytime.time())

# CSS Personalizado
st.markdown("""
<style>
    .stButton > button {
        background-color: #28a745; color: white; font-weight: bold;
        border-radius: 5px; padding: 10px 20px; width: 100%; border: none;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); transition: background-color 0.3s;
    }
    .stButton > button:hover { background-color: #218838; }
    .kpi-card {
        background-color: #333333; padding: 20px; border-radius: 10px;
        text-align: center; margin-bottom: 20px;
    }
    .kpi-value { font-size: 3em; font-weight: bold; color: #FF4B4B; }
    .kpi-label { font-size: 1.2em; color: #FFFFFF; }
    .card-container {
        background-color: #FF4B4B; color: white; padding: 15px;
        border-radius: 8px; margin-bottom: 15px; box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        height: 100%;
    }
    .card-title {
        font-size: 1.5em; font-weight: bold; color: white;
        border-bottom: 1px solid rgba(255,255,255,0.5); padding-bottom: 5px; margin-bottom: 10px;
    }
    .card-item { margin-bottom: 5px; font-size: 1em; }
    .card-label { font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- Carregamento de Dados ---
# Executa overlay na primeira carga
if st.session_state.ui_phase == 'init':
    utils.overlay_on()
    st.rerun()

df_todos_dados = utils.carregar_dados_completos(st.session_state.cache_buster)

if st.session_state.ui_phase != 'ready':
    utils.overlay_off()

# --- KPIs do Topo (Ativos) ---
if not df_todos_dados.empty:
    count_deslig = df_todos_dados[
        (df_todos_dados['Categoria'] == utils.SHEET_DESLIGAMENTOS) &
        (pd.isna(df_todos_dados['Normalização']) | (df_todos_dados['Normalização'] == ''))
    ].shape[0]

    count_equip = df_todos_dados[
        (df_todos_dados['Categoria'] == utils.SHEET_EQUIPAMENTOS) &
        (pd.isna(df_todos_dados['Normalização']) | (df_todos_dados['Normalização'] == ''))
    ].shape[0]
else:
    count_deslig = 0
    count_equip = 0

with st.container(border=True):
    st.markdown("<h1 style='margin:0'>OCORRÊNCIAS ATIVAS</h1>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"<div class='kpi-card'><div class='kpi-label'>USINAS DESLIGADAS</div><div class='kpi-value'>{count_deslig}</div></div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div class='kpi-card'><div class='kpi-label'>EQUIPAMENTOS PARADOS</div><div class='kpi-value'>{count_equip}</div></div>", unsafe_allow_html=True)

# --- Inicialização de Estados de Filtro ---
if 'filtros_meses' not in st.session_state:
    st.session_state.filtros_meses = [utils.MESES_TRADUCAO[datetime.now().strftime('%B')]]
if 'filtros_anos' not in st.session_state:
    if not df_todos_dados.empty:
        anos = sorted([a for a in df_todos_dados['Ano'].unique() if a != 0])
        st.session_state.filtros_anos = anos
    else:
        st.session_state.filtros_anos = []
if 'filtros_dias' not in st.session_state:
    st.session_state.filtros_dias = [] # Inicia vazio (todos) ou lógica específica
if 'filtros_clientes' not in st.session_state: st.session_state.filtros_clientes = []
if 'filtros_ugs' not in st.session_state: st.session_state.filtros_ugs = []
if 'filtros_tipos' not in st.session_state: st.session_state.filtros_tipos = []
if 'filtros_ativos' not in st.session_state: st.session_state.filtros_ativos = []
if 'filtros_ocorrencias' not in st.session_state: st.session_state.filtros_ocorrencias = []
if 'categoria_top' not in st.session_state: st.session_state.categoria_top = "Ambas"

st.header('OCORRÊNCIAS FILTRADAS')

# KPI Esquerdo (Banco Completo)
col_kpi1, col_kpi2 = st.columns(2)
with col_kpi1:
    if not df_todos_dados.empty:
        total_db = df_todos_dados[pd.isna(df_todos_dados['Normalização']) | (df_todos_dados['Normalização'] == '')].shape[0]
    else:
        total_db = 0
    st.markdown(f"<div class='kpi-card'><div class='kpi-label'>Total no Banco (Abertas)</div><div class='kpi-value'>{total_db}</div></div>", unsafe_allow_html=True)

# Botão Atualizar
col_btn, _ = st.columns([0.2, 0.8])
with col_btn:
    if st.button("Atualizar Dados"):
        st.cache_data.clear()
        utils.overlay_on()
        st.rerun()

# --- Lógica de Filtragem ---
if not df_todos_dados.empty:
    st.radio("Categoria:", ["Ambas", utils.SHEET_DESLIGAMENTOS, utils.SHEET_EQUIPAMENTOS], 
             horizontal=True, key="categoria_top", on_change=utils.overlay_on)

    # Filtros de Data
    st.subheader("Filtros de Período")
    c_ano, c_mes, c_dia = st.columns(3)
    
    with c_ano:
        anos_opts = sorted([a for a in df_todos_dados['Ano'].unique() if a != 0])
        st.session_state.filtros_anos = st.multiselect("Anos", options=anos_opts, default=st.session_state.filtros_anos)

    with c_mes:
        meses_opts = utils.MESES_CRONOLOGICOS
        st.session_state.filtros_meses = st.multiselect("Meses", options=meses_opts, default=st.session_state.filtros_meses)
        
    with c_dia:
        # Dias dinâmicos baseados no ano/mes selecionado
        mask_ano = df_todos_dados['Ano'].isin(st.session_state.filtros_anos) if st.session_state.filtros_anos else pd.Series([True]*len(df_todos_dados))
        mask_mes = df_todos_dados['Mês'].isin(st.session_state.filtros_meses) if st.session_state.filtros_meses else pd.Series([True]*len(df_todos_dados))
        dias_disp = sorted(df_todos_dados[mask_ano & mask_mes]['Dia'].dropna().unique().astype(int).tolist())
        dias_disp = [d for d in dias_disp if d > 0]
        st.session_state.filtros_dias = st.multiselect("Dias", options=dias_disp, default=st.session_state.filtros_dias)

    # Filtros Adicionais
    st.subheader("Filtros Adicionais")
    c1, c2, c3, c4, c5 = st.columns(5)
    
    with c1:
        opts = utils.options_from(df_todos_dados['Cliente'])
        st.session_state.filtros_clientes = st.multiselect("Cliente", opts, default=st.session_state.filtros_clientes)
    
    with c2:
        # UGs filtradas por cliente
        df_tmp = df_todos_dados
        if st.session_state.filtros_clientes:
            df_tmp = df_tmp[utils.matches_any_canon(df_tmp['Cliente'], st.session_state.filtros_clientes)]
        opts_ug = sorted(utils.options_from(df_tmp['UG']))
        st.session_state.filtros_ugs = st.multiselect("UG", opts_ug, default=st.session_state.filtros_ugs)
        
    with c3:
        opts_tipo = utils.options_from(df_todos_dados['Tipo de ocorrência'])
        st.session_state.filtros_tipos = st.multiselect("Tipo", opts_tipo, default=st.session_state.filtros_tipos)
        
    with c4:
        opts_ativo = utils.options_from(df_todos_dados['Ativo'])
        st.session_state.filtros_ativos = st.multiselect("Ativo", opts_ativo, default=st.session_state.filtros_ativos)
        
    with c5:
        opts_ocorr = utils.options_from(df_todos_dados['Ocorrência'])
        st.session_state.filtros_ocorrencias = st.multiselect("Ocorrência", opts_ocorr, default=st.session_state.filtros_ocorrencias)

    # --- Aplicação das Máscaras ---
    m_cat = pd.Series([True]*len(df_todos_dados))
    if st.session_state.categoria_top != "Ambas":
        m_cat = df_todos_dados['Categoria'] == st.session_state.categoria_top
        
    m_ano = df_todos_dados['Ano'].isin(st.session_state.filtros_anos) if st.session_state.filtros_anos else pd.Series([True]*len(df_todos_dados))
    m_mes = df_todos_dados['Mês'].isin(st.session_state.filtros_meses) if st.session_state.filtros_meses else pd.Series([True]*len(df_todos_dados))
    m_dia = df_todos_dados['Dia'].isin(st.session_state.filtros_dias) if st.session_state.filtros_dias else pd.Series([True]*len(df_todos_dados))
    
    m_cli = utils.matches_any_canon(df_todos_dados['Cliente'], st.session_state.filtros_clientes)
    m_ug = utils.matches_any_canon(df_todos_dados['UG'], st.session_state.filtros_ugs)
    m_tip = utils.matches_any_canon(df_todos_dados['Tipo de ocorrência'], st.session_state.filtros_tipos)
    m_atv = utils.matches_any_canon(df_todos_dados['Ativo'], st.session_state.filtros_ativos)
    m_ocr = utils.matches_any_canon(df_todos_dados['Ocorrência'], st.session_state.filtros_ocorrencias)
    
    df_filtrado = df_todos_dados[m_cat & m_ano & m_mes & m_dia & m_cli & m_ug & m_tip & m_atv & m_ocr].copy()
    
    # Apenas Abertas
    df_abertas = df_filtrado[pd.isna(df_filtrado['Normalização']) | (df_filtrado['Normalização'] == '')].copy()
    
    # KPI Direito
    with col_kpi2:
         st.markdown(f"<div class='kpi-card'><div class='kpi-label'>Total Filtrado (Abertas)</div><div class='kpi-value'>{len(df_abertas)}</div></div>", unsafe_allow_html=True)
         
    # --- Tabela e Ordenação ---
    if not df_abertas.empty:
        # Coluna Tempo
        mask_valid = df_abertas['Desligamento'].notna()
        df_abertas.loc[mask_valid, 'Tempo em Segundos'] = (
            (datetime.now() - df_abertas.loc[mask_valid, 'Desligamento']).dt.total_seconds().astype(int)
        )
        df_abertas.loc[~mask_valid, 'Tempo em Segundos'] = 0
        
        st.markdown("---")
        st.write("### Ordenar")
        c_sort1, c_sort2 = st.columns(2)
        with c_sort1:
            sort_col = st.selectbox("Ordenar por:", ["Desligamento", "Tempo em Segundos", "UG"], index=0)
        with c_sort2:
            sort_asc = st.radio("Ordem:", ["Descendente", "Ascendente"], horizontal=True) == "Ascendente"
            
        df_sorted = df_abertas.sort_values(by=sort_col, ascending=sort_asc)
        
        # Display para Tabela
        def formatar_tempo(s):
            d = s // 86400
            h = (s % 86400) // 3600
            m = (s % 3600) // 60
            return f"{int(d)}d {int(h)}h {int(m)}m"
            
        df_show = df_sorted.copy()
        df_show['Tempo'] = df_show['Tempo em Segundos'].apply(formatar_tempo)
        
        st.dataframe(df_show[['Categoria', 'Tempo', 'UG', 'Data', 'Hora', 'Tipo de ocorrência', 'Ocorrência', 'Descrição']], use_container_width=True)

        # --- Edição ---
        st.markdown("---")
        st.write("### Editar Ocorrência")
        
        df_sorted['Display'] = (
            df_sorted['UG'].astype(str) + " | " + 
            df_sorted['Ativo'].astype(str) + " | " + 
            df_sorted['Ocorrência'].astype(str) + " | " + 
            df_sorted['Desligamento'].dt.strftime('%d/%m %H:%M').fillna('')
        )
        
        opts = df_sorted['Display'].tolist()
        sel = st.selectbox("Selecione para editar:", options=opts, index=None, placeholder="Escolha uma ocorrência...")
        
        if sel:
             id_unico = df_sorted.loc[df_sorted['Display'] == sel, 'ID_Unico'].values[0]
             st.session_state['id_unico_para_editar'] = id_unico
        
        if st.button("📝 Editar Selecionada", disabled=not bool(sel)):
             st.switch_page("pages/3_Editar_Ocorrência.py")

        # --- Cards ---
        st.markdown("---")
        num_cols = 4
        rows = list(df_sorted.iterrows())
        
        for i in range(0, len(rows), num_cols):
            cols = st.columns(num_cols)
            for j in range(num_cols):
                if i + j < len(rows):
                    _, row = rows[i+j]
                    with cols[j]:
                        d_ocor = row['Desligamento'].strftime('%d/%m %H:%M') if pd.notna(row['Desligamento']) else ""
                        st.markdown(f"""
                        <div class="card-container">
                            <div class="card-title">{html.escape(str(row.get('UG')))}</div>
                            <div class="card-item"><b>Cliente:</b> {html.escape(str(row.get('Cliente')))}</div>
                            <div class="card-item"><b>Ocorrência:</b> {html.escape(str(row.get('Ocorrência')))}</div>
                            <div class="card-item"><b>Data:</b> {d_ocor}</div>
                            <div class="card-item"><b>Desc:</b> {html.escape(str(row.get('Descrição')))}</div>
                        </div>
                        """, unsafe_allow_html=True)
    else:
        st.info("Nenhuma ocorrência encontrada para os filtros.")
else:
    st.warning("Não foi possível carregar os dados.")