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

# CSS Personalizado (Mantendo o estilo dos cards e botões)
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
    
    /* CSS DETALHADO DOS CARDS (Restaurado) */
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
if st.session_state.ui_phase == 'init':
    utils.overlay_on()
    st.rerun()

df_todos_dados = utils.carregar_dados_completos(st.session_state.cache_buster)

if st.session_state.ui_phase != 'ready':
    utils.overlay_off()

# --- KPIs do Topo ---
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
    # Lógica original: calcular dias disponíveis com base no ano/mês inicial
    if not df_todos_dados.empty:
        mask = (df_todos_dados['Ano'].isin(st.session_state.filtros_anos)) & \
               (df_todos_dados['Mês'].isin(st.session_state.filtros_meses))
        dias = sorted(df_todos_dados[mask]['Dia'].unique().astype(int).tolist())
        st.session_state.filtros_dias = [d for d in dias if d > 0]
    else:
        st.session_state.filtros_dias = []

# Inicializa outros filtros vazios se não existirem
for k in ['filtros_clientes', 'filtros_ugs', 'filtros_tipos', 'filtros_ativos', 'filtros_ocorrencias']:
    if k not in st.session_state: st.session_state[k] = []
if 'categoria_top' not in st.session_state: st.session_state.categoria_top = "Ambas"

st.header('OCORRÊNCIAS FILTRADAS')

# KPI Esquerdo (Banco Completo)
col_kpi1, col_kpi2 = st.columns(2)
with col_kpi1:
    total_db = 0
    if not df_todos_dados.empty:
        total_db = df_todos_dados[pd.isna(df_todos_dados['Normalização']) | (df_todos_dados['Normalização'] == '')].shape[0]
    st.markdown(f"<div class='kpi-card'><div class='kpi-label'>Total no Banco (Abertas)</div><div class='kpi-value'>{total_db}</div></div>", unsafe_allow_html=True)

# Botão Atualizar
col_btn, _ = st.columns([0.2, 0.8])
with col_btn:
    if st.button("Atualizar Dados"):
        st.cache_data.clear()
        utils.overlay_on()
        st.rerun()

# --- LÓGICA DE FILTROS COM BOTÕES (RESTAURADA) ---
def _marcar(prefixo_key: str, itens: list, filtro_key: str, marcar_todos: bool):
    """Função auxiliar para marcar/desmarcar todos"""
    st.session_state[filtro_key] = list(itens) if marcar_todos else []
    utils.overlay_on() # Ativa loading visualmente

if not df_todos_dados.empty:
    st.markdown("#### Filtrar por categoria (planilha)")
    st.radio("Categoria:", ["Ambas", utils.SHEET_DESLIGAMENTOS, utils.SHEET_EQUIPAMENTOS], 
             horizontal=True, key="categoria_top", on_change=utils.overlay_on)

    # --- FILTROS DE PERÍODO (Ano, Mês, Dia) ---
    st.subheader("Selecione o período desejado")
    c_ano, c_mes, c_dia = st.columns(3)
    
    # 1. ANOS
    with c_ano:
        with st.container(border=True):
            st.write("### Ano(s):")
            anos_opts = sorted([a for a in df_todos_dados['Ano'].unique() if a != 0])
            with st.expander("Expandir anos"):
                # Checkboxes individuais poderiam ser recriados aqui, mas o multiselect abaixo resolve
                pass
            
            c_b1, c_b2 = st.columns(2)
            c_b1.button('Sel. Todos', key='sel_ano', use_container_width=True, on_click=_marcar, args=('', anos_opts, 'filtros_anos', True))
            c_b2.button('Desmarcar', key='des_ano', use_container_width=True, on_click=_marcar, args=('', anos_opts, 'filtros_anos', False))
            
            st.session_state.filtros_anos = st.multiselect(" ", anos_opts, default=st.session_state.filtros_anos, label_visibility='collapsed')

    # 2. MESES
    with c_mes:
        with st.container(border=True):
            st.write("### Mês(es):")
            meses_opts = utils.MESES_CRONOLOGICOS
            with st.expander("Expandir meses"):
                pass
            
            c_b1, c_b2 = st.columns(2)
            c_b1.button('Sel. Todos', key='sel_mes', use_container_width=True, on_click=_marcar, args=('', meses_opts, 'filtros_meses', True))
            c_b2.button('Desmarcar', key='des_mes', use_container_width=True, on_click=_marcar, args=('', meses_opts, 'filtros_meses', False))
            
            st.session_state.filtros_meses = st.multiselect(" ", meses_opts, default=st.session_state.filtros_meses, label_visibility='collapsed')

    # 3. DIAS (Dinâmico)
    with c_dia:
        with st.container(border=True):
            st.write("### Dia(s):")
            # Calcula dias disponíveis
            mask_ano = df_todos_dados['Ano'].isin(st.session_state.filtros_anos) if st.session_state.filtros_anos else pd.Series([True]*len(df_todos_dados))
            mask_mes = df_todos_dados['Mês'].isin(st.session_state.filtros_meses) if st.session_state.filtros_meses else pd.Series([True]*len(df_todos_dados))
            dias_disp = sorted(df_todos_dados[mask_ano & mask_mes]['Dia'].dropna().unique().astype(int).tolist())
            dias_disp = [d for d in dias_disp if d > 0]
            
            with st.expander("Expandir dias"):
                pass
            
            c_b1, c_b2 = st.columns(2)
            c_b1.button('Sel. Todos', key='sel_dia', use_container_width=True, on_click=_marcar, args=('', dias_disp, 'filtros_dias', True))
            c_b2.button('Desmarcar', key='des_dia', use_container_width=True, on_click=_marcar, args=('', dias_disp, 'filtros_dias', False))
            
            # Limpa seleção se dias não existirem mais
            st.session_state.filtros_dias = [d for d in st.session_state.filtros_dias if d in dias_disp]
            st.session_state.filtros_dias = st.multiselect(" ", dias_disp, default=st.session_state.filtros_dias, label_visibility='collapsed')

    # --- FILTROS ADICIONAIS (Com botões) ---
    st.subheader("Filtros Adicionais")
    
    # Helper para renderizar bloco de filtro com botões
    def render_filter_block(label, col_name, state_key):
        with st.container(border=True):
            st.write(f"{label}:")
            # Filtra opções baseado nos filtros anteriores? O original usava "matches_any_canon" no dataset base
            # Para simplicidade e performance, usaremos as opções totais ou levemente filtradas
            opts = sorted(utils.options_from(df_todos_dados[col_name]))
            
            c_a, c_b = st.columns(2)
            c_a.button('Sel. Todos', key=f's_{state_key}', use_container_width=True, on_click=_marcar, args=('', opts, state_key, True))
            c_b.button('Desmarcar', key=f'd_{state_key}', use_container_width=True, on_click=_marcar, args=('', opts, state_key, False))
            
            # Garante que o default esteja nas opções
            valid_defaults = [x for x in st.session_state[state_key] if x in opts]
            st.session_state[state_key] = st.multiselect(" ", opts, default=valid_defaults, label_visibility='collapsed', key=f'ms_{state_key}')

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: render_filter_block("Cliente", 'Cliente', 'filtros_clientes')
    with c2: render_filter_block("UG", 'UG', 'filtros_ugs')
    with c3: render_filter_block("Tipo", 'Tipo de ocorrência', 'filtros_tipos')
    with c4: render_filter_block("Ativo", 'Ativo', 'filtros_ativos')
    with c5: render_filter_block("Ocorrência", 'Ocorrência', 'filtros_ocorrencias')

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
        st.write("### Ordenar e Editar")
        c_sort1, c_sort2 = st.columns(2)
        with c_sort1:
            sort_col = st.selectbox("Ordenar por:", ["Desligamento", "Tempo em Segundos", "UG"], index=0)
        with c_sort2:
            sort_asc = st.radio("Ordem:", ["Descendente", "Ascendente"], horizontal=True) == "Ascendente"
            
        df_sorted = df_abertas.sort_values(by=sort_col, ascending=sort_asc)
        
        # Display para Edição (ID composto)
        df_sorted['Display'] = (
            df_sorted['UG'].astype(str) + " | " + 
            df_sorted['Ativo'].astype(str) + " | " + 
            df_sorted['Ocorrência'].astype(str) + " | " + 
            df_sorted['Desligamento'].dt.strftime('%d/%m %H:%M').fillna('')
        )
        
        # Seletor de Edição
        opts = df_sorted['Display'].tolist()
        sel = st.selectbox("Selecione para editar:", options=opts, index=None, placeholder="Escolha uma ocorrência...")
        
        if sel:
             id_unico = df_sorted.loc[df_sorted['Display'] == sel, 'ID_Unico'].values[0]
             st.session_state['id_unico_para_editar'] = id_unico
        
        if st.button("📝 Editar Ocorrência Selecionada", disabled=not bool(sel)):
             st.switch_page("pages/3_Editar_Ocorrência.py")

        # Tabela
        st.header("Lista de Ocorrências (Tabela)")
        def formatar_tempo(s):
            d = s // 86400
            h = (s % 86400) // 3600
            m = (s % 3600) // 60
            return f"{int(d)}d {int(h)}h {int(m)}m"
            
        df_show = df_sorted.copy()
        df_show['Tempo'] = df_show['Tempo em Segundos'].apply(formatar_tempo)
        st.dataframe(df_show[['Categoria', 'Tempo', 'UG', 'Data', 'Hora', 'Tipo de ocorrência', 'Ocorrência', 'Descrição']], use_container_width=True)

        # --- CARDS DETALHADOS (Restaurados) ---
        st.header("Detalhes por Ocorrência (Cards)")
        num_cols = 4
        rows = list(df_sorted.iterrows())
        
        def format_datetime_card(dt_obj):
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
                        ug = html.escape(str(row.get("UG", "N/A")))
                        tipo_ocorrencia = html.escape(str(row.get("Tipo de ocorrência", "")))
                        ativo = html.escape(str(row.get("Ativo", "")))
                        nome_ativo = html.escape(str(row.get("Nome Ativo", "")))
                        ocorrencia = html.escape(str(row.get("Ocorrência", "")))
                        operador = html.escape(str(row.get("Operador", "")))
                        descricao = html.escape(str(row.get("Descrição", ""))).replace('\n', '<br>')
                        protocolo = html.escape(str(row.get("Protocolo", "")))
                        os_val = html.escape(str(row.get("OS", "")))

                        data_ocor, hora_ocor = format_datetime_card(row.get('Desligamento'))
                        data_ca, hora_ca = format_datetime_card(row.get('Cliente Avisado'))
                        data_loop, hora_loop = format_datetime_card(row.get('Atendimento Loop'))
                        data_terc, hora_terc = format_datetime_card(row.get('Atendimento Terceiros'))
                        data_norm, hora_norm = format_datetime_card(row.get('Normalização'))

                        quantidade_html = ''
                        if row.get('Categoria') == utils.SHEET_EQUIPAMENTOS:
                            try:
                                qv = float(row.get('Quantidade', 0))
                                if qv > 0:
                                    quantidade_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {int(qv)}</div>'
                            except: pass

                        card_html = f"""
                        <div class="card-container">
                            <div class="card-title">{ug}</div>
                            <div class="card-item"><span class="card-label">Cliente:</span> {cliente}</div>
                            <div class="card-item"><span class="card-label">Categoria:</span> {categoria}</div>
                            <div class="card-item"><span class="card-label">Tipo de Ocorrência:</span> {tipo_ocorrencia}</div>
                            <div class="card-item"><span class="card-label">Ativo:</span> {ativo}</div>
                            <div class="card-item"><span class="card-label">Nome do ativo:</span> {nome_ativo}</div>
                            <div class="card-item"><span class="card-label">Ocorrência:</span> {ocorrencia}</div>
                            <div class="card-item"><span class="card-label">Operador:</span> {operador}</div>
                            {quantidade_html}
                            <br>
                            <div class="card-item"><span class="card-label">Data da ocorrência:</span> {data_ocor}</div>
                            <div class="card-item"><span class="card-label">Hora da ocorrência:</span> {hora_ocor}</div>
                            <div class="card-item"><span class="card-label">Data cliente avisado:</span> {data_ca}</div>
                            <div class="card-item"><span class="card-label">Hora cliente avisado:</span> {hora_ca}</div>
                            <div class="card-item"><span class="card-label">Data do atendimento LOOP:</span> {data_loop}</div>
                            <div class="card-item"><span class="card-label">Hora do atendimento LOOP:</span> {hora_loop}</div>
                            <div class="card-item"><span class="card-label">Data do atendimento de terceiros:</span> {data_terc}</div>
                            <div class="card-item"><span class="card-label">Hora do atendimento de terceiros:</span> {hora_terc}</div>
                            <div class="card-item"><span class="card-label">Data de normalização:</span> {data_norm}</div>
                            <div class="card-item"><span class="card-label">Hora de normalização:</span> {hora_norm}</div>
                            <br>
                            <div class="card-item"><span class="card-label">Descrição:</span> {descricao}</div>
                            <div class="card-item"><span class="card-label">Protocolo:</span> {protocolo}</div>
                            <div class="card-item"><span class="card-label">OS:</span> {os_val}</div>
                        </div>
                        """
                        st.html(card_html)
    else:
        st.info("Nenhuma ocorrência encontrada para os filtros.")
else:
    st.warning("Não foi possível carregar os dados.")