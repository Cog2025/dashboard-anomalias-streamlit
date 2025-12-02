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

# --- CSS (Restaurado conforme prints) ---
st.markdown("""
<style>
    /* Botões Verdes para Filtros */
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
    .kpi-value { font-size: 3em; font-weight: bold; color: #4b4eff; }
    .kpi-label { font-size: 1.2em; color: #FFFFFF; }
    
    /* Card Detalhado */
    .card-container {
        background-color: #089641; /* Verde conforme imagem */
        color: white; padding: 15px;
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

# --- Dados ---
if st.session_state.ui_phase == 'init':
    utils.overlay_on()
    st.rerun()

df = utils.carregar_dados_completos(st.session_state.cache_buster)

if st.session_state.ui_phase != 'ready':
    utils.overlay_off()

# --- KPIs ---
count_deslig = 0
count_equip = 0
if not df.empty:
    count_deslig = df[(df['Categoria'] == utils.SHEET_DESLIGAMENTOS) & (pd.isna(df['Normalização']) | (df['Normalização'] == ''))].shape[0]
    count_equip = df[(df['Categoria'] == utils.SHEET_EQUIPAMENTOS) & (pd.isna(df['Normalização']) | (df['Normalização'] == ''))].shape[0]

with st.container(border=True):
    st.markdown("<h1 style='margin:0'>OCORRÊNCIAS ATIVAS</h1>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    c1.markdown(f"<div class='kpi-card'><div class='kpi-label'>USINAS DESLIGADAS</div><div class='kpi-value'>{count_deslig}</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='kpi-card'><div class='kpi-label'>EQUIPAMENTOS PARADOS</div><div class='kpi-value'>{count_equip}</div></div>", unsafe_allow_html=True)

# --- Inicialização Filtros ---
if 'filtros_anos' not in st.session_state:
    st.session_state.filtros_anos = sorted([a for a in df['Ano'].unique() if a != 0]) if not df.empty else []
if 'filtros_meses' not in st.session_state:
    st.session_state.filtros_meses = [utils.MESES_TRADUCAO[datetime.now().strftime('%B')]]
if 'filtros_dias' not in st.session_state:
    # Lógica original: calcular dias disponíveis com base na seleção inicial
    if not df.empty and st.session_state.filtros_anos and st.session_state.filtros_meses:
        mask = (df['Ano'].isin(st.session_state.filtros_anos)) & (df['Mês'].isin(st.session_state.filtros_meses))
        dias = sorted(df[mask]['Dia'].unique().astype(int).tolist())
        st.session_state.filtros_dias = [d for d in dias if d > 0]
    else:
        st.session_state.filtros_dias = []

for k in ['filtros_clientes', 'filtros_ugs', 'filtros_tipos', 'filtros_ativos', 'filtros_ocorrencias']:
    if k not in st.session_state: st.session_state[k] = []
if 'categoria_top' not in st.session_state: st.session_state.categoria_top = "Ambas"

# --- Botões de Ação ---
def _marcar(prefixo, itens, key, valor):
    st.session_state[key] = list(itens) if valor else []
    utils.overlay_on()

# --- Filtros Visuais ---
st.header('OCORRÊNCIAS FILTRADAS')

# KPI Banco
total_db = df[(pd.isna(df['Normalização']) | (df['Normalização'] == ''))].shape[0] if not df.empty else 0
c_kpi, c_btn = st.columns([1, 1])
c_kpi.markdown(f"<div class='kpi-card'><div class='kpi-label'>Total no Banco (Abertas)</div><div class='kpi-value'>{total_db}</div></div>", unsafe_allow_html=True)
if c_btn.button("Atualizar Dados"):
    st.cache_data.clear()
    utils.overlay_on()
    st.rerun()

if not df.empty:
    st.markdown("#### Filtrar por categoria (planilha)")
    st.radio("Categoria:", ["Ambas", utils.SHEET_DESLIGAMENTOS, utils.SHEET_EQUIPAMENTOS], 
             horizontal=True, key="categoria_top", on_change=utils.overlay_on)

    # --- Seção de Período ---
    st.subheader("Selecione o período desejado")
    c_ano, c_mes, c_dia = st.columns(3)

    # ANOS
    with c_ano:
        with st.container(border=True):
            st.write("### Ano(s):")
            opts = sorted([a for a in df['Ano'].unique() if a != 0])
            with st.expander("Expandir anos"): pass
            b1, b2 = st.columns(2)
            b1.button("Sel. Todos", key="sa_ano", on_click=_marcar, args=('', opts, 'filtros_anos', True))
            b2.button("Desmarcar", key="dm_ano", on_click=_marcar, args=('', opts, 'filtros_anos', False))
            st.session_state.filtros_anos = st.multiselect(" ", opts, default=st.session_state.filtros_anos, label_visibility="collapsed")

    # MESES
    with c_mes:
        with st.container(border=True):
            st.write("### Mês(es):")
            opts = utils.MESES_CRONOLOGICOS
            with st.expander("Expandir meses"): pass
            b1, b2 = st.columns(2)
            b1.button("Sel. Todos", key="sa_mes", on_click=_marcar, args=('', opts, 'filtros_meses', True))
            b2.button("Desmarcar", key="dm_mes", on_click=_marcar, args=('', opts, 'filtros_meses', False))
            st.session_state.filtros_meses = st.multiselect(" ", opts, default=st.session_state.filtros_meses, label_visibility="collapsed")

    # DIAS
    with c_dia:
        with st.container(border=True):
            st.write("### Dia(s):")
            # Logica de dias disponiveis
            mask_ano = df['Ano'].isin(st.session_state.filtros_anos) if st.session_state.filtros_anos else pd.Series([True]*len(df))
            mask_mes = df['Mês'].isin(st.session_state.filtros_meses) if st.session_state.filtros_meses else pd.Series([True]*len(df))
            dias_disp = sorted(df[mask_ano & mask_mes]['Dia'].dropna().unique().astype(int).tolist())
            dias_disp = [d for d in dias_disp if d > 0]

            with st.expander("Expandir dias"): pass
            b1, b2 = st.columns(2)
            b1.button("Sel. Todos", key="sa_dia", on_click=_marcar, args=('', dias_disp, 'filtros_dias', True))
            b2.button("Desmarcar", key="dm_dia", on_click=_marcar, args=('', dias_disp, 'filtros_dias', False))
            
            # Limpa inválidos
            st.session_state.filtros_dias = [d for d in st.session_state.filtros_dias if d in dias_disp]
            st.session_state.filtros_dias = st.multiselect(" ", dias_disp, default=st.session_state.filtros_dias, label_visibility="collapsed")

    # --- Filtros Adicionais ---
    st.subheader("Filtros Adicionais")
    
    def render_filter(label, col, key_store, key_s, key_d):
        with st.container(border=True):
            st.write(f"{label}:")
            opts = sorted(utils.options_from(df[col]))
            b1, b2 = st.columns(2)
            b1.button("Sel. Todos", key=key_s, on_click=_marcar, args=('', opts, key_store, True))
            b2.button("Desmarcar", key=key_d, on_click=_marcar, args=('', opts, key_store, False))
            # Valida defaults
            valid = [x for x in st.session_state[key_store] if x in opts]
            st.session_state[key_store] = st.multiselect(" ", opts, default=valid, label_visibility="collapsed", key=f"ms_{key_store}")

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: render_filter("Cliente", "Cliente", "filtros_clientes", "sa_cli", "dm_cli")
    with c2: render_filter("UG", "UG", "filtros_ugs", "sa_ug", "dm_ug")
    with c3: render_filter("Tipo", "Tipo de ocorrência", "filtros_tipos", "sa_tip", "dm_tip")
    with c4: render_filter("Ativo", "Ativo", "filtros_ativos", "sa_atv", "dm_atv")
    with c5: render_filter("Ocorrência", "Ocorrência", "filtros_ocorrencias", "sa_ocr", "dm_ocr")

    # --- Aplicação Filtros ---
    m_cat = pd.Series([True]*len(df))
    if st.session_state.categoria_top != "Ambas":
        m_cat = df['Categoria'] == st.session_state.categoria_top

    m_ano = df['Ano'].isin(st.session_state.filtros_anos) if st.session_state.filtros_anos else pd.Series([True]*len(df))
    m_mes = df['Mês'].isin(st.session_state.filtros_meses) if st.session_state.filtros_meses else pd.Series([True]*len(df))
    m_dia = df['Dia'].isin(st.session_state.filtros_dias) if st.session_state.filtros_dias else pd.Series([True]*len(df))
    
    m_cli = utils.matches_any_canon(df['Cliente'], st.session_state.filtros_clientes)
    m_ug = utils.matches_any_canon(df['UG'], st.session_state.filtros_ugs)
    m_tip = utils.matches_any_canon(df['Tipo de ocorrência'], st.session_state.filtros_tipos)
    m_atv = utils.matches_any_canon(df['Ativo'], st.session_state.filtros_ativos)
    m_ocr = utils.matches_any_canon(df['Ocorrência'], st.session_state.filtros_ocorrencias)

    df_filt = df[m_cat & m_ano & m_mes & m_dia & m_cli & m_ug & m_tip & m_atv & m_ocr].copy()
    df_abertas = df_filt[pd.isna(df_filt['Normalização']) | (df_filt['Normalização'] == '')].copy()

    # KPI Filtrado
    st.markdown(f"<div class='kpi-card'><div class='kpi-label'>Total com Filtro</div><div class='kpi-value'>{len(df_abertas)}</div></div>", unsafe_allow_html=True)

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
        
        # Display para Tabela
        def formatar_tempo(s):
            d = s // 86400; h = (s % 86400) // 3600; m = (s % 3600) // 60
            return f"{int(d)}d {int(h)}h {int(m)}m"
        df_show = df_sorted.copy()
        df_show['Tempo'] = df_show['Tempo em Segundos'].apply(formatar_tempo)
        
        # Seletor Edição
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
        
        if st.button("📝 Editar Ocorrência Selecionada", disabled=not bool(sel)):
             st.switch_page("pages/3_Editar_Ocorrência.py")
             
        # Tabela
        st.header("Lista de Ocorrências (Tabela)")
        st.dataframe(df_show[['Categoria', 'Tempo', 'UG', 'Data', 'Hora', 'Tipo de ocorrência', 'Ocorrência', 'Descrição']], use_container_width=True)

        # --- CARDS DETALHADOS (RESTAURADOS) ---
        st.header("Detalhes por Ocorrência (Cards)")
        num_cols = 4
        rows = list(df_sorted.iterrows())
        
        def fmt_dt(dt):
            if pd.notna(dt): return dt.strftime('%d/%m/%Y'), dt.strftime('%H:%M')
            return '', ''

        for i in range(0, len(rows), num_cols):
            cols = st.columns(num_cols)
            for j in range(num_cols):
                if i + j < len(rows):
                    _, row = rows[i+j]
                    with cols[j]:
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
                        
                        qtd_html = ''
                        if row.get('Categoria') == utils.SHEET_EQUIPAMENTOS:
                             try:
                                 qv = float(row.get('Quantidade', 0))
                                 if qv > 0: qtd_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {int(qv)}</div>'
                             except: pass

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
                          {qtd_html}
                          <br>
                          <div class="card-item"><span class="card-label">Data Desligamento:</span> {d_des} {h_des}</div>
                          <div class="card-item"><span class="card-label">Data Normalização:</span> {d_norm} {h_norm}</div>
                          <div class="card-item"><span class="card-label">Cliente Avisado:</span> {d_ca} {h_ca}</div>
                          <div class="card-item"><span class="card-label">Atendimento Loop:</span> {d_loop} {h_loop}</div>
                          <div class="card-item"><span class="card-label">Atendimento Terc.:</span> {d_terc} {h_terc}</div>
                          <br>
                          <div class="card-item"><span class="card-label">Descrição:</span> {desc}</div>
                          <div class="card-item"><span class="card-label">Protocolo:</span> {prot}</div>
                          <div class="card-item"><span class="card-label">OS:</span> {osv}</div>
                        </div>
                        """, unsafe_allow_html=True)
    else:
        st.info("Nenhuma ocorrência encontrada para os filtros selecionados.")