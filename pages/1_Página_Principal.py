import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import html
import re
from collections import Counter, defaultdict
import utils

# --- 1. Configuração ---
st.set_page_config(layout="wide", page_title="Dashboard Ocorrências")

# Estado do overlay
if 'ui_phase' not in st.session_state:
    st.session_state.ui_phase = 'init'
if 'loading_ts' not in st.session_state:
    st.session_state.loading_ts = 0

# Renderiza Overlay
utils.render_loading_overlay(st.session_state.ui_phase)

def start_loading():
    st.session_state.ui_phase = 'loading'
    st.session_state.loading_ts = pytime.time()

if st.session_state.ui_phase == 'init':
    start_loading()

# Failsafe para destravar loading
if st.session_state.ui_phase == 'loading' and (pytime.time() - st.session_state.loading_ts) > 5:
    st.session_state.ui_phase = 'ready'
    st.session_state.loading_ts = 0

# --- Helpers de Texto ---
def _collapse_spaces(s: str) -> str:
    return " ".join(str(s).split())

def canon(s) -> str:
    if s is None: return ""
    return _collapse_spaces(str(s)).casefold()

def options_from(series: pd.Series) -> list:
    # Retorna opções únicas, sem vazios
    series = series.astype(str).map(_collapse_spaces)
    unique_vals = sorted([x for x in series.unique() if x and x.lower() != "nan" and x != "" and x != "-"])
    return ["-"] + unique_vals

def matches_any_canon(series: pd.Series, selected: list[str]) -> pd.Series:
    if not selected:
        return pd.Series([True]*len(series), index=series.index)
    sel_c = {canon(s) for s in selected if s and s != "-"}
    return series.astype(str).map(canon).isin(sel_c)

meses_traducao = {
    'January': 'Janeiro', 'February': 'Fevereiro', 'March': 'Março',
    'April': 'Abril', 'May': 'Maio', 'June': 'Junho',
    'July': 'Julho', 'August': 'Agosto', 'September': 'Setembro',
    'October': 'Outubro', 'November': 'Novembro', 'December': 'Dezembro'
}
meses_cronologicos = list(meses_traducao.values())

# --- CSS Personalizado ---
st.markdown("""
<style>
    /* Estilo dos Botões Verdes de Filtro */
    div[data-testid="stExpander"] .stButton > button {
        background-color: #28a745;
        color: white;
        font-weight: bold;
        border-radius: 4px;
        border: none;
        height: auto;
        padding: 4px 10px;
        font-size: 0.85em;
        width: 100%;
    }
    div[data-testid="stExpander"] .stButton > button:hover {
        background-color: #218838;
    }

    /* KPIs */
    .kpi-card {
        background-color: #333333; padding: 20px; border-radius: 10px;
        text-align: center; margin-bottom: 20px;
    }
    .kpi-value { font-size: 3em; font-weight: bold; color: #FF4B4B; }
    .kpi-label { font-size: 1.2em; color: #FFFFFF; }
    
    /* Cards de Ocorrência */
    .card-container {
        background-color: #FF4B4B; color: white; padding: 15px;
        border-radius: 8px; margin-bottom: 15px; box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        height: 100%;
    }
    .card-title {
        font-size: 1.5em; font-weight: bold; color: white;
        border-bottom: 1px solid rgba(255,255,255,0.5);
        padding-bottom: 5px; margin-bottom: 10px;
    }
    .card-item { margin-bottom: 5px; font-size: 1em; }
    .card-label { font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- Carregamento de Dados ---
@st.cache_data(ttl=600)
def carregar_dados(cb):
    try:
        client = utils.connect_to_google_sheets()
        if not client: return pd.DataFrame()
        
        wb = client.open_by_url(utils.SPREADSHEET_URL)
        df1 = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_DESLIGAMENTOS))
        df2 = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_EQUIPAMENTOS))
        
        df1['Categoria'] = 'DESLIGAMENTOS'
        df2['Categoria'] = 'EQUIPAMENTOS'
        df = pd.concat([df1, df2], ignore_index=True)
        
        # Mapa de Renomeação Seguro
        mapa = {
            'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência',
            'ATIVO': 'Ativo', 'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência', 'QUANTIDADE': 'Quantidade', 
            'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização', 'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 
            'DESCRIÇÃO': 'Descrição', 'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop',
            'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo', 'CLIENTE AVISADO': 'Cliente Avisado'
        }
        
        renomear = {}
        for col in df.columns:
            c_upper = col.strip().upper()
            if c_upper in mapa: renomear[col] = mapa[c_upper]
        df.rename(columns=renomear, inplace=True)
        
        # Limpeza e Datas
        if 'Cliente' in df.columns:
            df = df[(df['Cliente'] != '') & (df['UG'] != '')]
            
        cols_dt = ['Normalização', 'Desligamento', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']
        for c in cols_dt:
            if c in df.columns: df[c] = pd.to_datetime(df[c], errors='coerce', dayfirst=True)
            
        if 'Desligamento' in df.columns:
            df['Data'] = df['Desligamento'].dt.strftime('%Y-%m-%d')
            df['Hora'] = df['Desligamento'].dt.strftime('%H:%M:%S')
            df['Mês']  = df['Desligamento'].dt.strftime('%B').map(meses_traducao)
            df['Ano']  = df['Desligamento'].dt.year.fillna(0).astype(int)
            df['Dia']  = df['Desligamento'].dt.day.fillna(0).astype(int)
            
            df['ID_Unico'] = (
                df['UG'].astype(str).str.upper() + "|" +
                df['Ativo'].astype(str).str.upper() + "|" +
                df['Ocorrência'].astype(str).str.upper() + "|" +
                df['Desligamento'].astype(str)
            )
        
        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

if 'cache_buster' not in st.session_state:
    st.session_state.cache_buster = int(pytime.time())

df_todos = carregar_dados(st.session_state.cache_buster)

if st.session_state.ui_phase == 'loading':
    st.session_state.ui_phase = 'ready'
    st.session_state.loading_ts = 0
    utils.render_loading_overlay('ready')

# ==============================================================================
# --- BARRA LATERAL: FILTROS ---
# ==============================================================================
with st.sidebar:
    st.header("Filtros")
    
    if st.button("🔄 Atualizar Dados"):
        st.cache_data.clear()
        start_loading()
        st.rerun()
    
    st.markdown("---")
    
    # Categoria
    st.markdown("**Categoria**")
    if 'categoria_top' not in st.session_state: st.session_state['categoria_top'] = 'Ambas'
    st.radio("Cat", ["Ambas", "DESLIGAMENTOS", "EQUIPAMENTOS"], key="categoria_top", label_visibility="collapsed", on_change=start_loading)
    
    st.markdown("---")
    st.markdown("**Período**")
    
    # Inicializa variáveis de estado se não existirem
    if 'filtros_anos' not in st.session_state: 
        st.session_state.filtros_anos = sorted([a for a in df_todos['Ano'].unique() if a != 0]) if not df_todos.empty else []
    if 'filtros_meses' not in st.session_state: 
        st.session_state.filtros_meses = [meses_traducao[datetime.now().strftime('%B')]]
    if 'filtros_dias' not in st.session_state: st.session_state.filtros_dias = []
    
    # Opções disponíveis
    anos_disp = sorted([a for a in df_todos['Ano'].unique() if a != 0]) if not df_todos.empty else []
    
    # Função helper para botões
    def set_filtro(key, values):
        st.session_state[key] = values
        start_loading()
    
    # 1. Anos
    with st.expander("📅 Anos", expanded=True):
        c1, c2 = st.columns(2)
        c1.button("Sel. Todos", key="btn_all_ano", on_click=set_filtro, args=('filtros_anos', anos_disp))
        c2.button("Desmarcar", key="btn_none_ano", on_click=set_filtro, args=('filtros_anos', []))
        st.session_state.filtros_anos = st.multiselect("Selecione", anos_disp, default=st.session_state.filtros_anos, key="ms_anos", label_visibility="collapsed")

    # 2. Meses
    with st.expander("📆 Meses", expanded=False):
        c1, c2 = st.columns(2)
        c1.button("Sel. Todos", key="btn_all_mes", on_click=set_filtro, args=('filtros_meses', meses_cronologicos))
        c2.button("Desmarcar", key="btn_none_mes", on_click=set_filtro, args=('filtros_meses', []))
        st.session_state.filtros_meses = st.multiselect("Selecione", meses_cronologicos, default=st.session_state.filtros_meses, key="ms_meses", label_visibility="collapsed")

    # 3. Dias (Dinâmico)
    with st.expander("numeric Dias", expanded=False):
        # Calcula dias baseados na seleção atual de ano/mês
        if not df_todos.empty:
            mask = (df_todos['Ano'].isin(st.session_state.filtros_anos)) & (df_todos['Mês'].isin(st.session_state.filtros_meses))
            dias_disp = sorted(df_todos[mask]['Dia'].unique().astype(int).tolist())
            dias_disp = [d for d in dias_disp if d != 0]
        else:
            dias_disp = []
            
        c1, c2 = st.columns(2)
        c1.button("Sel. Todos", key="btn_all_dia", on_click=set_filtro, args=('filtros_dias', dias_disp))
        c2.button("Desmarcar", key="btn_none_dia", on_click=set_filtro, args=('filtros_dias', []))
        
        # Limpa seleção inválida
        st.session_state.filtros_dias = [d for d in st.session_state.filtros_dias if d in dias_disp]
        st.session_state.filtros_dias = st.multiselect("Selecione", dias_disp, default=st.session_state.filtros_dias, key="ms_dias", label_visibility="collapsed")

    st.markdown("---")
    st.markdown("**Filtros Adicionais**")
    
    # Inicializa adicionais
    for k in ['filtros_clientes', 'filtros_ugs', 'filtros_tipos', 'filtros_ativos', 'filtros_ocorrencias']:
        if k not in st.session_state: st.session_state[k] = []

    def render_sidebar_filter(label, key, col_name):
        if col_name not in df_todos.columns: return
        with st.expander(label):
            opts = options_from(df_todos[col_name])
            opts_clean = [o for o in opts if o != "-"]
            
            c1, c2 = st.columns(2)
            c1.button("Sel. Todos", key=f"all_{key}", on_click=set_filtro, args=(key, opts_clean))
            c2.button("Desmarcar", key=f"none_{key}", on_click=set_filtro, args=(key, []))
            
            st.session_state[key] = st.multiselect("Selecione", opts_clean, default=[x for x in st.session_state[key] if x in opts_clean], key=f"ms_{key}", label_visibility="collapsed")

    render_sidebar_filter("Clientes", 'filtros_clientes', 'Cliente')
    render_sidebar_filter("UGs", 'filtros_ugs', 'UG')
    render_sidebar_filter("Tipos", 'filtros_tipos', 'Tipo de ocorrência')
    render_sidebar_filter("Ativos", 'filtros_ativos', 'Ativo')
    render_sidebar_filter("Ocorrências", 'filtros_ocorrencias', 'Ocorrência')

# ==============================================================================
# --- ÁREA PRINCIPAL ---
# ==============================================================================

# KPIs Superiores (Ativas)
count_deslig = df_todos[(df_todos['Categoria'] == 'DESLIGAMENTOS') & (pd.isna(df_todos['Normalização']) | (df_todos['Normalização'] == ''))].shape[0] if not df_todos.empty else 0
count_equip = df_todos[(df_todos['Categoria'] == 'EQUIPAMENTOS') & (pd.isna(df_todos['Normalização']) | (df_todos['Normalização'] == ''))].shape[0] if not df_todos.empty else 0

with st.container(border=True):
    st.markdown("<h1 style='margin:0; text-align:center;'>OCORRÊNCIAS ATIVAS</h1>", unsafe_allow_html=True)
    st.write("")
    c1, c2 = st.columns(2)
    c1.markdown(f"<div class='kpi-card'><div class='kpi-label'>USINAS DESLIGADAS</div><div class='kpi-value'>{count_deslig}</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='kpi-card'><div class='kpi-label'>EQUIPAMENTOS PARADOS</div><div class='kpi-value'>{count_equip}</div></div>", unsafe_allow_html=True)

# Aplicação Filtros
m_cat = pd.Series([True]*len(df_todos))
if st.session_state.categoria_top != "Ambas":
    m_cat = df_todos['Categoria'] == st.session_state.categoria_top

m_ano = df_todos['Ano'].isin(st.session_state.filtros_anos) if st.session_state.filtros_anos else pd.Series([True]*len(df_todos))
m_mes = df_todos['Mês'].isin(st.session_state.filtros_meses) if st.session_state.filtros_meses else pd.Series([True]*len(df_todos))
m_dia = df_todos['Dia'].isin(st.session_state.filtros_dias) if st.session_state.filtros_dias else pd.Series([True]*len(df_todos))

m_cli = matches_any_canon(df_todos['Cliente'], st.session_state.filtros_clientes)
m_ug = matches_any_canon(df_todos['UG'], st.session_state.filtros_ugs) # Usando canon para UG também por segurança
m_tip = matches_any_canon(df_todos['Tipo de ocorrência'], st.session_state.filtros_tipos)
m_atv = matches_any_canon(df_todos['Ativo'], st.session_state.filtros_ativos)
m_ocr = matches_any_canon(df_todos['Ocorrência'], st.session_state.filtros_ocorrencias)

df_filt = df_todos[m_cat & m_ano & m_mes & m_dia & m_cli & m_ug & m_tip & m_atv & m_ocr].copy()
df_abertas = df_filt[pd.isna(df_filt['Normalização']) | (df_filt['Normalização'] == '')].copy()

# KPI Filtrado
st.markdown("---")
total_db = df_todos[(pd.isna(df_todos['Normalização']) | (df_todos['Normalização'] == ''))].shape[0] if not df_todos.empty else 0

c1, c2 = st.columns(2)
c1.markdown(f"<div class='kpi-card'><div class='kpi-label'>Total no Banco (Abertas)</div><div class='kpi-value'>{total_db}</div></div>", unsafe_allow_html=True)
c2.markdown(f"<div class='kpi-card'><div class='kpi-label'>Total com Filtro Selecionado</div><div class='kpi-value'>{len(df_abertas)}</div></div>", unsafe_allow_html=True)

if not df_abertas.empty:
    mask_valid = df_abertas['Desligamento'].notna()
    df_abertas.loc[mask_valid, 'Tempo em Segundos'] = (
        (datetime.now() - df_abertas.loc[mask_valid, 'Desligamento']).dt.total_seconds().astype(int)
    )
    df_abertas.loc[~mask_valid, 'Tempo em Segundos'] = 0

    st.markdown("---")
    
    # Ordenação e Edição
    c1, c2 = st.columns(2)
    sort_by = c1.selectbox("Ordenar por:", ["Data do Desligamento", "Tempo de Desligamento", "UG", "Ativo"])
    sort_ord = c2.radio("Ordem:", ["Descendente", "Ascendente"], horizontal=True)
    
    col_map = {'Data do Desligamento': 'Desligamento', 'Tempo de Desligamento': 'Tempo em Segundos', 'UG': 'UG', 'Ativo': 'Ativo'}
    df_sorted = df_abertas.sort_values(by=col_map[sort_by], ascending=(sort_ord == "Ascendente"))

    # Select Edição
    df_sorted['Display'] = (
        df_sorted['UG'].astype(str) + " | " + df_sorted['Ocorrência'].astype(str) + " | " +
        df_sorted['Desligamento'].dt.strftime('%d/%m %H:%M').fillna('')
    )
    st.session_state['df_lista_para_editar'] = df_sorted.copy()
    
    sel = st.selectbox("Selecione para editar:", df_sorted['Display'].tolist(), index=None, placeholder="Escolha uma ocorrência...")
    if sel:
        id_unico = df_sorted.loc[df_sorted['Display'] == sel, 'ID_Unico'].values[0]
        st.session_state['id_unico_para_editar'] = id_unico
    
    if st.button("📝 Editar Selecionada", disabled=not bool(sel)):
        st.switch_page("pages/3_Editar_Ocorrência.py")

    # Tabela
    st.markdown("### Lista de Ocorrências")
    
    def fmt_tempo(row):
        s = row['Tempo em Segundos']
        d, r = divmod(s, 86400); h, r = divmod(r, 3600); m, s = divmod(r, 60)
        return f"{int(d)}d {int(h)}h {int(m)}m"
    
    df_view = df_sorted.copy()
    df_view['Tempo'] = df_view.apply(fmt_tempo, axis=1)
    # Seleção de colunas SEGURA
    cols_table = ['Categoria', 'Tempo', 'UG', 'Data', 'Hora', 'Tipo de ocorrência', 'Ativo', 'Ocorrência', 'Descrição']
    cols_existentes = [c for c in cols_table if c in df_view.columns]
    st.dataframe(df_view[cols_existentes], use_container_width=True)

    # Cards (Mesmo código visual detalhado)
    st.markdown("### Detalhes (Cards)")
    num_cols = 4
    rows = list(df_sorted.iterrows())
    
    def fmt_dt(dt):
        if pd.notna(dt): return dt.strftime('%d/%m/%Y'), dt.strftime('%H:%M')
        return '', ''

    for i in range(0, len(rows), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j < len(rows):
                _, r = rows[i + j]
                with cols[j]:
                    ug = html.escape(str(r.get("UG", "N/A")))
                    cli = html.escape(str(r.get("Cliente", "")))
                    ocr = html.escape(str(r.get("Ocorrência", "")))
                    desc = html.escape(str(r.get("Descrição", ""))).replace('\n', '<br>')
                    d_des, h_des = fmt_dt(r.get('Desligamento'))
                    
                    st.markdown(f"""
                    <div class="card-container">
                        <div class="card-title">{ug}</div>
                        <div class="card-item"><span class="card-label">Cliente:</span> {cli}</div>
                        <div class="card-item"><span class="card-label">Ocorrência:</span> {ocr}</div>
                        <div class="card-item"><span class="card-label">Data:</span> {d_des} {h_des}</div>
                        <br>
                        <div class="card-item"><span class="card-label">Descrição:</span> {desc}</div>
                    </div>
                    """, unsafe_allow_html=True)
else:
    st.info("Nenhuma ocorrência encontrada.")