import os
import streamlit as st
import pandas as pd
from datetime import datetime
import html
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe

# --- 1. Configuração da Página e Layout ---
st.set_page_config(layout="wide")

# --- 2. Dicionário para tradução dos meses ---
meses_traducao = {
    'January': 'Janeiro', 'February': 'Fevereiro', 'March': 'Março',
    'April': 'Abril', 'May': 'Maio', 'June': 'Junho',
    'July': 'Julho', 'August': 'Agosto', 'September': 'Setembro',
    'October': 'Outubro', 'November': 'Novembro', 'December': 'Dezembro'
}
meses_cronologicos = list(meses_traducao.values())

# --- 3. CSS ---
st.markdown("""
<style>
    .stButton > button {
        background-color: #28a745;
        color: white;
        font-weight: bold;
        border-radius: 5px;
        padding: 10px 20px;
        width: 100%;
        border: none;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: background-color 0.3s;
    }
    .stButton > button:hover { background-color: #218838; }
    .kpi-card {
        background-color: #333333;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 20px;
    }
    .kpi-value { font-size: 3em; font-weight: bold; color: #FF4B4B; }
    .kpi-label { font-size: 1.2em; color: #AAAAAA; }
    .stMultiSelect { max-height: 200px; overflow-y: auto; }
    .column-header { font-weight: bold; font-size: 1.2em; }
    
    /* Estilo para os cards */
    .card-container {
        background-color: #FF4B4B;
        color: white;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 15px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        height: 100%;
    }
    .card-title {
        font-size: 1.5em; font-weight: bold; color: white;
        border-bottom: 1px solid rgba(255,255,255,0.5);
        padding-bottom: 5px; margin-bottom: 10px;
    }
    .card-item { margin-bottom: 5px; font-size: 1em; }
    .card-label { font-weight: bold; }
    
    .streamlit-dataframe table td {
        word-break: break-word;
        white-space: normal;
    }
</style>
""", unsafe_allow_html=True)

# --- 4. Carregar e Tratar os Dados ---

# Define os "escopos" - as permissões que nosso script solicitará.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Nome das planilhas que vamos ler
CREDS_FILE = "google_credentials.json"
PLANILHA_NOME_1 = "DESLIGAMENTOS"
PLANILHA_NOME_2 = "EQUIPAMENTOS"

@st.cache_resource(ttl=600)
def connect_to_google_sheets():
    # Verifica se está rodando localmente (o arquivo existe) ou na nuvem (usa st.secrets)
    if os.path.exists(CREDS_FILE):
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    
    client = gspread.authorize(creds)
    return client

def fetch_sheet_as_df(worksheet):
    data = worksheet.get_all_values()
    if not data:
        return pd.DataFrame()
    
    headers = [header.strip() for header in data.pop(0)]
    df = pd.DataFrame(data, columns=headers)
    return df

@st.cache_data(ttl=600)
def carregar_dados_google_sheets():
    try:
        client = connect_to_google_sheets()
        spreadsheet_url = "https://docs.google.com/spreadsheets/d/1KeJjbsLVP9DkxPCmNSN4VzbSBeG3SFSCAdPhir39iqg/edit?usp=sharing"
        workbook = client.open_by_url(spreadsheet_url)

        df_desligamentos = fetch_sheet_as_df(workbook.worksheet(PLANILHA_NOME_1))
        df_equipamentos = fetch_sheet_as_df(workbook.worksheet(PLANILHA_NOME_2))

        if 'IDENTIFICADOR' in df_desligamentos.columns:
            df_desligamentos['IDENTIFICADOR'] = df_desligamentos['IDENTIFICADOR'].astype(str)
        if 'IDENTIFICADOR' in df_equipamentos.columns:
            df_equipamentos['IDENTIFICADOR'] = df_equipamentos['IDENTIFICADOR'].astype(str)

        df_desligamentos.dropna(how='all', inplace=True)
        df_equipamentos.dropna(how='all', inplace=True)
        
        df_desligamentos['Categoria'] = 'DESLIGAMENTOS'
        df_equipamentos['Categoria']  = 'EQUIPAMENTOS'
        df_todos_dados = pd.concat([df_desligamentos, df_equipamentos], ignore_index=True)

        mapa_renomear = {
            'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência',
            'ATIVO': 'Ativo', 'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência',
            'QUANTIDADE': 'Quantidade', 'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização',
            'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 'DESCRIÇÃO': 'Descrição',
            'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop',
            'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo', 'CLIENTE AVISADO': 'Cliente Avisado'
        }
        colunas_atuais = df_todos_dados.columns
        renomear_final = {}
        for col in colunas_atuais:
            col_strip_upper = col.strip().upper()
            if col_strip_upper in mapa_renomear:
                renomear_final[col] = mapa_renomear[col_strip_upper]
        df_todos_dados.rename(columns=renomear_final, inplace=True)

        df_todos_dados.fillna('', inplace=True)

        # Garante que a coluna 'Cliente' existe antes de filtrar
        if 'Cliente' in df_todos_dados.columns:
            df_todos_dados = df_todos_dados[
                (df_todos_dados['Cliente'] != '') &
                (df_todos_dados['UG'] != '') &
                (df_todos_dados['Sigla'] != '')
            ].copy()
        
        colunas_datetime = ['Normalização', 'Desligamento', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']
        for col in colunas_datetime:
            if col in df_todos_dados.columns:
                df_todos_dados[col] = pd.to_datetime(df_todos_dados[col], errors='coerce')

        colunas_texto = ['Operador', 'Descrição', 'OS', 'Protocolo']
        for col in colunas_texto:
            if col in df_todos_dados.columns:
                 df_todos_dados[col] = df_todos_dados[col].astype(str).fillna('')

        # Verifica se a coluna 'Desligamento' existe e não está vazia antes de processar
        if 'Desligamento' in df_todos_dados.columns and not df_todos_dados['Desligamento'].isnull().all():
            df_todos_dados['Data'] = df_todos_dados['Desligamento'].dt.strftime('%Y-%m-%d')
            df_todos_dados['Hora'] = df_todos_dados['Desligamento'].dt.strftime('%H:%M:%S')
            df_todos_dados['Mês']  = df_todos_dados['Desligamento'].dt.strftime('%B').map(meses_traducao)
            df_todos_dados['Ano']  = df_todos_dados['Desligamento'].dt.year.fillna(0).astype(int)
            df_todos_dados['Dia']  = df_todos_dados['Desligamento'].dt.day.fillna(0).astype(int)

            df_todos_dados['ID_Unico'] = df_todos_dados['UG'].astype(str).str.upper() + "|" + \
                                    df_todos_dados['Ativo'].astype(str).str.upper() + "|" + \
                                    df_todos_dados['Ocorrência'].astype(str).str.upper() + "|" + \
                                    df_todos_dados['Desligamento'].astype(str)
        else:
            # Cria colunas vazias se 'Desligamento' não existir, para evitar erros posteriores
            for col in ['Data', 'Hora', 'Mês', 'Ano', 'Dia', 'ID_Unico']:
                df_todos_dados[col] = None


        return df_todos_dados

    except FileNotFoundError:
        st.error(f"Erro: O arquivo de credenciais '{CREDS_FILE}' não foi encontrado. Verifique se ele está na mesma pasta do seu script principal (app.py).")
        return pd.DataFrame()
    except gspread.exceptions.SpreadsheetNotFound:
        st.error("Erro: Planilha não encontrada. Verifique o link e se você compartilhou a planilha com o email da conta de serviço.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Ocorreu um erro ao carregar ou processar os dados do Google Sheets: {e}")
        return pd.DataFrame()



df_todos_dados = carregar_dados_google_sheets()


# Garante que a coluna de data/hora está no formato correto
df_todos_dados['Desligamento'] = pd.to_datetime(df_todos_dados['Desligamento'], errors='coerce')


# --- 5. Inicialização dos Filtros ---
if 'filtros_meses' not in st.session_state:
    st.session_state.filtros_meses = [meses_traducao[datetime.now().strftime('%B')]]
if 'filtros_anos' not in st.session_state:
    if not df_todos_dados.empty and 'Ano' in df_todos_dados.columns:
        anos_atuais = sorted(df_todos_dados['Ano'].unique().tolist())
        st.session_state.filtros_anos = [a for a in anos_atuais if a != 0]
    else:
        st.session_state.filtros_anos = []
if 'filtros_dias' not in st.session_state:
    if not df_todos_dados.empty and {'Mês','Ano'}.issubset(df_todos_dados.columns):
        dias_atuais = sorted(df_todos_dados[(df_todos_dados['Mês'].isin(st.session_state.filtros_meses)) & (df_todos_dados['Ano'].isin(st.session_state.filtros_anos))]['Dia'].unique().tolist())
        st.session_state.filtros_dias = [d for d in dias_atuais if d != 0]
    else:
        st.session_state.filtros_dias = []
if 'filtros_categorias' not in st.session_state:
    st.session_state.filtros_categorias = sorted(df_todos_dados['Categoria'].unique().tolist()) if not df_todos_dados.empty else []
if 'filtros_clientes' not in st.session_state:
    st.session_state.filtros_clientes = sorted(df_todos_dados['Cliente'].unique().tolist()) if not df_todos_dados.empty else []
if 'filtros_ugs' not in st.session_state:
    st.session_state.filtros_ugs = sorted(df_todos_dados['UG'].unique().tolist()) if not df_todos_dados.empty else []
if 'filtros_tipos' not in st.session_state:
    st.session_state.filtros_tipos = sorted(df_todos_dados['Tipo de ocorrência'].unique().tolist()) if not df_todos_dados.empty else []
if 'filtros_ativos' not in st.session_state:
    st.session_state.filtros_ativos = sorted(df_todos_dados['Ativo'].unique().tolist()) if not df_todos_dados.empty else []
if 'filtros_ocorrencias' not in st.session_state:
    st.session_state.filtros_ocorrencias = sorted(df_todos_dados['Ocorrência'].unique().tolist()) if not df_todos_dados.empty else []

# --- 6. Título e KPIs ---
st.title('Usinas desligadas no momento')
col_kpi1, col_kpi2 = st.columns(2)
with col_kpi1:
    if not df_todos_dados.empty and 'Normalização' in df_todos_dados.columns:
        df_desligadas_geral = df_todos_dados[pd.isna(df_todos_dados['Normalização']) | (df_todos_dados['Normalização'] == '')].copy()
        total_kpi_value = df_desligadas_geral.shape[0]
    else:
        total_kpi_value = 0
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">Total no Banco de Dados Completo</div>
        <div class="kpi-value">{total_kpi_value}</div>
    </div>
    """, unsafe_allow_html=True)

# --- 7. Botão de Atualização ---
col_top_left, col_top_right = st.columns([0.2, 0.8])
with col_top_left:
    if st.button('Atualizar Dados'):
        st.cache_data.clear()
        st.rerun()

# --- 8. Interface de Filtros ---
if not df_todos_dados.empty:
    st.subheader("Selecione o período desejado")
    col_ano, col_mes, col_dia = st.columns(3)
    
    with col_ano:
        st.write("### Ano(s):")
        anos_disponiveis = sorted([a for a in df_todos_dados['Ano'].unique() if a != 0])
        with st.expander("Expandir anos"):
            for ano in anos_disponiveis:
                st.checkbox(str(ano), key=f'cb_ano_{ano}', value=(ano in st.session_state.filtros_anos))
        col_botoes = st.columns(2)
        with col_botoes[0]:
            if st.button('Sel. Todos', key='sel_ano', use_container_width=True):
                st.session_state.filtros_anos = anos_disponiveis
                st.rerun()
        with col_botoes[1]:
            if st.button('Desmarcar', key='des_ano', use_container_width=True):
                st.session_state.filtros_anos = []
                st.rerun()

    with col_mes:
        st.write("### Mês(es):")
        meses_disponiveis = meses_cronologicos
        with st.expander("Expandir meses"):
            for mes in meses_disponiveis:
                st.checkbox(mes, key=f'cb_mes_{mes}', value=(mes in st.session_state.filtros_meses))
        col_botoes = st.columns(2)
        with col_botoes[0]:
            if st.button('Sel. Todos', key='sel_mes', use_container_width=True):
                st.session_state.filtros_meses = meses_disponiveis
                st.rerun()
        with col_botoes[1]:
            if st.button('Desmarcar', key='des_mes', use_container_width=True):
                st.session_state.filtros_meses = []
                st.rerun()

    with col_dia:
        st.write("### Dia(s):")
        meses_selecionados_input = [mes for mes in meses_cronologicos if st.session_state.get(f'cb_mes_{mes}')]
        anos_selecionados_input = [ano for ano in anos_disponiveis if st.session_state.get(f'cb_ano_{ano}')]
        dias_disponiveis_temp = df_todos_dados[
            df_todos_dados['Mês'].isin(meses_selecionados_input) & 
            df_todos_dados['Ano'].isin(anos_selecionados_input)
        ]['Dia'].unique().tolist()
        dias_disponiveis = sorted([d for d in dias_disponiveis_temp if d != 0])
        
        with st.expander("Expandir dias"):
            dias_cols = st.columns(7)
            for i, dia in enumerate(range(1, 32)):
                with dias_cols[i % 7]:
                    if dia in dias_disponiveis:
                        st.checkbox(str(dia), key=f'cb_dia_{dia}', value=(dia in st.session_state.filtros_dias))
                    else:
                        st.checkbox(str(dia), key=f'cb_dia_{dia}', disabled=True)
        col_botoes = st.columns(2)
        with col_botoes[0]:
            if st.button('Sel. Todos', key='sel_dia', use_container_width=True):
                st.session_state.filtros_dias = dias_disponiveis
                st.rerun()
        with col_botoes[1]:
            if st.button('Desmarcar', key='des_dia', use_container_width=True):
                st.session_state.filtros_dias = []
                st.rerun()

    st.subheader("Filtros Adicionais")
    col_cliente, col_ug, col_tipo, col_ativo, col_ocorrencia = st.columns(5)
    
    with col_cliente:
        st.write("Cliente:")
        col_botoes = st.columns(2)
        with col_botoes[0]:
            if st.button('Sel. Todos', key='sel_cli', use_container_width=True):
                st.session_state.filtros_clientes = sorted(df_todos_dados['Cliente'].unique().tolist())
                st.rerun()
        with col_botoes[1]:
            if st.button('Desmarcar', key='des_cli', use_container_width=True):
                st.session_state.filtros_clientes = []
                st.rerun()
        st.session_state.filtros_clientes = st.multiselect(
            ' ', options=sorted(df_todos_dados['Cliente'].unique().tolist()),
            default=st.session_state.filtros_clientes, label_visibility='hidden')

    with col_ug:
        st.write("UG:")
        df_temp = df_todos_dados[df_todos_dados['Cliente'].isin(st.session_state.filtros_clientes)]
        ugs_disponiveis = sorted(df_temp['UG'].unique().tolist())

        # --- LINHA ADICIONADA PARA A CORREÇÃO ---
        # Garante que apenas UGs válidas permaneçam selecionadas após a mudança do filtro de cliente.
        st.session_state.filtros_ugs = [ug for ug in st.session_state.filtros_ugs if ug in ugs_disponiveis]
        # -----------------------------------------

        col_botoes = st.columns(2)
        with col_botoes[0]:
            if st.button('Sel. Todos', key='sel_ug', use_container_width=True):
                st.session_state.filtros_ugs = ugs_disponiveis
                st.rerun()
        with col_botoes[1]:
            if st.button('Desmarcar', key='des_ug', use_container_width=True):
                st.session_state.filtros_ugs = []
                st.rerun()
        st.session_state.filtros_ugs = st.multiselect(
            ' ', options=ugs_disponiveis, default=st.session_state.filtros_ugs, label_visibility='hidden')

    with col_tipo:
        st.write("Tipo de Ocorrência:")
        col_botoes = st.columns(2)
        with col_botoes[0]:
            if st.button('Sel. Todos', key='sel_tipo', use_container_width=True):
                st.session_state.filtros_tipos = sorted(df_todos_dados['Tipo de ocorrência'].unique().tolist())
                st.rerun()
        with col_botoes[1]:
            if st.button('Desmarcar', key='des_tipo', use_container_width=True):
                st.session_state.filtros_tipos = []
                st.rerun()
        st.session_state.filtros_tipos = st.multiselect(
            ' ', options=sorted(df_todos_dados['Tipo de ocorrência'].unique().tolist()),
            default=st.session_state.filtros_tipos, label_visibility='hidden')

    with col_ativo:
        st.write("Ativo:")
        col_botoes = st.columns(2)
        with col_botoes[0]:
            if st.button('Sel. Todos', key='sel_ativo', use_container_width=True):
                st.session_state.filtros_ativos = sorted(df_todos_dados['Ativo'].unique().tolist())
                st.rerun()
        with col_botoes[1]:
            if st.button('Desmarcar', key='des_ativo', use_container_width=True):
                st.session_state.filtros_ativos = []
                st.rerun()
        st.session_state.filtros_ativos = st.multiselect(
            ' ', options=sorted(df_todos_dados['Ativo'].unique().tolist()),
            default=st.session_state.filtros_ativos, label_visibility='hidden')
    
    with col_ocorrencia:
        st.write("Ocorrência:")
        col_botoes = st.columns(2)
        with col_botoes[0]:
            if st.button('Sel. Todos', key='sel_ocorr', use_container_width=True):
                st.session_state.filtros_ocorrencias = sorted(df_todos_dados['Ocorrência'].unique().tolist())
                st.rerun()
        with col_botoes[1]:
            if st.button('Desmarcar', key='des_ocorr', use_container_width=True):
                st.session_state.filtros_ocorrencias = []
                st.rerun()
        st.session_state.filtros_ocorrencias = st.multiselect(
            ' ', options=sorted(df_todos_dados['Ocorrência'].unique().tolist()),
            default=st.session_state.filtros_ocorrencias, label_visibility='hidden')

    # --- Aplicação dos Filtros ---
    meses_selecionados = [mes for mes in meses_cronologicos if st.session_state.get(f'cb_mes_{mes}')]
    anos_selecionados = [ano for ano in anos_disponiveis if st.session_state.get(f'cb_ano_{ano}')]
    dias_selecionados = [dia for dia in dias_disponiveis if st.session_state.get(f'cb_dia_{dia}')]

    df_filtrado = df_todos_dados[
        (df_todos_dados['Mês'].isin(meses_selecionados)) &
        (df_todos_dados['Ano'].isin(anos_selecionados)) &
        (df_todos_dados['Dia'].isin(dias_selecionados)) &
        (df_todos_dados['Categoria'].isin(st.session_state.filtros_categorias)) &
        (df_todos_dados['Cliente'].isin(st.session_state.filtros_clientes)) &
        (df_todos_dados['UG'].isin(st.session_state.filtros_ugs)) &
        (df_todos_dados['Tipo de ocorrência'].isin(st.session_state.filtros_tipos)) &
        (df_todos_dados['Ativo'].isin(st.session_state.filtros_ativos)) &
        (df_todos_dados['Ocorrência'].isin(st.session_state.filtros_ocorrencias))
    ].copy()
    
    df_desligadas = df_filtrado[pd.isna(df_filtrado['Normalização']) | (df_filtrado['Normalização'] == '')].copy()
    
    with col_kpi2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total com Filtro Selecionado</div>
            <div class="kpi-value">{len(df_desligadas)}</div>
        </div>
        """, unsafe_allow_html=True)
    
    if not df_desligadas.empty:
        df_desligadas['Tempo em Segundos'] = (datetime.now() - df_desligadas['Desligamento']).dt.total_seconds().astype(int)

        # --- CONTROLES DE ORDENAÇÃO ---
        st.markdown("---")
        st.write("### Ordenar e Editar")
        sort_cols = st.columns(2)
        
        with sort_cols[0]:
            sort_options_display = {
                'Data do Desligamento': 'Desligamento',
                'Tempo de Desligamento': 'Tempo em Segundos',
                'UG': 'UG',
                'Ativo': 'Ativo'
            }
            sort_by_display = st.selectbox(
                "Ordenar por:",
                options=sort_options_display.keys(), index=0)
            sort_by_column = sort_options_display[sort_by_display]

        with sort_cols[1]:
            sort_order = st.radio(
                "Ordem:",
                options=['Descendente', 'Ascendente'], index=0, horizontal=True)
            is_ascending = (sort_order == 'Ascendente')

        df_sorted = df_desligadas.sort_values(by=sort_by_column, ascending=is_ascending)

        # ***** NOVO: SELEÇÃO PARA EDIÇÃO *****
        st.markdown("---")
        st.write("### Editar uma Ocorrência")

        # Criamos uma coluna 'Display' para facilitar a seleção no selectbox
        df_sorted['Display'] = df_sorted['UG'].astype(str) + " | " + \
                               df_sorted['Ativo'].astype(str) + " | " + \
                               df_sorted['Nome Ativo'].astype(str) + " | " + \
                               df_sorted['Ocorrência'].astype(str) + " | " + \
                               df_sorted['Desligamento'].dt.strftime('%d/%m/%Y %H:%M')

        ocorrencia_selecionada_display = st.selectbox(
            "Selecione a ocorrência para editar:",
            options=df_sorted['Display'],
            index=None, # Nenhum selecionado por padrão
            placeholder="Escolha uma ocorrência..."
        )

        if ocorrencia_selecionada_display:
            # AQUI, passamos a pegar o valor da nova coluna 'ID_Unico'
            id_unico_para_editar = df_sorted[df_sorted['Display'] == ocorrencia_selecionada_display].iloc[0]['ID_Unico']
    
            # E salvamos em uma nova variável de sessão para clareza
            st.session_state['id_unico_para_editar'] = id_unico_para_editar
    
            if st.button("📝 Editar Ocorrência Selecionada"):
                st.switch_page("pages/3_Editar_Ocorrência.py")
        
        # --- LISTA DE OCORRÊNCIAS (TABELA) ---
        st.header("Lista de Ocorrências (Tabela)")
        df_para_tabela = df_sorted.copy()
        
        def formatar_tempo_estatico(row):
            dias = row['Tempo em Segundos'] // 86400
            horas = (row['Tempo em Segundos'] % 86400) // 3600
            minutos = (row['Tempo em Segundos'] % 3600) // 60
            return f"{dias}d {horas}h {minutos}m"
        
        df_para_tabela['Tempo de Desligamento'] = df_para_tabela.apply(formatar_tempo_estatico, axis=1)
        df_para_tabela.reset_index(inplace=True, drop=True)
        df_para_tabela['Linha'] = df_para_tabela.index + 1
        
        st.dataframe(df_para_tabela[[
            'Linha', 'Categoria', 'Tempo de Desligamento', 'UG', 'Data', 'Hora', 'Tipo de ocorrência', 
            'Ativo', 'Ocorrência', 'Operador', 'Descrição', 'OS'
        ]], use_container_width=True)

        # --- DETALHES POR OCORRÊNCIA (CARDS) ---
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
                    index, row = rows[i + j]
                    with cols[j]:
                        categoria = html.escape(str(row.get("Categoria", "")))
                        ug = html.escape(str(row.get("UG", "N/A")))
                        tipo_ocorrencia = html.escape(str(row.get("Tipo de ocorrência", "")))
                        ativo = html.escape(str(row.get("Ativo", "")))
                        nome_ativo = html.escape(str(row.get("Nome Ativo", "")))
                        ocorrencia = html.escape(str(row.get("Ocorrência", "")))
                        operador = html.escape(str(row.get("Operador", "")))
                        descricao = html.escape(str(row.get("Descrição", ""))).replace('\n', '<br>')
                        protocolo = html.escape(str(row.get("Protocolo", "")))
                        os = html.escape(str(row.get("OS", "")))

                        data_ocor, hora_ocor = format_datetime_card(row.get('Desligamento'))
                        data_ca, hora_ca = format_datetime_card(row.get('Cliente Avisado'))
                        data_loop, hora_loop = format_datetime_card(row.get('Atendimento Loop'))
                        data_terc, hora_terc = format_datetime_card(row.get('Atendimento Terceiros'))
                        data_norm, hora_norm = format_datetime_card(row.get('Normalização'))

                        quantidade_html = ''
                        if row.get('Categoria') == 'EQUIPAMENTOS':
                            quantidade_val = row.get('Quantidade', 0)
                            try:
                                if pd.notna(quantidade_val) and float(quantidade_val) > 0:
                                    quantidade_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {int(float(quantidade_val))}</div>'
                            except (ValueError, TypeError):
                                quantidade_html = ''

                        card_html = f"""
                        <div class="card-container">
                            <div class="card-title">{ug}</div>
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
                            <div class="card-item"><span class="card-label">OS:</span> {os}</div>
                        </div>
                        """
                        st.html(card_html)
                
    else:
        st.info("Nenhuma usina encontrada com o campo 'Normalização' em branco para os filtros selecionados.")
else:
    st.warning("Não foi possível carregar os dados. Verifique o arquivo local ou os filtros aplicados.")