import os
import streamlit as st
import pandas as pd
from datetime import datetime, time
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(layout="wide")
st.title("📝 Editar Ocorrência")

# --- CONFIGURAÇÃO DE ACESSO AO GOOGLE SHEETS ---
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
CREDS_FILE = "google_credentials.json"
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1KeJjbsLVP9DkxPCmNSN4VzbSBeG3SFSCAdPhir39iqg/edit?usp=sharing"
PLANILHA_NOME_1 = "DESLIGAMENTOS"
PLANILHA_NOME_2 = "EQUIPAMENTOS"
PLANILHA_DADOS = "DADOS"
MAPA_RENOMEAR = {
    'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência',
    'ATIVO': 'Ativo', 'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência',
    'QUANTIDADE': 'Quantidade', 'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização',
    'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 'DESCRIÇÃO': 'Descrição',
    'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop', 
    'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo', 'CLIENTE AVISADO': 'Cliente Avisado'
}

def fetch_sheet_as_df(worksheet):
    data = worksheet.get_all_values()
    if not data: return pd.DataFrame()
    headers = [h.replace('\xa0', '').strip() for h in data.pop(0)]
    return pd.DataFrame(data, columns=headers)

@st.cache_resource(ttl=600)
def connect_to_google_sheets():
    # Verifica se está rodando localmente (o arquivo existe) ou na nuvem (usa st.secrets)
    if os.path.exists(CREDS_FILE):
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    
    client = gspread.authorize(creds)
    return client

@st.cache_data(ttl=60)
def carregar_dados_completos():
    try:
        client = connect_to_google_sheets()
        workbook = client.open_by_url(SPREADSHEET_URL)
        df_desligamentos = fetch_sheet_as_df(workbook.worksheet(PLANILHA_NOME_1))
        df_equipamentos = fetch_sheet_as_df(workbook.worksheet(PLANILHA_NOME_2))

        df_desligamentos['Categoria'] = 'DESLIGAMENTOS'
        df_equipamentos['Categoria']  = 'EQUIPAMENTOS'
        df_todos_dados = pd.concat([df_desligamentos, df_equipamentos], ignore_index=True)

        colunas_atuais = df_todos_dados.columns
        renomear_final = {}
        for col in colunas_atuais:
            col_strip_upper = col.strip().upper()
            if col_strip_upper in MAPA_RENOMEAR:
                renomear_final[col] = MAPA_RENOMEAR[col_strip_upper]
        df_todos_dados.rename(columns=renomear_final, inplace=True)
        df_todos_dados.fillna('', inplace=True)
        colunas_datetime = ['Normalização', 'Desligamento', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']
        for col in colunas_datetime:
            if col in df_todos_dados.columns:
                df_todos_dados[col] = pd.to_datetime(df_todos_dados[col], errors='coerce', dayfirst=False)
        
        df_todos_dados['ID_Unico'] = df_todos_dados['UG'].astype(str).str.upper() + "|" + \
                                    df_todos_dados['Ativo'].astype(str).str.upper() + "|" + \
                                    df_todos_dados['Ocorrência'].astype(str).str.upper() + "|" + \
                                    df_todos_dados['Desligamento'].astype(str)
        return df_todos_dados
    except Exception as e:
        st.error(f"Erro ao carregar dados do Google Sheets: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=600)
def carregar_opcoes_para_edicao():
    try:
        client = connect_to_google_sheets()
        workbook = client.open_by_url(SPREADSHEET_URL)
        df_dados = fetch_sheet_as_df(workbook.worksheet(PLANILHA_DADOS)).fillna('')
        
        for col in df_dados.columns:
            if df_dados[col].dtype == 'object': 
                df_dados[col] = df_dados[col].str.strip()

        opcoes = {
            'tipos_ocorrencia': sorted(df_dados[df_dados['TIPO DE OCORRÊNCIA'] != '']['TIPO DE OCORRÊNCIA'].unique().tolist()),
            'ocorrencias': sorted(df_dados[df_dados['OCORRÊNCIA'] != '']['OCORRÊNCIA'].unique().tolist()),
            'operadores': sorted(df_dados[df_dados['OPERADOR'] != '']['OPERADOR'].unique().tolist())
        }
        return opcoes
    except Exception as e:
        st.error(f"Erro ao carregar listas de opções: {e}")
        return {}

def combine_date_time(date_val, time_val):
    if date_val and time_val:
        return datetime.combine(date_val, time_val)
    return None

def split_datetime(dt_obj):
    if pd.notna(dt_obj) and isinstance(dt_obj, datetime):
        return dt_obj.date(), dt_obj.time()
    return None, None

# --- LÓGICA DA PÁGINA ---
if 'id_unico_para_editar' not in st.session_state or not st.session_state['id_unico_para_editar']:
    st.warning("Nenhuma ocorrência selecionada para edição.")
    st.page_link("pages/1_Página_Principal.py", label="Voltar para a Página Principal", icon="🏠")
else:
    id_para_editar = st.session_state['id_unico_para_editar']
    df_completo = carregar_dados_completos()
    opcoes_edicao = carregar_opcoes_para_edicao()

    if not df_completo.empty and opcoes_edicao:
        dados_ocorrencia = df_completo[df_completo['ID_Unico'] == id_para_editar]

        if not dados_ocorrencia.empty:
            ocorrencia = dados_ocorrencia.iloc[0].to_dict()
            categoria = ocorrencia.get('Categoria', 'DESLIGAMENTOS')

            with st.form("edit_form"):
                st.subheader(f"Editando Ocorrência em: {categoria}")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.text_input("UG", value=ocorrencia.get('UG'), key="ug")
                    st.text_input("Nome Ativo", value=ocorrencia.get('Nome Ativo'), key="nome_ativo")

                    tipo_ocorrencia_opts = opcoes_edicao.get('tipos_ocorrencia', [])
                    tipo_idx = 0
                    if ocorrencia.get('Tipo de ocorrência') in tipo_ocorrencia_opts:
                        tipo_idx = tipo_ocorrencia_opts.index(ocorrencia.get('Tipo de ocorrência'))
                    st.selectbox("Tipo de Ocorrência", options=tipo_ocorrencia_opts, index=tipo_idx, key="tipo_ocorrencia")

                    ocorrencia_opts = opcoes_edicao.get('ocorrencias', [])
                    ocorrencia_idx = 0
                    if ocorrencia.get('Ocorrência') in ocorrencia_opts:
                        ocorrencia_idx = ocorrencia_opts.index(ocorrencia.get('Ocorrência'))
                    st.selectbox("Ocorrência", options=ocorrencia_opts, index=ocorrencia_idx, key="ocorrencia")

                    operador_opts = opcoes_edicao.get('operadores', [])
                    operador_idx = 0
                    if ocorrencia.get('Operador') in operador_opts:
                        operador_idx = operador_opts.index(ocorrencia.get('Operador'))
                    st.selectbox("Operador", options=operador_opts, index=operador_idx, key="operador")

                    st.text_area("Descrição", value=ocorrencia.get('Descrição'), key="descricao")
                    st.text_input("OS", value=ocorrencia.get('OS'), key="os")
                    st.text_input("Protocolo", value=ocorrencia.get('Protocolo'), key="protocolo")

                with col2:
                    norm_date, norm_time = split_datetime(ocorrencia.get('Normalização'))
                    loop_date, loop_time = split_datetime(ocorrencia.get('Atendimento Loop'))
                    terc_date, terc_time = split_datetime(ocorrencia.get('Atendimento Terceiros'))
                    avis_date, avis_time = split_datetime(ocorrencia.get('Cliente Avisado'))
                    
                    st.date_input("Data Normalização", value=norm_date, key="norm_date")
                    st.time_input("Hora Normalização", value=norm_time, key="norm_time")
                    st.date_input("Data Atendimento Loop", value=loop_date, key="loop_date")
                    st.time_input("Hora Atendimento Loop", value=loop_time, key="loop_time")
                    st.date_input("Data Atendimento Terceiros", value=terc_date, key="terc_date")
                    st.time_input("Hora Atendimento Terceiros", value=terc_time, key="terc_time")
                    st.date_input("Data Cliente Avisado", value=avis_date, key="avis_date")
                    st.time_input("Hora Cliente Avisado", value=avis_time, key="avis_time")

                submitted = st.form_submit_button("✅ Salvar Alterações")

                if submitted:
                    try:
                        client = connect_to_google_sheets()
                        workbook = client.open_by_url(SPREADSHEET_URL)
                        worksheet = workbook.worksheet(categoria)
                        all_data = worksheet.get_all_values()
                        headers = all_data[0]
                        
                        row_to_edit = -1
                        idx_ug = headers.index('UG')
                        idx_ativo = headers.index('ATIVO')
                        idx_ocorrencia = headers.index('OCORRÊNCIA')
                        idx_desligamento = headers.index('DESLIGAMENTO')

                        for i, row in enumerate(all_data[1:], start=2):
                            try:
                                desligamento_dt = pd.to_datetime(row[idx_desligamento], errors='coerce')
                                if pd.isna(desligamento_dt): continue
                                current_id = f"{row[idx_ug].upper()}|{row[idx_ativo].upper()}|{row[idx_ocorrencia].upper()}|{desligamento_dt}"
                                if current_id == id_para_editar:
                                    row_to_edit = i
                                    break
                            except (IndexError, ValueError): continue
                        
                        if row_to_edit != -1:
                            dados_atualizados = ocorrencia.copy()
                            dados_atualizados['UG'] = st.session_state.ug
                            dados_atualizados['Nome Ativo'] = st.session_state.nome_ativo
                            dados_atualizados['Tipo de ocorrência'] = st.session_state.tipo_ocorrencia
                            dados_atualizados['Ocorrência'] = st.session_state.ocorrencia
                            dados_atualizados['Operador'] = st.session_state.operador
                            dados_atualizados['Descrição'] = st.session_state.descricao
                            dados_atualizados['OS'] = st.session_state.os
                            dados_atualizados['Protocolo'] = st.session_state.protocolo
                            
                            def format_dt(dt_obj):
                                return dt_obj.strftime('%Y-%m-%d %H:%M:%S') if dt_obj else ''

                            dados_atualizados['Normalização'] = format_dt(combine_date_time(st.session_state.norm_date, st.session_state.norm_time))
                            dados_atualizados['Atendimento Loop'] = format_dt(combine_date_time(st.session_state.loop_date, st.session_state.loop_time))
                            dados_atualizados['Atendimento Terceiros'] = format_dt(combine_date_time(st.session_state.terc_date, st.session_state.terc_time))
                            dados_atualizados['Cliente Avisado'] = format_dt(combine_date_time(st.session_state.avis_date, st.session_state.avis_time))
                            
                            mapa_renomear_inverso = {v: k for k, v in MAPA_RENOMEAR.items()}
                            linha_para_atualizar = []
                            for h in headers:
                                h_strip = h.strip()
                                key_title_case = MAPA_RENOMEAR.get(h_strip.upper(), h_strip)
                                valor = dados_atualizados.get(key_title_case, '')

                                # --- CORREÇÃO APLICADA AQUI ---
                                # Verifica se o valor é um objeto de data/hora e o converte para texto
                                if isinstance(valor, (datetime, pd.Timestamp)):
                                    valor = valor.strftime('%Y-%m-%d %H:%M:%S')
                                # ---------------------------------
                                
                                linha_para_atualizar.append(valor)
                            
                            worksheet.update(f'A{row_to_edit}', [linha_para_atualizar], value_input_option='USER_ENTERED')

                            st.success("Ocorrência atualizada com sucesso!")
                            st.cache_data.clear()
                        else:
                            st.error("Não foi possível encontrar a linha na Planilha Google para editar.")
                    except Exception as e:
                        st.error(f"Ocorreu um erro ao atualizar a Planilha Google: {e}")
        else:
            st.error("O ID da ocorrência selecionada não foi encontrado nos dados carregados.")