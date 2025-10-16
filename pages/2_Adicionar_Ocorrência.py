import os
from webdav3.client import Client
import io
import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import re

# --- 1. CONFIGURAÇÃO DA PÁGINA E CSS ---
st.set_page_config(layout="wide")
st.title("Adicionar Nova Ocorrência")
st.markdown("""
<style>
    .stButton > button {
        background-color: #28a745; color: white; font-weight: bold;
        border-radius: 5px; padding: 10px 20px; width: 100%; border: none;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); transition: background-color 0.3s;
    }
    .stButton > button:hover { background-color: #218838; }
    .card-container {
        background-color: #FF4B4B; /* Cor vermelha original para os cards */
        color: white; padding: 15px;
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

# --- 2. FUNÇÕES E CONFIGURAÇÕES DE DADOS ---
def sanitize_key(text):
    return re.sub(r'[^A-Za-z0-9_]', '_', str(text))

def format_datetime_card(dt_obj):
    if isinstance(dt_obj, datetime):
        return dt_obj.strftime('%d/%m/%Y'), dt_obj.strftime('%H:%M')
    if isinstance(dt_obj, str) and dt_obj:
        try:
            dt = pd.to_datetime(dt_obj)
            return dt.strftime('%d/%m/%Y'), dt.strftime('%H:%M')
        except (ValueError, TypeError): return '', ''
    return '', ''

PLANILHA_DESLIGAMENTOS = 'DESLIGAMENTOS'
PLANILHA_EQUIPAMENTOS = 'EQUIPAMENTOS'
PLANILHA_DADOS = 'DADOS'
PLANILHA_DETALHADA = 'Usinas_Detalhado'

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
CREDS_FILE = "google_credentials.json"
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1KeJjbsLVP9DkxPCmNSN4VzbSBeG3SFSCAdPhir39iqg/edit?usp=sharing"

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
def carregar_dados_e_opcoes():
    try:
        client = connect_to_google_sheets()
        workbook = client.open_by_url(SPREADSHEET_URL)
        df_dados = fetch_sheet_as_df(workbook.worksheet(PLANILHA_DADOS)).fillna('')
        df_detalhado = fetch_sheet_as_df(workbook.worksheet(PLANILHA_DETALHADA)).fillna('')

        for col in df_dados.columns:
            if df_dados[col].dtype == 'object': df_dados[col] = df_dados[col].str.strip()
        for col in df_detalhado.columns:
            if df_detalhado[col].dtype == 'object': df_detalhado[col] = df_detalhado[col].str.strip()
        
        op_cliente = ['-'] + sorted(df_dados[df_dados['CLIENTE'] != '']['CLIENTE'].unique().tolist())
        op_ocorrencia = ['-'] + sorted(df_dados[df_dados['OCORRÊNCIA'] != '']['OCORRÊNCIA'].unique().tolist())
        op_tipo = ['-'] + sorted(df_dados[df_dados['TIPO DE OCORRÊNCIA'] != '']['TIPO DE OCORRÊNCIA'].unique().tolist())
        op_ativo = ['-'] + sorted(df_dados[df_dados['ATIVO'] != '']['ATIVO'].unique().tolist())
        op_operador = ['-'] + sorted(df_dados[df_dados['OPERADOR'] != '']['OPERADOR'].unique().tolist())
        
        return { 'df_dados': df_dados, 'df_detalhado': df_detalhado, 'Cliente': op_cliente,
                 'Ocorrência': op_ocorrencia, 'Tipo de ocorrência': op_tipo, 'Ativo': op_ativo,
                 'Operador': op_operador }
    except Exception as e:
        st.error(f"Erro ao carregar os dados das planilhas: {e}"); return {}

# --- 3. INTERFACE DO STREAMLIT ---
dados_e_opcoes = carregar_dados_e_opcoes()
if not dados_e_opcoes: st.stop()
df_dados = dados_e_opcoes.get('df_dados', pd.DataFrame())
df_detalhado = dados_e_opcoes.get('df_detalhado', pd.DataFrame())

if 'last_submission_details' in st.session_state and st.session_state.last_submission_details:
    submitted_occurrences = st.session_state.last_submission_details
    st.success(f"{len(submitted_occurrences)} ocorrência(s) adicionada(s) com sucesso!")
    
    num_cols = 4 
    for i in range(0, len(submitted_occurrences), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j < len(submitted_occurrences):
                with cols[j]:
                    details = submitted_occurrences[i + j]
                    
                    data_ocor, hora_ocor = format_datetime_card(details.get('DESLIGAMENTO'))
                    data_loop, hora_loop = format_datetime_card(details.get('ATENDIMENTO LOOP'))
                    data_terc, hora_terc = format_datetime_card(details.get('ATENDIMENTO TERCEIROS'))
                    data_norm, hora_norm = format_datetime_card(details.get('NORMALIZAÇÃO'))

                    quantidade_html = ""
                    if "QUANTIDADE" in details and details['QUANTIDADE'] and str(details['QUANTIDADE']).strip() != '':
                        quantidade_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {details["QUANTIDADE"]}</div>'

                    # CORREÇÃO: Todas as chaves .get() agora estão em MAIÚSCULAS
                    card_html = f"""
                    <div class="card-container">
                        <div class="card-title">{details.get("UG", "N/A")}</div>
                        <div class="card-item"><span class="card-label">Categoria:</span> {details.get("Categoria", "")}</div>
                        <div class="card-item"><span class="card-label">Tipo de Ocorrência:</span> {details.get("TIPO DE OCORRÊNCIA", "")}</div>
                        <div class="card-item"><span class="card-label">Ativo:</span> {details.get("ATIVO", "")}</div>
                        <div class="card-item"><span class="card-label">Nome do ativo:</span> {details.get("NOME ATIVO", "")}</div>
                        <div class="card-item"><span class="card-label">Ocorrência:</span> {details.get("OCORRÊNCIA", "")}</div>
                        {quantidade_html}
                        <br>
                        <div class="card-item"><span class="card-label">Data da ocorrência:</span> {data_ocor}</div>
                        <div class="card-item"><span class="card-label">Hora da ocorrência:</span> {hora_ocor}</div>
                        <div class="card-item"><span class="card-label">Data do atendimento LOOP:</span> {data_loop}</div>
                        <div class="card-item"><span class="card-label">Hora do atendimento LOOP:</span> {hora_loop}</div>
                        <div class="card-item"><span class="card-label">Data do atendimento de terceiros:</span> {data_terc}</div>
                        <div class="card-item"><span class="card-label">Hora do atendimento de terceiros:</span> {hora_terc}</div>
                        <div class="card-item"><span class="card-label">Data de normalização:</span> {data_norm}</div>
                        <div class="card-item"><span class="card-label">Hora de normalização:</span> {hora_norm}</div>
                        <br>
                        <div class="card-item"><span class="card-label">Descrição:</span> {str(details.get("DESCRIÇÃO", "")).replace('\n', '<br>')}</div>
                        <div class="card-item"><span class="card-label">Protocolo:</span> {details.get("PROTOCOLO", "")}</div>
                        <div class="card-item"><span class="card-label">OS:</span> {details.get("OS", "")}</div>
                    </div>
                    """
                    st.html(card_html)
    del st.session_state['last_submission_details']

# Inicializar contador de reset (para forçar recriação de widgets)
if 'form_reset_counter' not in st.session_state:
    st.session_state.form_reset_counter = 0

# Obter o contador atual
reset_counter = st.session_state.form_reset_counter

categoria_selecionada = st.selectbox(
    "Selecione a Categoria da Ocorrência", 
    options=[PLANILHA_DESLIGAMENTOS, PLANILHA_EQUIPAMENTOS], 
    key=f'categoria_selecionada_{reset_counter}',
    index=None,
    placeholder="Selecione a categoria..."
)
st.subheader("Informações Gerais")
col1, col2 = st.columns(2)
with col1:
    cliente_selecionado = st.selectbox("Cliente", options=dados_e_opcoes.get('Cliente', []), key=f'cliente_select_{reset_counter}')
    
    op_ug = []
    if cliente_selecionado and cliente_selecionado != '-':
        cond_cliente = (df_dados['CLIENTE'] == cliente_selecionado)
        cond_ug_nao_vazia = (df_dados['UG'] != '')
        ug_filtradas = df_dados[cond_cliente & cond_ug_nao_vazia]['UG'].dropna().unique().tolist()
        op_ug = sorted(ug_filtradas)
    
    ug_selecionada = st.multiselect("UG (Unidade Geradora)", options=op_ug, key=f'ug_select_{reset_counter}')
    tipo_ocorrencia = st.selectbox("Tipo de Ocorrência", options=dados_e_opcoes.get('Tipo de ocorrência', []), key=f'tipo_ocorrencia_{reset_counter}')
    ativo = st.selectbox("Ativo", options=dados_e_opcoes.get('Ativo', []), key=f'ativo_{reset_counter}')
    
    items_para_processar = []
    if ativo and ativo != '-':
        if ativo.upper() in ['INVERSOR', 'TRACKER', 'STRING']:
            if ug_selecionada:
                df_filtrado = df_detalhado[df_detalhado['Usina'].isin(ug_selecionada)]
                col_map = {'INVERSOR': 'Inversor Conectado', 'TRACKER': 'Tracker Conectado', 'STRING': 'Nome String'}
                col_to_use = col_map.get(ativo.upper())
                if col_to_use in df_filtrado.columns:
                    opcoes_detalhadas = sorted(list(filter(None, df_filtrado[col_to_use].dropna().unique().tolist())))
                    nome_ativo_valor = st.multiselect("Nome Ativo", options=opcoes_detalhadas, key=f'nome_ativo_multi_{reset_counter}')
                    items_para_processar = nome_ativo_valor
            else: st.warning("Selecione uma UG para filtrar os ativos.")
        else:
            st.multiselect("Nome Ativo (Usinas Selecionadas)", options=ug_selecionada, disabled=True, default=ug_selecionada)
            items_para_processar = ug_selecionada
            
    if categoria_selecionada == PLANILHA_EQUIPAMENTOS:
        st.number_input("Quantidade", min_value=1, step=1, key=f'quantidade_{reset_counter}', value=1)
    
    st.session_state['items_para_processar'] = items_para_processar

with col2:
    ocorrencia = st.selectbox("Ocorrência", options=dados_e_opcoes.get('Ocorrência', []), key=f'ocorrencia_{reset_counter}')
    operador = st.selectbox("Operador", options=dados_e_opcoes.get('Operador', []), key=f'operador_{reset_counter}')
    protocolo = st.text_input("Protocolo", placeholder="Ex: 12346", key=f'protocolo_{reset_counter}')
    os_input = st.text_input("OS (Ordem de Serviço)", placeholder="Ex: OS12345", key=f'os_input_{reset_counter}')
    descricao = st.text_area("Descrição Detalhada", height=135, placeholder="Descreva a ocorrência...", key=f'descricao_{reset_counter}')

st.markdown("---")
st.subheader("Horários da Ocorrência")

eventos_map = {'Desligamento': 'desligamento', 'Cliente Avisado': 'ca', 'Atendimento Loop': 'al', 'Atendimento Terceiros': 'at', 'Normalização': 'norm'}
is_multiplos_itens = len(items_para_processar) > 1

with st.container(border=True):
    if is_multiplos_itens:
        st.markdown("**Controles para Múltiplos Itens** (marque para aplicar o mesmo valor a todos)")
        cols = st.columns(len(eventos_map))
        for i, (label, key) in enumerate(eventos_map.items()):
            with cols[i]:
                st.checkbox(f"Mesmo Dia? ({label})", key=f'mesmo_dia_{key}_{reset_counter}')
                st.checkbox(f"Mesmo Horário? ({label})", key=f'mesmo_horario_{key}_{reset_counter}')


    st.markdown("**Defina os horários abaixo:**")
    for label, key in eventos_map.items():
        if not is_multiplos_itens or st.session_state.get(f'mesmo_dia_{key}_{reset_counter}') or st.session_state.get(f'mesmo_horario_{key}_{reset_counter}'):
            with st.container():
                cols = st.columns([2, 1, 1])
                cols[0].markdown(f"**{label}**")
                default_date = None
                default_time = None
                if not is_multiplos_itens or st.session_state.get(f'mesmo_dia_{key}_{reset_counter}'):
                    cols[1].date_input(f"Data {label}", value=default_date, key=f'data_{key}_master_{reset_counter}', label_visibility="collapsed", format="DD/MM/YYYY")
                if not is_multiplos_itens or st.session_state.get(f'mesmo_horario_{key}_{reset_counter}'):
                    cols[2].time_input(f"Hora {label}", value=default_time, key=f'hora_{key}_master_{reset_counter}', label_visibility="collapsed")

show_specific_times = is_multiplos_itens and any(not st.session_state.get(f'mesmo_dia_{key}_{reset_counter}') or not st.session_state.get(f'mesmo_horario_{key}_{reset_counter}') for key in eventos_map.values())

if show_specific_times:
    with st.expander("**Preencher Datas/Horários Específicos**", expanded=True):
        for item in items_para_processar:
            st.markdown(f"--- \n **{item}**"); item_key = sanitize_key(item)
            for label, key in eventos_map.items():
                if not st.session_state.get(f'mesmo_dia_{key}_{reset_counter}') or not st.session_state.get(f'mesmo_horario_{key}_{reset_counter}'):
                    cols = st.columns([2, 1, 1]); cols[0].markdown(f"{label}:")
                    if not st.session_state.get(f'mesmo_dia_{key}_{reset_counter}'): cols[1].date_input(f"Data Específica {label}", value=None, key=f'data_{key}_{item_key}_{reset_counter}', label_visibility="collapsed", format="DD/MM/YYYY")
                    if not st.session_state.get(f'mesmo_horario_{key}_{reset_counter}'): cols[2].time_input(f"Hora Específica {label}", value=None, key=f'hora_{key}_{item_key}_{reset_counter}', label_visibility="collapsed")

st.markdown("---")
if st.button('Adicionar Ocorrência', type="primary", use_container_width=True):
    
    def find_ug_for_ativo(ativo_nome, df_detalhado_cache, ugs_filtradas):
        df_filtrado = df_detalhado_cache[df_detalhado_cache['Usina'].isin(ugs_filtradas)]
        for col_name in ['Inversor Conectado', 'Tracker Conectado', 'Nome String']:
            if col_name in df_filtrado.columns:
                match = df_filtrado[df_filtrado[col_name] == ativo_nome]
                if not match.empty: return match['Usina'].iloc[0]
        return None

    iter_list = st.session_state.get('items_para_processar', [])
    if not iter_list: 
        st.error("Por favor, selecione uma ou mais UGs ou Nomes de Ativo.")
    else:
        ocorrencias_para_salvar = []
        ativo_selecionado = st.session_state.get(f'ativo_{reset_counter}')
        ugs_selecionadas_no_form = st.session_state.get('ug_select', [])
        is_multi = len(iter_list) > 1
        
        erro_encontrado = False
        for item in iter_list:
            item_key_sanitized = sanitize_key(item)
            ug_final, nome_ativo_para_salvar = (None, item)

            if ativo_selecionado.upper() in ['INVERSOR', 'TRACKER', 'STRING']:
                ug_final = find_ug_for_ativo(item, df_detalhado, ugs_selecionadas_no_form)
            else:
                ug_final, nome_ativo_para_salvar = (item, item)
            
            if not ug_final:
                st.error(f"Não foi possível determinar a UG para o item '{item}'. A ocorrência não foi salva.")
                erro_encontrado = True; continue

            client_df = df_dados[df_dados['UG'] == ug_final]
            if client_df.empty:
                st.error(f"Não foi possível encontrar dados (Cliente, Sigla) para a UG '{ug_final}'.")
                erro_encontrado = True; continue
            
            cliente_final = client_df['CLIENTE'].iloc[0]
            sigla_final = client_df['SIGLA'].iloc[0]
            
            ocorrencia_base = {
                'CLIENTE': cliente_final, 'UG': ug_final, 'SIGLA': sigla_final,
                'TIPO DE OCORRÊNCIA': st.session_state.get(f'tipo_ocorrencia_{reset_counter}'),
                'ATIVO': st.session_state.get(f'ativo_{reset_counter}'),
                'NOME ATIVO': nome_ativo_para_salvar, 
                'OCORRÊNCIA': st.session_state.get(f'ocorrencia_{reset_counter}'),
                'OPERADOR': st.session_state.get(f'operador_{reset_counter}'), 
                'DESCRIÇÃO': st.session_state.get(f'descricao_{reset_counter}'),
                'PROTOCOLO': st.session_state.get(f'protocolo_{reset_counter}'),
                'OS': st.session_state.get(f'os_input_{reset_counter}')
            }
            if st.session_state.get(f'categoria_selecionada_{reset_counter}') == PLANILHA_EQUIPAMENTOS:
                ocorrencia_base['QUANTIDADE'] = st.session_state.get('quantidade', 1)

            for nome_evento, key_evento in eventos_map.items():
                data = st.session_state.get(f'data_{key_evento}_master_{reset_counter}') if (is_multi and st.session_state.get(f'mesmo_dia_{key_evento}')) or not is_multi else st.session_state.get(f'data_{key_evento}_{item_key_sanitized}_{reset_counter}')
                hora = st.session_state.get(f'hora_{key_evento}_master_{reset_counter}') if (is_multi and st.session_state.get(f'mesmo_horario_{key_evento}')) or not is_multi else st.session_state.get(f'hora_{key_evento}_{item_key_sanitized}_{reset_counter}')
                
                if data and hora:
                    ocorrencia_base[nome_evento.upper()] = datetime.combine(data, hora).strftime('%d/%m/%Y %H:%M:%S')
                else:
                    ocorrencia_base[nome_evento.upper()] = ''
            
            ocorrencias_para_salvar.append(ocorrencia_base)

        if not erro_encontrado and ocorrencias_para_salvar:
            try:
                client = connect_to_google_sheets()
                workbook = client.open_by_url(SPREADSHEET_URL)
                worksheet = workbook.worksheet(st.session_state.get(f'categoria_selecionada_{reset_counter}'))
                colunas_planilha = worksheet.row_values(1)
                
                COLUNAS_EDITAVEIS = [
                    "UG", "TIPO DE OCORRÊNCIA", "ATIVO", "NOME ATIVO", "OCORRÊNCIA", 
                    "OPERADOR", "DESLIGAMENTO", "CLIENTE AVISADO", "ATENDIMENTO LOOP", 
                    "ATENDIMENTO TERCEIROS", "NORMALIZAÇÃO", "DESCRIÇÃO", "PROTOCOLO", "OS",
                    "QUANTIDADE"
                ]

                ug_col_index = colunas_planilha.index('UG') + 1
                ug_column_values = worksheet.col_values(ug_col_index)
                start_row = len(ug_column_values) + 1
                try:
                    start_row = ug_column_values.index('') + 1
                except ValueError:
                    pass

                linhas_para_adicionar = []
                for occ in ocorrencias_para_salvar:
                    linha_ordenada = []
                    for col_header in colunas_planilha:
                        col_strip = col_header.strip()
                        if col_strip in COLUNAS_EDITAVEIS:
                            valor = occ.get(col_strip, '')
                            linha_ordenada.append('' if valor == '-' else valor)
                        else:
                            linha_ordenada.append(None)
                    linhas_para_adicionar.append(linha_ordenada)
                
                if linhas_para_adicionar:
                    worksheet.update(f'A{start_row}', linhas_para_adicionar, value_input_option='USER_ENTERED')
                    
                    # O dicionário 'ocorrencias_para_salvar' já está no formato correto.
                    for item_dict in ocorrencias_para_salvar:
                        item_dict['Categoria'] = st.session_state.get(f'categoria_selecionada_{reset_counter}')

                    st.session_state.last_submission_details = ocorrencias_para_salvar

                    # Limpar TODOS os campos do formulário (exceto last_submission_details e form_reset_counter)
                    keys_to_preserve = {'last_submission_details', 'form_reset_counter'}
                    keys_to_delete = [k for k in list(st.session_state.keys()) if k not in keys_to_preserve]

                    for key in keys_to_delete:
                        del st.session_state[key]

                    # Incrementar contador para forçar recriação de todos os widgets
                    st.session_state.form_reset_counter += 1


                    st.rerun()

            except Exception as e:
                st.error(f"Ocorreu um erro ao salvar na Planilha Google: {e}")