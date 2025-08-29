import streamlit as st
import pandas as pd
from datetime import datetime
import os
import win32com.client
import pythoncom
import re

st.set_page_config(layout="wide")

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

def sanitize_key(text):
    return re.sub(r'[^A-Za-z0-9_]', '_', str(text))

# --- Constantes e Mapas de Colunas ---
ARQUIVO_DADOS = 'Aviso de Anomalias 2025.xlsx'
PLANILHA_DESLIGAMENTOS = 'DESLIGAMENTOS'
PLANILHA_EQUIPAMENTOS = 'EQUIPAMENTOS'
PLANILHA_DADOS = 'DADOS'
PLANILHA_DETALHADA = 'Usinas_Detalhado'

MAPA_DESLIGAMENTOS = {
    'UG': 'C', 'Tipo de ocorrência': 'E', 'Ativo': 'F', 'Nome Ativo': 'G', 
    'Ocorrência': 'H', 'Operador': 'I', 'Desligamento': 'J', 'Cliente Avisado': 'K', 
    'Atendimento Loop': 'L', 'Atendimento Terceiros': 'M', 'Normalização': 'N', 
    'Descrição': 'O', 'Protocolo': 'P', 'OS': 'Q'
}

# --- CORREÇÃO APLICADA AQUI ---
MAPA_EQUIPAMENTOS = {
    'UG': 'C', 'Tipo de ocorrência': 'E', 'Ativo': 'F', 'Nome Ativo': 'G', 
    'Ocorrência': 'H', # Estava na coluna I
    'Quantidade': 'I', # Estava na coluna H
    'Operador': 'J', 
    'Desligamento': 'K', 
    'Cliente Avisado': 'L', 
    'Atendimento Loop': 'M', 
    'Atendimento Terceiros': 'N', 
    'Normalização': 'O', 
    'Descrição': 'P', 
    'Protocolo': 'Q', 
    'OS': 'R'
}


@st.cache_data(ttl=60)
def carregar_dados_e_opcoes():
    try:
        df_dados = pd.read_excel(ARQUIVO_DADOS, sheet_name=PLANILHA_DADOS, dtype=str).fillna('')
        df_detalhado = pd.read_excel(ARQUIVO_DADOS, sheet_name=PLANILHA_DETALHADA, dtype=str).fillna('')
        for col in df_dados.columns:
            if df_dados[col].dtype == 'object': df_dados[col] = df_dados[col].str.strip()
        for col in df_detalhado.columns:
            if df_detalhado[col].dtype == 'object': df_detalhado[col] = df_detalhado[col].str.strip()
        
        op_cliente = ['-'] + sorted(df_dados['CLIENTE'].dropna().unique().tolist())
        op_ocorrencia = ['-'] + sorted(df_dados['OCORRÊNCIA'].dropna().unique().tolist())
        op_tipo = ['-'] + sorted(df_dados['TIPO DE OCORRÊNCIA'].dropna().unique().tolist())
        op_ativo = ['-'] + sorted(df_dados['ATIVO'].dropna().unique().tolist())
        op_operador = ['-'] + sorted(df_dados['OPERADOR'].dropna().unique().tolist())
        
        return { 'df_dados': df_dados, 'df_detalhado': df_detalhado, 'Cliente': op_cliente,
                 'Ocorrência': op_ocorrencia, 'Tipo de ocorrência': op_tipo, 'Ativo': op_ativo,
                 'Operador': op_operador }
    except Exception as e:
        st.error(f"Erro ao carregar os dados das planilhas: {e}"); return {}

dados_e_opcoes = carregar_dados_e_opcoes()
if not dados_e_opcoes: st.stop()
df_dados = dados_e_opcoes.get('df_dados')
df_detalhado = dados_e_opcoes.get('df_detalhado')

def find_first_empty_row(sheet, map_colunas):
    col_letra_ug = map_colunas['UG']
    col_num = sheet.Range(f"{col_letra_ug}1").Column
    for row_idx in range(2, 20000):
        cell_value = sheet.Cells(row_idx, col_num).Value
        if cell_value is None or cell_value == 0 or cell_value == '': return row_idx
    return sheet.Cells(sheet.Rows.Count, col_num).End(-4162).Row + 1

def write_data_to_row(sheet, data_dict, row_number, map_colunas):
    COLUNAS_PROTEGIDAS = ['Cliente', 'Sigla']
    for key, value in data_dict.items():
        if key not in COLUNAS_PROTEGIDAS and key in map_colunas:
            col_letter = map_colunas[key]
            sheet.Range(f'{col_letter}{row_number}').Value = value

def format_datetime_card(dt_obj):
    """Formata um objeto datetime OU UMA STRING de data em data e hora para o card."""
    # Se já for um objeto datetime, formata diretamente
    if isinstance(dt_obj, datetime):
        return dt_obj.strftime('%d/%m/%Y'), dt_obj.strftime('%H:%M')
    
    # Se for um texto (string) e não estiver vazio, tenta converter e formatar
    if isinstance(dt_obj, str) and dt_obj:
        try:
            # Tenta converter o texto para um objeto datetime
            dt = pd.to_datetime(dt_obj)
            # Se conseguir, formata e retorna
            return dt.strftime('%d/%m/%Y'), dt.strftime('%H:%M')
        except (ValueError, TypeError):
            # Se a conversão falhar, retorna vazio
            return '', ''
            
    # Para qualquer outra coisa (None, NaT, etc.), retorna vazio
    return '', ''

st.title("Adicionar Nova Ocorrência")

if 'last_submission_details' in st.session_state and st.session_state.last_submission_details:
    # Agora estamos recebendo uma lista de ocorrências
    submitted_occurrences = st.session_state.last_submission_details
    
    st.success(f"{len(submitted_occurrences)} ocorrência(s) adicionada(s) com sucesso!")
    
    # Define o número de colunas para os cards, controlando a largura
    num_cols = 4 
    
    # Itera pela lista de ocorrências e cria um card para cada uma
    for i in range(0, len(submitted_occurrences), num_cols):
        cols = st.columns(num_cols)
        for j in range(num_cols):
            if i + j < len(submitted_occurrences):
                with cols[j]:
                    details = submitted_occurrences[i + j]

                    # Formata as datas e horas para exibição
                    data_ocor, hora_ocor = format_datetime_card(details.get('Desligamento'))
                    data_loop, hora_loop = format_datetime_card(details.get('Atendimento Loop'))
                    data_terc, hora_terc = format_datetime_card(details.get('Atendimento Terceiros'))
                    data_norm, hora_norm = format_datetime_card(details.get('Normalização'))

                    quantidade_html = ""
                    if "Quantidade" in details and details['Quantidade']:
                        quantidade_html = f'<div class="card-item"><span class="card-label">Quantidade:</span> {details["Quantidade"]}</div>'

                    # Estrutura HTML idêntica à da página principal
                    card_html = f"""
                    <div class="card-container">
                        <div class="card-title">{details.get("UG", "N/A")}</div>
                        <div class="card-item"><span class="card-label">Categoria:</span> {details.get("Categoria", "")}</div>
                        <div class="card-item"><span class="card-label">Tipo de Ocorrência:</span> {details.get("Tipo de ocorrência", "")}</div>
                        <div class="card-item"><span class="card-label">Ativo:</span> {details.get("Ativo", "")}</div>
                        <div class="card-item"><span class="card-label">Nome do ativo:</span> {details.get("Nome Ativo", "")}</div>
                        <div class="card-item"><span class="card-label">Ocorrência:</span> {details.get("Ocorrência", "")}</div>
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
                        <div class="card-item"><span class="card-label">Descrição:</span> {str(details.get("Descrição", "")).replace('\n', '<br>')}</div>
                        <div class="card-item"><span class="card-label">Protocolo:</span> {details.get("Protocolo", "")}</div>
                        <div class="card-item"><span class="card-label">OS:</span> {details.get("OS", "")}</div>
                    </div>
                    """
                    st.html(card_html)
                    
    del st.session_state['last_submission_details']

if st.session_state.get('sucesso'):
    st.success("Dados salvos na planilha com sucesso!")
    del st.session_state['sucesso']

categoria_selecionada = st.selectbox("Selecione a Categoria da Ocorrência", options=[PLANILHA_DESLIGAMENTOS, PLANILHA_EQUIPAMENTOS], key='categoria_selecionada')
st.subheader("Informações Gerais")
col1, col2 = st.columns(2)
with col1:
    cliente_selecionado = st.selectbox("Cliente", options=dados_e_opcoes.get('Cliente', []), key='cliente_select')
    op_ug = []
    if cliente_selecionado and cliente_selecionado != '-':
        cond_cliente = (df_dados['CLIENTE'] == cliente_selecionado); cond_ug_nao_vazia = (df_dados['UG'] != '')
        ug_filtradas = df_dados[cond_cliente & cond_ug_nao_vazia]['UG'].dropna().unique().tolist()
        op_ug = sorted(ug_filtradas)
    ug_selecionada = st.multiselect("UG (Unidade Geradora)", options=op_ug, key='ug_select')
    tipo_ocorrencia = st.selectbox("Tipo de Ocorrência", options=dados_e_opcoes.get('Tipo de ocorrência', []), key='tipo_ocorrencia')
    ativo = st.selectbox("Ativo", options=dados_e_opcoes.get('Ativo', []), key='ativo')
    items_para_processar = []
    if ativo and ativo != '-':
        if ativo.upper() in ['INVERSOR', 'TRACKER', 'STRING']:
            if ug_selecionada:
                df_filtrado = df_detalhado[df_detalhado['Usina'].isin(ug_selecionada)]
                col_map = {'INVERSOR': 'Inversor Conectado', 'TRACKER': 'Tracker Conectado', 'STRING': 'Nome String'}
                col_to_use = col_map.get(ativo.upper())
                opcoes_detalhadas = sorted(list(filter(None, df_filtrado[col_to_use].dropna().unique().tolist()))) if col_to_use else []
                nome_ativo_valor = st.multiselect("Nome Ativo", options=opcoes_detalhadas, key='nome_ativo_multi')
                items_para_processar = nome_ativo_valor
            else: st.warning("Selecione uma UG para filtrar os ativos.")
        else:
            st.multiselect("Nome Ativo (Usinas Selecionadas)", options=ug_selecionada, disabled=True, default=ug_selecionada)
            items_para_processar = ug_selecionada
    if categoria_selecionada == PLANILHA_EQUIPAMENTOS:
        st.number_input("Quantidade", min_value=1, step=1, key='quantidade', value=1)
    st.session_state['items_para_processar'] = items_para_processar
with col2:
    ocorrencia = st.selectbox("Ocorrência", options=dados_e_opcoes.get('Ocorrência', []), key='ocorrencia')
    operador = st.selectbox("Operador", options=dados_e_opcoes.get('Operador', []), key='operador')
    protocolo = st.text_input("Protocolo", placeholder="Ex: 12346", key='protocolo')
    os_input = st.text_input("OS (Ordem de Serviço)", placeholder="Ex: OS12345", key='os_input')
    descricao = st.text_area("Descrição Detalhada", height=135, placeholder="Descreva a ocorrência...", key='descricao')

st.markdown("---")
st.subheader("Horários da Ocorrência")

eventos_map = {'Desligamento': 'desligamento', 'Cliente Avisado': 'ca', 'Atendimento Loop': 'al', 'Atendimento Terceiros': 'at', 'Normalização': 'norm'}
is_multiplos_itens = len(items_para_processar) > 1

with st.container(border=True):
    if is_multiplos_itens:
        st.markdown("**Controles para Múltiplos Itens** (marque para aplicar o mesmo valor a todos)")
        cols = st.columns(5)
        for i, (label, key) in enumerate(eventos_map.items()):
            with cols[i]:
                st.checkbox(f"Mesmo Dia? ({label})", key=f'mesmo_dia_{key}')
                st.checkbox(f"Mesmo Horário? ({label})", key=f'mesmo_horario_{key}')

    st.markdown("**Defina os horários abaixo:**")
    for label, key in eventos_map.items():
        if not is_multiplos_itens or st.session_state.get(f'mesmo_dia_{key}') or st.session_state.get(f'mesmo_horario_{key}'):
            with st.container():
                cols = st.columns([2, 1, 1])
                cols[0].markdown(f"**{label}**")
                if not is_multiplos_itens or st.session_state.get(f'mesmo_dia_{key}'):
                    cols[1].date_input(f"Data {label}", key=f'data_{key}_master', label_visibility="collapsed")
                if not is_multiplos_itens or st.session_state.get(f'mesmo_horario_{key}'):
                    cols[2].time_input(f"Hora {label}", value=None, key=f'hora_{key}_master', label_visibility="collapsed")

show_specific_times = is_multiplos_itens and any(not st.session_state.get(f'mesmo_dia_{key}') or not st.session_state.get(f'mesmo_horario_{key}') for key in eventos_map.values())

if show_specific_times:
    with st.expander("**Preencher Datas/Horários Específicos**", expanded=True):
        for item in items_para_processar:
            st.markdown(f"--- \n **{item}**"); item_key = sanitize_key(item)
            for label, key in eventos_map.items():
                if not st.session_state.get(f'mesmo_dia_{key}') or not st.session_state.get(f'mesmo_horario_{key}'):
                    cols = st.columns([2, 1, 1]); cols[0].markdown(f"{label}:")
                    if not st.session_state.get(f'mesmo_dia_{key}'): cols[1].date_input(f"Data Específica {label}", value=None, key=f'data_{key}_{item_key}', label_visibility="collapsed")
                    if not st.session_state.get(f'mesmo_horario_{key}'): cols[2].time_input(f"Hora Específica {label}", value=None, key=f'hora_{key}_{item_key}', label_visibility="collapsed")

st.markdown("---")
if st.button('Adicionar Ocorrência', type="primary", use_container_width=True):
    @st.cache_data
    def find_ug_for_ativo(ativo_nome, df_detalhado_cache, ugs_filtradas):
        df_filtrado = df_detalhado_cache[df_detalhado_cache['Usina'].isin(ugs_filtradas)]
        for col_name in ['Inversor Conectado', 'Tracker Conectado', 'Nome String']:
            match = df_filtrado[df_filtrado[col_name] == ativo_nome]
            if not match.empty: return match['Usina'].iloc[0]
        return None

    iter_list = st.session_state.get('items_para_processar', [])
    if not iter_list: st.error("Por favor, selecione uma ou mais UGs ou Nomes de Ativo.")
    else:
        ocorrencias_para_salvar = []
        ativo_selecionado = st.session_state.ativo.upper()
        ugs_selecionadas_no_form = st.session_state.get('ug_select', [])
        is_multi = len(iter_list) > 1
        
        for item in iter_list:
            item_key_sanitized = sanitize_key(item)
            ug_final, nome_ativo_para_salvar = None, item
            if ativo_selecionado in ['INVERSOR', 'TRACKER', 'STRING']:
                ug_final = find_ug_for_ativo(item, df_detalhado, ugs_selecionadas_no_form)
                if ug_final is None and ugs_selecionadas_no_form: ug_final = ugs_selecionadas_no_form[0]
            else:
                ug_final, nome_ativo_para_salvar = item, item
            
            ocorrencia_base = {'UG': ug_final, 'Tipo de ocorrência': st.session_state.tipo_ocorrencia.upper(), 'Ativo': ativo_selecionado, 'Nome Ativo': nome_ativo_para_salvar, 'Ocorrência': st.session_state.ocorrencia.upper(), 'Operador': st.session_state.operador, 'Descrição': st.session_state.descricao, 'Protocolo': st.session_state.protocolo, 'OS': st.session_state.os_input}
            if st.session_state.categoria_selecionada == PLANILHA_EQUIPAMENTOS:
                ocorrencia_base['Quantidade'] = st.session_state.get('quantidade', 1)

            for nome_evento, key_evento in eventos_map.items():
                mesmo_dia_final = st.session_state.get(f'mesmo_dia_{key_evento}') if is_multi else True
                mesmo_horario_final = st.session_state.get(f'mesmo_horario_{key_evento}') if is_multi else True

                data = st.session_state.get(f'data_{key_evento}_master') if mesmo_dia_final else st.session_state.get(f'data_{key_evento}_{item_key_sanitized}')
                hora = st.session_state.get(f'hora_{key_evento}_master') if mesmo_horario_final else st.session_state.get(f'hora_{key_evento}_{item_key_sanitized}')
                
                if data and hora:
                    full_datetime = datetime.combine(data, hora)
                    ocorrencia_base[nome_evento] = full_datetime.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    ocorrencia_base[nome_evento] = None
            
            ocorrencias_para_salvar.append(ocorrencia_base)

        mapa_colunas_ativo = MAPA_EQUIPAMENTOS if st.session_state.categoria_selecionada == PLANILHA_EQUIPAMENTOS else MAPA_DESLIGAMENTOS
        
        excel, workbook, sheet, pythoncom_iniciado = None, None, None, False
        try:
            pythoncom.CoInitialize()
            pythoncom_iniciado = True
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            workbook = excel.Workbooks.Open(os.path.abspath(ARQUIVO_DADOS))
            sheet = workbook.Sheets(st.session_state.categoria_selecionada)
            start_row = find_first_empty_row(sheet, mapa_colunas_ativo)
            
            sucesso_geral = True
            for index, ocorrencia_final in enumerate(ocorrencias_para_salvar):
                try: write_data_to_row(sheet, ocorrencia_final, start_row + index, mapa_colunas_ativo)
                except Exception as e:
                    st.error(f"Erro ao escrever a linha {start_row + index}: {e}"); sucesso_geral = False; break
            
            if sucesso_geral:
                    # 'ocorrencias_para_salvar' já é a lista de que precisamos.
                    # Apenas adicionamos a categoria a cada item da lista.
                    categoria_para_salvar = st.session_state.categoria_selecionada
                    for item_dict in ocorrencias_para_salvar:
                        item_dict['Categoria'] = categoria_para_salvar

                    # Salvamos a LISTA COMPLETA no estado da sessão
                    st.session_state.last_submission_details = ocorrencias_para_salvar
                    
                    # Limpa os campos do formulário para a próxima entrada
                    keys_to_keep = ['sucesso', 'last_submission_details']
                    for key in list(st.session_state.keys()):
                        if key not in keys_to_keep:
                            del st.session_state[key]
                    st.session_state.sucesso = True

        except Exception as e: st.error(f"Ocorreu um erro durante a escrita no Excel: {e}")
        finally:
            if workbook: workbook.Save(); workbook.Close(SaveChanges=True)
            if excel: excel.Quit()
            if pythoncom_iniciado: pythoncom.CoUninitialize()
            if 'sucesso' in st.session_state: st.rerun()