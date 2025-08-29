import streamlit as st
import pandas as pd
from datetime import datetime, time
import os
import win32com.client
import pythoncom

# --- FUNÇÕES AUXILIARES (BOA PRÁTICA COLOCÁ-LAS NO TOPO) ---
def combine_date_time(date_val, time_val):
    """Combina data e hora em um objeto datetime, se ambos existirem."""
    return datetime.combine(date_val, time_val) if date_val and time_val else None

def find_row_by_virtual_id(sheet, virtual_id, map_colunas):
    """
    Itera pelas linhas do Excel, recria o ID Virtual para cada uma de forma robusta
    e compara para encontrar o número da linha correspondente.
    """
    id_parts = virtual_id.split('|')
    target_ug, target_ativo, target_ocorrencia, target_desligamento_str = id_parts

    for row_idx in range(2, sheet.UsedRange.Rows.Count + 2):
        try:
            # Pega os valores da linha atual
            ug_val = sheet.Range(f"{map_colunas['UG']}{row_idx}").Value
            ativo_val = sheet.Range(f"{map_colunas['Ativo']}{row_idx}").Value
            ocorrencia_val = sheet.Range(f"{map_colunas['Ocorrência']}{row_idx}").Value
            desligamento_val = sheet.Range(f"{map_colunas['Desligamento']}{row_idx}").Value

            # Se a linha estiver vazia, para a busca
            if ug_val is None and ativo_val is None:
                break

            # Limpa e padroniza os textos para bater com o Pandas (str e uppercase)
            ug_str = str(ug_val or '').upper()
            ativo_str = str(ativo_val or '').upper()
            ocorrencia_str = str(ocorrencia_val or '').upper()

            # Trata a data de forma robusta para bater com o Pandas
            desligamento_str_excel = ""
            if hasattr(desligamento_val, 'year'):  # Verifica se é um objeto de data
                desligamento_str_excel = desligamento_val.strftime('%Y-%m-%d %H:%M:%S')
            elif desligamento_val is not None:
                # Se não for data, tenta converter da mesma forma que o Pandas faria
                temp_date = pd.to_datetime(desligamento_val, errors='coerce')
                if pd.notna(temp_date):
                    desligamento_str_excel = temp_date.strftime('%Y-%m-%d %H:%M:%S')

            # Compara o ID reconstruído com o ID alvo
            if (ug_str == target_ug and
                ativo_str == target_ativo and
                ocorrencia_str == target_ocorrencia and
                desligamento_str_excel == target_desligamento_str):
                return row_idx  # ENCONTRAMOS!

        except Exception:
            # Se der qualquer erro em uma linha, simplesmente pula para a próxima
            continue
            
    return None # Não encontrou a linha

def write_data_to_row(sheet, data_dict, row_number, map_colunas):
    for key, value in data_dict.items():
        if key in map_colunas:
            col_letter = map_colunas[key]
            if isinstance(value, datetime):
                value = value.strftime('%Y-%m-%d %H:%M:%S')
            elif value is None:
                value = ''
            sheet.Range(f'{col_letter}{row_number}').Value = value

@st.cache_data(ttl=1)
def carregar_dados_completos():
    try:
        df_d = pd.read_excel(ARQUIVO_DADOS, sheet_name=PLANILHA_DESLIGAMENTOS)
        df_e = pd.read_excel(ARQUIVO_DADOS, sheet_name=PLANILHA_EQUIPAMENTOS)
        df_d['Categoria'] = PLANILHA_DESLIGAMENTOS
        df_e['Categoria'] = PLANILHA_EQUIPAMENTOS
        df_completo = pd.concat([df_d, df_e], ignore_index=True)

        mapa_renomear = {
            'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 
            'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência', 'ATIVO': 'Ativo', 
            'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência',
            'QUANTIDADE': 'Quantidade', 'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização',
            'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 'DESCRIÇÃO': 'Descrição',
            'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop',
            'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo',
            'CLIENTE AVISADO': 'Cliente Avisado'  # <-- ADICIONADO AQUI
        }
        renomear_final = {col: mapa_renomear[col.strip().upper()] for col in df_completo.columns if col.strip().upper() in mapa_renomear}
        df_completo.rename(columns=renomear_final, inplace=True)

        df_completo.fillna('', inplace=True)
        colunas_para_padronizar = ['Cliente', 'UG', 'Tipo de ocorrência', 'Ativo', 'Ocorrência']
        for col in colunas_para_padronizar:
            if col in df_completo.columns:
                df_completo[col] = df_completo[col].astype(str).str.upper()
        
        # Converte todas as colunas de data de uma vez
        for col in ['Desligamento', 'Normalização', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']:
             if col in df_completo.columns:
                df_completo[col] = pd.to_datetime(df_completo[col], errors='coerce')

        df_completo['ID_Unico'] = df_completo['UG'].astype(str) + "|" + \
                                  df_completo['Ativo'].astype(str) + "|" + \
                                  df_completo['Ocorrência'].astype(str) + "|" + \
                                  df_completo['Desligamento'].astype(str)
        return df_completo
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(layout="wide")
st.title("📝 Editar Ocorrência")

ARQUIVO_DADOS = 'Aviso de Anomalias 2025.xlsx'
PLANILHA_DESLIGAMENTOS = 'DESLIGAMENTOS'
PLANILHA_EQUIPAMENTOS = 'EQUIPAMENTOS'

MAPA_DESLIGAMENTOS = {
    'UG': 'C', 'Tipo de ocorrência': 'E', 'Ativo': 'F', 'Nome Ativo': 'G', 
    'Ocorrência': 'H', 'Operador': 'I', 'Desligamento': 'J', 
    'Cliente Avisado': 'K',  # <-- ADICIONADO AQUI
    'Atendimento Loop': 'L', 'Atendimento Terceiros': 'M', 'Normalização': 'N', 
    'Descrição': 'O', 'Protocolo': 'P', 'OS': 'Q'
}
MAPA_EQUIPAMENTOS = {
    'UG': 'C', 'Tipo de ocorrência': 'E', 'Ativo': 'F', 'Nome Ativo': 'G', 
    'Ocorrência': 'H', 'Quantidade': 'I', 'Operador': 'J', 
    'Desligamento': 'K', 
    'Cliente Avisado': 'L',  # <-- ADICIONADO AQUI
    'Atendimento Loop': 'M', 
    'Atendimento Terceiros': 'N', 'Normalização': 'O', 
    'Descrição': 'P', 'Protocolo': 'Q', 'OS': 'R'
}

# --- LÓGICA PRINCIPAL DA PÁGINA ---
if 'id_unico_para_editar' not in st.session_state or st.session_state['id_unico_para_editar'] is None:
    st.warning("Selecione uma ocorrência na Página Principal para editar.")
    st.page_link("pages/1_Página_Principal.py", label="Voltar", icon="🏠")
else:
    id_unico = st.session_state['id_unico_para_editar']
    df_completo = carregar_dados_completos()
    
    if not df_completo.empty:
        ocorrencia_df = df_completo[df_completo['ID_Unico'] == id_unico]

        if ocorrencia_df.empty:
            st.error(f"Erro: A ocorrência com ID '{id_unico}' não foi encontrada.")
        else:
            ocorrencia_data = ocorrencia_df.iloc[0].to_dict()
            categoria = ocorrencia_data.get('Categoria')

            st.subheader(f"Editando: {ocorrencia_data.get('UG')} | {ocorrencia_data.get('Ativo')}")

            with st.form("edit_form"):
                def get_date_time_values(field_name):
                    dt_val = ocorrencia_data.get(field_name)
                    return (dt_val.date(), dt_val.time()) if pd.notna(dt_val) and isinstance(dt_val, datetime) else (None, None)

                st.write("#### Detalhes da Ocorrência")
                cols1 = st.columns(3)
                ug = cols1[0].text_input("UG", value=ocorrencia_data.get("UG"))
                ativo = cols1[1].text_input("Ativo", value=ocorrencia_data.get("Ativo"))
                nome_ativo = cols1[2].text_input("Nome do Ativo", value=ocorrencia_data.get("Nome Ativo"))

                cols2 = st.columns(3)
                tipo_ocorrencia = cols2[0].text_input("Tipo de Ocorrência", value=ocorrencia_data.get("Tipo de ocorrência"))
                ocorrencia = cols2[1].text_input("Ocorrência", value=ocorrencia_data.get("Ocorrência"))
                quantidade_val = ocorrencia_data.get("Quantidade", 1)
                quantidade = cols2[2].number_input("Quantidade", value=int(quantidade_val if pd.notna(quantidade_val) and quantidade_val != '' else 1), step=1) if categoria == PLANILHA_EQUIPAMENTOS else None
                
                st.write("#### Datas e Horários")
                d_des, t_des = get_date_time_values('Desligamento')
                cols_des = st.columns(2)
                data_desligamento = cols_des[0].date_input("Data do Desligamento", value=d_des)
                hora_desligamento = cols_des[1].time_input("Hora do Desligamento", value=t_des)
                
                d_ca, t_ca = get_date_time_values('Cliente Avisado')
                cols_ca = st.columns(2)
                data_cliente_avisado = cols_ca[0].date_input("Data Cliente Avisado", value=d_ca)
                hora_cliente_avisado = cols_ca[1].time_input("Hora Cliente Avisado", value=t_ca)

                d_loop, t_loop = get_date_time_values('Atendimento Loop')
                cols_loop = st.columns(2)
                data_loop = cols_loop[0].date_input("Data Atendimento LOOP", value=d_loop)
                hora_loop = cols_loop[1].time_input("Hora Atendimento LOOP", value=t_loop)
                
                d_terc, t_terc = get_date_time_values('Atendimento Terceiros')
                cols_terc = st.columns(2)
                data_terceiros = cols_terc[0].date_input("Data Atendimento Terceiros", value=d_terc)
                hora_terceiros = cols_terc[1].time_input("Hora Atendimento Terceiros", value=t_terc)
                
                d_norm, t_norm = get_date_time_values('Normalização')
                cols_norm = st.columns(2)
                data_normalizacao = cols_norm[0].date_input("Data de Normalização", value=d_norm)
                hora_normalizacao = cols_norm[1].time_input("Hora de Normalização", value=t_norm)

                st.write("#### Outras Informações")
                descricao = st.text_area("Descrição", value=str(ocorrencia_data.get("Descrição", "")))
                cols7 = st.columns(3)
                protocolo = cols7[0].text_input("Protocolo", value=str(ocorrencia_data.get("Protocolo", "")))
                os_val = cols7[1].text_input("OS", value=str(ocorrencia_data.get("OS", "")))
                operador = cols7[2].text_input("Operador", value=str(ocorrencia_data.get("Operador", "")))
                
                submitted = st.form_submit_button("✅ Salvar Alterações")

                if submitted:
                    updated_data = {
                        'UG': ug, 'Ativo': ativo, 'Nome Ativo': nome_ativo,
                        'Tipo de ocorrência': tipo_ocorrencia, 'Ocorrência': ocorrencia,
                        'Desligamento': combine_date_time(data_desligamento, hora_desligamento),
                        'Cliente Avisado': combine_date_time(data_cliente_avisado, hora_cliente_avisado), # <-- LINHA CORRIGIDA E ADICIONADA AQUI
                        'Atendimento Loop': combine_date_time(data_loop, hora_loop),
                        'Atendimento Terceiros': combine_date_time(data_terceiros, hora_terceiros),
                        'Normalização': combine_date_time(data_normalizacao, hora_normalizacao),
                        'Descrição': descricao, 'Protocolo': protocolo, 'OS': os_val, 'Operador': operador
                    }
                    if quantidade is not None: updated_data['Quantidade'] = quantidade

                    excel, workbook, sheet = None, None, None
                    try:
                        pythoncom.CoInitialize()
                        excel = win32com.client.Dispatch("Excel.Application")
                        workbook = excel.Workbooks.Open(os.path.abspath(ARQUIVO_DADOS))
                        sheet = workbook.Sheets(categoria)
                        
                        mapa = MAPA_EQUIPAMENTOS if categoria == PLANILHA_EQUIPAMENTOS else MAPA_DESLIGAMENTOS
                        row_to_update = find_row_by_virtual_id(sheet, id_unico, mapa)
                        
                        if row_to_update:
                            write_data_to_row(sheet, updated_data, row_to_update, mapa)
                            st.success("Ocorrência atualizada com sucesso!")
                            st.cache_data.clear()
                        else:
                            st.error("Não foi possível encontrar a linha no Excel para ATUALIZAR. Verifique a função de busca.")
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")
                    finally:
                        if 'workbook' in locals() and workbook: workbook.Save(); workbook.Close(SaveChanges=True)
                        if 'excel' in locals() and excel: excel.Quit()
                        pythoncom.CoUninitialize()

            st.page_link("pages/1_Página_Principal.py", label="Voltar", icon="🏠")