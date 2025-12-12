import os
import streamlit as st
import pandas as pd
from datetime import datetime, date, time
import time as pytime
import gspread
from google.oauth2.service_account import Credentials
import html
import utils

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(layout="wide")

# [NOVO] Aplica o tema visual corrigido
utils.render_page_config_and_css()

# Estado mínimo do overlay
if 'ui_phase' not in st.session_state:
    st.session_state.ui_phase = 'ready'
if 'loading_ts' not in st.session_state:
    st.session_state.loading_ts = 0

utils.render_loading_overlay(st.session_state.ui_phase)

def overlay_on():
    st.session_state.ui_phase = 'loading'
    st.session_state.loading_ts = pytime.time()
    utils.render_loading_overlay('loading')

def overlay_off():
    st.session_state.ui_phase = 'ready'
    st.session_state.loading_ts = 0
    utils.render_loading_overlay('ready')

utils.render_loading_overlay('ready')

# CSS dos Cards
st.markdown("""
<style>
    .card-container {
        background-color: #FF4B4B;
        color: white;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 15px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
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

st.title("📝 Editar Ocorrência")

# --- FEEDBACK VISUAL APÓS EDIÇÃO ---
if 'edited_occurrence_feedback' in st.session_state:
    st.success("Ocorrência atualizada com sucesso!")
    d = st.session_state.pop('edited_occurrence_feedback')
    
    def _fmt(val):
        if not val: return ''
        if isinstance(val, str): return val
        return val.strftime('%d/%m/%Y %H:%M')

    st.markdown(f"""
    <div class="card-container">
        <div class="card-title">{d.get('UG', '')}</div>
        <div class="card-item"><span class="card-label">Cliente:</span> {d.get('Cliente', '')}</div>
        <div class="card-item"><span class="card-label">Ativo:</span> {d.get('Ativo', '')}</div>
        <div class="card-item"><span class="card-label">Nome Ativo:</span> {d.get('Nome Ativo', '')}</div>
        <div class="card-item"><span class="card-label">Tipo Ocorrência:</span> {d.get('Tipo de ocorrência', '')}</div>
        <div class="card-item"><span class="card-label">Ocorrência:</span> {d.get('Ocorrência', '')}</div>
        <div class="card-item"><span class="card-label">Operador:</span> {d.get('Operador', '')}</div>
        <br>
        <div class="card-item"><span class="card-label">Normalização:</span> {_fmt(d.get('Normalização'))}</div>
        <div class="card-item"><span class="card-label">Descrição:</span> {d.get('Descrição', '')}</div>
    </div>
    """, unsafe_allow_html=True)

MAPA_RENOMEAR = {
    'IDENTIFICADOR': 'Identificador', 'CLIENTE': 'Cliente', 'UG': 'UG', 'TIPO DE OCORRÊNCIA': 'Tipo de ocorrência',
    'ATIVO': 'Ativo', 'NOME ATIVO': 'Nome Ativo', 'OCORRÊNCIA': 'Ocorrência',
    'QUANTIDADE': 'Quantidade', 'SIGLA': 'Sigla', 'NORMALIZAÇÃO': 'Normalização',
    'DESLIGAMENTO': 'Desligamento', 'OPERADOR': 'Operador', 'DESCRIÇÃO': 'Descrição',
    'OS': 'OS', 'ATENDIMENTO LOOP': 'Atendimento Loop',
    'ATENDIMENTO TERCEIROS': 'Atendimento Terceiros', 'PROTOCOLO': 'Protocolo', 'CLIENTE AVISADO': 'Cliente Avisado'
}

@st.cache_data(ttl=60)
def carregar_dados_completos():
    try:
        client = utils.connect_to_google_sheets()
        workbook = client.open_by_url(utils.SPREADSHEET_URL)
        df_desligamentos = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_DESLIGAMENTOS))
        df_equipamentos = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_EQUIPAMENTOS))

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
        client = utils.connect_to_google_sheets()
        workbook = client.open_by_url(utils.SPREADSHEET_URL)
        df_dados = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_DADOS)).fillna('')

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
    if pd.notna(dt_obj) and isinstance(dt_obj, (datetime, pd.Timestamp)):
        return dt_obj.date(), dt_obj.time()
    return None, None

# --- 3. INTERFACE DO STREAMLIT ---
overlay_on()

def _stop_with_overlay_off(msg: str | None = None, kind: str = "warning"):
    if msg:
        if kind == "warning": st.warning(msg)
        elif kind == "error": st.error(msg)
        elif kind == "info": st.info(msg)
    overlay_off()
    st.stop()

try:
    if ('id_unico_para_editar' not in st.session_state) or (not st.session_state['id_unico_para_editar']):
        df_lista = st.session_state.get('df_lista_para_editar')

        if df_lista is not None and not df_lista.empty:
            if 'Display' not in df_lista.columns:
                if 'Desligamento' in df_lista.columns:
                    df_tmp = df_lista.copy()
                    if not pd.api.types.is_datetime64_any_dtype(df_tmp['Desligamento']):
                        df_tmp['Desligamento'] = pd.to_datetime(df_tmp['Desligamento'], errors='coerce')
                    disp_data = df_tmp['Desligamento'].dt.strftime('%d/%m/%Y %H:%M').fillna('')
                else:
                    disp_data = pd.Series([''] * len(df_lista))

                df_lista['Display'] = (
                    df_lista.get('UG', '').astype(str) + " | " +
                    df_lista.get('Ativo', '').astype(str) + " | " +
                    df_lista.get('Nome Ativo', '').astype(str) + " | " +
                    df_lista.get('Ocorrência', '').astype(str) + " | " +
                    disp_data.astype(str)
                )

            st.subheader("Lista de Ocorrências (da página principal)")
            st.dataframe(
                df_lista.drop(columns=[c for c in ['ID_Unico'] if c in df_lista.columns]),
                use_container_width=True
            )

            options = df_lista['Display'].tolist()
            occ_disp = st.selectbox(
                "Selecione a ocorrência para editar:",
                options=options,
                index=None,
                placeholder="Escolha uma ocorrência..."
            )
            if occ_disp:
                sel = df_lista.loc[df_lista['Display'] == occ_disp].iloc[0]
                st.session_state['id_unico_para_editar'] = sel['ID_Unico']
                overlay_off()
                st.rerun()

            if not st.session_state.get('id_unico_para_editar'):
                _stop_with_overlay_off("Nenhuma ocorrência selecionada para edição.")

        else:
            df_full = carregar_dados_completos()
            if df_full.empty:
                _stop_with_overlay_off("Falha ao carregar dados para montar a lista de edição.", kind="error")

            df_tmp = df_full.copy()
            if 'Desligamento' in df_tmp.columns and not pd.api.types.is_datetime64_any_dtype(df_tmp['Desligamento']):
                df_tmp['Desligamento'] = pd.to_datetime(df_tmp['Desligamento'], errors='coerce')
            disp_data = df_tmp['Desligamento'].dt.strftime('%d/%m/%Y %H:%M').fillna('') if 'Desligamento' in df_tmp.columns else ''
            df_full['Display'] = (
                df_full.get('UG','').astype(str) + " | " +
                df_full.get('Ativo','').astype(str) + " | " +
                df_full.get('Nome Ativo','').astype(str) + " | " +
                df_full.get('Ocorrência','').astype(str) + " | " + disp_data.astype(str)
            )

            st.subheader("Selecione a ocorrência para editar")
            options = df_full['Display'].tolist()
            occ_disp = st.selectbox(
                "Ocorrências",
                options=options,
                index=None,
                placeholder="Escolha uma ocorrência..."
            )
            if occ_disp:
                sel = df_full.loc[df_full['Display'] == occ_disp].iloc[0]
                st.session_state['id_unico_para_editar'] = sel['ID_Unico']
                overlay_off()
                st.rerun()
            else:
                overlay_off()
                st.info("Escolha uma ocorrência para prosseguir.")
                st.stop()

    id_para_editar = st.session_state['id_unico_para_editar']
    df_completo = carregar_dados_completos()
    opcoes_edicao = carregar_opcoes_para_edicao()

    if df_completo.empty or not opcoes_edicao:
        _stop_with_overlay_off("Falha ao carregar dados completos ou listas de opções para edição.", kind="error")

    dados_ocorrencia = df_completo[df_completo['ID_Unico'] == id_para_editar]
    if dados_ocorrencia.empty:
        _stop_with_overlay_off("O ID selecionado não foi encontrado nos dados carregados.", kind="error")

    ocorrencia = dados_ocorrencia.iloc[0].to_dict()
    categoria = ocorrencia.get('Categoria', 'DESLIGAMENTOS')

    overlay_off()

    with st.form("edit_form"):
        st.subheader(f"Editando Ocorrência em: {categoria}")

        col1, col2 = st.columns(2)
        with col1:
            st.text_input("UG", value=ocorrencia.get('UG'), key="ug")
            st.text_input("Nome Ativo", value=ocorrencia.get('Nome Ativo'), key="nome_ativo")

            tipo_ocorrencia_opts = opcoes_edicao.get('tipos_ocorrencia', [])
            tipo_idx = tipo_ocorrencia_opts.index(ocorrencia.get('Tipo de ocorrência')) if ocorrencia.get('Tipo de ocorrência') in tipo_ocorrencia_opts else 0
            st.selectbox("Tipo de Ocorrência", options=tipo_ocorrencia_opts, index=tipo_idx, key="tipo_ocorrencia")

            ocorrencia_opts = opcoes_edicao.get('ocorrencias', [])
            ocorrencia_idx = ocorrencia_opts.index(ocorrencia.get('Ocorrência')) if ocorrencia.get('Ocorrência') in ocorrencia_opts else 0
            st.selectbox("Ocorrência", options=ocorrencia_opts, index=ocorrencia_idx, key="ocorrencia")

            operador_opts = opcoes_edicao.get('operadores', [])
            operador_idx = operador_opts.index(ocorrencia.get('Operador')) if ocorrencia.get('Operador') in operador_opts else 0
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
                client = utils.connect_to_google_sheets()
                workbook = client.open_by_url(utils.SPREADSHEET_URL)
                worksheet = workbook.worksheet(categoria)
                all_data = worksheet.get_all_values()
                headers = all_data[0]

                headers_map = {h.strip().upper(): i for i, h in enumerate(headers)}
                idx_ug = headers_map['UG']
                idx_ativo = headers_map['ATIVO']
                idx_ocorrencia = headers_map['OCORRÊNCIA']
                idx_desligamento = headers_map['DESLIGAMENTO']

                row_to_edit = -1
                for i, row in enumerate(all_data[1:], start=2):
                    try:
                        desligamento_dt = pd.to_datetime(row[idx_desligamento], errors='coerce')
                        if pd.isna(desligamento_dt): continue
                        current_id = f"{row[idx_ug].upper()}|{row[idx_ativo].upper()}|{row[idx_ocorrencia].upper()}|{desligamento_dt}"
                        if current_id == id_para_editar:
                            row_to_edit = i
                            break
                    except (IndexError, ValueError):
                        continue

                if row_to_edit == -1:
                    _stop_with_overlay_off("Não foi possível localizar a linha correspondente na planilha.", kind="error")

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

                linha_para_atualizar = []
                for h in headers:
                    h_strip = h.strip()
                    key_title_case = MAPA_RENOMEAR.get(h_strip.upper(), h_strip)
                    valor = dados_atualizados.get(key_title_case, '')
                    if isinstance(valor, (datetime, pd.Timestamp)):
                        valor = valor.strftime('%Y-%m-%d %H:%M:%S')
                    linha_para_atualizar.append(valor)

                worksheet.update(f'A{row_to_edit}', [linha_para_atualizar], value_input_option='USER_ENTERED')

                st.session_state['edited_occurrence_feedback'] = dados_atualizados
                st.session_state.pop('id_unico_para_editar', None)
                st.cache_data.clear()
                
                st.rerun()

            except Exception as e:
                st.error(f"Ocorreu um erro ao atualizar a Planilha Google: {e}")

finally:
    overlay_off()