import os
import io
import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import re
import utils

# --- Configuração Inicial ---
st.set_page_config(layout="wide", page_title="Adicionar Ocorrência")

# Renderiza Overlay (Loading)
utils.render_loading_overlay(st.session_state.get('ui_phase', 'ready'))

def overlay_on():
    st.session_state.ui_phase = 'loading'
    st.session_state.loading_ts = pytime.time()
    utils.render_loading_overlay('loading')

def overlay_off():
    st.session_state.ui_phase = 'ready'
    st.session_state.loading_ts = 0
    utils.render_loading_overlay('ready')

# CSS Personalizado
st.markdown("""
<style>
    .stButton > button { background-color: #28a745; color: white; font-weight: bold; width: 100%; }
    .stButton > button:hover { background-color: #218838; }
    
    /* Estilo para as Abas */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        border-radius: 4px 4px 0px 0px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

st.title("Adicionar Nova Ocorrência")

@st.cache_data(ttl=60)
def carregar_opcoes():
    try:
        client = utils.connect_to_google_sheets()
        if not client: return {}
        
        wb = client.open_by_url(utils.SPREADSHEET_URL)
        df_dados = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_DADOS)).fillna('')
        df_detalhado = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_DETALHADA)).fillna('')
        
        def make_opt(series):
            # Adiciona o traço no início para evitar seleção automática
            return ["-"] + sorted([x for x in series.unique() if x and x != "-"])

        return {
            'df_dados': df_dados,
            'df_detalhado': df_detalhado,
            'Cliente': make_opt(df_dados['CLIENTE']),
            'Ocorrência': make_opt(df_dados['OCORRÊNCIA']),
            'Tipo de ocorrência': make_opt(df_dados['TIPO DE OCORRÊNCIA']),
            'Ativo': make_opt(df_dados['ATIVO']),
            'Operador': make_opt(df_dados['OPERADOR'])
        }
    except Exception as e:
        st.error(f"Erro: {e}")
        return {}

dados = carregar_opcoes()

# Feedback de Sucesso no Topo
if 'last_submission_details' in st.session_state:
    count = len(st.session_state.last_submission_details)
    st.success(f"✅ Sucesso! {count} ocorrência(s) adicionada(s).")
    del st.session_state['last_submission_details']

if not dados:
    st.warning("Não foi possível carregar as opções.")
    st.stop()

if 'form_reset_counter' not in st.session_state: st.session_state.form_reset_counter = 0
rc = st.session_state.form_reset_counter

# --- CATEGORIA ---
st.info("Selecione a categoria para iniciar o preenchimento.")
cat = st.selectbox("Categoria", ["-", utils.SHEET_DESLIGAMENTOS, utils.SHEET_EQUIPAMENTOS], key=f'cat_{rc}')

if cat != "-":
    st.markdown("---")
    
    # --- ABAS ---
    tab1, tab2, tab3 = st.tabs(["📋 Dados Básicos", "⚙️ Detalhes & OS", "📅 Data & Hora"])
    
    # Aba 1: Dados
    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            cli = st.selectbox("Cliente", dados['Cliente'], key=f'cliente_select_{rc}')
            
            op_ug = []
            if cli != "-":
                df_d = dados['df_dados']
                op_ug = sorted(df_d[df_d['CLIENTE'] == cli]['UG'].unique().tolist())
                
            ugs = st.multiselect("UG (Unidade Geradora)", op_ug, key=f'ug_select_{rc}')
            st.write("") 
            ativo = st.selectbox("Ativo", dados['Ativo'], key=f'ativo_{rc}')
            
            items_proc = ugs
            if ativo in ['INVERSOR', 'TRACKER', 'STRING'] and ugs:
                 df_det = dados['df_detalhado']
                 df_filt = df_det[df_det['Usina'].isin(ugs)]
                 cmap = {'INVERSOR': 'Inversor Conectado', 'TRACKER': 'Tracker Conectado', 'STRING': 'Nome String'}
                 target_col = cmap.get(ativo)
                 if target_col in df_filt.columns:
                     op_det = sorted(list(filter(None, df_filt[target_col].dropna().unique().tolist())))
                     items_proc = st.multiselect(f"Selecione {ativo}(s)", op_det, key=f'det_{rc}')
            
            st.session_state['items_proc'] = items_proc

        with col2:
            tipo = st.selectbox("Tipo de Ocorrência", dados['Tipo de ocorrência'], key=f'tipo_ocorrencia_{rc}')
            ocorr = st.selectbox("Ocorrência", dados['Ocorrência'], key=f'ocr_{rc}')
            oper = st.selectbox("Operador", dados['Operador'], key=f'opr_{rc}')

    # Aba 2: Detalhes
    with tab2:
        desc = st.text_area("Descrição Detalhada", height=150, placeholder="Descreva o problema aqui...", key=f'desc_{rc}')
        c_tec1, c_tec2 = st.columns(2)
        with c_tec1:
            os_in = st.text_input("Ordem de Serviço (OS)", key=f'os_{rc}')
            if cat == utils.SHEET_EQUIPAMENTOS:
                st.number_input("Quantidade", min_value=1, key=f'qtd_{rc}')
        with c_tec2:
            prot = st.text_input("Protocolo", key=f'prot_{rc}')

    # Aba 3: Datas (RESTAURADO: LÓGICA VARIADA)
    with tab3:
        st.markdown("**Registro do Desligamento**")
        
        # Chave para alternar entre "Tudo igual" ou "Individual"
        is_multi = len(items_proc) > 1
        usar_mesmo_tempo = True
        
        if is_multi:
            usar_mesmo_tempo = st.checkbox("Usar a mesma data e hora para todos os itens selecionados?", value=True, key=f'check_mesmo_tempo_{rc}')
        
        if usar_mesmo_tempo:
            c_h = st.columns(2)
            with c_h[0]:
                data_master = st.date_input("Data (Geral)", key=f'dd_master_{rc}')
            with c_h[1]:
                hora_master = st.time_input("Hora (Geral)", key=f'hd_master_{rc}')
        else:
            st.info("Preencha os horários específicos abaixo:")
            with st.expander("Expandir para preencher horários individuais", expanded=True):
                for i, item in enumerate(items_proc):
                    st.markdown(f"**{item}**")
                    c_ind = st.columns(2)
                    c_ind[0].date_input(f"Data - {item}", key=f'dd_{rc}_{i}')
                    c_ind[1].time_input(f"Hora - {item}", key=f'hd_{rc}_{i}')

    # --- BOTÃO SALVAR ---
    st.markdown("---")
    if st.button("💾 Adicionar Ocorrência", type="primary"):
        erros = []
        if cli == "-": erros.append("Cliente")
        if not ugs: erros.append("UGs")
        if tipo == "-": erros.append("Tipo de Ocorrência")
        if ativo == "-": erros.append("Ativo")
        if ocorr == "-": erros.append("Ocorrência")
        if oper == "-": erros.append("Operador")
        if not desc: erros.append("Descrição")
        if not items_proc: erros.append("Itens (UGs ou Ativos específicos)")
        
        if erros:
            st.error(f"⚠️ Preencha os campos obrigatórios: {', '.join(erros)}")
        else:
            overlay_on()
            try:
                client = utils.connect_to_google_sheets()
                ws = client.open_by_url(utils.SPREADSHEET_URL).worksheet(cat)
                
                rows_add = []
                header = ws.row_values(1)
                hmap = {h.strip().upper(): i for i, h in enumerate(header)}
                
                for i, item in enumerate(items_proc):
                    # Recupera UG pai
                    ug_final = item if item in ugs else ugs[0] 
                    
                    df_d = dados['df_dados']
                    sigla = ""
                    if not df_d.empty:
                        match = df_d[df_d['UG'] == ug_final]
                        if not match.empty: sigla = match['SIGLA'].iloc[0]
                    
                    # Recupera data/hora correta (Master ou Individual)
                    if usar_mesmo_tempo:
                        d_val = data_master
                        h_val = hora_master
                    else:
                        d_val = st.session_state.get(f'dd_{rc}_{i}')
                        h_val = st.session_state.get(f'hd_{rc}_{i}')
                    
                    if not d_val or not h_val:
                        # Fallback se algo falhar na captura individual
                        d_val = datetime.now().date()
                        h_val = datetime.now().time()

                    d_row = {
                        'UG': ug_final, 'CLIENTE': cli, 'SIGLA': sigla,
                        'TIPO DE OCORRÊNCIA': tipo, 'ATIVO': ativo, 'NOME ATIVO': item,
                        'OCORRÊNCIA': ocorr, 'OPERADOR': oper, 'DESCRIÇÃO': desc,
                        'PROTOCOLO': prot, 'OS': os_in, 
                        'DESLIGAMENTO': f"{d_val.strftime('%d/%m/%Y')} {h_val.strftime('%H:%M:%S')}"
                    }
                    if cat == utils.SHEET_EQUIPAMENTOS: 
                        d_row['QUANTIDADE'] = st.session_state.get(f'qtd_{rc}', 1)
                    
                    line = [''] * len(header)
                    for k, v in d_row.items():
                        if k in hmap: line[hmap[k]] = v
                    rows_add.append(line)
                
                if rows_add:
                    ws.append_rows(rows_add, value_input_option='USER_ENTERED')
                    st.session_state.last_submission_details = rows_add
                    st.session_state.form_reset_counter += 1
                    overlay_off()
                    st.rerun()
                else:
                    st.warning("Nenhum dado para salvar.")
                    overlay_off()

            except Exception as e:
                st.error(f"Erro ao salvar: {e}")
                overlay_off()