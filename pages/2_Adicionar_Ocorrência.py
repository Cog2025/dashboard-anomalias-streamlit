import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import utils

st.set_page_config(layout="wide")
utils.init_overlay()

st.title("Adicionar Nova Ocorrência")

# CSS para botão verde
st.markdown("""
<style>
    .stButton > button { background-color: #28a745; color: white; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# Carrega opções (Cacheado)
@st.cache_data(ttl=60)
def carregar_opcoes_adicao():
    try:
        client = utils.connect_to_google_sheets()
        wb = client.open_by_url(utils.SPREADSHEET_URL)
        df_dados = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_DADOS)).fillna('')
        df_detalhado = utils.fetch_sheet_as_df(wb.worksheet(utils.SHEET_DETALHADA)).fillna('')
        
        return {
            'df_dados': df_dados, 
            'df_detalhado': df_detalhado,
            'clientes': sorted(utils.options_from(df_dados['CLIENTE'])),
            'ocorrencias': sorted(utils.options_from(df_dados['OCORRÊNCIA'])),
            'tipos': sorted(utils.options_from(df_dados['TIPO DE OCORRÊNCIA'])),
            'ativos': sorted(utils.options_from(df_dados['ATIVO'])),
            'operadores': sorted(utils.options_from(df_dados['OPERADOR']))
        }
    except Exception as e:
        st.error(f"Erro ao carregar opções: {e}")
        return {}

dados = carregar_opcoes_adicao()

if not dados:
    st.stop()

if 'form_reset_counter' not in st.session_state:
    st.session_state.form_reset_counter = 0

reset_counter = st.session_state.form_reset_counter

# --- Formulário ---
cat = st.selectbox("Categoria", [utils.SHEET_DESLIGAMENTOS, utils.SHEET_EQUIPAMENTOS], key=f'cat_{reset_counter}')

col1, col2 = st.columns(2)

with col1:
    cli = st.selectbox("Cliente", dados['clientes'], key=f'cliente_select_{reset_counter}')
    
    # Filtro Dinâmico de UGs
    op_ug = []
    if cli:
        df_d = dados['df_dados']
        op_ug = sorted(df_d[df_d['CLIENTE'] == cli]['UG'].unique().tolist())
    
    ugs = st.multiselect("UGs", op_ug, key=f'ug_select_{reset_counter}')
    tipo = st.selectbox("Tipo de Ocorrência", dados['tipos'], key=f'tipo_ocorrencia_{reset_counter}')
    ativo = st.selectbox("Ativo", dados['ativos'], key=f'ativo_{reset_counter}')
    
    # Lógica de Ativos Detalhados
    items_processar = ugs # Default
    if ativo in ['INVERSOR', 'TRACKER', 'STRING'] and ugs:
        df_det = dados['df_detalhado']
        df_filt = df_det[df_det['Usina'].isin(ugs)]
        col_map = {'INVERSOR': 'Inversor Conectado', 'TRACKER': 'Tracker Conectado', 'STRING': 'Nome String'}
        col_target = col_map.get(ativo)
        if col_target and col_target in df_filt.columns:
            op_detalhe = sorted(list(filter(None, df_filt[col_target].dropna().unique().tolist())))
            items_processar = st.multiselect(f"Selecione {ativo}(s)", op_detalhe, key=f'detalhe_{reset_counter}')
    
    st.session_state['items_para_processar'] = items_processar

with col2:
    ocorr = st.selectbox("Ocorrência", dados['ocorrencias'], key=f'ocorrencia_{reset_counter}')
    oper = st.selectbox("Operador", dados['operadores'], key=f'operador_{reset_counter}')
    desc = st.text_area("Descrição Detalhada", height=135, key=f'descricao_{reset_counter}')
    prot = st.text_input("Protocolo", key=f'protocolo_{reset_counter}')
    os_val = st.text_input("OS", key=f'os_input_{reset_counter}')
    
    qtd = 1
    if cat == utils.SHEET_EQUIPAMENTOS:
        qtd = st.number_input("Quantidade", min_value=1, key=f'quantidade_{reset_counter}')

st.markdown("---")
st.write("### Horários")
# Simplificado para Data/Hora Master para todos os itens
c_h = st.columns(4)
data_des = c_h[0].date_input("Data Desligamento", key=f'data_des_master_{reset_counter}')
hora_des = c_h[1].time_input("Hora Desligamento", key=f'hora_des_master_{reset_counter}')

# --- Botão Salvar ---
if st.button("Adicionar Ocorrência", type="primary", use_container_width=True):
    # --- VALIDAÇÃO DE CAMPOS OBRIGATÓRIOS ---
    erros = []
    if not cli: erros.append("Cliente")
    if not ugs: erros.append("UGs") # Verifica seleção base
    if not items_processar: erros.append("Itens (UGs ou Ativos)") # Verifica itens finais
    if not tipo: erros.append("Tipo de Ocorrência")
    if not ativo: erros.append("Ativo")
    if not ocorr: erros.append("Ocorrência")
    if not oper: erros.append("Operador")
    if not desc or len(desc) < 5: erros.append("Descrição (min 5 chars)")

    if erros:
        st.error(f"⚠️ Por favor, preencha: {', '.join(erros)}")
    else:
        # Prossegue
        utils.overlay_on()
        try:
            client = utils.connect_to_google_sheets()
            ws = client.open_by_url(utils.SPREADSHEET_URL).worksheet(cat)
            
            # Helper para encontrar a UG de um ativo (Inversor/Tracker)
            def find_ug(item_name):
                # Se o item já é uma UG, retorna ele mesmo
                if item_name in ugs: return item_name
                # Se não, busca no detalhado
                df_det = dados['df_detalhado']
                for c in ['Inversor Conectado', 'Tracker Conectado', 'Nome String']:
                     if c in df_det.columns:
                         match = df_det[df_det[c] == item_name]
                         if not match.empty: return match['Usina'].iloc[0]
                return None

            linhas_para_adicionar = []
            
            # Mapeamento do Header da Planilha
            header = ws.row_values(1)
            # Mapa: {NOME_COLUNA_UPPER: index}
            h_map = {h.strip().upper(): i for i, h in enumerate(header)}
            
            for item in items_processar:
                ug_final = find_ug(item)
                if not ug_final: continue
                
                # Busca Sigla
                df_d = dados['df_dados']
                row_cli = df_d[df_d['UG'] == ug_final]
                sigla = row_cli['SIGLA'].iloc[0] if not row_cli.empty else ""
                
                # Monta objeto de dados
                dados_row = {
                    'UG': ug_final,
                    'CLIENTE': cli,
                    'SIGLA': sigla,
                    'TIPO DE OCORRÊNCIA': tipo,
                    'ATIVO': ativo,
                    'NOME ATIVO': item,
                    'OCORRÊNCIA': ocorr,
                    'OPERADOR': oper,
                    'DESCRIÇÃO': desc,
                    'PROTOCOLO': prot,
                    'OS': os_val,
                    'QUANTIDADE': qtd,
                    'DESLIGAMENTO': f"{data_des.strftime('%d/%m/%Y')} {hora_des.strftime('%H:%M:%S')}"
                }
                
                # Cria array ordenado conforme colunas da planilha
                linha = [''] * len(header)
                for col_name, val in dados_row.items():
                    if col_name in h_map:
                        linha[h_map[col_name]] = val
                
                linhas_para_adicionar.append(linha)
            
            # Append no Google Sheets
            if linhas_para_adicionar:
                ws.append_rows(linhas_para_adicionar, value_input_option='USER_ENTERED')
                
                st.success(f"{len(linhas_para_adicionar)} ocorrências salvas!")
                st.session_state.form_reset_counter += 1
                utils.overlay_off()
                pytime.sleep(1)
                st.rerun()
            else:
                st.warning("Nenhum dado gerado para salvar.")
                utils.overlay_off()

        except Exception as e:
            st.error(f"Erro crítico ao salvar: {e}")
            utils.overlay_off()