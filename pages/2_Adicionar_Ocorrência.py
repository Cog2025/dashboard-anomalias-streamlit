import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import utils

st.set_page_config(layout="wide")
utils.init_overlay()

st.title("Adicionar Nova Ocorrência")

# CSS botão
st.markdown("""
<style>
    .stButton > button { background-color: #28a745; color: white; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

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
            # Importante: Lists puras sem traços manuais, controlaremos no selectbox
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
# CORREÇÃO: index=None evita pré-seleção
cat = st.selectbox(
    "Selecione a Categoria da Ocorrência", 
    [utils.SHEET_DESLIGAMENTOS, utils.SHEET_EQUIPAMENTOS], 
    index=None, 
    placeholder="Selecione a categoria...",
    key=f'cat_{reset_counter}'
)

if not cat:
    st.info("Por favor, selecione uma categoria acima para continuar.")
    st.stop()

col1, col2 = st.columns(2)

with col1:
    # CORREÇÃO: index=None em todos os campos chaves
    cli = st.selectbox("Cliente", dados['clientes'], index=None, placeholder="Selecione...", key=f'cliente_select_{reset_counter}')
    
    op_ug = []
    if cli:
        df_d = dados['df_dados']
        op_ug = sorted(df_d[df_d['CLIENTE'] == cli]['UG'].unique().tolist())
    
    ugs = st.multiselect("UGs", op_ug, placeholder="Escolha as UGs...", key=f'ug_select_{reset_counter}')
    tipo = st.selectbox("Tipo de Ocorrência", dados['tipos'], index=None, placeholder="Selecione...", key=f'tipo_ocorrencia_{reset_counter}')
    ativo = st.selectbox("Ativo", dados['ativos'], index=None, placeholder="Selecione...", key=f'ativo_{reset_counter}')
    
    items_processar = ugs 
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
    ocorr = st.selectbox("Ocorrência", dados['ocorrencias'], index=None, placeholder="Selecione...", key=f'ocorrencia_{reset_counter}')
    oper = st.selectbox("Operador", dados['operadores'], index=None, placeholder="Selecione...", key=f'operador_{reset_counter}')
    desc = st.text_area("Descrição Detalhada", height=135, key=f'descricao_{reset_counter}')
    prot = st.text_input("Protocolo", key=f'protocolo_{reset_counter}')
    os_val = st.text_input("OS", key=f'os_input_{reset_counter}')
    
    qtd = 1
    if cat == utils.SHEET_EQUIPAMENTOS:
        qtd = st.number_input("Quantidade", min_value=1, key=f'quantidade_{reset_counter}')

st.markdown("---")
st.write("### Horários")
c_h = st.columns(4)
data_des = c_h[0].date_input("Data Desligamento", key=f'data_des_master_{reset_counter}')
hora_des = c_h[1].time_input("Hora Desligamento", key=f'hora_des_master_{reset_counter}')

if st.button("Adicionar Ocorrência", type="primary", use_container_width=True):
    erros = []
    if not cli: erros.append("Cliente")
    if not ugs: erros.append("UGs") 
    if not items_processar: erros.append("Itens (UGs ou Ativos)")
    if not tipo: erros.append("Tipo de Ocorrência")
    if not ativo: erros.append("Ativo")
    if not ocorr: erros.append("Ocorrência")
    if not oper: erros.append("Operador")
    if not desc or len(desc) < 5: erros.append("Descrição (min 5 chars)")

    if erros:
        st.error(f"⚠️ Por favor, preencha: {', '.join(erros)}")
    else:
        utils.overlay_on()
        try:
            client = utils.connect_to_google_sheets()
            ws = client.open_by_url(utils.SPREADSHEET_URL).worksheet(cat)
            
            def find_ug(item_name):
                if item_name in ugs: return item_name
                df_det = dados['df_detalhado']
                for c in ['Inversor Conectado', 'Tracker Conectado', 'Nome String']:
                     if c in df_det.columns:
                         match = df_det[df_det[c] == item_name]
                         if not match.empty: return match['Usina'].iloc[0]
                return None

            linhas_para_adicionar = []
            header = ws.row_values(1)
            h_map = {h.strip().upper(): i for i, h in enumerate(header)}
            
            for item in items_processar:
                ug_final = find_ug(item)
                if not ug_final: continue
                
                df_d = dados['df_dados']
                row_cli = df_d[df_d['UG'] == ug_final]
                sigla = row_cli['SIGLA'].iloc[0] if not row_cli.empty else ""
                
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
                
                linha = [''] * len(header)
                for col_name, val in dados_row.items():
                    if col_name in h_map:
                        linha[h_map[col_name]] = val
                
                linhas_para_adicionar.append(linha)
            
            if linhas_para_adicionar:
                ws.append_rows(linhas_para_adicionar, value_input_option='USER_ENTERED')
                
                # Feedback de sucesso detalhado (similar ao Editar)
                st.success(f"{len(linhas_para_adicionar)} ocorrência(s) adicionada(s)!")
                
                # Visualização em cards dos itens adicionados
                st.write("---")
                cols = st.columns(4)
                for i, row_data in enumerate(linhas_para_adicionar):
                    # Recupera dados para exibição basica
                    ug_idx = h_map.get('UG')
                    ocr_idx = h_map.get('OCORRÊNCIA')
                    
                    ug_val = row_data[ug_idx] if ug_idx < len(row_data) else "?"
                    ocr_val = row_data[ocr_idx] if ocr_idx < len(row_data) else "?"
                    
                    with cols[i % 4]:
                        st.markdown(f"""
                        <div style="background-color:#28a745; padding:10px; border-radius:5px; margin-bottom:10px;">
                            <b>{ug_val}</b><br>{ocr_val}
                        </div>
                        """, unsafe_allow_html=True)
                
                st.session_state.form_reset_counter += 1
                pytime.sleep(3)
                utils.overlay_off()
                st.rerun()
            else:
                st.warning("Nenhum dado gerado.")
                utils.overlay_off()

        except Exception as e:
            st.error(f"Erro crítico ao salvar: {e}")
            utils.overlay_off()