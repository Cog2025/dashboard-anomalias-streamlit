import streamlit as st
import pandas as pd
from datetime import datetime
import time as pytime
import utils

st.set_page_config(layout="wide")
utils.init_overlay()
st.title("📝 Editar Ocorrência")

# --- Feedback Visual de Sucesso ---
if 'edited_occurrence_feedback' in st.session_state:
    st.success("Ocorrência atualizada com sucesso!")
    d = st.session_state.pop('edited_occurrence_feedback')
    
    st.markdown(f"""
    <div style="background-color:#FF4B4B; padding:15px; border-radius:8px; color:white; margin-bottom:15px;">
        <h3>{d.get('UG', '')}</h3>
        <p><b>Ocorrência:</b> {d.get('Ocorrência', '')}</p>
        <p><b>Status:</b> Dados salvos e normalização atualizada.</p>
    </div>
    """, unsafe_allow_html=True)

# --- Seleção da Ocorrência ---
# Se não veio da Home com ID, mostra lista para selecionar
if 'id_unico_para_editar' not in st.session_state:
    df = utils.carregar_dados_completos()
    if not df.empty:
        # Filtra apenas abertas ou mostra todas? Mostra todas para permitir correção
        df['Display'] = (
            df['UG'] + " | " + df['Ocorrência'] + " | " + 
            df['Desligamento'].dt.strftime('%d/%m/%Y %H:%M').fillna('')
        )
        opts = df['Display'].tolist()
        sel = st.selectbox("Buscar Ocorrência:", options=opts, index=None, placeholder="Digite para buscar...")
        
        if sel:
             id_unico = df.loc[df['Display'] == sel, 'ID_Unico'].values[0]
             st.session_state['id_unico_para_editar'] = id_unico
             st.rerun()
    else:
        st.error("Sem dados para carregar.")
        st.stop()

# --- Formulário de Edição ---
if 'id_unico_para_editar' in st.session_state:
    id_alvo = st.session_state['id_unico_para_editar']
    df_full = utils.carregar_dados_completos()
    
    # Encontra a linha no DF local
    row = df_full[df_full['ID_Unico'] == id_alvo]
    
    if row.empty:
        st.error("ID não encontrado. A ocorrência pode ter sido deletada ou a data alterada.")
        if st.button("Voltar"):
            del st.session_state['id_unico_para_editar']
            st.rerun()
    else:
        dados = row.iloc[0].to_dict()
        cat_atual = dados['Categoria']
        
        st.subheader(f"Editando: {dados['UG']} ({cat_atual})")
        
        with st.form("edit_form"):
            c1, c2 = st.columns(2)
            with c1:
                # Campos de Texto
                n_ug = st.text_input("UG", value=dados['UG'])
                n_nome_atv = st.text_input("Nome Ativo", value=dados['Nome Ativo'])
                # Seria ideal carregar listas de opções aqui, mas text_input permite edição livre
                n_tipo = st.text_input("Tipo", value=dados['Tipo de ocorrência'])
                n_ocorr = st.text_input("Ocorrência", value=dados['Ocorrência'])
                n_oper = st.text_input("Operador", value=dados['Operador'])
                n_desc = st.text_area("Descrição", value=dados['Descrição'])
                n_prot = st.text_input("Protocolo", value=dados['Protocolo'])
                n_os = st.text_input("OS", value=dados['OS'])
                
            with c2:
                # Datas
                def split_dt(val):
                    if pd.isna(val) or val == "": return None, None
                    return val.date(), val.time()

                dn_d, dn_t = split_dt(dados['Normalização'])
                d_norm = st.date_input("Data Normalização", value=dn_d)
                h_norm = st.time_input("Hora Normalização", value=dn_t)
                
                dl_d, dl_t = split_dt(dados['Atendimento Loop'])
                d_loop = st.date_input("Data Loop", value=dl_d)
                h_loop = st.time_input("Hora Loop", value=dl_t)
                
                dt_d, dt_t = split_dt(dados['Atendimento Terceiros'])
                d_terc = st.date_input("Data Terceiros", value=dt_d)
                h_terc = st.time_input("Hora Terceiros", value=dt_t)

                da_d, da_t = split_dt(dados['Cliente Avisado'])
                d_avis = st.date_input("Data Aviso", value=da_d)
                h_avis = st.time_input("Hora Aviso", value=da_t)

            submitted = st.form_submit_button("✅ Salvar Alterações", type="primary")
            
            if submitted:
                utils.overlay_on()
                try:
                    client = utils.connect_to_google_sheets()
                    ws = client.open_by_url(utils.SPREADSHEET_URL).worksheet(cat_atual)
                    
                    # --- Lógica para encontrar a linha no Google Sheets ---
                    # Precisamos encontrar a linha que bate com UG, ATIVO, OCORRÊNCIA e DESLIGAMENTO
                    all_values = ws.get_all_values()
                    header = all_values[0]
                    h_map = {h.strip().upper(): i for i, h in enumerate(header)}
                    
                    # Indices chaves
                    idx_ug = h_map.get('UG')
                    idx_atv = h_map.get('ATIVO')
                    idx_ocr = h_map.get('OCORRÊNCIA')
                    idx_des = h_map.get('DESLIGAMENTO')
                    
                    target_row_index = -1
                    
                    # Recria o ID original para comparação
                    original_deslig_str = str(dados['Desligamento']) # Formato timestamp pandas
                    
                    for i, r in enumerate(all_values):
                        if i == 0: continue # Pula header
                        
                        # Tenta parsear a data da planilha para comparar com o ID
                        try:
                            # A planilha está em texto d/m/y h:m:s, o ID_Unico usa o formato do pandas
                            # Estratégia: Criar o ID da linha atual e comparar com id_alvo
                            r_ug = r[idx_ug].upper()
                            r_atv = r[idx_atv].upper()
                            r_ocr = r[idx_ocr].upper()
                            r_des_raw = r[idx_des]
                            
                            r_dt = pd.to_datetime(r_des_raw, dayfirst=True, errors='coerce')
                            if pd.isna(r_dt): continue
                            
                            # ID Recriado
                            row_id = f"{r_ug}|{r_atv}|{r_ocr}|{r_dt}"
                            
                            if row_id == id_alvo:
                                target_row_index = i + 1 # GSpread é 1-based
                                break
                        except Exception:
                            continue
                            
                    if target_row_index == -1:
                        st.error("Linha não encontrada na planilha. Os dados chaves (UG, Data) podem ter mudado.")
                        utils.overlay_off()
                    else:
                        # Atualiza a linha
                        # Helper para formatar data str
                        def fmt_join(d, t):
                            if d and t: return datetime.combine(d, t).strftime('%d/%m/%Y %H:%M:%S')
                            return ""
                        
                        # Monta dict de update
                        updates = {
                            'UG': n_ug, 'NOME ATIVO': n_nome_atv, 'TIPO DE OCORRÊNCIA': n_tipo,
                            'OCORRÊNCIA': n_ocorr, 'OPERADOR': n_oper, 'DESCRIÇÃO': n_desc,
                            'PROTOCOLO': n_prot, 'OS': n_os,
                            'NORMALIZAÇÃO': fmt_join(d_norm, h_norm),
                            'ATENDIMENTO LOOP': fmt_join(d_loop, h_loop),
                            'ATENDIMENTO TERCEIROS': fmt_join(d_terc, h_terc),
                            'CLIENTE AVISADO': fmt_join(d_avis, h_avis)
                        }
                        
                        # Precisamos ler a linha inteira, atualizar e escrever de volta
                        # Ou atualizar celula por celula. Row update é mais seguro.
                        current_row_data = all_values[target_row_index - 1] # 0-based para lista local
                        new_row_data = current_row_data[:]
                        
                        for col_name, val in updates.items():
                            if col_name in h_map:
                                new_row_data[h_map[col_name]] = val
                        
                        # Update range
                        cell_range = f"A{target_row_index}"
                        ws.update(cell_range, [new_row_data], value_input_option='USER_ENTERED')
                        
                        # Sucesso
                        st.session_state['edited_occurrence_feedback'] = updates
                        del st.session_state['id_unico_para_editar']
                        st.cache_data.clear()
                        utils.overlay_off()
                        st.rerun()

                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")
                    utils.overlay_off()

    if st.button("Cancelar Edição"):
        del st.session_state['id_unico_para_editar']
        st.rerun()