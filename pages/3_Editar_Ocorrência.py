import streamlit as st
import pandas as pd
from datetime import datetime
import utils
import html

# --- Configuração Inicial ---
st.set_page_config(layout="wide")
utils.init_overlay()

st.title("📝 Editar Ocorrência")

# --- CSS Personalizado (Mantendo consistência com as outras páginas) ---
st.markdown("""
<style>
    .stButton > button {
        background-color: #28a745; color: white; font-weight: bold;
        border-radius: 5px; padding: 10px 20px; width: 100%; border: none;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); transition: background-color 0.3s;
    }
    .stButton > button:hover { background-color: #218838; }
    
    /* CSS DETALHADO DOS CARDS */
    .card-container {
        background-color: #FF4B4B; color: white; padding: 15px;
        border-radius: 8px; margin-bottom: 15px; box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    .card-title {
        font-size: 1.5em; font-weight: bold; color: white;
        border-bottom: 1px solid rgba(255,255,255,0.5); padding-bottom: 5px; margin-bottom: 10px;
    }
    .card-item { margin-bottom: 5px; font-size: 1em; }
    .card-label { font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- Feedback Visual de Sucesso (Card Detalhado) ---
if 'edited_occurrence_feedback' in st.session_state:
    st.success("Ocorrência atualizada com sucesso!")
    d = st.session_state.pop('edited_occurrence_feedback')
    
    # Helper para formatar data/hora no card
    def _fmt(val):
        if not val: return ''
        if isinstance(val, str): return val
        # Se for objeto datetime/timestamp
        try: return val.strftime('%d/%m/%Y %H:%M')
        except: return str(val)

    st.markdown(f"""
    <div class="card-container">
        <div class="card-title">{html.escape(str(d.get('UG', '')))}</div>
        <div class="card-item"><span class="card-label">Cliente:</span> {html.escape(str(d.get('Cliente', '')))}</div>
        <div class="card-item"><span class="card-label">Ativo:</span> {html.escape(str(d.get('Ativo', '')))}</div>
        <div class="card-item"><span class="card-label">Nome Ativo:</span> {html.escape(str(d.get('Nome Ativo', '')))}</div>
        <div class="card-item"><span class="card-label">Tipo Ocorrência:</span> {html.escape(str(d.get('Tipo de ocorrência', '')))}</div>
        <div class="card-item"><span class="card-label">Ocorrência:</span> {html.escape(str(d.get('Ocorrência', '')))}</div>
        <div class="card-item"><span class="card-label">Operador:</span> {html.escape(str(d.get('Operador', '')))}</div>
        <br>
        <div class="card-item"><span class="card-label">Normalização:</span> {_fmt(d.get('Normalização'))}</div>
        <div class="card-item"><span class="card-label">Descrição:</span> {html.escape(str(d.get('Descrição', '')))}</div>
    </div>
    """, unsafe_allow_html=True)

# --- Seleção da Ocorrência ---
# Se não veio da Home com ID, mostra lista para selecionar
if 'id_unico_para_editar' not in st.session_state:
    df = utils.carregar_dados_completos()
    if not df.empty:
        # Cria coluna de exibição amigável
        df['Display'] = (
            df['UG'].astype(str) + " | " + 
            df['Ocorrência'].astype(str) + " | " + 
            df['Desligamento'].dt.strftime('%d/%m/%Y %H:%M').fillna('')
        )
        
        opts = df['Display'].tolist()
        sel = st.selectbox(
            "Buscar Ocorrência:", 
            options=opts, 
            index=None, 
            placeholder="Digite para buscar ou selecione..."
        )
        
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
                # Campos de Texto (Mantendo edição livre para flexibilidade)
                n_ug = st.text_input("UG", value=dados['UG'])
                n_nome_atv = st.text_input("Nome Ativo", value=dados['Nome Ativo'])
                n_tipo = st.text_input("Tipo", value=dados['Tipo de ocorrência'])
                n_ocorr = st.text_input("Ocorrência", value=dados['Ocorrência'])
                n_oper = st.text_input("Operador", value=dados['Operador'])
                n_desc = st.text_area("Descrição", value=dados['Descrição'], height=150)
                n_prot = st.text_input("Protocolo", value=dados['Protocolo'])
                n_os = st.text_input("OS", value=dados['OS'])
                
            with c2:
                # Datas
                def split_dt(val):
                    if pd.isna(val) or val == "": return None, None
                    try: return val.date(), val.time()
                    except AttributeError: return None, None

                dn_d, dn_t = split_dt(dados['Normalização'])
                st.markdown("**Normalização**")
                col_d1, col_d2 = st.columns(2)
                d_norm = col_d1.date_input("Data Norm.", value=dn_d)
                h_norm = col_d2.time_input("Hora Norm.", value=dn_t)
                
                dl_d, dl_t = split_dt(dados['Atendimento Loop'])
                st.markdown("**Atendimento Loop**")
                col_d3, col_d4 = st.columns(2)
                d_loop = col_d3.date_input("Data Loop", value=dl_d)
                h_loop = col_d4.time_input("Hora Loop", value=dl_t)
                
                dt_d, dt_t = split_dt(dados['Atendimento Terceiros'])
                st.markdown("**Atendimento Terceiros**")
                col_d5, col_d6 = st.columns(2)
                d_terc = col_d5.date_input("Data Terc.", value=dt_d)
                h_terc = col_d6.time_input("Hora Terc.", value=dt_t)

                da_d, da_t = split_dt(dados['Cliente Avisado'])
                st.markdown("**Cliente Avisado**")
                col_d7, col_d8 = st.columns(2)
                d_avis = col_d7.date_input("Data Aviso", value=da_d)
                h_avis = col_d8.time_input("Hora Aviso", value=da_t)

            st.markdown("---")
            submitted = st.form_submit_button("✅ Salvar Alterações", type="primary", use_container_width=True)
            
            if submitted:
                utils.overlay_on()
                try:
                    client = utils.connect_to_google_sheets()
                    ws = client.open_by_url(utils.SPREADSHEET_URL).worksheet(cat_atual)
                    
                    # --- Lógica para encontrar a linha no Google Sheets ---
                    all_values = ws.get_all_values()
                    header = all_values[0]
                    h_map = {h.strip().upper(): i for i, h in enumerate(header)}
                    
                    # Indices chaves para identificar a linha única
                    idx_ug = h_map.get('UG')
                    idx_atv = h_map.get('ATIVO')
                    idx_ocr = h_map.get('OCORRÊNCIA')
                    idx_des = h_map.get('DESLIGAMENTO')
                    
                    target_row_index = -1
                    
                    for i, r in enumerate(all_values):
                        if i == 0: continue # Pula header
                        try:
                            r_ug = r[idx_ug].upper()
                            r_atv = r[idx_atv].upper()
                            r_ocr = r[idx_ocr].upper()
                            r_des_raw = r[idx_des]
                            
                            r_dt = pd.to_datetime(r_des_raw, dayfirst=True, errors='coerce')
                            if pd.isna(r_dt): continue
                            
                            # Recria ID para comparação
                            row_id = f"{r_ug}|{r_atv}|{r_ocr}|{r_dt}"
                            
                            if row_id == id_alvo:
                                target_row_index = i + 1 # GSpread é 1-based
                                break
                        except Exception:
                            continue
                            
                    if target_row_index == -1:
                        st.error("Linha não encontrada na planilha. Os dados chaves podem ter mudado.")
                        utils.overlay_off()
                    else:
                        # Helper para formatar data str para salvar
                        def fmt_join(d, t):
                            if d and t: return datetime.combine(d, t).strftime('%d/%m/%Y %H:%M:%S')
                            return ""
                        
                        # Dados Normalizados para Salvar
                        norm_str = fmt_join(d_norm, h_norm)
                        
                        updates = {
                            'UG': n_ug, 'NOME ATIVO': n_nome_atv, 'TIPO DE OCORRÊNCIA': n_tipo,
                            'OCORRÊNCIA': n_ocorr, 'OPERADOR': n_oper, 'DESCRIÇÃO': n_desc,
                            'PROTOCOLO': n_prot, 'OS': n_os,
                            'NORMALIZAÇÃO': norm_str,
                            'ATENDIMENTO LOOP': fmt_join(d_loop, h_loop),
                            'ATENDIMENTO TERCEIROS': fmt_join(d_terc, h_terc),
                            'CLIENTE AVISADO': fmt_join(d_avis, h_avis)
                        }
                        
                        # Atualização Segura (Lê linha atual -> Modifica -> Escreve linha)
                        current_row_data = all_values[target_row_index - 1]
                        new_row_data = current_row_data[:]
                        
                        for col_name, val in updates.items():
                            if col_name in h_map:
                                new_row_data[h_map[col_name]] = val
                        
                        ws.update(f"A{target_row_index}", [new_row_data], value_input_option='USER_ENTERED')
                        
                        # Prepara dados para o Card de Feedback
                        # Precisamos enriquecer 'updates' com dados que não mudaram (como Cliente) para o card ficar bonito
                        feedback_data = dados.copy()
                        feedback_data.update(updates)
                        
                        st.session_state['edited_occurrence_feedback'] = feedback_data
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