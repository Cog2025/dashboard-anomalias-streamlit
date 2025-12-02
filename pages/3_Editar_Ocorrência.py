import os
import streamlit as st
import pandas as pd
from datetime import datetime, date, time
import time as pytime
import gspread
from google.oauth2.service_account import Credentials
import html
import utils

# --- Configuração ---
st.set_page_config(layout="wide", page_title="Editar Ocorrência")
utils.render_loading_overlay(st.session_state.get('ui_phase', 'ready'))

def overlay_on():
    st.session_state.ui_phase = 'loading'
    st.session_state.loading_ts = pytime.time()
    utils.render_loading_overlay('loading')

def overlay_off():
    st.session_state.ui_phase = 'ready'
    st.session_state.loading_ts = 0
    utils.render_loading_overlay('ready')

# CSS Cards e Tabs
st.markdown("""
<style>
    .stButton > button { background-color: #28a745; color: white; font-weight: bold; width: 100%; }
    .stButton > button:hover { background-color: #218838; }
    
    .card-container { background-color: #FF4B4B; color: white; padding: 15px; border-radius: 8px; margin-bottom: 15px; }
    .card-title { font-size: 1.5em; font-weight: bold; border-bottom: 1px solid rgba(255,255,255,0.5); padding-bottom: 5px; margin-bottom: 10px; }
    .card-item { margin-bottom: 5px; }
    .card-label { font-weight: bold; }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { height: 50px; font-weight: 600; border-radius: 4px 4px 0px 0px; }
</style>
""", unsafe_allow_html=True)

st.title("📝 Editar Ocorrência")

# Feedback Visual
if 'edited_occurrence_feedback' in st.session_state:
    st.success("✅ Ocorrência atualizada com sucesso!")
    d = st.session_state.pop('edited_occurrence_feedback')
    
    def _fmt(val):
        if not val: return ''
        if isinstance(val, str): return val
        return val.strftime('%d/%m/%Y %H:%M')

    st.markdown(f"""
    <div class="card-container">
        <div class="card-title">{html.escape(str(d.get('UG', '')))}</div>
        <div class="card-item"><span class="card-label">Ocorrência:</span> {html.escape(str(d.get('Ocorrência', '')))}</div>
        <div class="card-item"><span class="card-label">Normalização:</span> {_fmt(d.get('Normalização'))}</div>
        <div class="card-item"><span class="card-label">Descrição:</span> {html.escape(str(d.get('Descrição', '')))}</div>
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
        if not client: return pd.DataFrame()
        
        workbook = client.open_by_url(utils.SPREADSHEET_URL)
        df1 = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_DESLIGAMENTOS))
        df2 = utils.fetch_sheet_as_df(workbook.worksheet(utils.SHEET_EQUIPAMENTOS))
        df1['Categoria'] = 'DESLIGAMENTOS'; df2['Categoria'] = 'EQUIPAMENTOS'
        df = pd.concat([df1, df2], ignore_index=True)

        renomear = {}
        for col in df.columns:
            c_upper = col.strip().upper()
            if c_upper in MAPA_RENOMEAR: renomear[col] = MAPA_RENOMEAR[c_upper]
        df.rename(columns=renomear, inplace=True)
        df.fillna('', inplace=True)

        for c in ['Normalização', 'Desligamento', 'Atendimento Loop', 'Atendimento Terceiros', 'Cliente Avisado']:
            if c in df.columns: df[c] = pd.to_datetime(df[c], errors='coerce', dayfirst=False)

        # ID Único
        df['ID_Unico'] = (df['UG'].astype(str).str.upper() + "|" + 
                          df['Ativo'].astype(str).str.upper() + "|" + 
                          df['Ocorrência'].astype(str).str.upper() + "|" + 
                          df['Desligamento'].astype(str))
        return df
    except Exception as e:
        st.error(f"Erro: {e}")
        return pd.DataFrame()

def split_datetime(dt_obj):
    if pd.notna(dt_obj) and isinstance(dt_obj, (datetime, pd.Timestamp)):
        return dt_obj.date(), dt_obj.time()
    return None, None

def combine_date_time(date_val, time_val):
    if date_val and time_val: return datetime.combine(date_val, time_val)
    return None

# --- Seleção ---
if ('id_unico_para_editar' not in st.session_state) or (not st.session_state['id_unico_para_editar']):
    df_full = carregar_dados_completos()
    if not df_full.empty:
        # Tenta usar dataframe limpo se vier da home
        if 'df_lista_para_editar' in st.session_state:
             df_full = st.session_state['df_lista_para_editar']
        
        # Garante coluna display
        if 'Display' not in df_full.columns:
             # Fallback simples se não tiver display pronto
             df_full['Display'] = df_full['UG'].astype(str) + " | " + df_full['Ocorrência'].astype(str)

        st.subheader("Selecione a ocorrência")
        opts = df_full['Display'].tolist()
        sel = st.selectbox("Buscar:", options=opts, index=None, placeholder="Digite para buscar...")
        
        if sel:
            try:
                row = df_full[df_full['Display'] == sel].iloc[0]
                st.session_state['id_unico_para_editar'] = row['ID_Unico']
                st.rerun()
            except: pass
    else:
        st.warning("Sem dados.")
        st.stop()

# --- Formulário ---
id_alvo = st.session_state.get('id_unico_para_editar')
if id_alvo:
    df_full = carregar_dados_completos()
    row = df_full[df_full['ID_Unico'] == id_alvo]
    
    if row.empty:
        st.error("Ocorrência não encontrada (pode ter sido alterada).")
        if st.button("Voltar"):
            del st.session_state['id_unico_para_editar']
            st.rerun()
    else:
        dados = row.iloc[0].to_dict()
        st.info(f"Editando: **{dados['UG']}** - {dados['Ocorrência']}")
        
        with st.form("edit_form"):
            # --- USO DE ABAS (CHANGE 1) ---
            tab_dados, tab_datas, tab_tec = st.tabs(["📋 Dados Gerais", "📅 Datas & Horários", "⚙️ Técnico & OS"])
            
            with tab_dados:
                c1, c2 = st.columns(2)
                with c1:
                    n_ug = st.text_input("UG", value=dados['UG'])
                    n_nome_atv = st.text_input("Nome Ativo", value=dados['Nome Ativo'])
                    n_tipo = st.text_input("Tipo", value=dados['Tipo de ocorrência'])
                with c2:
                    n_ativo = st.text_input("Ativo", value=dados['Ativo'])
                    n_ocorr = st.text_input("Ocorrência", value=dados['Ocorrência'])
                    n_oper = st.text_input("Operador", value=dados['Operador'])

            with tab_datas:
                c_d1, c_d2 = st.columns(2)
                with c_d1:
                    st.markdown("##### Normalização (Finalização)")
                    dn_d, dn_t = split_datetime(dados['Normalização'])
                    d_norm = st.date_input("Data Normalização", value=dn_d)
                    h_norm = st.time_input("Hora Normalização", value=dn_t)
                    
                    st.markdown("##### Cliente Avisado")
                    da_d, da_t = split_datetime(dados['Cliente Avisado'])
                    d_avis = st.date_input("Data Aviso", value=da_d)
                    h_avis = st.time_input("Hora Aviso", value=da_t)

                with c_d2:
                    st.markdown("##### Atendimento Loop")
                    dl_d, dl_t = split_datetime(dados['Atendimento Loop'])
                    d_loop = st.date_input("Data Loop", value=dl_d)
                    h_loop = st.time_input("Hora Loop", value=dl_t)
                    
                    st.markdown("##### Atendimento Terceiros")
                    dt_d, dt_t = split_datetime(dados['Atendimento Terceiros'])
                    d_terc = st.date_input("Data Terceiros", value=dt_d)
                    h_terc = st.time_input("Hora Terceiros", value=dt_t)

            with tab_tec:
                n_desc = st.text_area("Descrição", value=dados['Descrição'], height=150)
                ct1, ct2 = st.columns(2)
                with ct1: n_prot = st.text_input("Protocolo", value=dados['Protocolo'])
                with ct2: n_os = st.text_input("OS", value=dados['OS'])

            st.markdown("---")
            if st.form_submit_button("💾 Salvar Alterações", type="primary"):
                overlay_on()
                try:
                    client = utils.connect_to_google_sheets()
                    ws = client.open_by_url(utils.SPREADSHEET_URL).worksheet(dados['Categoria'])
                    
                    # Lógica de Update
                    all_values = ws.get_all_values()
                    header = all_values[0]
                    h_map = {h.strip().upper(): i for i, h in enumerate(header)}
                    
                    # Indices chaves
                    idx_ug = h_map.get('UG'); idx_atv = h_map.get('ATIVO'); idx_ocr = h_map.get('OCORRÊNCIA'); idx_des = h_map.get('DESLIGAMENTO')
                    
                    target_row = -1
                    for i, r in enumerate(all_values):
                        if i == 0: continue
                        try:
                            # Recria ID para match
                            r_dt = pd.to_datetime(r[idx_des], dayfirst=True, errors='coerce')
                            if pd.isna(r_dt): continue
                            rid = f"{r[idx_ug].upper()}|{r[idx_atv].upper()}|{r[idx_ocr].upper()}|{r_dt}"
                            if rid == id_alvo:
                                target_row = i + 1; break
                        except: continue
                    
                    if target_row == -1:
                        st.error("Linha não encontrada.")
                        overlay_off()
                    else:
                        def jdt(d, t): return datetime.combine(d, t).strftime('%d/%m/%Y %H:%M:%S') if d and t else ""
                        
                        upd = {
                            'UG': n_ug, 'NOME ATIVO': n_nome_atv, 'TIPO DE OCORRÊNCIA': n_tipo,
                            'OCORRÊNCIA': n_ocorr, 'OPERADOR': n_oper, 'DESCRIÇÃO': n_desc,
                            'PROTOCOLO': n_prot, 'OS': n_os,
                            'NORMALIZAÇÃO': jdt(d_norm, h_norm), 'ATENDIMENTO LOOP': jdt(d_loop, h_loop),
                            'ATENDIMENTO TERCEIROS': jdt(d_terc, h_terc), 'CLIENTE AVISADO': jdt(d_avis, h_avis)
                        }
                        
                        # Update Row seguro
                        curr = all_values[target_row-1]
                        new_row = curr[:]
                        for k, v in upd.items():
                            if k in h_map: new_row[h_map[k]] = v
                        
                        ws.update(f"A{target_row}", [new_row], value_input_option='USER_ENTERED')
                        
                        # Feedback
                        fb = dados.copy(); fb.update(upd)
                        st.session_state['edited_occurrence_feedback'] = fb
                        del st.session_state['id_unico_para_editar']
                        st.cache_data.clear()
                        overlay_off()
                        st.rerun()

                except Exception as e:
                    st.error(f"Erro: {e}")
                    overlay_off()

    if st.button("Cancelar"):
        del st.session_state['id_unico_para_editar']
        st.rerun()