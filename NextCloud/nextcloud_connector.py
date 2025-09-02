# nextcloud_connector.py
import streamlit as st
import pandas as pd
from webdav3.client import Client
import io

# --- Função para conectar ao cliente WebDAV ---
# O @st.cache_resource garante que a conexão seja feita apenas uma vez.
@st.cache_resource
def get_nextcloud_client():
    """Conecta ao servidor Nextcloud usando as credenciais dos segredos."""
    try:
        options = {
            'webdav_hostname': st.secrets["nextcloud"]["url"],
            'webdav_login':    st.secrets["nextcloud"]["login"],
            'webdav_password': st.secrets["nextcloud"]["password"]
        }
        return Client(options)
    except Exception as e:
        st.error(f"Erro ao conectar ao Nextcloud. Verifique suas credenciais em st.secrets: {e}")
        return None

# --- Função para ler o arquivo Excel do Nextcloud ---
# O @st.cache_data garante que o arquivo não seja baixado repetidamente sem necessidade.
@st.cache_data(ttl=60)
def read_excel_from_nextcloud():
    """Baixa o arquivo Excel da nuvem e o carrega em um dicionário de DataFrames (um por aba)."""
    try:
        client = get_nextcloud_client()
        if client:
            remote_path = st.secrets["nextcloud"]["path"]
            # Baixa o arquivo para a memória do aplicativo
            response = client.resource(remote_path).read()
            file_in_memory = io.BytesIO(response)
            # Lê todas as abas de uma vez
            all_sheets = pd.read_excel(file_in_memory, sheet_name=None, engine='openpyxl')
            return all_sheets
        return None
    except Exception as e:
        st.error(f"Erro ao ler o arquivo do Nextcloud. Verifique o caminho do arquivo e as permissões: {e}")
        return None

# --- Função para salvar o arquivo Excel de volta no Nextcloud ---
def write_excel_to_nextcloud(all_sheets_dict):
    """Recebe um dicionário de DataFrames e salva como um arquivo .xlsx, substituindo o antigo."""
    try:
        client = get_nextcloud_client()
        if client:
            remote_path = st.secrets["nextcloud"]["path"]
            buffer = io.BytesIO()
            
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                for sheet_name, df in all_sheets_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            buffer.seek(0) # Retorna ao início do buffer para leitura
            
            # Envia o arquivo em memória para o Nextcloud
            client.resource(remote_path).write(buffer.read())
            
            # Limpa o cache para forçar a releitura dos dados na próxima interação
            st.cache_data.clear()
            return True
        return False
    except Exception as e:
        st.error(f"Erro ao salvar o arquivo no Nextcloud: {e}")
        return False