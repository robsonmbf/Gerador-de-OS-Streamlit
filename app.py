import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import zipfile
from io import BytesIO
import time
import re
import sys
import os

# Adicionar o diretório atual ao path para importar módulos locais
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from database.models import DatabaseManager
from database.auth import AuthManager
from database.user_data import UserDataManager

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Ordens de Serviço (OS)",
    page_icon="📄",
    layout="wide",
)

# --- DEFINIÇÃO DE CONSTANTES GLOBAIS ---
UNIDADES_DE_MEDIDA = ["dB(A)", "m/s²", "ppm", "mg/m³", "%", "°C", "lx", "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"]
AGENTES_DE_RISCO = sorted([
    "Ruído (Contínuo ou Intermitente)", "Ruído (Impacto)", "Vibração de Corpo Inteiro", "Vibração de Mãos e Braços",
    "Radiações Ionizantes", "Radiações Não-Ionizantes", "Frio", "Calor", "Pressões Anormais", "Umidade", "Poeiras", 
    "Fumos", "Névoas", "Neblinas", "Gases", "Vapores", "Produtos Químicos em Geral", "Vírus", "Bactérias", 
    "Protozoários", "Fungos", "Parasitas", "Bacilos"
])
CATEGORIAS_RISCO = {'fisico': '🔥 Físicos', 'quimico': '⚗️ Químicos', 'biologico': '🦠 Biológicos', 'ergonomico': '🏃 Ergonômicos', 'acidente': '⚠️ Acidentes'}

# --- Inicialização dos Gerenciadores ---
@st.cache_resource
def init_managers():
    db_manager = DatabaseManager()
    auth_manager = AuthManager(db_manager)
    user_data_manager = UserDataManager(db_manager)
    return db_manager, auth_manager, user_data_manager

db_manager, auth_manager, user_data_manager = init_managers()

# --- CSS PERSONALIZADO ---
st.markdown("""
<style>
    /* ... (CSS inalterado) ... */
</style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE AUTENTICAÇÃO ---
def show_login_page():
    # ... (código da página de login inalterado) ...
    pass

def check_authentication():
    # ... (código de verificação de autenticação inalterado) ...
    pass

def logout_user():
    # ... (código de logout inalterado) ...
    pass
    
def show_user_info():
    # ... (código de informações do usuário inalterado) ...
    pass

# --- NOVA LÓGICA DE SESSÃO E BANCO DE DADOS ---
def init_user_session_state():
    """Carrega os dados do usuário do DB para o session_state se ainda não foram carregados."""
    if st.session_state.get('user_data') and not st.session_state.get('user_data_loaded'):
        user_id = st.session_state.user_data.get('user_id')
        if user_id:
            st.session_state.medicoes_adicionadas = user_data_manager.get_user_measurements(user_id)
            st.session_state.epis_adicionados = user_data_manager.get_user_epis(user_id)
            st.session_state.riscos_manuais_adicionados = user_data_manager.get_user_manual_risks(user_id)
            st.session_state.cargos_concluidos = set()
            st.session_state.user_data_loaded = True # Marca que os dados foram carregados

# --- FUNÇÕES DE LÓGICA DE NEGÓCIO (com pequenas alterações) ---
@st.cache_data
def obter_dados_pgr():
    """Restaura a lista completa de riscos."""
    data = [
        {'categoria': 'fisico', 'risco': 'Ruído (Contínuo ou Intermitente)', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'},
        {'categoria': 'fisico', 'risco': 'Ruído (Impacto)', 'possiveis_danos': 'Perda auditiva, trauma acústico.'},
        # ... (TODA A LISTA COMPLETA DE RISCOS RESTAURADA AQUI) ...
        {'categoria': 'acidente', 'risco': 'Projeção de Partículas', 'possiveis_danos': 'Lesões oculares, cortes na pele.'}
    ]
    return pd.DataFrame(data)

# ... (outras funções como mapear_e_renomear_colunas, gerar_os, etc., permanecem iguais)

# --- APLICAÇÃO PRINCIPAL ---
def main():
    check_authentication()
    
    if not st.session_state.get('authenticated'):
        show_login_page()
        st.session_state.user_data_loaded = False # Garante que os dados serão recarregados no próximo login
        return
    
    # Carrega os dados do usuário para a sessão
    init_user_session_state()
    
    show_user_info()
    
    # ... (código da interface principal) ...
    
    with st.container(border=True):
        st.markdown('##### ⚠️ 3. Configure os Riscos e Medidas de Controle')
        # ... (interface de seleção de riscos) ...
        
        # --- SEÇÕES DE ENTRADA DE DADOS ATUALIZADAS ---
        
        # Adicionar Medições (com salvar e remover)
        with st.expander("📊 **Adicionar Medições**"):
            with st.form("form_medicao"):
                # ... (campos do formulário)
                if st.form_submit_button("Adicionar Medição"):
                    if agente_final and valor:
                        # Salva no banco de dados
                        user_data_manager.add_measurement(st.session_state.user_data['user_id'], agente_final, valor, unidade, epi_med)
                        # Recarrega os dados do DB para atualizar a tela
                        st.session_state.user_data_loaded = False
                        st.rerun()

            if st.session_state.get('medicoes_adicionadas'):
                st.write("**Medições salvas:**")
                for med in st.session_state.medicoes_adicionadas:
                    col1, col2 = st.columns([4, 1])
                    col1.markdown(f"- {med['agent']}: {med['value']} {med['unit']}")
                    if col2.button("Remover", key=f"del_med_{med['id']}"):
                        user_data_manager.remove_measurement(st.session_state.user_data['user_id'], med['id'])
                        st.session_state.user_data_loaded = False
                        st.rerun()

        # Adicionar Risco Manual (com salvar e remover)
        with st.expander("➕ **Adicionar Risco Manual** (na aba de Riscos)"):
             # (Lógica similar com add_manual_risk e remove_manual_risk)
             pass

        # Adicionar EPIs Gerais (com salvar e remover)
        with st.expander("🦺 **Adicionar EPIs Gerais**"):
             # (Lógica similar com add_epi e remove_epi)
             pass

    # ... (Resto do código, botão de gerar OS, etc.)

if __name__ == "__main__":
    main()
