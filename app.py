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
    [data-testid="stSidebar"] {
        display: none;
    }
    .main-header {
        text-align: center;
        padding-bottom: 20px;
    }
    .auth-container {
        max-width: 400px;
        margin: 0 auto;
        padding: 2rem;
        border: 1px solid #ddd;
        border-radius: 10px;
        background-color: #f9f9f9;
    }
    .user-info {
        background-color: #262730; 
        color: white;            
        padding: 1rem;
        border-radius: 5px;
        margin-bottom: 1rem;
        border: 1px solid #3DD56D; 
    }
    /* --- INÍCIO DO NOVO ESTILO PARA OS CARDS DE RISCO --- */
    .risk-card {
        border: 1px solid #e0e0e0;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 10px;
        background-color: #f9f9f9;
        transition: box-shadow 0.3s ease-in-out;
    }
    .risk-card:hover {
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .risk-description {
        font-size: 0.9rem;
        color: #666;
    }
    /* --- FIM DO NOVO ESTILO --- */
</style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE AUTENTICAÇÃO E LÓGICA DE NEGÓCIO ---
# (Todas as funções como show_login_page, obter_dados_pgr, gerar_os, etc., permanecem aqui sem alteração)
def show_login_page():
    # ... (código inalterado)
    pass
def check_authentication():
    # ... (código inalterado)
    pass
def logout_user():
    # ... (código inalterado)
    pass
def show_user_info():
    # ... (código inalterado)
    pass
def init_user_session_state():
    # ... (código inalterado)
    pass
def normalizar_texto(texto):
    # ... (código inalterado)
    pass
def mapear_e_renomear_colunas_funcionarios(df):
    # ... (código inalterado)
    pass
@st.cache_data
def carregar_planilha(arquivo):
    # ... (código inalterado)
    pass
@st.cache_data
def obter_dados_pgr():
    # ... (código com a lista completa)
    data = [
        {'categoria': 'fisico', 'risco': 'Ruído (Contínuo ou Intermitente)', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'},
        # ... (todos os outros riscos) ...
        {'categoria': 'acidente', 'risco': 'Projeção de Partículas', 'possiveis_danos': 'Lesões oculares, cortes na pele.'}
    ]
    return pd.DataFrame(data)
def substituir_placeholders(doc, contexto):
    # ... (código inalterado)
    pass
def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, riscos_manuais, modelo_doc_carregado):
    # ... (código inalterado)
    pass


# --- APLICAÇÃO PRINCIPAL ---
def main():
    check_authentication()
    init_user_session_state()
    
    if not st.session_state.get('authenticated'):
        show_login_page()
        return
    
    show_user_info()
    
    st.markdown("""<div class="main-header"><h1>📄 Gerador de Ordens de Serviço (OS)</h1><p>Gere OS em lote a partir de um modelo Word (.docx) e uma planilha de funcionários.</p></div>""", unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown("##### 📂 1. Carregue os Documentos")
        col1, col2 = st.columns(2)
        with col1:
            arquivo_funcionarios = st.file_uploader("📄 **Planilha de Funcionários (.xlsx)**", type="xlsx")
        with col2:
            arquivo_modelo_os = st.file_uploader("📝 **Modelo de OS (.docx)**", type="docx")

    if not arquivo_funcionarios or not arquivo_modelo_os:
        st.info("📋 Por favor, carregue a Planilha de Funcionários e o Modelo de OS para continuar.")
        return
    
    df_funcionarios_raw = carregar_planilha(arquivo_funcionarios)
    if df_funcionarios_raw is None:
        st.stop()

    df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_raw)
    df_pgr = obter_dados_pgr()

    with st.container(border=True):
        st.markdown('##### 👥 2. Selecione os Funcionários')
        # ... (código de filtro de funcionários inalterado)
        setores = []
        funcao_sel = []
        df_final_filtrado = df_funcionarios

    with st.container(border=True):
        st.markdown('##### ⚠️ 3. Configure os Riscos e Medidas de Controle')
        st.info("Os riscos, medições e EPIs configurados aqui serão aplicados a **TODOS** os funcionários selecionados acima.")

        # --- INÍCIO DA NOVA INTERFACE DE SELEÇÃO DE RISCOS ---
        riscos_selecionados = []
        
        # Inicializa o estado da sessão para 'selecionar_todos' se não existir
        if 'selecionar_todos' not in st.session_state:
            st.session_state.selecionar_todos = {}

        for categoria_key, categoria_nome in CATEGORIAS_RISCO.items():
            st.subheader(categoria_nome)
            riscos_da_categoria = df_pgr[df_pgr['categoria'] == categoria_key]
            
            # Botão "Selecionar Todos" para a categoria
            if st.button(f"Selecionar Todos - {categoria_nome.split(' ')[1]}", key=f"select_all_{categoria_key}"):
                # Inverte o estado de seleção para a categoria
                current_state = st.session_state.selecionar_todos.get(categoria_key, False)
                st.session_state.selecionar_todos[categoria_key] = not current_state

            select_all_value = st.session_state.selecionar_todos.get(categoria_key, False)

            # Divide os riscos em duas colunas
            col1, col2 = st.columns(2)
            metade = len(riscos_da_categoria) // 2
            
            for i, row in riscos_da_categoria.iterrows():
                col = col1 if i < (riscos_da_categoria.index[0] + metade) else col2
                with col:
                    # Usamos HTML/Markdown para criar o "cartão"
                    st.markdown('<div class="risk-card">', unsafe_allow_html=True)
                    
                    selecionado = st.checkbox(
                        label=f"**{row['risco']}**",
                        key=f"risk_{row['risco']}",
                        value=select_all_value, # Define o estado baseado no botão "Selecionar Todos"
                        help=row['possiveis_danos']
                    )
                    st.markdown(f"<div class='risk-description'>{row['possiveis_danos']}</div>", unsafe_allow_html=True)
                    
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    if selecionado:
                        riscos_selecionados.append(row['risco'])
            st.divider()
        # --- FIM DA NOVA INTERFACE ---

        # Seção para Adicionar Medições, Riscos Manuais e EPIs
        # ... (código dos expanders inalterado)
        col_exp1, col_exp2, col_exp3 = st.columns(3)
        with col_exp1:
            with st.expander("📊 **Adicionar Medições**"):
                # ...
                pass
        with col_exp2:
            with st.expander("➕ **Adicionar Risco Manual**"):
                # ...
                pass
        with col_exp3:
            with st.expander("🦺 **Adicionar EPIs Gerais**"):
                # ...
                pass

    st.divider()
    if st.button("🚀 Gerar OS para Funcionários Selecionados", type="primary", use_container_width=True, disabled=df_final_filtrado.empty):
        # ... (lógica de geração do ZIP inalterada)
        pass

if __name__ == "__main__":
    main()
