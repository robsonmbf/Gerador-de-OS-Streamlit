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
</style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE AUTENTICAÇÃO E LÓGICA DE NEGÓCIO ---
def show_login_page():
    st.markdown("""<div class="main-header"><h1>🔐 Acesso ao Sistema</h1><p>Faça login ou registre-se para acessar o Gerador de OS</p></div>""", unsafe_allow_html=True)
    tab1, tab2 = st.tabs(["Login", "Registro"])
    with tab1:
        with st.form("login_form"):
            email = st.text_input("Email", placeholder="seu@email.com")
            password = st.text_input("Senha", type="password")
            if st.form_submit_button("Entrar", use_container_width=True):
                if email and password:
                    success, message, session_data = auth_manager.login_user(email, password)
                    if success:
                        st.session_state.authenticated = True
                        st.session_state.user_data = session_data
                        st.session_state.user_data_loaded = False 
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)
                else:
                    st.error("Por favor, preencha todos os campos")
    with tab2:
        with st.form("register_form"):
            reg_email = st.text_input("Email", placeholder="seu@email.com", key="reg_email")
            reg_password = st.text_input("Senha", type="password", key="reg_password")
            reg_password_confirm = st.text_input("Confirmar Senha", type="password")
            if st.form_submit_button("Registrar", use_container_width=True):
                if reg_email and reg_password and reg_password_confirm:
                    if reg_password != reg_password_confirm:
                        st.error("As senhas não coincidem")
                    else:
                        success, message = auth_manager.register_user(reg_email, reg_password)
                        if success:
                            st.success(message)
                            st.info("Agora você pode fazer login com suas credenciais")
                        else:
                            st.error(message)
                else:
                    st.error("Por favor, preencha todos os campos")

def check_authentication():
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'user_data' not in st.session_state:
        st.session_state.user_data = None
    if st.session_state.authenticated and st.session_state.user_data:
        session_token = st.session_state.user_data.get('session_token')
        if session_token:
            is_valid, _ = auth_manager.validate_session(session_token)
            if not is_valid:
                st.session_state.authenticated = False
                st.session_state.user_data = None
                st.rerun()

def logout_user():
    if st.session_state.user_data and st.session_state.user_data.get('session_token'):
        auth_manager.logout_user(st.session_state.user_data['session_token'])
    st.session_state.authenticated = False
    st.session_state.user_data = None
    st.session_state.user_data_loaded = False
    st.rerun()

def show_user_info():
    if st.session_state.get('authenticated'):
        user_email = st.session_state.user_data.get('email', 'N/A')
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown(f'<div class="user-info">👤 <strong>Usuário:</strong> {user_email}</div>', unsafe_allow_html=True)
        with col2:
            if st.button("Sair", type="secondary"):
                logout_user()

def init_user_session_state():
    if st.session_state.get('authenticated') and not st.session_state.get('user_data_loaded'):
        user_id = st.session_state.user_data.get('user_id')
        if user_id:
            st.session_state.medicoes_adicionadas = user_data_manager.get_user_measurements(user_id)
            st.session_state.epis_adicionados = user_data_manager.get_user_epis(user_id)
            st.session_state.riscos_manuais_adicionados = user_data_manager.get_user_manual_risks(user_id)
            st.session_state.user_data_loaded = True
    
    if 'medicoes_adicionadas' not in st.session_state:
        st.session_state.medicoes_adicionadas = []
    if 'epis_adicionados' not in st.session_state:
        st.session_state.epis_adicionados = []
    if 'riscos_manuais_adicionados' not in st.session_state:
        st.session_state.riscos_manuais_adicionados = []
    if 'cargos_concluidos' not in st.session_state:
        st.session_state.cargos_concluidos = set()

def normalizar_texto(texto):
    if not isinstance(texto, str): return ""
    return re.sub(r'[\s\W_]+', '', texto.lower().strip())

def mapear_e_renomear_colunas_funcionarios(df):
    df_copia = df.copy()
    mapeamento = {
        'nome_do_funcionario': ['nomedofuncionario', 'nome', 'funcionario', 'funcionário', 'colaborador', 'nomecompleto'],
        'funcao': ['funcao', 'função', 'cargo'],
        'data_de_admissao': ['datadeadmissao', 'dataadmissao', 'admissao', 'admissão'],
        'setor': ['setordetrabalho', 'setor', 'area', 'área', 'departamento'],
        'descricao_de_atividades': ['descricaodeatividades', 'atividades', 'descricaoatividades', 'descriçãodeatividades', 'tarefas', 'descricaodastarefas'],
        'empresa': ['empresa'],
        'unidade': ['unidade']
    }
    colunas_renomeadas = {}
    colunas_df_normalizadas = {normalizar_texto(col): col for col in df_copia.columns}
    for nome_padrao, nomes_possiveis in mapeamento.items():
        for nome_possivel in nomes_possiveis:
            if nome_possivel in colunas_df_normalizadas:
                coluna_original = colunas_df_normalizadas[nome_possivel]
                colunas_renomeadas[coluna_original] = nome_padrao
                break
    df_copia.rename(columns=colunas_renomeadas, inplace=True)
    return df_copia

@st.cache_data
def carregar_planilha(arquivo):
    if arquivo is None: return None
    try:
        return pd.read_excel(arquivo)
    except Exception as e:
        st.error(f"Erro ao ler o ficheiro Excel: {e}")
        return None

@st.cache_data
def obter_dados_pgr():
    data = [
        {'categoria': 'fisico', 'risco': 'Ruído (Contínuo ou Intermitente)', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'},
        {'categoria': 'fisico', 'risco': 'Ruído (Impacto)', 'possiveis_danos': 'Perda auditiva, trauma acústico.'},
        {'categoria': 'fisico', 'risco': 'Vibração de Corpo Inteiro', 'possiveis_danos': 'Problemas na coluna, dores lombares.'},
        {'categoria': 'fisico', 'risco': 'Vibração de Mãos e Braços', 'possiveis_danos': 'Doenças osteomusculares, problemas circulatórios.'},
        {'categoria': 'fisico', 'risco': 'Calor', 'possiveis_danos': 'Desidratação, insolação, cãibras, exaustão, intermação.'},
        {'categoria': 'fisico', 'risco': 'Frio', 'possiveis_danos': 'Hipotermia, congelamento, doenças respiratórias.'},
        {'categoria': 'fisico', 'risco': 'Radiações Ionizantes', 'possiveis_danos': 'Câncer, mutações genéticas, queimaduras.'},
        {'categoria': 'fisico', 'risco': 'Radiações Não-Ionizantes', 'possiveis_danos': 'Queimaduras, lesões oculares, câncer de pele.'},
        {'categoria': 'fisico', 'risco': 'Pressões Anormais', 'possiveis_danos': 'Doença descompressiva, barotrauma.'},
        {'categoria': 'fisico', 'risco': 'Umidade', 'possiveis_danos': 'Doenças respiratórias, dermatites, micoses.'},
        {'categoria': 'quimico', 'risco': 'Poeiras', 'possiveis_danos': 'Pneumoconioses (silicose, asbestose), irritação respiratória.'},
        {'categoria': 'quimico', 'risco': 'Fumos', 'possiveis_danos': 'Doenças respiratórias (febre dos fumos metálicos), intoxicações.'},
        {'categoria': 'quimico', 'risco': 'Névoas', 'possiveis_danos': 'Irritação respiratória, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Gases', 'possiveis_danos': 'Asfixia, intoxicações, irritação respiratória.'},
        {'categoria': 'quimico', 'risco': 'Vapores', 'possiveis_danos': 'Irritação respiratória, intoxicações, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Produtos Químicos em Geral', 'possiveis_danos': 'Queimaduras, irritações, intoxicações, dermatites, câncer.'},
        {'categoria': 'biologico', 'risco': 'Bactérias', 'possiveis_danos': 'Infecções, doenças infecciosas (tétano, tuberculose).'},
        {'categoria': 'biologico', 'risco': 'Fungos', 'possiveis_danos': 'Micoses, alergias, infecções respiratórias.'},
        {'categoria': 'biologico', 'risco': 'Vírus', 'possiveis_danos': 'Doenças virais (hepatite, HIV), infecções.'},
        {'categoria': 'ergonomico', 'risco': 'Levantamento e Transporte Manual de Peso', 'possiveis_danos': 'Lesões musculoesqueléticas, dores na coluna.'},
        {'categoria': 'ergonomico', 'risco': 'Posturas Inadequadas', 'possiveis_danos': 'Dores musculares, lesões na coluna, LER/DORT.'},
        {'categoria': 'ergonomico', 'risco': 'Repetitividade', 'possiveis_danos': 'LER/DORT, tendinites, síndrome do túnel do carpo.'},
        {'categoria': 'acidente', 'risco': 'Máquinas e Equipamentos sem Proteção', 'possiveis_danos': 'Amputações, cortes, esmagamentos, prensamentos.'},
        {'categoria': 'acidente', 'risco': 'Eletricidade', 'possiveis_danos': 'Choque elétrico, queimaduras, fibrilação ventricular.'},
        {'categoria': 'acidente', 'risco': 'Trabalho em Altura', 'possiveis_danos': 'Quedas, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Projeção de Partículas', 'possiveis_danos': 'Lesões oculares, cortes na pele.'}
    ]
    return pd.DataFrame(data)

# --- FUNÇÃO CORRIGIDA DE SUBSTITUIÇÃO DE PLACEHOLDERS ---
def substituir_placeholders(doc, contexto):
    def substituir_em_paragrafo(paragrafo, contexto):
        texto = paragrafo.text
        for chave, valor in contexto.items():
            if chave in texto:
                texto = texto.replace(chave, str(valor))
        if paragrafo.runs:
            paragrafo.runs[0].text = texto
            for extra_run in paragrafo.runs[1:]:
                extra_run.text = ""
    
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for p in celula.paragraphs:
                    substituir_em_paragrafo(p, contexto)

    for p in doc.paragraphs:
        substituir_em_paragrafo(p, contexto)

# --- RESTANTE DO APP ---
# (mantém igual ao seu app original, só alterei a função acima)

# Funções de geração de OS e main continuam as mesmas...
