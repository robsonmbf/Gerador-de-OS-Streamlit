import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
UNIDADES_DE_MEDIDA = ["dB(A)", "m/s²", "m/s¹⁷⁵", "ppm", "mg/m³", "%", "°C", "lx", "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"]
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
        {'categoria': 'quimico', 'risco': 'Exposição a Produto Químico', 'possiveis_danos': 'Irritação/lesão ocular, na pele e mucosas; Dermatites; Queimadura Química; Intoxicação; Náuseas; Vômitos.'},
        {'categoria': 'fisico', 'risco': 'Ambiente Artificialmente Frio', 'possiveis_danos': 'Estresse, desconforto, dormência, rigidez nas partes com maior intensidade de exposição ao frio, redução da destreza, formigamento, redução da sensibilidade dos dedos e flexibilidade das articulações.'},
        {'categoria': 'fisico', 'risco': 'Exposição ao Ruído', 'possiveis_danos': 'Perda Auditiva Induzida pelo Ruído Ocupacional (PAIRO).'},
        {'categoria': 'fisico', 'risco': 'Vibrações Localizadas (mão/braço)', 'possiveis_danos': 'Alterações articulares e vasomotoras.'},
        {'categoria': 'fisico', 'risco': 'Vibração de Corpo Inteiro (AREN/VDVR)', 'possiveis_danos': 'Alterações no sistema digestivo, sistema musculoesquelético, sistema nervoso, alterações na visão, enjoos, náuseas, palidez.'},
        {'categoria': 'fisico', 'risco': 'Radiações Ionizantes', 'possiveis_danos': 'Câncer, mutações genéticas, queimaduras.'},
        {'categoria': 'fisico', 'risco': 'Radiações Não Ionizantes', 'possiveis_danos': 'Queimaduras, lesões oculares, câncer de pele.'},
        {'categoria': 'fisico', 'risco': 'Pressões Anormais', 'possiveis_danos': 'Doença descompressiva, barotrauma.'},
        {'categoria': 'fisico', 'risco': 'Umidade', 'possiveis_danos': 'Doenças respiratórias, dermatites, micoses.'},
        {'categoria': 'fisico', 'risco': 'Calor', 'possiveis_danos': 'Desidratação, insolação, cãibras, exaustão, intermação.'},
        {'categoria': 'fisico', 'risco': 'Frio', 'possiveis_danos': 'Hipotermia, congelamento, doenças respiratórias.'},
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
        {'categoria': 'acidente', 'risco': 'Projeção de Partículas', 'possiveis_danos': 'Lesões oculares, cortes na pele.'},
        {'categoria': 'acidente', 'risco': 'Animais peçonhentos', 'possiveis_danos': 'Picadas, mordidas, reações alérgicas, infecções, dor, inchaço, necrose, paralisia, morte.'},
        {'categoria': 'acidente', 'risco': 'Armazenamento inadequado de materiais', 'possiveis_danos': 'Quedas, soterramento, esmagamento, lesões por esforço repetitivo.'},
        {'categoria': 'acidente', 'risco': 'Atropelamento', 'possiveis_danos': 'Fraturas, lacerações, traumatismos, morte.'},
        {'categoria': 'acidente', 'risco': 'Choque contra objetos', 'possiveis_danos': 'Contusões, fraturas, lacerações.'},
        {'categoria': 'acidente', 'risco': 'Colisão', 'possiveis_danos': 'Contusões, fraturas, lacerações, traumatismos.'},
        {'categoria': 'acidente', 'risco': 'Contato com eletricidade', 'possiveis_danos': 'Choque elétrico, queimaduras, fibrilação ventricular, morte.'},
        {'categoria': 'acidente', 'risco': 'Contato com superfície quente', 'possiveis_danos': 'Queimaduras de 1º, 2º ou 3º grau.'},
        {'categoria': 'acidente', 'risco': 'Contato com superfície fria', 'possiveis_danos': 'Queimaduras por frio, hipotermia.'},
        {'categoria': 'acidente', 'risco': 'Corte/Laceração', 'possiveis_danos': 'Hemorragia, infecção, perda de função.'},
        {'categoria': 'acidente', 'risco': 'Empilhamento inadequado', 'possiveis_danos': 'Quedas, soterramento, esmagamento.'},
        {'categoria': 'acidente', 'risco': 'Equipamento com defeito/sem manutenção', 'possiveis_danos': 'Falha do equipamento, acidentes, lesões.'},
        {'categoria': 'acidente', 'risco': 'Explosão', 'possiveis_danos': 'Queimaduras, traumatismos, projeção de fragmentos, morte.'},
        {'categoria': 'acidente', 'risco': 'Incêndio', 'possiveis_danos': 'Queimaduras, inalação de fumaça, asfixia, morte.'},
        {'categoria': 'acidente', 'risco': 'Impacto de objetos', 'possiveis_danos': 'Contusões, fraturas, lacerações.'},
        {'categoria': 'acidente', 'risco': 'Máquinas e equipamentos sem proteção', 'possiveis_danos': 'Amputações, cortes, esmagamentos, prensamentos.'},
        {'categoria': 'acidente', 'risco': 'Manuseio de produtos químicos sem EPI', 'possiveis_danos': 'Queimaduras químicas, irritações, intoxicações.'},
        {'categoria': 'acidente', 'risco': 'Queda de altura', 'possiveis_danos': 'Fraturas, traumatismos, morte.'},
        {'categoria': 'acidente', 'risco': 'Queda em mesmo nível', 'possiveis_danos': 'Contusões, entorses, fraturas.'},
        {'categoria': 'acidente', 'risco': 'Soterramento', 'possiveis_danos': 'Asfixia, traumatismos, morte.'},
        {'categoria': 'acidente', 'risco': 'Veículos em movimento', 'possiveis_danos': 'Atropelamento, colisão, esmagamento, morte.'},
        {'categoria': 'acidente', 'risco': 'Agressão física', 'possiveis_danos': 'Lesões corporais, traumatismos, estresse psicológico.'},
        {'categoria': 'acidente', 'risco': 'Animais (ataque de)', 'possiveis_danos': 'Mordidas, arranhões, infecções, reações alérgicas.'},
        {'categoria': 'acidente', 'risco': 'Desabamento/colapso', 'possiveis_danos': 'Soterramento, esmagamento, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Exposição a temperaturas extremas', 'possiveis_danos': 'Hipotermia, hipertermia, queimaduras, insolação.'},
        {'categoria': 'acidente', 'risco': 'Ferramentas manuais (uso inadequado)', 'possiveis_danos': 'Cortes, perfurações, contusões, fraturas.'},
        {'categoria': 'acidente', 'risco': 'Incêndio/explosão', 'possiveis_danos': 'Queimaduras, inalação de fumaça, traumatismos, morte.'},
        {'categoria': 'acidente', 'risco': 'Objetos cortantes/perfurocortantes', 'possiveis_danos': 'Cortes, perfurações, lacerações, infecções.'},
        {'categoria': 'acidente', 'risco': 'Produtos químicos (derramamento/vazamento)', 'possiveis_danos': 'Queimaduras químicas, irritações, intoxicações, problemas respiratórios.'},
        {'categoria': 'acidente', 'risco': 'Queda de objetos', 'possiveis_danos': 'Impacto, contusões, fraturas, traumatismos.'},
        {'categoria': 'acidente', 'risco': 'Ruído excessivo', 'possiveis_danos': 'Perda auditiva, zumbido, estresse.'},
        {'categoria': 'acidente', 'risco': 'Superfícies escorregadias/irregulares', 'possiveis_danos': 'Quedas, contusões, fraturas.'},
        {'categoria': 'acidente', 'risco': 'Trabalho em altura', 'possiveis_danos': 'Quedas, fraturas, traumatismos, morte.'},
        {'categoria': 'acidente', 'risco': 'Trabalho em espaço confinado', 'possiveis_danos': 'Asfixia, intoxicação, desmaio, morte.'},
        {'categoria': 'acidente', 'risco': 'Veículos e máquinas (operação)', 'possiveis_danos': 'Atropelamento, colisão, esmagamento, amputações, morte.'},
        {'categoria': 'biologico', 'risco': 'Vírus, bactérias, fungos, parasitas', 'possiveis_danos': 'Infecções, doenças, reações alérgicas.'},
        {'categoria': 'ergonomico', 'risco': 'Posturas inadequadas', 'possiveis_danos': 'Dores musculares, lesões na coluna, LER/DORT.'},
        {'categoria': 'ergonomico', 'risco': 'Repetitividade', 'possiveis_danos': 'LER/DORT, tendinites, síndrome do túnel do carpo.'},
        {'categoria': 'ergonomico', 'risco': 'Levantamento e transporte manual de peso', 'possiveis_danos': 'Lesões musculoesqueléticas, dores na coluna.'},
        {'categoria': 'fisico', 'risco': 'Ruído (contínuo ou intermitente)', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'},
        {'categoria': 'fisico', 'risco': 'Vibração (corpo inteiro)', 'possiveis_danos': 'Problemas na coluna, dores lombares.'},
        {'categoria': 'fisico', 'risco': 'Vibração (mãos e braços)', 'possiveis_danos': 'Doenças osteomusculares, problemas circulatórios.'},
        {'categoria': 'fisico', 'risco': 'Calor', 'possiveis_danos': 'Desidratação, insolação, cãibras, exaustão, intermação.'},
        {'categoria': 'fisico', 'risco': 'Frio', 'possiveis_danos': 'Hipotermia, congelamento, doenças respiratórias.'},
        {'categoria': 'fisico', 'risco': 'Radiações ionizantes', 'possiveis_danos': 'Câncer, mutações genéticas, queimaduras.'},
        {'categoria': 'fisico', 'risco': 'Radiações não ionizantes', 'possiveis_danos': 'Queimaduras, lesões oculares, câncer de pele.'},
        {'categoria': 'fisico', 'risco': 'Pressões anormais', 'possiveis_danos': 'Doença descompressiva, barotrauma.'},
        {'categoria': 'fisico', 'risco': 'Umidade', 'possiveis_danos': 'Doenças respiratórias, dermatites, micoses.'},
        {'categoria': 'quimico', 'risco': 'Poeiras', 'possiveis_danos': 'Pneumoconioses (silicose, asbestose), irritação respiratória.'},
        {'categoria': 'quimico', 'risco': 'Fumos', 'possiveis_danos': 'Doenças respiratórias (febre dos fumos metálicos), intoxicações.'},
        {'categoria': 'quimico', 'risco': 'Névoas', 'possiveis_danos': 'Irritação respiratória, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Gases', 'possiveis_danos': 'Asfixia, intoxicações, irritação respiratória.'},
        {'categoria': 'quimico', 'risco': 'Vapores', 'possiveis_danos': 'Irritação respiratória, intoxicações, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Produtos químicos em geral', 'possiveis_danos': 'Queimaduras, irritações, intoxicações, dermatites, câncer.'},
        {'categoria': 'acidente', 'risco': 'Animais peçonhentos e insetos', 'possiveis_danos': 'Ferimento, corte, contusão, reação alérgica, infecção, morte.'},
        {'categoria': 'acidente', 'risco': 'Atropelamento', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Batida contra', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas.'},
        {'categoria': 'acidente', 'risco': 'Choque elétrico', 'possiveis_danos': 'Queimadura de 1º, 2º ou 3º grau, fibrilação ventricular, morte.'},
        {'categoria': 'acidente', 'risco': 'Contato com o sistema elétrico energizado.', 'possiveis_danos': 'Queimadura de 1º, 2º ou 3º grau, fibrilação ventricular, morte.'},
        {'categoria': 'acidente', 'risco': 'Contato com ferramentas cortantes e/ou perfurantes.', 'possiveis_danos': 'Corte, laceração, ferida contusa, punctura (ferida aberta), perfuração.'},
        {'categoria': 'acidente', 'risco': 'Contato com partes móveis de máquinas e equipamentos.', 'possiveis_danos': 'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.'},
        {'categoria': 'acidente', 'risco': 'Contato com produtos químicos.', 'possiveis_danos': 'Intoxicação, envenenamento, queimadura, irritação ou reação alérgica.'},
        {'categoria': 'acidente', 'risco': 'Contato com substância cáustica, tóxica ou nociva.', 'possiveis_danos': 'Intoxicação, envenenamento, queimadura, irritação ou reação alérgica.'},
        {'categoria': 'acidente', 'risco': 'Ingestão de substância cáustica, tóxica ou nociva.', 'possiveis_danos': 'Intoxicação, envenenamento, queimadura, irritação ou reação alérgica.'},
        {'categoria': 'acidente', 'risco': 'Inalação, ingestão e/ou absorção.', 'possiveis_danos': 'Intoxicação, envenenamento, queimadura, irritação ou reação alérgica.'},
        {'categoria': 'acidente', 'risco': 'Incêndio/Explosão', 'possiveis_danos': 'Queimadura de 1º, 2º ou 3º grau, asfixia,  arremessos, cortes, escoriações, luxações, fraturas.'},
        {'categoria': 'acidente', 'risco': 'Objetos cortantes/perfurocortantes', 'possiveis_danos': 'Corte, laceração, ferida contusa, punctura (ferida aberta), perfuração.'},
        {'categoria': 'acidente', 'risco': 'Pessoas não autorizadas e/ou visitantes no local de trabalho', 'possiveis_danos': 'Escoriação, ferimento, corte, luxação, fratura, entre outros danos devido às características do local e atividades realizadas.'},
        {'categoria': 'acidente', 'risco': 'Portas, escotilhas, tampas, "bocas de visita", flanges', 'possiveis_danos': 'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações, exposição à gases tóxicos.'},
        {'categoria': 'acidente', 'risco': 'Projeção de Partículas sólidas e/ou líquidas', 'possiveis_danos': 'Ferimento, corte, queimadura, perfuração, intoxicação.'},
        {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível de andaime, passarela, plataforma, etc.', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível de escada (móvel ou fixa).', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível de material empilhado.', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível de veículo.', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível em poço, escavação, abertura no piso, etc.', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível ≤ 2m', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível > 2m', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Queda de pessoa em mesmo nível', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas.'},
        {'categoria': 'acidente', 'risco': 'Reação do corpo a seus movimentos (escorregão sem queda, etc.)', 'possiveis_danos': 'Torções, distensções, rupturas ou outras lesões musculares internas.'},
        {'categoria': 'acidente', 'risco': 'Vidro (recipientes, portas, bancadas, janelas, objetos diversos).', 'possiveis_danos': 'Corte, ferimento, perfuração.'},
        {'categoria': 'acidente', 'risco': 'Soterramento', 'possiveis_danos': 'Asfixia, desconforto respiratório, nível de consciência alterado, letargia, palidez, pele azulada, tosse, transtorno neurológico.'},
        {'categoria': 'acidente', 'risco': 'Substâncias tóxicas e/ou inflamáveis', 'possiveis_danos': 'Intoxicação, asfixia, queimaduras de  1º, 2º ou 3º grau.'},
        {'categoria': 'acidente', 'risco': 'Superfícies, substâncias e/ou objetos aquecidos ', 'possiveis_danos': 'Queimadura de 1º, 2º ou 3º grau.'},
        {'categoria': 'acidente', 'risco': 'Superfícies, substâncias e/ou objetos em baixa temperatura ', 'possiveis_danos': 'Queimadura de 1º, 2º ou 3º grau.'},
        {'categoria': 'acidente', 'risco': 'Tombamento, quebra e/ou ruptura de estrutura (fixa ou móvel)', 'possiveis_danos': 'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.'},
        {'categoria': 'acidente', 'risco': 'Tombamento de máquina/equipamento', 'possiveis_danos': 'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.'},
        {'categoria': 'acidente', 'risco': 'Trabalho à céu aberto', 'possiveis_danos': 'Intermação, insolação, cãibra, exaustão, desidratação, resfriados.'},
        {'categoria': 'acidente', 'risco': 'Trabalho em espaços confinados', 'possiveis_danos': 'Asfixia, hiperóxia, contaminação por poeiras e/ou gases tóxicos, queimadura de 1º, 2º ou 3º grau, arremessos, cortes, escoriações, luxações, fraturas.'},
        {'categoria': 'acidente', 'risco': 'Trabalho com máquinas portáteis rotativas.', 'possiveis_danos': 'Cortes, ferimentos, escoriações, amputações.'},
        {'categoria': 'acidente', 'risco': 'Trabalho com máquinas e/ou equipamentos', 'possiveis_danos': 'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações, choque elétrico.'}
    ]
    return pd.DataFrame(data)

def substituir_placeholders(doc, contexto):
    """
    Substitui placeholders preservando a formatação do template.
    """
    def aplicar_formatacao_padrao(run):
        """Aplica formatação Segoe UI 9pt"""
        run.font.name = 'Segoe UI'
        run.font.size = Pt(9)
        return run

    def processar_paragrafo(p):
        texto_original_paragrafo = p.text

        # --- Lógica CORRIGIDA E MANTIDA para [MEDIÇÕES] ---
        if "[MEDIÇÕES]" in texto_original_paragrafo:
            for run in p.runs:
                run.text = ''
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            medicoes_valor = contexto.get("[MEDIÇÕES]", "Não aplicável")
            if medicoes_valor == "Não aplicável" or not medicoes_valor.strip():
                run = aplicar_formatacao_padrao(p.add_run("Não aplicável"))
                run.font.bold = False
            else:
                linhas = medicoes_valor.split('\n')
                for i, linha in enumerate(linhas):
                    if not linha.strip(): continue
                    if i > 0: p.add_run().add_break()
                    if ":" in linha:
                        partes = linha.split(":", 1)
                        agente_texto = partes[0].strip() + ":"
                        valor_texto = partes[1].strip()
                        run_agente = aplicar_formatacao_padrao(p.add_run(agente_texto + " "))
                        run_agente.font.bold = True
                        run_valor = aplicar_formatacao_padrao(p.add_run(valor_texto))
                        run_valor.font.bold = False
                    else:
                        run_simples = aplicar_formatacao_padrao(p.add_run(linha))
                        run_simples.font.bold = False
            return

        # --- Lógica RESTAURADA E CORRIGIDA para outros placeholders ---
        placeholders_no_paragrafo = [key for key in contexto if key in texto_original_paragrafo]
        if not placeholders_no_paragrafo:
            return

        # Preserva o estilo do primeiro 'run', que geralmente define o estilo do rótulo no template
        estilo_rotulo = {
            'bold': p.runs[0].bold if p.runs else False,
            'italic': p.runs[0].italic if p.runs else False,
            'underline': p.runs[0].underline if p.runs else False,
        }

        # Substitui todos os placeholders para obter o texto final
        texto_final = texto_original_paragrafo
        for key in placeholders_no_paragrafo:
            texto_final = texto_final.replace(key, str(contexto[key]))
        
        # Limpa o parágrafo para reescrevê-lo com a formatação correta
        p.clear()

        # Reconstrói o parágrafo, aplicando o estilo do rótulo e deixando os valores sem formatação
        texto_restante = texto_final
        for i, key in enumerate(placeholders_no_paragrafo):
            valor_placeholder = str(contexto[key])
            partes = texto_restante.split(valor_placeholder, 1)
            
            # Adiciona o texto antes do valor (que é o rótulo) com o estilo preservado
            if partes[0]:
                run_rotulo = aplicar_formatacao_padrao(p.add_run(partes[0]))
                run_rotulo.font.bold = estilo_rotulo['bold']
                run_rotulo.font.italic = estilo_rotulo['italic']
                run_rotulo.underline = estilo_rotulo['underline']

            # Adiciona o valor do placeholder sem formatação
            run_valor = aplicar_formatacao_padrao(p.add_run(valor_placeholder))
            run_valor.font.bold = False
            run_valor.font.italic = False
            run_valor.font.underline = False
            
            texto_restante = partes[1]

        # Adiciona qualquer texto que sobrar no final, com o estilo do rótulo
        if texto_restante:
            run_final = aplicar_formatacao_padrao(p.add_run(texto_restante))
            run_final.font.bold = estilo_rotulo['bold']
            run_final.font.italic = estilo_rotulo['italic']
            run_final.underline = estilo_rotulo['underline']

    # Processar parágrafos em tabelas e no corpo do documento
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    processar_paragrafo(p)
    for p in doc.paragraphs:
        processar_paragrafo(p)


def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, riscos_manuais, modelo_doc_carregado):
    doc = Document(modelo_doc_carregado)
    riscos_info = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]
    riscos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}
    danos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}

    # Processar riscos selecionados
    for _, risco_row in riscos_info.iterrows():
        categoria = str(risco_row.get("categoria", "")).lower()
        if categoria in riscos_por_categoria:
            riscos_por_categoria[categoria].append(str(risco_row.get("risco", "")))
            danos = risco_row.get("possiveis_danos")
            if pd.notna(danos): 
                danos_por_categoria[categoria].append(str(danos))

    # Processar riscos manuais
    if riscos_manuais:
        map_categorias_rev = {v: k for k, v in CATEGORIAS_RISCO.items()}
        for risco_manual in riscos_manuais:
            categoria_display = risco_manual.get('category')
            categoria_alvo = map_categorias_rev.get(categoria_display)
            if categoria_alvo:
                riscos_por_categoria[categoria_alvo].append(risco_manual.get('risk_name', ''))
                if risco_manual.get('possible_damages'):
                    danos_por_categoria[categoria_alvo].append(risco_manual.get('possible_damages'))

    # Limpar duplicatas
    for cat in danos_por_categoria:
        danos_por_categoria[cat] = sorted(list(set(danos_por_categoria[cat])))

    # FORMATAÇÃO SIMPLES DAS MEDIÇÕES
    medicoes_formatadas = []
    for med in medicoes_manuais:
        agente = str(med.get('agent', '')).strip()
        valor = str(med.get('value', '')).strip()
        unidade = str(med.get('unit', '')).strip()
       
        if agente and agente not in ['', 'N/A', 'nan', 'None'] and valor and valor not in ['', 'N/A', 'nan', 'None']:
            linha = f"{agente}: {valor}"
            if unidade and unidade not in ['', 'N/A', 'nan', 'None']:
                linha += f" {unidade}"
            if epi and epi not in ['', 'N/A', 'nan', 'None']:
                linha += f" | EPI: {epi}"
            medicoes_formatadas.append(linha)
    medicoes_texto = "\n".join(medicoes_formatadas) if medicoes_formatadas else "Não aplicável"

    # Processar data de admissão
    data_admissao = "Não informado"
    if 'data_de_admissao' in funcionario and pd.notna(funcionario['data_de_admissao']):
        try: 
            data_admissao = pd.to_datetime(funcionario['data_de_admissao']).strftime('%d/%m/%Y')
        except Exception: 
            data_admissao = str(funcionario['data_de_admissao'])
    elif 'Data de Admissão' in funcionario and pd.notna(funcionario['Data de Admissão']):
        try: 
            data_admissao = pd.to_datetime(funcionario['Data de Admissão']).strftime('%d/%m/%Y')
        except Exception: 
            data_admissao = str(funcionario['Data de Admissão'])

    # Processar descrição de atividades
    descricao_atividades = "Não informado"
    if 'descricao_de_atividades' in funcionario and pd.notna(funcionario['descricao_de_atividades']):
        descricao_atividades = str(funcionario['descricao_de_atividades']).strip()
    elif 'Descrição de Atividades' in funcionario and pd.notna(funcionario['Descrição de Atividades']):
        descricao_atividades = str(funcionario['Descrição de Atividades']).strip()

    if descricao_atividades == "Não informado" or descricao_atividades == "" or descricao_atividades == "nan":
        funcao = str(funcionario.get('funcao', funcionario.get('Função', 'N/A')))
        setor = str(funcionario.get('setor', funcionario.get('Setor', 'N/A')))
        if funcao != 'N/A' and setor != 'N/A':
            descricao_atividades = f"Atividades relacionadas à função de {funcao} no setor {setor}, incluindo todas as tarefas operacionais, administrativas e de apoio inerentes ao cargo."
        else:
            descricao_atividades = "Atividades operacionais, administrativas e de apoio conforme definido pela chefia imediata."

    def tratar_lista_vazia(lista, separador=", "):
        if not lista or all(not item.strip() for item in lista): 
            return "Não identificado"
        return separador.join(sorted(list(set(item for item in lista if item and item.strip()))))

    # Contexto
    contexto = {
        "[NOME EMPRESA]": str(funcionario.get("empresa", funcionario.get("Empresa", "N/A"))), 
        "[UNIDADE]": str(funcionario.get("unidade", funcionario.get("Unidade", "N/A"))),
        "[NOME FUNCIONÁRIO]": str(funcionario.get("nome_do_funcionario", funcionario.get("Nome", "N/A"))), 
        "[DATA DE ADMISSÃO]": data_admissao,
        "[SETOR]": str(funcionario.get("setor", funcionario.get("Setor", "N/A"))), 
        "[FUNÇÃO]": str(funcionario.get("funcao", funcionario.get("Função", "N/A"))),
        "[DESCRIÇÃO DE ATIVIDADES]": descricao_atividades,
        "[RISCOS FÍSICOS]": tratar_lista_vazia(riscos_por_categoria["fisico"]),
        "[RISCOS DE ACIDENTE]": tratar_lista_vazia(riscos_por_categoria["acidente"]),
        "[RISCOS QUÍMICOS]": tratar_lista_vazia(riscos_por_categoria["quimico"]),
        "[RISCOS BIOLÓGICOS]": tratar_lista_vazia(riscos_por_categoria["biologico"]),
        "[RISCOS ERGONÔMICOS]": tratar_lista_vazia(riscos_por_categoria["ergonomico"]),
        "[POSSÍVEIS DANOS RISCOS FÍSICOS]": tratar_lista_vazia(danos_por_categoria["fisico"], "; "),
        "[POSSÍVEIS DANOS RISCOS ACIDENTE]": tratar_lista_vazia(danos_por_categoria["acidente"], "; "),
        "[POSSÍVEIS DANOS RISCOS QUÍMICOS]": tratar_lista_vazia(danos_por_categoria["quimico"], "; "),
        "[POSSÍVEIS DANOS RISCOS BIOLÓGICOS]": tratar_lista_vazia(danos_por_categoria["biologico"], "; "),
        "[POSSÍVEIS DANOS RISCOS ERGONÔMICOS]": tratar_lista_vazia(danos_por_categoria["ergonomico"], "; "),
        "[EPIS]": tratar_lista_vazia([epi['epi_name'] for epi in epis_manuais]),
        "[MEDIÇÕES]": medicoes_texto,
    }

    substituir_placeholders(doc, contexto)
    return doc

# --- APLICAÇÃO PRINCIPAL ---
def main():
    check_authentication()
    init_user_session_state()
    
    if not st.session_state.get('authenticated'):
        show_login_page()
        return
    
    user_id = st.session_state.user_data['user_id']
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
        setores = sorted(df_funcionarios['setor'].dropna().unique().tolist()) if 'setor' in df_funcionarios.columns else []
        setor_sel = st.multiselect("Filtrar por Setor(es)", setores)
        df_filtrado_setor = df_funcionarios[df_funcionarios['setor'].isin(setor_sel)] if setor_sel else df_funcionarios
        st.caption(f"{len(df_filtrado_setor)} funcionário(s) no(s) setor(es) selecionado(s).")
        funcoes_disponiveis = sorted(df_filtrado_setor['funcao'].dropna().unique().tolist()) if 'funcao' in df_filtrado_setor.columns else []
        funcoes_formatadas = []
        if setor_sel:
            for funcao in funcoes_disponiveis:
                concluido = all((s, funcao) in st.session_state.cargos_concluidos for s in setor_sel)
                if concluido:
                    funcoes_formatadas.append(f"{funcao} ✅ Concluído")
                else:
                    funcoes_formatadas.append(funcao)
        else:
            funcoes_formatadas = funcoes_disponiveis
        funcao_sel_formatada = st.multiselect("Filtrar por Função/Cargo(s)", funcoes_formatadas)
        funcao_sel = [f.replace(" ✅ Concluído", "") for f in funcao_sel_formatada]
        df_final_filtrado = df_filtrado_setor[df_filtrado_setor['funcao'].isin(funcao_sel)] if funcao_sel else df_filtrado_setor
        st.success(f"**{len(df_final_filtrado)} funcionário(s) selecionado(s) para gerar OS.**")
        st.dataframe(df_final_filtrado[['nome_do_funcionario', 'setor', 'funcao']])

    with st.container(border=True):
        st.markdown('##### ⚠️ 3. Configure os Riscos e Medidas de Controle')
        st.info("Os riscos configurados aqui serão aplicados a TODOS os funcionários selecionados.")
        riscos_selecionados = []
        nomes_abas = list(CATEGORIAS_RISCO.values()) + ["➕ Manual"]
        tabs = st.tabs(nomes_abas)
        for i, (categoria_key, categoria_nome) in enumerate(CATEGORIAS_RISCO.items()):
            with tabs[i]:
                riscos_da_categoria = df_pgr[df_pgr['categoria'] == categoria_key]['risco'].tolist()
                selecionados = st.multiselect("Selecione os riscos:", options=riscos_da_categoria, key=f"riscos_{categoria_key}")
                riscos_selecionados.extend(selecionados)
        with tabs[-1]:
            with st.form("form_risco_manual", clear_on_submit=True):
                st.markdown("###### Adicionar um Risco que não está na lista")
                risco_manual_nome = st.text_input("Descrição do Risco")
                categoria_manual = st.selectbox("Categoria do Risco Manual", list(CATEGORIAS_RISCO.values()))
                danos_manuais = st.text_area("Possíveis Danos (Opcional)")
                if st.form_submit_button("Adicionar Risco Manual"):
                    if risco_manual_nome and categoria_manual:
                        user_data_manager.add_manual_risk(user_id, categoria_manual, risco_manual_nome, danos_manuais)
                        st.session_state.user_data_loaded = False
                        st.rerun()
            if st.session_state.riscos_manuais_adicionados:
                st.write("**Riscos manuais salvos:**")
                for r in st.session_state.riscos_manuais_adicionados:
                    col1, col2 = st.columns([4, 1])
                    col1.markdown(f"- **{r['risk_name']}** ({r['category']})")
                    if col2.button("Remover", key=f"rem_risco_{r['id']}"):
                        user_data_manager.remove_manual_risk(user_id, r['id'])
                        st.session_state.user_data_loaded = False
                        st.rerun()
        
        total_riscos = len(riscos_selecionados) + len(st.session_state.riscos_manuais_adicionados)
        if total_riscos > 0:
            with st.expander(f"📖 Resumo de Riscos Selecionados ({total_riscos} no total)", expanded=True):
                riscos_para_exibir = {cat: [] for cat in CATEGORIAS_RISCO.values()}
                for risco_nome in riscos_selecionados:
                    categoria_key_series = df_pgr[df_pgr['risco'] == risco_nome]['categoria']
                    if not categoria_key_series.empty:
                        categoria_key = categoria_key_series.iloc[0]
                        categoria_display = CATEGORIAS_RISCO.get(categoria_key)
                        if categoria_display:
                            riscos_para_exibir[categoria_display].append(risco_nome)
                for risco_manual in st.session_state.riscos_manuais_adicionados:
                    riscos_para_exibir[risco_manual['category']].append(risco_manual['risk_name'])
                for categoria, lista_riscos in riscos_para_exibir.items():
                    if lista_riscos:
                        st.markdown(f"**{categoria}**")
                        for risco in sorted(list(set(lista_riscos))):
                            st.markdown(f"- {risco}")
        
        st.divider()

        col_exp1, col_exp2 = st.columns(2)
        with col_exp1:
            with st.expander("📊 **Adicionar Medições**"):
                with st.form("form_medicao", clear_on_submit=True):
                    opcoes_agente = ["-- Digite um novo agente abaixo --"] + AGENTES_DE_RISCO
                    agente_selecionado = st.selectbox("Selecione um Agente/Fonte da lista...", options=opcoes_agente)
                    agente_manual = st.text_input("...ou digite um novo aqui:")
                    valor = st.text_input("Valor Medido")
                    unidade = st.selectbox("Unidade", UNIDADES_DE_MEDIDA)
                    epi_med = st.text_input("EPI Associado (Opcional)")
                    if st.form_submit_button("Adicionar Medição"):
                        agente_a_salvar = agente_manual.strip() if agente_manual.strip() else agente_selecionado
                        if agente_a_salvar != "-- Digite um novo agente abaixo --" and valor:
                            user_data_manager.add_measurement(user_id, agente_a_salvar, valor, unidade, epi_med)
                            st.session_state.user_data_loaded = False
                            st.rerun()
                        else:
                            st.warning("Por favor, preencha o Agente e o Valor.")
                if st.session_state.medicoes_adicionadas:
                    st.write("**Medições salvas:**")
                    for med in st.session_state.medicoes_adicionadas:
                        col1, col2 = st.columns([4, 1])
                        col1.markdown(f"- {med['agent']}: {med['value']} {med['unit']}")
                        if col2.button("Remover", key=f"rem_med_{med['id']}"):
                            user_data_manager.remove_measurement(user_id, med['id'])
                            st.session_state.user_data_loaded = False
                            st.rerun()
        with col_exp2:
            with st.expander("🦺 **Adicionar EPIs Gerais**"):
                with st.form("form_epi", clear_on_submit=True):
                    epi_nome = st.text_input("Nome do EPI")
                    if st.form_submit_button("Adicionar EPI"):
                        if epi_nome:
                            user_data_manager.add_epi(user_id, epi_nome)
                            st.session_state.user_data_loaded = False
                            st.rerun()
                if st.session_state.epis_adicionados:
                    st.write("**EPIs salvos:**")
                    for epi in st.session_state.epis_adicionados:
                        col1, col2 = st.columns([4, 1])
                        col1.markdown(f"- {epi['epi_name']}")
                        if col2.button("Remover", key=f"rem_epi_{epi['id']}"):
                            user_data_manager.remove_epi(user_id, epi['id'])
                            st.session_state.user_data_loaded = False
                            st.rerun()

    st.divider()
    if st.button("🚀 Gerar OS para Funcionários Selecionados", type="primary", use_container_width=True, disabled=df_final_filtrado.empty):
        with st.spinner(f"Gerando {len(df_final_filtrado)} documentos..."):
            documentos_gerados = []
            combinacoes_processadas = set()
            for _, func in df_final_filtrado.iterrows():
                combinacoes_processadas.add((func['setor'], func['funcao']))
                doc = gerar_os(
                    func, 
                    df_pgr, 
                    riscos_selecionados, 
                    st.session_state.epis_adicionados,
                    st.session_state.medicoes_adicionadas, 
                    st.session_state.riscos_manuais_adicionados, 
                    arquivo_modelo_os
                )
                doc_io = BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                nome_limpo = re.sub(r'[^\w\s-]', '', func.get("nome_do_funcionario", "Func_Sem_Nome")).strip().replace(" ", "_")
                caminho_no_zip = f"{func.get('setor', 'SemSetor')}/{func.get('funcao', 'SemFuncao')}/OS_{nome_limpo}.docx"
                documentos_gerados.append((caminho_no_zip, doc_io.getvalue()))
            st.session_state.cargos_concluidos.update(combinacoes_processadas)
            if documentos_gerados:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for nome_arquivo, conteudo_doc in documentos_gerados:
                        zip_file.writestr(nome_arquivo, conteudo_doc)
                nome_arquivo_zip = f"OS_Geradas_{time.strftime('%Y%m%d')}.zip"
                st.success(f"🎉 **{len(documentos_gerados)} Ordens de Serviço geradas!**")
                st.download_button(
                    label="📥 Baixar Todas as OS (.zip)", 
                    data=zip_buffer.getvalue(), 
                    file_name=nome_arquivo_zip, 
                    mime="application/zip",
                    use_container_width=True
                )

if __name__ == "__main__":
    main()
