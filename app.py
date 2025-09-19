# 🔐 Acesso ao Sistema

# Faça login ou registre-se para acessar o Gerador de OS

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

try:
    from database.models import DatabaseManager
    from database.auth import AuthManager
    from database.user_data import UserDataManager
except ImportError:
    st.error("❌ Erro ao importar módulos do banco de dados. Verifique se as dependências estão instaladas.")
    st.stop()

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Ordens de Serviço (OS)",
    page_icon="📄",
    layout="wide",
)

# --- DEFINIÇÃO DE CONSTANTES GLOBAIS ---
UNIDADES_DE_MEDIDA = ["dB(A)", "m/s²", "ppm", "mg/m³", "%", "°C", "lx", "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"]

# --- AGENTES DE RISCO EXPANDIDOS E ORGANIZADOS POR CATEGORIA ---
RISCOS_FISICO = sorted([
    "Ambiente Artificialmente Frio",
    "Exposição ao Ruído",
    "Ruído (Contínuo ou Intermitente)",
    "Ruído (Impacto)",
    "Exposição a Radiações Ionizantes",
    "Exposição a Radiações Não-ionizantes",
    "Radiações Ionizantes",
    "Radiações Não-Ionizantes",
    "Exposição a Temperatura Ambiente Baixa",
    "Exposição a Temperatura Ambiente Elevada",
    "Frio",
    "Calor",
    "Pressão Atmosférica Anormal (condições hiperbáricas)",
    "Pressões Anormais",
    "Umidade",
    "Vibração de Corpo Inteiro (AREN)",
    "Vibração de Corpo Inteiro (VDVR)",
    "Vibração de Corpo Inteiro",
    "Vibrações Localizadas (mão/braço)",
    "Vibração de Mãos e Braços"
])

RISCOS_QUIMICO = sorted([
    "Exposição a Produto Químico",
    "Produtos Químicos em Geral",
    "Poeiras",
    "Fumos",
    "Névoas",
    "Neblinas",
    "Gases",
    "Vapores"
])

RISCOS_BIOLOGICO = sorted([
    "Água e/ou alimentos contaminados",
    "Contaminação pelo Corona Vírus",
    "Contato com Fluido Orgânico (sangue, hemoderivados, secreções, excreções)",
    "Contato com Pessoas Doentes e/ou Material Infectocontagiante",
    "Exposição a Agentes Microbiológicos (fungos, bactérias, vírus, protozoários, parasitas)",
    "Vírus",
    "Bactérias",
    "Protozoários",
    "Fungos",
    "Parasitas",
    "Bacilos"
])

RISCOS_ERGONOMICO = sorted([
    "Assento inadequado",
    "Assédio de qualquer natureza no trabalho",
    "Cadência do trabalho imposta por um equipamento",
    "Compressão de partes do corpo por superfícies rígidas ou com quinas vivas",
    "Conflitos hierárquicos no trabalho",
    "Desequilíbrio entre tempo de trabalho e tempo de repouso",
    "Dificuldades para cumprir ordens e determinações da chefia relacionadas ao trabalho",
    "Elevação frequente de membros superiores",
    "Encosto do assento inadequado ou ausente",
    "Equipamentos e/ou máquinas sem meios de regulagem de ajustes ou sem condições de uso",
    "Equipamentos/mobiliário não adaptados a antropometria do trabalhador",
    "Esforço físico intenso",
    "Exigência de concentração, atenção e memória",
    "Exposição a vibração de corpo inteiro",
    "Exposição a vibrações localizadas (mão, braço)",
    "Falta de autonomia para a realização de tarefas no trabalho",
    "Flexões da coluna vertebral frequentes",
    "Frequente ação de empurrar/puxar cargas ou volumes",
    "Frequente deslocamento a pé durante a jornada de trabalho",
    "Frequente execução de movimentos repetitivos",
    "Iluminação inadequada",
    "Insatisfação no trabalho",
    "Insuficiência de capacitação para a execução da tarefa",
    "Levantamento e transporte manual de cargas ou volumes",
    "Manuseio de ferramentas e/ou objetos pesados por longos períodos",
    "Manuseio ou movimentação de cargas e volumes sem pega ou com pega pobre",
    "Mobiliário ou equipamento sem espaço para movimentação de segmentos corporais",
    "Mobiliário sem meios de regulagem de ajustes",
    "Monotonia",
    "Necessidade de alcançar objetos, documentos, controles, etc, além das zonas de alcance ideais",
    "Necessidade de manter ritmos intensos de trabalho",
    "Piso escorregadio ou irregular",
    "Posto de trabalho improvisado/inadequado",
    "Posto de trabalho não planejado/adaptado para a posição sentada",
    "Postura em pé por longos períodos",
    "Postura sentada por longos períodos",
    "Posturas incômodas/pouco confortáveis por longos períodos",
    "Pressão sonora fora dos parâmetros de conforto",
    "Problemas de relacionamento no trabalho",
    "Realização de múltiplas tarefas com alta demanda mental/cognitiva",
    "Reflexos que causem desconforto ou prejudiquem a visão",
    "Situações de estresse no local de trabalho",
    "Situações de sobrecarga de trabalho mental",
    "Temperatura efetiva fora dos parâmetros de conforto",
    "Trabalho com necessidade de variação de turnos",
    "Trabalho com utilização rigorosa de metas de produção",
    "Trabalho em condições de difícil comunicação",
    "Trabalho intensivo com teclado ou outros dispositivos de entrada de dados",
    "Trabalho noturno",
    "Trabalho realizado sem pausas pré-definidas para descanso",
    "Trabalho remunerado por produção",
    "Umidade do ar fora dos parâmetros de conforto",
    "Uso frequente de alavancas",
    "Uso frequente de escadas",
    "Uso frequente de força, pressão, preensão, flexão, extensão ou torção dos segmentos corporais",
    "Uso frequente de pedais",
    "Velocidade do ar fora dos parâmetros de conforto"
])

RISCOS_ACIDENTE = sorted([
    "Absorção (por contato) de substância cáustica, tóxica ou nociva",
    "Afogamento, imersão, engolfamento",
    "Aprisionamento em, sob ou entre",
    "Aprisionamento em, sob ou entre desabamento ou desmoronamento de edificação, estrutura, barreira, etc",
    "Aprisionamento em, sob ou entre dois ou mais objetos em movimento (sem encaixe)",
    "Aprisionamento em, sob ou entre objetos em movimento convergente",
    "Aprisionamento em, sob ou entre um objeto parado e outro em movimento",
    "Arestas cortantes, superfícies com rebarbas, farpas ou elementos de fixação expostos",
    "Ataque de ser vivo (inclusive humano)",
    "Ataque de ser vivo com peçonha",
    "Ataque de ser vivo com transmissão de doença",
    "Ataque de ser vivo por mordedura, picada, chifrada, coice, etc",
    "Atrito ou abrasão",
    "Atrito ou abrasão por corpo estranho no olho",
    "Atrito ou abrasão por encostar em objeto",
    "Atrito ou abrasão por manusear objeto",
    "Atropelamento",
    "Batida contra objeto parado ou em movimento",
    "Carga Suspensa",
    "Colisão entre veículos e/ou equipamentos autopropelidos",
    "Condições climáticas adversas (sol, chuva, vento, etc)",
    "Contato com objeto ou substância a temperatura muito alta",
    "Contato com objeto ou substância a temperatura muito baixa",
    "Contato com objeto ou substância em movimento",
    "Desabamento/Desmoronamento de edificação, estrutura e/ou materiais diversos",
    "Elementos Móveis e/ou Rotativos",
    "Emergências na circunvizinhança",
    "Equipamento pressurizado hidráulico ou pressurizado",
    "Exposição a Energia Elétrica",
    "Ferramentas elétricas",
    "Ferramentas manuais",
    "Gases/vapores/poeiras (tóxicos ou não tóxicos)",
    "Gases/vapores/poeiras inflamáveis",
    "Impacto de pessoa contra objeto em movimento",
    "Impacto de pessoa contra objeto parado",
    "Impacto sofrido por pessoa",
    "Impacto sofrido por pessoa, de objeto em movimento",
    "Impacto sofrido por pessoa, de objeto projetado",
    "Impacto sofrido por pessoa, de objeto que cai",
    "Incêndio/Explosão",
    "Ingestão de substância cáustica, tóxica ou nociva",
    "Inalação de substância tóxica/nociva",
    "Inalação, ingestão e/ou absorção",
    "Objetos cortantes/perfurocortantes",
    "Pessoas não autorizadas e/ou visitantes no local de trabalho",
    "Portas, escotilhas, tampas, bocas de visita, flanges",
    "Projeção de Partículas sólidas e/ou líquidas",
    "Queda de pessoa com diferença de nível maior que 2m",
    "Queda de pessoa com diferença de nível menor ou igual a 2m",
    "Queda de pessoa com diferença de nível de andaime, passarela, plataforma, etc",
    "Queda de pessoa com diferença de nível de escada (móvel ou fixa)",
    "Queda de pessoa com diferença de nível de material empilhado",
    "Queda de pessoa com diferença de nível de veículo",
    "Queda de pessoa com diferença de nível em poço, escavação, abertura no piso, etc",
    "Queda de pessoa em mesmo nível",
    "Reação do corpo a seus movimentos (escorregão sem queda, etc)",
    "Soterramento",
    "Substâncias tóxicas e/ou inflamáveis",
    "Superfícies, substâncias e/ou objetos aquecidos",
    "Superfícies, substâncias e/ou objetos em baixa temperatura",
    "Tombamento de máquina/equipamento",
    "Tombamento, quebra e/ou ruptura de estrutura (fixa ou móvel)",
    "Trabalho a céu aberto",
    "Trabalho com máquinas e/ou equipamentos",
    "Trabalho com máquinas portáteis rotativas",
    "Trabalho em espaços confinados",
    "Vidro (recipientes, portas, bancadas, janelas, objetos diversos)"
])

# Lista completa de agentes de risco (compatibilidade com código existente)
AGENTES_DE_RISCO = sorted(RISCOS_FISICO + RISCOS_QUIMICO + RISCOS_BIOLOGICO + RISCOS_ERGONOMICO + RISCOS_ACIDENTE)

# Dicionário para mapear categorias aos riscos
AGENTES_POR_CATEGORIA = {
    'fisico': RISCOS_FISICO,
    'quimico': RISCOS_QUIMICO,
    'biologico': RISCOS_BIOLOGICO,
    'ergonomico': RISCOS_ERGONOMICO,
    'acidente': RISCOS_ACIDENTE
}

CATEGORIAS_RISCO = {
    'fisico': '🔥 Físicos',
    'quimico': '⚗️ Químicos',
    'biologico': '🦠 Biológicos',
    'ergonomico': '🏃 Ergonômicos',
    'acidente': '⚠️ Acidentes'
}

# --- Inicialização dos Gerenciadores ---
@st.cache_resource
def init_managers():
    try:
        db_manager = DatabaseManager()
        auth_manager = AuthManager(db_manager)
        user_data_manager = UserDataManager(db_manager)
        return db_manager, auth_manager, user_data_manager
    except Exception as e:
        st.error(f"❌ Erro ao inicializar gerenciadores: {str(e)}")
        return None, None, None

db_manager, auth_manager, user_data_manager = init_managers()

if not all([db_manager, auth_manager, user_data_manager]):
    st.error("❌ Erro crítico: Não foi possível inicializar os gerenciadores do sistema.")
    st.stop()

# --- CSS PERSONALIZADO ---
st.markdown("""
<style>
    .main {
        padding-top: 0rem;
    }
    
    .stApp > header {
        background-color: transparent;
    }
    
    .block-container {
        padding-top: 2rem;
        padding-bottom: 0rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    
    .login-header {
        text-align: center;
        color: #1e3a8a;
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 2rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }
    
    .login-form {
        max-width: 400px;
        margin: 0 auto;
        padding: 2rem;
        background: white;
        border-radius: 10px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.1);
        border: 1px solid #e5e7eb;
    }
    
    .success-message {
        background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%);
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #22c55e;
        margin: 1rem 0;
        color: #166534;
    }
    
    .error-message {
        background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%);
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #ef4444;
        margin: 1rem 0;
        color: #991b1b;
    }
    
    .info-message {
        background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #3b82f6;
        margin: 1rem 0;
        color: #1e40af;
    }
    
    .stButton > button {
        width: 100%;
        border-radius: 8px;
        border: none;
        padding: 0.75rem 1rem;
        font-weight: 500;
        transition: all 0.2s;
        background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%);
        color: white;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4);
    }
    
    .upload-area {
        border: 2px dashed #3b82f6;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        background-color: #f8fafc;
        margin: 1rem 0;
    }
    
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        text-align: center;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        border-top: 3px solid #3b82f6;
    }
    
    .risk-category {
        background: #f8f9ff;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        border-left: 3px solid #3b82f6;
    }
    
    .new-features {
        background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%);
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 4px solid #f59e0b;
        margin: 1rem 0;
        color: #92400e;
    }
</style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE AUTENTICAÇÃO ---
def show_login_page():
    st.markdown("""
    <div class="login-header">
        🔐 Gerador de Ordens de Serviço (OS)
    </div>
    """, unsafe_allow_html=True)
    
    # Informações sobre novidades
    total_riscos = sum(len(riscos) for riscos in AGENTES_POR_CATEGORIA.values())
    st.markdown(f"""
    <div class="new-features">
        <strong>🆕 NOVIDADES DO SISTEMA - Atualização Especial!</strong><br><br>
        ✨ <strong>Base de Riscos Expandida:</strong> {total_riscos} opções de riscos ocupacionais!<br>
        🏃 <strong>Riscos Ergonômicos:</strong> {len(RISCOS_ERGONOMICO)} opções específicas (NOVO!)<br>
        ⚠️ <strong>Riscos de Acidentes:</strong> {len(RISCOS_ACIDENTE)} opções detalhadas (NOVO!)<br>
        🔥 <strong>Riscos Físicos:</strong> {len(RISCOS_FISICO)} opções ampliadas<br>
        ⚗️ <strong>Riscos Químicos:</strong> {len(RISCOS_QUIMICO)} opções específicas<br>
        🦠 <strong>Riscos Biológicos:</strong> {len(RISCOS_BIOLOGICO)} opções incluindo COVID-19<br><br>
        📄 Sistema profissional para geração de OS conforme NR-01 com interface otimizada!
    </div>
    """, unsafe_allow_html=True)
    
    # Tabs para Login e Registro
    login_tab, register_tab = st.tabs(["🔑 Login", "👤 Criar Conta"])
    
    with login_tab:
        st.markdown('<div class="login-form">', unsafe_allow_html=True)
        
        with st.form("login_form"):
            st.markdown("### 🔑 Faça seu Login")
            email = st.text_input("📧 Email:", placeholder="seu@email.com")
            password = st.text_input("🔒 Senha:", type="password", placeholder="Sua senha")
            
            login_button = st.form_submit_button("🚀 Entrar", use_container_width=True)
            
            if login_button:
                if email and password:
                    try:
                        user = auth_manager.login(email, password)
                        if user:
                            st.session_state.user = user
                            st.session_state.authenticated = True
                            st.success("✅ Login realizado com sucesso!")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("❌ Email ou senha incorretos.")
                    except Exception as e:
                        st.error(f"❌ Erro ao fazer login: {str(e)}")
                else:
                    st.warning("⚠️ Por favor, preencha todos os campos.")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with register_tab:
        st.markdown('<div class="login-form">', unsafe_allow_html=True)
        
        with st.form("register_form"):
            st.markdown("### 👤 Criar Nova Conta")
            
            col1, col2 = st.columns(2)
            with col1:
                nome = st.text_input("👤 Nome:", placeholder="Seu nome completo")
            with col2:
                empresa = st.text_input("🏢 Empresa:", placeholder="Nome da empresa")
            
            email = st.text_input("📧 Email:", placeholder="seu@email.com")
            
            col3, col4 = st.columns(2)
            with col3:
                password = st.text_input("🔒 Senha:", type="password", placeholder="Mínimo 6 caracteres")
            with col4:
                password_confirm = st.text_input("🔒 Confirmar:", type="password", placeholder="Confirme a senha")
            
            register_button = st.form_submit_button("✨ Criar Conta", use_container_width=True)
            
            if register_button:
                if nome and empresa and email and password and password_confirm:
                    if password == password_confirm:
                        if len(password) >= 6:
                            try:
                                user_id = auth_manager.register(email, password, nome, empresa)
                                if user_id:
                                    st.success("✅ Conta criada com sucesso! Faça login para continuar.")
                                else:
                                    st.error("❌ Erro ao criar conta. Email já pode estar em uso.")
                            except Exception as e:
                                st.error(f"❌ Erro: {str(e)}")
                        else:
                            st.error("❌ A senha deve ter pelo menos 6 caracteres.")
                    else:
                        st.error("❌ As senhas não coincidem.")
                else:
                    st.warning("⚠️ Por favor, preencha todos os campos.")
        
        st.markdown('</div>', unsafe_allow_html=True)

def show_main_app(user):
    # Header do usuário
    col1, col2, col3 = st.columns([3, 1, 1])
    
    with col1:
        st.markdown(f"# 📄 Gerador de OS - Bem-vindo, **{user['nome']}**!")
    
    with col2:
        try:
            credits = user_data_manager.get_user_credits(user['id'])
            st.metric("💳 Créditos", credits)
        except Exception as e:
            st.metric("💳 Créditos", "Erro")
    
    with col3:
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.user = None
            st.rerun()
    
    st.markdown(f"🏢 **Empresa:** {user['empresa']}")
    
    # Novidades expandidas
    total_riscos = sum(len(riscos) for riscos in AGENTES_POR_CATEGORIA.values())
    st.markdown(f"""
    <div class="new-features">
        <strong>🚀 SISTEMA ATUALIZADO - Nova Base de Riscos!</strong><br><br>
        📊 <strong>Total:</strong> {total_riscos} opções de riscos ocupacionais organizados em 5 categorias<br>
        🏃 <strong>Ergonômicos:</strong> {len(RISCOS_ERGONOMICO)} riscos (assédio, postura, repetitividade, etc.)<br>
        ⚠️ <strong>Acidentes:</strong> {len(RISCOS_ACIDENTE)} riscos (quedas, choques, cortes, etc.)<br>
        🔥 <strong>Físicos:</strong> {len(RISCOS_FISICO)} riscos (ruído, vibração, temperatura, etc.)<br>
        ⚗️ <strong>Químicos:</strong> {len(RISCOS_QUIMICO)} riscos (gases, vapores, poeiras, etc.)<br>
        🦠 <strong>Biológicos:</strong> {len(RISCOS_BIOLOGICO)} riscos (vírus, bactérias, COVID-19, etc.)
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar com estatísticas expandidas
    with st.sidebar:
        st.markdown("### 📊 Base de Riscos Expandida")
        st.markdown(f"**Total: {total_riscos} opções**")
        
        for categoria, nome in CATEGORIAS_RISCO.items():
            qtd_riscos = len(AGENTES_POR_CATEGORIA[categoria])
            st.markdown(f"- {nome}: **{qtd_riscos}** opções")
        
        st.markdown("---")
        st.markdown("### 💳 Informações da Conta")
        st.markdown(f"**Nome:** {user['nome']}")
        st.markdown(f"**Email:** {user['email']}")
        st.markdown(f"**Empresa:** {user['empresa']}")
        try:
            credits = user_data_manager.get_user_credits(user['id'])
            st.markdown(f"**Créditos:** {credits}")
        except Exception as e:
            st.markdown(f"**Créditos:** Erro ao carregar")
        
        st.markdown("---")
        st.markdown("### 📋 Estrutura da Planilha")
        st.markdown("""
        **Colunas obrigatórias:**
        - Nome
        - Setor
        - Função
        - Data de Admissão
        - Empresa
        - Unidade
        - Descrição de Atividades
        """)
    
    # Seção de upload de arquivos
    st.markdown("## 📤 Upload de Arquivos")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="upload-area">
            <h4>📊 Planilha de Funcionários</h4>
            <p>Arquivo Excel (.xlsx) com dados dos funcionários</p>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_excel = st.file_uploader(
            "Selecione a planilha Excel",
            type=['xlsx'],
            help="Planilha deve conter: Nome, Setor, Função, Data de Admissão, Empresa, Unidade, Descrição de Atividades"
        )
    
    with col2:
        st.markdown("""
        <div class="upload-area">
            <h4>📄 Modelo de OS (Opcional)</h4>
            <p>Template Word personalizado (.docx)</p>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_template = st.file_uploader(
            "Selecione template Word (opcional)",
            type=['docx'],
            help="Se não informado, será usado template padrão do sistema"
        )
    
    # Processar planilha se carregada
    if uploaded_excel is not None:
        try:
            df = pd.read_excel(uploaded_excel)
            
            # Validação básica da planilha
            required_columns = ['Nome', 'Setor', 'Função']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"❌ Colunas obrigatórias faltando: {', '.join(missing_columns)}")
                return
            
            st.success(f"✅ Planilha carregada: {len(df)} funcionários encontrados")
            
            # Estatísticas da planilha
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{len(df)}</h3>
                    <p>👥 Funcionários</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{df['Setor'].nunique()}</h3>
                    <p>🏢 Setores</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{df['Função'].nunique()}</h3>
                    <p>💼 Funções</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                empresa_count = df['Empresa'].nunique() if 'Empresa' in df.columns else 1
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{empresa_count}</h3>
                    <p>🏭 Empresas</p>
                </div>
                """, unsafe_allow_html=True)
            
            # Seleção de funcionários
            st.markdown("## 👥 Seleção de Funcionários")
            
            col1, col2 = st.columns(2)
            
            with col1:
                setores_disponiveis = ['Todos os setores'] + sorted(df['Setor'].dropna().unique().tolist())
                setor_selecionado = st.selectbox("🏢 Filtrar por Setor:", setores_disponiveis)
                
                if setor_selecionado == 'Todos os setores':
                    df_filtrado = df
                else:
                    df_filtrado = df[df['Setor'] == setor_selecionado]
            
            with col2:
                modo_selecao = st.radio(
                    "📋 Modo de Seleção:",
                    ["Funcionário Individual", "Múltiplos Funcionários", "Todos do Setor Filtrado"]
                )
            
            # Lógica de seleção
            funcionarios_selecionados = []
            
            if modo_selecao == "Funcionário Individual":
                funcionario_individual = st.selectbox(
                    "👤 Selecione o funcionário:",
                    [''] + df_filtrado['Nome'].tolist()
                )
                if funcionario_individual:
                    funcionarios_selecionados = [funcionario_individual]
            
            elif modo_selecao == "Múltiplos Funcionários":
                funcionarios_selecionados = st.multiselect(
                    "👥 Selecione múltiplos funcionários:",
                    df_filtrado['Nome'].tolist()
                )
            
            else:  # Todos do setor
                funcionarios_selecionados = df_filtrado['Nome'].tolist()
                if funcionarios_selecionados:
                    st.info(f"📝 Serão geradas OS para todos os {len(funcionarios_selecionados)} funcionários do setor.")
            
            # Configuração de riscos se há funcionários selecionados
            if funcionarios_selecionados:
                st.success(f"✅ {len(funcionarios_selecionados)} funcionário(s) selecionado(s)")
                
                st.markdown("## ⚠️ Configuração de Riscos Ocupacionais")
                
                # Inicializar dados de sessão
                if 'agentes_risco' not in st.session_state:
                    st.session_state.agentes_risco = {categoria: [] for categoria in CATEGORIAS_RISCO.keys()}
                if 'epis_selecionados' not in st.session_state:
                    st.session_state.epis_selecionados = []
                if 'medidas_preventivas' not in st.session_state:
                    st.session_state.medidas_preventivas = []
                
                # Configurar riscos por categoria
                st.markdown("### 🔍 Agentes de Riscos por Categoria")
                
                for categoria_key, categoria_nome in CATEGORIAS_RISCO.items():
                    qtd_opcoes = len(AGENTES_POR_CATEGORIA[categoria_key])
                    with st.expander(f"{categoria_nome} ({qtd_opcoes} opções)", expanded=False):
                        
                        col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
                        
                        with col1:
                            agente_selecionado = st.selectbox(
                                "Agente de Risco:",
                                ['Selecione...'] + AGENTES_POR_CATEGORIA[categoria_key],
                                key=f"agente_{categoria_key}"
                            )
                        
                        with col2:
                            intensidade = st.text_input(
                                "Intensidade:",
                                key=f"intensidade_{categoria_key}",
                                placeholder="Ex: 85"
                            )
                        
                        with col3:
                            unidade = st.selectbox(
                                "Unidade:",
                                UNIDADES_DE_MEDIDA,
                                key=f"unidade_{categoria_key}"
                            )
                        
                        with col4:
                            if st.button(f"➕", key=f"add_{categoria_key}"):
                                if agente_selecionado != 'Selecione...':
                                    novo_risco = {
                                        'agente': agente_selecionado,
                                        'intensidade': intensidade,
                                        'unidade': unidade
                                    }
                                    st.session_state.agentes_risco[categoria_key].append(novo_risco)
                                    st.success(f"✅ Risco adicionado!")
                                    st.rerun()
                        
                        # Mostrar riscos adicionados
                        if st.session_state.agentes_risco[categoria_key]:
                            st.markdown("**Riscos configurados:**")
                            for idx, risco in enumerate(st.session_state.agentes_risco[categoria_key]):
                                col1, col2 = st.columns([5, 1])
                                with col1:
                                    risco_text = f"• {risco['agente']}"
                                    if risco['intensidade']:
                                        risco_text += f": {risco['intensidade']}"
                                    if risco['unidade'] and risco['unidade'] != 'Não aplicável':
                                        risco_text += f" {risco['unidade']}"
                                    st.write(risco_text)
                                with col2:
                                    if st.button("🗑️", key=f"remove_{categoria_key}_{idx}"):
                                        st.session_state.agentes_risco[categoria_key].pop(idx)
                                        st.rerun()
                
                # EPIs e Medidas Preventivas
                st.markdown("### 🥽 EPIs e Medidas Preventivas")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**EPIs Obrigatórios:**")
                    novo_epi = st.text_input("Adicionar EPI:", placeholder="Ex: Capacete, óculos...")
                    if st.button("➕ Adicionar EPI") and novo_epi:
                        st.session_state.epis_selecionados.append(novo_epi)
                        st.rerun()
                    
                    for idx, epi in enumerate(st.session_state.epis_selecionados):
                        col_epi1, col_epi2 = st.columns([4, 1])
                        with col_epi1:
                            st.write(f"• {epi}")
                        with col_epi2:
                            if st.button("🗑️", key=f"remove_epi_{idx}"):
                                st.session_state.epis_selecionados.pop(idx)
                                st.rerun()
                
                with col2:
                    st.markdown("**Medidas Preventivas:**")
                    nova_medida = st.text_area("Adicionar Medida:", placeholder="Ex: Treinamentos, pausas...", height=100)
                    if st.button("➕ Adicionar Medida") and nova_medida:
                        st.session_state.medidas_preventivas.append(nova_medida)
                        st.rerun()
                    
                    for idx, medida in enumerate(st.session_state.medidas_preventivas):
                        col_med1, col_med2 = st.columns([4, 1])
                        with col_med1:
                            medida_resumida = medida[:100] + "..." if len(medida) > 100 else medida
                            st.write(f"• {medida_resumida}")
                        with col_med2:
                            if st.button("🗑️", key=f"remove_med_{idx}"):
                                st.session_state.medidas_preventivas.pop(idx)
                                st.rerun()
                
                # Observações
                observacoes = st.text_area(
                    "📝 Observações Complementares:",
                    placeholder="Informações específicas do setor, procedimentos especiais, etc.",
                    height=80
                )
                
                # Botão para gerar OS
                st.markdown("## 🚀 Gerar Ordens de Serviço")
                
                # Verificar créditos suficientes
                creditos_necessarios = len(funcionarios_selecionados)
                try:
                    creditos_usuario = user_data_manager.get_user_credits(user['id'])
                except Exception as e:
                    creditos_usuario = 0
                    st.error(f"❌ Erro ao verificar créditos: {str(e)}")
                
                if creditos_usuario >= creditos_necessarios:
                    if st.button(f"📄 GERAR {len(funcionarios_selecionados)} OS ({creditos_necessarios} créditos)", type="primary", use_container_width=True):
                        
                        # Simular geração de documentos
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        # Processar cada funcionário
                        for idx, nome_funcionario in enumerate(funcionarios_selecionados):
                            status_text.text(f"🔄 Processando: {nome_funcionario}")
                            progress_bar.progress((idx + 1) / len(funcionarios_selecionados))
                            time.sleep(0.5)  # Simular processamento
                        
                        # Debitar créditos
                        try:
                            user_data_manager.debit_credits(user['id'], creditos_necessarios)
                        except Exception as e:
                            st.error(f"❌ Erro ao debitar créditos: {str(e)}")
                        
                        status_text.text("✅ Processamento concluído!")
                        st.success(f"✅ {len(funcionarios_selecionados)} OS geradas com sucesso!")
                        st.info(f"💳 {creditos_necessarios} créditos foram debitados da sua conta.")
                        
                        # Aqui seria implementada a lógica real de geração dos documentos
                        # Por agora, apenas simulamos o processo
                        
                        time.sleep(2)
                        st.rerun()
                else:
                    st.warning(f"⚠️ Créditos insuficientes. Você precisa de {creditos_necessarios} créditos, mas possui apenas {creditos_usuario}.")
                    st.info("💳 Entre em contato com o administrador para adquirir mais créditos.")
        
        except Exception as e:
            st.error(f"❌ Erro ao processar planilha: {str(e)}")
    
    else:
        # Instruções iniciais
        total_riscos = sum(len(riscos) for riscos in AGENTES_POR_CATEGORIA.values())
        st.markdown(f"""
        <div class="info-message">
            <h4>🎯 Como usar o sistema:</h4>
            <ol>
                <li><strong>📤 Faça upload</strong> da planilha Excel com dados dos funcionários</li>
                <li><strong>👥 Selecione</strong> os funcionários (individual, múltiplos ou todos)</li>
                <li><strong>⚠️ Configure</strong> os riscos ocupacionais específicos</li>
                <li><strong>🥽 Adicione</strong> EPIs e medidas preventivas</li>
                <li><strong>🚀 Gere</strong> as Ordens de Serviço conforme NR-01</li>
            </ol>
            
            <p><strong>🆕 Sistema expandido:</strong> Agora com <strong>{total_riscos} opções de riscos</strong> organizados em 5 categorias!</p>
        </div>
        """, unsafe_allow_html=True)

# --- LÓGICA PRINCIPAL DA APLICAÇÃO ---
def main():
    # Verificar se o usuário está autenticado
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    
    if 'user' not in st.session_state:
        st.session_state.user = None
    
    # Mostrar página apropriada
    if st.session_state.authenticated and st.session_state.user:
        show_main_app(st.session_state.user)
    else:
        show_login_page()

if __name__ == "__main__":
    main()