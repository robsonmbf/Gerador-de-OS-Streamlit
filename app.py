# 🔐 Gerador de Ordens de Serviço (OS)

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
import datetime
import hashlib

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Ordens de Serviço (OS)",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
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

# --- CSS DARK MINIMALISTA ---
st.markdown("""
<style>
    /* TEMA DARK MINIMALISTA - SEM CORES VIBRANTES */
    .stApp {
        background: #0d1117 !important;
        color: #e6edf3 !important;
    }
    
    .main {
        background: #0d1117 !important;
        color: #e6edf3 !important;
    }
    
    /* OCULTAR HEADER COMPLETAMENTE */
    header[data-testid="stHeader"] {
        height: 0px;
        max-height: 0px;
        overflow: hidden;
    }
    
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
    }
    
    /* TÍTULO PRINCIPAL - MINIMALISTA */
    .title-header {
        text-align: center;
        color: #f0f6fc;
        font-size: 2.5rem;
        font-weight: 300;
        margin: 2rem 0;
        letter-spacing: 1px;
    }
    
    /* FORMULÁRIO DE LOGIN - MINIMALISTA */
    .login-container {
        max-width: 450px;
        margin: 2rem auto;
        padding: 2rem;
        background: #161b22;
        border-radius: 8px;
        border: 1px solid #30363d;
    }
    
    .login-title {
        text-align: center;
        color: #f0f6fc;
        font-size: 1.5rem;
        font-weight: 400;
        margin-bottom: 2rem;
    }
    
    /* BOTÕES MINIMALISTAS */
    .stButton > button {
        background: #21262d;
        color: #f0f6fc;
        border: 1px solid #30363d;
        border-radius: 6px;
        padding: 0.5rem 1rem;
        font-size: 1rem;
        font-weight: 400;
        width: 100%;
        transition: all 0.2s ease;
    }
    
    .stButton > button:hover {
        background: #30363d;
        border-color: #484f58;
    }
    
    .stButton > button:focus {
        background: #30363d;
        border-color: #58a6ff;
        box-shadow: 0 0 0 3px rgba(88, 166, 255, 0.1);
    }
    
    /* INPUTS MINIMALISTAS */
    .stTextInput > div > div > input {
        background: #0d1117 !important;
        color: #f0f6fc !important;
        border: 1px solid #30363d !important;
        border-radius: 6px !important;
        padding: 0.5rem !important;
        font-size: 14px !important;
    }
    
    .stTextInput > div > div > input:focus {
        border-color: #58a6ff !important;
        box-shadow: 0 0 0 3px rgba(88, 166, 255, 0.1) !important;
        outline: none !important;
    }
    
    /* SELECTBOX MINIMALISTA */
    .stSelectbox > div > div {
        background: #0d1117 !important;
        color: #f0f6fc !important;
        border: 1px solid #30363d !important;
        border-radius: 6px !important;
    }
    
    .stSelectbox [data-baseweb="select"] {
        background: #0d1117 !important;
        border: 1px solid #30363d !important;
        border-radius: 6px !important;
    }
    
    /* MULTISELECT MINIMALISTA */
    .stMultiSelect > div > div {
        background: #0d1117 !important;
        color: #f0f6fc !important;
        border: 1px solid #30363d !important;
        border-radius: 6px !important;
    }
    
    .stMultiSelect [data-baseweb="select"] {
        background: #0d1117 !important;
        border: 1px solid #30363d !important;
        border-radius: 6px !important;
    }
    
    /* MENSAGENS MINIMALISTAS */
    .success-msg {
        background: rgba(46, 160, 67, 0.1);
        border: 1px solid #238636;
        border-radius: 6px;
        padding: 12px;
        color: #7ee787;
        margin: 1rem 0;
        font-size: 14px;
    }
    
    .error-msg {
        background: rgba(248, 81, 73, 0.1);
        border: 1px solid #da3633;
        border-radius: 6px;
        padding: 12px;
        color: #f85149;
        margin: 1rem 0;
        font-size: 14px;
    }
    
    .info-msg {
        background: rgba(56, 139, 253, 0.1);
        border: 1px solid #1f6feb;
        border-radius: 6px;
        padding: 12px;
        color: #79c0ff;
        margin: 1rem 0;
        font-size: 14px;
    }
    
    .warning-msg {
        background: rgba(187, 128, 9, 0.1);
        border: 1px solid #bb8009;
        border-radius: 6px;
        padding: 12px;
        color: #f2cc60;
        margin: 1rem 0;
        font-size: 14px;
    }
    
    /* TABS MINIMALISTAS */
    .stTabs [data-baseweb="tab-list"] {
        background: transparent;
        border-bottom: 1px solid #30363d;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        color: #7d8590;
        border: none;
        padding: 8px 16px;
    }
    
    .stTabs [aria-selected="true"] {
        background: transparent !important;
        color: #f0f6fc !important;
        border-bottom: 2px solid #58a6ff !important;
    }
    
    /* SIDEBAR MINIMALISTA */
    .css-1d391kg {
        background: #161b22 !important;
        border-right: 1px solid #30363d;
    }
    
    /* UPLOAD AREAS MINIMALISTAS */
    .upload-area {
        background: #161b22;
        border: 2px dashed #30363d;
        border-radius: 8px;
        padding: 2rem;
        text-align: center;
        margin: 1rem 0;
        color: #7d8590;
    }
    
    .upload-area:hover {
        border-color: #58a6ff;
        background: #0d1117;
    }
    
    /* CARDS MINIMALISTAS */
    .metric-card {
        background: #161b22;
        padding: 1rem;
        border-radius: 8px;
        text-align: center;
        border: 1px solid #30363d;
        margin: 0.5rem 0;
    }
    
    .metric-card h3 {
        color: #f0f6fc;
        font-size: 1.5rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .metric-card p {
        color: #7d8590;
        margin: 0;
        font-size: 14px;
    }
    
    /* EXPANSORES MINIMALISTAS */
    .streamlit-expanderHeader {
        background: #161b22 !important;
        border: 1px solid #30363d !important;
        border-radius: 6px !important;
        color: #f0f6fc !important;
    }
    
    .streamlit-expanderContent {
        background: #0d1117 !important;
        border: 1px solid #30363d !important;
        color: #f0f6fc !important;
        border-radius: 0 0 6px 6px !important;
    }
    
    /* TEXT AREAS MINIMALISTAS */
    .stTextArea > div > div > textarea {
        background: #0d1117 !important;
        color: #f0f6fc !important;
        border: 1px solid #30363d !important;
        border-radius: 6px !important;
    }
    
    .stTextArea > div > div > textarea:focus {
        border-color: #58a6ff !important;
        box-shadow: 0 0 0 3px rgba(88, 166, 255, 0.1) !important;
    }
    
    /* RADIO BUTTONS MINIMALISTAS */
    .stRadio > div {
        background: transparent !important;
        color: #f0f6fc !important;
    }
    
    /* PROGRESS BAR MINIMALISTA */
    .stProgress > div > div > div {
        background: #58a6ff !important;
    }
    
    /* MÉTRICAS MINIMALISTAS */
    [data-testid="metric-container"] {
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 1rem;
    }
    
    [data-testid="metric-container"] > div {
        color: #f0f6fc;
    }
    
    [data-testid="metric-container"] [data-testid="metric-value"] {
        color: #f0f6fc;
    }
    
    [data-testid="metric-container"] [data-testid="metric-label"] {
        color: #7d8590;
    }
    
    /* REMOVER CORES DE FUNDO PADRÃO */
    .stApp > div > div > div > div {
        background: transparent !important;
    }
</style>
""", unsafe_allow_html=True)

# --- VALIDAÇÃO DE EMAIL ---
def is_valid_email(email):
    """Valida se o email é de um provedor real (Gmail, Outlook, etc.)"""
    if not email or '@' not in email:
        return False
    
    # Padrão básico de email
    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    if not re.match(email_pattern, email):
        return False
    
    # Verificar provedores válidos
    valid_domains = [
        'gmail.com', 'outlook.com', 'hotmail.com', 'yahoo.com', 'yahoo.com.br',
        'uol.com.br', 'terra.com.br', 'bol.com.br', 'ig.com.br', 'globo.com',
        'live.com', 'msn.com', 'icloud.com', 'me.com', 'mac.com',
        'protonmail.com', 'zoho.com', 'yandex.com'
    ]
    
    domain = email.split('@')[1].lower()
    return domain in valid_domains

# --- SISTEMA DE AUTENTICAÇÃO ---
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def initialize_users():
    if 'users_db' not in st.session_state:
        st.session_state.users_db = {
            'robsonmbf@hotmail.com': {
                'id': 1,
                'password': hash_password('123456'),
                'nome': 'Robson',
                'empresa': 'Minha Empresa',
                'email': 'robsonmbf@hotmail.com',
                'credits': 999999,
                'is_admin': True
            }
        }

def authenticate_user(email, password):
    initialize_users()
    hashed_password = hash_password(password)
    
    if email in st.session_state.users_db:
        user_data = st.session_state.users_db[email]
        if user_data['password'] == hashed_password:
            return user_data
    return None

def register_user(email, password, nome, empresa):
    initialize_users()
    
    if email in st.session_state.users_db:
        return None
    
    user_id = len(st.session_state.users_db) + 1
    st.session_state.users_db[email] = {
        'id': user_id,
        'password': hash_password(password),
        'nome': nome,
        'empresa': empresa,
        'email': email,
        'credits': 100,
        'is_admin': False
    }
    return user_id

def get_user_credits(user_id):
    initialize_users()
    for user in st.session_state.users_db.values():
        if user['id'] == user_id:
            if user.get('is_admin', False):
                return "∞"
            return user['credits']
    return 0

def debit_credits(user_id, amount):
    initialize_users()
    for user in st.session_state.users_db.values():
        if user['id'] == user_id:
            if user.get('is_admin', False):
                return True
            user['credits'] = max(0, user['credits'] - amount)
            return True
    return False

def check_sufficient_credits(user_id, amount):
    initialize_users()
    for user in st.session_state.users_db.values():
        if user['id'] == user_id:
            if user.get('is_admin', False):
                return True
            return user['credits'] >= amount
    return False

# --- FUNÇÕES AUXILIARES ---
def create_sample_data():
    sample_data = {
        'Nome': ['JOÃO SILVA SANTOS', 'MARIA OLIVEIRA COSTA', 'PEDRO ALVES FERREIRA'],
        'Setor': ['PRODUCAO DE LA DE ACO', 'ADMINISTRACAO DE RH', 'MANUTENCAO QUIMICA'],
        'Função': ['OPERADOR PRODUCAO I', 'ANALISTA ADM PESSOAL PL', 'MECANICO MANUT II'],
        'Data de Admissão': ['15/03/2020', '22/08/2019', '10/01/2021'],
        'Empresa': ['SUA EMPRESA', 'SUA EMPRESA', 'SUA EMPRESA'],
        'Unidade': ['Matriz', 'Matriz', 'Matriz'],
        'Descrição de Atividades': [
            'Operar equipamentos de produção nível I, controlar parâmetros operacionais, realizar inspeções visuais e registrar dados de produção.',
            'Executar atividades de administração de pessoal, controlar documentos trabalhistas, elaborar relatórios e dar suporte às equipes.',
            'Executar manutenção preventiva e corretiva em equipamentos, diagnosticar falhas, trocar componentes e registrar intervenções.'
        ]
    }
    return pd.DataFrame(sample_data)

def validate_excel_structure(df):
    required_columns = ['Nome', 'Setor', 'Função', 'Data de Admissão', 'Empresa', 'Unidade', 'Descrição de Atividades']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        return False, f"Colunas obrigatórias faltando: {', '.join(missing_columns)}"
    
    if df.empty:
        return False, "A planilha está vazia"
    
    return True, "Estrutura válida"

def substituir_placeholders_no_documento(doc, dados_funcionario, agentes_risco, epis, medidas_preventivas, observacoes, medicoes=""):
    """
    Substitui os placeholders no documento template pelos dados reais
    CORRIGIDO: Se não houver riscos selecionados, preenche com "Ausência de Fator de Risco"
    """
    try:
        # Dicionário de substituições básicas
        substituicoes = {
            '[NOME EMPRESA]': dados_funcionario.get('Empresa', ''),
            '[UNIDADE]': dados_funcionario.get('Unidade', ''),
            '[NOME FUNCIONÁRIO]': dados_funcionario.get('Nome', ''),
            '[DATA DE ADMISSÃO]': dados_funcionario.get('Data de Admissão', ''),
            '[SETOR]': dados_funcionario.get('Setor', ''),
            '[FUNÇÃO]': dados_funcionario.get('Função', ''),
            '[DESCRIÇÃO DE ATIVIDADES]': dados_funcionario.get('Descrição de Atividades', ''),
            '[MEDIÇÕES]': medicoes if medicoes else "Não aplicável para esta função."
        }
        
        # Preparar textos dos riscos com "Ausência de Fator de Risco" quando vazio
        riscos_texto = {}
        danos_texto = {}
        
        for categoria in ['fisico', 'quimico', 'biologico', 'ergonomico', 'acidente']:
            categoria_upper = categoria.upper()
            if categoria == 'fisico':
                categoria_nome = 'FÍSICOS'
            elif categoria == 'quimico':
                categoria_nome = 'QUÍMICOS'
            elif categoria == 'biologico':
                categoria_nome = 'BIOLÓGICOS'
            elif categoria == 'ergonomico':
                categoria_nome = 'ERGONÔMICOS'
            elif categoria == 'acidente':
                categoria_nome = 'ACIDENTE'
            
            # Montar texto dos riscos - CORREÇÃO AQUI
            if categoria in agentes_risco and agentes_risco[categoria]:
                riscos_lista = []
                for risco in agentes_risco[categoria]:
                    risco_text = risco['agente']
                    if risco.get('intensidade'):
                        risco_text += f": {risco['intensidade']}"
                    if risco.get('unidade') and risco['unidade'] != 'Não aplicável':
                        risco_text += f" {risco['unidade']}"
                    riscos_lista.append(risco_text)
                riscos_texto[f'[RISCOS {categoria_nome}]'] = '; '.join(riscos_lista)
                
                # Possíveis danos (texto genérico baseado na categoria)
                if categoria == 'fisico':
                    danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "Perda auditiva, lesões por vibração, queimaduras, hipotermia, hipertermia"
                elif categoria == 'quimico':
                    danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "Intoxicação, dermatoses, pneumoconioses, alergias respiratórias"
                elif categoria == 'biologico':
                    danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "Infecções, doenças infectocontagiosas, alergias"
                elif categoria == 'ergonomico':
                    danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "LER/DORT, fadiga, estresse, dores musculares"
                elif categoria == 'acidente':
                    danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "Fraturas, cortes, contusões, queimaduras, morte"
            else:
                # *** CORREÇÃO PRINCIPAL: Usar "Ausência de Fator de Risco" em vez de "Não identificados" ***
                riscos_texto[f'[RISCOS {categoria_nome}]'] = "Ausência de Fator de Risco"
                danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "Não aplicável"
        
        # Adicionar riscos e danos às substituições
        substituicoes.update(riscos_texto)
        substituicoes.update(danos_texto)
        
        # Preparar texto dos EPIs
        if epis:
            substituicoes['[EPIS]'] = '; '.join(epis)
        else:
            substituicoes['[EPIS]'] = "Conforme análise de risco específica da função"
        
        # Substituir nos parágrafos
        for paragrafo in doc.paragraphs:
            for placeholder, valor in substituicoes.items():
                if placeholder in paragrafo.text:
                    paragrafo.text = paragrafo.text.replace(placeholder, valor)
        
        # Substituir nas tabelas (se houver)
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for placeholder, valor in substituicoes.items():
                        if placeholder in celula.text:
                            celula.text = celula.text.replace(placeholder, valor)
        
        return doc
        
    except Exception as e:
        st.error(f"Erro ao substituir placeholders: {str(e)}")
        return None

def gerar_documento_os(dados_funcionario, agentes_risco, epis, medidas_preventivas, observacoes, uploaded_template=None):
    """
    Função para gerar documento da OS - CORRIGIDA para usar template uploaded
    """
    try:
        # Se um template foi fornecido, usar ele
        if uploaded_template is not None:
            doc = Document(uploaded_template)
            # Substituir placeholders no template
            doc = substituir_placeholders_no_documento(
                doc, dados_funcionario, agentes_risco, epis, medidas_preventivas, observacoes
            )
        else:
            # Código original para criar documento do zero (fallback)
            doc = Document()
            
            titulo = doc.add_heading('ORDEM DE SERVIÇO', 0)
            titulo.alignment = 1
            
            subtitulo = doc.add_paragraph('Informações sobre Condições de Segurança e Saúde no Trabalho - NR-01')
            subtitulo.alignment = 1
            
            doc.add_paragraph()
            
            # Informações do Funcionário
            info_func = doc.add_paragraph()
            info_func.add_run(f"Empresa: {dados_funcionario.get('Empresa', '')}\t\t")
            info_func.add_run(f"Unidade: {dados_funcionario.get('Unidade', '')}")
            
            info_func2 = doc.add_paragraph()
            info_func2.add_run(f"Nome do Funcionário: {dados_funcionario.get('Nome', '')}")
            
            info_func3 = doc.add_paragraph()
            info_func3.add_run(f"Data de Admissão: {dados_funcionario.get('Data de Admissão', '')}")
            
            info_func4 = doc.add_paragraph()
            info_func4.add_run(f"Setor de Trabalho: {dados_funcionario.get('Setor', '')}\t\t")
            info_func4.add_run(f"Função: {dados_funcionario.get('Função', '')}")
            
            doc.add_paragraph()
            
            doc.add_heading('TAREFAS DA FUNÇÃO', level=1)
            doc.add_paragraph(dados_funcionario.get('Descrição de Atividades', 'Atividades relacionadas à função exercida.'))
            
            # CORREÇÃO: Verificar se há riscos, senão usar "Ausência de Fator de Risco"
            tem_riscos = any(agentes_risco.get(categoria, []) for categoria in agentes_risco.keys()) if agentes_risco else False
            
            doc.add_heading('AGENTES DE RISCOS OCUPACIONAIS', level=1)
            
            if tem_riscos:
                for categoria, riscos in agentes_risco.items():
                    if riscos:
                        categoria_titulo = categoria.replace('_', ' ').title()
                        doc.add_heading(f'Riscos {categoria_titulo}', level=2)
                        
                        for risco in riscos:
                            risco_para = doc.add_paragraph()
                            risco_para.add_run(f"• {risco['agente']}")
                            if risco.get('intensidade'):
                                risco_para.add_run(f": {risco['intensidade']}")
                            if risco.get('unidade'):
                                risco_para.add_run(f" {risco['unidade']}")
            else:
                # Se não há riscos, adicionar "Ausência de Fator de Risco"
                doc.add_paragraph("Ausência de Fator de Risco")
            
            if epis:
                doc.add_heading('EQUIPAMENTOS DE PROTEÇÃO INDIVIDUAL (EPIs)', level=1)
                for epi in epis:
                    doc.add_paragraph(f"• {epi}", style='List Bullet')
            else:
                doc.add_heading('EQUIPAMENTOS DE PROTEÇÃO INDIVIDUAL (EPIs)', level=1)
                doc.add_paragraph("Conforme análise de risco específica da função")
            
            if medidas_preventivas:
                doc.add_heading('MEDIDAS PREVENTIVAS E DE CONTROLE', level=1)
                for medida in medidas_preventivas:
                    doc.add_paragraph(f"• {medida}", style='List Bullet')
            
            doc.add_heading('PROCEDIMENTOS EM SITUAÇÕES DE EMERGÊNCIA', level=1)
            emergencia_texto = """• Comunique imediatamente o acidente à chefia imediata ou responsável pela área;
• Preserve as condições do local de acidente até a comunicação com a autoridade competente;
• Procure atendimento médico no ambulatório da empresa ou serviço médico de emergência;
• Siga as orientações do Plano de Emergência da empresa;
• Registre a ocorrência conforme procedimentos estabelecidos."""
            doc.add_paragraph(emergencia_texto)
            
            doc.add_heading('ORIENTAÇÕES SOBRE GRAVE E IMINENTE RISCO', level=1)
            gir_texto = """• Sempre que constatar condição de grave e iminente risco, interrompa imediatamente as atividades;
• Comunique de forma urgente ao seu superior hierárquico;
• Aguarde as providências necessárias e autorização para retorno;
• É direito do trabalhador recusar-se a trabalhar em condições de risco grave e iminente."""
            doc.add_paragraph(gir_texto)
            
            if observacoes:
                doc.add_heading('OBSERVAÇÕES COMPLEMENTARES', level=1)
                doc.add_paragraph(observacoes)
            
            doc.add_paragraph()
            nota_legal = doc.add_paragraph()
            nota_legal.add_run("IMPORTANTE: ").bold = True
            nota_legal.add_run(
                "Conforme Art. 158 da CLT e NR-01, o descumprimento das disposições "
                "sobre segurança e saúde no trabalho sujeita o empregado às penalidades "
                "legais, inclusive demissão por justa causa."
            )
            
            doc.add_paragraph()
            doc.add_paragraph("_" * 40 + "\t\t" + "_" * 40)
            doc.add_paragraph("Funcionário\t\t\t\t\tResponsável pela Área")
            doc.add_paragraph(f"Data: {datetime.date.today().strftime('%d/%m/%Y')}")
        
        return doc
        
    except Exception as e:
        st.error(f"Erro ao gerar documento: {str(e)}")
        return None

# --- FUNÇÃO DE LOGIN ---
def show_login_page():
    st.markdown('<div class="title-header">Gerador de Ordens de Serviço</div>', unsafe_allow_html=True)
    
    total_riscos = sum(len(riscos) for riscos in AGENTES_POR_CATEGORIA.values())
    
    st.markdown(f"""
    <div class="info-msg">
        <strong>Sistema Atualizado - Base Expandida de Riscos</strong><br><br>
        ✅ <strong>{total_riscos} opções</strong> de riscos ocupacionais organizados em 5 categorias<br>
        ✅ <strong>Ergonômicos:</strong> {len(RISCOS_ERGONOMICO)} opções específicas<br>
        ✅ <strong>Acidentes:</strong> {len(RISCOS_ACIDENTE)} opções detalhadas<br>
        ✅ <strong>Físicos:</strong> {len(RISCOS_FISICO)} opções ampliadas<br>
        ✅ <strong>Químicos:</strong> {len(RISCOS_QUIMICO)} opções específicas<br>
        ✅ <strong>Biológicos:</strong> {len(RISCOS_BIOLOGICO)} incluindo COVID-19<br><br>
        Sistema profissional conforme NR-01 com tema minimalista
    </div>
    """, unsafe_allow_html=True)
    
    login_tab, register_tab = st.tabs(["Login", "Criar Conta"])
    
    with login_tab:
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        st.markdown('<div class="login-title">Acesso ao Sistema</div>', unsafe_allow_html=True)
        
        # MOSTRAR DADOS DE ACESSO ADMIN
        st.markdown("""
        <div class="warning-msg">
            <strong>👑 Dados de Acesso Administrador:</strong><br>
            📧 <strong>Email:</strong> robsonmbf@hotmail.com<br>
            🔒 <strong>Senha:</strong> 123456<br>
            💳 <strong>Créditos:</strong> Ilimitados
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            email = st.text_input("Email:", placeholder="seu@gmail.com")
            password = st.text_input("Senha:", type="password", placeholder="Sua senha")
            
            login_button = st.form_submit_button("Entrar no Sistema")
            
            if login_button:
                if email and password:
                    if is_valid_email(email):
                        user = authenticate_user(email, password)
                        if user:
                            st.session_state.user = user
                            st.session_state.authenticated = True
                            st.markdown('<div class="success-msg">Login realizado com sucesso</div>', unsafe_allow_html=True)
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.markdown('<div class="error-msg">Email ou senha incorretos</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="error-msg">Email deve ser de um provedor válido (Gmail, Outlook, Yahoo, etc.)</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="error-msg">Por favor, preencha todos os campos</div>', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with register_tab:
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        st.markdown('<div class="login-title">Criar Nova Conta</div>', unsafe_allow_html=True)
        
        with st.form("register_form"):
            col1, col2 = st.columns(2)
            with col1:
                nome = st.text_input("Nome:", placeholder="Seu nome completo")
            with col2:
                empresa = st.text_input("Empresa:", placeholder="Nome da empresa")
            
            email = st.text_input("Email:", placeholder="seu@gmail.com")
            
            col3, col4 = st.columns(2)
            with col3:
                password = st.text_input("Senha:", type="password", placeholder="Mínimo 6 caracteres")
            with col4:
                password_confirm = st.text_input("Confirmar:", type="password", placeholder="Confirme a senha")
            
            register_button = st.form_submit_button("Criar Conta")
            
            if register_button:
                if nome and empresa and email and password and password_confirm:
                    if is_valid_email(email):
                        if password == password_confirm:
                            if len(password) >= 6:
                                user_id = register_user(email, password, nome, empresa)
                                if user_id:
                                    st.markdown('<div class="success-msg">Conta criada com sucesso! Faça login para continuar</div>', unsafe_allow_html=True)
                                else:
                                    st.markdown('<div class="error-msg">Erro ao criar conta. Email já pode estar em uso</div>', unsafe_allow_html=True)
                            else:
                                st.markdown('<div class="error-msg">A senha deve ter pelo menos 6 caracteres</div>', unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="error-msg">As senhas não coincidem</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="error-msg">Email deve ser de um provedor válido (Gmail, Outlook, Yahoo, etc.)</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="error-msg">Por favor, preencha todos os campos</div>', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

def show_main_app(user):
    # Header do usuário
    col1, col2, col3 = st.columns([3, 1, 1])
    
    with col1:
        st.markdown(f"# Gerador de OS - Bem-vindo, **{user['nome']}**")
    
    with col2:
        credits = get_user_credits(user['id'])
        st.metric("Créditos", credits)
    
    with col3:
        if st.button("Sair"):
            st.session_state.authenticated = False
            st.session_state.user = None
            st.rerun()
    
    st.markdown(f"**Empresa:** {user['empresa']}")
    
    # Mostrar status de admin se for o caso
    if user.get('is_admin', False):
        st.markdown("""
        <div class="warning-msg">
            <strong>Conta Administrador</strong><br>
            • Créditos ilimitados<br>
            • Não há cobrança de créditos<br>
            • Acesso completo ao sistema
        </div>
        """, unsafe_allow_html=True)
    
    total_riscos = sum(len(riscos) for riscos in AGENTES_POR_CATEGORIA.values())
    st.markdown(f"""
    <div class="info-msg">
        <strong>Sistema Atualizado</strong><br><br>
        📊 <strong>Total:</strong> {total_riscos} opções de riscos ocupacionais organizados em 5 categorias<br>
        ✅ <strong>Preenchimento automático:</strong> "Ausência de Fator de Risco" quando necessário
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar com informações
    with st.sidebar:
        st.markdown("### Base de Riscos")
        st.markdown(f"**Total: {total_riscos} opções**")
        
        for categoria, nome in CATEGORIAS_RISCO.items():
            qtd_riscos = len(AGENTES_POR_CATEGORIA[categoria])
            st.markdown(f"- {nome}: **{qtd_riscos}** opções")
        
        st.markdown("---")
        st.markdown("### Informações da Conta")
        st.markdown(f"**Nome:** {user['nome']}")
        st.markdown(f"**Email:** {user['email']}")
        st.markdown(f"**Empresa:** {user['empresa']}")
        st.markdown(f"**Créditos:** {credits}")
        if user.get('is_admin', False):
            st.markdown("**Status:** Administrador")
        
        st.markdown("---")
        st.markdown("### Estrutura da Planilha")
        st.markdown("""
        **Colunas obrigatórias:**
        - Nome, Setor, Função
        - Data de Admissão
        - Empresa, Unidade  
        - Descrição de Atividades
        """)
        
        # Botão para baixar planilha exemplo
        sample_df = create_sample_data()
        sample_buffer = BytesIO()
        sample_df.to_excel(sample_buffer, index=False)
        sample_buffer.seek(0)
        
        st.download_button(
            "Baixar Planilha Exemplo",
            data=sample_buffer.getvalue(),
            file_name="modelo_funcionarios.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    # Seção de upload de arquivos
    st.markdown("## Upload de Arquivos")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="upload-area">
            <h4>Planilha de Funcionários</h4>
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
            <h4>Modelo de OS (Opcional)</h4>
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
            is_valid, message = validate_excel_structure(df)
            
            if not is_valid:
                st.markdown(f'<div class="error-msg">{message}</div>', unsafe_allow_html=True)
                return
            
            st.markdown(f'<div class="success-msg">Planilha carregada: {len(df)} funcionários encontrados</div>', unsafe_allow_html=True)
            
            # Estatísticas da planilha
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{len(df)}</h3>
                    <p>Funcionários</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{df['Setor'].nunique()}</h3>
                    <p>Setores</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{df['Função'].nunique()}</h3>
                    <p>Funções</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                empresa_count = df['Empresa'].nunique() if 'Empresa' in df.columns else 1
                st.markdown(f"""
                <div class="metric-card">
                    <h3>{empresa_count}</h3>
                    <p>Empresas</p>
                </div>
                """, unsafe_allow_html=True)
            
            # SELEÇÃO DE FUNCIONÁRIOS COM FILTROS MÚLTIPLOS
            st.markdown("## Seleção de Funcionários")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # FILTRO POR SETORES (MÚLTIPLA SELEÇÃO)
                todos_setores = sorted(df['Setor'].dropna().unique().tolist())
                setores_selecionados = st.multiselect(
                    "Filtrar por Setores:",
                    todos_setores,
                    help="Selecione um ou mais setores. Deixe vazio para incluir todos."
                )
            
            with col2:
                # FILTRO POR FUNÇÕES (MÚLTIPLA SELEÇÃO)
                todas_funcoes = sorted(df['Função'].dropna().unique().tolist())
                funcoes_selecionadas = st.multiselect(
                    "Filtrar por Funções:",
                    todas_funcoes,
                    help="Selecione uma ou mais funções. Deixe vazio para incluir todas."
                )
            
            # APLICAR FILTROS
            df_filtrado = df.copy()
            
            # Filtrar por setores se selecionados
            if setores_selecionados:
                df_filtrado = df_filtrado[df_filtrado['Setor'].isin(setores_selecionados)]
            
            # Filtrar por funções se selecionadas
            if funcoes_selecionadas:
                df_filtrado = df_filtrado[df_filtrado['Função'].isin(funcoes_selecionadas)]
            
            # Mostrar informações dos filtros aplicados
            if setores_selecionados or funcoes_selecionadas:
                filtros_aplicados = []
                if setores_selecionados:
                    filtros_aplicados.append(f"Setores: {', '.join(setores_selecionados)}")
                if funcoes_selecionadas:
                    filtros_aplicados.append(f"Funções: {', '.join(funcoes_selecionadas)}")
                
                st.markdown(f'<div class="info-msg"><strong>Filtros aplicados:</strong><br>• {("<br>• ".join(filtros_aplicados))}<br><strong>Funcionários encontrados:</strong> {len(df_filtrado)}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="info-msg"><strong>Nenhum filtro aplicado</strong> - Mostrando todos os {len(df_filtrado)} funcionários</div>', unsafe_allow_html=True)
            
            # MODO DE SELEÇÃO
            modo_selecao = st.radio(
                "Modo de Seleção:",
                ["Funcionário Individual", "Múltiplos Funcionários", "Todos os Funcionários Filtrados"],
                help="Escolha como selecionar os funcionários para gerar as OS"
            )
            
            # LÓGICA DE SELEÇÃO
            funcionarios_selecionados = []
            
            if modo_selecao == "Funcionário Individual":
                if len(df_filtrado) > 0:
                    funcionario_individual = st.selectbox(
                        "Selecione o funcionário:",
                        [''] + df_filtrado['Nome'].tolist(),
                        help="Escolha um funcionário específico da lista filtrada"
                    )
                    if funcionario_individual:
                        funcionarios_selecionados = [funcionario_individual]
                else:
                    st.warning("Nenhum funcionário encontrado com os filtros aplicados.")
            
            elif modo_selecao == "Múltiplos Funcionários":
                if len(df_filtrado) > 0:
                    funcionarios_selecionados = st.multiselect(
                        "Selecione múltiplos funcionários:",
                        df_filtrado['Nome'].tolist(),
                        help="Escolha vários funcionários da lista filtrada"
                    )
                else:
                    st.warning("Nenhum funcionário encontrado com os filtros aplicados.")
            
            else:  # Todos os funcionários filtrados
                funcionarios_selecionados = df_filtrado['Nome'].tolist()
                if funcionarios_selecionados:
                    st.info(f"Serão geradas OS para todos os {len(funcionarios_selecionados)} funcionários filtrados.")
            
            # Configuração de riscos se há funcionários selecionados
            if funcionarios_selecionados:
                st.markdown(f'<div class="success-msg">{len(funcionarios_selecionados)} funcionário(s) selecionado(s)</div>', unsafe_allow_html=True)
                
                st.markdown("## Configuração de Riscos Ocupacionais")
                
                st.markdown("""
                <div class="warning-msg">
                    <strong>Importante:</strong><br>
                    • Se nenhum risco for selecionado, o sistema preencherá automaticamente com <strong>"Ausência de Fator de Risco"</strong><br>
                    • Isso garante conformidade com as normas de segurança do trabalho
                </div>
                """, unsafe_allow_html=True)
                
                # Inicializar dados de sessão
                if 'agentes_risco' not in st.session_state:
                    st.session_state.agentes_risco = {categoria: [] for categoria in CATEGORIAS_RISCO.keys()}
                if 'epis_selecionados' not in st.session_state:
                    st.session_state.epis_selecionados = []
                if 'medidas_preventivas' not in st.session_state:
                    st.session_state.medidas_preventivas = []
                
                # Configurar riscos por categoria
                st.markdown("### Agentes de Riscos por Categoria")
                
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
                            if st.button(f"Adicionar", key=f"add_{categoria_key}"):
                                if agente_selecionado != 'Selecione...':
                                    novo_risco = {
                                        'agente': agente_selecionado,
                                        'intensidade': intensidade,
                                        'unidade': unidade
                                    }
                                    st.session_state.agentes_risco[categoria_key].append(novo_risco)
                                    st.markdown('<div class="success-msg">Risco adicionado</div>', unsafe_allow_html=True)
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
                                    if st.button("Remover", key=f"remove_{categoria_key}_{idx}"):
                                        st.session_state.agentes_risco[categoria_key].pop(idx)
                                        st.rerun()
                
                # EPIs e Medidas Preventivas
                st.markdown("### EPIs e Medidas Preventivas")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**EPIs Obrigatórios:**")
                    novo_epi = st.text_input("Adicionar EPI:", placeholder="Ex: Capacete, óculos...")
                    if st.button("Adicionar EPI") and novo_epi:
                        st.session_state.epis_selecionados.append(novo_epi)
                        st.rerun()
                    
                    for idx, epi in enumerate(st.session_state.epis_selecionados):
                        col_epi1, col_epi2 = st.columns([4, 1])
                        with col_epi1:
                            st.write(f"• {epi}")
                        with col_epi2:
                            if st.button("X", key=f"remove_epi_{idx}"):
                                st.session_state.epis_selecionados.pop(idx)
                                st.rerun()
                
                with col2:
                    st.markdown("**Medidas Preventivas:**")
                    nova_medida = st.text_area("Adicionar Medida:", placeholder="Ex: Treinamentos, pausas...", height=100)
                    if st.button("Adicionar Medida") and nova_medida:
                        st.session_state.medidas_preventivas.append(nova_medida)
                        st.rerun()
                    
                    for idx, medida in enumerate(st.session_state.medidas_preventivas):
                        col_med1, col_med2 = st.columns([4, 1])
                        with col_med1:
                            medida_resumida = medida[:100] + "..." if len(medida) > 100 else medida
                            st.write(f"• {medida_resumida}")
                        with col_med2:
                            if st.button("X", key=f"remove_med_{idx}"):
                                st.session_state.medidas_preventivas.pop(idx)
                                st.rerun()
                
                # Observações
                observacoes = st.text_area(
                    "Observações Complementares:",
                    placeholder="Informações específicas do setor, procedimentos especiais, etc.",
                    height=80
                )
                
                # Botão para gerar OS
                st.markdown("## Gerar Ordens de Serviço")
                
                creditos_necessarios = len(funcionarios_selecionados)
                tem_creditos_suficientes = check_sufficient_credits(user['id'], creditos_necessarios)
                
                if tem_creditos_suficientes:
                    if user.get('is_admin', False):
                        button_text = f"Gerar {len(funcionarios_selecionados)} OS (Gratuito - Admin)"
                    else:
                        button_text = f"Gerar {len(funcionarios_selecionados)} OS ({creditos_necessarios} créditos)"
                    
                    if st.button(button_text, type="primary"):
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        documentos_gerados = []
                        
                        # Processar cada funcionário selecionado
                        for idx, nome_funcionario in enumerate(funcionarios_selecionados):
                            status_text.text(f"Gerando OS para: {nome_funcionario}")
                            
                            # Buscar dados do funcionário
                            dados_funcionario = df_filtrado[df_filtrado['Nome'] == nome_funcionario].iloc[0].to_dict()
                            
                            # Gerar documento
                            doc = gerar_documento_os(
                                dados_funcionario=dados_funcionario,
                                agentes_risco=st.session_state.agentes_risco,
                                epis=st.session_state.epis_selecionados,
                                medidas_preventivas=st.session_state.medidas_preventivas,
                                observacoes=observacoes,
                                uploaded_template=uploaded_template
                            )
                            
                            if doc:
                                # Salvar documento em buffer
                                buffer = BytesIO()
                                doc.save(buffer)
                                buffer.seek(0)
                                
                                documentos_gerados.append({
                                    'nome': nome_funcionario.replace(' ', '_').replace('/', '_'),
                                    'buffer': buffer
                                })
                            
                            # Atualizar progresso
                            progress_bar.progress((idx + 1) / len(funcionarios_selecionados))
                            time.sleep(0.3)
                        
                        # Debitar créditos (só se não for admin)
                        if not user.get('is_admin', False):
                            debit_credits(user['id'], creditos_necessarios)
                        
                        status_text.text("Geração concluída!")
                        
                        # Disponibilizar downloads
                        if documentos_gerados:
                            if len(documentos_gerados) == 1:
                                st.markdown('<div class="success-msg">Ordem de Serviço gerada com sucesso</div>', unsafe_allow_html=True)
                                st.download_button(
                                    label="Download da Ordem de Serviço",
                                    data=documentos_gerados[0]['buffer'].getvalue(),
                                    file_name=f"OS_{documentos_gerados[0]['nome']}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    use_container_width=True
                                )
                            else:
                                st.markdown(f'<div class="success-msg">{len(documentos_gerados)} Ordens de Serviço geradas com sucesso</div>', unsafe_allow_html=True)
                                
                                # Criar arquivo ZIP
                                zip_buffer = BytesIO()
                                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                    for doc_info in documentos_gerados:
                                        zip_file.writestr(
                                            f"OS_{doc_info['nome']}.docx",
                                            doc_info['buffer'].getvalue()
                                        )
                                
                                zip_buffer.seek(0)
                                
                                st.download_button(
                                    label=f"Download de {len(documentos_gerados)} Ordens de Serviço (ZIP)",
                                    data=zip_buffer.getvalue(),
                                    file_name=f"Lote_OS_{datetime.date.today().strftime('%d%m%Y')}.zip",
                                    mime="application/zip",
                                    use_container_width=True
                                )
                                
                            # Mostrar créditos restantes (apenas se não for admin)
                            if not user.get('is_admin', False):
                                creditos_restantes = get_user_credits(user['id'])
                                st.markdown(f'<div class="info-msg">{creditos_necessarios} créditos foram debitados. Créditos restantes: {creditos_restantes}</div>', unsafe_allow_html=True)
                            else:
                                st.markdown(f'<div class="info-msg">Geração realizada sem custo (conta administrador)</div>', unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="error-msg">Erro: Nenhum documento foi gerado. Verifique as configurações</div>', unsafe_allow_html=True)
                        
                        time.sleep(2)
                        st.rerun()
                else:
                    creditos_usuario = get_user_credits(user['id'])
                    st.markdown(f'<div class="error-msg">Créditos insuficientes. Você precisa de {creditos_necessarios} créditos, mas possui apenas {creditos_usuario}</div>', unsafe_allow_html=True)
        
        except Exception as e:
            st.markdown(f'<div class="error-msg">Erro ao processar planilha: {str(e)}</div>', unsafe_allow_html=True)
    
    else:
        # Instruções iniciais
        total_riscos = sum(len(riscos) for riscos in AGENTES_POR_CATEGORIA.values())
        st.markdown(f"""
        <div class="info-msg">
            <h4>Como usar o sistema:</h4>
            <ol>
                <li><strong>Faça upload</strong> da planilha Excel com dados dos funcionários</li>
                <li><strong>Faça upload</strong> do template Word (opcional) - use seu modelo personalizado</li>
                <li><strong>Selecione setores</strong> (múltipla seleção disponível)</li>
                <li><strong>Selecione funções</strong> (múltipla seleção disponível)</li>
                <li><strong>Escolha</strong> o modo de seleção de funcionários</li>
                <li><strong>Configure</strong> os riscos ocupacionais específicos</li>
                <li><strong>Adicione</strong> EPIs e medidas preventivas</li>
                <li><strong>Gere</strong> as Ordens de Serviço conforme NR-01</li>
            </ol>
            
            <p><strong>Funcionalidades implementadas:</strong></p>
            <ul>
                <li>Filtro múltiplo por setores - Selecione vários setores simultaneamente</li>
                <li>Filtro múltiplo por funções - Selecione várias funções simultaneamente</li>
                <li>Template personalizado - Upload do seu modelo Word</li>
                <li>Preenchimento automático - "Ausência de Fator de Risco" quando necessário</li>
                <li>Validação de email - Apenas emails de provedores válidos</li>
                <li>Tema minimalista - Interface clean e profissional</li>
                <li>{total_riscos} opções de riscos organizados em 5 categorias</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

# --- LÓGICA PRINCIPAL ---
def main():
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    
    if 'user' not in st.session_state:
        st.session_state.user = None
    
    if st.session_state.authenticated and st.session_state.user:
        show_main_app(st.session_state.user)
    else:
        show_login_page()

if __name__ == "__main__":
    main()
