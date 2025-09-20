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

# --- CSS PERSONALIZADO LIMPO ---
st.markdown("""
<style>
    /* TEMA DARK MODERNO */
    .stApp {
        background: #1a1a2e;
        color: #ffffff;
    }
    
    .main {
        background: #1a1a2e;
        color: #ffffff;
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
    
    /* TÍTULO PRINCIPAL */
    .title-header {
        text-align: center;
        color: #4CAF50;
        font-size: 3rem;
        font-weight: bold;
        margin: 2rem 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
    }
    
    /* FORMULÁRIO DE LOGIN */
    .login-container {
        max-width: 500px;
        margin: 2rem auto;
        padding: 2rem;
        background: #16213e;
        border-radius: 15px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.3);
        border: 1px solid #0f3460;
    }
    
    .login-title {
        text-align: center;
        color: #4CAF50;
        font-size: 2rem;
        margin-bottom: 2rem;
    }
    
    /* BOTÕES */
    .stButton > button {
        background: linear-gradient(135deg, #4CAF50, #45a049);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.75rem 1.5rem;
        font-size: 1.1rem;
        font-weight: 600;
        width: 100%;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #45a049, #3d8b40);
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(76, 175, 80, 0.4);
    }
    
    /* INPUTS */
    .stTextInput > div > div > input {
        background: #0f1419;
        color: white;
        border: 2px solid #0f3460;
        border-radius: 8px;
        padding: 0.75rem;
    }
    
    .stTextInput > div > div > input:focus {
        border-color: #4CAF50;
        box-shadow: 0 0 0 2px rgba(76, 175, 80, 0.2);
    }
    
    /* MENSAGENS */
    .success-msg {
        background: rgba(76, 175, 80, 0.1);
        border: 1px solid #4CAF50;
        border-radius: 8px;
        padding: 1rem;
        color: #4CAF50;
        margin: 1rem 0;
    }
    
    .error-msg {
        background: rgba(244, 67, 54, 0.1);
        border: 1px solid #f44336;
        border-radius: 8px;
        padding: 1rem;
        color: #f44336;
        margin: 1rem 0;
    }
    
    .warning-msg {
        background: rgba(255, 193, 7, 0.1);
        border: 1px solid #ffc107;
        border-radius: 8px;
        padding: 1rem;
        color: #ffc107;
        margin: 1rem 0;
    }
    
    .info-msg {
        background: rgba(33, 150, 243, 0.1);
        border: 1px solid #2196F3;
        border-radius: 8px;
        padding: 1rem;
        color: #2196F3;
        margin: 1rem 0;
    }
    
    /* TABS */
    .stTabs [data-baseweb="tab-list"] {
        background: #16213e;
        border-radius: 10px 10px 0 0;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        color: #ffffff;
    }
    
    .stTabs [aria-selected="true"] {
        background: #4CAF50 !important;
        color: white !important;
    }
    
    /* SIDEBAR */
    .css-1d391kg {
        background: #16213e;
    }
    
    /* UPLOAD AREAS */
    .upload-area {
        background: #16213e;
        border: 2px dashed #4CAF50;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        margin: 1rem 0;
        color: white;
    }
    
    /* CARDS */
    .metric-card {
        background: #16213e;
        padding: 1.5rem;
        border-radius: 10px;
        text-align: center;
        border: 1px solid #0f3460;
        margin: 0.5rem 0;
    }
    
    .metric-card h3 {
        color: #4CAF50;
        font-size: 2rem;
        margin-bottom: 0.5rem;
    }
    
    .metric-card p {
        color: #ffffff;
        margin: 0;
    }
    
    /* EXPANSORES */
    .streamlit-expanderHeader {
        background: #16213e !important;
        border: 1px solid #0f3460 !important;
        border-radius: 8px !important;
        color: white !important;
    }
    
    .streamlit-expanderContent {
        background: #1a1a2e !important;
        border: 1px solid #0f3460 !important;
        color: white !important;
    }
    
    /* MULTISELECT CUSTOMIZADO */
    .stMultiSelect [data-baseweb="select"] {
        background: #16213e;
        border: 2px solid #0f3460;
        border-radius: 8px;
    }
    
    .stMultiSelect [data-baseweb="select"]:hover {
        border-color: #4CAF50;
    }
    
    /* SELECTBOX CUSTOMIZADO */
    .stSelectbox [data-baseweb="select"] {
        background: #16213e;
        border: 2px solid #0f3460;
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# --- SISTEMA DE AUTENTICAÇÃO APRIMORADO ---
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def initialize_users():
    if 'users_db' not in st.session_state:
        st.session_state.users_db = {
            'admin@teste.com': {
                'id': 1,
                'password': hash_password('admin123'),
                'nome': 'Administrador',
                'empresa': 'Empresa Teste',
                'email': 'admin@teste.com',
                'credits': 1000,
                'is_admin': False
            },
            'robsonmbf@hotmail.com': {
                'id': 2,
                'password': hash_password('123456'),
                'nome': 'Robson',
                'empresa': 'Minha Empresa',
                'email': 'robsonmbf@hotmail.com',
                'credits': 999999,  # Créditos ilimitados
                'is_admin': True    # Admin não consome créditos
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
                return "∞"  # Mostrar infinito para admin
            return user['credits']
    return 0

def debit_credits(user_id, amount):
    initialize_users()
    for user in st.session_state.users_db.values():
        if user['id'] == user_id:
            if user.get('is_admin', False):
                return True  # Admin não consome créditos
            user['credits'] = max(0, user['credits'] - amount)
            return True
    return False

def check_sufficient_credits(user_id, amount):
    initialize_users()
    for user in st.session_state.users_db.values():
        if user['id'] == user_id:
            if user.get('is_admin', False):
                return True  # Admin sempre tem créditos suficientes
            return user['credits'] >= amount
    return False

# --- FUNÇÕES AUXILIARES APRIMORADAS ---
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

def check_duplicate_functions_across_sectors(df):
    """Verifica se há funções duplicadas em setores diferentes"""
    function_sector_map = {}
    duplicates = []
    
    for _, row in df.iterrows():
        funcao = row['Função']
        setor = row['Setor']
        
        if funcao in function_sector_map:
            if setor not in function_sector_map[funcao]:
                function_sector_map[funcao].append(setor)
                if len(function_sector_map[funcao]) == 2:  # Primeira duplicata encontrada
                    duplicates.append({
                        'funcao': funcao,
                        'setores': function_sector_map[funcao].copy()
                    })
                elif len(function_sector_map[funcao]) > 2:  # Mais setores para a mesma função
                    # Atualizar a lista de setores para essa função
                    for dup in duplicates:
                        if dup['funcao'] == funcao:
                            dup['setores'] = function_sector_map[funcao].copy()
        else:
            function_sector_map[funcao] = [setor]
    
    return duplicates

def gerar_documento_os(dados_funcionario, agentes_risco, epis, medidas_preventivas, observacoes, template_doc=None):
    try:
        if template_doc:
            doc = template_doc
        else:
            doc = Document()
        
        if not template_doc:
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
        
        if agentes_risco:
            doc.add_heading('AGENTES DE RISCOS OCUPACIONAIS', level=1)
            
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
        
        if epis:
            doc.add_heading('EQUIPAMENTOS DE PROTEÇÃO INDIVIDUAL (EPIs)', level=1)
            for epi in epis:
                doc.add_paragraph(f"• {epi}", style='List Bullet')
        
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

# --- FUNÇÃO DE LOGIN ATUALIZADA ---
def show_login_page():
    st.markdown('<div class="title-header">🔐 Gerador de Ordens de Serviço (OS)</div>', unsafe_allow_html=True)
    
    total_riscos = sum(len(riscos) for riscos in AGENTES_POR_CATEGORIA.values())
    
    st.markdown(f"""
    <div class="info-msg">
        <strong>🆕 SISTEMA ATUALIZADO - Base Expandida de Riscos!</strong><br><br>
        ✨ <strong>{total_riscos} opções</strong> de riscos ocupacionais organizados em 5 categorias<br>
        🏃 <strong>Ergonômicos:</strong> {len(RISCOS_ERGONOMICO)} opções específicas<br>
        ⚠️ <strong>Acidentes:</strong> {len(RISCOS_ACIDENTE)} opções detalhadas<br>
        🔥 <strong>Físicos:</strong> {len(RISCOS_FISICO)} opções ampliadas<br>
        ⚗️ <strong>Químicos:</strong> {len(RISCOS_QUIMICO)} opções específicas<br>
        🦠 <strong>Biológicos:</strong> {len(RISCOS_BIOLOGICO)} incluindo COVID-19<br><br>
        📄 Sistema profissional conforme NR-01 com interface dark!
    </div>
    """, unsafe_allow_html=True)
    
    login_tab, register_tab = st.tabs(["🔑 Login", "👤 Criar Conta"])
    
    with login_tab:
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        st.markdown('<div class="login-title">🔑 Faça seu Login</div>', unsafe_allow_html=True)
        
        with st.form("login_form"):
            # NÃO mostrar contas de teste publicamente
            email = st.text_input("📧 Email:", placeholder="seu@email.com")
            password = st.text_input("🔒 Senha:", type="password", placeholder="Sua senha")
            
            login_button = st.form_submit_button("🚀 Entrar")
            
            if login_button:
                if email and password:
                    user = authenticate_user(email, password)
                    if user:
                        st.session_state.user = user
                        st.session_state.authenticated = True
                        st.markdown('<div class="success-msg">✅ Login realizado com sucesso!</div>', unsafe_allow_html=True)
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.markdown('<div class="error-msg">❌ Email ou senha incorretos.</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="error-msg">⚠️ Por favor, preencha todos os campos.</div>', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with register_tab:
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        st.markdown('<div class="login-title">👤 Criar Nova Conta</div>', unsafe_allow_html=True)
        
        with st.form("register_form"):
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
            
            register_button = st.form_submit_button("✨ Criar Conta")
            
            if register_button:
                if nome and empresa and email and password and password_confirm:
                    if password == password_confirm:
                        if len(password) >= 6:
                            user_id = register_user(email, password, nome, empresa)
                            if user_id:
                                st.markdown('<div class="success-msg">✅ Conta criada com sucesso! Faça login para continuar.</div>', unsafe_allow_html=True)
                            else:
                                st.markdown('<div class="error-msg">❌ Erro ao criar conta. Email já pode estar em uso.</div>', unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="error-msg">❌ A senha deve ter pelo menos 6 caracteres.</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="error-msg">❌ As senhas não coincidem.</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="error-msg">⚠️ Por favor, preencha todos os campos.</div>', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

def show_main_app(user):
    # Header do usuário
    col1, col2, col3 = st.columns([3, 1, 1])
    
    with col1:
        st.markdown(f"# 📄 Gerador de OS - Bem-vindo, **{user['nome']}**!")
    
    with col2:
        credits = get_user_credits(user['id'])
        st.metric("💳 Créditos", credits)
    
    with col3:
        if st.button("🚪 Logout"):
            st.session_state.authenticated = False
            st.session_state.user = None
            st.rerun()
    
    st.markdown(f"🏢 **Empresa:** {user['empresa']}")
    
    # Mostrar status de admin se for o caso
    if user.get('is_admin', False):
        st.markdown("""
        <div class="warning-msg">
            <strong>👑 CONTA ADMINISTRADOR</strong><br>
            • Créditos ilimitados<br>
            • Não há cobrança de créditos<br>
            • Acesso completo ao sistema
        </div>
        """, unsafe_allow_html=True)
    
    total_riscos = sum(len(riscos) for riscos in AGENTES_POR_CATEGORIA.values())
    st.markdown(f"""
    <div class="info-msg">
        <strong>🚀 SISTEMA ATUALIZADO - Nova Base de Riscos!</strong><br><br>
        📊 <strong>Total:</strong> {total_riscos} opções de riscos ocupacionais organizados em 5 categorias<br>
        🏃 <strong>Ergonômicos:</strong> {len(RISCOS_ERGONOMICO)} riscos específicos<br>
        ⚠️ <strong>Acidentes:</strong> {len(RISCOS_ACIDENTE)} riscos detalhados<br>
        🔥 <strong>Físicos:</strong> {len(RISCOS_FISICO)} riscos ampliados<br>
        ⚗️ <strong>Químicos:</strong> {len(RISCOS_QUIMICO)} opções específicas<br>
        🦠 <strong>Biológicos:</strong> {len(RISCOS_BIOLOGICO)} incluindo COVID-19
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar com informações
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
        st.markdown(f"**Créditos:** {credits}")
        if user.get('is_admin', False):
            st.markdown("**Status:** 👑 Administrador")
        
        st.markdown("---")
        st.markdown("### 📋 Estrutura da Planilha")
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
            "📥 Baixar Planilha Exemplo",
            data=sample_buffer.getvalue(),
            file_name="modelo_funcionarios.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
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
            is_valid, message = validate_excel_structure(df)
            
            if not is_valid:
                st.markdown(f'<div class="error-msg">❌ {message}</div>', unsafe_allow_html=True)
                return
            
            st.markdown(f'<div class="success-msg">✅ Planilha carregada: {len(df)} funcionários encontrados</div>', unsafe_allow_html=True)
            
            # Verificar funções duplicadas em setores diferentes
            duplicates = check_duplicate_functions_across_sectors(df)
            if duplicates:
                st.markdown('<div class="warning-msg"><strong>⚠️ ATENÇÃO - Funções Duplicadas Encontradas:</strong><br><br>', unsafe_allow_html=True)
                for dup in duplicates:
                    setores_text = ", ".join(dup['setores'])
                    st.markdown(f"• **{dup['funcao']}** encontrada nos setores: **{setores_text}**<br>", unsafe_allow_html=True)
                st.markdown('<br>Recomenda-se revisar se os riscos são os mesmos para esta função em setores diferentes.</div>', unsafe_allow_html=True)
            
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
            
            # Seleção de funcionários MELHORADA
            st.markdown("## 👥 Seleção de Funcionários")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # SELEÇÃO MÚLTIPLA DE SETORES
                todos_setores = sorted(df['Setor'].dropna().unique().tolist())
                setores_selecionados = st.multiselect(
                    "🏢 Selecione os Setores:",
                    todos_setores,
                    help="Selecione um ou mais setores. Se nenhum for selecionado, todos serão incluídos."
                )
                
                # Se nenhum setor for selecionado, incluir todos
                if not setores_selecionados:
                    df_filtrado = df
                    st.info("📋 Todos os setores estão incluídos (nenhum selecionado)")
                else:
                    df_filtrado = df[df['Setor'].isin(setores_selecionados)]
                    setores_text = ", ".join(setores_selecionados)
                    st.success(f"📋 Filtrando por: {setores_text}")
            
            with col2:
                modo_selecao = st.radio(
                    "📋 Modo de Seleção:",
                    ["Funcionário Individual", "Múltiplos Funcionários", "Todos dos Setores Selecionados"]
                )
            
            # Lógica de seleção APRIMORADA
            funcionarios_selecionados = []
            
            if modo_selecao == "Funcionário Individual":
                funcionario_individual = st.selectbox(
                    "👤 Selecione o funcionário:",
                    [''] + df_filtrado['Nome'].tolist(),
                    help="Escolha um funcionário específico da lista filtrada"
                )
                if funcionario_individual:
                    funcionarios_selecionados = [funcionario_individual]
            
            elif modo_selecao == "Múltiplos Funcionários":
                funcionarios_selecionados = st.multiselect(
                    "👥 Selecione múltiplos funcionários:",
                    df_filtrado['Nome'].tolist(),
                    help="Escolha vários funcionários mantendo Ctrl pressionado"
                )
            
            else:  # Todos dos setores selecionados
                funcionarios_selecionados = df_filtrado['Nome'].tolist()
                if funcionarios_selecionados:
                    if setores_selecionados:
                        st.info(f"📝 Serão geradas OS para todos os {len(funcionarios_selecionados)} funcionários dos setores selecionados.")
                    else:
                        st.info(f"📝 Serão geradas OS para todos os {len(funcionarios_selecionados)} funcionários de todos os setores.")
            
            # Configuração de riscos se há funcionários selecionados
            if funcionarios_selecionados:
                st.markdown(f'<div class="success-msg">✅ {len(funcionarios_selecionados)} funcionário(s) selecionado(s)</div>', unsafe_allow_html=True)
                
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
                                    st.markdown('<div class="success-msg">✅ Risco adicionado!</div>', unsafe_allow_html=True)
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
                
                creditos_necessarios = len(funcionarios_selecionados)
                tem_creditos_suficientes = check_sufficient_credits(user['id'], creditos_necessarios)
                
                if tem_creditos_suficientes:
                    if user.get('is_admin', False):
                        button_text = f"📄 GERAR {len(funcionarios_selecionados)} OS (GRATUITO - ADMIN)"
                    else:
                        button_text = f"📄 GERAR {len(funcionarios_selecionados)} OS ({creditos_necessarios} créditos)"
                    
                    if st.button(button_text, type="primary"):
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        documentos_gerados = []
                        
                        # Processar cada funcionário selecionado
                        for idx, nome_funcionario in enumerate(funcionarios_selecionados):
                            status_text.text(f"🔄 Gerando OS para: {nome_funcionario}")
                            
                            # Buscar dados do funcionário
                            dados_funcionario = df_filtrado[df_filtrado['Nome'] == nome_funcionario].iloc[0].to_dict()
                            
                            # Gerar documento
                            doc = gerar_documento_os(
                                dados_funcionario=dados_funcionario,
                                agentes_risco=st.session_state.agentes_risco,
                                epis=st.session_state.epis_selecionados,
                                medidas_preventivas=st.session_state.medidas_preventivas,
                                observacoes=observacoes,
                                template_doc=uploaded_template
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
                        
                        status_text.text("✅ Geração concluída!")
                        
                        # Disponibilizar downloads
                        if documentos_gerados:
                            if len(documentos_gerados) == 1:
                                st.markdown('<div class="success-msg">✅ Ordem de Serviço gerada com sucesso!</div>', unsafe_allow_html=True)
                                st.download_button(
                                    label="📥 Download da Ordem de Serviço",
                                    data=documentos_gerados[0]['buffer'].getvalue(),
                                    file_name=f"OS_{documentos_gerados[0]['nome']}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    use_container_width=True
                                )
                            else:
                                st.markdown(f'<div class="success-msg">✅ {len(documentos_gerados)} Ordens de Serviço geradas com sucesso!</div>', unsafe_allow_html=True)
                                
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
                                    label=f"📥 Download de {len(documentos_gerados)} Ordens de Serviço (ZIP)",
                                    data=zip_buffer.getvalue(),
                                    file_name=f"Lote_OS_{datetime.date.today().strftime('%d%m%Y')}.zip",
                                    mime="application/zip",
                                    use_container_width=True
                                )
                                
                            # Mostrar créditos restantes (apenas se não for admin)
                            if not user.get('is_admin', False):
                                creditos_restantes = get_user_credits(user['id'])
                                st.markdown(f'<div class="info-msg">💳 {creditos_necessarios} créditos foram debitados. Créditos restantes: {creditos_restantes}</div>', unsafe_allow_html=True)
                            else:
                                st.markdown(f'<div class="info-msg">👑 Geração realizada sem custo (conta administrador)</div>', unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="error-msg">❌ Erro: Nenhum documento foi gerado. Verifique as configurações.</div>', unsafe_allow_html=True)
                        
                        time.sleep(2)
                        st.rerun()
                else:
                    creditos_usuario = get_user_credits(user['id'])
                    st.markdown(f'<div class="error-msg">⚠️ Créditos insuficientes. Você precisa de {creditos_necessarios} créditos, mas possui apenas {creditos_usuario}.</div>', unsafe_allow_html=True)
        
        except Exception as e:
            st.markdown(f'<div class="error-msg">❌ Erro ao processar planilha: {str(e)}</div>', unsafe_allow_html=True)
    
    else:
        # Instruções iniciais
        total_riscos = sum(len(riscos) for riscos in AGENTES_POR_CATEGORIA.values())
        st.markdown(f"""
        <div class="info-msg">
            <h4>🎯 Como usar o sistema:</h4>
            <ol>
                <li><strong>📤 Faça upload</strong> da planilha Excel com dados dos funcionários</li>
                <li><strong>🏢 Selecione</strong> um ou mais setores (ou deixe vazio para todos)</li>
                <li><strong>👥 Escolha</strong> o modo de seleção de funcionários</li>
                <li><strong>⚠️ Configure</strong> os riscos ocupacionais específicos</li>
                <li><strong>🥽 Adicione</strong> EPIs e medidas preventivas</li>
                <li><strong>🚀 Gere</strong> as Ordens de Serviço conforme NR-01</li>
            </ol>
            
            <p><strong>🆕 Novidades desta versão:</strong></p>
            <ul>
                <li>✅ <strong>Seleção múltipla de setores</strong> - Escolha vários setores simultaneamente</li>
                <li>✅ <strong>Detecção de funções duplicadas</strong> - Sistema alerta sobre mesma função em setores diferentes</li>
                <li>✅ <strong>{total_riscos} opções de riscos</strong> organizados em 5 categorias</li>
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
