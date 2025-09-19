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
import datetime

# Adicionar o diretório atual ao path para importar módulos locais
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Tentar importar módulos do banco de dados, se não conseguir, usar sistema simplificado
try:
    from database.models import DatabaseManager
    from database.auth import AuthManager
    from database.user_data import UserDataManager
    USE_DATABASE = True
except ImportError:
    USE_DATABASE = False
    st.warning("⚠️ Sistema funcionando no modo simplificado (sem banco de dados)")

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
if USE_DATABASE:
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
else:
    db_manager, auth_manager, user_data_manager = None, None, None

# --- CSS PERSONALIZADO CORRIGIDO ---
st.markdown("""
<style>
    /* CORRIGIR BARRA BRANCA DO TOPO */
    .stApp > header {
        background-color: transparent !important;
        height: 0px !important;
        position: fixed !important;
        top: -100px !important;
    }
    
    /* REMOVER PADDING SUPERIOR */
    .main .block-container {
        padding-top: 1rem !important;
        padding-bottom: 0rem !important;
        max-width: 100% !important;
    }
    
    /* FUNDO PRINCIPAL */
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
    }
    
    .main {
        background: transparent !important;
        padding-top: 0rem !important;
    }
    
    /* TÍTULOS E CABEÇALHOS */
    .login-header {
        text-align: center;
        color: white !important;
        font-size: 2.8rem !important;
        font-weight: bold !important;
        margin-bottom: 2rem !important;
        text-shadow: 3px 3px 6px rgba(0,0,0,0.3) !important;
        padding: 2rem 0 !important;
    }
    
    /* FORMULÁRIOS DE LOGIN */
    .login-form {
        max-width: 450px !important;
        margin: 2rem auto !important;
        padding: 2.5rem !important;
        background: rgba(255, 255, 255, 0.95) !important;
        border-radius: 15px !important;
        box-shadow: 0 8px 32px rgba(0,0,0,0.2) !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        backdrop-filter: blur(10px) !important;
    }
    
    /* MENSAGENS DE STATUS */
    .success-message {
        background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%) !important;
        padding: 1.2rem !important;
        border-radius: 10px !important;
        border-left: 4px solid #22c55e !important;
        margin: 1rem 0 !important;
        color: #166534 !important;
        box-shadow: 0 2px 8px rgba(34, 197, 94, 0.2) !important;
    }
    
    .error-message {
        background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%) !important;
        padding: 1.2rem !important;
        border-radius: 10px !important;
        border-left: 4px solid #ef4444 !important;
        margin: 1rem 0 !important;
        color: #991b1b !important;
        box-shadow: 0 2px 8px rgba(239, 68, 68, 0.2) !important;
    }
    
    .info-message, .new-features {
        background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%) !important;
        padding: 1.5rem !important;
        border-radius: 10px !important;
        border-left: 4px solid #3b82f6 !important;
        margin: 1rem 0 !important;
        color: #1e40af !important;
        box-shadow: 0 2px 8px rgba(59, 130, 246, 0.2) !important;
    }
    
    /* BOTÕES */
    .stButton > button {
        width: 100% !important;
        border-radius: 10px !important;
        border: none !important;
        padding: 0.8rem 1.5rem !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        transition: all 0.3s ease !important;
        background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%) !important;
        color: white !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 25px rgba(59, 130, 246, 0.4) !important;
        background: linear-gradient(135deg, #2563eb 0%, #1e40af 100%) !important;
    }
    
    /* INPUTS */
    .stTextInput > div > div > input {
        border-radius: 8px !important;
        border: 2px solid #e5e7eb !important;
        padding: 0.8rem !important;
        font-size: 1rem !important;
        transition: all 0.3s ease !important;
    }
    
    .stTextInput > div > div > input:focus {
        border-color: #3b82f6 !important;
        box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1) !important;
    }
    
    /* TABS */
    .stTabs > div > div > div > div {
        background: rgba(255, 255, 255, 0.1) !important;
        border-radius: 10px 10px 0 0 !important;
        padding: 1rem !important;
    }
    
    /* ÁREAS DE UPLOAD */
    .upload-area {
        border: 2px dashed #3b82f6 !important;
        border-radius: 15px !important;
        padding: 2.5rem !important;
        text-align: center !important;
        background: rgba(255, 255, 255, 0.95) !important;
        margin: 1rem 0 !important;
        backdrop-filter: blur(10px) !important;
        transition: all 0.3s ease !important;
    }
    
    .upload-area:hover {
        background: rgba(59, 130, 246, 0.05) !important;
        border-color: #2563eb !important;
    }
    
    /* CARDS DE MÉTRICAS */
    .metric-card {
        background: rgba(255, 255, 255, 0.95) !important;
        padding: 2rem !important;
        border-radius: 15px !important;
        text-align: center !important;
        box-shadow: 0 4px 20px rgba(0,0,0,0.1) !important;
        border-top: 4px solid #3b82f6 !important;
        margin: 1rem 0 !important;
        backdrop-filter: blur(10px) !important;
        transition: all 0.3s ease !important;
    }
    
    .metric-card:hover {
        transform: translateY(-5px) !important;
        box-shadow: 0 8px 30px rgba(59, 130, 246, 0.2) !important;
    }
    
    /* SIDEBAR */
    .css-1d391kg {
        background: rgba(255, 255, 255, 0.95) !important;
        backdrop-filter: blur(10px) !important;
    }
    
    /* REMOVER ESPAÇAMENTOS DESNECESSÁRIOS */
    .css-18e3th9 {
        padding-top: 0 !important;
    }
    
    /* EXPANSORES */
    .streamlit-expanderHeader {
        background: rgba(248, 250, 252, 0.95) !important;
        border-radius: 10px !important;
        border: 1px solid #e5e7eb !important;
        backdrop-filter: blur(10px) !important;
    }
    
    .streamlit-expanderContent {
        background: rgba(255, 255, 255, 0.95) !important;
        border: 1px solid #e5e7eb !important;
        border-top: none !important;
        border-radius: 0 0 10px 10px !important;
        backdrop-filter: blur(10px) !important;
    }
</style>
""", unsafe_allow_html=True)

# --- SISTEMA DE AUTENTICAÇÃO SIMPLIFICADO ---
class SimpleAuthManager:
    def __init__(self):
        if 'users' not in st.session_state:
            st.session_state.users = {
                'admin@teste.com': {
                    'id': 1,
                    'password': 'admin123',
                    'nome': 'Administrador',
                    'empresa': 'Empresa Teste',
                    'email': 'admin@teste.com'
                }
            }
    
    def authenticate(self, email, password):
        users = st.session_state.users
        if email in users and users[email]['password'] == password:
            return users[email]
        return None
    
    def register_user(self, email, password, nome, empresa):
        if email not in st.session_state.users:
            user_id = len(st.session_state.users) + 1
            st.session_state.users[email] = {
                'id': user_id,
                'password': password,
                'nome': nome,
                'empresa': empresa,
                'email': email
            }
            return user_id
        return None

class SimpleUserDataManager:
    def get_user_credits(self, user_id):
        return 100  # Créditos ilimitados para modo simplificado
    
    def debit_credits(self, user_id, amount):
        return True  # Sempre sucesso no modo simplificado

# --- INICIALIZAR SISTEMA SIMPLIFICADO SE NECESSÁRIO ---
if not USE_DATABASE:
    auth_manager = SimpleAuthManager()
    user_data_manager = SimpleUserDataManager()

# --- FUNÇÕES AUXILIARES ---
def create_sample_data():
    """Cria dados de exemplo para demonstração"""
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
    """Valida se a planilha tem a estrutura necessária"""
    required_columns = ['Nome', 'Setor', 'Função', 'Data de Admissão', 'Empresa', 'Unidade', 'Descrição de Atividades']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        return False, f"Colunas obrigatórias faltando: {', '.join(missing_columns)}"
    
    if df.empty:
        return False, "A planilha está vazia"
    
    return True, "Estrutura válida"

def gerar_documento_os(dados_funcionario, agentes_risco, epis, medidas_preventivas, observacoes, template_doc=None):
    """Gera a Ordem de Serviço com base nos dados fornecidos"""
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

# --- FUNÇÃO DE LOGIN ---
def show_login_page():
    st.markdown("""
    <div class="login-header">
        🔐 Gerador de Ordens de Serviço (OS)
    </div>
    """, unsafe_allow_html=True)
    
    total_riscos = sum(len(riscos) for riscos in AGENTES_POR_CATEGORIA.values())
    st.markdown(f"""
    <div class="new-features">
        <strong>🆕 SISTEMA ATUALIZADO - Base Expandida de Riscos!</strong><br><br>
        ✨ <strong>{total_riscos} opções</strong> de riscos ocupacionais organizados em 5 categorias<br>
        🏃 <strong>Riscos Ergonômicos:</strong> {len(RISCOS_ERGONOMICO)} opções específicas<br>
        ⚠️ <strong>Riscos de Acidentes:</strong> {len(RISCOS_ACIDENTE)} opções detalhadas<br>
        🔥 <strong>Riscos Físicos:</strong> {len(RISCOS_FISICO)} opções ampliadas<br>
        ⚗️ <strong>Riscos Químicos:</strong> {len(RISCOS_QUIMICO)} opções específicas<br>
        🦠 <strong>Riscos Biológicos:</strong> {len(RISCOS_BIOLOGICO)} opções incluindo COVID-19<br><br>
        📄 Sistema profissional conforme NR-01 com interface otimizada!
    </div>
    """, unsafe_allow_html=True)
    
    login_tab, register_tab = st.tabs(["🔑 Login", "👤 Criar Conta"])
    
    with login_tab:
        st.markdown('<div class="login-form">', unsafe_allow_html=True)
        
        with st.form("login_form"):
            st.markdown("### 🔑 Faça seu Login")
            
            if not USE_DATABASE:
                st.info("**💡 Modo Demo:** Use `admin@teste.com` / `admin123`")
            
            email = st.text_input("📧 Email:", placeholder="seu@email.com")
            password = st.text_input("🔒 Senha:", type="password", placeholder="Sua senha")
            
            login_button = st.form_submit_button("🚀 Entrar", use_container_width=True)
            
            if login_button:
                if email and password:
                    try:
                        if USE_DATABASE:
                            user = auth_manager.login(email, password)
                        else:
                            user = auth_manager.authenticate(email, password)
                        
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
                                if USE_DATABASE:
                                    user_id = auth_manager.register(email, password, nome, empresa)
                                else:
                                    user_id = auth_manager.register_user(email, password, nome, empresa)
                                
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
        except:
            st.metric("💳 Créditos", "∞")
    
    with col3:
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.user = None
            st.rerun()
    
    st.markdown(f"🏢 **Empresa:** {user['empresa']}")
    
    total_riscos = sum(len(riscos) for riscos in AGENTES_POR_CATEGORIA.values())
    st.markdown(f"""
    <div class="new-features">
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
        
        try:
            credits = user_data_manager.get_user_credits(user['id'])
            st.markdown(f"**Créditos:** {credits}")
        except:
            st.markdown(f"**Créditos:** ∞ (Modo demo)")
        
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
                st.error(f"❌ {message}")
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
            
            else:
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
                
                creditos_necessarios = len(funcionarios_selecionados)
                
                if st.button(f"📄 GERAR {len(funcionarios_selecionados)} OS", type="primary", use_container_width=True):
                    
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
                    
                    # Debitar créditos
                    try:
                        user_data_manager.debit_credits(user['id'], creditos_necessarios)
                    except:
                        pass  # Ignorar erro no modo simplificado
                    
                    status_text.text("✅ Geração concluída!")
                    
                    # Disponibilizar downloads
                    if documentos_gerados:
                        if len(documentos_gerados) == 1:
                            st.success(f"✅ Ordem de Serviço gerada com sucesso!")
                            st.download_button(
                                label="📥 Download da Ordem de Serviço",
                                data=documentos_gerados[0]['buffer'].getvalue(),
                                file_name=f"OS_{documentos_gerados[0]['nome']}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                        else:
                            st.success(f"✅ {len(documentos_gerados)} Ordens de Serviço geradas com sucesso!")
                            
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
                    else:
                        st.error("❌ Erro: Nenhum documento foi gerado. Verifique as configurações.")
        
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
