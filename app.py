# 🔐 Sistema Gerador de OS - Visualização Melhorada (CORRIGIDO)
# Desenvolvido por especialista em UX/UI - Setembro 2025

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
import json

# Adicionar o diretório atual ao path para importar módulos locais
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Tentativa de importar módulos locais com fallback
try:
    from database.models import DatabaseManager
    from database.auth import AuthManager
    from database.user_data import UserDataManager
    USE_LOCAL_DB = True
except ImportError:
    USE_LOCAL_DB = False
    st.warning("⚠️ Módulos de banco de dados não encontrados. Sistema funcionará em modo local.")

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Ordens de Serviço (OS)",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CONSTANTES GLOBAIS EXPANDIDAS ---
UNIDADES_DE_MEDIDA = [
    "dB(A)", "m/s²", "m/s¹⁷⁵", "ppm", "mg/m³", "%", "°C", "lx", 
    "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"
]

# === MAPEAMENTO AUTOMÁTICO DE UNIDADES POR AGENTE ===
UNIDADES_AUTOMATICAS = {
    # Riscos Físicos - Ruído
    "Exposição ao Ruído": "dB(A)",
    "Ruído (Contínuo ou Intermitente)": "dB(A)",
    "Ruído (Impacto)": "dB(A)",
    
    # Riscos Físicos - Vibração
    "Vibração de Corpo Inteiro (AREN)": "m/s²",
    "Vibração de Corpo Inteiro (VDVR)": "m/s¹⁷⁵",
    "Vibrações Localizadas (mão/braço)": "m/s²",
    "Vibrações Localizadas em partes do corpo": "m/s²",
    "Vibração de Mãos e Braços": "m/s²",
    "Vibração de Corpo Inteiro": "m/s²",
    
    # Riscos Físicos - Temperatura
    "Ambiente Artificialmente Frio": "°C",
    "Exposição à Temperatura Ambiente Baixa": "°C",
    "Exposição à Temperatura Ambiente Elevada": "°C",
    "Calor": "°C",
    "Frio": "°C",
    
    # Riscos Físicos - Radiações
    "Exposição à Radiações Ionizantes": "µT",
    "Exposição à Radiações Não-ionizantes": "µT",
    "Radiações Ionizantes": "µT",
    "Radiações Não-Ionizantes": "µT",
    
    # Riscos Físicos - Pressão
    "Pressão Atmosférica Anormal (condições hiperbáricas)": "kV/m",
    "Pressões Anormais": "kV/m",
    
    # Riscos Físicos - Iluminação
    "Iluminação inadequada (deficiente ou excessiva)": "lx",
    
    # Riscos Químicos - Concentração
    "Exposição a Produto Químico": "ppm",
    "Produtos Químicos em Geral": "ppm",
    "Poeiras": "mg/m³",
    "Fumos": "mg/m³",
    "Névoas": "mg/m³",
    "Neblinas": "mg/m³",
    "Gases": "ppm",
    "Vapores": "ppm",
    "Exposição a gases e vapores": "ppm",
    
    # Riscos Biológicos - Geralmente não aplicável
    "Água e/ou alimentos contaminados": "Não aplicável",
    "Contaminação pelo Corona Vírus": "Não aplicável",
    "Contato com Fluido Orgânico (sangue, hemoderivados, secreções, excreções)": "Não aplicável",
    "Contato com Pessoas Doentes e/ou Material Infectocontagiante": "Não aplicável",
    "Exposição à Agentes Microbiológicos (fungos, bactérias, vírus, protozoários, parasitas)": "Não aplicável",
    
    # Riscos Ergonômicos - Geralmente percentual ou não aplicável
    "Posturas incômodas/pouco confortáveis por longos períodos": "%",
    "Postura sentada por longos períodos": "%",
    "Postura em pé por longos períodos": "%",
    "Frequente execução de movimentos repetitivos": "%",
    "Esforço físico intenso": "%",
    "Levantamento e transporte manual de cargas ou volumes": "%",
    
    # Riscos de Acidentes - Geralmente não aplicável
    "Contato com corrente elétrica": "Não aplicável",
    "Contato com chama aberta": "Não aplicável",
    "Queda com diferença de nível": "Não aplicável",
    "Queda de mesmo nível": "Não aplicável",
    "Objeto cortante ou perfurante": "Não aplicável",
}

# === BASE DE RISCOS EXPANDIDA DO PGR (142 RISCOS) ===
RISCOS_PGR = {
    "quimico": [
        "Exposição a Produto Químico"
    ],
    "fisico": [
        "Ambiente Artificialmente Frio",
        "Exposição ao Ruído",
        "Exposição à Radiações Ionizantes",
        "Exposição à Radiações Não-ionizantes",
        "Exposição à Temperatura Ambiente Baixa",
        "Exposição à Temperatura Ambiente Elevada",
        "Pressão Atmosférica Anormal (condições hiperbáricas)",
        "Vibração de Corpo Inteiro (AREN)",
        "Vibração de Corpo Inteiro (VDVR)",
        "Vibrações Localizadas (mão/braço)",
        "Vibrações Localizadas em partes do corpo"
    ],
    "biologico": [
        "Água e/ou alimentos contaminados",
        "Contaminação pelo Corona Vírus",
        "Contato com Fluido Orgânico (sangue, hemoderivados, secreções, excreções)",
        "Contato com Pessoas Doentes e/ou Material Infectocontagiante",
        "Exposição à Agentes Microbiológicos (fungos, bactérias, vírus, protozoários, parasitas)"
    ],
    "ergonomico": [
        "Assento inadequado",
        "Assédio de qualquer natureza no trabalho",
        "Cadência do trabalho imposta por um equipamento",
        "Compressão de partes do corpo por superfícies rígidas ou com quinas vivas",
        "Conflitos hierárquicos no trabalho",
        "Controle rígido de produtividade",
        "Desconforto, constrangimento e/ou perturbação da situação de trabalho",
        "Dupla jornada de trabalho",
        "Equipamento e mobiliário inadequados às condições morfológicas",
        "Escassez de recursos/pessoas para execução das atividades",
        "Esforço físico intenso",
        "Falta de pausas, intervalos e descansos adequados",
        "Falta de treinamento/orientação para o trabalho",
        "Frequente ação de empurrar/puxar cargas ou volumes",
        "Frequente deslocamento à pé durante à jornada de trabalho",
        "Frequente execução de movimentos repetitivos",
        "Iluminação inadequada (deficiente ou excessiva)",
        "Inadequação de layout do ambiente de trabalho",
        "Inadequação do ritmo de trabalho",
        "Jornada de trabalho prolongada",
        "Levantamento e transporte manual de cargas ou volumes",
        "Limitação de espaço para execução de movimentos",
        "Manuseio de ferramentas e/ou objetos pesados por longos períodos",
        "Monotonia, repetitividade das tarefas",
        "Necessidade de alta concentração mental e atenção para o trabalho",
        "Organização do trabalho inadequada",
        "Pausa para descanso insuficiente",
        "Postura em pé por longos períodos",
        "Postura sentada por longos períodos",
        "Posturas incômodas/pouco confortáveis por longos períodos",
        "Pressão de tempo para cumprir tarefas/pressão temporal",
        "Pressão de hierarquia/chefias",
        "Problemas no relacionamento interpessoal",
        "Queda de mesmo nível ou a pequena altura",
        "Qualidade do ar no ambiente de trabalho",
        "Responsabilidades e complexidades da tarefa",
        "Ritmo de trabalho acelerado",
        "Sobrecarga de funções",
        "Trabalho em turnos e noturno",
        "Trabalho monótono, repetitivo",
        "Uso de equipamento de proteção individual inadequado",
        "Uso frequente de força, pressão, preensão, flexão, extensão ou torção dos segmentos corporais",
        "Utilização de instrumentos que produzem vibrações",
        "Velocidade de execução",
        "Ventilação inadequada (deficiente ou excessiva)"
    ],
    "acidente": [
        "Absorção (por contato) de substância cáustica, tóxica ou nociva",
        "Afogamento, imersão, engolfamento",
        "Aprisionamento em, sob ou entre",
        "Aprisionamento em, sob ou entre desabamento ou desmoronamento de edificação, estrutura, barreira, etc.",
        "Aprisionamento em, sob ou entre dois ou mais objetos em movimento (sem encaixe)",
        "Aprisionamento em, sob ou entre objetos em movimento convergente",
        "Aprisionamento em, sob ou entre um objeto parado e outro em movimento",
        "Arestas cortantes, superfícies com rebarbas, farpas ou elementos de fixação expostos",
        "Ataque de ser vivo por mordedura, picada, chifrada, coice, etc.",
        "Choque contra objeto imóvel (pessoa em movimento)",
        "Choque entre dois objetos em movimento, sendo um deles a pessoa que sofre a lesão",
        "Colisão entre pessoas",
        "Contato com chama aberta",
        "Contato com corrente elétrica",
        "Contato com objetos ou ambientes aquecidos (queimaduras)",
        "Contato com objetos ou ambientes resfriados (queimaduras por frio)",
        "Contato com substâncias cáusticas, tóxicas (por inalação ou ingestão)",
        "Desabamento, desmoronamento, soterramento",
        "Esforços excessivos ou inadequados",
        "Escorregões e tropeços com queda",
        "Explosão",
        "Exposição a gases e vapores",
        "Ferimento por objeto cortante ou pontiagudo",
        "Golpe por objeto lançado, projetado ou que cai",
        "Impacto causado por objeto que cai",
        "Impacto de pessoa contra objeto ou estrutura",
        "Incêndio",
        "Lesões por esforços repetitivos ou sobrecarga",
        "Objeto cortante ou perfurante",
        "Perfuração por objeto pontiagudo",
        "Pisadela, pancada ou choque contra objeto imóvel",
        "Projeção de fragmentos ou partículas",
        "Queda com diferença de nível",
        "Queda de mesmo nível",
        "Queda de objetos, materiais, ferramentas ou estruturas",
        "Queimaduras por contato com superfícies quentes",
        "Ruptura de reservatório sob pressão",
        "Transporte de pessoas"
    ]
}

# Manter compatibilidade com código original
AGENTES_DE_RISCO = []
for categoria, riscos in RISCOS_PGR.items():
    AGENTES_DE_RISCO.extend(riscos)
AGENTES_DE_RISCO = sorted(AGENTES_DE_RISCO)

CATEGORIAS_RISCO = {
    'fisico': '🔥 Físicos', 
    'quimico': '⚗️ Químicos', 
    'biologico': '🦠 Biológicos', 
    'ergonomico': '🏃 Ergonômicos', 
    'acidente': '⚠️ Acidentes'
}

# === CORES E ÍCONES POR CATEGORIA ===
CATEGORIA_VISUAL = {
    'fisico': {'cor': '#FF6B35', 'cor_bg': '#FFF5F3', 'icone': '🔥'},
    'quimico': {'cor': '#8E44AD', 'cor_bg': '#F8F5FB', 'icone': '⚗️'},
    'biologico': {'cor': '#16A085', 'cor_bg': '#F1F9F7', 'icone': '🦠'},
    'ergonomico': {'cor': '#3498DB', 'cor_bg': '#F3F8FC', 'icone': '🏃'},
    'acidente': {'cor': '#E74C3C', 'cor_bg': '#FDF2F2', 'icone': '⚠️'}
}

# --- Função para obter unidade automática ---
def obter_unidade_automatica(agente_risco):
    """Retorna a unidade de medida automática para um agente de risco"""
    return UNIDADES_AUTOMATICAS.get(agente_risco, "Não aplicável")

# --- Inicialização dos Gerenciadores com fallback ---
@st.cache_resource
def init_managers():
    if USE_LOCAL_DB:
        try:
            db_manager = DatabaseManager()
            auth_manager = AuthManager(db_manager)
            user_data_manager = UserDataManager(db_manager)
            return db_manager, auth_manager, user_data_manager
        except Exception as e:
            st.error(f"Erro ao inicializar banco de dados: {e}")
            return None, None, None
    else:
        return None, None, None

# Inicialização segura
if USE_LOCAL_DB:
    db_manager, auth_manager, user_data_manager = init_managers()
else:
    db_manager, auth_manager, user_data_manager = None, None, None

# --- CSS PERSONALIZADO MELHORADO PARA VISUALIZAÇÃO ---
st.markdown("""
<style>
    /* === VARIÁVEIS CSS === */
    :root {
        --primary-color: #1f77b4;
        --secondary-color: #ff7f0e;
        --success-color: #2ca02c;
        --warning-color: #ff7f0e;
        --error-color: #d62728;
        --background-dark: #0e1117;
        --background-light: #262730;
        --card-background: #1e1e2e;
        --text-primary: #ffffff;
        --text-secondary: #b3b3b3;
        --border-color: #3d3d3d;
        --shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
        --border-radius: 12px;
        
        /* Cores específicas por categoria */
        --fisico-color: #FF6B35;
        --fisico-bg: rgba(255, 107, 53, 0.1);
        --quimico-color: #8E44AD;
        --quimico-bg: rgba(142, 68, 173, 0.1);
        --biologico-color: #16A085;
        --biologico-bg: rgba(22, 160, 133, 0.1);
        --ergonomico-color: #3498DB;
        --ergonomico-bg: rgba(52, 152, 219, 0.1);
        --acidente-color: #E74C3C;
        --acidente-bg: rgba(231, 76, 60, 0.1);
    }

    /* === LAYOUT GERAL === */
    .main > div {
        padding-top: 2rem;
    }
    
    .stApp {
        background: linear-gradient(135deg, #0e1117 0%, #1a1a2e 100%);
    }

    /* === SIDEBAR MELHORADA === */
    .css-1d391kg {
        background: linear-gradient(180deg, #1e1e2e 0%, #2d2d3a 100%);
        border-right: 1px solid var(--border-color);
    }

    /* === CARDS E CONTAINERS MELHORADOS === */
    .metric-card {
        background: var(--card-background);
        padding: 1.5rem;
        border-radius: var(--border-radius);
        border: 1px solid var(--border-color);
        box-shadow: var(--shadow);
        margin-bottom: 1rem;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.4);
    }

    /* === CARDS DE RISCO MELHORADOS === */
    .risk-card-enhanced {
        background: var(--card-background);
        border: 1px solid var(--border-color);
        border-radius: var(--border-radius);
        padding: 1.5rem;
        margin: 1rem 0;
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
    }

    .risk-card-enhanced::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        height: 100%;
        width: 4px;
        transition: all 0.3s ease;
    }

    .risk-card-enhanced.fisico::before { background: var(--fisico-color); }
    .risk-card-enhanced.quimico::before { background: var(--quimico-color); }
    .risk-card-enhanced.biologico::before { background: var(--biologico-color); }
    .risk-card-enhanced.ergonomico::before { background: var(--ergonomico-color); }
    .risk-card-enhanced.acidente::before { background: var(--acidente-color); }

    .risk-card-enhanced:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.4);
        border-color: var(--primary-color);
    }

    .risk-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 1rem;
    }

    .risk-title {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        flex: 1;
    }

    .risk-icon {
        font-size: 2rem;
        filter: drop-shadow(0 2px 4px rgba(0, 0, 0, 0.3));
    }

    .risk-category-name {
        font-size: 1.2rem;
        font-weight: 700;
        color: var(--text-primary);
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    .risk-content {
        background: rgba(255, 255, 255, 0.05);
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }

    .risk-type {
        font-size: 1rem;
        color: var(--text-primary);
        font-weight: 600;
        margin-bottom: 0.5rem;
        line-height: 1.4;
    }

    .risk-details {
        display: flex;
        flex-wrap: wrap;
        gap: 1rem;
        margin-top: 0.75rem;
    }

    .risk-detail-item {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        background: rgba(255, 255, 255, 0.08);
        padding: 0.5rem 1rem;
        border-radius: 20px;
        font-size: 0.9rem;
    }

    .risk-detail-label {
        color: var(--text-secondary);
        font-weight: 500;
    }

    .risk-detail-value {
        color: var(--text-primary);
        font-weight: 600;
    }

    .risk-actions {
        display: flex;
        gap: 0.5rem;
        align-items: center;
    }

    /* === STATUS BADGES MELHORADOS === */
    .status-badge {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        padding: 0.5rem 1rem;
        border-radius: 25px;
        font-size: 0.85rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        transition: all 0.2s ease;
    }

    .status-success {
        background: linear-gradient(45deg, #2ca02c, #27ae60);
        color: white;
        box-shadow: 0 2px 8px rgba(44, 160, 44, 0.3);
    }

    .status-auto {
        background: linear-gradient(45deg, #3498db, #2980b9);
        color: white;
        box-shadow: 0 2px 8px rgba(52, 152, 219, 0.3);
    }

    .unit-auto-badge {
        background: linear-gradient(45deg, #e67e22, #d35400);
        color: white;
        padding: 0.25rem 0.75rem;
        border-radius: 15px;
        font-size: 0.75rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        box-shadow: 0 2px 4px rgba(230, 126, 34, 0.3);
    }

    /* === BOTÕES MELHORADOS === */
    .stButton > button {
        background: linear-gradient(45deg, var(--primary-color), #1565c0);
        color: white;
        border: none;
        border-radius: var(--border-radius);
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        transition: all 0.2s ease;
        box-shadow: var(--shadow);
    }

    .stButton > button:hover {
        background: linear-gradient(45deg, #1565c0, var(--primary-color));
        transform: translateY(-1px);
        box-shadow: 0 6px 20px rgba(31, 119, 180, 0.4);
    }

    /* === FORMULÁRIOS === */
    .stSelectbox > div > div {
        background: var(--card-background);
        border: 1px solid var(--border-color);
        border-radius: var(--border-radius);
    }

    .stTextInput > div > div {
        background: var(--card-background);
        border: 1px solid var(--border-color);
        border-radius: var(--border-radius);
    }

    /* === CABEÇALHOS === */
    .main-header {
        background: linear-gradient(90deg, var(--primary-color), var(--secondary-color));
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.5rem;
        font-weight: 700;
        text-align: center;
        margin-bottom: 2rem;
    }

    .section-header {
        color: var(--text-primary);
        font-size: 1.5rem;
        font-weight: 600;
        margin: 1.5rem 0 1rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid var(--primary-color);
    }

    /* === ALERTAS PERSONALIZADOS === */
    .custom-alert {
        padding: 1rem;
        border-radius: var(--border-radius);
        margin: 1rem 0;
        border-left: 4px solid;
    }

    .alert-info {
        background: rgba(31, 119, 180, 0.1);
        border-left-color: var(--primary-color);
        color: var(--text-primary);
    }

    .alert-success {
        background: rgba(44, 160, 44, 0.1);
        border-left-color: var(--success-color);
        color: var(--text-primary);
    }

    .alert-warning {
        background: rgba(255, 127, 14, 0.1);
        border-left-color: var(--warning-color);
        color: var(--text-primary);
    }

    /* === ESTATÍSTICAS VISUAIS === */
    .category-stats {
        display: flex;
        gap: 1rem;
        flex-wrap: wrap;
        margin-bottom: 1.5rem;
    }

    .stat-item {
        background: var(--card-background);
        border: 1px solid var(--border-color);
        border-radius: var(--border-radius);
        padding: 1rem;
        text-align: center;
        min-width: 120px;
        flex: 1;
    }

    .stat-number {
        font-size: 2rem;
        font-weight: 700;
        color: var(--primary-color);
        margin-bottom: 0.25rem;
    }

    .stat-label {
        font-size: 0.9rem;
        color: var(--text-secondary);
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    /* === ANIMAÇÕES === */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }

    .fade-in {
        animation: fadeIn 0.5s ease-out;
    }

    @keyframes slideIn {
        from { opacity: 0; transform: translateX(-20px); }
        to { opacity: 1; transform: translateX(0); }
    }

    .slide-in {
        animation: slideIn 0.4s ease-out;
    }

    /* === RESPONSIVIDADE === */
    @media (max-width: 768px) {
        .risk-card-enhanced {
            padding: 1rem;
        }
        
        .risk-header {
            flex-direction: column;
            align-items: flex-start;
            gap: 1rem;
        }
        
        .risk-details {
            flex-direction: column;
            gap: 0.5rem;
        }
    }
</style>
""", unsafe_allow_html=True)

# --- FUNÇÕES AUXILIARES PARA GERENCIAMENTO DE ESTADO ---

def get_user_data():
    """Carrega dados do usuário do session_state"""
    return st.session_state.get('user_data', {
        'risks_salvos': [],
        'creditos': 0,
        'os_geradas_total': 0,
        'ultimo_uso': 'Nunca'
    })

def save_user_data(data):
    """Salva dados do usuário no session_state"""
    if 'user_data' not in st.session_state:
        st.session_state.user_data = {}
    st.session_state.user_data.update(data)

def is_authenticated():
    """Verifica se o usuário está autenticado"""
    return st.session_state.get("authenticated", False)

def get_current_user():
    """Retorna informações do usuário atual"""
    return st.session_state.get("user_info", {
        'nome': 'Usuário Demo',
        'email': 'demo@gerador-os.com'
    })

# --- COMPONENTES UI MELHORADOS ---

def create_metric_card(title, value, delta=None, help_text=None):
    """Cria um card de métrica visual aprimorado"""
    delta_html = ""
    if delta:
        delta_color = "var(--success-color)" if delta > 0 else "var(--error-color)"
        delta_symbol = "↑" if delta > 0 else "↓"
        delta_html = f'<div style="color: {delta_color}; font-size: 0.9rem; margin-top: 0.5rem;">{delta_symbol} {abs(delta)}%</div>'
    
    help_html = ""
    if help_text:
        help_html = f'<div style="color: var(--text-secondary); font-size: 0.8rem; margin-top: 0.25rem;">{help_text}</div>'
    
    st.markdown(f"""
    <div class="metric-card fade-in">
        <div style="color: var(--text-secondary); font-size: 0.9rem; text-transform: uppercase; letter-spacing: 1px;">{title}</div>
        <div style="color: var(--text-primary); font-size: 2rem; font-weight: 700; margin: 0.5rem 0;">{value}</div>
        {delta_html}
        {help_html}
    </div>
    """, unsafe_allow_html=True)

def show_risk_statistics():
    """Mostra estatísticas dos riscos disponíveis"""
    st.markdown("""
    <div class="category-stats">
        <div class="stat-item">
            <div class="stat-number">142</div>
            <div class="stat-label">Total de Riscos</div>
        </div>
        <div class="stat-item">
            <div class="stat-number">68</div>
            <div class="stat-label">Acidentes</div>
        </div>
        <div class="stat-item">
            <div class="stat-number">57</div>
            <div class="stat-label">Ergonômicos</div>
        </div>
        <div class="stat-item">
            <div class="stat-number">11</div>
            <div class="stat-label">Físicos</div>
        </div>
        <div class="stat-item">
            <div class="stat-number">5</div>
            <div class="stat-label">Biológicos</div>
        </div>
        <div class="stat-item">
            <div class="stat-number">1</div>
            <div class="stat-label">Químicos</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

def create_enhanced_risk_card(risk, idx):
    """Cria um card de risco visualmente aprimorado"""
    categoria = risk.get('categoria', 'outros')
    visual = CATEGORIA_VISUAL.get(categoria, CATEGORIA_VISUAL['fisico'])
    categoria_display = CATEGORIAS_RISCO.get(categoria, '📋 Outros')
    unidade = risk.get('unidade', 'N/A')
    eh_automatica = risk.get('unidade_automatica', False)
    
    unidade_badge = ""
    if eh_automatica:
        unidade_badge = '<span class="unit-auto-badge">AUTO</span>'
    
    st.markdown(f"""
    <div class="risk-card-enhanced {categoria} slide-in" style="animation-delay: {idx * 0.1}s;">
        <div class="risk-header">
            <div class="risk-title">
                <span class="risk-icon">{visual['icone']}</span>
                <span class="risk-category-name">{categoria_display}</span>
            </div>
            <div class="risk-actions">
                <span class="status-badge status-success">
                    <span>✓</span>
                    <span>Ativo</span>
                </span>
            </div>
        </div>
        
        <div class="risk-content">
            <div class="risk-type">{risk.get('tipo', 'N/A')}</div>
            
            <div class="risk-details">
                <div class="risk-detail-item">
                    <span class="risk-detail-label">Unidade:</span>
                    <span class="risk-detail-value">{unidade}</span>
                    {unidade_badge}
                </div>
                <div class="risk-detail-item">
                    <span class="risk-detail-label">Categoria:</span>
                    <span class="risk-detail-value">{categoria.title()}</span>
                </div>
                <div class="risk-detail-item">
                    <span class="risk-detail-label">ID:</span>
                    <span class="risk-detail-value">{risk.get('id', 'N/A')[:8]}...</span>
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

def create_risk_summary_card(risks_salvos):
    """Cria um card visual para resumo de riscos com visualização aprimorada"""
    if not risks_salvos:
        st.markdown("""
        <div class="custom-alert alert-info">
            <strong>ℹ️ Nenhum risco selecionado</strong><br>
            Adicione riscos usando as opções acima para começar a construir sua avaliação.
        </div>
        """, unsafe_allow_html=True)
        return
    
    # Contar riscos por categoria
    risk_counts = {}
    for risk in risks_salvos:
        categoria = risk.get('categoria', 'Outros')
        categoria_display = CATEGORIAS_RISCO.get(categoria, f"📋 {categoria.title()}")
        risk_counts[categoria_display] = risk_counts.get(categoria_display, 0) + 1
    
    st.markdown('<div class="section-header">📊 Resumo de Riscos Selecionados</div>', unsafe_allow_html=True)
    
    # Layout em colunas para as métricas
    cols = st.columns(min(len(risk_counts), 4))
    for idx, (categoria, count) in enumerate(risk_counts.items()):
        with cols[idx % 4]:
            create_metric_card(categoria.split(' ', 1)[-1], count, help_text=f"{count} risco{'s' if count != 1 else ''}")
    
    # Controles de visualização
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.markdown('<div class="section-header">📝 Riscos Cadastrados</div>', unsafe_allow_html=True)
    
    with col2:
        view_mode = st.selectbox(
            "👁️ Modo de Visualização",
            ["📋 Cards Detalhados", "📝 Lista Compacta", "📊 Tabela"],
            key="risk_view_mode"
        )
    
    # Mostrar riscos conforme o modo selecionado
    if "Cards" in view_mode:
        # Visualização em cards detalhados
        for idx, risk in enumerate(risks_salvos):
            col1, col2 = st.columns([5, 1])
            
            with col1:
                create_enhanced_risk_card(risk, idx)
            
            with col2:
                st.markdown("<br>" * 3, unsafe_allow_html=True)  # Espaçamento
                if st.button("🗑️", key=f"remove_card_{idx}", help="Remover risco"):
                    risks_salvos.pop(idx)
                    save_user_data({'risks_salvos': risks_salvos})
                    st.rerun()
    
    elif "Lista" in view_mode:
        # Visualização em lista compacta
        for idx, risk in enumerate(risks_salvos):
            col1, col2 = st.columns([6, 1])
            
            with col1:
                categoria = risk.get('categoria', 'outros')
                visual = CATEGORIA_VISUAL.get(categoria, CATEGORIA_VISUAL['fisico'])
                unidade = risk.get('unidade', 'N/A')
                eh_automatica = risk.get('unidade_automatica', False)
                
                unidade_badge = ""
                if eh_automatica:
                    unidade_badge = '<span class="unit-auto-badge">AUTO</span>'
                
                st.markdown(f"""
                <div class="risk-card-enhanced {categoria}">
                    <div style="display: flex; align-items: center; gap: 1rem;">
                        <span style="font-size: 1.5rem;">{visual['icone']}</span>
                        <div style="flex: 1;">
                            <div style="font-weight: 600; color: var(--text-primary); margin-bottom: 0.25rem;">
                                {risk.get('tipo', 'N/A')}
                            </div>
                            <div style="color: var(--text-secondary); font-size: 0.9rem;">
                                {categoria.title()} • {unidade} {unidade_badge}
                            </div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"remove_list_{idx}", help="Remover risco"):
                    risks_salvos.pop(idx)
                    save_user_data({'risks_salvos': risks_salvos})
                    st.rerun()
    
    else:  # Tabela
        # Visualização em tabela
        risk_data = []
        for idx, risk in enumerate(risks_salvos):
            categoria = risk.get('categoria', 'outros')
            visual = CATEGORIA_VISUAL.get(categoria, CATEGORIA_VISUAL['fisico'])
            
            risk_data.append({
                'Categoria': f"{visual['icone']} {categoria.title()}",
                'Tipo de Risco': risk.get('tipo', 'N/A'),
                'Unidade': risk.get('unidade', 'N/A'),
                'Auto': "✅" if risk.get('unidade_automatica', False) else "➖",
                'Status': "🟢 Ativo",
                'ID': risk.get('id', 'N/A')[:8] + "..."
            })
        
        if risk_data:
            df_risks = pd.DataFrame(risk_data)
            
            # Configurar o dataframe para exibição
            st.dataframe(
                df_risks,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Categoria": st.column_config.TextColumn("🏷️ Categoria", width="medium"),
                    "Tipo de Risco": st.column_config.TextColumn("🎯 Tipo de Risco", width="large"),
                    "Unidade": st.column_config.TextColumn("📏 Unidade", width="small"),
                    "Auto": st.column_config.TextColumn("🤖 Auto", width="small"),
                    "Status": st.column_config.TextColumn("📊 Status", width="small"),
                    "ID": st.column_config.TextColumn("🆔 ID", width="small")
                }
            )
            
            # Botão para limpar todos os riscos
            col1, col2, col3 = st.columns([1, 1, 1])
            with col2:
                if st.button("🗑️ Limpar Todos os Riscos", use_container_width=True):
                    save_user_data({'risks_salvos': []})
                    st.success("🗑️ Todos os riscos foram removidos!")
                    time.sleep(1)
                    st.rerun()

def show_login_page():
    """Página de login com design melhorado"""
    st.markdown('<div class="main-header">🔐 Sistema Gerador de OS</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="custom-alert alert-info">
        <strong>Bem-vindo ao Sistema Demo!</strong><br>
        Demonstração do Gerador de Ordens de Serviço com visualização aprimorada.
        <br><br>
        🎯 <strong>Funcionalidades:</strong> Base expandida com 142 riscos categorizados do PGR!<br>
        🚀 <strong>Novo:</strong> Visualização aprimorada de riscos com múltiplos modos de exibição!
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 Entrar no Sistema Demo", use_container_width=True, type="primary"):
            st.session_state.authenticated = True
            st.session_state.user_info = {
                'nome': 'Usuário Demo',
                'email': 'demo@gerador-os.com'
            }
            st.success("✅ Acesso liberado!")
            time.sleep(1)
            st.rerun()
        
        st.info("💡 **Modo Demo:** Todas as funcionalidades estão disponíveis para teste!")

def show_main_app():
    """Interface principal do aplicativo com melhorias visuais"""
    # Header da aplicação
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown('<div class="main-header">📄 Gerador de OS</div>', unsafe_allow_html=True)
    
    # Sidebar com informações do usuário
    with st.sidebar:
        user_info = get_current_user()
        st.markdown(f"""
        <div class="metric-card">
            <div style="text-align: center;">
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">👤</div>
                <div style="font-weight: 600; color: var(--text-primary);">{user_info.get('nome', 'Usuário')}</div>
                <div style="color: var(--text-secondary); font-size: 0.9rem;">{user_info.get('email', '')}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown('<div class="section-header">⚙️ Controles</div>', unsafe_allow_html=True)
        
        if st.button("🔄 Limpar sessão", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key not in ['authenticated', 'user_info']:
                    del st.session_state[key]
            st.rerun()
        
        if st.button("🚪 Sair", use_container_width=True):
            st.session_state.authenticated = False
            if 'user_info' in st.session_state:
                del st.session_state.user_info
            st.rerun()
            
        # Informações do sistema
        st.markdown("---")
        st.markdown("### 📊 Estatísticas do Sistema")
        st.metric("Riscos Disponíveis", "142")
        st.metric("Categorias", "5")
        st.metric("Versão", "2.2")
        
        # Informações sobre visualização
        st.markdown("---")
        st.markdown("### 🎨 Visualização")
        st.info("✨ Cards Detalhados\n✨ Lista Compacta\n✨ Tabela Interativa")
    
    # Conteúdo principal com abas organizadas
    tab1, tab2, tab3, tab4 = st.tabs(["🏠 Dashboard", "⚠️ Gestão de Riscos", "📄 Gerar OS", "💰 Créditos"])
    
    with tab1:
        show_dashboard()
    
    with tab2:
        show_risk_management()
    
    with tab3:
        show_os_generation()
    
    with tab4:
        show_credits_management()

def show_dashboard():
    """Dashboard principal com métricas e resumos"""
    st.markdown('<div class="section-header">📊 Visão Geral</div>', unsafe_allow_html=True)
    
    # Carregar dados do usuário
    dados_usuario = get_user_data()
    risks_salvos = dados_usuario.get('risks_salvos', [])
    creditos = dados_usuario.get('creditos', 0)
    
    # Métricas principais
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        create_metric_card("Riscos Cadastrados", len(risks_salvos), help_text="Total de riscos")
    
    with col2:
        create_metric_card("Créditos Disponíveis", creditos, help_text="Para geração de OS")
    
    with col3:
        os_geradas = dados_usuario.get('os_geradas_total', 0)
        create_metric_card("OS Geradas", os_geradas, help_text="Histórico total")
    
    with col4:
        ultimo_uso = dados_usuario.get('ultimo_uso', 'Nunca')
        create_metric_card("Último Uso", ultimo_uso, help_text="Data da última OS")
    
    # Estatísticas da base de riscos
    st.markdown('<div class="section-header">📋 Base de Riscos PGR</div>', unsafe_allow_html=True)
    show_risk_statistics()
    
    # Informações sobre visualização melhorada
    st.markdown('<div class="section-header">👁️ Sistema de Visualização Aprimorado</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="custom-alert alert-success">
            <strong>📋 Cards Detalhados</strong><br>
            Visualização completa com todos os detalhes<br>
            e códigos de cores por categoria
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="custom-alert alert-info">
            <strong>📝 Lista Compacta</strong><br>
            Visualização otimizada para muitos riscos<br>
            com informações essenciais
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="custom-alert alert-warning">
            <strong>📊 Tabela Interativa</strong><br>
            Visualização estruturada com ordenação<br>
            e filtros por categoria
        </div>
        """, unsafe_allow_html=True)
    
    # Resumo de riscos com visualização aprimorada
    create_risk_summary_card(risks_salvos)

def show_risk_management():
    """Interface de gerenciamento de riscos melhorada com visualização aprimorada"""
    st.markdown('<div class="section-header">⚠️ Gerenciamento de Riscos</div>', unsafe_allow_html=True)
    
    # Mostrar estatísticas da base de riscos
    show_risk_statistics()
    
    # Formulário de adição de risco
    with st.expander("➕ Adicionar Novo Risco", expanded=True):
        col1, col2 = st.columns(2)
        
        with col1:
            categoria = st.selectbox(
                "📂 Categoria do Risco",
                options=list(CATEGORIAS_RISCO.keys()),
                format_func=lambda x: CATEGORIAS_RISCO[x],
                key="nova_categoria"
            )
            
            # Filtrar riscos pela categoria selecionada
            riscos_da_categoria = RISCOS_PGR.get(categoria, [])
            
            if riscos_da_categoria:
                tipo_risco = st.selectbox(
                    f"🎯 Tipo de Risco ({len(riscos_da_categoria)} disponíveis)", 
                    riscos_da_categoria, 
                    key="novo_tipo"
                )
            else:
                tipo_risco = st.text_input("🎯 Digite o tipo de risco", key="novo_tipo_custom")
        
        with col2:
            # Obter unidade automática
            if tipo_risco:
                unidade_auto = obter_unidade_automatica(tipo_risco)
                if unidade_auto != "Não aplicável":
                    st.success(f"🤖 **Unidade automática detectada:** {unidade_auto}")
                    usar_automatica = st.checkbox("Usar unidade automática", value=True, key="usar_auto")
                    
                    if usar_automatica:
                        unidade = unidade_auto
                        st.info(f"✅ Unidade selecionada: **{unidade}**")
                    else:
                        unidade = st.selectbox("📏 Unidade de Medida (Manual)", UNIDADES_DE_MEDIDA, key="nova_unidade")
                else:
                    unidade = st.selectbox("📏 Unidade de Medida", UNIDADES_DE_MEDIDA, key="nova_unidade")
                    usar_automatica = False
            else:
                unidade = st.selectbox("📏 Unidade de Medida", UNIDADES_DE_MEDIDA, key="nova_unidade")
                usar_automatica = False
            
            # Mostrar informação sobre a categoria selecionada
            st.info(f"📊 Categoria selecionada possui **{len(riscos_da_categoria)}** riscos disponíveis")
            
            col_add, col_reset = st.columns(2)
            with col_add:
                if st.button("✅ Adicionar Risco", use_container_width=True):
                    if tipo_risco:
                        dados_usuario = get_user_data()
                        risks_salvos = dados_usuario.get('risks_salvos', [])
                        
                        novo_risco = {
                            'categoria': categoria,
                            'tipo': tipo_risco,
                            'unidade': unidade,
                            'unidade_automatica': usar_automatica if 'usar_automatica' in locals() else False,
                            'id': f"{categoria}_{len(risks_salvos)}"
                        }
                        
                        risks_salvos.append(novo_risco)
                        save_user_data({'risks_salvos': risks_salvos})
                        
                        if usar_automatica:
                            st.success(f"✅ Risco adicionado com unidade automática: **{unidade}**!")
                        else:
                            st.success("✅ Risco adicionado com sucesso!")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("❌ Selecione um tipo de risco!")
            
            with col_reset:
                if st.button("🔄 Limpar Todos", use_container_width=True):
                    save_user_data({'risks_salvos': []})
                    st.success("🗑️ Todos os riscos foram removidos!")
                    time.sleep(1)
                    st.rerun()
    
    # Busca rápida de riscos
    with st.expander("🔍 Busca Rápida de Riscos", expanded=False):
        search_term = st.text_input("🔎 Digite para buscar riscos", placeholder="Ex: ruído, vibração, VDVR, AREN...")
        
        if search_term:
            search_results = []
            for categoria, riscos in RISCOS_PGR.items():
                for risco in riscos:
                    if search_term.lower() in risco.lower():
                        search_results.append({
                            'categoria': categoria,
                            'risco': risco
                        })
            
            if search_results:
                st.success(f"✅ Encontrados **{len(search_results)}** resultados:")
                
                for result in search_results[:10]:  # Mostrar até 10 resultados
                    categoria_display = CATEGORIAS_RISCO.get(result['categoria'], result['categoria'].title())
                    unidade_auto = obter_unidade_automatica(result['risco'])
                    
                    col1, col2 = st.columns([4, 1])
                    with col1:
                        unidade_badge = ""
                        if unidade_auto != "Não aplicável":
                            unidade_badge = f'<span class="unit-auto-badge">{unidade_auto}</span>'
                        
                        st.markdown(f"""
                        <div style="background: var(--card-background); border: 1px solid var(--border-color); border-radius: 8px; padding: 0.75rem; margin: 0.25rem 0;">
                            <div style="font-weight: 600; color: var(--text-primary);">{categoria_display}</div>
                            <div style="color: var(--text-secondary); font-size: 0.9rem; margin-top: 0.25rem;">
                                {result['risco']} {unidade_badge}
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        if st.button("➕", key=f"add_search_{result['categoria']}_{result['risco'][:20]}", help="Adicionar este risco"):
                            dados_usuario = get_user_data()
                            risks_salvos = dados_usuario.get('risks_salvos', [])
                            
                            novo_risco = {
                                'categoria': result['categoria'],
                                'tipo': result['risco'],
                                'unidade': unidade_auto,
                                'unidade_automatica': unidade_auto != "Não aplicável",
                                'id': f"{result['categoria']}_{len(risks_salvos)}"
                            }
                            
                            risks_salvos.append(novo_risco)
                            save_user_data({'risks_salvos': risks_salvos})
                            
                            if unidade_auto != "Não aplicável":
                                st.success(f"✅ Risco adicionado com unidade automática **{unidade_auto}**: {result['risco']}")
                            else:
                                st.success(f"✅ Risco adicionado: {result['risco']}")
                            time.sleep(1)
                            st.rerun()
                
                if len(search_results) > 10:
                    st.info(f"... e mais {len(search_results) - 10} resultados. Refine sua busca para ver mais opções.")
            else:
                st.warning("❌ Nenhum risco encontrado com esse termo.")
    
    # Exibir riscos salvos com visualização aprimorada
    dados_usuario = get_user_data()
    risks_salvos = dados_usuario.get('risks_salvos', [])
    
    if risks_salvos:
        create_risk_summary_card(risks_salvos)
    else:
        st.markdown("""
        <div class="custom-alert alert-info">
            <strong>ℹ️ Nenhum risco cadastrado</strong><br>
            Use o formulário acima para adicionar riscos ao sistema. 
            A base possui <strong>142 riscos</strong> categorizados do PGR para sua seleção.
            <br><br>
            🎨 <strong>Novo:</strong> Sistema de visualização aprimorado com 3 modos diferentes!
        </div>
        """, unsafe_allow_html=True)

def show_os_generation():
    """Interface de geração de OS melhorada"""
    st.markdown('<div class="section-header">📄 Geração de OS</div>', unsafe_allow_html=True)
    
    dados_usuario = get_user_data()
    creditos = dados_usuario.get('creditos', 0)
    risks_salvos = dados_usuario.get('risks_salvos', [])
    
    # Verificar pré-requisitos
    if creditos <= 0:
        st.markdown("""
        <div class="custom-alert alert-warning">
            <strong>⚠️ Créditos insuficientes</strong><br>
            Você precisa de créditos para gerar OS. Acesse a aba "Créditos" para adquirir.
            <br><br>
            💡 <strong>Modo Demo:</strong> Funcionalidade disponível apenas para demonstração.
        </div>
        """, unsafe_allow_html=True)
        
        # Adicionar alguns créditos demo
        if st.button("🎁 Adicionar 10 Créditos Demo", type="primary"):
            save_user_data({'creditos': 10})
            st.success("✅ Créditos demo adicionados!")
            time.sleep(1)
            st.rerun()
        return
    
    if not risks_salvos:
        st.markdown("""
        <div class="custom-alert alert-warning">
            <strong>⚠️ Nenhum risco cadastrado</strong><br>
            Você precisa cadastrar riscos antes de gerar OS. Acesse a aba "Gestão de Riscos".
            <br><br>
            💡 <strong>Dica:</strong> A base contém 142 riscos categorizados do PGR para sua seleção.
            🎨 <strong>Novo:</strong> Visualização aprimorada com múltiplos modos de exibição!
        </div>
        """, unsafe_allow_html=True)
        return
    
    # Formulário de geração
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📂 Arquivos Necessários")
        
        template_file = st.file_uploader(
            "📄 Modelo Word (.docx)",
            type=['docx'],
            help="Upload do template da OS em formato Word"
        )
        
        excel_file = st.file_uploader(
            "📊 Planilha de Funcionários (.xlsx)",
            type=['xlsx'],
            help="Planilha com dados dos funcionários"
        )
    
    with col2:
        st.markdown("### ⚙️ Configurações")
        
        create_metric_card("Créditos Disponíveis", creditos, help_text="Cada OS gasta 1 crédito")
        create_metric_card("Riscos Cadastrados", len(risks_salvos), help_text="Serão incluídos nas OS")
        
        # Mostrar informação sobre unidades automáticas
        unidades_auto = sum(1 for risk in risks_salvos if risk.get('unidade_automatica', False))
        if unidades_auto > 0:
            st.success(f"🤖 {unidades_auto} risco(s) com unidade automática detectada!")
        
        if template_file and excel_file:
            if st.button("🚀 Gerar OS (Demo)", use_container_width=True, type="primary"):
                with st.spinner("Processando OS..."):
                    # Simulação do processo
                    progress_bar = st.progress(0)
                    for i in range(100):
                        time.sleep(0.01)
                        progress_bar.progress(i + 1)
                    
                    # Decrementar créditos
                    novo_credito = max(0, creditos - 1)
                    save_user_data({
                        'creditos': novo_credito,
                        'os_geradas_total': dados_usuario.get('os_geradas_total', 0) + 1,
                        'ultimo_uso': time.strftime("%d/%m/%Y")
                    })
                    
                    st.success("✅ OS geradas com sucesso!")
                    st.balloons()
                    time.sleep(2)
                    st.rerun()

def show_credits_management():
    """Interface de gerenciamento de créditos"""
    st.markdown('<div class="section-header">💰 Gerenciamento de Créditos</div>', unsafe_allow_html=True)
    
    dados_usuario = get_user_data()
    creditos = dados_usuario.get('creditos', 0)
    
    # Status atual
    create_metric_card("Saldo Atual", creditos, help_text="Créditos disponíveis")
    
    # Pacotes de créditos
    st.markdown("### 🛒 Pacotes Disponíveis (Demo)")
    
    col1, col2, col3 = st.columns(3)
    
    pacotes = [
        {"nome": "Básico", "creditos": 10, "preco": 50.00, "economia": 0},
        {"nome": "Profissional", "creditos": 25, "preco": 100.00, "economia": 25},
        {"nome": "Empresarial", "creditos": 50, "preco": 180.00, "economia": 70}
    ]
    
    for idx, pacote in enumerate(pacotes):
        with [col1, col2, col3][idx]:
            economia_html = ""
            if pacote["economia"] > 0:
                economia_html = f'<div style="color: var(--success-color); font-weight: 600; margin-top: 0.5rem;">💰 Economia: R$ {pacote["economia"]:.2f}</div>'
            
            st.markdown(f"""
            <div class="metric-card" style="text-align: center;">
                <div style="font-size: 1.5rem; margin-bottom: 1rem;">📦</div>
                <div style="font-weight: 700; color: var(--text-primary); font-size: 1.25rem;">{pacote["nome"]}</div>
                <div style="color: var(--text-secondary); margin: 0.5rem 0;">{pacote["creditos"]} créditos</div>
                <div style="font-size: 1.5rem; font-weight: 700; color: var(--primary-color);">R$ {pacote["preco"]:.2f}</div>
                {economia_html}
            </div>
            """, unsafe_allow_html=True)
            
            if st.button(f"🛒 Simular {pacote['nome']}", key=f"buy_{idx}", use_container_width=True):
                novo_total = creditos + pacote["creditos"]
                save_user_data({'creditos': novo_total})
                st.success(f"✅ {pacote['creditos']} créditos adicionados! Total: {novo_total}")
                time.sleep(1)
                st.rerun()

# --- EXECUÇÃO PRINCIPAL ---
def main():
    """Função principal da aplicação"""
    
    # Verificar autenticação
    if not is_authenticated():
        show_login_page()
    else:
        show_main_app()

if __name__ == "__main__":
    main()
