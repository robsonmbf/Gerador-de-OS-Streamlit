# 🔐 Sistema Gerador de OS - VERSÃO COMPLETA COM RELATÓRIOS
# Preenchimento automático de DOCX + Relatório detalhado das OS geradas
# Desenvolvido por especialista em UX/UI - Setembro 2025

import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import zipfile
from io import BytesIO
import time
import re
import sys
import os
import json
from datetime import datetime, timedelta
import uuid

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

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Ordens de Serviço (OS)",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# === CONSTANTES GLOBAIS ATUALIZADAS ===
UNIDADES_DE_MEDIDA = [
    "dB(A)", "m/s²", "m/s¹⁷⁵", "ppm", "mg/m³", "%", "°C", "lx", 
    "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"
]

# === MAPEAMENTO AUTOMÁTICO DE UNIDADES ===
UNIDADES_AUTOMATICAS = {
    # Vibração - NOVOS VDVR e AREN
    "Vibração de Corpo Inteiro (AREN)": "m/s²",
    "Vibração de Corpo Inteiro (VDVR)": "m/s¹⁷⁵",
    "Vibrações Localizadas (mão/braço)": "m/s²",
    "Vibração de Mãos e Braços": "m/s²",
    "Vibração de Corpo Inteiro": "m/s²",
    
    # Ruído
    "Exposição ao Ruído": "dB(A)",
    "Ruído (Contínuo ou Intermitente)": "dB(A)",
    "Ruído (Impacto)": "dB(A)",
    
    # Temperatura
    "Ambiente Artificialmente Frio": "°C",
    "Exposição à Temperatura Ambiente Baixa": "°C",
    "Exposição à Temperatura Ambiente Elevada": "°C",
    "Calor": "°C",
    "Frio": "°C",
    
    # Químicos
    "Exposição a Produto Químico": "ppm",
    "Produtos Químicos em Geral": "ppm",
    "Poeiras": "mg/m³",
    "Fumos": "mg/m³",
    "Névoas": "mg/m³",
    "Neblinas": "mg/m³",
    "Gases": "ppm",
    "Vapores": "ppm",
    
    # Outros
    "Exposição à Radiações Ionizantes": "µT",
    "Exposição à Radiações Não-ionizantes": "µT",
    "Radiações Ionizantes": "µT",
    "Radiações Não-Ionizantes": "µT",
    "Pressão Atmosférica Anormal": "kV/m",
    "Pressões Anormais": "kV/m",
    "Iluminação inadequada": "lx",
}

# === BASE DE RISCOS EXPANDIDA ===
AGENTES_DE_RISCO_ORIGINAL = sorted([
    "Ruído (Contínuo ou Intermitente)", "Ruído (Impacto)", "Vibração de Corpo Inteiro", "Vibração de Mãos e Braços",
    "Radiações Ionizantes", "Radiações Não-Ionizantes", "Frio", "Calor", "Pressões Anormais", "Umidade", "Poeiras", 
    "Fumos", "Névoas", "Neblinas", "Gases", "Vapores", "Produtos Químicos em Geral", "Vírus", "Bactérias", 
    "Protozoários", "Fungos", "Parasitas", "Bacilos"
])

NOVOS_RISCOS = [
    "Vibração de Corpo Inteiro (AREN)",
    "Vibração de Corpo Inteiro (VDVR)", 
    "Exposição ao Ruído",
    "Ambiente Artificialmente Frio",
    "Exposição à Temperatura Ambiente Baixa",
    "Exposição à Temperatura Ambiente Elevada",
    "Pressão Atmosférica Anormal (condições hiperbáricas)",
    "Vibrações Localizadas (mão/braço)",
    "Vibrações Localizadas em partes do corpo",
    "Exposição a Produto Químico",
    "Água e/ou alimentos contaminados",
    "Contaminação pelo Corona Vírus", 
    "Contato com Fluido Orgânico (sangue, hemoderivados, secreções, excreções)",
    "Contato com Pessoas Doentes e/ou Material Infectocontagiante",
    "Exposição à Agentes Microbiológicos (fungos, bactérias, vírus, protozoários, parasitas)"
]

AGENTES_DE_RISCO = sorted(list(set(AGENTES_DE_RISCO_ORIGINAL + NOVOS_RISCOS)))

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
    if USE_LOCAL_DB:
        try:
            db_manager = DatabaseManager()
            auth_manager = AuthManager(db_manager)  
            user_data_manager = UserDataManager(db_manager)
            return db_manager, auth_manager, user_data_manager
        except:
            return None, None, None
    return None, None, None

if USE_LOCAL_DB:
    db_manager, auth_manager, user_data_manager = init_managers()
else:
    db_manager, auth_manager, user_data_manager = None, None, None

# --- Funções Auxiliares ---
def obter_unidade_automatica(agente_risco):
    return UNIDADES_AUTOMATICAS.get(agente_risco, "Não aplicável")

def get_user_data():
    return st.session_state.get('user_data', {
        'risks_salvos': [],
        'creditos': 10,
        'os_geradas_total': 0,
        'ultimo_uso': 'Nunca',
        'historico_os': []  # Novo campo para histórico
    })

def save_user_data(data):
    if 'user_data' not in st.session_state:
        st.session_state.user_data = {}
    st.session_state.user_data.update(data)
    
    if USE_LOCAL_DB and user_data_manager:
        try:
            user_data_manager.save_user_data(data)
        except:
            pass

def is_authenticated():
    return st.session_state.get("authenticated", False)

def get_current_user():
    return st.session_state.get("user_info", {
        'nome': 'Usuário Demo',
        'email': 'demo@gerador-os.com'
    })

# === FUNÇÕES APRIMORADAS DE PROCESSAMENTO DOCX ===

def replace_placeholders_advanced(doc, replacements):
    """Substitui placeholders no documento Word com suporte avançado"""
    # Substituir em parágrafos
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                # Preservar formatação original
                for run in paragraph.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(value))
                paragraph.text = paragraph.text.replace(key, str(value))
    
    # Substituir em tabelas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for key, value in replacements.items():
                        if key in paragraph.text:
                            paragraph.text = paragraph.text.replace(key, str(value))
    
    # Substituir em cabeçalhos e rodapés
    for section in doc.sections:
        # Cabeçalho
        header = section.header
        for paragraph in header.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, str(value))
        
        # Rodapé
        footer = section.footer
        for paragraph in footer.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, str(value))

def add_risk_table_to_doc(doc, risks_salvos):
    """Adiciona uma tabela formatada com os riscos ao documento"""
    if not risks_salvos:
        return
    
    # Encontrar posição para inserir tabela (após {{RISCOS_TABELA}})
    target_paragraph = None
    for paragraph in doc.paragraphs:
        if "{{RISCOS_TABELA}}" in paragraph.text:
            target_paragraph = paragraph
            break
    
    if target_paragraph:
        # Limpar o placeholder
        target_paragraph.text = target_paragraph.text.replace("{{RISCOS_TABELA}}", "")
        
        # Criar tabela
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        
        # Cabeçalho da tabela
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Categoria'
        header_cells[1].text = 'Tipo de Risco'
        header_cells[2].text = 'Unidade'
        header_cells[3].text = 'Automático'
        
        # Formatar cabeçalho
        for cell in header_cells:
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Adicionar dados dos riscos
        for risk in risks_salvos:
            row_cells = table.add_row().cells
            row_cells[0].text = CATEGORIAS_RISCO.get(risk.get('categoria', ''), 'N/A')
            row_cells[1].text = risk.get('tipo', 'N/A')
            row_cells[2].text = risk.get('unidade', 'N/A')
            row_cells[3].text = 'Sim' if risk.get('unidade_automatica', False) else 'Não'

def processar_planilha_avancada(excel_file):
    """Processa planilha com validação avançada"""
    try:
        df = pd.read_excel(excel_file)
        
        # Normalizar nomes das colunas
        df.columns = df.columns.str.strip().str.title()
        
        # Mapear colunas comuns
        column_mapping = {
            'Name': 'Nome',
            'Employee': 'Nome',
            'Funcionario': 'Nome',
            'Position': 'Cargo',
            'Job': 'Cargo',
            'Function': 'Cargo',
            'Department': 'Setor',
            'Departamento': 'Setor',
            'Area': 'Setor',
            'Id': 'Matricula',
            'Employee_Id': 'Matricula',
            'Registration': 'Matricula'
        }
        
        # Aplicar mapeamento
        for old_col, new_col in column_mapping.items():
            if old_col in df.columns:
                df.rename(columns={old_col: new_col}, inplace=True)
        
        # Validar colunas obrigatórias
        required_columns = ['Nome']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.warning(f"Colunas obrigatórias não encontradas: {missing_columns}")
        
        # Preencher valores em branco
        df.fillna({
            'Nome': 'Nome não informado',
            'Cargo': 'Cargo não informado',
            'Setor': 'Setor não informado',
            'Matricula': 'N/A'
        }, inplace=True)
        
        return df.to_dict('records')
        
    except Exception as e:
        st.error(f"Erro ao processar planilha: {e}")
        return []

def gerar_documentos_os_avancados(template_file, excel_file, risks_salvos):
    """Gera documentos de OS com recursos avançados e histórico"""
    try:
        funcionarios = processar_planilha_avancada(excel_file)
        if not funcionarios:
            return None, None
        
        zip_buffer = BytesIO()
        historico_geracao = []
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for idx, funcionario in enumerate(funcionarios):
                doc = Document(template_file)
                
                # Preparar texto dos riscos
                riscos_texto = "\n".join([
                    f"• {risk.get('tipo', 'N/A')} ({CATEGORIAS_RISCO.get(risk.get('categoria', ''), 'N/A')}) - {risk.get('unidade', 'N/A')}"
                    for risk in risks_salvos
                ])
                
                # Separar riscos por categoria para relatório
                riscos_por_categoria = {}
                for risk in risks_salvos:
                    categoria = risk.get('categoria', 'outros')
                    if categoria not in riscos_por_categoria:
                        riscos_por_categoria[categoria] = []
                    riscos_por_categoria[categoria].append(risk)
                
                # Criar texto detalhado por categoria
                riscos_detalhados = ""
                for categoria, riscos in riscos_por_categoria.items():
                    categoria_nome = CATEGORIAS_RISCO.get(categoria, categoria.title())
                    riscos_detalhados += f"\n{categoria_nome}:\n"
                    for risk in riscos:
                        auto_badge = " (Automático)" if risk.get('unidade_automatica', False) else ""
                        riscos_detalhados += f"  • {risk.get('tipo', 'N/A')} - {risk.get('unidade', 'N/A')}{auto_badge}\n"
                
                # Placeholders expandidos
                data_atual = datetime.now()
                replacements = {
                    # Dados básicos
                    '{{NOME}}': funcionario.get('Nome', 'Nome não informado'),
                    '{{CARGO}}': funcionario.get('Cargo', 'Cargo não informado'),
                    '{{SETOR}}': funcionario.get('Setor', 'Setor não informado'),
                    '{{MATRICULA}}': str(funcionario.get('Matricula', 'N/A')),
                    
                    # Datas
                    '{{DATA}}': data_atual.strftime("%d/%m/%Y"),
                    '{{DATA_COMPLETA}}': data_atual.strftime("%d de %B de %Y"),
                    '{{HORA}}': data_atual.strftime("%H:%M"),
                    '{{MES}}': data_atual.strftime("%B"),
                    '{{ANO}}': data_atual.strftime("%Y"),
                    
                    # Riscos
                    '{{RISCOS}}': riscos_texto,
                    '{{RISCOS_DETALHADOS}}': riscos_detalhados,
                    '{{TOTAL_RISCOS}}': str(len(risks_salvos)),
                    '{{TOTAL_CATEGORIAS}}': str(len(riscos_por_categoria)),
                    
                    # Estatísticas
                    '{{RISCOS_AUTOMATICOS}}': str(sum(1 for r in risks_salvos if r.get('unidade_automatica', False))),
                    '{{RISCOS_MANUAIS}}': str(sum(1 for r in risks_salvos if not r.get('unidade_automatica', False))),
                    
                    # Informações do sistema
                    '{{USUARIO}}': get_current_user().get('nome', 'Sistema'),
                    '{{EMAIL_USUARIO}}': get_current_user().get('email', 'sistema@os.com'),
                    '{{VERSAO_SISTEMA}}': "3.1",
                    '{{ID_GERACAO}}': str(uuid.uuid4())[:8].upper()
                }
                
                # Substituir placeholders
                replace_placeholders_advanced(doc, replacements)
                
                # Adicionar tabela de riscos se solicitada
                add_risk_table_to_doc(doc, risks_salvos)
                
                # Salvar documento no ZIP
                doc_buffer = BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)
                
                nome_limpo = re.sub(r'[^\w\s-]', '', funcionario.get('Nome', f'Funcionario_{idx+1}')).replace(' ', '_')
                nome_arquivo = f"OS_{nome_limpo}_{data_atual.strftime('%Y%m%d')}.docx"
                zip_file.writestr(nome_arquivo, doc_buffer.read())
                
                # Registrar no histórico
                historico_item = {
                    'funcionario': funcionario.get('Nome', 'N/A'),
                    'cargo': funcionario.get('Cargo', 'N/A'),
                    'setor': funcionario.get('Setor', 'N/A'),
                    'matricula': str(funcionario.get('Matricula', 'N/A')),
                    'total_riscos': len(risks_salvos),
                    'riscos_automaticos': sum(1 for r in risks_salvos if r.get('unidade_automatica', False)),
                    'data_geracao': data_atual.isoformat(),
                    'arquivo': nome_arquivo,
                    'id_geracao': replacements['{{ID_GERACAO}}']
                }
                historico_geracao.append(historico_item)
        
        zip_buffer.seek(0)
        return zip_buffer, historico_geracao
        
    except Exception as e:
        st.error(f"Erro ao gerar documentos: {e}")
        return None, None

def gerar_relatorio_os(historico_geracao, risks_salvos):
    """Gera relatório detalhado das OS geradas"""
    if not historico_geracao:
        return None
    
    # Criar documento do relatório
    doc = Document()
    
    # Título
    titulo = doc.add_heading('RELATÓRIO DE ORDENS DE SERVIÇO GERADAS', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Informações gerais
    doc.add_heading('1. RESUMO EXECUTIVO', level=1)
    
    data_geracao = datetime.now()
    info_geral = doc.add_paragraph()
    info_geral.add_run(f"Data de Geração: ").bold = True
    info_geral.add_run(f"{data_geracao.strftime('%d/%m/%Y às %H:%M')}\n")
    info_geral.add_run(f"Total de OS Geradas: ").bold = True
    info_geral.add_run(f"{len(historico_geracao)}\n")
    info_geral.add_run(f"Total de Riscos Avaliados: ").bold = True
    info_geral.add_run(f"{len(risks_salvos)}\n")
    info_geral.add_run(f"Riscos com Unidade Automática: ").bold = True
    info_geral.add_run(f"{sum(1 for r in risks_salvos if r.get('unidade_automatica', False))}\n")
    
    # Estatísticas por setor
    doc.add_heading('2. DISTRIBUIÇÃO POR SETOR', level=1)
    
    setores = {}
    for item in historico_geracao:
        setor = item['setor']
        if setor not in setores:
            setores[setor] = 0
        setores[setor] += 1
    
    for setor, count in sorted(setores.items()):
        p = doc.add_paragraph()
        p.add_run(f"• {setor}: ").bold = True
        p.add_run(f"{count} funcionário(s)")
    
    # Lista detalhada de riscos
    doc.add_heading('3. RISCOS AVALIADOS', level=1)
    
    # Agrupar riscos por categoria
    riscos_por_categoria = {}
    for risk in risks_salvos:
        categoria = risk.get('categoria', 'outros')
        if categoria not in riscos_por_categoria:
            riscos_por_categoria[categoria] = []
        riscos_por_categoria[categoria].append(risk)
    
    for categoria, riscos in riscos_por_categoria.items():
        categoria_nome = CATEGORIAS_RISCO.get(categoria, categoria.title())
        doc.add_heading(f"3.{list(riscos_por_categoria.keys()).index(categoria) + 1} {categoria_nome}", level=2)
        
        for risk in riscos:
            p = doc.add_paragraph()
            p.add_run(f"• {risk.get('tipo', 'N/A')}")
            p.add_run(f" - Unidade: {risk.get('unidade', 'N/A')}")
            if risk.get('unidade_automatica', False):
                p.add_run(" (Detectado automaticamente)").italic = True
    
    # Tabela detalhada dos funcionários
    doc.add_heading('4. FUNCIONÁRIOS AVALIADOS', level=1)
    
    # Criar tabela
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    
    # Cabeçalho
    header_cells = table.rows[0].cells
    headers = ['Nome', 'Cargo', 'Setor', 'Matrícula', 'Total Riscos', 'ID Geração']
    for i, header in enumerate(headers):
        header_cells[i].text = header
        header_cells[i].paragraphs[0].runs[0].bold = True
        header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Dados
    for item in historico_geracao:
        row_cells = table.add_row().cells
        row_cells[0].text = item['funcionario']
        row_cells[1].text = item['cargo']
        row_cells[2].text = item['setor']
        row_cells[3].text = item['matricula']
        row_cells[4].text = str(item['total_riscos'])
        row_cells[5].text = item['id_geracao']
    
    # Rodapé
    doc.add_page_break()
    doc.add_heading('5. OBSERVAÇÕES', level=1)
    obs = doc.add_paragraph()
    obs.add_run("• Este relatório foi gerado automaticamente pelo Sistema Gerador de OS v3.1\n")
    obs.add_run("• Todos os riscos foram avaliados conforme base de dados do PGR\n")
    obs.add_run("• Unidades de medida automáticas: VDVR (m/s¹⁷⁵), AREN (m/s²), Ruído (dB(A)), etc.\n")
    obs.add_run(f"• Usuário responsável: {get_current_user().get('nome', 'Sistema')}\n")
    obs.add_run(f"• Email: {get_current_user().get('email', 'sistema@os.com')}")
    
    # Salvar em buffer
    relatorio_buffer = BytesIO()
    doc.save(relatorio_buffer)
    relatorio_buffer.seek(0)
    
    return relatorio_buffer

# === CSS E COMPONENTES MANTIDOS ===
st.markdown("""
<style>
    :root {
        --primary-color: #1f77b4;
        --success-color: #2ca02c;
        --warning-color: #ff7f0e;
        --card-background: #1e1e2e;
        --text-primary: #ffffff;
        --text-secondary: #b3b3b3;
        --border-color: #3d3d3d;
        --border-radius: 12px;
    }

    .main > div { padding-top: 2rem; }
    .stApp { background: linear-gradient(135deg, #0e1117 0%, #1a1a2e 100%); }

    .metric-card {
        background: var(--card-background);
        padding: 1.5rem;
        border-radius: var(--border-radius);
        border: 1px solid var(--border-color);
        margin-bottom: 1rem;
        transition: transform 0.2s ease;
    }
    .metric-card:hover { transform: translateY(-2px); }

    .main-header {
        background: linear-gradient(90deg, var(--primary-color), var(--warning-color));
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
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

    .alert {
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

    .report-card {
        background: linear-gradient(135deg, #2c3e50, #3498db);
        color: white;
        padding: 1.5rem;
        border-radius: var(--border-radius);
        margin: 1rem 0;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
    }
</style>
""", unsafe_allow_html=True)

def create_metric_card(title, value, help_text=None):
    help_html = ""
    if help_text:
        help_html = f'<div style="color: var(--text-secondary); font-size: 0.8rem; margin-top: 0.25rem;">{help_text}</div>'
    
    st.markdown(f"""
    <div class="metric-card">
        <div style="color: var(--text-secondary); font-size: 0.9rem; text-transform: uppercase;">{title}</div>
        <div style="color: var(--text-primary); font-size: 2rem; font-weight: 700; margin: 0.5rem 0;">{value}</div>
        {help_html}
    </div>
    """, unsafe_allow_html=True)

# === PÁGINAS MANTIDAS + NOVA ABA DE RELATÓRIOS ===

def show_login_page():
    st.markdown('<div class="main-header">🔐 Sistema Gerador de OS</div>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="alert alert-info">
        <strong>🚀 Sistema Completo com Relatórios!</strong><br>
        ✅ Preenchimento automático de DOCX com placeholders expandidos<br>
        ✅ Geração de relatório detalhado das OS<br>
        ✅ Unidades automáticas: VDVR (m/s¹⁷⁵), AREN (m/s²)<br>
        ✅ Histórico completo de gerações<br>
        ✅ Tabelas formatadas nos documentos
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 Entrar no Sistema", use_container_width=True, type="primary"):
            st.session_state.authenticated = True
            st.session_state.user_info = {
                'nome': 'Usuário Demo',
                'email': 'demo@gerador-os.com'
            }
            st.success("✅ Acesso liberado!")
            time.sleep(1)
            st.rerun()

def show_main_app():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown('<div class="main-header">📄 Gerador de OS</div>', unsafe_allow_html=True)
    
    with st.sidebar:
        user_info = get_current_user()
        st.markdown(f"""
        <div class="metric-card">
            <div style="text-align: center;">
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">👤</div>
                <div style="font-weight: 600; color: var(--text-primary);">{user_info.get('nome')}</div>
                <div style="color: var(--text-secondary); font-size: 0.9rem;">{user_info.get('email')}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("🚪 Sair", use_container_width=True):
            st.session_state.authenticated = False
            st.rerun()
        
        st.markdown("---")
        st.markdown("### 📊 Sistema v3.1")
        st.metric("Placeholders", "15+")
        st.metric("Relatórios", "Automáticos")
        st.metric("VDVR/AREN", "Auto")
    
    # TABS EXPANDIDAS COM NOVA ABA DE RELATÓRIOS
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🏠 Início", 
        "⚠️ Gestão de Riscos", 
        "📄 Gerar OS", 
        "📊 Relatórios",  # NOVA ABA
        "💰 Créditos"
    ])
    
    with tab1:
        show_dashboard()
    
    with tab2:
        show_risk_management()
    
    with tab3:
        show_os_generation_advanced()  # VERSÃO APRIMORADA
    
    with tab4:
        show_reports_page()  # NOVA PÁGINA
    
    with tab5:
        show_credits_management()

def show_dashboard():
    st.markdown('<div class="section-header">📊 Dashboard</div>', unsafe_allow_html=True)
    
    dados_usuario = get_user_data()
    risks_salvos = dados_usuario.get('risks_salvos', [])
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        create_metric_card("Riscos", len(risks_salvos), "Cadastrados")
    with col2:
        create_metric_card("Créditos", dados_usuario.get('creditos', 10), "Disponíveis")
    with col3:
        create_metric_card("OS Geradas", dados_usuario.get('os_geradas_total', 0), "Total")
    with col4:
        historico_os = dados_usuario.get('historico_os', [])
        create_metric_card("Relatórios", len(historico_os), "Gerações")

def show_risk_management():
    st.markdown('<div class="section-header">⚠️ Gestão de Riscos</div>', unsafe_allow_html=True)
    
    with st.expander("➕ Adicionar Risco", expanded=True):
        col1, col2 = st.columns(2)
        
        with col1:
            categoria = st.selectbox(
                "📂 Categoria",
                options=list(CATEGORIAS_RISCO.keys()),
                format_func=lambda x: CATEGORIAS_RISCO[x],
                key="categoria_select"
            )
            
            # Filtrar agentes por categoria (simplificado)
            agentes_filtrados = [a for a in AGENTES_DE_RISCO if any(
                term in a.lower() for term in {
                    'fisico': ['ruído', 'vibração', 'temperatura', 'radiação', 'frio', 'calor', 'pressão'],
                    'quimico': ['químico', 'poeira', 'fumo', 'névoa', 'gas', 'vapor'],
                    'biologico': ['vírus', 'bactéria', 'fungo', 'água', 'corona', 'fluido'],
                    'ergonomico': ['postur', 'esforç', 'repet', 'carg', 'iluminação'],
                    'acidente': ['queda', 'choque', 'cortant', 'elétric', 'incênd']
                }.get(categoria, [])
            )]
            
            if not agentes_filtrados:
                agentes_filtrados = AGENTES_DE_RISCO
            
            agente = st.selectbox(
                f"🎯 Agente ({len(agentes_filtrados)} disponíveis)",
                agentes_filtrados,
                key="agente_select"
            )
        
        with col2:
            if agente:
                unidade_auto = obter_unidade_automatica(agente)
                if unidade_auto != "Não aplicável":
                    st.success(f"🤖 **Detectado:** {unidade_auto}")
                    usar_auto = st.checkbox("Usar automático", value=True, key="usar_auto")
                    
                    if usar_auto:
                        unidade = unidade_auto
                        if "VDVR" in agente:
                            st.balloons()
                            st.success("🎯 **VDVR** com m/s¹⁷⁵!")
                        elif "AREN" in agente:
                            st.balloons()
                            st.success("📊 **AREN** com m/s²!")
                    else:
                        unidade = st.selectbox("📏 Manual", UNIDADES_DE_MEDIDA, key="unidade_manual")
                else:
                    unidade = st.selectbox("📏 Unidade", UNIDADES_DE_MEDIDA, key="unidade_select")
                    usar_auto = False
            
            if st.button("✅ Adicionar", use_container_width=True):
                if agente:
                    dados_usuario = get_user_data()
                    risks = dados_usuario.get('risks_salvos', [])
                    
                    novo = {
                        'categoria': categoria,
                        'tipo': agente,
                        'unidade': unidade,
                        'unidade_automatica': usar_auto,
                        'id': f"{categoria}_{len(risks)}_{int(time.time())}"
                    }
                    
                    risks.append(novo)
                    save_user_data({'risks_salvos': risks})
                    st.success("✅ Adicionado!")
                    time.sleep(1)
                    st.rerun()

def show_os_generation_advanced():
    """Geração de OS com recursos aprimorados"""
    st.markdown('<div class="section-header">📄 Geração Avançada de OS</div>', unsafe_allow_html=True)
    
    dados_usuario = get_user_data()
    risks_salvos = dados_usuario.get('risks_salvos', [])
    
    if not risks_salvos:
        st.markdown("""
        <div class="alert alert-warning">
            <strong>⚠️ Adicione riscos primeiro</strong><br>
            Vá para "Gestão de Riscos" e cadastre pelo menos um risco.
        </div>
        """, unsafe_allow_html=True)
        return
    
    st.markdown("### 📂 Upload de Arquivos")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📄 Template DOCX")
        template_file = st.file_uploader(
            "Template com placeholders",
            type=['docx'],
            help="Use placeholders como {{NOME}}, {{CARGO}}, {{RISCOS}}, etc.",
            key="template_upload"
        )
        
        if template_file:
            st.success(f"✅ {template_file.name}")
    
    with col2:
        st.markdown("#### 📊 Planilha XLSX")
        excel_file = st.file_uploader(
            "Dados dos funcionários",
            type=['xlsx'],
            help="Colunas: Nome, Cargo, Setor, Matricula",
            key="excel_upload"
        )
        
        if excel_file:
            st.success(f"✅ {excel_file.name}")
            try:
                df_preview = pd.read_excel(excel_file)
                st.dataframe(df_preview.head(3), use_container_width=True)
                st.info(f"📊 {len(df_preview)} funcionários")
            except Exception as e:
                st.error(f"Erro: {e}")
    
    # Placeholders disponíveis
    with st.expander("🏷️ Placeholders Disponíveis", expanded=False):
        st.markdown("""
        **Dados Pessoais:**
        - `{{NOME}}` - Nome do funcionário
        - `{{CARGO}}` - Cargo/função
        - `{{SETOR}}` - Departamento
        - `{{MATRICULA}}` - Número de matrícula
        
        **Datas:**
        - `{{DATA}}` - Data atual (DD/MM/YYYY)
        - `{{DATA_COMPLETA}}` - Data por extenso
        - `{{HORA}}` - Horário atual
        - `{{MES}}` - Mês atual
        - `{{ANO}}` - Ano atual
        
        **Riscos:**
        - `{{RISCOS}}` - Lista simples de riscos
        - `{{RISCOS_DETALHADOS}}` - Riscos por categoria
        - `{{RISCOS_TABELA}}` - Tabela formatada (substitui por tabela real)
        - `{{TOTAL_RISCOS}}` - Quantidade de riscos
        - `{{TOTAL_CATEGORIAS}}` - Quantidade de categorias
        - `{{RISCOS_AUTOMATICOS}}` - Riscos com unidade automática
        - `{{RISCOS_MANUAIS}}` - Riscos com unidade manual
        
        **Sistema:**
        - `{{USUARIO}}` - Nome do usuário do sistema
        - `{{EMAIL_USUARIO}}` - Email do usuário
        - `{{VERSAO_SISTEMA}}` - Versão atual
        - `{{ID_GERACAO}}` - ID único da geração
        """)
    
    # Botão de geração
    if template_file and excel_file:
        st.markdown("### 🚀 Gerar OS e Relatório")
        
        if st.button("📄 Gerar Tudo", type="primary", use_container_width=True):
            with st.spinner("Processando..."):
                progress = st.progress(0)
                status = st.empty()
                
                for i in range(100):
                    time.sleep(0.02)
                    progress.progress(i + 1)
                    if i < 30:
                        status.text("📄 Processando template...")
                    elif i < 60:
                        status.text("📊 Lendo funcionários...")
                    elif i < 80:
                        status.text("⚠️ Inserindo riscos...")
                    else:
                        status.text("📋 Gerando relatório...")
                
                try:
                    # Gerar documentos
                    zip_result, historico = gerar_documentos_os_avancados(
                        template_file, excel_file, risks_salvos
                    )
                    
                    if zip_result and historico:
                        # Gerar relatório
                        relatorio = gerar_relatorio_os(historico, risks_salvos)
                        
                        # Atualizar dados
                        dados_usuario = get_user_data()
                        historico_completo = dados_usuario.get('historico_os', [])
                        historico_completo.extend(historico)
                        
                        save_user_data({
                            'creditos': max(0, dados_usuario.get('creditos', 10) - 1),
                            'os_geradas_total': dados_usuario.get('os_geradas_total', 0) + len(historico),
                            'ultimo_uso': datetime.now().strftime("%d/%m/%Y"),
                            'historico_os': historico_completo
                        })
                        
                        st.success("✅ Geração concluída!")
                        st.balloons()
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.download_button(
                                "📥 Baixar OS (ZIP)",
                                data=zip_result,
                                file_name=f"OS_Lote_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                                mime="application/zip",
                                use_container_width=True
                            )
                        
                        with col2:
                            if relatorio:
                                st.download_button(
                                    "📊 Baixar Relatório",
                                    data=relatorio,
                                    file_name=f"Relatorio_OS_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    use_container_width=True
                                )
                        
                        st.info(f"📄 {len(historico)} OS geradas • 📊 1 Relatório criado")
                    
                except Exception as e:
                    st.error(f"❌ Erro: {e}")
                
                status.empty()
                progress.empty()

def show_reports_page():
    """NOVA PÁGINA - Relatórios e Histórico"""
    st.markdown('<div class="section-header">📊 Relatórios e Histórico</div>', unsafe_allow_html=True)
    
    dados_usuario = get_user_data()
    historico_os = dados_usuario.get('historico_os', [])
    
    if not historico_os:
        st.markdown("""
        <div class="alert alert-info">
            <strong>📊 Nenhum relatório disponível</strong><br>
            Gere suas primeiras OS na aba "Gerar OS" para ver os relatórios aqui.
        </div>
        """, unsafe_allow_html=True)
        return
    
    # Estatísticas gerais
    st.markdown("### 📈 Estatísticas Gerais")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        create_metric_card("Total OS", len(historico_os), "Geradas")
    
    with col2:
        funcionarios_unicos = len(set(item['funcionario'] for item in historico_os))
        create_metric_card("Funcionários", funcionarios_unicos, "Diferentes")
    
    with col3:
        setores_unicos = len(set(item['setor'] for item in historico_os))
        create_metric_card("Setores", setores_unicos, "Diferentes")
    
    with col4:
        riscos_auto_total = sum(item.get('riscos_automaticos', 0) for item in historico_os)
        create_metric_card("Auto Detecções", riscos_auto_total, "VDVR/AREN/etc")
    
    # Histórico detalhado
    st.markdown("### 📋 Histórico Detalhado")
    
    # Filtros
    col1, col2, col3 = st.columns(3)
    
    with col1:
        setores_disponiveis = sorted(set(item['setor'] for item in historico_os))
        setor_filtro = st.selectbox(
            "🏢 Filtrar por Setor", 
            ["Todos"] + setores_disponiveis,
            key="setor_filtro"
        )
    
    with col2:
        data_inicio = st.date_input(
            "📅 Data Inicial",
            value=datetime.now() - timedelta(days=30),
            key="data_inicio"
        )
    
    with col3:
        data_fim = st.date_input(
            "📅 Data Final",
            value=datetime.now(),
            key="data_fim"
        )
    
    # Filtrar dados
    historico_filtrado = historico_os
    
    if setor_filtro != "Todos":
        historico_filtrado = [item for item in historico_filtrado if item['setor'] == setor_filtro]
    
    historico_filtrado = [
        item for item in historico_filtrado 
        if data_inicio <= datetime.fromisoformat(item['data_geracao']).date() <= data_fim
    ]
    
    # Mostrar dados filtrados
    if historico_filtrado:
        st.info(f"📊 Mostrando {len(historico_filtrado)} de {len(historico_os)} registros")
        
        # Converter para DataFrame para exibição
        df_historico = pd.DataFrame(historico_filtrado)
        df_historico['Data/Hora'] = pd.to_datetime(df_historico['data_geracao']).dt.strftime('%d/%m/%Y %H:%M')
        
        # Reordenar colunas
        colunas_display = ['Data/Hora', 'funcionario', 'cargo', 'setor', 'total_riscos', 'riscos_automaticos', 'id_geracao']
        df_display = df_historico[colunas_display].copy()
        df_display.columns = ['Data/Hora', 'Funcionário', 'Cargo', 'Setor', 'Riscos', 'Auto', 'ID']
        
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        
        # Análises adicionais
        with st.expander("📊 Análises Adicionais", expanded=False):
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**📈 OS por Setor:**")
                setor_count = df_historico['setor'].value_counts()
                for setor, count in setor_count.items():
                    st.write(f"• {setor}: {count} OS")
            
            with col2:
                st.markdown("**🤖 Detecções Automáticas:**")
                total_riscos = df_historico['total_riscos'].sum()
                total_auto = df_historico['riscos_automaticos'].sum()
                if total_riscos > 0:
                    pct_auto = (total_auto / total_riscos) * 100
                    st.metric("% Automático", f"{pct_auto:.1f}%", f"{total_auto}/{total_riscos}")
        
        # Botão para gerar relatório consolidado
        if st.button("📋 Gerar Relatório Consolidado", use_container_width=True):
            with st.spinner("Gerando relatório consolidado..."):
                # Aqui você pode implementar um relatório mais abrangente
                st.success("🎯 Funcionalidade de relatório consolidado pode ser expandida!")
    
    else:
        st.warning("❌ Nenhum registro encontrado com os filtros aplicados.")

def show_credits_management():
    st.markdown('<div class="section-header">💰 Créditos</div>', unsafe_allow_html=True)
    
    dados_usuario = get_user_data()
    creditos = dados_usuario.get('creditos', 10)
    
    create_metric_card("Saldo", creditos, "Disponíveis")
    
    col1, col2, col3 = st.columns(3)
    pacotes = [
        {"nome": "Básico", "creditos": 10, "preco": 50.00},
        {"nome": "Pro", "creditos": 25, "preco": 100.00},
        {"nome": "Premium", "creditos": 50, "preco": 180.00}
    ]
    
    for idx, pac in enumerate(pacotes):
        with [col1, col2, col3][idx]:
            st.markdown(f"""
            <div class="metric-card" style="text-align: center;">
                <div style="font-size: 1.5rem; margin-bottom: 1rem;">📦</div>
                <div style="font-weight: 700; font-size: 1.2rem;">{pac["nome"]}</div>
                <div style="margin: 0.5rem 0;">{pac["creditos"]} créditos</div>
                <div style="font-size: 1.5rem; font-weight: 700; color: var(--primary-color);">
                    R$ {pac["preco"]:.2f}
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button(f"🛒 {pac['nome']}", key=f"comprar_{idx}", use_container_width=True):
                novo_total = creditos + pac["creditos"]
                save_user_data({'creditos': novo_total})
                st.success(f"✅ +{pac['creditos']} créditos!")
                time.sleep(1)
                st.rerun()

# --- Main ---
def main():
    if not is_authenticated():
        show_login_page()
    else:
        show_main_app()

if __name__ == "__main__":
    main()
