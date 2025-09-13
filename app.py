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

# --- FUNÇÕES DE AUTENTICAÇÃO E LÓGICA DE NEGÓCIO ---
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
    # ... (código da função inalterado) ...
    return df

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
    # ... (código da função com lista completa de riscos inalterado) ...
    return pd.DataFrame(data)

# --- INÍCIO DA ALTERAÇÃO 1: FUNÇÃO DE SUBSTITUIÇÃO SIMPLIFICADA E ROBUSTA ---
def substituir_placeholders(doc, contexto):
    """
    Substitui placeholders de forma simples e direta em parágrafos e tabelas.
    Esta abordagem é mais robusta contra erros de formatação complexa.
    """
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    # Combina o texto de todos os 'runs' para ter o texto completo
                    texto_completo_paragrafo = "".join(run.text for run in p.runs)
                    for key, value in contexto.items():
                        if key in texto_completo_paragrafo:
                            # Faz a substituição no texto completo
                            novo_texto = texto_completo_paragrafo.replace(key, str(value))
                            # Limpa o parágrafo e adiciona o novo texto
                            p.clear()
                            run = p.add_run(novo_texto)
                            # Garante a formatação correta para todo o run
                            run.font.name = 'Segoe UI'
                            run.font.size = Pt(9)
                            # Atualiza o texto base para futuras substituições no mesmo parágrafo
                            texto_completo_paragrafo = novo_texto

    for p in doc.paragraphs:
        # Mesma lógica para parágrafos fora de tabelas
        texto_completo_paragrafo = "".join(run.text for run in p.runs)
        for key, value in contexto.items():
            if key in texto_completo_paragrafo:
                novo_texto = texto_completo_paragrafo.replace(key, str(value))
                p.clear()
                run = p.add_run(novo_texto)
                run.font.name = 'Segoe UI'
                run.font.size = Pt(9)
                texto_completo_paragrafo = novo_texto
# --- FIM DA ALTERAÇÃO 1 ---


def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, riscos_manuais, modelo_doc_carregado):
    doc = Document(modelo_doc_carregado)
    # ... (lógica de processamento de riscos inalterada) ...
    
    # --- INÍCIO DA ALTERAÇÃO 2: FORMATAÇÃO DE MEDIÇÕES EM ESTILO DE TABELA ---
    medicoes_ordenadas = sorted(medicoes_manuais, key=lambda med: med.get('agent', ''))
    
    medicoes_formatadas = []
    # Encontra o texto mais longo para alinhar
    max_len_agente = 0
    max_len_valor = 0
    if medicoes_ordenadas:
        max_len_agente = max(len(med.get('agent', '')) for med in medicoes_ordenadas)
        max_len_valor = max(len(med.get('value', '')) for med in medicoes_ordenadas)

    for med in medicoes_ordenadas:
        agente = med.get('agent', 'N/A')
        valor = med.get('value', 'N/A')
        unidade = med.get('unit', '')
        
        # Adiciona espaços para alinhar as colunas
        padding_agente = ' ' * (max_len_agente - len(agente))
        padding_valor = ' ' * (max_len_valor - len(valor))
        
        # Usa tabulação para criar um espaçamento consistente
        linha = f"{agente}:{padding_agente}\t{valor}{padding_valor}\t{unidade}"
        medicoes_formatadas.append(linha)
    
    medicoes_texto = "\n".join(medicoes_formatadas) if medicoes_formatadas else "Não aplicável"
    # --- FIM DA ALTERAÇÃO 2 ---

    # ... (resto da função gerar_os e criação do contexto inalterado) ...
    contexto = {
        # ... (todos os outros placeholders) ...
        "[MEDIÇÕES]": medicoes_texto,
    }
    
    substituir_placeholders(doc, contexto)
    return doc


# --- APLICAÇÃO PRINCIPAL ---
def main():
    # ... (código da aplicação principal inalterado) ...
    pass

if __name__ == "__main__":
    main()
