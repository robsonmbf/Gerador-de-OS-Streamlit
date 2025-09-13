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

# --- Inicialização dos Gerenciadores ---
@st.cache_resource
def init_managers():
    """Inicializa os gerenciadores de banco de dados"""
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
    /* --- CORREÇÃO DE ESTILO DA BARRA DE USUÁRIO --- */
    .user-info {
        background-color: #262730; /* Cor de fundo cinza escuro, combinando com o tema */
        color: white;             /* Cor do texto para branco */
        padding: 1rem;
        border-radius: 5px;
        margin-bottom: 1rem;
        border: 1px solid #3DD56D; /* Borda verde para destaque */
    }
    .activity-log {
        background-color: #f0f0f0;
        padding: 1rem;
        border-radius: 5px;
        max-height: 300px;
        overflow-y: auto;
    }
</style>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE AUTENTICAÇÃO ---

def show_login_page():
    """Exibe a página de login/registro"""
    st.markdown("""<div class="main-header"><h1>🔐 Acesso ao Sistema</h1><p>Faça login ou registre-se para acessar o Gerador de OS</p></div>""", unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["Login", "Registro"])
    
    with tab1:
        st.markdown('<div class="auth-container">', unsafe_allow_html=True)
        st.subheader("Fazer Login")
        
        with st.form("login_form"):
            email = st.text_input("Email", placeholder="seu@email.com")
            password = st.text_input("Senha", type="password")
            submit_login = st.form_submit_button("Entrar", use_container_width=True)
            
            if submit_login:
                if email and password:
                    success, message, session_data = auth_manager.login_user(email, password)
                    if success:
                        st.session_state.authenticated = True
                        st.session_state.user_data = session_data
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)
                else:
                    st.error("Por favor, preencha todos os campos")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with tab2:
        st.markdown('<div class="auth-container">', unsafe_allow_html=True)
        st.subheader("Criar Conta")
        
        with st.form("register_form"):
            reg_email = st.text_input("Email", placeholder="seu@email.com", key="reg_email")
            reg_password = st.text_input("Senha", type="password", key="reg_password")
            reg_password_confirm = st.text_input("Confirmar Senha", type="password")
            submit_register = st.form_submit_button("Registrar", use_container_width=True)
            
            if submit_register:
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
        
        st.markdown('</div>', unsafe_allow_html=True)

def check_authentication():
    """Verifica se o usuário está autenticado"""
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    
    if 'user_data' not in st.session_state:
        st.session_state.user_data = None
    
    # Verificar sessão se houver token
    if st.session_state.authenticated and st.session_state.user_data:
        session_token = st.session_state.user_data.get('session_token')
        if session_token:
            is_valid, session_info = auth_manager.validate_session(session_token)
            if not is_valid:
                st.session_state.authenticated = False
                st.session_state.user_data = None
                st.rerun()

def logout_user():
    """Faz logout do usuário"""
    if st.session_state.user_data and st.session_state.user_data.get('session_token'):
        auth_manager.logout_user(st.session_state.user_data['session_token'])
    
    st.session_state.authenticated = False
    st.session_state.user_data = None
    st.rerun()

def show_user_info():
    """Exibe informações do usuário logado"""
    if st.session_state.user_data:
        user_email = st.session_state.user_data.get('email', 'N/A')
        user_id = st.session_state.user_data.get('user_id')
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.markdown(f'<div class="user-info">👤 <strong>Usuário:</strong> {user_email}</div>', unsafe_allow_html=True)
        
        with col2:
            if st.button("Sair", type="secondary"):
                logout_user()

# --- INICIALIZAÇÃO DO SESSION STATE PARA DADOS DO USUÁRIO ---
def init_user_session_state():
    """Inicializa o session state com dados do usuário do banco de dados"""
    if not st.session_state.authenticated or not st.session_state.user_data:
        return
    
    user_id = st.session_state.user_data.get('user_id')
    if not user_id:
        return
    
    # Carregar dados do usuário do banco de dados
    user_summary = user_data_manager.get_user_summary(user_id)
    
    # Inicializar listas no session_state se não existirem
    if 'medicoes_adicionadas' not in st.session_state:
        # Converter medições do banco para o formato esperado pelo app
        st.session_state.medicoes_adicionadas = []
        for measurement in user_summary['measurements']:
            st.session_state.medicoes_adicionadas.append({
                'agente': measurement['agent'],
                'valor': measurement['value'],
                'unidade': measurement['unit'],
                'epi': measurement['epi'] or ''
            })
    
    if 'epis_adicionados' not in st.session_state:
        # Converter EPIs do banco para lista de strings
        st.session_state.epis_adicionados = [epi['epi_name'] for epi in user_summary['epis']]
    
    if 'riscos_manuais_adicionados' not in st.session_state:
        # Converter riscos manuais do banco para o formato esperado
        st.session_state.riscos_manuais_adicionados = []
        for risk in user_summary['manual_risks']:
            st.session_state.riscos_manuais_adicionados.append({
                'categoria': risk['category'],
                'risco': risk['risk_name'],
                'danos': risk['possible_damages'] or ''
            })
    
    if 'setores_concluidos' not in st.session_state:
        st.session_state.setores_concluidos = set()
    
    if 'cargos_concluidos' not in st.session_state:
        st.session_state.cargos_concluidos = set()

# --- FUNÇÕES PARA SINCRONIZAR COM BANCO DE DADOS ---

def sync_measurement_to_db(measurement):
    """Sincroniza uma medição com o banco de dados"""
    if not st.session_state.authenticated:
        return
    
    user_id = st.session_state.user_data.get('user_id')
    if user_id:
        user_data_manager.add_measurement(
            user_id,
            measurement['agente'],
            measurement['valor'],
            measurement['unidade'],
            measurement['epi'] if measurement['epi'] else None
        )

def sync_epi_to_db(epi_name):
    """Sincroniza um EPI com o banco de dados"""
    if not st.session_state.authenticated:
        return
    
    user_id = st.session_state.user_data.get('user_id')
    if user_id:
        user_data_manager.add_epi(user_id, epi_name)

def sync_manual_risk_to_db(risk):
    """Sincroniza um risco manual com o banco de dados"""
    if not st.session_state.authenticated:
        return
    
    user_id = st.session_state.user_data.get('user_id')
    if user_id:
        user_data_manager.add_manual_risk(
            user_id,
            risk['categoria'],
            risk['risco'],
            risk['danos'] if risk['danos'] else None
        )

# --- LISTAS DE ДАНЫ CONSTANTES (mantidas do código original) ---
UNIDADES_DE_MEDIDA = ["dB(A)", "m/s²", "ppm", "mg/m³", "%", "°C", "lx", "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"]
AGENTES_DE_RISCO = sorted([
    "Ruído (Contínuo ou Intermitente)", "Ruído (Impacto)", "Vibração de Corpo Inteiro", "Vibração de Mãos e Braços",
    "Radiações Ionizantes", "Radiações Não-Ionizantes", "Frio", "Calor", "Pressões Anormais", "Umidade", "Poeiras", 
    "Fumos", "Névoas", "Neblinas", "Gases", "Vapores", "Produtos Químicos em Geral", "Vírus", "Bactérias", 
    "Protozoários", "Fungos", "Parasitas", "Bacilos"
])
CATEGORIAS_RISCO = {'fisico': '🔥 Físicos', 'quimico': '⚗️ Químicos', 'biologico': '🦠 Biológicos', 'ergonomico': '🏃 Ergonômicos', 'acidente': '⚠️ Acidentes'}

# --- Funções de Lógica de Negócio (mantidas do código original) ---

def normalizar_texto(texto):
    """Remove acentos, espaços e caracteres especiais para comparação de strings."""
    if not isinstance(texto, str): return ""
    texto = texto.lower().strip()
    texto = re.sub(r'[\s\W_]+', '', texto) 
    return texto

def mapear_e_renomear_colunas_funcionarios(df):
    """Renomeia as colunas da planilha de funcionários para um padrão conhecido."""
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
    """Carrega e armazena em cache a planilha para evitar recarregamentos."""
    if arquivo is None: return None
    try:
        return pd.read_excel(arquivo)
    except Exception as e:
        st.error(f"Erro ao ler o ficheiro Excel: {e}")
        return None

def obter_dados_pgr():
    """Simula a obtenção de dados de um PGR. Em um caso real, isso viria de um banco de dados ou outra planilha."""
    data = [
        # Riscos Físicos
        {'categoria': 'fisico', 'risco': 'Ruído (Contínuo ou Intermitente)', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'},
        {'categoria': 'fisico', 'risco': 'Ruído (Impacto)', 'possiveis_danos': 'Perda auditiva, trauma acústico.'},
        {'categoria': 'fisico', 'risco': 'Vibração de Corpo Inteiro', 'possiveis_danos': 'Problemas na coluna, dores lombares.'},
        {'categoria': 'fisico', 'risco': 'Vibração de Mãos e Braços', 'possiveis_danos': 'Doenças osteomusculares, problemas circulatórios (síndrome de Raynaud).'},
        {'categoria': 'fisico', 'risco': 'Calor', 'possiveis_danos': 'Desidratação, insolação, cãibras, exaustão, intermação.'},
        {'categoria': 'fisico', 'risco': 'Frio', 'possiveis_danos': 'Hipotermia, congelamento, doenças respiratórias.'},
        {'categoria': 'fisico', 'risco': 'Radiações Ionizantes', 'possiveis_danos': 'Câncer, mutações genéticas, queimaduras.'},
        {'categoria': 'fisico', 'risco': 'Radiações Não-Ionizantes', 'possiveis_danos': 'Queimaduras, lesões oculares, câncer de pele.'},
        {'categoria': 'fisico', 'risco': 'Pressões Anormais', 'possiveis_danos': 'Doença descompressiva, barotrauma.'},
        {'categoria': 'fisico', 'risco': 'Umidade', 'possiveis_danos': 'Doenças respiratórias, dermatites, micoses.'},

        # Riscos Químicos
        {'categoria': 'quimico', 'risco': 'Poeiras', 'possiveis_danos': 'Pneumoconioses (silicose, asbestose), irritação respiratória, alergias.'},
        {'categoria': 'quimico', 'risco': 'Fumos', 'possiveis_danos': 'Doenças respiratórias (febre dos fumos metálicos), intoxicações.'},
        {'categoria': 'quimico', 'risco': 'Névoas', 'possiveis_danos': 'Irritação respiratória, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Neblinas', 'possiveis_danos': 'Irritação do trato respiratório.'},
        {'categoria': 'quimico', 'risco': 'Gases', 'possiveis_danos': 'Asfixia, intoxicações, irritação respiratória.'},
        {'categoria': 'quimico', 'risco': 'Vapores', 'possiveis_danos': 'Irritação respiratória, intoxicações, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Produtos Químicos em Geral', 'possiveis_danos': 'Queimaduras, irritações, intoxicações, dermatites, câncer.'},

        # Riscos Biológicos
        {'categoria': 'biologico', 'risco': 'Bactérias', 'possiveis_danos': 'Infecções, doenças infecciosas (tétano, tuberculose).'},
        {'categoria': 'biologico', 'risco': 'Fungos', 'possiveis_danos': 'Micoses, alergias, infecções respiratórias.'},
        {'categoria': 'biologico', 'risco': 'Vírus', 'possiveis_danos': 'Doenças virais (hepatite, HIV), infecções.'},
        {'categoria': 'biologico', 'risco': 'Parasitas', 'possiveis_danos': 'Doenças parasitárias, infecções.'},
        {'categoria': 'biologico', 'risco': 'Protozoários', 'possiveis_danos': 'Doenças parasitárias (leishmaniose, malária).'},
        {'categoria': 'biologico', 'risco': 'Bacilos', 'possiveis_danos': 'Infecções diversas, como tuberculose.'},
        
        # Riscos Ergonômicos
        {'categoria': 'ergonomico', 'risco': 'Levantamento e Transporte Manual de Peso', 'possiveis_danos': 'Lesões musculoesqueléticas, dores na coluna.'},
        {'categoria': 'ergonomico', 'risco': 'Posturas Inadequadas', 'possiveis_danos': 'Dores musculares, lesões na coluna, LER/DORT.'},
        {'categoria': 'ergonomico', 'risco': 'Repetitividade', 'possiveis_danos': 'LER/DORT, tendinites, síndrome do túnel do carpo.'},
        {'categoria': 'ergonomico', 'risco': 'Jornada de Trabalho Prolongada', 'possiveis_danos': 'Fadiga, estresse, acidentes de trabalho.'},
        {'categoria': 'ergonomico', 'risco': 'Iluminação Inadequada', 'possiveis_danos': 'Fadiga visual, dores de cabeça, acidentes.'},

        # Riscos de Acidentes
        {'categoria': 'acidente', 'risco': 'Arranjo Físico Inadequado', 'possiveis_danos': 'Quedas, colisões, esmagamentos.'},
        {'categoria': 'acidente', 'risco': 'Máquinas e Equipamentos sem Proteção', 'possiveis_danos': 'Amputações, cortes, esmagamentos, prensamentos.'},
        {'categoria': 'acidente', 'risco': 'Ferramentas Inadequadas ou Defeituosas', 'possiveis_danos': 'Cortes, perfurações, fraturas.'},
        {'categoria': 'acidente', 'risco': 'Eletricidade', 'possiveis_danos': 'Choque elétrico, queimaduras, fibrilação ventricular.'},
        {'categoria': 'acidente', 'risco': 'Incêndio e Explosão', 'possiveis_danos': 'Queimaduras, asfixia, lesões por impacto.'},
        {'categoria': 'acidente', 'risco': 'Animais Peçonhentos', 'possiveis_danos': 'Picadas, mordidas, reações alérgicas, envenenamento.'},
        {'categoria': 'acidente', 'risco': 'Armazenamento Inadequado', 'possiveis_danos': 'Quedas de materiais, esmagamentos.'},
        {'categoria': 'acidente', 'risco': 'Trabalho em Altura', 'possiveis_danos': 'Quedas, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Espaços Confinados', 'possiveis_danos': 'Asfixia, intoxicações, explosões.'},
        {'categoria': 'acidente', 'risco': 'Projeção de Partículas', 'possiveis_danos': 'Lesões oculares, cortes na pele.'}
    ]
    return pd.DataFrame(data)

# --- Continuação das funções do código original (substituir_placeholders, gerar_os) ---
# [O resto das funções permanece igual ao código original]

def substituir_placeholders(doc, contexto):
    """Substitui chaves de texto (ex: [NOME]) no documento Word pelos valores do contexto."""
    # Itera por parágrafos, tabelas e cabeçalhos para uma substituição completa.
    all_elements = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_elements.extend(cell.paragraphs)
    for section in doc.sections:
        header = section.header
        all_elements.extend(header.paragraphs)
        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_elements.extend(cell.paragraphs)

    for p in all_elements:
        # Constrói o texto completo do parágrafo a partir de seus 'runs'
        inline_text = "".join(r.text for r in p.runs)
        
        # Pula para o próximo parágrafo se não houver placeholders nele
        if not any(key in inline_text for key in contexto):
            continue

        original_text = inline_text
        
        # Realiza as substituições no texto em memória
        for key, value in contexto.items():
            inline_text = inline_text.replace(str(key), str(value))

        # Se o texto foi alterado, reescreve o parágrafo tratando as quebras de linha
        if original_text != inline_text:
            p.text = ""  # Limpa todo o conteúdo do parágrafo (todos os 'runs')
            
            lines = inline_text.split('\n')
            for i, line in enumerate(lines):
                if i > 0:
                    p.add_run().add_break()  # Adiciona uma quebra de linha suave (Shift+Enter)
                
                # Adiciona o texto da linha atual
                run = p.add_run(line)
                # Reaplica a formatação padrão, garantindo que não fique em negrito
                font = run.font
                font.name = 'Segoe UI'
                font.size = Pt(9)
                font.bold = False # Garante que o texto não será negrito

def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, riscos_manuais, modelo_doc_carregado):
    """Função principal que gera um documento de Ordem de Serviço para um funcionário."""
    doc = Document(modelo_doc_carregado)
    
    # Processa riscos do PGR
    riscos_info = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]
    riscos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}
    danos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}
    for _, risco_row in riscos_info.iterrows():
        categoria = str(risco_row.get("categoria", "")).lower()
        if categoria in riscos_por_categoria:
            riscos_por_categoria[categoria].append(str(risco_row.get("risco", "")))
            danos = risco_row.get("possiveis_danos")
            if pd.notna(danos): danos_por_categoria[categoria].append(str(danos))

    # Adiciona riscos inseridos manualmente
    if riscos_manuais:
        map_categorias_rev = {v: k for k, v in CATEGORIAS_RISCO.items()}
        for risco_manual in riscos_manuais:
            categoria_display = risco_manual.get('categoria')
            categoria_alvo = map_categorias_rev.get(categoria_display)
            if categoria_alvo:
                riscos_por_categoria[categoria_alvo].append(risco_manual.get('risco', ''))
                if risco_manual.get('danos'):
                    danos_por_categoria[categoria_alvo].append(risco_manual.get('danos'))

    # Formata a lista de medições, incluindo o EPI associado
    medicoes_formatadas = []
    for med in medicoes_manuais:
        epi_info = f" | EPI: {med['epi']}" if med.get("epi") else ""
        medicoes_formatadas.append(f"{med['agente']}: {med['valor']} {med['unidade']}{epi_info}")
    medicoes_texto = "\n".join(medicoes_formatadas) if medicoes_formatadas else "Não aplicável"

    # Trata e formata campos do funcionário
    data_admissao = "Não informado"
    if 'data_de_admissao' in funcionario and pd.notna(funcionario['data_de_admissao']):
        try: data_admissao = pd.to_datetime(funcionario['data_de_admissao']).strftime('%d/%m/%Y')
        except Exception: data_admissao = str(funcionario['data_de_admissao'])

    descricao_atividades = "Não informado"
    if 'descricao_de_atividades' in funcionario and pd.notna(funcionario['descricao_de_atividades']):
        descricao_atividades = str(funcionario['descricao_de_atividades'])
    
    def tratar_lista_vazia(lista, separador=", "):
        if not lista or all(not item.strip() for item in lista): return "Não identificado"
        return separador.join(sorted(list(set(item for item in lista if item and item.strip()))))

    # Cria o dicionário de contexto para substituição no Word
    contexto = {
        "[NOME EMPRESA]": str(funcionario.get("empresa", "N/A")), 
        "[UNIDADE]": str(funcionario.get("unidade", "N/A")),
        "[NOME FUNCIONÁRIO]": str(funcionario.get("nome_do_funcionario", "N/A")), 
        "[DATA DE ADMISSÃO]": data_admissao,
        "[SETOR]": str(funcionario.get("setor", "N/A")), 
        "[FUNÇÃO]": str(funcionario.get("funcao", "N/A")),
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
        "[EPIS]": tratar_lista_vazia(epis_manuais.split(',')) or "Não aplicável",
        "[MEDIÇÕES]": medicoes_texto,
    }
    
    substituir_placeholders(doc, contexto)
    return doc

# --- APLICAÇÃO PRINCIPAL ---

def main():
    """Função principal da aplicação"""
    # Verificar autenticação
    check_authentication()
    
    if not st.session_state.authenticated:
        show_login_page()
        return
    
    # Usuário autenticado - mostrar aplicação principal
    if st.session_state.authenticated:
        show_user_info()
    
    init_user_session_state()
    
    # Log da atividade de acesso
    user_id = st.session_state.user_data.get('user_id')
    if user_id:
        db_manager.log_activity(user_id, 'app_access')
    
    # Interface principal (código original adaptado)
    st.markdown("""<div class="main-header"><h1>📄 Gerador de Ordens de Serviço (OS)</h1><p>Gere OS em lote a partir de um modelo Word (.docx) e uma planilha de funcionários.</p></div>""", unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown("##### 📂 1. Carregue os Documentos")
        col1, col2 = st.columns(2)
        with col1:
            arquivo_funcionarios = st.file_uploader("📄 **Planilha de Funcionários (.xlsx)**", type="xlsx", help="Planilha com colunas como: Nome, Função, Setor, Empresa, etc.")
        with col2:
            arquivo_modelo_os = st.file_uploader("📝 **Modelo de OS (.docx)**", type="docx", help="Documento Word com placeholders como [NOME FUNCIONÁRIO], [SETOR], etc.")

    if not arquivo_funcionarios or not arquivo_modelo_os:
        st.info("📋 Por favor, carregue a Planilha de Funcionários e o Modelo de OS para continuar.")
        return

    # Log do upload de arquivos
    if user_id:
        db_manager.log_activity(user_id, 'upload_employee_file', {'file_name': arquivo_funcionarios.name})
        db_manager.log_activity(user_id, 'upload_os_template', {'file_name': arquivo_modelo_os.name})

    # Carrega os dados após os uploads
    df_funcionarios_raw = carregar_planilha(arquivo_funcionarios)
    if df_funcionarios_raw is None:
        st.stop()

    df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_raw)
    df_pgr = obter_dados_pgr()

    # Resto da interface (adaptada do código original)
    # [Continuar com o resto da implementação...]

if __name__ == "__main__":
    main()
