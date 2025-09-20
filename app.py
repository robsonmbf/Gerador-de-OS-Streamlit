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

def substituir_placeholders(doc, contexto):
    """
    Substitui placeholders com ALINHAMENTO CORRETO.
    Força alinhamento à esquerda para medições.
    """
    from docx.shared import Pt
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    # Obter texto completo
                    full_text = p.text
                    texto_modificado = full_text

                    # Substituir placeholders
                    for key, value in contexto.items():
                        if key in texto_modificado:
                            texto_modificado = texto_modificado.replace(key, str(value))

                    # Se mudou, recriar o parágrafo
                    if texto_modificado != full_text:
                        # CORREÇÃO ESPECIAL PARA MEDIÇÕES
                        if "[MEDIÇÕES]" in full_text:
                            # Forçar alinhamento à esquerda
                            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                            print(f"✅ Alinhamento à esquerda aplicado para medições")

                        # Salvar formatação original
                        font_info = None
                        if p.runs:
                            font = p.runs[0].font
                            font_info = {
                                'name': font.name,
                                'size': font.size,
                                'bold': font.bold,
                                'italic': font.italic
                            }

                        # Limpar runs
                        for run in p.runs[:]:
                            p._element.remove(run._element)

                        # Criar novo run
                        new_run = p.add_run(texto_modificado)

                        # Aplicar formatação Segoe UI 9pt
                        new_run.font.name = 'Segoe UI'
                        new_run.font.size = Pt(9)
                        new_run.font.bold = False
                        if font_info and font_info['italic']:
                            new_run.font.italic = font_info['italic']

    # Processar parágrafos fora de tabelas
    for p in doc.paragraphs:
        # Obter texto completo
        full_text = p.text
        texto_modificado = full_text

        # Substituir placeholders
        for key, value in contexto.items():
            if key in texto_modificado:
                texto_modificado = texto_modificado.replace(key, str(value))

        # Se mudou, recriar o parágrafo
        if texto_modificado != full_text:
            # CORREÇÃO ESPECIAL PARA MEDIÇÕES
            if "[MEDIÇÕES]" in full_text:
                # Forçar alinhamento à esquerda
                p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                print(f"✅ Alinhamento à esquerda aplicado para medições")

            # Salvar formatação original
            font_info = None
            if p.runs:
                font = p.runs[0].font
                font_info = {
                    'name': font.name,
                    'size': font.size,
                    'bold': font.bold,
                    'italic': font.italic
                }

            # Limpar runs
            for run in p.runs[:]:
                p._element.remove(run._element)

            # Criar novo run
            new_run = p.add_run(texto_modificado)

            # Aplicar formatação Segoe UI 9pt
            new_run.font.name = 'Segoe UI'
            new_run.font.size = Pt(9)
            new_run.font.bold = False
            if font_info and font_info['italic']:
                new_run.font.italic = font_info['italic']

def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, riscos_manuais, modelo_doc_carregado, ausencia_fator_risco=False):
    """
    Função modificada para incluir a opção 'Ausência de Fator de Risco'
    """
    doc = Document(modelo_doc_carregado)
    
    # Se 'Ausência de Fator de Risco' estiver selecionada, preencher todos os riscos com essa frase
    if ausencia_fator_risco:
        riscos_por_categoria = {cat: ["Ausência de Fator de Risco"] for cat in CATEGORIAS_RISCO.keys()}
        danos_por_categoria = {cat: ["Não aplicável"] for cat in CATEGORIAS_RISCO.keys()}
    else:
        # Lógica original
        riscos_info = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]
        riscos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}
        danos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}

        # Processar riscos selecionados
        for _, risco_row in riscos_info.iterrows():
            categoria = str(risco_row.get("categoria", "")).lower()
            if categoria in riscos_por_categoria:
                riscos_por_categoria[categoria].append(str(risco_row.get("risco", "")))
                danos = risco_row.get("possiveis_danos", "")
                if danos and str(danos).strip():
                    danos_por_categoria[categoria].append(str(danos))

        # Processar riscos manuais
        for risco_manual in riscos_manuais:
            categoria_display = risco_manual.get('category', '')
            categoria_key = None
            for key, value in CATEGORIAS_RISCO.items():
                if value == categoria_display:
                    categoria_key = key
                    break
            if categoria_key:
                riscos_por_categoria[categoria_key].append(risco_manual.get('risk_name', ''))
                danos_manual = risco_manual.get('possible_damages', '')
                if danos_manual and str(danos_manual).strip():
                    danos_por_categoria[categoria_key].append(str(danos_manual))

    # Preparar medições
    medicoes_texto = ""
    if medicoes_manuais:
        medicoes_lista = []
        for med in medicoes_manuais:
            agent = med.get('agent', '')
            value = med.get('value', '')
            unit = med.get('unit', '')
            epi = med.get('associated_epi', '')
            medicao_str = f"{agent}: {value} {unit}"
            if epi:
                medicao_str += f" (EPI: {epi})"
            medicoes_lista.append(medicao_str)
        medicoes_texto = "\n".join(medicoes_lista)
    else:
        medicoes_texto = "Não foram realizadas medições específicas para esta função."

    # Preparar EPIs
    epis_texto = ""
    if epis_manuais:
        epis_lista = [epi.get('epi_name', '') for epi in epis_manuais if epi.get('epi_name')]
        epis_texto = ", ".join(epis_lista)
    else:
        epis_texto = "Conforme análise de risco específica da função."

    # Preparar contexto para substituição
    contexto = {
        "[NOME_FUNCIONARIO]": funcionario.get("nome_do_funcionario", ""),
        "[FUNCAO]": funcionario.get("funcao", ""),
        "[SETOR]": funcionario.get("setor", ""),
        "[DATA_ADMISSAO]": funcionario.get("data_de_admissao", ""),
        "[DESCRICAO_ATIVIDADES]": funcionario.get("descricao_de_atividades", ""),
        "[EMPRESA]": funcionario.get("empresa", ""),
        "[UNIDADE]": funcionario.get("unidade", ""),
        "[MEDIÇÕES]": medicoes_texto,
        "[EPIS]": epis_texto,
    }

    # Adicionar riscos por categoria
    for categoria_key, categoria_nome in CATEGORIAS_RISCO.items():
        riscos_lista = riscos_por_categoria.get(categoria_key, [])
        danos_lista = danos_por_categoria.get(categoria_key, [])
        
        if riscos_lista:
            contexto[f"[RISCOS_{categoria_key.upper()}]"] = "; ".join(riscos_lista)
        else:
            contexto[f"[RISCOS_{categoria_key.upper()}]"] = "Não identificados para esta função"
        
        if danos_lista:
            contexto[f"[DANOS_{categoria_key.upper()}]"] = "; ".join(danos_lista)
        else:
            contexto[f"[DANOS_{categoria_key.upper()}]"] = "Não aplicável"

    # Substituir placeholders
    substituir_placeholders(doc, contexto)
    return doc

def main():
    check_authentication()
    if not st.session_state.authenticated:
        show_login_page()
        return
    
    show_user_info()
    init_user_session_state()
    user_id = st.session_state.user_data.get('user_id')
    
    st.title("📋 Gerador de Ordens de Serviço (OS)")
    st.markdown("---")
    
    # Upload do modelo de OS
    with st.container(border=True):
        st.markdown('##### 📄 1. Carregue o Modelo de OS')
        arquivo_modelo_os = st.file_uploader("Selecione o ficheiro modelo da OS (.docx)", type=["docx"])
        if not arquivo_modelo_os:
            st.warning("⚠️ Por favor, carregue o modelo de OS para continuar.")
            return
        st.success("✅ Modelo de OS carregado com sucesso!")
    
    # Upload da planilha de funcionários
    with st.container(border=True):
        st.markdown('##### 📊 2. Carregue a Planilha de Funcionários')
        arquivo_funcionarios = st.file_uploader("Selecione a planilha de funcionários (.xlsx)", type=["xlsx"])
        if not arquivo_funcionarios:
            st.warning("⚠️ Por favor, carregue a planilha de funcionários para continuar.")
            return
        
        df_funcionarios_original = carregar_planilha(arquivo_funcionarios)
        if df_funcionarios_original is None:
            return
        
        df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_original)
        colunas_obrigatorias = ['nome_do_funcionario', 'funcao', 'setor', 'data_de_admissao', 'descricao_de_atividades', 'empresa', 'unidade']
        colunas_faltando = [col for col in colunas_obrigatorias if col not in df_funcionarios.columns]
        
        if colunas_faltando:
            st.error(f"❌ Colunas obrigatórias em falta: {', '.join(colunas_faltando)}")
            st.info("💡 Certifique-se de que a planilha contém as colunas: Nome do Funcionário, Função, Setor, Data de Admissão, Descrição de Atividades, Empresa, Unidade")
            return
        
        st.success(f"✅ Planilha carregada com sucesso! {len(df_funcionarios)} funcionários encontrados.")
        
        # Filtros
        col1, col2 = st.columns(2)
        with col1:
            setores_unicos = sorted(df_funcionarios['setor'].dropna().unique())
            setores_selecionados = st.multiselect("Filtrar por Setores:", setores_unicos, default=setores_unicos)
        with col2:
            funcoes_unicas = sorted(df_funcionarios['funcao'].dropna().unique())
            funcoes_selecionadas = st.multiselect("Filtrar por Funções:", funcoes_unicas, default=funcoes_unicas)
        
        df_final_filtrado = df_funcionarios[
            (df_funcionarios['setor'].isin(setores_selecionados)) & 
            (df_funcionarios['funcao'].isin(funcoes_selecionadas))
        ]
        
        if df_final_filtrado.empty:
            st.warning("⚠️ Nenhum funcionário corresponde aos filtros selecionados.")
            return
        
        st.info(f"📊 {len(df_final_filtrado)} funcionário(s) selecionado(s) após aplicar os filtros.")
        with st.expander("👀 Visualizar funcionários selecionados"):
            st.dataframe(df_final_filtrado[['nome_do_funcionario', 'setor', 'funcao']], use_container_width=True)

    df_pgr = obter_dados_pgr()
    
    # Configuração de riscos com nova opção
    with st.container(border=True):
        st.markdown('##### ⚠️ 3. Configure os Riscos e Medidas de Controle')
        st.info("Os riscos configurados aqui serão aplicados a TODOS os funcionários selecionados.")
        
        # NOVA FUNCIONALIDADE: Checkbox para "Ausência de Fator de Risco"
        ausencia_fator_risco = st.checkbox(
            "🚫 **Ausência de Fator de Risco**", 
            help="Marque esta opção se não há fatores de risco identificados para os funcionários selecionados. Isso preencherá todos os campos de risco com 'Ausência de Fator de Risco'."
        )
        
        if ausencia_fator_risco:
            st.success("✅ Opção 'Ausência de Fator de Risco' selecionada. Todos os campos de risco serão preenchidos com esta informação.")
            # Desabilitar a seleção de riscos quando "Ausência de Fator de Risco" estiver marcada
            st.info("ℹ️ A seleção de riscos específicos foi desabilitada pois 'Ausência de Fator de Risco' está marcada.")
            riscos_selecionados = []
        else:
            # Lógica original para seleção de riscos
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
                    arquivo_modelo_os,
                    ausencia_fator_risco  # Novo parâmetro
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

