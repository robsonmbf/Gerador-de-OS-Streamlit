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

# --- MOCKUP PARA TESTE SEM BANCO DE DADOS ---
# Classes para simular o comportamento do banco de dados e permitir
# que o script rode de forma independente.
class MockDBManager:
    def get_connection(self): return None

class MockAuthManager:
    def __init__(self, db_manager): pass
    def login_user(self, email, password): return True, "Login bem-sucedido!", {'user_id': 'test_user', 'email': email, 'session_token': 'fake_token'}
    def register_user(self, email, password): return True, "Registro bem-sucedido!"
    def validate_session(self, token): return True, "Sessão válida"
    def logout_user(self, token): pass

class MockUserDataManager:
    def __init__(self, db_manager):
        if 'mock_data' not in st.session_state:
            st.session_state.mock_data = {
                'measurements': [],
                'epis': [],
                'manual_risks': []
            }
    def get_user_measurements(self, user_id): return st.session_state.mock_data['measurements']
    def add_measurement(self, user_id, agent, value, unit, epi): st.session_state.mock_data['measurements'].append({'id': time.time(), 'agent': agent, 'value': value, 'unit': unit, 'epi': epi})
    def remove_measurement(self, user_id, med_id): st.session_state.mock_data['measurements'] = [m for m in st.session_state.mock_data['measurements'] if m['id'] != med_id]
    def get_user_epis(self, user_id): return st.session_state.mock_data['epis']
    def add_epi(self, user_id, epi_name): st.session_state.mock_data['epis'].append({'id': time.time(), 'epi_name': epi_name})
    def remove_epi(self, user_id, epi_id): st.session_state.mock_data['epis'] = [e for e in st.session_state.mock_data['epis'] if e['id'] != epi_id]
    def get_user_manual_risks(self, user_id): return st.session_state.mock_data['manual_risks']
    def add_manual_risk(self, user_id, category, risk_name, damages): st.session_state.mock_data['manual_risks'].append({'id': time.time(), 'category': category, 'risk_name': risk_name, 'possible_damages': damages})
    def remove_manual_risk(self, user_id, risk_id): st.session_state.mock_data['manual_risks'] = [r for r in st.session_state.mock_data['manual_risks'] if r['id'] != risk_id]

# --- FIM DO MOCKUP ---

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
    # Para usar seu banco de dados real, comente as 3 linhas do Mock
    # e descomente as 4 linhas seguintes.
    db_manager = MockDBManager()
    auth_manager = MockAuthManager(db_manager)
    user_data_manager = MockUserDataManager(db_manager)
    
    # from database.models import DatabaseManager
    # from database.auth import AuthManager
    # from database.user_data import UserDataManager
    # db_manager = DatabaseManager()
    # auth_manager = AuthManager(db_manager)
    # user_data_manager = UserDataManager(db_manager)
    return db_manager, auth_manager, user_data_manager

db_manager, auth_manager, user_data_manager = init_managers()

# --- CSS PERSONALIZADO ---
st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none; }
    .main-header { text-align: center; padding-bottom: 20px; }
    .auth-container { max-width: 400px; margin: 0 auto; padding: 2rem; border: 1px solid #ddd; border-radius: 10px; background-color: #f9f9f9; }
    .user-info { background-color: #262730; color: white; padding: 1rem; border-radius: 5px; margin-bottom: 1rem; border: 1px solid #3DD56D; }
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
                    else: st.error(message)
                else: st.error("Por favor, preencha todos os campos")
    with tab2:
        with st.form("register_form"):
            reg_email = st.text_input("Email", placeholder="seu@email.com", key="reg_email")
            reg_password = st.text_input("Senha", type="password", key="reg_password")
            reg_password_confirm = st.text_input("Confirmar Senha", type="password")
            if st.form_submit_button("Registrar", use_container_width=True):
                if reg_email and reg_password and reg_password_confirm:
                    if reg_password != reg_password_confirm: st.error("As senhas não coincidem")
                    else:
                        success, message = auth_manager.register_user(reg_email, reg_password)
                        if success:
                            st.success(message)
                            st.info("Agora você pode fazer login com suas credenciais")
                        else: st.error(message)
                else: st.error("Por favor, preencha todos os campos")

def check_authentication():
    if 'authenticated' not in st.session_state: st.session_state.authenticated = False
    if 'user_data' not in st.session_state: st.session_state.user_data = None
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
        with col1: st.markdown(f'<div class="user-info">👤 <strong>Usuário:</strong> {user_email}</div>', unsafe_allow_html=True)
        with col2:
            if st.button("Sair", type="secondary"): logout_user()

def init_user_session_state():
    if st.session_state.get('authenticated') and not st.session_state.get('user_data_loaded'):
        user_id = st.session_state.user_data.get('user_id')
        if user_id:
            st.session_state.medicoes_adicionadas = user_data_manager.get_user_measurements(user_id)
            st.session_state.epis_adicionados = user_data_manager.get_user_epis(user_id)
            st.session_state.riscos_manuais_adicionados = user_data_manager.get_user_manual_risks(user_id)
            st.session_state.user_data_loaded = True
    if 'medicoes_adicionadas' not in st.session_state: st.session_state.medicoes_adicionadas = []
    if 'epis_adicionados' not in st.session_state: st.session_state.epis_adicionados = []
    if 'riscos_manuais_adicionados' not in st.session_state: st.session_state.riscos_manuais_adicionados = []
    if 'cargos_concluidos' not in st.session_state: st.session_state.cargos_concluidos = set()

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
    try: return pd.read_excel(arquivo)
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

def substituir_placeholders_com_logica_medicoes(doc, contexto):
    """
    Substitui placeholders em um documento Word, com lógica especial para
    a chave '[MEDIÇÕES]' para evitar problemas de espaçamento.
    """
    # Processa placeholders em tabelas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # Itera sobre uma cópia da lista de parágrafos para modificá-la com segurança
                for p in list(cell.paragraphs):
                    # Lógica especial para o placeholder de medições
                    if '[MEDIÇÕES]' in p.text:
                        valor_medicoes = contexto.get('[MEDIÇÕES]')
                        p.text = ""  # Limpa o parágrafo do placeholder

                        if isinstance(valor_medicoes, list) and valor_medicoes:
                            for i, linha_medicao in enumerate(valor_medicoes):
                                # Usa o parágrafo existente para a primeira linha, cria novos para as demais
                                par_atual = p if i == 0 else cell.add_paragraph()
                                run = par_atual.add_run(linha_medicao)
                                run.font.name = 'Segoe UI'
                                run.font.size = Pt(9)
                                par_atual.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        else:  # Se for uma string (ex: "Não aplicável")
                            run = p.add_run(str(valor_medicoes))
                            run.font.name = 'Segoe UI'
                            run.font.size = Pt(9)
                    
                    # Lógica normal para outros placeholders
                    try:
                        inline_text = "".join(run.text for run in p.runs)
                        for key, value in contexto.items():
                            if key != '[MEDIÇÕES]' and key in inline_text:
                                new_text = inline_text.replace(key, str(value))
                                p.clear()
                                run = p.add_run(new_text)
                                run.font.name = 'Segoe UI'
                                run.font.size = Pt(9)
                                inline_text = new_text  # Atualiza para a próxima iteração
                    except Exception:
                        continue

    # Processa placeholders em parágrafos fora de tabelas
    for p in doc.paragraphs:
        inline_text = "".join(run.text for run in p.runs)
        for key, value in contexto.items():
            if key != '[MEDIÇÕES]' and key in inline_text:
                new_text = inline_text.replace(key, str(value))
                p.clear()
                run = p.add_run(new_text)
                run.font.name = 'Segoe UI'
                run.font.size = Pt(9)
                inline_text = new_text

def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, riscos_manuais, modelo_doc_carregado):
    doc = Document(modelo_doc_carregado)
    riscos_info = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]
    riscos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}
    danos_por_categoria = {cat: [] for cat in CATEGORias_RISCO.keys()}
    for _, risco_row in riscos_info.iterrows():
        categoria = str(risco_row.get("categoria", "")).lower()
        if categoria in riscos_por_categoria:
            riscos_por_categoria[categoria].append(str(risco_row.get("risco", "")))
            danos = risco_row.get("possiveis_danos")
            if pd.notna(danos): danos_por_categoria[categoria].append(str(danos))
    if riscos_manuais:
        map_categorias_rev = {v: k for k, v in CATEGORIAS_RISCO.items()}
        for risco_manual in riscos_manuais:
            categoria_display = risco_manual.get('category')
            categoria_alvo = map_categorias_rev.get(categoria_display)
            if categoria_alvo:
                riscos_por_categoria[categoria_alvo].append(risco_manual.get('risk_name', ''))
                if risco_manual.get('possible_damages'): danos_por_categoria[categoria_alvo].append(risco_manual.get('possible_damages'))
    for cat in danos_por_categoria: danos_por_categoria[cat] = sorted(list(set(danos_por_categoria[cat])))

    medicoes_ordenadas = sorted(medicoes_manuais, key=lambda med: med.get('agent', ''))
    medicoes_formatadas = [f"{med.get('agent', 'N/A')}: {med.get('value', 'N/A')} {med.get('unit', '')}" for med in medicoes_ordenadas]

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
        "[EPIS]": tratar_lista_vazia([epi['epi_name'] for epi in epis_manuais]),
        "[MEDIÇÕES]": medicoes_formatadas if medicoes_formatadas else "Não aplicável",
    }

    substituir_placeholders_com_logica_medicoes(doc, contexto)
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
        with col1: arquivo_funcionarios = st.file_uploader("📄 **Planilha de Funcionários (.xlsx)**", type="xlsx")
        with col2: arquivo_modelo_os = st.file_uploader("📝 **Modelo de OS (.docx)**", type="docx")

    if not arquivo_funcionarios or not arquivo_modelo_os:
        st.info("📋 Por favor, carregue a Planilha de Funcionários e o Modelo de OS para continuar.")
        return

    df_funcionarios_raw = carregar_planilha(arquivo_funcionarios)
    if df_funcionarios_raw is None: st.stop()

    df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_raw)
    df_pgr = obter_dados_pgr()

    with st.container(border=True):
        st.markdown('##### 👥 2. Selecione os Funcionários')
        setores = sorted(df_funcionarios['setor'].dropna().unique().tolist()) if 'setor' in df_funcionarios.columns else []
        setor_sel = st.multiselect("Filtrar por Setor(es)", setores)
        df_filtrado_setor = df_funcionarios[df_funcionarios['setor'].isin(setor_sel)] if setor_sel else df_funcionarios
        st.caption(f"{len(df_filtrado_setor)} funcionário(s) no(s) setor(es) selecionado(s).")
        funcoes_disponiveis = sorted(df_filtrado_setor['funcao'].dropna().unique().tolist()) if 'funcao' in df_filtrado_setor.columns else []
        funcoes_formatadas = [f"{f} ✅ Concluído" if setor_sel and all((s, f) in st.session_state.cargos_concluidos for s in setor_sel) else f for f in funcoes_disponiveis]
        funcao_sel_formatada = st.multiselect("Filtrar por Função/Cargo(s)", funcoes_formatadas)
        funcao_sel = [f.replace(" ✅ Concluído", "") for f in funcao_sel_formatada]
        df_final_filtrado = df_filtrado_setor[df_filtrado_setor['funcao'].isin(funcao_sel)] if funcao_sel else df_filtrado_setor
        st.success(f"**{len(df_final_filtrado)} funcionário(s) selecionado(s) para gerar OS.**")
        if not df_final_filtrado.empty: st.dataframe(df_final_filtrado[['nome_do_funcionario', 'setor', 'funcao']])

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
                    categoria_key_series = df_pgr[df_pgr['risco'] == risco_nome
