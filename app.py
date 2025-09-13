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

# --- INÍCIO DA ALTERAÇÃO 1: FUNÇÃO DE SUBSTITUIÇÃO MELHORADA ---
def substituir_placeholders(doc, contexto):
    """Substitui placeholders em parágrafos e tabelas, preservando melhor a formatação."""
    # Processa parágrafos
    for p in doc.paragraphs:
        # Pula parágrafos vazios
        if not p.text.strip():
            continue
            
        # Junta o texto de todos os 'runs' para ter o texto completo do parágrafo
        texto_completo_paragrafo = "".join(run.text for run in p.runs)
        
        for key, value in contexto.items():
            if key in texto_completo_paragrafo:
                # Armazena a formatação original do primeiro run do parágrafo
                # para reaplicar nas partes que não são o valor do placeholder
                primeiro_run = p.runs[0]
                formato_original = {
                    'bold': primeiro_run.bold,
                    'italic': primeiro_run.italic,
                    'underline': primeiro_run.underline,
                    'font_name': primeiro_run.font.name,
                    'font_size': primeiro_run.font.size
                }

                # Limpa o parágrafo
                p.clear()

                # Divide o texto em partes antes e depois do placeholder
                partes = texto_completo_paragrafo.split(key)

                for i, parte in enumerate(partes):
                    # Adiciona a parte do texto antes/depois do placeholder com a formatação original
                    run_parte = p.add_run(parte)
                    run_parte.bold = formato_original['bold']
                    run_parte.italic = formato_original['italic']
                    run_parte.underline = formato_original['underline']
                    if formato_original['font_name']:
                        run_parte.font.name = formato_original['font_name']
                    if formato_original['font_size']:
                        run_parte.font.size = formato_original['font_size']
                    
                    # Adiciona o valor do placeholder (com formatação padrão), exceto para a última parte
                    if i < len(partes) - 1:
                        linhas_valor = str(value).split('\n')
                        for j, linha in enumerate(linhas_valor):
                            run_valor = p.add_run(linha)
                            run_valor.bold = False
                            run_valor.italic = False
                            run_valor.underline = False
                            run_valor.font.name = 'Segoe UI'
                            run_valor.font.size = Pt(9)
                            if j < len(linhas_valor) - 1:
                                run_valor.add_break()
                
                # Atualiza o texto completo para a próxima iteração no mesmo parágrafo
                texto_completo_paragrafo = "".join(run.text for run in p.runs)
    # Processa tabelas da mesma forma
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if not p.text.strip():
                        continue
                    texto_completo_paragrafo = "".join(run.text for run in p.runs)
                    for key, value in contexto.items():
                        if key in texto_completo_paragrafo:
                            primeiro_run = p.runs[0]
                            formato_original = {
                                'bold': primeiro_run.bold, 'italic': primeiro_run.italic, 'underline': primeiro_run.underline,
                                'font_name': primeiro_run.font.name, 'font_size': primeiro_run.font.size
                            }
                            p.clear()
                            partes = texto_completo_paragrafo.split(key)
                            for i, parte in enumerate(partes):
                                run_parte = p.add_run(parte)
                                run_parte.bold = formato_original['bold']
                                run_parte.italic = formato_original['italic']
                                run_parte.underline = formato_original['underline']
                                if formato_original['font_name']: run_parte.font.name = formato_original['font_name']
                                if formato_original['font_size']: run_parte.font.size = formato_original['font_size']
                                if i < len(partes) - 1:
                                    linhas_valor = str(value).split('\n')
                                    for j, linha in enumerate(linhas_valor):
                                        run_valor = p.add_run(linha)
                                        run_valor.bold = False; run_valor.italic = False; run_valor.underline = False
                                        run_valor.font.name = 'Segoe UI'; run_valor.font.size = Pt(9)
                                        if j < len(linhas_valor) - 1:
                                            run_valor.add_break()
                            texto_completo_paragrafo = "".join(run.text for run in p.runs)
# --- FIM DA ALTERAÇÃO 1 ---

def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, riscos_manuais, modelo_doc_carregado):
    doc = Document(modelo_doc_carregado)
    riscos_info = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]
    riscos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}
    danos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}
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
                if risco_manual.get('possible_damages'):
                    danos_por_categoria[categoria_alvo].append(risco_manual.get('possible_damages'))
    for cat in danos_por_categoria:
        danos_por_categoria[cat] = sorted(list(set(danos_por_categoria[cat])))
        
    medicoes_ordenadas = sorted(medicoes_manuais, key=lambda med: med.get('agent', ''))
    
    medicoes_formatadas = []
    # --- INÍCIO DA ALTERAÇÃO 2: ALINHAMENTO CONSISTENTE DAS MEDIÇÕES ---
    for med in medicoes_ordenadas:
        agente = med.get('agent', 'N/A')
        valor = med.get('value', 'N/A')
        unidade = med.get('unit', '')
        epi = med.get('epi', '')
        
        epi_info = f" | EPI: {epi}" if epi and epi.strip() else ""
        # Adiciona tabulação para um alinhamento consistente
        medicoes_formatadas.append(f"{agente}:\t{valor} {unidade}{epi_info}")
    # --- FIM DA ALTERAÇÃO 2 ---

    medicoes_texto = "\n".join(medicoes_formatadas) if medicoes_formatadas else "Não aplicável"
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
        "[MEDIÇÕES]": medicoes_texto,
    }
    substituir_placeholders(doc, contexto)
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
        with col1:
            arquivo_funcionarios = st.file_uploader("📄 **Planilha de Funcionários (.xlsx)**", type="xlsx")
        with col2:
            arquivo_modelo_os = st.file_uploader("📝 **Modelo de OS (.docx)**", type="docx")

    if not arquivo_funcionarios or not arquivo_modelo_os:
        st.info("📋 Por favor, carregue a Planilha de Funcionários e o Modelo de OS para continuar.")
        return
    
    df_funcionarios_raw = carregar_planilha(arquivo_funcionarios)
    if df_funcionarios_raw is None:
        st.stop()

    df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_raw)
    df_pgr = obter_dados_pgr()

    with st.container(border=True):
        st.markdown('##### 👥 2. Selecione os Funcionários')
        setores = sorted(df_funcionarios['setor'].dropna().unique().tolist()) if 'setor' in df_funcionarios.columns else []
        setor_sel = st.multiselect("Filtrar por Setor(es)", setores)
        df_filtrado_setor = df_funcionarios[df_funcionarios['setor'].isin(setor_sel)] if setor_sel else df_funcionarios
        st.caption(f"{len(df_filtrado_setor)} funcionário(s) no(s) setor(es) selecionado(s).")
        funcoes_disponiveis = sorted(df_filtrado_setor['funcao'].dropna().unique().tolist()) if 'funcao' in df_filtrado_setor.columns else []
        funcoes_formatadas = []
        if setor_sel:
            for funcao in funcoes_disponiveis:
                concluido = all((s, funcao) in st.session_state.cargos_concluidos for s in setor_sel)
                if concluido:
                    funcoes_formatadas.append(f"{funcao} ✅ Concluído")
                else:
                    funcoes_formatadas.append(funcao)
        else:
            funcoes_formatadas = funcoes_disponiveis
        funcao_sel_formatada = st.multiselect("Filtrar por Função/Cargo(s)", funcoes_formatadas)
        funcao_sel = [f.replace(" ✅ Concluído", "") for f in funcao_sel_formatada]
        df_final_filtrado = df_filtrado_setor[df_filtrado_setor['funcao'].isin(funcao_sel)] if funcao_sel else df_filtrado_setor
        st.success(f"**{len(df_final_filtrado)} funcionário(s) selecionado(s) para gerar OS.**")
        st.dataframe(df_final_filtrado[['nome_do_funcionario', 'setor', 'funcao']])

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

        col_exp1, col_exp2, col_exp3 = st.columns(3)
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
            with st.expander("➕ **Adicionar Risco Manual (Alternativo)**"):
                 st.info("Para adicionar riscos manuais, por favor, use a aba '➕ Manual' na seção de seleção de riscos acima.")
        with col_exp3:
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
                    arquivo_modelo_os
                )
                doc_io = BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                nome_limpo = re.sub(r'[^\w\s-]', '', func.get("nome_do_funcionario", "Func_Sem_Nome")).strip().replace(" ", "_")
                documentos_gerados.append((f"OS_{nome_limpo}.docx", doc_io.getvalue()))
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
