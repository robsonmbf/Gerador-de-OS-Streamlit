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
UNIDADES_DE_MEDIDA = ["dB(A)", "m/s²", "m/s¹⁷⁵", "ppm", "mg/m³", "%", "°C", "lx", "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"]
AGENTES_DE_RISCO = sorted([
    "Ruído (Contínuo ou Intermitente)", "Ruído (Impacto)", "Vibração de Corpo Inteiro", "Vibração de Mãos e Braços",
    "Radiações Ionizantes", "Radiações Não-Ionizantes", "Frio", "Calor", "Pressões Anormais", "Umidade", "Poeiras", 
    "Fumos", "Névoas", "Neblinas", "Gases", "Vapores", "Produtos Químicos em Geral", "Vírus", "Bactérias", 
    "Protozoários", "Fungos", "Parasitas", "Bacilos"
])
CATEGORIAS_RISCO = {"fisico": "🔥 Físicos", "quimico": "⚗️ Químicos", "biologico": "🦠 Biológicos", "ergonomico": "🏃 Ergonômicos", "acidente": "⚠️ Acidentes"}

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
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "user_data" not in st.session_state:
        st.session_state.user_data = None
    if st.session_state.authenticated and st.session_state.user_data:
        session_token = st.session_state.user_data.get("session_token")
        if session_token:
            is_valid, _ = auth_manager.validate_session(session_token)
            if not is_valid:
                st.session_state.authenticated = False
                st.session_state.user_data = None
                st.rerun()

def logout_user():
    if st.session_state.user_data and st.session_state.user_data.get("session_token"):
        auth_manager.logout_user(st.session_state.user_data["session_token"])
    st.session_state.authenticated = False
    st.session_state.user_data = None
    st.session_state.user_data_loaded = False
    st.rerun()

def show_user_info():
    if st.session_state.get("authenticated"):
        user_email = st.session_state.user_data.get("email", "N/A")
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown(f'''<div class="user-info">👤 <strong>Usuário:</strong> {user_email}</div>''', unsafe_allow_html=True)
        with col2:
            if st.button("Sair", type="secondary"):
                logout_user()

def init_user_session_state():
    if st.session_state.get("authenticated") and not st.session_state.get("user_data_loaded"):
        user_id = st.session_state.user_data.get("user_id")
        if user_id:
            st.session_state.medicoes_adicionadas = user_data_manager.get_user_measurements(user_id)
            st.session_state.epis_adicionados = user_data_manager.get_user_epis(user_id)
            st.session_state.riscos_manuais_adicionados = user_data_manager.get_user_manual_risks(user_id)
            st.session_state.user_data_loaded = True
    
    if "medicoes_adicionadas" not in st.session_state:
        st.session_state.medicoes_adicionadas = []
    if "epis_adicionados" not in st.session_state:
        st.session_state.epis_adicionados = []
    if "riscos_manuais_adicionados" not in st.session_state:
        st.session_state.riscos_manuais_adicionados = []
    if "cargos_concluidos" not in st.session_state:
        st.session_state.cargos_concluidos = set()

def normalizar_texto(texto):
    if not isinstance(texto, str): return ""
    return re.sub(r"[\s\W_]+", "", texto.lower().strip())

def mapear_e_renomear_colunas_funcionarios(df):
    df_copia = df.copy()
    mapeamento = {
        "nome_do_funcionario": ["nomedofuncionario", "nome", "funcionario", "funcionário", "colaborador", "nomecompleto"],
        "funcao": ["funcao", "função", "cargo"],
        "data_de_admissao": ["datadeadmissao", "dataadmissao", "admissao", "admissão"],
        "setor": ["setordetrabalho", "setor", "area", "área", "departamento"],
        "descricao_de_atividades": ["descricaodeatividades", "atividades", "descricaoatividades", "descriçãodeatividades", "tarefas", "descricaodastarefas"],
        "empresa": ["empresa"],
        "unidade": ["unidade"]
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
    # Dados existentes
    data_existente = [
        {"categoria": "fisico", "risco": "Ruído (Contínuo ou Intermitente)", "possiveis_danos": "Perda auditiva, zumbido, estresse, irritabilidade."},
        {"categoria": "fisico", "risco": "Ruído (Impacto)", "possiveis_danos": "Perda auditiva, trauma acústico."},
        {"categoria": "fisico", "risco": "Vibração de Corpo Inteiro", "possiveis_danos": "Problemas na coluna, dores lombares."},
        {"categoria": "fisico", "risco": "Vibração de Mãos e Braços", "possiveis_danos": "Doenças osteomusculares, problemas circulatórios."},
        {"categoria": "fisico", "risco": "Calor", "possiveis_danos": "Desidratação, insolação, cãibras, exaustão, intermação."},
        {"categoria": "fisico", "risco": "Frio", "possiveis_danos": "Hipotermia, congelamento, doenças respiratórias."},
        {"categoria": "fisico", "risco": "Radiações Ionizantes", "possiveis_danos": "Câncer, mutações genéticas, queimaduras."},
        {"categoria": "fisico", "risco": "Radiações Não-Ionizantes", "possiveis_danos": "Queimaduras, lesões oculares, câncer de pele."},
        {"categoria": "fisico", "risco": "Pressões Anormais", "possiveis_danos": "Doença descompressiva, barotrauma."},
        {"categoria": "fisico", "risco": "Umidade", "possiveis_danos": "Doenças respiratórias, dermatites, micoses."},
        {"categoria": "quimico", "risco": "Poeiras", "possiveis_danos": "Pneumoconioses (silicose, asbestose), irritação respiratória."},
        {"categoria": "quimico", "risco": "Fumos", "possiveis_danos": "Doenças respiratórias (febre dos fumos metálicos), intoxicações."},
        {"categoria": "quimico", "risco": "Névoas", "possiveis_danos": "Irritação respiratória, dermatites."},
        {"categoria": "quimico", "risco": "Gases", "possiveis_danos": "Asfixia, intoxicações, irritação respiratória."},
        {"categoria": "quimico", "risco": "Vapores", "possiveis_danos": "Irritação respiratória, intoxicações, dermatites."},
        {"categoria": "quimico", "risco": "Produtos Químicos em Geral", "possiveis_danos": "Queimaduras, irritações, intoxicações, dermatites, câncer."},
        {"categoria": "biologico", "risco": "Bactérias", "possiveis_danos": "Infecções, doenças infecciosas (tétano, tuberculose)."},
        {"categoria": "biologico", "risco": "Fungos", "possiveis_danos": "Micoses, alergias, infecções respiratórias."},
        {"categoria": "biologico", "risco": "Vírus", "possiveis_danos": "Doenças virais (hepatite, HIV), infecções."},
        {"categoria": "ergonomico", "risco": "Levantamento e Transporte Manual de Peso", "possiveis_danos": "Lesões musculoesqueléticas, dores na coluna."},
        {"categoria": "ergonomico", "risco": "Posturas Inadequadas", "possiveis_danos": "Dores musculares, lesões na coluna, LER/DORT."},
        {"categoria": "ergonomico", "risco": "Repetitividade", "possiveis_danos": "LER/DORT, tendinites, síndrome do túnel do carpo."},
        {"categoria": "acidente", "risco": "Máquinas e Equipamentos sem Proteção", "possiveis_danos": "Amputações, cortes, esmagamentos, prensamentos."},
        {"categoria": "acidente", "risco": "Eletricidade", "possiveis_danos": "Choque elétrico, queimaduras, fibrilação ventricular."},
        {"categoria": "acidente", "risco": "Trabalho em Altura", "possiveis_danos": "Quedas, fraturas, morte."},
        {"categoria": "acidente", "risco": "Projeção de Partículas", "possiveis_danos": "Lesões oculares, cortes na pele."}
    ]
    # Carregar dados da planilha
    try:
        df_excel = pd.read_excel("/home/ubuntu/upload/PerigoseRiscosPGR.xlsx")
        df_excel.rename(columns={
            "Categoria": "categoria",
            "Perigo  (Fator de Risco/Agente Nocivo/Situação Perigosa)": "risco",
            "Possíveis Danos ou Agravos à Saúde": "possiveis_danos"
        }, inplace=True)
        df_excel["categoria"] = df_excel["categoria"].str.lower()
        data_excel = df_excel.to_dict(orient="records")
        data_existente.extend(data_excel)
    except FileNotFoundError:
        # Se o arquivo não for encontrado, apenas os dados existentes são usados
        pass
    return pd.DataFrame(data_existente)

def substituir_placeholders(doc, contexto):
    """
    Substitui placeholders preservando a formatação do template.
    """
    def aplicar_formatacao_padrao(run):
        """Aplica formatação Segoe UI 9pt"""
        run.font.name = "Segoe UI"
        run.font.size = Pt(9)
        return run

    def processar_paragrafo(p):
        texto_original_paragrafo = p.text

        # --- Lógica CORRIGIDA E MANTIDA para [MEDIÇÕES] ---
        if "[MEDIÇÕES]" in texto_original_paragrafo:
            for run in p.runs:
                run.text = ""
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            medicoes_valor = contexto.get("[MEDIÇÕES]", "Não aplicável")
            if medicoes_valor == "Não aplicável" or not medicoes_valor.strip():
                run = aplicar_formatacao_padrao(p.add_run("Não aplicável"))
                run.font.bold = False
            else:
                linhas = medicoes_valor.split("\n")
                for i, linha in enumerate(linhas):
                    if not linha.strip(): continue
                    if i > 0: p.add_run().add_break()
                    if ":" in linha:
                        partes = linha.split(":", 1)
                        agente_texto = partes[0].strip() + ":"
                        valor_texto = partes[1].strip()
                        run_agente = aplicar_formatacao_padrao(p.add_run(agente_texto + " "))
                        run_agente.font.bold = True
                        run_valor = aplicar_formatacao_padrao(p.add_run(valor_texto))
                        run_valor.font.bold = False
                    else:
                        run_simples = aplicar_formatacao_padrao(p.add_run(linha))
                        run_simples.font.bold = False
            return

        # --- Lógica RESTAURADA E CORRIGIDA para outros placeholders ---
        placeholders_no_paragrafo = [key for key in contexto if key in texto_original_paragrafo]
        if not placeholders_no_paragrafo:
            return

        # Preserva o estilo do primeiro "run", que geralmente define o estilo do rótulo no template
        estilo_rotulo = {
            "bold": p.runs[0].bold if p.runs else False,
            "italic": p.runs[0].italic if p.runs else False,
            "underline": p.runs[0].underline if p.runs else False,
        }

        # Substitui todos os placeholders para obter o texto final
        texto_final = texto_original_paragrafo
        for key in placeholders_no_paragrafo:
            texto_final = texto_final.replace(key, str(contexto[key]))
        
        # Limpa o parágrafo para reescrevê-lo com a formatação correta
        p.clear()

        # Reconstrói o parágrafo, aplicando o estilo do rótulo e deixando os valores sem formatação
        texto_restante = texto_final
        for i, key in enumerate(placeholders_no_paragrafo):
            valor_placeholder = str(contexto[key])
            partes = texto_restante.split(valor_placeholder, 1)
            
            # Adiciona o texto antes do valor (que é o rótulo) com o estilo preservado
            if partes[0]:
                run_rotulo = aplicar_formatacao_padrao(p.add_run(partes[0]))
                run_rotulo.font.bold = estilo_rotulo["bold"]
                run_rotulo.font.italic = estilo_rotulo["italic"]
                run_rotulo.underline = estilo_rotulo["underline"]

            # Adiciona o valor do placeholder sem formatação
            run_valor = aplicar_formatacao_padrao(p.add_run(valor_placeholder))
            run_valor.font.bold = False
            run_valor.font.italic = False
            run_valor.font.underline = False
            
            texto_restante = partes[1]

        # Adiciona qualquer texto que sobrar no final, com o estilo do rótulo
        if texto_restante:
            run_final = aplicar_formatacao_padrao(p.add_run(texto_restante))
            run_final.font.bold = estilo_rotulo["bold"]
            run_final.font.italic = estilo_rotulo["italic"]
            run_final.underline = estilo_rotulo["underline"]

    # Processar parágrafos em tabelas e no corpo do documento
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    processar_paragrafo(p)
    for p in doc.paragraphs:
        processar_paragrafo(p)


def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_selecionados, medicoes_selecionadas, data_elaboracao, local_data, template_path="/home/ubuntu/upload/TEMPLATES DE OS.docx"):
    try:
        doc = Document(template_path)
    except Exception as e:
        st.error(f"Erro ao carregar o template: {e}")
        return None

    riscos_formatados = ""
    if not riscos_selecionados:
        riscos_formatados = "Não identificado"
    else:
        for i, r in enumerate(riscos_selecionados):
            riscos_formatados += f"{r["risco"]} - {r["danos"]}"
            if i < len(riscos_selecionados) - 1:
                riscos_formatados += "\n"

    epis_formatados = ""
    if not epis_selecionados:
        epis_formatados = "Não necessário"
    else:
        for i, epi in enumerate(epis_selecionados):
            epis_formatados += epi
            if i < len(epis_selecionados) - 1:
                epis_formatados += ", "

    medicoes_formatadas = ""
    if not medicoes_selecionadas:
        medicoes_formatadas = "Não aplicável"
    else:
        for i, medicao in enumerate(medicoes_selecionadas):
            medicoes_formatadas += f"{medicao["agente"]}: {medicao["valor"]} {medicao["unidade"]}"
            if i < len(medicoes_selecionadas) - 1:
                medicoes_formatadas += "\n"

    contexto = {
        "[EMPRESA]": funcionario.get("empresa", ""),
        "[UNIDADE]": funcionario.get("unidade", ""),
        "[SETOR]": funcionario.get("setor", ""),
        "[NOME]": funcionario.get("nome_do_funcionario", ""),
        "[FUNÇÃO]": funcionario.get("funcao", ""),
        "[ADMISSÃO]": funcionario.get("data_de_admissao", ""),
        "[DESCRIÇÃO DAS ATIVIDADES]": funcionario.get("descricao_de_atividades", ""),
        "[RISCOS]": riscos_formatados,
        "[EPI]": epis_formatados,
        "[MEDIÇÕES]": medicoes_formatadas,
        "[DATA_ELABORACAO]": data_elaboracao,
        "[LOCAL_DATA]": local_data,
    }

    substituir_placeholders(doc, contexto)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def main():
    check_authentication()

    if not st.session_state.get("authenticated"):
        show_login_page()
        return

    show_user_info()
    init_user_session_state()

    st.title("Gerador de Ordens de Serviço (OS) 📄")

    uploaded_file = st.file_uploader("Carregue a planilha de funcionários (Excel)", type=["xlsx"])
    df_funcionarios = carregar_planilha(uploaded_file)

    if df_funcionarios is not None:
        df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios)
        lista_funcionarios = df_funcionarios["nome_do_funcionario"].tolist()
        funcionario_selecionado_nome = st.selectbox("Selecione o Funcionário", lista_funcionarios)
        funcionario_selecionado = df_funcionarios[df_funcionarios["nome_do_funcionario"] == funcionario_selecionado_nome].iloc[0].to_dict()

        st.subheader(f"Informações de {funcionario_selecionado_nome}")
        st.write(f"**Função:** {funcionario_selecionado.get("funcao", "N/A")}")
        st.write(f"**Setor:** {funcionario_selecionado.get("setor", "N/A")}")
        st.write(f"**Descrição das Atividades:** {funcionario_selecionado.get("descricao_de_atividades", "N/A")}")

        df_pgr = obter_dados_pgr()
        riscos_disponiveis = df_pgr["risco"].tolist()

        st.subheader("Seleção de Riscos e Medidas")
        
        # --- SELEÇÃO DE RISCOS DO PGR ---
        with st.expander("Adicionar Riscos do PGR"):
            col1, col2 = st.columns(2)
            with col1:
                categoria_selecionada = st.selectbox("Filtrar por Categoria", ["Todas"] + list(CATEGORIAS_RISCO.values()))
            
            riscos_filtrados = df_pgr
            if categoria_selecionada != "Todas":
                categoria_key = [key for key, value in CATEGORIAS_RISCO.items() if value == categoria_selecionada][0]
                riscos_filtrados = df_pgr[df_pgr["categoria"] == categoria_key]

            with col2:
                risco_selecionado_pgr = st.selectbox("Selecione o Risco", riscos_filtrados["risco"].tolist())

            if st.button("Adicionar Risco do PGR"):
                if risco_selecionado_pgr:
                    dano_associado = df_pgr[df_pgr["risco"] == risco_selecionado_pgr]["possiveis_danos"].iloc[0]
                    st.session_state.riscos_manuais_adicionados.append({"risco": risco_selecionado_pgr, "danos": dano_associado})
                    user_data_manager.save_user_manual_risks(st.session_state.user_data["user_id"], st.session_state.riscos_manuais_adicionados)

        # --- ADIÇÃO DE RISCOS MANUAIS ---
        with st.expander("Adicionar Risco Manualmente"):
            with st.form("manual_risk_form"):
                risco_manual = st.text_input("Descrição do Risco")
                danos_manuais = st.text_area("Possíveis Danos")
                if st.form_submit_button("Adicionar Risco Manual"):
                    if risco_manual and danos_manuais:
                        st.session_state.riscos_manuais_adicionados.append({"risco": risco_manual, "danos": danos_manuais})
                        user_data_manager.save_user_manual_risks(st.session_state.user_data["user_id"], st.session_state.riscos_manuais_adicionados)

        # --- EXIBIÇÃO E REMOÇÃO DE RISCOS ---
        if st.session_state.riscos_manuais_adicionados:
            st.write("**Riscos Adicionados:**")
            for i, risco in enumerate(st.session_state.riscos_manuais_adicionados):
                col1, col2, col3 = st.columns([3, 4, 1])
                with col1:
                    st.write(risco["risco"])
                with col2:
                    st.write(risco["danos"])
                with col3:
                    if st.button(f"Remover##{i}", key=f"rem_risk_{i}"):
                        st.session_state.riscos_manuais_adicionados.pop(i)
                        user_data_manager.save_user_manual_risks(st.session_state.user_data["user_id"], st.session_state.riscos_manuais_adicionados)
                        st.rerun()

        # --- SELEÇÃO DE EPIs ---
        with st.expander("Adicionar EPIs"):
            with st.form("epi_form") :
                novo_epi = st.text_input("Nome do EPI")
                if st.form_submit_button("Adicionar EPI"):
                    if novo_epi and novo_epi not in st.session_state.epis_adicionados:
                        st.session_state.epis_adicionados.append(novo_epi)
                        user_data_manager.save_user_epis(st.session_state.user_data["user_id"], st.session_state.epis_adicionados)

        if st.session_state.epis_adicionados:
            st.write("**EPIs Adicionados:**")
            epis_selecionados_geracao = st.multiselect("Selecione os EPIs para esta OS", st.session_state.epis_adicionados, default=st.session_state.epis_adicionados)

        # --- SELEÇÃO DE MEDIÇÕES ---
        with st.expander("Adicionar Medições de Agentes"):
            with st.form("medicao_form"):
                agente = st.selectbox("Agente de Risco", AGENTES_DE_RISCO)
                valor = st.number_input("Valor da Medição", format="%.4f")
                unidade = st.selectbox("Unidade de Medida", UNIDADES_DE_MEDIDA)
                if st.form_submit_button("Adicionar Medição"):
                    st.session_state.medicoes_adicionadas.append({"agente": agente, "valor": valor, "unidade": unidade})
                    user_data_manager.save_user_measurements(st.session_state.user_data["user_id"], st.session_state.medicoes_adicionadas)

        if st.session_state.medicoes_adicionadas:
            st.write("**Medições Adicionadas:**")
            medicoes_para_remover = []
            for i, medicao in enumerate(st.session_state.medicoes_adicionadas):
                col1, col2, col3, col4 = st.columns([3, 2, 2, 1])
                with col1:
                    st.write(medicao["agente"])
                with col2:
                    st.write(medicao["valor"])
                with col3:
                    st.write(medicao["unidade"])
                with col4:
                    if st.button(f"Remover##{i}", key=f"rem_med_{i}"):
                        medicoes_para_remover.append(i)
            
            if medicoes_para_remover:
                st.session_state.medicoes_adicionadas = [m for i, m in enumerate(st.session_state.medicoes_adicionadas) if i not in medicoes_para_remover]
                user_data_manager.save_user_measurements(st.session_state.user_data["user_id"], st.session_state.medicoes_adicionadas)
                st.rerun()

        st.subheader("Geração da Ordem de Serviço")
        data_elaboracao = st.date_input("Data de Elaboração", pd.to_datetime("today"))
        local_data = st.text_input("Local e Data para Assinatura", "Cidade, DD de Mês de AAAA")

        if st.button("Gerar Ordem de Serviço"):
            with st.spinner("Gerando documento..."):
                os_bytes = gerar_os(
                    funcionario_selecionado,
                    df_pgr,
                    st.session_state.riscos_manuais_adicionados,
                    epis_selecionados_geracao,
                    st.session_state.medicoes_adicionadas,
                    data_elaboracao.strftime("%d/%m/%Y"),
                    local_data
                )
                if os_bytes:
                    st.success("Ordem de Serviço gerada com sucesso!")
                    st.download_button(
                        label="Baixar Ordem de Serviço",
                        data=os_bytes,
                        file_name=f"OS_{funcionario_selecionado_nome}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

if __name__ == "__main__":
    main()


