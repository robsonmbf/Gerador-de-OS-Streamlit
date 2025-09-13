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

# --- INÍCIO DA CORREÇÃO: DEFINIÇÃO DE CONSTANTES GLOBAIS ---
# Movendo estas listas para o escopo global para que fiquem acessíveis em todo o app
UNIDADES_DE_MEDIDA = ["dB(A)", "m/s²", "ppm", "mg/m³", "%", "°C", "lx", "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"]
AGENTES_DE_RISCO = sorted([
    "Ruído (Contínuo ou Intermitente)", "Ruído (Impacto)", "Vibração de Corpo Inteiro", "Vibração de Mãos e Braços",
    "Radiações Ionizantes", "Radiações Não-Ionizantes", "Frio", "Calor", "Pressões Anormais", "Umidade", "Poeiras", 
    "Fumos", "Névoas", "Neblinas", "Gases", "Vapores", "Produtos Químicos em Geral", "Vírus", "Bactérias", 
    "Protozoários", "Fungos", "Parasitas", "Bacilos"
])
CATEGORIAS_RISCO = {'fisico': '🔥 Físicos', 'quimico': '⚗️ Químicos', 'biologico': '🦠 Biológicos', 'ergonomico': '🏃 Ergonômicos', 'acidente': '⚠️ Acidentes'}
# --- FIM DA CORREÇÃO ---


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
    .user-info {
        background-color: #262730; 
        color: white;            
        padding: 1rem;
        border-radius: 5px;
        margin-bottom: 1rem;
        border: 1px solid #3DD56D; 
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
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.markdown(f'<div class="user-info">👤 <strong>Usuário:</strong> {user_email}</div>', unsafe_allow_html=True)
        
        with col2:
            if st.button("Sair", type="secondary"):
                logout_user()

def init_user_session_state():
    """Inicializa o session state para os dados do app"""
    if 'medicoes_adicionadas' not in st.session_state:
        st.session_state.medicoes_adicionadas = []
    if 'epis_adicionados' not in st.session_state:
        st.session_state.epis_adicionados = []
    if 'riscos_manuais_adicionados' not in st.session_state:
        st.session_state.riscos_manuais_adicionados = []

# --- FUNÇÕES DE LÓGICA DE NEGÓCIO ---
def normalizar_texto(texto):
    if not isinstance(texto, str): return ""
    texto = texto.lower().strip()
    texto = re.sub(r'[\s\W_]+', '', texto) 
    return texto

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
        {'categoria': 'fisico', 'risco': 'Ruído (Contínuo ou Intermitente)', 'possiveis_danos': 'Perda auditiva, zumbido, estresse.'},
        {'categoria': 'fisico', 'risco': 'Calor', 'possiveis_danos': 'Desidratação, insolação, cãibras.'},
        {'categoria': 'quimico', 'risco': 'Poeiras', 'possiveis_danos': 'Doenças respiratórias, alergias.'},
        {'categoria': 'biologico', 'risco': 'Bactérias', 'possiveis_danos': 'Infecções diversas.'},
        {'categoria': 'ergonomico', 'risco': 'Posturas Inadequadas', 'possiveis_danos': 'Dores musculares, lesões na coluna.'},
        {'categoria': 'acidente', 'risco': 'Trabalho em Altura', 'possiveis_danos': 'Quedas, fraturas, morte.'}
    ]
    return pd.DataFrame(data)

def substituir_placeholders(doc, contexto):
    elementos = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                elementos.extend(cell.paragraphs)
    for p in elementos:
        full_text = "".join(run.text for run in p.runs)
        for key, value in contexto.items():
            if key in full_text:
                style = p.style
                p.clear()
                p.style = style
                parts = full_text.split(key)
                for i, part in enumerate(parts):
                    p.add_run(part)
                    if i < len(parts) - 1:
                        value_lines = str(value).split('\n')
                        for j, line in enumerate(value_lines):
                            run_valor = p.add_run(line)
                            run_valor.bold = False
                            run_valor.font.name = 'Segoe UI'
                            run_valor.font.size = Pt(9)
                            if j < len(value_lines) - 1:
                                run_valor.add_break()
                full_text = "".join(run.text for run in p.runs)

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
            categoria_display = risco_manual.get('categoria')
            categoria_alvo = map_categorias_rev.get(categoria_display)
            if categoria_alvo:
                riscos_por_categoria[categoria_alvo].append(risco_manual.get('risco', ''))
                if risco_manual.get('danos'):
                    danos_por_categoria[categoria_alvo].append(risco_manual.get('danos'))
    medicoes_ordenadas = sorted(medicoes_manuais, key=lambda med: med['agente'])
    medicoes_formatadas = []
    for med in medicoes_ordenadas:
        epi_info = f" | EPI: {med['epi']}" if med.get("epi") else ""
        medicoes_formatadas.append(f"{med['agente']}: {med['valor']} {med['unidade']}{epi_info}")
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
        "[EPIS]": tratar_lista_vazia(epis_manuais.split(',')) if epis_manuais else "Não aplicável",
        "[MEDIÇÕES]": medicoes_texto,
    }
    substituir_placeholders(doc, contexto)
    return doc

# --- APLICAÇÃO PRINCIPAL ---
def main():
    """Função principal da aplicação"""
    check_authentication()
    init_user_session_state()
    
    if not st.session_state.authenticated:
        show_login_page()
        return
    
    if st.session_state.authenticated:
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
        st.error("Não foi possível ler a planilha de funcionários. Verifique o arquivo.")
        st.stop()

    df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_raw)
    df_pgr = obter_dados_pgr()

    with st.container(border=True):
        st.markdown('##### 👥 2. Selecione os Funcionários')
        setores = sorted(df_funcionarios['setor'].dropna().unique().tolist()) if 'setor' in df_funcionarios.columns else []
        setor_sel = st.multiselect("Filtrar por Setor(es)", setores, placeholder="Selecione um ou mais setores")
        
        df_filtrado_setor = df_funcionarios[df_funcionarios['setor'].isin(setor_sel)] if setor_sel else df_funcionarios
        
        funcoes = sorted(df_filtrado_setor['funcao'].dropna().unique().tolist()) if 'funcao' in df_filtrado_setor.columns else []
        funcao_sel = st.multiselect("Filtrar por Função/Cargo(s)", funcoes, placeholder="Selecione uma ou mais funções")

        df_final_filtrado = df_filtrado_setor[df_filtrado_setor['funcao'].isin(funcao_sel)] if funcao_sel else df_filtrado_setor
        
        st.success(f"**{len(df_final_filtrado)} funcionários selecionados.**")
        st.dataframe(df_final_filtrado[['nome_do_funcionario', 'setor', 'funcao']], use_container_width=True, height=250)

    with st.container(border=True):
        st.markdown('##### ⚠️ 3. Configure os Riscos e Medidas de Controle')
        st.info("Os riscos, medições e EPIs configurados aqui serão aplicados a **TODOS** os funcionários selecionados acima.")

        st.markdown("**Riscos Identificados (PGR)**")
        riscos_selecionados = []
        tabs = st.tabs(list(CATEGORIAS_RISCO.values()))
        for i, (key, nome) in enumerate(CATEGORIAS_RISCO.items()):
            with tabs[i]:
                riscos_categoria = df_pgr[df_pgr['categoria'] == key]['risco'].tolist()
                selecionados = st.multiselect(f"Selecione os riscos:", options=riscos_categoria, key=f"riscos_{key}")
                riscos_selecionados.extend(selecionados)

        col_exp1, col_exp2, col_exp3 = st.columns(3)
        with col_exp1:
            with st.expander("📊 **Adicionar Medições**"):
                with st.form("form_medicao"):
                    agente = st.selectbox("Agente/Fonte", AGENTES_DE_RISCO)
                    valor = st.text_input("Valor Medido")
                    unidade = st.selectbox("Unidade", UNIDADES_DE_MEDIDA)
                    epi_med = st.text_input("EPI Associado")
                    submitted = st.form_submit_button("Adicionar Medição")
                    if submitted and agente and valor:
                        st.session_state.medicoes_adicionadas.append({"agente": agente, "valor": valor, "unidade": unidade, "epi": epi_med})
                        st.experimental_rerun()
                if st.session_state.medicoes_adicionadas:
                    st.write("**Adicionadas:**")
                    for med in st.session_state.medicoes_adicionadas:
                        st.markdown(f"- {med['agente']}: {med['valor']} {med['unidade']}")

        with col_exp2:
            with st.expander("➕ **Adicionar Risco Manual**"):
                 with st.form("form_risco_manual"):
                    risco = st.text_input("Descrição do Risco")
                    categoria = st.selectbox("Categoria", list(CATEGORIAS_RISCO.values()))
                    danos = st.text_area("Possíveis Danos")
                    submitted = st.form_submit_button("Adicionar Risco")
                    if submitted and risco and categoria:
                        st.session_state.riscos_manuais_adicionados.append({"risco": risco, "categoria": categoria, "danos": danos})
                        st.experimental_rerun()
                 if st.session_state.riscos_manuais_adicionados:
                    st.write("**Adicionados:**")
                    for r in st.session_state.riscos_manuais_adicionados:
                        st.markdown(f"- **{r['risco']}** ({r['categoria']})")

        with col_exp3:
            with st.expander("🦺 **Adicionar EPIs Gerais**"):
                with st.form("form_epi"):
                    epi_nome = st.text_input("Nome do EPI")
                    submitted = st.form_submit_button("Adicionar EPI")
                    if submitted and epi_nome:
                        st.session_state.epis_adicionados.append(epi_nome)
                        st.experimental_rerun()
                if st.session_state.epis_adicionados:
                    st.write("**Adicionados:**")
                    for epi_item in st.session_state.epis_adicionados:
                        st.markdown(f"- {epi_item}")

    st.divider()
    if st.button("🚀 Gerar OS para Funcionários Selecionados", type="primary", use_container_width=True, disabled=df_final_filtrado.empty):
        epis_finais = ", ".join(st.session_state.epis_adicionados)
        with st.spinner(f"Gerando {len(df_final_filtrado)} documentos..."):
            documentos_gerados = []
            for _, func in df_final_filtrado.iterrows():
                doc = gerar_os(func, df_pgr, riscos_selecionados, epis_finais, st.session_state.medicoes_adicionadas, st.session_state.riscos_manuais_adicionados, arquivo_modelo_os)
                doc_io = BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                nome_limpo = re.sub(r'[^\w\s-]', '', func.get("nome_do_funcionario", "Func_Sem_Nome")).strip().replace(" ", "_")
                documentos_gerados.append((f"OS_{nome_limpo}.docx", doc_io.getvalue()))

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
