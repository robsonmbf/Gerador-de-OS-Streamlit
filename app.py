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
# Lista de agentes de risco combinada e atualizada
AGENTES_DE_RISCO = sorted([
    'Abrasão',
    'Ambiente Artificialmente Frio',
    'Animais Peçonhentos',
    'Animais peçonhentos',
    'Arranjo Físico Inadequado',
    'Armazenamento Inadequado',
    'Atropelamento',
    'Bacilos',
    'Bactérias',
    'Biotipologia do Trabalhador',
    'Calor',
    'Choque elétrico',
    'Controle Rígido de Produtividade',
    'Corte e/ou perfuração',
    'Eletricidade',
    'Eletricidade (contato direto ou indireto)',
    'Equipamentos e/ou ferramentas cortantes e/ou perfurantes',
    'Equipamentos energizados',
    'Esforço Físico Intenso',
    'Exigência de Postura Inadequada',
    'Exigência e/ou execução de movimentos repetitivos',
    'Explosão',
    'Exposição a Agentes Biológicos (vírus, bactérias, protozoários, fungos, parasitas, bacilos)',
    'Exposição a Produto Químico',
    'Exposição a Temperaturas Elevadas (Calor)',
    'Exposição ao Ruído',
    'Exposição à Radiações Ionizantes',
    'Exposição à Radiações Não-ionizantes',
    'Ferramentas Inadequadas ou Defeituosas',
    'Frio',
    'Fumos',
    'Fungos',
    'Gases',
    'Iluminação Inadequada',
    'Impacto contra',
    'Imposição de Ritmos Excessivos',
    'Incêndio',
    'Inundação',
    'Jornada de Trabalho Prolongada',
    'Levantamento e Transporte Manual de Peso',
    'Levantamento, transporte e descarga individual de material',
    'Monotonia e Repetitividade',
    'Máquinas e Equipamentos sem Proteção',
    'Máquinas e/ou equipamentos em movimento (partes móveis)',
    'Neblinas',
    'Névoas',
    'Outras Situações Causadoras de Estresse Físico e/ou Psíquico',
    'Outros (Acidentes)',
    'Pancada por',
    'Parasitas',
    'Poeiras',
    'Poeiras, fumos e/ou gases tóxicos, asfixiantes e/ou inflamáveis',
    'Prensagem de membros',
    'Pressões Anormais',
    'Probabilidade de Incêndio ou Explosão',
    'Projeção de Partículas',
    'Projeção de materiais, peças e/ou partículas',
    'Protozoários',
    'Queda de Mesmo Nível',
    'Queda de Nível Diferente',
    'Queda de mesmo nível',
    'Queda de nível diferente',
    'Queda de objetos e/ou materiais',
    'Queimadura (térmica, elétrica ou química)',
    'Radiações Ionizantes',
    'Radiações Não-Ionizantes',
    'Ruptura de mangueira e/ou tubulação',
    'Ruptura e/ou projeção de materiais',
    'Ruptura, rompimento e/ou cisalhamento de estrutura e/ou componente',
    'Ruído (Contínuo ou Intermitente)',
    'Ruído (Impacto)',
    'Soterramento',
    'Substâncias Químicas (geral)',
    'Substâncias tóxicas e/ou inflamáveis',
    'Superfícies, substâncias e/ou objetos aquecidos',
    'Superfícies, substâncias e/ou objetos em baixa temperatura',
    'Tombamento de máquina/equipamento',
    'Tombamento, quebra e/ou ruptura de estrutura (fixa ou móvel)',
    'Trabalho à céu aberto',
    'Trabalho em Turno e Noturno',
    'Trabalho em espaços confinados',
    'Trabalho em turnos (diurno/noturno) e/ou jornadas de trabalho prolongadas',
    'Umidade',
    'Vapores',
    'Vibração de Corpo Inteiro',
    'Vibração de Corpo Inteiro (AREN)',
    'Vibração de Corpo Inteiro (VDVR)',
    'Vibração de Mãos e Braços',
    'Vibrações Localizadas (mão/braço)',
    'Vidro (recipientes, portas, bancadas, janelas, objetos diversos).',
    'Vírus'
])

# --- Funções Auxiliares ---

def substituir_placeholders(doc, dados):
    """Substitui os placeholders no documento Word pelos dados fornecidos."""
    for p in doc.paragraphs:
        for key, value in dados.items():
            if f'[{key}]' in p.text:
                # Usar inline para manter a formatação
                texto_original = p.text
                p.text = ""
                partes = texto_original.split(f'[{key}]')
                for i, parte in enumerate(partes):
                    p.add_run(parte)
                    if i < len(partes) - 1:
                        p.add_run(str(value)).bold = True

def adicionar_tabela_medicoes(doc, medicoes):
    """Adiciona a tabela de medições ao documento."""
    placeholder_tabela = "[TABELA_MEDICOES]"
    for p in doc.paragraphs:
        if placeholder_tabela in p.text:
            p.text = p.text.replace(placeholder_tabela, "")
            tabela = doc.add_table(rows=1, cols=4)
            tabela.style = 'Table Grid'
            hdr_cells = tabela.rows[0].cells
            hdr_cells[0].text = 'Agente de Risco'
            hdr_cells[1].text = 'Intensidade/Concentração'
            hdr_cells[2].text = 'Unidade de Medida'
            hdr_cells[3].text = 'Técnica Utilizada'

            for medicao in medicoes:
                row_cells = tabela.add_row().cells
                row_cells[0].text = medicao.get('agente_risco', '')
                row_cells[1].text = str(medicao.get('intensidade', ''))
                row_cells[2].text = medicao.get('unidade_medida', '')
                row_cells[3].text = medicao.get('tecnica', '')
            
            # Formatação do cabeçalho
            for cell in hdr_cells:
                cell.paragraphs[0].runs[0].font.bold = True
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            return # Para a função após encontrar e substituir o placeholder

def adicionar_riscos_manuais(doc, riscos_manuais):
    """Adiciona a lista de riscos manuais ao documento."""
    placeholder_riscos = "[LISTA_RISCOS_ADICIONAIS]"
    for p in doc.paragraphs:
        if placeholder_riscos in p.text:
            p.text = p.text.replace(placeholder_riscos, "")
            # Adiciona os riscos como uma lista com marcadores
            for risco in riscos_manuais:
                doc.add_paragraph(risco, style='List Bullet')
            return # Para a função após encontrar e substituir o placeholder

def gerar_os_para_funcionario(dados_func, medicoes, riscos_manuais, arquivo_modelo):
    """Gera um único documento de Ordem de Serviço."""
    doc = Document(arquivo_modelo)
    substituir_placeholders(doc, dados_func)
    
    if medicoes:
        adicionar_tabela_medicoes(doc, medicoes)
    else: # Limpa o placeholder se não houver medições
        for p in doc.paragraphs:
            if "[TABELA_MEDICOES]" in p.text:
                p.text = p.text.replace("[TABELA_MEDICOES]", "Não se aplica.")

    if riscos_manuais:
        adicionar_riscos_manuais(doc, riscos_manuais)
    else: # Limpa o placeholder se não houver riscos
        for p in doc.paragraphs:
            if "[LISTA_RISCOS_ADICIONAIS]" in p.text:
                p.text = p.text.replace("[LISTA_RISCOS_ADICIONAIS]", "Nenhum risco adicional informado.")
    
    return doc

# --- Inicialização do Estado da Sessão ---
def inicializar_session_state():
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
    if 'user_email' not in st.session_state:
        st.session_state.user_email = ""
    if 'db_manager' not in st.session_state:
        st.session_state.db_manager = DatabaseManager()
        st.session_state.db_manager.connect()
    if 'auth_manager' not in st.session_state:
        st.session_state.auth_manager = AuthManager(st.session_state.db_manager)
    if 'user_data_manager' not in st.session_state:
        st.session_state.user_data_manager = None
    if 'medicoes_adicionadas' not in st.session_state:
        st.session_state.medicoes_adicionadas = []
    if 'riscos_manuais_adicionados' not in st.session_state:
        st.session_state.riscos_manuais_adicionados = []
    if 'cargos_concluidos' not in st.session_state:
        st.session_state.cargos_concluidos = set()

# --- Interface de Login ---
def mostrar_tela_login():
    st.title("Bem-vindo ao Gerador de OS")
    st.subheader("Por favor, faça o login para continuar")

    with st.form("login_form"):
        email = st.text_input("Email")
        password = st.text_input("Senha", type="password")
        submit_button = st.form_submit_button("Login")

        if submit_button:
            if st.session_state.auth_manager.login_user(email, password):
                st.session_state.logged_in = True
                st.session_state.user_email = email
                st.session_state.user_data_manager = UserDataManager(st.session_state.db_manager, email)
                st.success("Login realizado com sucesso!")
                st.rerun()
            else:
                st.error("Email ou senha inválidos.")

# --- Interface Principal da Aplicação ---
def mostrar_app_principal():
    st.title("📄 Gerador de Ordens de Serviço (OS)")
    st.markdown("Faça o upload dos arquivos necessários e preencha os campos para gerar as Ordens de Serviço.")

    # --- Upload de Arquivos ---
    st.sidebar.header("1. Upload de Arquivos")
    arquivo_funcionarios = st.sidebar.file_uploader("📂 Planilha de Funcionários (.xlsx)", type=['xlsx'])
    arquivo_modelo_os = st.sidebar.file_uploader("📄 Modelo de OS (.docx)", type=['docx'])

    # --- Seção de Medições Quantitativas ---
    st.sidebar.header("2. Medições Quantitativas (Opcional)")
    with st.sidebar.expander("Adicionar Medição"):
        with st.form("form_medicao", clear_on_submit=True):
            agente_risco = st.selectbox("Agente de Risco", AGENTES_DE_RISCO, key="agente_risco_med")
            intensidade = st.number_input("Intensidade/Concentração", format="%.4f", step=0.0001)
            unidade_medida = st.selectbox("Unidade de Medida", UNIDADES_DE_MEDIDA)
            tecnica = st.text_input("Técnica Utilizada")
            submitted_medicao = st.form_submit_button("Adicionar Medição")

            if submitted_medicao:
                nova_medicao = {
                    "id": time.time(),
                    "agente_risco": agente_risco,
                    "intensidade": intensidade,
                    "unidade_medida": unidade_medida,
                    "tecnica": tecnica
                }
                st.session_state.medicoes_adicionadas.append(nova_medicao)
                st.success(f"Medição para '{agente_risco}' adicionada!")

    # Exibir medições adicionadas
    if st.session_state.medicoes_adicionadas:
        st.sidebar.write("**Medições Adicionadas:**")
        for medicao in st.session_state.medicoes_adicionadas:
            cols = st.sidebar.columns([4, 1])
            cols[0].info(f"{medicao['agente_risco']}: {medicao['intensidade']} {medicao['unidade_medida']}")
            if cols[1].button("🗑️", key=f"del_med_{medicao['id']}", help="Remover medição"):
                st.session_state.medicoes_adicionadas = [m for m in st.session_state.medicoes_adicionadas if m['id'] != medicao['id']]
                st.rerun()

    # --- Seção de Riscos Manuais ---
    st.sidebar.header("3. Riscos Adicionais (Opcional)")
    with st.sidebar.expander("Adicionar Risco Manualmente"):
        risco_manual = st.text_input("Descreva o risco")
        if st.button("Adicionar Risco"):
            if risco_manual:
                st.session_state.riscos_manuais_adicionados.append(risco_manual)
                st.success(f"Risco '{risco_manual}' adicionado!")
            else:
                st.warning("Por favor, descreva o risco.")

    # Exibir riscos manuais adicionados
    if st.session_state.riscos_manuais_adicionados:
        st.sidebar.write("**Riscos Adicionais Adicionados:**")
        for i, risco in enumerate(st.session_state.riscos_manuais_adicionados):
            cols_risco = st.sidebar.columns([4, 1])
            cols_risco[0].warning(risco)
            if cols_risco[1].button("🗑️", key=f"del_risco_{i}", help="Remover risco"):
                st.session_state.riscos_manuais_adicionados.pop(i)
                st.rerun()
                
    st.sidebar.markdown("---")
    if st.sidebar.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.user_email = ""
        st.session_state.user_data_manager = None
        # Limpar estado da sessão ao sair
        for key in list(st.session_state.keys()):
            if key not in ['db_manager', 'auth_manager']: # Mantém conexões
                del st.session_state[key]
        st.rerun()


    # --- Área Principal ---
    if not arquivo_funcionarios or not arquivo_modelo_os:
        st.info("Por favor, faça o upload da planilha de funcionários e do modelo de OS para começar.")
        st.stop()

    try:
        df_funcionarios = pd.read_excel(arquivo_funcionarios)
        st.success("Planilha de funcionários carregada com sucesso!")
        
        # Mapeamento de colunas
        st.subheader("Mapeamento de Colunas")
        st.write("Selecione as colunas da sua planilha que correspondem aos campos necessários.")
        
        colunas = df_funcionarios.columns.tolist()
        
        col_empresa = st.selectbox("Nome da Empresa", colunas, index=colunas.index('empresa') if 'empresa' in colunas else 0)
        col_cnpj = st.selectbox("CNPJ da Empresa", colunas, index=colunas.index('cnpj') if 'cnpj' in colunas else 0)
        col_nome_func = st.selectbox("Nome do Funcionário", colunas, index=colunas.index('nome_do_funcionario') if 'nome_do_funcionario' in colunas else 0)
        col_cpf = st.selectbox("CPF do Funcionário", colunas, index=colunas.index('cpf') if 'cpf' in colunas else 0)
        col_setor = st.selectbox("Setor", colunas, index=colunas.index('setor') if 'setor' in colunas else 0)
        col_funcao = st.selectbox("Função", colunas, index=colunas.index('funcao') if 'funcao' in colunas else 0)
        col_cbo = st.selectbox("CBO", colunas, index=colunas.index('cbo') if 'cbo' in colunas else 0)
        col_atividades = st.selectbox("Descrição das Atividades", colunas, index=colunas.index('atividades') if 'atividades' in colunas else 0)
        
        # Mapeamento para os placeholders
        mapa_colunas = {
            "EMPRESA": col_empresa,
            "CNPJ": col_cnpj,
            "NOME_FUNCIONARIO": col_nome_func,
            "CPF": col_cpf,
            "SETOR": col_setor,
            "FUNCAO": col_funcao,
            "CBO": col_cbo,
            "DESCRICAO_ATIVIDADES": col_atividades
        }
        
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel: {e}")
        st.stop()

    st.markdown("---")
    
    # Seleção de Cargos para Geração
    st.subheader("Seleção de Cargos para Geração")
    df_funcionarios['Setor/Função'] = df_funcionarios[col_setor] + " / " + df_funcionarios[col_funcao]
    cargos_unicos = df_funcionarios['Setor/Função'].unique().tolist()
    
    cargos_selecionados = st.multiselect(
        "Selecione os cargos para os quais deseja gerar a OS. A OS será a mesma para todos os funcionários do mesmo cargo.",
        cargos_unicos,
        default=cargos_unicos
    )

    if st.button("🚀 Gerar Ordens de Serviço"):
        if not cargos_selecionados:
            st.warning("Nenhum cargo selecionado. Por favor, selecione pelo menos um cargo para gerar as OS.")
            st.stop()

        with st.spinner("Gerando documentos... Por favor, aguarde."):
            documentos_gerados = []
            combinacoes_processadas = set()
            
            # Filtra o DataFrame para incluir apenas os cargos selecionados
            df_filtrado = df_funcionarios[df_funcionarios['Setor/Função'].isin(cargos_selecionados)]

            for index, row in df_filtrado.iterrows():
                # Preparar dados do funcionário
                dados_func = {key: row[col] for key, col in mapa_colunas.items()}
                
                # Gerar OS
                doc = gerar_os_para_funcionario(
                    dados_func, 
                    st.session_state.medicoes_adicionadas, 
                    st.session_state.riscos_manuais_adicionados, 
                    arquivo_modelo_os
                )
                
                # Salvar em memória
                doc_io = BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)

                # Limpar nome do arquivo para evitar caracteres inválidos
                nome_limpo = re.sub(r'[^\w\s-]', '', dados_func.get("NOME_FUNCIONARIO", "Func_Sem_Nome")).strip().replace(" ", "_")
                setor_limpo = re.sub(r'[^\w\s-]', '', dados_func.get("SETOR", "SemSetor")).strip().replace(" ", "_")
                funcao_limpa = re.sub(r'[^\w\s-]', '', dados_func.get("FUNCAO", "SemFuncao")).strip().replace(" ", "_")
                
                caminho_no_zip = f"{setor_limpo}/{funcao_limpa}/OS_{nome_limpo}.docx"
                documentos_gerados.append((caminho_no_zip, doc_io.getvalue()))

            if documentos_gerados:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for nome_arquivo, conteudo_doc in documentos_gerados:
                        zip_file.writestr(nome_arquivo, conteudo_doc)
                
                nome_arquivo_zip = f"OS_Geradas_{time.strftime('%Y%m%d')}.zip"
                st.success(f"🎉 **{len(documentos_gerados)} Ordens de Serviço geradas com sucesso!**")
                
                st.download_button(
                    label="📥 Baixar Todas as OS (.zip)",
                    data=zip_buffer.getvalue(),
                    file_name=nome_arquivo_zip,
                    mime="application/zip"
                )
            else:
                st.error("Nenhum documento foi gerado. Verifique os dados e seleções.")

# --- Lógica Principal ---
inicializar_session_state()

if st.session_state.logged_in:
    mostrar_app_principal()
else:
    mostrar_tela_login()
