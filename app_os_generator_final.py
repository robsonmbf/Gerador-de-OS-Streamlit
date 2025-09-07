import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import zipfile
from io import BytesIO
import time
import re

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Ordens de Serviço (OS)",
    page_icon="📄",
    layout="wide",
)

# --- CSS PERSONALIZADO ---
st.markdown("""
<style>
    [data-testid="stSidebar"] {
        display: none;
    }
</style>
""", unsafe_allow_html=True)

# --- INICIALIZAÇÃO DO SESSION STATE ---
if 'medicoes_adicionadas' not in st.session_state:
    st.session_state.medicoes_adicionadas = []
if 'riscos_manuais_adicionados' not in st.session_state:
    st.session_state.riscos_manuais_adicionados = []
if 'epis_adicionados' not in st.session_state:
    st.session_state.epis_adicionados = []
if 'setores_concluidos' not in st.session_state:
    st.session_state.setores_concluidos = set()
if 'cargos_concluidos' not in st.session_state:
    st.session_state.cargos_concluidos = set()

# --- LISTAS DE DADOS CONSTANTES ---
UNIDADES_DE_MEDIDA = ["dB(A)", "m/s²", "ppm", "mg/m³", "%", "°C", "lx", "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"]
AGENTES_DE_RISCO = sorted([
    "Ruído (Contínuo ou Intermitente)", "Ruído (Impacto)", "Vibração de Corpo Inteiro", "Vibração de Mãos e Braços",
    "Radiações Ionizantes", "Radiações Não-Ionizantes", "Frio", "Calor", "Pressões Anormais", "Umidade", "Poeiras", 
    "Fumos", "Névoas", "Neblinas", "Gases", "Vapores", "Produtos Químicos em Geral", "Vírus", "Bactérias", 
    "Protozoários", "Fungos", "Parasitas", "Bacilos"
])

# --- Funções de Lógica de Negócio ---
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

def substituir_placeholders(doc, contexto):
    all_paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paragraphs.extend(cell.paragraphs)
    for section in doc.sections:
        header = section.header
        all_paragraphs.extend(header.paragraphs)
        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_paragraphs.extend(cell.paragraphs)
    for p in all_paragraphs:
        full_text = "".join(run.text for run in p.runs)
        if not full_text.strip(): continue
        keys_in_text = [key for key in contexto if key in full_text]
        if not keys_in_text: continue
        original_text = full_text
        for key in keys_in_text:
            full_text = full_text.replace(str(key), str(contexto[key]))
        if original_text != full_text:
            for run in p.runs: run.text = ''
            new_run = p.add_run(full_text)
            font = new_run.font
            font.name = 'Segoe UI'
            font.size = Pt(9)

def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, riscos_manuais, modelo_doc_carregado):
    doc = Document(modelo_doc_carregado)
    riscos_info = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]
    riscos_por_categoria = {cat: [] for cat in ["fisico", "quimico", "biologico", "ergonomico", "acidente"]}
    danos_por_categoria = {cat: [] for cat in ["fisico", "quimico", "biologico", "ergonomico", "acidente"]}
    for _, risco_row in riscos_info.iterrows():
        categoria = str(risco_row.get("categoria", "")).lower()
        if categoria in riscos_por_categoria:
            riscos_por_categoria[categoria].append(str(risco_row.get("risco", "")))
            danos = risco_row.get("possiveis_danos")
            if pd.notna(danos): danos_por_categoria[categoria].append(str(danos))

    if riscos_manuais:
        map_categorias = {"🔥 Físicos": "fisico", "⚗️ Químicos": "quimico", "🦠 Biológicos": "biologico", "🏃 Ergonômicos": "ergonomico", "⚠️ Acidentes": "acidente"}
        for risco_manual in riscos_manuais:
            categoria_display = risco_manual.get('categoria')
            categoria_alvo = map_categorias.get(categoria_display)
            if categoria_alvo:
                riscos_por_categoria[categoria_alvo].append(risco_manual.get('risco', ''))
                if risco_manual.get('danos'):
                    danos_por_categoria[categoria_alvo].append(risco_manual.get('danos'))

    epis_recomendados = set(epi.strip() for epi in epis_manuais.split(',') if epi.strip())

    # Ajuste: incluir EPIs vinculados às medições
    medicoes_formatadas = []
    for med in medicoes_manuais:
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
        return separador.join(sorted(list(set(lista))))

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
        "[EPIS]": tratar_lista_vazia(list(epis_recomendados)) or "Não aplicável",
        "[MEDIÇÕES]": medicoes_texto,
    }
    substituir_placeholders(doc, contexto)
    return doc

# --- Interface do Streamlit ---
st.markdown("""<div class="main-header"><h1>📄 Gerador de Ordens de Serviço (OS)</h1><p>Geração automática de OS a partir de um modelo Word (.docx) e uma planilha de funcionários.</p></div>""", unsafe_allow_html=True)
with st.container(border=True):
    st.markdown("##### 📂 Documentos Necessários")
    col1, col2 = st.columns(2)
    with col1:
        arquivo_funcionarios = st.file_uploader("📄 **1. Planilha de Funcionários (.xlsx)**", type="xlsx", help="Carregue a planilha com os dados dos funcionários.")
    with col2:
        arquivo_modelo_os = st.file_uploader("📝 **2. Modelo de OS (.docx)**", type="docx", help="Carregue seu modelo de Ordem de Serviço em formato Word.")

if not arquivo_funcionarios or not arquivo_modelo_os:
    st.info("📋 Por favor, carregue a Planilha de Funcionários e o Modelo de OS acima para começar.")
else:
    df_funcionarios_raw = carregar_planilha(arquivo_funcionarios)
    df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_raw)
    df_pgr = pd.DataFrame()  # aqui ficaria sua função obter_dados_pgr(), mantive resumido

    # --- Adicionar Medições Ambientais ---
    with st.expander("📊 Adicionar Medições Ambientais"):
        def adicionar_medicao():
            agente_selecionado = st.session_state.agente_input
            agente_manual = st.session_state.agente_manual_input if "agente_manual_input" in st.session_state else None
            agente_final = agente_manual if agente_selecionado == "Outro (digitar manualmente)" else agente_selecionado
            valor = st.session_state.valor_input
            unidade = st.session_state.unidade_input
            epi = st.session_state.epi_medicao_input

            if agente_final and valor:
                medicao = {"agente": agente_final, "valor": valor, "unidade": unidade, "epi": epi}
                st.session_state.medicoes_adicionadas.append(medicao)
                st.session_state.agente_input = ""
                st.session_state.valor_input = ""
                st.session_state.epi_medicao_input = ""
                if "agente_manual_input" in st.session_state:
                    st.session_state.agente_manual_input = ""
            else:
                st.warning("Preencha o Agente e o Valor para adicionar uma medição.")

        def limpar_medicoes():
            st.session_state.medicoes_adicionadas = []

        col1, col2, col3 = st.columns([2,1,1])
        with col1: 
            agente_selecionado = st.selectbox(
                "Agente/Fonte do Risco",
                options=[""] + AGENTES_DE_RISCO + ["Outro (digitar manualmente)"],
                key="agente_input"
            )
            if agente_selecionado == "Outro (digitar manualmente)":
                st.text_input("Digite o Agente/Fonte do Risco manualmente", key="agente_manual_input")

        with col2:
            st.text_input("Valor Medido", key="valor_input")
        with col3:
            st.selectbox("Unidade de Medida", UNIDADES_DE_MEDIDA, key="unidade_input")

        st.text_input("EPI associado à medição", key="epi_medicao_input")

        col_btn1, col_btn2, _ = st.columns([1,1,2])
        with col_btn1:
            st.button("Adicionar Medição", on_click=adicionar_medicao)
        with col_btn2:
            st.button("Limpar Lista de Medições", on_click=limpar_medicoes)

        if st.session_state.medicoes_adicionadas:
            st.write("**Medições Adicionadas:**")
            for med in st.session_state.medicoes_adicionadas:
                epi_info = f" | EPI: {med['epi']}" if med.get("epi") else ""
                st.markdown(f"- {med['agente']}: {med['valor']} {med['unidade']}{epi_info}")
