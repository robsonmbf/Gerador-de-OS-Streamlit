import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
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

# --- DEFINIÇÃO DE CONSTANTES ---
UNIDADES_DE_MEDIDA = ["dB(A)", "m/s²", "m/s¹⁷⁵", "ppm", "mg/m³", "%", "°C", "lx", "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"]
AGENTES_DE_RISCO = sorted([
    "Ruído (Contínuo ou Intermitente)", "Ruído (Impacto)", "Vibração de Corpo Inteiro", "Vibração de Mãos e Braços",
    "Radiações Ionizantes", "Radiações Não-Ionizantes", "Frio", "Calor", "Pressões Anormais", "Umidade", "Poeiras", 
    "Fumos", "Névoas", "Neblinas", "Gases", "Vapores", "Produtos Químicos em Geral", "Vírus", "Bactérias", 
    "Protozoários", "Fungos", "Parasitas", "Bacilos"
])
CATEGORIAS_RISCO = {
    'fisico': '🔥 Físicos',
    'quimico': '⚗️ Químicos',
    'biologico': '🦠 Biológicos',
    'ergonomico': '🏃 Ergonômicos',
    'acidente': '⚠️ Acidentes'
}

# --- CSS ---
st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none; }
    .main-header { text-align: center; padding-bottom: 20px; }
</style>
""", unsafe_allow_html=True)


# --- FUNÇÕES AUXILIARES ---
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
        st.error(f"Erro ao ler o Excel: {e}")
        return None

@st.cache_data
def obter_dados_pgr(arquivo_pgr=None):
    """Carrega riscos do PGR enviado ou usa lista padrão."""
    if arquivo_pgr:
        try:
            df = pd.read_excel(arquivo_pgr)
            if not {'categoria','risco','possiveis_danos'}.issubset(df.columns.str.lower()):
                st.warning("Planilha de riscos inválida. Usando lista padrão.")
            else:
                return df
        except Exception as e:
            st.error(f"Erro ao ler riscos do Excel: {e}")
    # fallback padrão
    data = [
        {'categoria': 'fisico', 'risco': 'Ruído (Contínuo ou Intermitente)', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'},
        {'categoria': 'fisico', 'risco': 'Calor', 'possiveis_danos': 'Desidratação, insolação, exaustão.'},
        {'categoria': 'quimico', 'risco': 'Poeiras', 'possiveis_danos': 'Irritação respiratória, pneumoconioses.'},
        {'categoria': 'biologico', 'risco': 'Bactérias', 'possiveis_danos': 'Infecções bacterianas diversas.'},
        {'categoria': 'ergonomico', 'risco': 'Posturas Inadequadas', 'possiveis_danos': 'LER/DORT, dores musculares.'},
        {'categoria': 'acidente', 'risco': 'Trabalho em Altura', 'possiveis_danos': 'Quedas, fraturas, morte.'},
    ]
    return pd.DataFrame(data)


def substituir_placeholders(doc, contexto):
    def aplicar_formatacao_padrao(run):
        run.font.name = 'Segoe UI'
        run.font.size = Pt(9)
        return run
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in contexto.items():
                        if key in p.text:
                            for run in p.runs: run.text = run.text.replace(key, str(value))
    for p in doc.paragraphs:
        for key, value in contexto.items():
            if key in p.text:
                for run in p.runs: run.text = run.text.replace(key, str(value))

def gerar_os(funcionario, df_pgr, riscos_selecionados, epis, medicoes, modelo_doc_carregado):
    doc = Document(modelo_doc_carregado)
    riscos_info = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]

    riscos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}
    danos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}

    for _, risco_row in riscos_info.iterrows():
        categoria = str(risco_row.get("categoria", "")).lower()
        if categoria in riscos_por_categoria:
            riscos_por_categoria[categoria].append(str(risco_row.get("risco", "")))
            if pd.notna(risco_row.get("possiveis_danos")):
                danos_por_categoria[categoria].append(str(risco_row["possiveis_danos"]))

    def tratar_lista(lista):
        return ", ".join(sorted(set(lista))) if lista else "Não identificado"

    contexto = {
        "[NOME EMPRESA]": str(funcionario.get("empresa", "N/A")),
        "[UNIDADE]": str(funcionario.get("unidade", "N/A")),
        "[NOME FUNCIONÁRIO]": str(funcionario.get("nome_do_funcionario", "N/A")),
        "[DATA DE ADMISSÃO]": str(funcionario.get("data_de_admissao", "N/A")),
        "[SETOR]": str(funcionario.get("setor", "N/A")),
        "[FUNÇÃO]": str(funcionario.get("funcao", "N/A")),
        "[DESCRIÇÃO DE ATIVIDADES]": str(funcionario.get("descricao_de_atividades", "Não informado")),
        "[RISCOS FÍSICOS]": tratar_lista(riscos_por_categoria["fisico"]),
        "[RISCOS DE ACIDENTE]": tratar_lista(riscos_por_categoria["acidente"]),
        "[RISCOS QUÍMICOS]": tratar_lista(riscos_por_categoria["quimico"]),
        "[RISCOS BIOLÓGICOS]": tratar_lista(riscos_por_categoria["biologico"]),
        "[RISCOS ERGONÔMICOS]": tratar_lista(riscos_por_categoria["ergonomico"]),
        "[POSSÍVEIS DANOS RISCOS FÍSICOS]": tratar_lista(danos_por_categoria["fisico"]),
        "[POSSÍVEIS DANOS RISCOS ACIDENTE]": tratar_lista(danos_por_categoria["acidente"]),
        "[POSSÍVEIS DANOS RISCOS QUÍMICOS]": tratar_lista(danos_por_categoria["quimico"]),
        "[POSSÍVEIS DANOS RISCOS BIOLÓGICOS]": tratar_lista(danos_por_categoria["biologico"]),
        "[POSSÍVEIS DANOS RISCOS ERGONÔMICOS]": tratar_lista(danos_por_categoria["ergonomico"]),
        "[EPIS]": tratar_lista([e['epi_name'] for e in epis]) if epis else "Não aplicável",
        "[MEDIÇÕES]": tratar_lista([f"{m['agent']}: {m['value']} {m['unit']}" for m in medicoes]) if medicoes else "Não aplicável",
    }

    substituir_placeholders(doc, contexto)
    return doc


# --- APLICAÇÃO PRINCIPAL ---
def main():
    st.markdown("""<div class="main-header"><h1>📄 Gerador de Ordens de Serviço (OS)</h1><p>Gere OS em lote a partir de um modelo Word (.docx) e planilhas.</p></div>""", unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown("##### 📂 1. Carregue os Documentos")
        col1, col2, col3 = st.columns(3)
        with col1:
            arquivo_funcionarios = st.file_uploader("📄 **Planilha de Funcionários (.xlsx)**", type="xlsx")
        with col2:
            arquivo_modelo_os = st.file_uploader("📝 **Modelo de OS (.docx)**", type="docx")
        with col3:
            arquivo_pgr = st.file_uploader("⚠️ **Perigos e Riscos PGR (.xlsx)**", type="xlsx")

    if not arquivo_funcionarios or not arquivo_modelo_os:
        st.info("📋 Carregue funcionários e modelo de OS para continuar.")
        return

    df_funcionarios_raw = carregar_planilha(arquivo_funcionarios)
    if df_funcionarios_raw is None: st.stop()

    df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_raw)
    df_pgr = obter_dados_pgr(arquivo_pgr)

    with st.container(border=True):
        st.markdown('##### 👥 2. Selecione os Funcionários')
        setores = sorted(df_funcionarios['setor'].dropna().unique().tolist()) if 'setor' in df_funcionarios else []
        setor_sel = st.multiselect("Filtrar por Setor(es)", setores)
        df_filtrado_setor = df_funcionarios[df_funcionarios['setor'].isin(setor_sel)] if setor_sel else df_funcionarios
        funcoes_disponiveis = sorted(df_filtrado_setor['funcao'].dropna().unique().tolist()) if 'funcao' in df_filtrado_setor else []
        funcao_sel = st.multiselect("Filtrar por Função/Cargo(s)", funcoes_disponiveis)
        df_final = df_filtrado_setor[df_filtrado_setor['funcao'].isin(funcao_sel)] if funcao_sel else df_filtrado_setor
        st.success(f"{len(df_final)} funcionário(s) selecionado(s).")
        st.dataframe(df_final[['nome_do_funcionario','setor','funcao']])

    with st.container(border=True):
        st.markdown('##### ⚠️ 3. Selecione os Riscos')
        riscos_selecionados = []
        tabs = st.tabs(CATEGORIAS_RISCO.values())
        for i, (categoria_key, categoria_nome) in enumerate(CATEGORIAS_RISCO.items()):
            with tabs[i]:
                riscos_cat = df_pgr[df_pgr['categoria'] == categoria_key]['risco'].tolist()
                selecionados = st.multiselect("Riscos:", options=riscos_cat, key=f"riscos_{categoria_key}")
                riscos_selecionados.extend(selecionados)

    if st.button("🚀 Gerar OS", disabled=df_final.empty):
        with st.spinner(f"Gerando {len(df_final)} documentos..."):
            documentos_gerados = []
            for _, func in df_final.iterrows():
                doc = gerar_os(func, df_pgr, riscos_selecionados, [], [], arquivo_modelo_os)
                doc_io = BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                nome_limpo = re.sub(r'[^\w\s-]', '', func.get("nome_do_funcionario","SemNome")).strip().replace(" ","_")
                caminho_no_zip = f"{func.get('setor','SemSetor')}/{func.get('funcao','SemFuncao')}/OS_{nome_limpo}.docx"
                documentos_gerados.append((caminho_no_zip, doc_io.getvalue()))

            if documentos_gerados:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for nome_arquivo, conteudo in documentos_gerados:
                        zip_file.writestr(nome_arquivo, conteudo)
                nome_zip = f"OS_Geradas_{time.strftime('%Y%m%d')}.zip"
                st.success(f"{len(documentos_gerados)} OS geradas!")
                st.download_button("📥 Baixar Todas", data=zip_buffer.getvalue(), file_name=nome_zip, mime="application/zip")

if __name__ == "__main__":
    main()
