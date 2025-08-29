import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import os
import zipfile
from io import BytesIO
import base64
import tempfile
import time

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Ordens de Serviço (OS)",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- Inicialização do Session State ---
if 'documentos_gerados' not in st.session_state:
    st.session_state.documentos_gerados = []

# --- Unidades de Medida ---
UNIDADES_MEDIDA = [
    "dB Linear", "dB(C)", "dB(A)", "m/s²", "m/s1,75", "ppm", "mg/m³", "g/m³",
    "f/cm³", "°C", "m/s", "%", "lx", "ufc/m³", "W/m²", "A/m", "mT", "µT",
    "mA", "kV/m", "V/m", "J/m²", "mJ/cm²", "mSv", "mppdc", "UR(%)", "Lux"
]

# --- Funções de Lógica de Negócio ---

def normalizar_colunas(df):
    """Normaliza os nomes das colunas de um DataFrame."""
    if df is None:
        return None
    df.columns = (
        df.columns.str.lower()
        .str.strip()
        .str.replace(" ", "_")
        .str.replace("ç", "c").str.replace("ã", "a").str.replace("é", "e")
        .str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("utf-8")
    )
    rename_map = {
        'perigo__(fator_de_risco/agente_nocivo/situacao_perigosa)': 'risco',
        'perigo_(fator_de_risco/agente_nocivo/situacao_perigosa)': 'risco',
        'possiveis_danos_ou_agravos_a_saude': 'possiveis_danos'
    }
    df = df.rename(columns=rename_map)

    if 'categoria' in df.columns:
        df['categoria'] = df['categoria'].str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("utf-8").str.lower()
    return df

def mapear_e_renomear_colunas_funcionarios(df):
    """Tenta adivinhar e renomear colunas da planilha de funcionários."""
    mapeamento = {
        'nome_do_funcionario': ['nome', 'funcionario', 'colaborador', 'nome_completo'],
        'funcao': ['funcao', 'cargo', 'carga'],
        'data_de_admissao': ['data_admissao', 'data_de_admissao', 'admissao'],
        'setor': ['setor', 'area', 'departamento'],
        'matricula': ['matricula', 'registro', 'id'],
        'descricao_de_atividades': ['descricao_de_atividades', 'atividades', 'descricao_atividades', 'tarefas', 'funcoes'],
    }
    colunas_renomeadas = {}
    for nome_padrao, nomes_possiveis in mapeamento.items():
        for nome_possivel in nomes_possiveis:
            if nome_possivel in df.columns:
                colunas_renomeadas[nome_possivel] = nome_padrao
                break
    df = df.rename(columns=colunas_renomeadas)
    return df

@st.cache_data
def carregar_planilha(arquivo):
    """Carrega e processa uma planilha genérica."""
    if arquivo is None:
        return None
    try:
        df = pd.read_excel(arquivo)
        df = normalizar_colunas(df)
        return df
    except Exception as e:
        st.error(f"Erro ao ler o ficheiro Excel: {e}")
        return None

def replace_text_in_paragraph(paragraph, contexto):
    """Substitui placeholders em um único parágrafo."""
    for key, value in contexto.items():
        if key in paragraph.text:
            inline = paragraph.runs
            for i in range(len(inline)):
                if key in inline[i].text:
                    text = inline[i].text.replace(key, str(value))
                    inline[i].text = text

def substituir_placeholders(doc, contexto):
    """Substitui os placeholders em todo o documento (parágrafos e tabelas)."""
    for p in doc.paragraphs:
        replace_text_in_paragraph(p, contexto)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_text_in_paragraph(p, contexto)

def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, perigo_manual, danos_manuais, categoria_manual, modelo_doc_carregado, logo_path=None):
    """Gera uma única Ordem de Serviço para um funcionário usando um modelo .docx carregado."""
    # Abre o modelo .docx a partir do arquivo carregado na memória
    doc = Document(modelo_doc_carregado)

    if logo_path:
        try:
            # Tenta inserir na primeira tabela (cabeçalho)
            header_table = doc.tables[0]
            cell = header_table.cell(0, 0)
            cell.text = "" 
            p = cell.paragraphs[0]
            run = p.add_run()
            run.add_picture(logo_path, width=Inches(2.0))
        except (IndexError, KeyError):
            st.warning("Aviso: Não foi encontrada uma tabela no cabeçalho do modelo para inserir a logo. A imagem será inserida no topo do documento.")
            p = doc.paragraphs[0]
            run = p.insert_paragraph_before().add_run()
            run.add_picture(logo_path, width=Inches(2.0))
            
    riscos_info = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]
    
    riscos_por_categoria = {"fisico": [], "quimico": [], "biologico": [], "ergonomico": [], "acidente": []}
    danos_por_categoria = {"fisico": [], "quimico": [], "biologico": [], "ergonomico": [], "acidente": []}
    epis_recomendados = set()
    
    for _, risco_row in riscos_info.iterrows():
        categoria = str(risco_row.get("categoria", "")).strip().lower()
        risco_nome = str(risco_row.get("risco", "")).strip()
        if categoria in riscos_por_categoria:
            riscos_por_categoria[categoria].append(risco_nome)
            danos = risco_row.get("possiveis_danos")
            epis = risco_row.get("epis_recomendados")
            if pd.notna(danos): danos_por_categoria[categoria].append(str(danos))
            if pd.notna(epis): epis_recomendados.update([epi.strip() for epi in str(epis).split(',')])

    if perigo_manual and categoria_manual:
        map_categorias = {"Acidentes": "acidente", "Biológicos": "biologico", "Ergonômicos": "ergonomico", "Físicos": "fisico", "Químicos": "quimico"}
        categoria_alvo = map_categorias.get(categoria_manual)
        if categoria_alvo:
            riscos_por_categoria[categoria_alvo].append(perigo_manual)
            if danos_manuais:
                danos_por_categoria[categoria_alvo].append(danos_manuais)

    if epis_manuais:
        epis_extras = [epi.strip() for epi in epis_manuais.split(',')]
        epis_recomendados.update(epis_extras)

    medicoes_lista = []
    if 'medicoes' in df_pgr.columns:
        medicoes_df = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]
        medicoes_lista = [f"{row['risco']}: {row['medicoes']}" for _, row in medicoes_df.iterrows() if 'medicoes' in row and pd.notna(row['medicoes'])]

    if medicoes_manuais:
        medicoes_lista.extend([med.strip() for med in medicoes_manuais.split('\n') if med.strip()])
        
    data_admissao = "não informado"
    if 'data_de_admissao' in funcionario and pd.notna(funcionario['data_de_admissao']):
        try:
            data_admissao = pd.to_datetime(funcionario['data_de_admissao']).strftime('%d/%m/%Y')
        except Exception:
            data_admissao = str(funcionario['data_de_admissao'])

    nome_funcionario = str(funcionario.get("nome_do_funcionario", "N/A")).replace("[", "").replace("]", "")

    descricao_atividades = "Não informado"
    if 'descricao_de_atividades' in funcionario and pd.notna(funcionario['descricao_de_atividades']):
        descricao_atividades = str(funcionario['descricao_de_atividades'])

    def tratar_risco_vazio(lista_riscos):
        if not lista_riscos or all(not r.strip() for r in lista_riscos):
            return "Não identificado"
        return ", ".join(lista_riscos)

    def tratar_danos_vazios(lista_danos):
        if not lista_danos or all(not d.strip() for d in lista_danos):
            return "Não identificado"
        return "; ".join(set(lista_danos))

    contexto = {
        "[NOME EMPRESA]": str(funcionario.get("empresa", "N/A")), 
        "[UNIDADE]": str(funcionario.get("unidade", "N/A")),
        "[NOME FUNCIONÁRIO]": nome_funcionario, 
        "[DATA DE ADMISSÃO]": data_admissao,
        "[SETOR]": str(funcionario.get("setor", "N/A")), 
        "[FUNÇÃO]": str(funcionario.get("funcao", "N/A")),
        "[DESCRIÇÃO DE ATIVIDADES]": descricao_atividades,
        "[RISCOS FÍSICOS]": tratar_risco_vazio(riscos_por_categoria["fisico"]),
        "[RISCOS DE ACIDENTE]": tratar_risco_vazio(riscos_por_categoria["acidente"]),
        "[RISCOS QUÍMICOS]": tratar_risco_vazio(riscos_por_categoria["quimico"]),
        "[RISCOS BIOLÓGICOS]": tratar_risco_vazio(riscos_por_categoria["biologico"]),
        "[RISCOS ERGONÔMICOS]": tratar_risco_vazio(riscos_por_categoria["ergonomico"]),
        "[POSSÍVEIS DANOS RISCOS FÍSICOS]": tratar_danos_vazios(danos_por_categoria["fisico"]),
        "[POSSÍVEIS DANOS RISCOS ACIDENTE]": tratar_danos_vazios(danos_por_categoria["acidente"]),
        "[POSSÍVEIS DANOS RISCOS QUÍMICOS]": tratar_danos_vazios(danos_por_categoria["quimico"]),
        "[POSSÍVEIS DANOS RISCOS BIOLÓGICOS]": tratar_danos_vazios(danos_por_categoria["biologico"]),
        "[POSSÍVEIS DANOS RISCOS ERGONÔMICOS]": tratar_danos_vazios(danos_por_categoria["ergonomico"]),
        "[EPIS]": ", ".join(sorted(list(epis_recomendados))) or "Não aplicável",
        "[MEDIÇÕES]": "\n".join(medicoes_lista) or "Não aplicável",
    }
    
    substituir_placeholders(doc, contexto)
    
    return doc

# --- Base de dados PGR incorporada ---
def obter_dados_pgr():
    """Retorna os dados PGR padrão incorporados no sistema."""
    return pd.DataFrame([
        # (A lista de riscos continua a mesma, sem alterações)
        {'categoria': 'fisico', 'risco': 'Ruído', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'},
        {'categoria': 'quimico', 'risco': 'Poeiras', 'possiveis_danos': 'Pneumoconioses, irritação respiratória, alergias.'},
        {'categoria': 'biologico', 'risco': 'Bactérias', 'possiveis_danos': 'Infecções, doenças infecciosas.'},
        {'categoria': 'ergonomico', 'risco': 'Posturas Inadequadas', 'possiveis_danos': 'Dores musculares, lesões na coluna, LER/DORT.'},
        {'categoria': 'acidente', 'risco': 'Trabalho em Altura', 'possiveis_danos': 'Quedas, fraturas, morte.'}
        # ... (imagine a lista completa aqui para economizar espaço)
    ])

# --- Interface do Streamlit ---

st.markdown("""
<style>
    /* O CSS continua o mesmo */
    .main-header { background: linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%); padding: 2rem; border-radius: 10px; margin-bottom: 2rem; text-align: center; color: white; }
    .section-header { background: #f8fafc; padding: 1rem; border-left: 4px solid #3b82f6; margin: 1rem 0; border-radius: 5px; }
    .info-box { background: #e0f2fe; padding: 1rem; border-radius: 8px; border: 1px solid #0284c7; margin: 1rem 0; }
    .success-box { background: #dcfce7; padding: 1rem; border-radius: 8px; border: 1px solid #16a34a; margin: 1rem 0; }
    .stButton > button { background: linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%); color: white; border: none; padding: 0.5rem 2rem; border-radius: 8px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>📄 Gerador de Ordens de Serviço (OS)</h1>
    <p>Sistema para geração automática de OS a partir de um modelo Word (.docx) e uma planilha de funcionários.</p>
</div>
""", unsafe_allow_html=True)

# --- Sidebar para upload de arquivos ---
st.sidebar.markdown("### 📁 Carregar Arquivos")
arquivo_funcionarios = st.sidebar.file_uploader(
    "1. Planilha de Funcionários (.xlsx)", 
    type="xlsx", 
    help="Ficheiro .xlsx com os dados dos funcionários."
)

# MUDANÇA: Uploader de modelo agora é obrigatório
arquivo_modelo_os = st.sidebar.file_uploader(
    "2. Modelo de OS (.docx)",
    type="docx",
    help="Carregue o arquivo .docx que servirá como modelo para todas as OS geradas. Este arquivo é obrigatório."
)

arquivo_logo = st.sidebar.file_uploader(
    "3. Logo da Empresa (Opcional)", 
    type=["png", "jpg", "jpeg"],
    help="Imagem da logo para inserir no cabeçalho do documento."
)

st.sidebar.markdown("---")
st.sidebar.markdown("### ℹ️ Instruções")
st.sidebar.info(
    "Para gerar as OS, é **obrigatório** carregar a planilha de funcionários e o seu modelo de documento Word (.docx)."
)


# --- Lógica principal da Interface ---
df_pgr = obter_dados_pgr()

if arquivo_funcionarios is None or arquivo_modelo_os is None:
    st.markdown('<div class="info-box">📋 Por favor, carregue a <strong>Planilha de Funcionários</strong> e o <strong>Modelo de OS (.docx)</strong> na barra lateral para começar.</div>', unsafe_allow_html=True)
else:
    # A partir daqui, o código só é executado se ambos os arquivos forem carregados.
    df_funcionarios = carregar_planilha(arquivo_funcionarios)
    if df_funcionarios is not None:
        df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios)
        
        st.markdown('<div class="section-header"><h3>👥 Seleção de Funcionários</h3></div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            setores = ["Todos"] + (df_funcionarios['setor'].dropna().unique().tolist() if 'setor' in df_funcionarios.columns else [])
            setor_sel = st.selectbox("Filtrar por Setor", setores)
        with col2:
            df_filtrado = df_funcionarios[df_funcionarios['setor'] == setor_sel] if setor_sel != "Todos" else df_funcionarios
            funcoes = ["Todos"] + (df_filtrado['funcao'].dropna().unique().tolist() if 'funcao' in df_filtrado.columns else [])
            funcao_sel = st.selectbox("Filtrar por Função/Cargo", funcoes)

        df_final_filtrado = df_filtrado[df_filtrado['funcao'] == funcao_sel] if funcao_sel != "Todos" else df_filtrado
        
        st.markdown(f'<div class="info-box">✅ <strong>{len(df_final_filtrado)} funcionários</strong> correspondem aos filtros.</div>', unsafe_allow_html=True)
        
        if not df_final_filtrado.empty:
            cols_mostrar = [c for c in ['nome_do_funcionario', 'setor', 'funcao'] if c in df_final_filtrado.columns]
            st.dataframe(df_final_filtrado[cols_mostrar], use_container_width=True)
        
        # O restante da interface (seleção de riscos, etc.)
        st.markdown('<div class="section-header"><h3>⚠️ Configuração de Riscos</h3></div>', unsafe_allow_html=True)
        
        # ... (A lógica de seleção de riscos, EPIs, etc., continua a mesma) ...
        categorias = {'fisico': '🔥 Físicos', 'quimico': '⚗️ Químicos', 'biologico': '🦠 Biológicos', 'ergonomico': '🏃 Ergonômicos', 'acidente': '⚠️ Acidentes'}
        riscos_selecionados = []
        
        tabs = st.tabs(list(categorias.values()))
        for i, (categoria, nome_categoria) in enumerate(categorias.items()):
            with tabs[i]:
                riscos_categoria = df_pgr[df_pgr['categoria'] == categoria]['risco'].tolist()
                selecionados_categoria = st.multiselect(f"Selecione ({nome_categoria}):", options=riscos_categoria, key=f"riscos_{categoria}")
                riscos_selecionados.extend(selecionados_categoria)

        with st.expander("➕ Adicionar Risco Manual / EPIs / Medições"):
            perigo_manual = st.text_input("Descrição do Risco Manual")
            categoria_manual = st.selectbox("Categoria do Risco Manual", [""] + list(categorias.values()))
            danos_manuais = st.text_area("Possíveis Danos do Risco Manual")
            epis_manuais = st.text_area("EPIs Adicionais (separados por vírgula)")
            medicoes_manuais = st.text_area("Medições (uma por linha)")
        
        st.markdown('<div class="section-header"><h3>🚀 Gerar Ordens de Serviço</h3></div>', unsafe_allow_html=True)
        
        if st.button("🔄 Gerar OSs para Funcionários Selecionados", type="primary"):
            if df_final_filtrado.empty:
                st.error("Nenhum funcionário selecionado! Ajuste os filtros.")
            else:
                with st.spinner("Gerando Ordens de Serviço..."):
                    documentos_gerados = []
                    logo_path = None
                    
                    if arquivo_logo:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(arquivo_logo.name)[1]) as temp_logo:
                            temp_logo.write(arquivo_logo.getbuffer())
                            logo_path = temp_logo.name
                    
                    progress_bar = st.progress(0)
                    total = len(df_final_filtrado)
                    
                    for i, (_, func) in enumerate(df_final_filtrado.iterrows()):
                        progress_bar.progress((i + 1) / total, text=f"Gerando para: {func.get('nome_do_funcionario', '')}")
                        try:
                            # A chamada agora passa o arquivo de modelo obrigatoriamente
                            doc = gerar_os(
                                func, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais,
                                perigo_manual, danos_manuais, categoria_manual,
                                modelo_doc_carregado=arquivo_modelo_os, logo_path=logo_path
                            )
                            
                            doc_io = BytesIO()
                            doc.save(doc_io)
                            doc_io.seek(0)
                            
                            nome_limpo = str(func.get("nome_do_funcionario", f"Func_{i+1}")).replace(" ", "_")
                            documentos_gerados.append((f"OS_{nome_limpo}.docx", doc_io.getvalue()))
                            
                        except Exception as e:
                            st.error(f"Erro ao gerar OS para {func.get('nome_do_funcionario', 'desconhecido')}: {e}")
                    
                    if logo_path: os.unlink(logo_path)
                    
                    if documentos_gerados:
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for nome, conteudo in documentos_gerados: zf.writestr(nome, conteudo)
                        zip_buffer.seek(0)
                        
                        st.markdown(f'<div class="success-box">✅ <strong>{len(documentos_gerados)} Ordens de Serviço</strong> geradas!</div>', unsafe_allow_html=True)
                        st.download_button(
                            "📥 Baixar Todas as OSs (.zip)",
                            data=zip_buffer,
                            file_name=f"Ordens_de_Servico_{time.strftime('%Y%m%d')}.zip",
                            mime="application/zip",
                        )
