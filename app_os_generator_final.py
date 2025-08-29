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
        #df = normalizar_colunas(df) # Removido para manter nomes originais na exibição
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
    doc = Document(modelo_doc_carregado)

    if logo_path:
        try:
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
        map_categorias = {"Físicos": "fisico", "Químicos": "quimico", "Biológicos": "biologico", "Ergonômicos": "ergonomico", "Acidentes": "acidente"}
        categoria_alvo = map_categorias.get(categoria_manual)
        if categoria_alvo:
            riscos_por_categoria[categoria_alvo].append(perigo_manual)
            if danos_manuais:
                danos_por_categoria[categoria_alvo].append(danos_manuais)

    if epis_manuais:
        epis_extras = [epi.strip() for epi in epis_manuais.split(',')]
        epis_recomendados.update(epis_extras)

    medicoes_lista = []
    if medicoes_manuais:
        medicoes_lista.extend([med.strip() for med in medicoes_manuais.split('\n') if med.strip()])
        
    data_admissao = "não informado"
    if 'data_de_admissao' in funcionario and pd.notna(funcionario['data_de_admissao']):
        try:
            data_admissao = pd.to_datetime(funcionario['data_de_admissao']).strftime('%d/%m/%Y')
        except Exception:
            data_admissao = str(funcionario['data_de_admissao'])

    nome_funcionario = str(funcionario.get("nome_do_funcionario", "N/A"))
    descricao_atividades = str(funcionario.get("descricao_de_atividades", "Não informado"))

    def tratar_lista_vazia(lista, separador=", "):
        if not lista or all(not item.strip() for item in lista):
            return "Não identificado"
        return separador.join(set(lista))

    contexto = {
        "[NOME EMPRESA]": str(funcionario.get("empresa", "N/A")), 
        "[UNIDADE]": str(funcionario.get("unidade", "N/A")),
        "[NOME FUNCIONÁRIO]": nome_funcionario, 
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
        "[EPIS]": ", ".join(sorted(list(epis_recomendados))) or "Não aplicável",
        "[MEDIÇÕES]": "\n".join(medicoes_lista) or "Não aplicável",
    }
    
    substituir_placeholders(doc, contexto)
    
    return doc

# --- Base de dados PGR incorporada (CORRIGIDA) ---
def obter_dados_pgr():
    """Retorna os dados PGR padrão incorporados no sistema."""
    return pd.DataFrame([
        {'categoria': 'fisico', 'risco': 'Ruído', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'},
        {'categoria': 'fisico', 'risco': 'Vibração', 'possiveis_danos': 'Doenças osteomusculares, problemas circulatórios.'},
        {'categoria': 'fisico', 'risco': 'Calor', 'possiveis_danos': 'Desidratação, insolação, cãibras, exaustão, intermação.'},
        {'categoria': 'fisico', 'risco': 'Frio', 'possiveis_danos': 'Hipotermia, congelamento, doenças respiratórias.'},
        {'categoria': 'fisico', 'risco': 'Radiações Ionizantes', 'possiveis_danos': 'Câncer, mutações genéticas, queimaduras.'},
        {'categoria': 'fisico', 'risco': 'Radiações Não Ionizantes', 'possiveis_danos': 'Queimaduras, lesões oculares, câncer de pele.'},
        {'categoria': 'fisico', 'risco': 'Pressões Anormais', 'possiveis_danos': 'Doença descompressiva, barotrauma.'},
        {'categoria': 'fisico', 'risco': 'Umidade', 'possiveis_danos': 'Doenças respiratórias, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Poeiras', 'possiveis_danos': 'Pneumoconioses, irritação respiratória, alergias.'},
        {'categoria': 'quimico', 'risco': 'Fumos', 'possiveis_danos': 'Doenças respiratórias, intoxicações.'},
        {'categoria': 'quimico', 'risco': 'Névoas', 'possiveis_danos': 'Irritação respiratória, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Gases', 'possiveis_danos': 'Asfixia, intoxicações, irritação respiratória.'},
        {'categoria': 'quimico', 'risco': 'Vapores', 'possiveis_danos': 'Irritação respiratória, intoxicações, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Substâncias Químicas (líquidos e sólidos)', 'possiveis_danos': 'Queimaduras, irritações, intoxicações, dermatites, câncer.'},
        {'categoria': 'quimico', 'risco': 'Agrotóxicos', 'possiveis_danos': 'Intoxicações, dermatites, câncer.'},
        {'categoria': 'biologico', 'risco': 'Bactérias', 'possiveis_danos': 'Infecções, doenças infecciosas.'},
        {'categoria': 'biologico', 'risco': 'Fungos', 'possiveis_danos': 'Micoses, alergias, infecções respiratórias.'},
        {'categoria': 'biologico', 'risco': 'Vírus', 'possiveis_danos': 'Doenças virais, infecções.'},
        {'categoria': 'biologico', 'risco': 'Parasitas', 'possiveis_danos': 'Doenças parasitárias, infecções.'},
        {'categoria': 'biologico', 'risco': 'Protozoários', 'possiveis_danos': 'Doenças parasitárias.'},
        {'categoria': 'biologico', 'risco': 'Bacilos', 'possiveis_danos': 'Infecções diversas, como tuberculose.'},
        {'categoria': 'ergonomico', 'risco': 'Levantamento e Transporte Manual de Peso', 'possiveis_danos': 'Lesões musculoesqueléticas, dores na coluna.'},
        {'categoria': 'ergonomico', 'risco': 'Posturas Inadequadas', 'possiveis_danos': 'Dores musculares, lesões na coluna, LER/DORT.'},
        {'categoria': 'ergonomico', 'risco': 'Repetitividade', 'possiveis_danos': 'LER/DORT, tendinites, síndrome do túnel do carpo.'},
        {'categoria': 'ergonomico', 'risco': 'Jornada de Trabalho Prolongada', 'possiveis_danos': 'Fadiga, estresse, acidentes de trabalho.'},
        {'categoria': 'ergonomico', 'risco': 'Monotonia e Ritmo Excessivo', 'possiveis_danos': 'Estresse, fadiga mental, desmotivação.'},
        {'categoria': 'ergonomico', 'risco': 'Controle Rígido de Produtividade', 'possiveis_danos': 'Estresse, ansiedade, burnout.'},
        {'categoria': 'ergonomico', 'risco': 'Iluminação Inadequada', 'possiveis_danos': 'Fadiga visual, dores de cabeça.'},
        {'categoria': 'ergonomico', 'risco': 'Mobiliário Inadequado', 'possiveis_danos': 'Dores musculares, lesões na coluna.'},
        {'categoria': 'acidente', 'risco': 'Arranjo Físico Inadequado', 'possiveis_danos': 'Quedas, colisões, esmagamentos.'},
        {'categoria': 'acidente', 'risco': 'Máquinas e Equipamentos sem Proteção', 'possiveis_danos': 'Amputações, cortes, esmagamentos, prensamentos.'},
        {'categoria': 'acidente', 'risco': 'Ferramentas Inadequadas ou Defeituosas', 'possiveis_danos': 'Cortes, perfurações, fraturas.'},
        {'categoria': 'acidente', 'risco': 'Eletricidade', 'possiveis_danos': 'Choque elétrico, queimaduras, fibrilação ventricular.'},
        {'categoria': 'acidente', 'risco': 'Incêndio e Explosão', 'possiveis_danos': 'Queimaduras, asfixia, lesões por impacto.'},
        {'categoria': 'acidente', 'risco': 'Animais Peçonhentos', 'possiveis_danos': 'Picadas, mordidas, reações alérgicas, envenenamento.'},
        {'categoria': 'acidente', 'risco': 'Armazenamento Inadequado', 'possiveis_danos': 'Quedas de materiais, esmagamentos, soterramentos.'},
        {'categoria': 'acidente', 'risco': 'Trabalho em Altura', 'possiveis_danos': 'Quedas, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Espaços Confinados', 'possiveis_danos': 'Asfixia, intoxicações, explosões.'},
        {'categoria': 'acidente', 'risco': 'Condução de Veículos', 'possiveis_danos': 'Acidentes de trânsito, lesões diversas.'},
        {'categoria': 'acidente', 'risco': 'Projeção de Partículas', 'possiveis_danos': 'Lesões oculares, cortes na pele.'}
    ])

# --- Interface do Streamlit ---
st.markdown("""
<style>
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
    <p>Geração automática de OS a partir de um modelo Word (.docx) e uma planilha de funcionários.</p>
</div>
""", unsafe_allow_html=True)

# --- Sidebar ---
st.sidebar.markdown("### 📁 Arquivos Necessários")
arquivo_funcionarios = st.sidebar.file_uploader("1. Planilha de Funcionários (.xlsx)", type="xlsx")
arquivo_modelo_os = st.sidebar.file_uploader("2. Modelo de OS (.docx)", type="docx")
arquivo_logo = st.sidebar.file_uploader("3. Logo da Empresa (Opcional)", type=["png", "jpg", "jpeg"])

st.sidebar.markdown("---")
st.sidebar.info("É obrigatório carregar a planilha de funcionários e o modelo de OS para iniciar.")


# --- Lógica Principal da Interface ---
if arquivo_funcionarios is None or arquivo_modelo_os is None:
    st.markdown('<div class="info-box">📋 Por favor, carregue a <strong>Planilha de Funcionários</strong> e o <strong>Modelo de OS (.docx)</strong> na barra lateral para começar.</div>', unsafe_allow_html=True)
else:
    df_funcionarios_raw = carregar_planilha(arquivo_funcionarios)
    df_funcionarios = mapear_e_renomear_colunas_funcionarios(normalizar_colunas(df_funcionarios_raw.copy()))
    df_pgr = obter_dados_pgr()
    
    if df_funcionarios is not None:
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
            st.dataframe(df_final_filtrado[['nome_do_funcionario', 'setor', 'funcao']], use_container_width=True)
        
        st.markdown('<div class="section-header"><h3>⚠️ Configuração de Riscos</h3></div>', unsafe_allow_html=True)
        
        categorias = {'fisico': '🔥 Físicos', 'quimico': '⚗️ Químicos', 'biologico': '🦠 Biológicos', 'ergonomico': '🏃 Ergonômicos', 'acidente': '⚠️ Acidentes'}
        riscos_selecionados = []
        
        tabs = st.tabs(list(categorias.values()))
        for i, (key, nome) in enumerate(categorias.items()):
            with tabs[i]:
                riscos_categoria = df_pgr[df_pgr['categoria'] == key]['risco'].tolist()
                selecionados = st.multiselect(f"Selecione os riscos para {nome}:", options=riscos_categoria, key=f"riscos_{key}")
                riscos_selecionados.extend(selecionados)

        with st.expander("➕ Adicionar Risco Manual, EPIs e Medições (Opcional)"):
            perigo_manual = st.text_input("Descrição do Risco Manual")
            categoria_manual = st.selectbox("Categoria do Risco Manual", [""] + list(categorias.values()))
            danos_manuais = st.text_area("Possíveis Danos do Risco Manual")
            epis_manuais = st.text_area("EPIs Adicionais (separados por vírgula)")
            medicoes_manuais = st.text_area("Medições Ambientais (uma por linha)")
        
        st.markdown('<div class="section-header"><h3>🚀 Gerar Ordens de Serviço</h3></div>', unsafe_allow_html=True)
        
        if st.button("🔄 Gerar OSs para Funcionários Selecionados", type="primary"):
            if df_final_filtrado.empty:
                st.error("Nenhum funcionário selecionado! Ajuste os filtros.")
            else:
                with st.spinner("Gerando Ordens de Serviço... Por favor, aguarde."):
                    documentos_gerados = []
                    logo_path = None
                    
                    if arquivo_logo:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(arquivo_logo.name)[1]) as temp_logo:
                            temp_logo.write(arquivo_logo.getbuffer())
                            logo_path = temp_logo.name
                    
                    progress_bar = st.progress(0)
                    total = len(df_final_filtrado)
                    
                    for i, (_, func) in enumerate(df_final_filtrado.iterrows()):
                        nome_func = func.get('nome_do_funcionario', f'Func_{i+1}')
                        progress_bar.progress((i + 1) / total, text=f"Processando: {nome_func}")
                        try:
                            doc = gerar_os(
                                func, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais,
                                perigo_manual, danos_manuais, categoria_manual,
                                modelo_doc_carregado=arquivo_modelo_os, logo_path=logo_path
                            )
                            
                            doc_io = BytesIO()
                            doc.save(doc_io)
                            doc_io.seek(0)
                            
                            nome_limpo = "".join(c for c in nome_func if c.isalnum() or c in (' ', '_')).rstrip()
                            documentos_gerados.append((f"OS_{nome_limpo.replace(' ', '_')}.docx", doc_io.getvalue()))
                            
                        except Exception as e:
                            st.error(f"Erro ao gerar OS para {nome_func}: {e}")
                    
                    if logo_path: os.unlink(logo_path)
                    
                    if documentos_gerados:
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                            for nome, conteudo in documentos_gerados: zf.writestr(nome, conteudo)
                        zip_buffer.seek(0)
                        
                        st.markdown(f'<div class="success-box">✅ <strong>{len(documentos_gerados)} Ordens de Serviço</strong> geradas com sucesso!</div>', unsafe_allow_html=True)
                        st.download_button(
                            "📥 Baixar Todas as OSs (.zip)",
                            data=zip_buffer,
                            file_name=f"Ordens_de_Servico_{time.strftime('%Y%m%d')}.zip",
                            mime="application/zip",
                        )
