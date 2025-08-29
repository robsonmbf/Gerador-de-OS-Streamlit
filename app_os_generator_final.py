import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import os
import zipfile
from io import BytesIO
import tempfile
import time
import re

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Ordens de Serviço (OS)",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- Funções de Lógica de Negócio ---

def normalizar_texto(texto):
    """Função auxiliar para limpar e padronizar strings."""
    if not isinstance(texto, str):
        return ""
    texto = texto.lower().strip()
    texto = re.sub(r'[^a-z0-9]', '', texto) # Remove acentos, espaços e caracteres especiais
    return texto

def mapear_e_renomear_colunas_funcionarios(df):
    """Mapeia e renomeia colunas da planilha de funcionários de forma mais robusta."""
    mapeamento = {
        'nome_do_funcionario': ['nome', 'nomedofuncionario', 'funcionario', 'colaborador', 'nomecompleto'],
        'funcao': ['funcao', 'funçao', 'cargo'],
        'data_de_admissao': ['datadeadmissao', 'dataadmissao', 'admissao'],
        'setor': ['setor', 'area', 'departamento'],
        'descricao_de_atividades': ['descricaodeatividades', 'atividades', 'descricaoatividades', 'tarefas', 'descricaodastarefas'],
        'empresa': ['empresa'],
        'unidade': ['unidade']
    }
    
    colunas_renomeadas = {}
    colunas_df = {normalizar_texto(col): col for col in df.columns}

    for nome_padrao, nomes_possiveis in mapeamento.items():
        for nome_possivel in nomes_possiveis:
            if nome_possivel in colunas_df:
                coluna_original = colunas_df[nome_possivel]
                colunas_renomeadas[coluna_original] = nome_padrao
                break # Pula para o próximo nome padrão assim que encontrar uma correspondência
                
    df.rename(columns=colunas_renomeadas, inplace=True)
    return df

@st.cache_data
def carregar_planilha(arquivo):
    """Carrega uma planilha do Excel."""
    if arquivo is None:
        return None
    try:
        return pd.read_excel(arquivo)
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
    """Substitui os placeholders em todo o documento."""
    for p in doc.paragraphs:
        replace_text_in_paragraph(p, contexto)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_text_in_paragraph(p, contexto)

def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, perigo_manual, danos_manuais, categoria_manual, modelo_doc_carregado, logo_path=None):
    """Gera uma única Ordem de Serviço."""
    doc = Document(modelo_doc_carregado)

    if logo_path:
        try:
            doc.tables[0].cell(0, 0).paragraphs[0].add_run().add_picture(logo_path, width=Inches(2.0))
        except Exception:
            st.warning("Aviso: Não foi possível inserir a logo na tabela do cabeçalho. Verifique o modelo.")
            
    riscos_info = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]
    
    riscos_por_categoria = {"fisico": [], "quimico": [], "biologico": [], "ergonomico": [], "acidente": []}
    danos_por_categoria = {"fisico": [], "quimico": [], "biologico": [], "ergonomico": [], "acidente": []}
    
    for _, risco_row in riscos_info.iterrows():
        categoria = str(risco_row.get("categoria", "")).lower()
        if categoria in riscos_por_categoria:
            riscos_por_categoria[categoria].append(str(risco_row.get("risco", "")))
            danos = risco_row.get("possiveis_danos")
            if pd.notna(danos):
                danos_por_categoria[categoria].append(str(danos))

    if perigo_manual and categoria_manual:
        map_categorias = {"Físicos": "fisico", "Químicos": "quimico", "Biológicos": "biologico", "Ergonômicos": "ergonomico", "Acidentes": "acidente"}
        categoria_alvo = map_categorias.get(categoria_manual)
        if categoria_alvo:
            riscos_por_categoria[categoria_alvo].append(perigo_manual)
            if danos_manuais:
                danos_por_categoria[categoria_alvo].append(danos_manuais)

    epis_recomendados = set(epi.strip() for epi in epis_manuais.split(',') if epi.strip())

    # CORREÇÃO: Trata a data de admissão e descrição de atividades
    data_admissao = "Não informado"
    if 'data_de_admissao' in funcionario and pd.notna(funcionario['data_de_admissao']):
        try:
            data_admissao = pd.to_datetime(funcionario['data_de_admissao']).strftime('%d/%m/%Y')
        except Exception:
            data_admissao = str(funcionario['data_de_admissao'])

    descricao_atividades = "Não informado"
    if 'descricao_de_atividades' in funcionario and pd.notna(funcionario['descricao_de_atividades']):
        descricao_atividades = str(funcionario['descricao_de_atividades'])

    def tratar_lista_vazia(lista, separador=", "):
        if not lista or all(not item.strip() for item in lista):
            return "Não identificado"
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
        "[MEDIÇÕES]": medicoes_manuais or "Não aplicável",
    }
    
    substituir_placeholders(doc, contexto)
    return doc

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
# (O restante do código da interface continua o mesmo)
st.markdown("""<style>...</style>""", unsafe_allow_html=True) # CSS omitido para brevidade
st.markdown("""<div class="main-header">...</div>""", unsafe_allow_html=True) # Header omitido

st.sidebar.markdown("### 📁 Arquivos Necessários")
arquivo_funcionarios = st.sidebar.file_uploader("1. Planilha de Funcionários (.xlsx)", type="xlsx")
arquivo_modelo_os = st.sidebar.file_uploader("2. Modelo de OS (.docx)", type="docx")
arquivo_logo = st.sidebar.file_uploader("3. Logo da Empresa (Opcional)", type=["png", "jpg", "jpeg"])
st.sidebar.info("É obrigatório carregar a planilha e o modelo de OS para iniciar.")

if not arquivo_funcionarios or not arquivo_modelo_os:
    st.info("📋 Por favor, carregue a Planilha de Funcionários e o Modelo de OS (.docx) na barra lateral para começar.")
else:
    df_funcionarios_raw = carregar_planilha(arquivo_funcionarios)
    df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_raw)
    df_pgr = obter_dados_pgr()

    st.markdown('<div class="section-header"><h3>👥 Seleção de Funcionários</h3></div>', unsafe_allow_html=True)
    # (Filtros e exibição da tabela)
    setores = ["Todos"] + (df_funcionarios['setor'].dropna().unique().tolist() if 'setor' in df_funcionarios.columns else [])
    setor_sel = st.selectbox("Filtrar por Setor", setores)
    df_filtrado = df_funcionarios[df_funcionarios['setor'] == setor_sel] if setor_sel != "Todos" else df_funcionarios
    funcoes = ["Todos"] + (df_filtrado['funcao'].dropna().unique().tolist() if 'funcao' in df_filtrado.columns else [])
    funcao_sel = st.selectbox("Filtrar por Função/Cargo", funcoes)
    df_final_filtrado = df_filtrado[df_filtrado['funcao'] == funcao_sel] if funcao_sel != "Todos" else df_filtrado
    
    st.success(f"✅ {len(df_final_filtrado)} funcionários selecionados.")
    if 'nome_do_funcionario' in df_final_filtrado.columns:
        st.dataframe(df_final_filtrado[['nome_do_funcionario', 'setor', 'funcao']], use_container_width=True)

    # (Seleção de riscos)
    st.markdown('<div class="section-header"><h3>⚠️ Configuração de Riscos</h3></div>', unsafe_allow_html=True)
    categorias = {'fisico': '🔥 Físicos', 'quimico': '⚗️ Químicos', 'biologico': '🦠 Biológicos', 'ergonomico': '🏃 Ergonômicos', 'acidente': '⚠️ Acidentes'}
    riscos_selecionados = []
    tabs = st.tabs(list(categorias.values()))
    for i, (key, nome) in enumerate(categorias.items()):
        with tabs[i]:
            riscos_categoria = df_pgr[df_pgr['categoria'] == key]['risco'].tolist()
            selecionados = st.multiselect(f"Selecione os riscos:", options=riscos_categoria, key=f"riscos_{key}")
            riscos_selecionados.extend(selecionados)

    with st.expander("➕ Adicionar Risco Manual, EPIs e Medições"):
        # ... (Inputs manuais)
        perigo_manual = st.text_input("Descrição do Risco Manual")
        categoria_manual = st.selectbox("Categoria do Risco Manual", [""] + list(categorias.values()))
        danos_manuais = st.text_area("Possíveis Danos do Risco Manual")
        epis_manuais = st.text_area("EPIs (separados por vírgula)")
        medicoes_manuais = st.text_area("Medições (uma por linha)")

    if st.button("🚀 Gerar OSs para Funcionários Selecionados", type="primary"):
        # (Lógica do botão de gerar)
        with st.spinner("Gerando documentos..."):
            # ... (código de geração e zip)
            documentos_gerados = []
            logo_path = None
            if arquivo_logo:
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(arquivo_logo.name)[1]) as temp:
                    temp.write(arquivo_logo.getbuffer())
                    logo_path = temp.name

            for _, func in df_final_filtrado.iterrows():
                doc = gerar_os(func, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, perigo_manual, danos_manuais, categoria_manual, arquivo_modelo_os, logo_path)
                doc_io = BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                nome_limpo = re.sub(r'[^a-zA-Z0-9_]', '', func.get("nome_do_funcionario", "Funcionario_Sem_Nome").replace(" ", "_"))
                documentos_gerados.append((f"OS_{nome_limpo}.docx", doc_io.getvalue()))

            if logo_path: os.unlink(logo_path)

            if documentos_gerados:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for nome, conteudo in documentos_gerados: zf.writestr(nome, conteudo)
                
                st.success(f"🎉 {len(documentos_gerados)} Ordens de Serviço geradas!")
                st.download_button("📥 Baixar Todas as OSs (.zip)", data=zip_buffer, file_name=f"Ordens_de_Servico_{time.strftime('%Y%m%d')}.zip", mime="application/zip")
