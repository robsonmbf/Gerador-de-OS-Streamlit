import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
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
    /* Oculta completamente a barra lateral e o botão de menu */
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
AGENTES_DE_RISCO = sorted(["Ruído (Contínuo ou Intermitente)", "Ruído (Impacto)", "Vibração de Corpo Inteiro", "Vibração de Mãos e Braços", "Radiações Ionizantes", "Radiações Não-Ionizantes", "Frio", "Calor", "Pressões Anormais", "Umidade", "Poeiras", "Fumos", "Névoas", "Neblinas", "Gases", "Vapores", "Produtos Químicos em Geral", "Vírus", "Bactérias", "Protozoários", "Fungos", "Parasitas", "Bacilos"])


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
        "[MEDIÇÕES]": medicoes_manuais or "Não aplicável",
    }
    substituir_placeholders(doc, contexto)
    return doc

def obter_dados_pgr():
    return pd.DataFrame([{'categoria': 'fisico', 'risco': 'Ruído', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'}, {'categoria': 'fisico', 'risco': 'Vibração', 'possiveis_danos': 'Doenças osteomusculares, problemas circulatórios.'}, {'categoria': 'fisico', 'risco': 'Calor', 'possiveis_danos': 'Desidratação, insolação, cãibras, exaustão, intermação.'}, {'categoria': 'fisico', 'risco': 'Frio', 'possiveis_danos': 'Hipotermia, congelamento, doenças respiratórias.'}, {'categoria': 'fisico', 'risco': 'Radiações Ionizantes', 'possiveis_danos': 'Câncer, mutações genéticas, queimaduras.'}, {'categoria': 'fisico', 'risco': 'Radiações Não Ionizantes', 'possiveis_danos': 'Queimaduras, lesões oculares, câncer de pele.'}, {'categoria': 'fisico', 'risco': 'Pressões Anormais', 'possiveis_danos': 'Doença descompressiva, barotrauma.'}, {'categoria': 'fisico', 'risco': 'Umidade', 'possiveis_danos': 'Doenças respiratórias, dermatites.'}, {'categoria': 'quimico', 'risco': 'Poeiras', 'possiveis_danos': 'Pneumoconioses, irritação respiratória, alergias.'}, {'categoria': 'quimico', 'risco': 'Fumos', 'possiveis_danos': 'Doenças respiratórias, intoxicações.'}, {'categoria': 'quimico', 'risco': 'Névoas', 'possiveis_danos': 'Irritação respiratória, dermatites.'}, {'categoria': 'quimico', 'risco': 'Gases', 'possiveis_danos': 'Asfixia, intoxicações, irritação respiratória.'}, {'categoria': 'quimico', 'risco': 'Vapores', 'possiveis_danos': 'Irritação respiratória, intoxicações, dermatites.'}, {'categoria': 'quimico', 'risco': 'Substâncias Químicas (líquidos e sólidos)', 'possiveis_danos': 'Queimaduras, irritações, intoxicações, dermatites, câncer.'}, {'categoria': 'quimico', 'risco': 'Agrotóxicos', 'possiveis_danos': 'Intoxicações, dermatites, câncer.'}, {'categoria': 'biologico', 'risco': 'Bactérias', 'possiveis_danos': 'Infecções, doenças infecciosas.'}, {'categoria': 'biologico', 'risco': 'Fungos', 'possiveis_danos': 'Micoses, alergias, infecções respiratórias.'}, {'categoria': 'biologico', 'risco': 'Vírus', 'possiveis_danos': 'Doenças virais, infecções.'}, {'categoria': 'biologico', 'risco': 'Parasitas', 'possiveis_danos': 'Doenças parasitárias, infecções.'}, {'categoria': 'biologico', 'risco': 'Protozoários', 'possiveis_danos': 'Doenças parasitárias.'}, {'categoria': 'biologico', 'risco': 'Bacilos', 'possiveis_danos': 'Infecções diversas, como tuberculose.'}, {'categoria': 'ergonomico', 'risco': 'Levantamento e Transporte Manual de Peso', 'possiveis_danos': 'Lesões musculoesqueléticas, dores na coluna.'}, {'categoria': 'ergonomico', 'risco': 'Posturas Inadequadas', 'possiveis_danos': 'Dores musculares, lesões na coluna, LER/DORT.'}, {'categoria': 'ergonomico', 'risco': 'Repetitividade', 'possiveis_danos': 'LER/DORT, tendinites, síndrome do túnel do carpo.'}, {'categoria': 'ergonomico', 'risco': 'Jornada de Trabalho Prolongada', 'possiveis_danos': 'Fadiga, estresse, acidentes de trabalho.'}, {'categoria': 'ergonomico', 'risco': 'Monotonia e Ritmo Excessivo', 'possiveis_danos': 'Estresse, fadiga mental, desmotivação.'}, {'categoria': 'ergonomico', 'risco': 'Controle Rígido de Produtividade', 'possiveis_danos': 'Estresse, ansiedade, burnout.'}, {'categoria': 'ergonomico', 'risco': 'Iluminação Inadequada', 'possiveis_danos': 'Fadiga visual, dores de cabeça.'}, {'categoria': 'ergonomico', 'risco': 'Mobiliário Inadequado', 'possiveis_danos': 'Dores musculares, lesões na coluna.'}, {'categoria': 'acidente', 'risco': 'Arranjo Físico Inadequado', 'possiveis_danos': 'Quedas, colisões, esmagamentos.'}, {'categoria': 'acidente', 'risco': 'Máquinas e Equipamentos sem Proteção', 'possiveis_danos': 'Amputações, cortes, esmagamentos, prensamentos.'}, {'categoria': 'acidente', 'risco': 'Ferramentas Inadequadas ou Defeituosas', 'possiveis_danos': 'Cortes, perfurações, fraturas.'}, {'categoria': 'acidente', 'risco': 'Eletricidade', 'possiveis_danos': 'Choque elétrico, queimaduras, fibrilação ventricular.'}, {'categoria': 'acidente', 'risco': 'Incêndio e Explosão', 'possiveis_danos': 'Queimaduras, asfixia, lesões por impacto.'}, {'categoria': 'acidente', 'risco': 'Animais Peçonhentos', 'possiveis_danos': 'Picadas, mordidas, reações alérgicas, envenenamento.'}, {'categoria': 'acidente', 'risco': 'Armazenamento Inadequado', 'possiveis_danos': 'Quedas de materiais, esmagamentos, soterramentos.'}, {'categoria': 'acidente', 'risco': 'Trabalho em Altura', 'possiveis_danos': 'Quedas, fraturas, morte.'}, {'categoria': 'acidente', 'risco': 'Espaços Confinados', 'possiveis_danos': 'Asfixia, intoxicações, explosões.'}, {'categoria': 'acidente', 'risco': 'Condução de Veículos', 'possiveis_danos': 'Acidentes de trânsito, lesões diversas.'}, {'categoria': 'acidente', 'risco': 'Projeção de Partículas', 'possiveis_danos': 'Lesões oculares, cortes na pele.'}])

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
    df_pgr = obter_dados_pgr()

    def formatar_setor(setor):
        return f"{setor} ✅" if setor in st.session_state.setores_concluidos else setor
    def formatar_cargo(cargo):
        return f"{cargo} ✅" if cargo in st.session_state.cargos_concluidos else cargo

    with st.container():
        st.markdown('### 👥 Seleção de Funcionários')
        if st.button("Limpar Indicadores de Conclusão (✅)"):
            st.session_state.setores_concluidos.clear()
            st.session_state.cargos_concluidos.clear()
            st.rerun()

        setores = sorted(df_funcionarios['setor'].dropna().unique().tolist()) if 'setor' in df_funcionarios.columns else []
        setor_sel = st.multiselect("Filtrar por Setor(es)", setores, format_func=formatar_setor)
        
        df_filtrado_setor = df_funcionarios[df_funcionarios['setor'].isin(setor_sel)] if setor_sel else df_funcionarios
        funcoes = sorted(df_filtrado_setor['funcao'].dropna().unique().tolist()) if 'funcao' in df_filtrado_setor.columns else []
        funcao_sel = st.multiselect("Filtrar por Função/Cargo(s)", funcoes, format_func=formatar_cargo)
        
        df_final_filtrado = df_filtrado_setor[df_filtrado_setor['funcao'].isin(funcao_sel)] if funcao_sel else df_filtrado_setor
        st.success(f"✅ {len(df_final_filtrado)} funcionários selecionados para o próximo lote.")
        
        colunas_desejadas = ['nome_do_funcionario', 'setor', 'funcao']
        colunas_existentes = [col for col in colunas_desejadas if col in df_final_filtrado.columns]
        if not colunas_existentes:
            st.error("❌ Nenhuma das colunas essenciais (nome, setor, função) foi encontrada.")
        else:
            st.dataframe(df_final_filtrado[colunas_existentes], use_container_width=True)
            st.markdown('### ⚠️ Configuração de Riscos')
            categorias = {'fisico': '🔥 Físicos', 'quimico': '⚗️ Químicos', 'biologico': '🦠 Biológicos', 'ergonomico': '🏃 Ergonômicos', 'acidente': '⚠️ Acidentes'}
            riscos_selecionados = []
            tabs = st.tabs(list(categorias.values()))
            for i, (key, nome) in enumerate(categorias.items()):
                with tabs[i]:
                    riscos_categoria = df_pgr[df_pgr['categoria'] == key]['risco'].tolist()
                    selecionados = st.multiselect(f"Selecione os riscos:", options=riscos_categoria, key=f"riscos_{key}")
                    riscos_selecionados.extend(selecionados)
            if riscos_selecionados:
                st.info(f"**Resumo de Riscos Selecionados ({len(riscos_selecionados)} no total)**")
                riscos_categorizados_para_display = {}
                for risco_nome in sorted(riscos_selecionados):
                    categoria_key_series = df_pgr[df_pgr['risco'] == risco_nome]['categoria']
                    if not categoria_key_series.empty:
                        categoria_key = categoria_key_series.iloc[0]
                        categoria_display = categorias.get(categoria_key, "Outros")
                        if categoria_display not in riscos_categorizados_para_display:
                            riscos_categorizados_para_display[categoria_display] = []
                        riscos_categorizados_para_display[categoria_display].append(risco_nome)
                for categoria, lista_riscos in riscos_categorizados_para_display.items():
                    with st.container():
                        st.markdown(f"**{categoria}**")
                        for risk in lista_riscos:
                            st.markdown(f"&nbsp;&nbsp;&nbsp; - {risk}")

            # --- NOMENCLATURA CORRIGIDA ---
            with st.expander("➕ Inclusão de Risco Manual"):
                def adicionar_risco_manual():
                    risco = st.session_state.risco_input
                    categoria = st.session_state.categoria_input
                    danos = st.session_state.danos_input
                    if risco and categoria:
                        st.session_state.riscos_manuais_adicionados.append({"risco": risco,"categoria": categoria,"danos": danos})
                        st.session_state.risco_input = ""; st.session_state.danos_input = ""
                    else: st.warning("Preencha a Descrição e a Categoria do Risco para adicioná-lo.")
                def limpar_riscos_manuais(): st.session_state.riscos_manuais_adicionados = []
                st.text_input("Descrição do Risco", key="risco_input")
                st.selectbox("Categoria do Risco", [""] + list(categorias.values()), key="categoria_input")
                st.text_area("Possíveis Danos", key="danos_input")
                c1, c2, _ = st.columns([1,1,2])
                with c1: st.button("Adicionar Risco", on_click=adicionar_risco_manual)
                with c2: st.button("Limpar Lista de Riscos", on_click=limpar_riscos_manuais)
                if st.session_state.riscos_manuais_adicionados:
                    st.write("**Riscos Manuais Adicionados:**")
                    for r in st.session_state.riscos_manuais_adicionados:
                        st.markdown(f"- **{r['risco']}** ({r['categoria']}): {r['danos']}")

            with st.expander("🦺 Adicionar EPIs"):
                def adicionar_epi():
                    epi = st.session_state.epi_input
                    if epi and epi.strip():
                        st.session_state.epis_adicionados.append(epi.strip())
                        st.session_state.epi_input = ""
                def limpar_epis():
                    st.session_state.epis_adicionados = []
                st.text_input("Digite um EPI para adicionar à lista", key="epi_input")
                c1_epi, c2_epi, _ = st.columns([1,1,2])
                with c1_epi: st.button("Adicionar EPI", on_click=adicionar_epi)
                with c2_epi: st.button("Limpar Lista de EPIs", on_click=limpar_epis)
                if st.session_state.epis_adicionados:
                    st.write("**EPIs Adicionados:**")
                    for epi_item in st.session_state.epis_adicionados:
                        st.markdown(f"- {epi_item}")

            with st.expander("📊 Adicionar Medições Ambientais"):
                def adicionar_medicao():
                    agente = st.session_state.agente_input
                    valor = st.session_state.valor_input
                    unidade = st.session_state.unidade_input
                    if agente and valor:
                        medicao_str = f"{agente}: {valor} {unidade}"
                        st.session_state.medicoes_adicionadas.append(medicao_str)
                        st.session_state.agente_input = ""; st.session_state.valor_input = ""
                    else: st.warning("Preencha o Agente e o Valor para adicionar uma medição.")
                def limpar_medicoes(): st.session_state.medicoes_adicionadas = []
                col1, col2, col3 = st.columns([2,1,1])
                with col1: st.selectbox("Agente/Fonte do Risco", options=[""] + AGENTES_DE_RISCO, key="agente_input")
                with col2: st.text_input("Valor Medido", key="valor_input")
                with col3: st.selectbox("Unidade de Medida", UNIDADES_DE_MEDIDA, key="unidade_input")
                col_btn1, col_btn2, _ = st.columns([1,1,2])
                with col_btn1: st.button("Adicionar Medição", on_click=adicionar_medicao)
                with c2: st.button("Limpar Lista de Medições", on_click=limpar_medicoes)
                if st.session_state.medicoes_adicionadas:
                    st.write("**Medições Adicionadas:**")
                    for med in st.session_state.medicoes_adicionadas:
                        st.markdown(f"- {med}")
            
            if st.button("🚀 Gerar OSs para Funcionários Selecionados", type="primary"):
                epis_finais = ", ".join(st.session_state.epis_adicionados)
                medicoes_finais = "\n".join(st.session_state.medicoes_adicionadas)
                riscos_manuais_finais = st.session_state.riscos_manuais_adicionados
                with st.spinner("Gerando documentos..."):
                    documentos_gerados = []
                    os_geradas_info_batch = [] 
                    for _, func in df_final_filtrado.iterrows():
                        doc = gerar_os(func, df_pgr, riscos_selecionados, epis_finais, medicoes_finais, riscos_manuais_finais, arquivo_modelo_os)
                        doc_io = BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        nome_limpo = re.sub(r'[^\w\s-]', '', func.get("nome_do_funcionario", "Func_Sem_Nome")).strip().replace(" ", "_")
                        documentos_gerados.append((f"OS_{nome_limpo}.docx", doc_io.getvalue()))
                        os_geradas_info_batch.append({'Funcionário': func.get("nome_do_funcionario", "N/A"),'Setor': func.get("setor", "N/A"),'Cargo/Função': func.get("funcao", "N/A")})

                    if documentos_gerados:
                        # Marcar setores e funções selecionados como concluídos
                        for setor in setor_sel:
                            st.session_state.setores_concluidos.add(setor)
                        for funcao in funcao_sel:
                            st.session_state.cargos_concluidos.add(funcao)
                        
                        df_resumo_batch = pd.DataFrame(os_geradas_info_batch)
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for nome_arquivo, conteudo_doc in documentos_gerados:
                                zip_file.writestr(nome_arquivo, conteudo_doc)
                        # Criar nome do arquivo baseado nas seleções
                        setores_nome = "_".join(setor_sel[:2]) if len(setor_sel) <= 2 else f"{setor_sel[0]}_e_outros"
                        funcoes_nome = "_".join(funcao_sel[:2]) if len(funcao_sel) <= 2 else f"{funcao_sel[0]}_e_outros"
                        nome_arquivo_zip = f"OS_{setores_nome}_{funcoes_nome}_{time.strftime('%Y%m%d')}.zip".replace(' ', '_')
                        
                        st.success(f"🎉 {len(documentos_gerados)} Ordens de Serviço geradas!")
                        st.download_button(label="📥 Baixar Lote Atual (.zip)", data=zip_buffer.getvalue(), file_name=nome_arquivo_zip, mime="application/zip")
                        with st.expander("📄 Resumo do Lote Gerado", expanded=True):
                            st.dataframe(df_resumo_batch, use_container_width=True)
                        
                        # --- LINHA REMOVIDA PARA GARANTIR O FUNCIONAMENTO DO DOWNLOAD ---
                        # st.rerun()
