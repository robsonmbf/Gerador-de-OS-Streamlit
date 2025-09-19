importar streamlit como st
importar pandas como pd
de docx importar documento
de docx.shared importar Pt
importar arquivo zip
de io importar BytesIO
tempo de importação
importar re
importar sistema
importar sistema operacional

# Adicione o diretório atual ao caminho para importar módulos locais
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

de database.models importar DatabaseManager
de database.auth importar AuthManager
de database.user_data importar UserDataManager

# --- Configuração da página ---
st.set_page_config(
    page_title="Gerador de Ordens de Serviço (OS)",
    ícone_de_página="ðŸ“„",
    layout="amplo",
)

# --- DEFINIÇÃO DE CONSTANTES GLOBAIS ---
UNIDADES_DE_MEDIDA = ["dB(A)", "m/sÂ²", "ppm", "mg/mÂ³", "%", "Â°C", "lx", "cal/cmÂ²", "ÂµT", "kV/m", "W/mÂ²", "f/cmÂ³", "Não aplicável"]
AGENTES_DE_RISCO = ordenado([
    "Ruído (Contínuo ou Intermitente)", "Ruído (Impacto)", "Vibração de Corpo Inteiro", "Vibração de Mãos e Braços",
    "Radiações Ionizantes", "Radiações Não-Ionizantes", "Frio", "Calor", "Pressões Anormais", "Umidade", "Poeiras",
    "Fumos", "Névoas", "Neblinas", "Gases", "Vapores", "Produtos Químicos em Geral", "Vírus", "Bactérias",
    "Protozoários", "Fungos", "Parasitas", "Bacilos"
])
CATEGORIAS_RISCO = {'físico': 'ðŸ”¥ Físicos', 'químico': 'âš—ï¸ Químicos', 'biológico': 'ðŸ¦ Biológicos', 'ergonômico': 'ðŸ ƒ Ergonômicos', 'acidente': 'âš ï¸ Acidentes'}

# --- Inicialização dos Gerenciadores ---
@st.cache_resource
definição init_managers():
    db_manager = Gerenciador de Banco de Dados()
    auth_manager = Gerenciador de Autenticação(gerenciador_de_banco_de_dados)
    user_data_manager = Gerenciador de Dados do Usuário(gerenciador_de_banco_de_dados)
    retornar db_manager, auth_manager, user_data_manager

gerenciador_de_banco_de_dados, gerenciador_de_autenticação, gerenciador_de_dados_do_usuário = gerenciadores_de_inicialização()

# --- CSS PERSONALIZADO ---
st.markdown("""
<estilo>
    [data-testid="stSidebar"] {
        exibição: nenhuma;
    }
    .cabeçalho-principal {
        alinhamento de texto: centro;
        preenchimento inferior: 20px;
    }
    .auth-container {
        largura máxima: 400px;
        margem: 0 automático;
        preenchimento: 2rem;
        borda: 1px sólido #ddd;
        raio da borda: 10px;
        cor de fundo: #f9f9f9;
    }
    .informações do usuário {
        cor de fundo: #262730;
        cor: branco;            
        preenchimento: 1rem;
        raio da borda: 5px;
        margem inferior: 1rem;
        borda: 1px sólido #3DD56D;
    }
</estilo>
""", unsafe_allow_html=Verdadeiro)

# --- FUNÇÕES DE AUTENTICAÇAO E LÓGICA DE NEGÓCIO ---
definição show_login_page():
    st.markdown("""<div class="main-header"><h1>ðŸ” Acesso ao Sistema</h1><p>Faça login ou registre-se para acessar o Gerador de SO</p></div>""", unsafe_allow_html=True)
    tab1, tab2 = st.tabs(["Login", "Registro"])
    com tab1:
        com st.form("login_form"):
            email = st.text_input("E-mail", espaço reservado="seu@email.com")
            senha = st.text_input("Senha", type="senha")
            se st.form_submit_button("Entrar", use_container_width=True):
                se e-mail e senha:
                    sucesso, mensagem, dados_da_sessão = auth_manager.login_user(email, senha)
                    se sucesso:
                        st.session_state.authenticated = Verdadeiro
                        st.session_state.user_data = dados_da_sessão
                        st.session_state.user_data_loaded = Falso
                        st.success(mensagem)
                        st.rerun()
                    outro:
                        st.error(mensagem)
                outro:
                    st.error("Por favor, preencha todos os campos")
    com tab2:
        com st.form("register_form"):
            reg_email = st.text_input("E-mail", placeholder="seu@email.com", key="reg_email")
            reg_password = st.text_input("Senha", type="senha", key="reg_password")
            reg_password_confirm = st.text_input("Confirmar Senha", type="password")
            se st.form_submit_button("Registrador", use_container_width=True):
                se reg_email e reg_password e reg_password_confirm:
                    se reg_password != reg_password_confirm:
                        st.error("As senhas não coincidem")
                    outro:
                        sucesso, mensagem = auth_manager.register_user(reg_email, reg_password)
                        se sucesso:
                            st.success(mensagem)
                            st.info("Agora você pode fazer login com suas credenciais")
                        outro:
                            st.error(mensagem)
                outro:
                    st.error("Por favor, preencha todos os campos")

definição check_authentication():
    se 'autenticado' não estiver em st.session_state:
        st.session_state.authenticated = Falso
    se 'user_data' não estiver em st.session_state:
        st.session_state.user_data = Nenhum
    se st.session_state.authenticated e st.session_state.user_data:
        token_de_sessão = st.session_state.user_data.get('token_de_sessão')
        se session_token:
            é_válido, _ = auth_manager.validate_session(token_de_sessão)
            se não for is_valid:
                st.session_state.authenticated = Falso
                st.session_state.user_data = Nenhum
                st.rerun()

def logout_user():
    se st.session_state.user_data e st.session_state.user_data.get('session_token'):
        auth_manager.logout_user(st.session_state.user_data['token_de_sessão'])
    st.session_state.authenticated = Falso
    st.session_state.user_data = Nenhum
    st.session_state.user_data_loaded = Falso
    st.rerun()

definição show_user_info():
    se st.session_state.get('autenticado'):
        user_email = st.session_state.user_data.get('email', 'N/D')
        col1, col2 = st.columns([3, 1])
        com col1:
            st.markdown(f'<div class="user-info">ðŸ'¤ <strong>Usuário:</strong> {user_email}</div>', unsafe_allow_html=True)
        com col2:
            se st.button("Sair", tipo="secundário"):
                logout_user()

definição init_user_session_state():
    se st.session_state.get('authenticated') e não st.session_state.get('user_data_loaded'):
        user_id = st.session_state.user_data.get('user_id')
        se user_id:
            st.session_state.medicoes_adicionadas = user_data_manager.get_user_measurements(user_id)
            st.session_state.epis_adicionados = gerenciador_de_dados_do_usuário.get_user_epis(id_do_usuário)
            st.session_state.riscos_manuais_adicionados = user_data_manager.get_user_manual_risks(user_id)
            st.session_state.user_data_loaded = Verdadeiro
    
    se 'medicoes_adicionadas' não estiver em st.session_state:
        st.session_state.medicoes_adicionadas = []
    se 'epis_adicionados' não estiver em st.session_state:
        st.session_state.epis_adicionados = []
    se 'riscos_manuais_adicionados' não estiver em st.session_state:
        st.session_state.riscos_manuais_adicionados = []
    se 'cargos_concluidos' não estiver em st.session_state:
        st.session_state.cargos_concluidos = set()

def normalizar_texto(texto):
    se não isinstance(texto, str): retorne ""
    retornar re.sub(r'[\s\W_]+', '', texto.lower().strip())

def mapear_e_renomear_colunas_funcionarios(df):
    df_copia = df.copy()
    mapeamento = {
        'nome_do_funcionario': ['nomedofuncionario', 'nome', 'funcionario', 'funcionário', 'colaborador', 'nomecompleto'],
        'função': ['função', 'função', 'carga'],
        'data_de_admissao': ['datadeadmissao', 'dataadmissao', 'admissao', 'admissao'],
        'setor': ['setordetrabalho', 'setor', 'área', 'Área', 'departamento'],
        'descricao_de_atividades': ['descricaodeatividades', 'atividades', 'descricaoatividades', 'descricaodeatividades', 'tarefas', 'descricaodastarefas'],
        'empresa': ['empresa'],
        'unidade': ['unidade']
    }
    colunas_renomeadas = {}
    colunas_df_normalizadas = {normalizar_texto(col): col para col em df_copia.columns}
    para nome_padrao, nomes_possiveis em mapeamento.items():
        para nome_possivel em nomes_possiveis:
            se nome_possivel em colunas_df_normalizadas:
                coluna_original = colunas_df_normalizadas[nome_possivel]
                colunas_renomeadas[coluna_original] = nome_padrao
                quebrar
    df_copia.rename(colunas=colunas_renomeadas, inplace=True)
    retornar df_copia

@st.cache_data
def carregar_planilha(arquivo):
    se arquivo for Nenhum: retorne Nenhum
    tentar:
        retornar pd.read_excel(arquivo)
    exceto Exceção como e:
        st.error(f"Erro ao ler o arquivo Excel: {e}")
        retornar Nenhum

@st.cache_data
def obter_dados_pgr():
    dados = [
        {'categoria': 'físico', 'risco': 'Ruído (Contínuo ou Intermitente)', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'},
        {'categoria': 'físico', 'risco': 'Ruído (Impacto)', 'possiveis_danos': 'Perda auditiva, trauma acústico.'},
        {'categoria': 'físico', 'risco': 'Vibração de Corpo Inteiro', 'possiveis_danos': 'Problemas na coluna, dores lombares.'},
        {'categoria': 'fisico', 'risco': 'Vibração de Mãos e Braços', 'possiveis_danos': 'Doenças osteomusculares, problemas circulatórios.'},
        {'categoria': 'físico', 'risco': 'Calor', 'possiveis_danos': 'Desidratação, insolação, céibras, exaustão, interação.'},
        {'categoria': 'físico', 'risco': 'Frio', 'possiveis_danos': 'Hipotermia, congelamento, doenças respiratórias.'},
        {'categoria': 'físico', 'risco': 'Radiações Ionizantes', 'possiveis_danos': 'Câncer, mutações genéticas, queimaduras.'},
        {'categoria': 'físico', 'risco': 'Radiações Não-Ionizantes', 'possiveis_danos': 'Queimaduras, lesões oculares, câncer de pele.'},
        {'categoria': 'físico', 'risco': 'Pressões Anormais', 'possiveis_danos': 'Doença descompressiva, barotrauma.'},
        {'categoria': 'físico', 'risco': 'Umidade', 'possiveis_danos': 'Doenças respiratórias, dermatites, micoses.'},
        {'categoria': 'quimico', 'risco': 'Poeiras', 'possiveis_danos': 'Pneumoconioses (silicose, asbestose), irritação respiratória.'},
        {'categoria': 'quimico', 'risco': 'Fumos', 'possiveis_danos': 'Doenças respiratórias (febre dos fumos metálicos), intoxicações.'},
        {'categoria': 'quimico', 'risco': 'Névoas', 'possiveis_danos': 'Irritação respiratória, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Gases', 'possiveis_danos': 'Asfixia, intoxicações, irritação respiratória.'},
        {'categoria': 'quimico', 'risco': 'Vapores', 'possiveis_danos': 'Irritação respiratória, intoxicações, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Produtos Químicos em Geral', 'possiveis_danos': 'Queimaduras, irritações, intoxicações, dermatites, câncer.'},
        {'categoria': 'biológico', 'risco': 'Bactérias', 'possiveis_danos': 'Infecções, doenças infecciosas (tétano, tuberculose).'},
        {'categoria': 'biológico', 'risco': 'Fungos', 'possiveis_danos': 'Micoses, alergias, infecções respiratórias.'},
        {'categoria': 'biológico', 'risco': 'Vírus', 'possiveis_danos': 'Doenças virais (hepatite, HIV), infecções.'},
        {'categoria': 'ergonômico', 'risco': 'Levantamento e Transporte Manual de Peso', 'possiveis_danos': 'Lesões musculoesqueléticas, dores na coluna.'},
        {'categoria': 'ergonômico', 'risco': 'Posturas Inadequadas', 'possiveis_danos': 'Dores musculares, lesões na coluna, LER/DORT.'},
        {'categoria': 'ergonômico', 'risco': 'Repetitividade', 'possiveis_danos': 'LER/DORT, tendinites, síndrome do túnel do carpo.'},
        {'categoria': 'acidente', 'risco': 'Máquinas e Equipamentos sem Proteção', 'possiveis_danos': 'Amputações, cortes, esmagamentos, prensamentos.'},
        {'categoria': 'acidente', 'risco': 'Eletricidade', 'possiveis_danos': 'Choque elétrico, queimaduras, fibrilação ventricular.'},
        {'categoria': 'acidente', 'risco': 'Trabalho em Altura', 'possiveis_danos': 'Quedas, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Projeção de Partículas', 'possiveis_danos': 'Lesões oculares, cortes na pele.'}
    ]
    retornar pd.DataFrame(dados)

def substituir_placeholders(doc, contexto):
    para tabela em doc.tables:
        para linha em table.rows:
            para célula em row.cells:
                para p em célula.parágrafos:
                    # Usando uma abordagem mais simples e direta
                    inline = p.runs
                    # Substitui o texto preservando a formatação do primeiro run
                    para i em intervalo(len(inline)):
                        para chave, valor em contexto.items():
                            se a chave estiver em inline[i].text:
                                texto = inline[i].text.replace(chave, str(valor))
                                inline[i].text = texto
    para p em doc.paragraphs:
        # Mesma lógica para parágrafos fora de tabelas
        inline = p.runs
        para i em intervalo(len(inline)):
            para chave, valor em contexto.items():
                se a chave estiver em inline[i].text:
                    texto = inline[i].text.replace(chave, str(valor))
                    inline[i].text = texto


def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, riscos_manuais, modelo_doc_carregado):
    doc = Documento(modelo_doc_carregado)
    riscos_info = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]
    riscos_por_categoria = {cat: [] para gato em CATEGORIAS_RISCO.keys()}
    danos_por_categoria = {cat: [] para gato em CATEGORIAS_RISCO.keys()}
    para _, risco_row em riscos_info.iterrows():
        categoria = str(risco_row.get("categoria", "")).lower()
        if categoria em riscos_por_categoria:
            riscos_por_categoria[categoria].append(str(risco_row.get("risco", "")))
            danos = risco_row.get("possiveis_danos")
            if pd.notna(danos): danos_por_categoria[categoria].append(str(danos))
    se riscos_manuais:
        map_categorias_rev = {v: k para k, v em CATEGORIAS_RISCO.items()}
        para risco_manual em riscos_manuais:
            categoria_display = risco_manual.get('categoria')
            categoria_alvo = map_categorias_rev.get(categoria_display)
            se categoria_alvo:
                riscos_por_categoria[categoria_alvo].append(risco_manual.get('risk_name', ''))
                se risco_manual.get('possíveis_danos'):
                    danos_por_categoria[categoria_alvo].append(risco_manual.get('possible_damages'))
    para gato em danos_por_categoria:
        danos_por_categoria[cat] = classificado(lista(set(danos_por_categoria[cat])))
    medicoes_ordenadas = sorted(medicoes_manuais, key=lambda med: med.get('agent', ''))
    medicoes_formatadas = []
    max_len = 0
    se medicoes_ordenadas:
        max_len = max(len(med.get('agent', '')) para med em medicoes_ordenadas)
    para med em medicoes_ordenadas:
        agente = med.get('agente', 'N/D')
        valor = med.get('valor', 'N/D')
        unidade = med.get('unidade', '')
        epi = med.get('epi', '')
        padding = ' ' * (max_len - len(agente))
        epi_info = f" | EPI: {epi}" if epi e epi.strip() else ""
        medicoes_formatadas.append(f"{agente}:{padding}\t{valor} {unidade}{epi_info}")
    medicoes_texto = "\n".join(medicoes_formatadas) if medicoes_formatadas else "Não é aplicável"
    data_admissao = "Não informado"
    if 'data_de_admissao' em funcionario e pd.notna(funcionario['data_de_admissao']):
        tente: data_admissao = pd.to_datetime(funcionario['data_de_admissao']).strftime('%d/%m/%Y')
        exceto Exceção: data_admissao = str(funcionario['data_de_admissao'])
    descricao_atividades = "Não informado"
    if 'descricao_de_atividades' em funcionario e pd.notna(funcionario['descricao_de_atividades']):
        descrição_atividades = str(funcionario['descricao_de_atividades'])
    def tratar_lista_vazia(lista, separador=", "):
        se não for lista ou all(não item.strip() para item na lista): retorne "Não identificado"
        retornar separador.join(classificado(lista(definir(item para item na lista se item e item.strip()))))
    contexto = {
        "[NOME EMPRESA]": str(funcionario.get("empresa", "N/A")),
        "[UNIDADE]": str(funcionario.get("unidade", "N/A")),
        "[NOME FUNCIONÁRIO]": str(funcionario.get("nome_do_funcionario", "N/A")),
        "[DADOS DE ADMISSÃO]": data_admissao,
        "[SETOR]": str(funcionario.get("setor", "N/A")),
        "[FUNÇÃO]": str(funcionario.get("funcao", "N/A")),
        "[DESCRIÇÃO DE ATIVIDADES]": descricao_atividades,
        "[RISCOS FÍSICOS]": tratar_lista_vazia(riscos_por_categoria["físico"]),
        "[RISCOS DE ACIDENTE]": tratar_lista_vazia(riscos_por_categoria["acidente"]),
        "[RISCOS QUÍMICOS]": tratar_lista_vazia(riscos_por_categoria["quimico"]),
        "[RISCOS BIOLÓGICOS]": tratar_lista_vazia(riscos_por_categoria["biológico"]),
        "[RISCOS ERGONÔMICOS]": tratar_lista_vazia(riscos_por_categoria["ergonômico"]),
        "[POSSÃ VEIS DANOS RISCOS FÁ SICOS]": tratar_lista_vazia(danos_por_categoria["físico"], "; "),
        "[POSSÃ VEIS DANOS RISCOS ACIDENTE]": tratar_lista_vazia(danos_por_categoria["acidente"], "; "),
        "[POSSÃ VEIS DANOS RISCOS QUÃ MICOS]": tratar_lista_vazia(danos_por_categoria["quimico"], "; "),
        "[POSSÃ VEIS DANOS RISCOS BIOLÓGICOS]": tratar_lista_vazia(danos_por_categoria["biologico"], "; "),
        "[POSSÃ VEIS DANOS RISCOS ERGONÔMICOS]": tratar_lista_vazia(danos_por_categoria["ergonômico"], "; "),
        "[EPIS]": tratar_lista_vazia([epi['epi_name'] para epi em epis_manuais]),
        "[MEDIÃ‡Ã•ES]": medicoes_texto,
    }
    substituir_placeholders(doc, contexto)
    documento de retorno

# --- PRINCIPAL DA APLICAÇÃO ---
def principal():
    verificar_autenticação()
    estado_da_sessão_do_usuário_inicial()
    
    se não st.session_state.get('autenticado'):
        mostrar_página_de_login()
        retornar
    
    user_id = st.session_state.user_data['user_id']
    mostrar_informações_do_usuário()
    
    st.markdown("""<div class="main-header"><h1>ðŸ“„ Gerador de Ordens de Serviço (OS)</h1><p>Gere OS em lote a partir de um modelo Word (.docx) e uma planilha de funções.</p></div>""", unsafe_allow_html=True)

    com st.container(border=True):
        st.markdown("##### ðŸ“‚ 1. Carregar os Documentos")
        col1, col2 = st.columns(2)
        com col1:
            arquivo_funcionarios = st.file_uploader("ðŸ“„ **Planilha de Funções (.xlsx)**", type="xlsx")
        com col2:
            arquivo_modelo_os = st.file_uploader("ðŸ“ **Modelo de SO (.docx)**", type="docx")

    se não for arquivo_funcionarios ou não arquivo_modelo_os:
        st.info("ðŸ“‹ Por favor, carregue a Planilha de Funções e o Modelo de OS para continuar.")
        retornar
    
    df_funcionarios_raw = carregar_planilha(arquivo_funcionarios)
    se df_funcionarios_raw for Nenhum:
        st.stop()

    df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_raw)
    df_pgr = obter_dados_pgr()

    com st.container(border=True):
        st.markdown('##### ðŸ'¥ 2. Selecione os Funcionários')
        setores = sorted(df_funcionarios['setor'].dropna().unique().tolist()) if 'setor' em df_funcionarios.columns else []
        setor_sel = st.multiselect("Filtrar por Setor(es)", setores)
        df_filtrado_setor = df_funcionarios[df_funcionarios['setor'].isin(setor_sel)] if setor_sel else df_funcionarios
        st.caption(f"{len(df_filtrado_setor)} função(s) no(s) setor(es) selecionado(s).")
        funcoes_disponiveis = sorted(df_filtrado_setor['funcao'].dropna().unique().tolist()) if 'funcao' em df_filtrado_setor.columns else []
        funcoes_formatadas = []
        se setor_sel:
            para funcao em funcoes_disponiveis:
                concluído = all((s, funcao) em st.session_state.cargos_concluidos for s em setor_sel)
                se concluído:
                    funcoes_formatadas.append(f"{funcao} âœ…Concluído")
                outro:
                    funcoes_formatadas.append(função)
        outro:
            funcoes_formatadas = funcoes_disponiveis
        funcao_sel_formatada = st.multiselect("Filtrar por Função/Carga(s)", funcoes_formatadas)
        funcao_sel = [f.replace(" âœ…Concluído", "") for f in funcao_sel_formatada]
        df_final_filtrado = df_filtrado_setor[df_filtrado_setor['funcao'].isin(funcao_sel)] if funcao_sel else df_filtrado_setor
        st.success(f"**{len(df_final_filtrado)} função(s) selecionada(s) para gerar OS.**")
        st.dataframe(df_final_filtrado[['nome_do_funcionario', 'setor', 'funcao']])

    com st.container(border=True):
        st.markdown('##### âš ï¸ 3. Configurar os Riscos e Medidas de Controle')
        st.info("Os riscos configurados aqui serão aplicados a TODOS os recursos selecionados.")
        riscos_selecionados = []
        nomes_abas = list(CATEGORIAS_RISCO.values()) + ["âž• Manual"]
        guias = st.tabs(nomes_abas)
        para i, (categoria_key, categoria_nome) em enumerar(CATEGORIAS_RISCO.items()):
            com abas[i]:
                riscos_da_categoria = df_pgr[df_pgr['categoria'] == categoria_key]['risco'].tolist()
                selecionados = st.multiselect("Seleção de riscos:", options=riscos_da_categoria, key=f"riscos_{categoria_key}")
                riscos_selecionados.extend(selecionados)
        com abas[-1]:
            com st.form("form_risco_manual", clear_on_submit=True):
                st.markdown("###### Adicionar um Risco que não está na lista")
                risco_manual_nome = st.text_input("Descrição do Risco")
                categoria_manual = st.selectbox("Categoria do Manual de Risco", list(CATEGORIAS_RISCO.values()))
                danos_manuais = st.text_area("Possíveis Danos (Opcional)")
                if st.form_submit_button("Adicionar Manual de Risco"):
                    se risco_manual_nome e categoria_manual:
                        user_data_manager.add_manual_risk(user_id, categoria_manual, risco_manual_nome, danos_manuais)
                        st.session_state.user_data_loaded = Falso
                        st.rerun()
            if st.session_state.riscos_manuais_adicionados:
                st.write("**Riscos manuais salvos:**")
                para r em st.session_state.riscos_manuais_adicionados:
                    col1, col2 = st.columns([4, 1])
                    col1.markdown(f"- **{r['nome_do_risco']}** ({r['categoria']})")
                    se col2.button("Removedor", chave=f"rem_risco_{r['id']}"):
                        gerenciador_de_dados_do_usuário.remove_manual_risk(id_do_usuário, r['id'])
                        st.session_state.user_data_loaded = Falso
                        st.rerun()
        
        total_riscos = len(riscos_selecionados) + len(st.session_state.riscos_manuais_adicionados)
        se total_riscos > 0:
            with st.expander(f"ðŸ“– Resumo de Riscos Selecionados ({total_riscos} no total)", expandido=True):
                riscos_para_exibir = {cat: [] para gato em CATEGORIAS_RISCO.values()}
                para risco_nome em riscos_selecionados:
                    categoria_key_series = df_pgr[df_pgr['risco'] == risco_nome]['categoria']
                    se não categoria_key_series.empty:
                        categoria_key = categoria_key_series.iloc[0]
                        categoria_display = CATEGORIAS_RISCO.get(categoria_key)
                        se categoria_display:
                            riscos_para_exibir[categoria_display].append(risco_nome)
                para risco_manual em st.session_state.riscos_manuais_adicionados:
                    riscos_para_exibir[risco_manual['category']].append(risco_manual['risk_name'])
                para categoria, lista_riscos em riscos_para_exibir.items():
                    se lista_riscos:
                        st.markdown(f"**{categoria}**")
                        para risco em sorted(list(set(lista_riscos))):
                            st.markdown(f"- {risco}")
        
        st.divisor()

        col_exp1, col_exp2 = st.columns(2)
        com col_exp1:
            with st.expander("ðŸ“Š **Adicionar Medições**"):
                com st.form("form_medicao", clear_on_submit=True):
                    opcoes_agente = ["-- Digite um novo agente abaixo --"] + AGENTES_DE_RISCO
                    agente_selecionado = st.selectbox("Selecione um Agente/Fonte da lista...", options=opcoes_agente)
                    agente_manual = st.text_input("...ou digite um novo aqui:")
                    valor = st.text_input("Valor Medido")
                    unidade = st.selectbox("Unidade", UNIDADES_DE_MEDIDA)
                    epi_med = st.text_input("EPI Associado (Opcional)")
                    if st.form_submit_button("Adicionar Medição"):
                        agente_a_salvar = agente_manual.strip() if agente_manual.strip() else agente_selecionado
                        if agente_a_salvar != "-- Digite um novo agente abaixo --" e valor:
                            user_data_manager.add_measurement(user_id, agente_a_salvar, valor, unidade, epi_med)
                            st.session_state.user_data_loaded = Falso
                            st.rerun()
                        outro:
                            st.warning("Por favor, preencha o Agente e o Valor.")
                if st.session_state.medicoes_adicionadas:
                    st.write("**Médias salvas:**")
                    para med em st.session_state.medicoes_adicionadas:
                        col1, col2 = st.columns([4, 1])
                        col1.markdown(f"- {med['agente']}: {med['valor']} {med['unidade']}")
                        se col2.button("Removedor", chave=f"rem_med_{med['id']}"):
                            user_data_manager.remove_measurement(id_do_usuário, med['id'])
                            st.session_state.user_data_loaded = Falso
                            st.rerun()
        com col_exp2:
            with st.expander("ðŸ¦º **Adicionar EPIs Gerais**"):
                com st.form("form_epi", clear_on_submit=True):
                    epi_nome = st.text_input("Nome do EPI")
                    se st.form_submit_button("Adicionar EPI"):
                        se epi_nome:
                            user_data_manager.add_epi(id_do_usuário, nome_do_episódio)
                            st.session_state.user_data_loaded = Falso
                            st.rerun()
                se st.session_state.epis_adicionados:
                    st.write("**EPIs salvos:**")
                    para epi em st.session_state.epis_adicionados:
                        col1, col2 = st.columns([4, 1])
                        col1.markdown(f"- {epi['nome_epi']}")
                        se col2.button("Removedor", chave=f"rem_epi_{epi['id']}"):
                            user_data_manager.remove_epi(id_do_usuário, epi['id'])
                            st.session_state.user_data_loaded = Falso
                            st.rerun()

    st.divisor()
    if st.button("ðŸš€ Gerar OS para Funções Selecionadas", type="primary", use_container_width=True, disabled=df_final_filtrado.empty):
        with st.spinner(f"Gerando {len(df_final_filtrado)} documentos..."):
            documentos_gerados = []
            combinacoes_processadas = set()
            para _, função em df_final_filtrado.iterrows():
                combinacoes_processadas.add((func['setor'], func['funcao']))
                doc = gerar_os(
                    função,
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
                caminho_no_zip = f"{func.get('setor', 'SemSetor')}/{func.get('funcao', 'SemFuncao')}/OS_{nome_limpo}.docx"
                documentos_gerados.append((caminho_no_zip, doc_io.getvalue()))
            st.session_state.cargos_concluidos.update(combinacoes_processadas)
            se documentos_gerados:
                zip_buffer = BytesIO()
                com zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) como zip_file:
                    para nome_arquivo, conteudo_doc em documentos_gerados:
                        zip_file.writestr(nome_arquivo, conteudo_doc)
                nome_arquivo_zip = f"OS_Geradas_{time.strftime('%Y%m%d')}.zip"
                st.success(f"ðŸŽ‰ **{len(documentos_gerados)} Ordens de Serviço geradas!**")
                st.botão_de_download(
                    label="ðŸ“¥ Baixar Tudo como OS (.zip)",
                    dados=zip_buffer.getvalue(),
                    nome_do_arquivo=nome_arquivo_zip,
                    mime="aplicativo/zip",
                    use_container_width=Verdadeiro
                )

se __nome__ == "__principal__":
    principal()
