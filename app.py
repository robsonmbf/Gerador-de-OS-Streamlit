import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import datetime

# Dados simplificados dos Riscos PGR (principais de cada categoria)
RISCOS_PGR_DADOS = {
    'Físico': {
        'riscos': [
            'Exposição ao Ruído',
            'Vibrações Localizadas (mão/braço)', 
            'Vibração de Corpo Inteiro (AREN)',
            'Vibração de Corpo Inteiro (VDVR)',
            'Exposição à Temperatura Ambiente Elevada',
            'Exposição à Temperatura Ambiente Baixa',
            'Exposição à Radiações Ionizantes',
            'Exposição à Radiações Não-ionizantes',
            'Pressão Atmosférica Anormal',
            'Umidade',
            'Ambiente Artificialmente Frio'
        ],
        'danos': [
            'Perda Auditiva Induzida pelo Ruído Ocupacional (PAIRO)',
            'Alterações articulares e vasomotoras',
            'Alterações no sistema digestivo, musculoesquelético, nervoso, visão, enjoos, náuseas, palidez',
            'Alterações no sistema digestivo, musculoesquelético, nervoso, visão, enjoos, náuseas, palidez',
            'Desidratação, erupções cutâneas, cãibras, fadiga física, problemas cardiocirculatórios',
            'Estresse, desconforto, dormência, rigidez, redução da destreza, formigamento',
            'Dano às células do corpo humano, causando doenças graves, inclusive câncer',
            'Depressão imunológica, fotoenvelhecimento, lesões oculares, doenças graves',
            'Barotrauma pulmonar, lesão de tecido pulmonar, embolia arterial gasosa',
            'Doenças do aparelho respiratório, quedas, doenças de pele, doenças circulatórias',
            'Estresse, desconforto, dormência, rigidez, redução da destreza'
        ]
    },
    'Químico': {
        'riscos': ['Exposição a Produto Químico'],
        'danos': ['Irritação/lesão ocular, na pele e mucosas; Dermatites; Queimadura Química; Intoxicação; Náuseas; Vômitos']
    },
    'Biológico': {
        'riscos': [
            'Água e/ou alimentos contaminados',
            'Contato com Fluido Orgânico',
            'Contato com Pessoas Doentes e/ou Material Infectocontagiante',
            'Contaminação pelo Corona Vírus',
            'Exposição à Agentes Microbiológicos'
        ],
        'danos': [
            'Intoxicação, diarreias, infecções intestinais',
            'Doenças infectocontagiosas',
            'Doenças infectocontagiosas',
            'COVID-19, gripes, febre, tosse seca, cansaço, dores, dor de garganta, diarreia',
            'Doenças infectocontagiosas, dermatites, irritação, desconforto, infecção respiratória'
        ]
    },
    'Ergonômico': {
        'riscos': [
            'Posturas incômodas por longos períodos',
            'Postura sentada por longos períodos',
            'Postura em pé por longos períodos',
            'Esforço físico intenso',
            'Levantamento e transporte manual de cargas',
            'Movimentos repetitivos',
            'Trabalho com necessidade de ritmos intensos',
            'Trabalho noturno',
            'Iluminação inadequada',
            'Temperatura fora dos parâmetros de conforto'
        ],
        'danos': [
            'Distúrbios musculoesqueléticos em músculos e articulações',
            'Sobrecarga dos membros superiores e coluna vertebral, dor localizada',
            'Sobrecarga corporal, dores nos membros inferiores e coluna, cansaço físico',
            'Distúrbios musculoesqueléticos, fadiga, dor localizada, redução da produtividade',
            'Distúrbios musculoesqueléticos, fadiga, dor localizada, redução da produtividade',
            'Distúrbios osteomusculares em músculos e articulações dos membros utilizados',
            'Sobrecarga e fadiga física e cognitiva, redução da percepção de risco',
            'Alterações psicofisiológicas e/ou sociais',
            'Fadiga visual e cognitiva, desconforto, redução da percepção de riscos',
            'Irritabilidade, estresse, dores de cabeça, perda de foco no trabalho'
        ]
    },
    'Acidente': {
        'riscos': [
            'Exposição à Energia Elétrica',
            'Queda de pessoa com diferença de nível',
            'Queda de pessoa em mesmo nível',
            'Objetos cortantes/perfurocortantes',
            'Incêndio/Explosão',
            'Máquinas e equipamentos sem proteção',
            'Trabalho em altura',
            'Trabalho em espaços confinados',
            'Atropelamento',
            'Impacto de objeto que cai'
        ],
        'danos': [
            'Choque elétrico e eletroplessão (eletrocussão)',
            'Escoriações, ferimentos, cortes, luxações, fraturas, morte',
            'Escoriações, ferimentos, cortes, luxações, fraturas',
            'Corte, laceração, ferida contusa, punctura (ferida aberta), perfuração',
            'Queimadura de 1º, 2º ou 3º grau, asfixia, arremessos, cortes, escoriações',
            'Prensamento, aprisionamento, cortes, escoriações, luxações, fraturas, amputações',
            'Escoriações, ferimentos, cortes, luxações, fraturas, morte',
            'Asfixia, hiperóxia, contaminação, queimadura, arremessos, cortes, escoriações',
            'Compressão/esmagamento, cortes, escoriações, luxações, fraturas, amputações',
            'Esmagamento, prensamento, aprisionamento, cortes, escoriações, luxações, fraturas'
        ]
    }
}

def get_danos_por_riscos_pgr(categoria, riscos_selecionados):
    """Retorna os danos associados aos riscos selecionados da planilha PGR"""
    if categoria not in RISCOS_PGR_DADOS or not riscos_selecionados:
        return ""

    danos_lista = []
    riscos_categoria = RISCOS_PGR_DADOS[categoria]["riscos"]
    danos_categoria = RISCOS_PGR_DADOS[categoria]["danos"]

    for risco in riscos_selecionados:
        if risco in riscos_categoria:
            indice = riscos_categoria.index(risco)
            if indice < len(danos_categoria):
                danos_lista.append(danos_categoria[indice])

    return "; ".join(danos_lista) if danos_lista else ""

# Configuração da página
st.set_page_config(
    page_title="📋 Gerador de Ordem de Serviço - NR01",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado
st.markdown("""
<style>
    .stApp {
        background-color: #1a1a1a;
        color: #ffffff;
    }
    .stSelectbox > div > div {
        background-color: #2d2d2d;
        color: #ffffff;
    }
    .stTextArea > div > div > textarea {
        background-color: #2d2d2d;
        color: #ffffff;
    }
    .stButton > button {
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
    }
    .stButton > button:hover {
        background-color: #45a049;
    }
    .stMultiSelect > div > div {
        background-color: #2d2d2d;
        color: #ffffff;
    }
</style>
""", unsafe_allow_html=True)

def create_sample_data():
    """Cria dados de exemplo para demonstração"""
    sample_data = {
        'Nome': ['JOÃO SILVA SANTOS', 'MARIA OLIVEIRA COSTA', 'PEDRO ALVES FERREIRA'],
        'Setor': ['PRODUCAO DE LA DE ACO', 'ADMINISTRACAO DE RH', 'MANUTENCAO QUIMICA'],
        'Função': ['OPERADOR PRODUCAO I', 'ANALISTA ADM PESSOAL PL', 'MECANICO MANUT II'],
        'Data de Admissão': ['15/03/2020', '22/08/2019', '10/01/2021'],
        'Empresa': ['SUA EMPRESA', 'SUA EMPRESA', 'SUA EMPRESA'],
        'Unidade': ['Matriz', 'Matriz', 'Matriz'],
        'Descrição de Atividades': [
            'Operar equipamentos de produção nível I, controlar parâmetros operacionais, realizar inspeções visuais e registrar dados de produção.',
            'Executar atividades de administração de pessoal, controlar documentos trabalhistas, elaborar relatórios e dar suporte às equipes.',
            'Executar manutenção preventiva e corretiva em equipamentos, diagnosticar falhas, trocar componentes e registrar intervenções.'
        ]
    }
    return pd.DataFrame(sample_data)

def validate_excel_structure(df):
    """Valida se a planilha tem a estrutura necessária"""
    required_columns = ['Nome', 'Setor', 'Função', 'Data de Admissão', 'Empresa', 'Unidade', 'Descrição de Atividades']
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        return False, f"Colunas obrigatórias faltando: {', '.join(missing_columns)}"

    if df.empty:
        return False, "A planilha está vazia"

    return True, "Estrutura válida"

def create_os_document(employee_data, riscos_data, medidas_data, assessment_data, measurements_data):
    """Cria o documento da Ordem de Serviço"""
    doc = Document()

    # Título principal
    heading = doc.add_heading('ORDEM DE SERVIÇO - SEGURANÇA E MEDICINA DO TRABALHO', level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Informações do funcionário
    doc.add_heading('DADOS DO FUNCIONÁRIO', level=2)

    info_table = doc.add_table(rows=6, cols=2)
    info_table.style = 'Table Grid'

    # Preencher tabela de informações
    info_data = [
        ('Nome:', employee_data.get('Nome', 'N/A')),
        ('Setor:', employee_data.get('Setor', 'N/A')),
        ('Função:', employee_data.get('Função', 'N/A')),
        ('Data de Admissão:', employee_data.get('Data de Admissão', 'N/A')),
        ('Empresa:', employee_data.get('Empresa', 'N/A')),
        ('Unidade:', employee_data.get('Unidade', 'N/A'))
    ]

    for i, (label, value) in enumerate(info_data):
        info_table.cell(i, 0).text = label
        info_table.cell(i, 1).text = str(value)
        info_table.cell(i, 0).paragraphs[0].runs[0].bold = True

    # Descrição das atividades
    doc.add_heading('DESCRIÇÃO DAS ATIVIDADES', level=2)
    desc_para = doc.add_paragraph(employee_data.get('Descrição de Atividades', 'Não informado'))

    # Agentes de riscos ocupacionais
    heading = doc.add_heading('AGENTES DE RISCOS OCUPACIONAIS - NR01 item 1.4.1 b) I / item 1.4.4 a)', level=2)

    # Usar os danos da planilha PGR quando disponíveis
    for tipo_risco, riscos_texto in riscos_data.items():
        if riscos_texto and riscos_texto.strip():
            # Separar os riscos individuais
            riscos_lista = [r.strip() for r in riscos_texto.split(';') if r.strip()]

            # Buscar danos específicos da planilha PGR
            danos_pgr = get_danos_por_riscos_pgr(tipo_risco, riscos_lista)

            para = doc.add_paragraph()
            para.add_run(f"➤ {tipo_risco}: ").bold = True

            if danos_pgr:
                para.add_run(danos_pgr)
            else:
                # Fallback para danos genéricos se não encontrar na planilha PGR
                danos_genericos = {
                    'Físico': 'Perda auditiva, fadiga, stress térmico',
                    'Químico': 'Intoxicação, irritação, dermatoses', 
                    'Acidente': 'Ferimentos, fraturas, queimaduras',
                    'Ergonômico': 'Dores musculares, LER/DORT, fadiga',
                    'Biológico': 'Infecções, alergias, doenças'
                }
                para.add_run(danos_genericos.get(tipo_risco, 'Possíveis danos relacionados a este tipo de risco'))

    # Medidas de controle
    heading = doc.add_heading('MEIOS PARA PREVENIR E CONTROLAR RISCOS - NR01 item 1.4.4 b)', level=2)

    for categoria, medidas in medidas_data.items():
        if medidas.strip():
            para = doc.add_paragraph()
            para.add_run(f"➤ {categoria}: ").bold = True
            para.add_run(medidas)

    # Data e assinatura
    doc.add_paragraph('\n')
    doc.add_paragraph(f'Data: {datetime.datetime.now().strftime("%d/%m/%Y")}')
    doc.add_paragraph('\n\n')
    doc.add_paragraph('_' * 50)
    doc.add_paragraph('Assinatura do Responsável pela Segurança')

    return doc

def main():
    st.title("📋 Gerador de Ordem de Serviço - NR01")
    st.markdown("Sistema para geração de Ordens de Serviço de Segurança e Medicina do Trabalho")

    # Sidebar para upload ou dados de exemplo
    st.sidebar.header("📊 Dados dos Funcionários")

    option = st.sidebar.radio(
        "Escolha a fonte dos dados:",
        ["📤 Upload de planilha Excel", "🧪 Usar dados de exemplo"]
    )

    df = None

    if option == "📤 Upload de planilha Excel":
        uploaded_file = st.sidebar.file_uploader(
            "Carregue a planilha com os dados dos funcionários",
            type=['xlsx', 'xls'],
            help="A planilha deve conter as colunas: Nome, Setor, Função, Data de Admissão, Empresa, Unidade, Descrição de Atividades"
        )

        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file)
                is_valid, message = validate_excel_structure(df)

                if is_valid:
                    st.sidebar.success(f"✅ Planilha carregada: {len(df)} funcionários")
                else:
                    st.sidebar.error(f"❌ {message}")
                    df = None
            except Exception as e:
                st.sidebar.error(f"❌ Erro ao carregar planilha: {str(e)}")

    else:  # Usar dados de exemplo
        df = create_sample_data()
        st.sidebar.success("✅ Usando dados de exemplo")

    if df is not None:
        # Seleção de funcionários
        st.subheader("👥 Seleção de Funcionários")

        selected_employees = st.multiselect(
            "Selecione os funcionários para gerar as OS:",
            options=df['Nome'].tolist(),
            default=df['Nome'].tolist()[0:1] if len(df) > 0 else [],
            help="Você pode selecionar múltiplos funcionários"
        )

        if selected_employees:
            st.success(f"✅ {len(selected_employees)} funcionário(s) selecionado(s)")

            # Interface de Riscos baseada na planilha PGR
            st.subheader("📋 Seleção de Riscos Ocupacionais (PGR)")
            st.info("Selecione os riscos aplicáveis baseados na planilha PGR:")

            # Criar tabs para cada categoria
            tab_fisico, tab_quimico, tab_biologico, tab_ergonomico, tab_acidente = st.tabs([
                "🔥 Físicos", "⚗️ Químicos", "🦠 Biológicos", "🏃 Ergonômicos", "⚠️ Acidentes"
            ])

            riscos_selecionados_pgr = {}

            # Tab Físicos
            with tab_fisico:
                st.write(f"**Riscos Físicos Disponíveis:** {len(RISCOS_PGR_DADOS['Físico']['riscos'])} opções")
                riscos_selecionados_pgr['Físico'] = st.multiselect(
                    "Selecione os Riscos Físicos:",
                    options=RISCOS_PGR_DADOS['Físico']['riscos'],
                    key="riscos_pgr_fisico",
                    help="Riscos físicos da planilha PGR"
                )
                if riscos_selecionados_pgr['Físico']:
                    danos = get_danos_por_riscos_pgr('Físico', riscos_selecionados_pgr['Físico'])
                    if danos:
                        st.info(f"**Possíveis Danos:** {danos}")

            # Tab Químicos  
            with tab_quimico:
                st.write(f"**Riscos Químicos Disponíveis:** {len(RISCOS_PGR_DADOS['Químico']['riscos'])} opções")
                riscos_selecionados_pgr['Químico'] = st.multiselect(
                    "Selecione os Riscos Químicos:",
                    options=RISCOS_PGR_DADOS['Químico']['riscos'],
                    key="riscos_pgr_quimico",
                    help="Riscos químicos da planilha PGR"
                )
                if riscos_selecionados_pgr['Químico']:
                    danos = get_danos_por_riscos_pgr('Químico', riscos_selecionados_pgr['Químico'])
                    if danos:
                        st.info(f"**Possíveis Danos:** {danos}")

            # Tab Biológicos
            with tab_biologico:
                st.write(f"**Riscos Biológicos Disponíveis:** {len(RISCOS_PGR_DADOS['Biológico']['riscos'])} opções")
                riscos_selecionados_pgr['Biológico'] = st.multiselect(
                    "Selecione os Riscos Biológicos:",
                    options=RISCOS_PGR_DADOS['Biológico']['riscos'],
                    key="riscos_pgr_biologico",
                    help="Riscos biológicos da planilha PGR"
                )
                if riscos_selecionados_pgr['Biológico']:
                    danos = get_danos_por_riscos_pgr('Biológico', riscos_selecionados_pgr['Biológico'])
                    if danos:
                        st.info(f"**Possíveis Danos:** {danos}")

            # Tab Ergonômicos
            with tab_ergonomico:
                st.write(f"**Riscos Ergonômicos Disponíveis:** {len(RISCOS_PGR_DADOS['Ergonômico']['riscos'])} opções")
                riscos_selecionados_pgr['Ergonômico'] = st.multiselect(
                    "Selecione os Riscos Ergonômicos:",
                    options=RISCOS_PGR_DADOS['Ergonômico']['riscos'],
                    key="riscos_pgr_ergonomico",
                    help="Riscos ergonômicos da planilha PGR"
                )
                if riscos_selecionados_pgr['Ergonômico']:
                    danos = get_danos_por_riscos_pgr('Ergonômico', riscos_selecionados_pgr['Ergonômico'])
                    if danos:
                        st.info(f"**Possíveis Danos:** {danos}")

            # Tab Acidentes
            with tab_acidente:
                st.write(f"**Riscos de Acidente Disponíveis:** {len(RISCOS_PGR_DADOS['Acidente']['riscos'])} opções")
                riscos_selecionados_pgr['Acidente'] = st.multiselect(
                    "Selecione os Riscos de Acidente:",
                    options=RISCOS_PGR_DADOS['Acidente']['riscos'],
                    key="riscos_pgr_acidente",
                    help="Riscos de acidente da planilha PGR"
                )
                if riscos_selecionados_pgr['Acidente']:
                    danos = get_danos_por_riscos_pgr('Acidente', riscos_selecionados_pgr['Acidente'])
                    if danos:
                        st.info(f"**Possíveis Danos:** {danos}")

            # Converter para o formato esperado pelo resto do código
            riscos = {}
            for categoria, lista_riscos in riscos_selecionados_pgr.items():
                if lista_riscos:
                    riscos[categoria] = "; ".join(lista_riscos)
                else:
                    riscos[categoria] = ""

            # Medidas de controle
            st.subheader("🛡️ Medidas de Controle")

            col3, col4 = st.columns([1, 1])

            with col3:
                medidas = {}
                medidas['EPC (Proteção Coletiva)'] = st.text_area(
                    "🏭 Equipamentos de Proteção Coletiva:",
                    placeholder="Ex: Ventilação, isolamento acústico, proteções em máquinas",
                    height=80
                )
                medidas['Administrativas'] = st.text_area(
                    "📋 Medidas Administrativas:",
                    placeholder="Ex: Treinamentos, procedimentos, pausas, rotação de tarefas",
                    height=80
                )

            with col4:
                medidas['EPI (Proteção Individual)'] = st.text_area(
                    "🦺 Equipamentos de Proteção Individual:",
                    placeholder="Ex: Protetor auricular, luvas, capacete, óculos de proteção",
                    height=80
                )
                medidas['Monitoramento'] = st.text_area(
                    "📊 Monitoramento:",
                    placeholder="Ex: Avaliações ambientais, exames médicos, inspeções",
                    height=80
                )

            # Botão para gerar documentos
            st.subheader("📄 Geração de Documentos")

            col5, col6 = st.columns([1, 1])

            with col5:
                if st.button("📋 Gerar Ordens de Serviço", type="primary"):
                    if any(riscos.values()) or any(medidas.values()):
                        documents = []

                        for employee_name in selected_employees:
                            employee_data = df[df['Nome'] == employee_name].iloc[0].to_dict()

                            doc = create_os_document(
                                employee_data,
                                riscos,
                                medidas,
                                {},  # assessment_data - não usado nesta versão simplificada
                                {}   # measurements_data - não usado nesta versão simplificada
                            )

                            # Salvar documento em memória
                            doc_buffer = BytesIO()
                            doc.save(doc_buffer)
                            doc_buffer.seek(0)

                            documents.append({
                                'name': f"OS_{employee_name.replace(' ', '_')}.docx",
                                'data': doc_buffer.getvalue()
                            })

                        # Download individual ou em lote
                        if len(documents) == 1:
                            st.download_button(
                                label="📥 Baixar Ordem de Serviço",
                                data=documents[0]['data'],
                                file_name=documents[0]['name'],
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        else:
                            # Criar ZIP com múltiplos documentos
                            import zipfile
                            zip_buffer = BytesIO()

                            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                for doc_info in documents:
                                    zip_file.writestr(doc_info['name'], doc_info['data'])

                            zip_buffer.seek(0)

                            st.download_button(
                                label=f"📥 Baixar {len(documents)} Ordens de Serviço (ZIP)",
                                data=zip_buffer.getvalue(),
                                file_name=f"OS_Lote_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                                mime="application/zip"
                            )

                        st.success(f"✅ {len(documents)} documento(s) gerado(s) com sucesso!")

                    else:
                        st.error("❌ É necessário preencher pelo menos um risco ou medida de controle.")

            with col6:
                st.info("💡 **Dica:** Selecione os riscos e medidas aplicáveis, depois clique em gerar para criar as Ordens de Serviço.")

    else:
        st.info("👆 Carregue uma planilha ou use os dados de exemplo para começar.")

if __name__ == "__main__":
    main()
