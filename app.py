import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import datetime

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
    .main-title {
        text-align: center;
        color: #1e3a8a;
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 30px;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }
    .section-header {
        color: #1e40af;
        font-size: 1.3rem;
        font-weight: bold;
        margin-top: 20px;
        margin-bottom: 10px;
        border-bottom: 2px solid #3b82f6;
        padding-bottom: 5px;
    }
    .info-box {
        background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
        padding: 20px;
        border-radius: 10px;
        border-left: 4px solid #3b82f6;
        margin: 15px 0;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    .success-box {
        background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%);
        padding: 20px;
        border-radius: 10px;
        border-left: 4px solid #22c55e;
        margin: 15px 0;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    .warning-box {
        background: linear-gradient(135deg, #fefce8 0%, #fef3c7 100%);
        padding: 20px;
        border-radius: 10px;
        border-left: 4px solid #eab308;
        margin: 15px 0;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    .upload-area {
        border: 2px dashed #3b82f6;
        border-radius: 10px;
        padding: 30px;
        text-align: center;
        background-color: #f8fafc;
        margin: 20px 0;
    }
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        text-align: center;
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

def create_os_document(funcionario_data, riscos_data, medidas_data, avaliacoes_data):
    """Cria documento da Ordem de Serviço baseado no modelo NR-01"""
    doc = Document()
    
    # Configurar margens
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
    
    # Título principal
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run('ORDEM DE SERVIÇO - NR01')
    title_run.bold = True
    title_run.font.size = Inches(0.2)
    
    # Subtítulo
    subtitle_para = doc.add_paragraph()
    subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle_para.add_run('Informações sobre Condições de Segurança e Saúde no Trabalho')
    subtitle_run.italic = True
    
    doc.add_paragraph()  # Espaço
    
    # Cabeçalho com informações do funcionário
    doc.add_paragraph(f"Empresa: {funcionario_data.get('Empresa', '')}\t\t\tUnidade: {funcionario_data.get('Unidade', '')}")
    doc.add_paragraph(f"Nome do Funcionário: {funcionario_data.get('Nome', '')}")
    doc.add_paragraph(f"Data de Admissão: {funcionario_data.get('Data de Admissão', '')}")
    doc.add_paragraph(f"Setor de Trabalho: {funcionario_data.get('Setor', '')}\t\t\tFunção: {funcionario_data.get('Função', '')}")
    
    doc.add_paragraph()  # Espaço
    
    # Tarefas da função
    heading = doc.add_heading('TAREFAS DA FUNÇÃO', level=2)
    heading.runs[0].font.size = Inches(0.15)
    doc.add_paragraph(funcionario_data.get('Descrição de Atividades', 'Atividades conforme descrição da função.'))
    
    # Agentes de riscos ocupacionais
    heading = doc.add_heading('AGENTES DE RISCOS OCUPACIONAIS - NR01 item 1.4.1 b) I / item 1.4.4 a)', level=2)
    heading.runs[0].font.size = Inches(0.12)
    
    for tipo_risco, riscos in riscos_data.items():
        if riscos.strip():
            para = doc.add_paragraph()
            para.add_run(f"➤ {tipo_risco}: ").bold = True
            para.add_run(riscos)
    
    # Possíveis danos à saúde
    heading = doc.add_heading('POSSÍVEIS DANOS À SAÚDE - NR01 item 1.4.1 b) I', level=2)
    heading.runs[0].font.size = Inches(0.12)
    
    danos_mapping = {
        'Físico': 'Perda auditiva, fadiga, stress térmico',
        'Químico': 'Intoxicação, irritação, dermatoses',
        'Acidente': 'Ferimentos, fraturas, queimaduras',
        'Ergonômico': 'Dores musculares, LER/DORT, fadiga',
        'Biológico': 'Infecções, alergias, doenças'
    }
    
    for tipo_risco, riscos in riscos_data.items():
        if riscos.strip():
            para = doc.add_paragraph()
            para.add_run(f"➤ {tipo_risco}: ").bold = True
            para.add_run(danos_mapping.get(tipo_risco, 'Possíveis danos relacionados a este tipo de risco'))
    
    # Meios de prevenção e controle
    heading = doc.add_heading('MEIOS PARA PREVENIR E CONTROLAR RISCOS - NR01 item 1.4.4 b)', level=2)
    heading.runs[0].font.size = Inches(0.12)
    doc.add_paragraph(medidas_data.get('prevencao', 'Seguir rigorosamente os procedimentos de segurança estabelecidos pela empresa.'))
    
    # Medidas adotadas pela empresa
    heading = doc.add_heading('MEDIDAS ADOTADAS PELA EMPRESA - NR01 item 1.4.1 b) II / item 1.4.4 c)', level=2)
    heading.runs[0].font.size = Inches(0.12)
    
    epi_para = doc.add_paragraph()
    epi_para.add_run("EPI Obrigatórios: ").bold = True
    epi_para.add_run(medidas_data.get('epi', 'Conforme necessidade da função e análise de riscos'))
    
    doc.add_paragraph(medidas_data.get('empresa', 
        'Treinamentos periódicos, supervisão constante, fornecimento e exigência de uso de EPIs, '
        'manutenção preventiva de equipamentos, monitoramento do ambiente de trabalho, '
        'exames médicos ocupacionais conforme PCMSO.'))
    
    # Avaliações ambientais
    heading = doc.add_heading('AVALIAÇÕES AMBIENTAIS - NR01 item 1.4.1 b) IV', level=2)
    heading.runs[0].font.size = Inches(0.12)
    doc.add_paragraph(avaliacoes_data.get('medicoes', 
        'As avaliações ambientais são realizadas conforme cronograma estabelecido pelo PPRA/PGR, '
        'com medições de agentes físicos, químicos e biológicos quando aplicável.'))
    
    # Procedimentos de emergência
    heading = doc.add_heading('PROCEDIMENTOS EM SITUAÇÕES DE EMERGÊNCIA - NR01 item 1.4.4 d) / item 1.4.1 e)', level=2)
    heading.runs[0].font.size = Inches(0.12)
    
    emergencia_text = """• Comunique imediatamente o acidente à chefia imediata ou pessoa designada;
• Preserve as condições do local de acidente até comunicação com autoridade competente;
• Em caso de ferimento, procure o ambulatório médico ou primeiros socorros;
• Siga as orientações do Plano de Emergência da empresa;
• Participe dos treinamentos de abandono e emergência."""
    
    doc.add_paragraph(emergencia_text)
    
    # Grave e iminente risco
    heading = doc.add_heading('ORIENTAÇÕES SOBRE GRAVE E IMINENTE RISCO - NR01 item 1.4.4 e) / item 1.4.3', level=2)
    heading.runs[0].font.size = Inches(0.12)
    
    gir_text = """• Sempre que constatar situação de Grave e Iminente Risco, interrompa imediatamente as atividades;
• Informe imediatamente ao seu superior hierárquico ou responsável pela área;
• Registre a ocorrência conforme procedimento estabelecido pela empresa;
• Aguarde as providências e liberação formal para retorno às atividades;
• Todo empregado tem o direito de recusar trabalho em condições de risco grave e iminente."""
    
    doc.add_paragraph(gir_text)
    
    doc.add_paragraph()  # Espaço
    
    # Nota legal
    nota_para = doc.add_paragraph()
    nota_run = nota_para.add_run(
        "IMPORTANTE: Conforme Art. 158 da CLT e NR-01 item 1.4.2.1, o descumprimento das "
        "disposições legais sobre segurança e saúde no trabalho sujeita o empregado às "
        "penalidades legais, inclusive demissão por justa causa."
    )
    nota_run.bold = True
    
    doc.add_paragraph()  # Espaço
    
    # Assinaturas
    doc.add_paragraph("_" * 40 + "\t\t" + "_" * 40)
    doc.add_paragraph("Funcionário\t\t\t\t\tResponsável pela Área")
    doc.add_paragraph(f"Data: {datetime.date.today().strftime('%d/%m/%Y')}")
    
    return doc

def main():
    # Título principal
    st.markdown('<h1 class="main-title">📋 Gerador de Ordem de Serviço - NR01</h1>', 
                unsafe_allow_html=True)
    
    # Sidebar com informações
    with st.sidebar:
        st.markdown("### 🛡️ Sistema de OS - NR01")
        st.markdown("---")
        
        st.markdown("### 📋 Como usar:")
        st.markdown("""
        1. **Upload** da planilha de funcionários
        2. **Selecione** o funcionário
        3. **Configure** os riscos e medidas
        4. **Gere** a Ordem de Serviço
        """)
        
        st.markdown("### 📁 Estrutura da Planilha:")
        st.markdown("""
        **Colunas obrigatórias:**
        - Nome
        - Setor
        - Função
        - Data de Admissão
        - Empresa
        - Unidade
        - Descrição de Atividades
        """)
        
        # Botão para baixar exemplo
        sample_df = create_sample_data()
        sample_buffer = BytesIO()
        sample_df.to_excel(sample_buffer, index=False)
        sample_buffer.seek(0)
        
        st.download_button(
            "📥 Baixar Planilha Exemplo",
            data=sample_buffer.getvalue(),
            file_name="modelo_funcionarios.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    # Área de upload
    st.markdown('<div class="upload-area">', unsafe_allow_html=True)
    st.markdown("### 📤 Upload da Planilha de Funcionários")
    
    uploaded_file = st.file_uploader(
        "Selecione sua planilha Excel (.xlsx)",
        type=['xlsx'],
        help="A planilha deve conter as colunas: Nome, Setor, Função, Data de Admissão, Empresa, Unidade, Descrição de Atividades"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_file is not None:
        try:
            # Carregar dados
            df = pd.read_excel(uploaded_file)
            
            # Validar estrutura
            is_valid, message = validate_excel_structure(df)
            
            if not is_valid:
                st.error(f"❌ {message}")
                st.info("💡 Use a planilha exemplo como modelo para estruturar seus dados corretamente.")
                return
            
            st.success(f"✅ Planilha carregada com sucesso! {len(df)} funcionários encontrados.")
            
            # Layout principal em colunas
            col1, col2 = st.columns([1, 1])
            
            with col1:
                st.markdown('<div class="section-header">👤 Seleção do Funcionário</div>', 
                            unsafe_allow_html=True)
                
                # Filtros
                setores = ['Todos'] + sorted(df['Setor'].dropna().unique().tolist())
                setor_selecionado = st.selectbox("🏢 Filtrar por Setor:", setores)
                
                if setor_selecionado != 'Todos':
                    df_filtrado = df[df['Setor'] == setor_selecionado]
                else:
                    df_filtrado = df
                
                # Seleção do funcionário
                funcionarios = [''] + df_filtrado['Nome'].dropna().unique().tolist()
                funcionario_selecionado = st.selectbox("👤 Selecionar Funcionário:", funcionarios)
                
                if funcionario_selecionado:
                    funcionario_info = df_filtrado[df_filtrado['Nome'] == funcionario_selecionado].iloc[0]
                    
                    # Exibir informações do funcionário
                    st.markdown('<div class="info-box">', unsafe_allow_html=True)
                    st.markdown(f"**👤 Nome:** {funcionario_info['Nome']}")
                    st.markdown(f"**🏢 Setor:** {funcionario_info['Setor']}")
                    st.markdown(f"**💼 Função:** {funcionario_info['Função']}")
                    st.markdown(f"**🏭 Empresa:** {funcionario_info['Empresa']}")
                    st.markdown(f"**📍 Unidade:** {funcionario_info['Unidade']}")
                    st.markdown(f"**📅 Admissão:** {funcionario_info['Data de Admissão']}")
                    st.markdown('</div>', unsafe_allow_html=True)
            
            with col2:
                st.markdown('<div class="section-header">⚠️ Agentes de Riscos Ocupacionais</div>', 
                            unsafe_allow_html=True)
                
                # Riscos ocupacionais com exemplos
                riscos = {}
                riscos['Físico'] = st.text_area(
                    "🔊 Riscos Físicos:", 
                    placeholder="Ex: Ruído acima de 85dB, vibração, temperaturas extremas, radiação",
                    height=60
                )
                riscos['Químico'] = st.text_area(
                    "🧪 Riscos Químicos:", 
                    placeholder="Ex: Solventes, ácidos, gases, vapores, poeiras químicas",
                    height=60
                )
                riscos['Acidente'] = st.text_area(
                    "⚠️ Riscos de Acidente:", 
                    placeholder="Ex: Máquinas sem proteção, eletricidade, trabalho em altura",
                    height=60
                )
                riscos['Ergonômico'] = st.text_area(
                    "🦴 Riscos Ergonômicos:", 
                    placeholder="Ex: Levantamento de peso, postura inadequada, movimentos repetitivos",
                    height=60
                )
                riscos['Biológico'] = st.text_area(
                    "🦠 Riscos Biológicos:", 
                    placeholder="Ex: Vírus, bactérias, fungos, parasitas",
                    height=60
                )
            
            # Segunda linha de colunas
            col3, col4 = st.columns([1, 1])
            
            with col3:
                st.markdown('<div class="section-header">🛡️ Medidas de Proteção e Controle</div>', 
                            unsafe_allow_html=True)
                
                medidas = {}
                medidas['epi'] = st.text_area(
                    "🥽 EPIs Obrigatórios:", 
                    placeholder="Ex: Capacete, óculos de proteção, luvas, calçado de segurança, protetor auricular",
                    height=80
                )
                medidas['prevencao'] = st.text_area(
                    "🔒 Medidas de Prevenção:", 
                    placeholder="Ex: Treinamentos, procedimentos operacionais, supervisão, manutenção preventiva",
                    height=80
                )
                medidas['empresa'] = st.text_area(
                    "🏢 Medidas da Empresa:", 
                    placeholder="Ex: CIPA, SIPAT, controle médico, monitoramento ambiental",
                    height=80
                )
            
            with col4:
                st.markdown('<div class="section-header">📊 Avaliações e Informações</div>', 
                            unsafe_allow_html=True)
                
                avaliacoes = {}
                avaliacoes['medicoes'] = st.text_area(
                    "📈 Avaliações Ambientais:", 
                    placeholder="Ex: Dosimetria de ruído: 82dB / Iluminação: 300 lux / Calor: IBUTG 25°C",
                    height=100
                )
                
                # Data da OS
                data_os = st.date_input("📅 Data da Ordem de Serviço:", datetime.date.today())
            
            # Botão para gerar OS
            if funcionario_selecionado:
                st.markdown("---")
                col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
                
                with col_btn2:
                    if st.button("📄 GERAR ORDEM DE SERVIÇO", use_container_width=True, type="primary"):
                        try:
                            # Preparar dados do funcionário
                            funcionario_data = {
                                'Nome': funcionario_info['Nome'],
                                'Setor': funcionario_info['Setor'],
                                'Função': funcionario_info['Função'],
                                'Empresa': funcionario_info['Empresa'],
                                'Unidade': funcionario_info['Unidade'],
                                'Data de Admissão': str(funcionario_info.get('Data de Admissão', 'A definir')),
                                'Descrição de Atividades': funcionario_info.get('Descrição de Atividades', 
                                                                              'Atividades relacionadas à função.')
                            }
                            
                            # Criar documento
                            doc = create_os_document(funcionario_data, riscos, medidas, avaliacoes)
                            
                            # Salvar em buffer
                            buffer = BytesIO()
                            doc.save(buffer)
                            buffer.seek(0)
                            
                            # Success message
                            st.markdown('<div class="success-box">', unsafe_allow_html=True)
                            st.markdown("### ✅ Ordem de Serviço Gerada com Sucesso!")
                            st.markdown("O documento foi criado conforme os requisitos da NR-01 e está pronto para download.")
                            st.markdown('</div>', unsafe_allow_html=True)
                            
                            # Botão de download
                            st.download_button(
                                label="💾 📥 BAIXAR ORDEM DE SERVIÇO",
                                data=buffer.getvalue(),
                                file_name=f"OS_{funcionario_info['Nome'].replace(' ', '_')}_{data_os.strftime('%d%m%Y')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                            
                        except Exception as e:
                            st.error(f"❌ Erro ao gerar documento: {e}")
                            st.exception(e)
            
            # Estatísticas da planilha
            st.markdown("---")
            st.markdown("### 📊 Estatísticas da Planilha Carregada")
            
            col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
            
            with col_stat1:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.metric("👥 Funcionários", len(df))
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col_stat2:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.metric("🏢 Setores", df['Setor'].nunique())
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col_stat3:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.metric("💼 Funções", df['Função'].nunique())
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col_stat4:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.metric("🏭 Unidades", df['Unidade'].nunique())
                st.markdown('</div>', unsafe_allow_html=True)
            
        except Exception as e:
            st.error(f"❌ Erro ao processar planilha: {e}")
            st.info("💡 Verifique se o arquivo está no formato Excel (.xlsx) e não está corrompido.")
    
    else:
        # Instruções quando não há arquivo
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        st.markdown("### 🎯 Como começar:")
        st.markdown("""
        1. **📥 Baixe** a planilha exemplo na barra lateral
        2. **✏️ Preencha** com os dados dos seus funcionários  
        3. **📤 Faça upload** da planilha preenchida
        4. **🎯 Selecione** um funcionário e configure os riscos
        5. **📄 Gere** a Ordem de Serviço em formato Word
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="warning-box">', unsafe_allow_html=True)
        st.markdown("### ⚠️ Importante:")
        st.markdown("""
        - A planilha deve estar no formato **.xlsx**
        - Todas as colunas obrigatórias devem estar preenchidas
        - Os dados devem estar organizados em linhas (cada funcionário = uma linha)
        - Não deixe linhas vazias entre os dados
        """)
        st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
