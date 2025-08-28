import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches

import os
import zipfile
from io import BytesIO
import base64
import tempfile

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Ordens de Serviço (OS)",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- Inicialização do Session State ---
if 'descricoes' not in st.session_state:
    st.session_state.descricoes = {}

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
    # Renomeia colunas específicas da planilha PGR para um padrão interno
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
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, str(value))

def substituir_placeholders(doc, contexto):
    """Substitui os placeholders em todo o documento (parágrafos e tabelas)."""
    for p in doc.paragraphs:
        replace_text_in_paragraph(p, contexto)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_text_in_paragraph(p, contexto)

def gerar_os(funcionario, df_pgr, modelo_path, descricao_atividades, riscos_selecionados, epis_manuais, medicoes_manuais, perigo_manual, danos_manuais, categoria_manual, logo_path=None):
    """Gera uma única Ordem de Serviço para um funcionário."""
    doc = Document(modelo_path)

    if logo_path:
        try:
            header_table = doc.tables[0]
            cell = header_table.cell(0, 0)
            cell.text = "" 
            p = cell.paragraphs[0]
            run = p.add_run()
            run.add_picture(logo_path, width=Inches(2.0))
        except IndexError:
            st.warning("Aviso: Não foi encontrada uma tabela de cabeçalho no modelo para inserir a logo.")
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

    data_admissao = "Não informado"
    if 'data_de_admissao' in funcionario and pd.notna(funcionario['data_de_admissao']):
        try:
            data_admissao = pd.to_datetime(funcionario['data_de_admissao']).strftime('%d/%m/%Y')
        except Exception:
            data_admissao = str(funcionario['data_de_admissao'])

    contexto = {
        "[NOME EMPRESA]": str(funcionario.get("empresa", "N/A")), "[UNIDADE]": str(funcionario.get("unidade", "N/A")),
        "[NOME FUNCIONÁRIO]": str(funcionario.get("nome_do_funcionario", "N/A")), "[DATA DE ADMISSÃO]": data_admissao,
        "[SETOR]": str(funcionario.get("setor", "N/A")), "[FUNÇÃO]": str(funcionario.get("funcao", "N/A")),
        "[DESCRIÇÃO DE ATIVIDADES]": descricao_atividades or "Não informado",
        "[RISCOS FÍSICOS]": ", ".join(riscos_por_categoria["fisico"]) or "Nenhum",
        "[RISCOS DE ACIDENTE]": ", ".join(riscos_por_categoria["acidente"]) or "Nenhum",
        "[RISCOS QUÍMICOS]": ", ".join(riscos_por_categoria["quimico"]) or "Nenhum",
        "[RISCOS BIOLÓGICOS]": ", ".join(riscos_por_categoria["biologico"]) or "Nenhum",
        "[RISCOS ERGONÔMICOS]": ", ".join(riscos_por_categoria["ergonomico"]) or "Nenhum",
        "[POSSÍVEIS DANOS RISCOS FÍSICOS]": "; ".join(set(danos_por_categoria["fisico"])) or "Nenhum",
        "[POSSÍVEIS DANOS RISCOS ACIDENTE]": "; ".join(set(danos_por_categoria["acidente"])) or "Nenhum",
        "[POSSÍVEIS DANOS RISCOS QUÍMICOS]": "; ".join(set(danos_por_categoria["quimico"])) or "Nenhum",
        "[POSSÍVEIS DANOS RISCOS BIOLÓGICOS]": "; ".join(set(danos_por_categoria["biologico"])) or "Nenhum",
        "[POSSÍVEIS DANOS RISCOS ERGONÔMICOS]": "; ".join(set(danos_por_categoria["ergonomico"])) or "Nenhum",
        "[EPIS]": ", ".join(sorted(list(epis_recomendados))) or "Nenhum",
        "[MEDIÇÕES]": "\n".join(medicoes_lista) or "Nenhuma medição aplicável.",
    }
    
    substituir_placeholders(doc, contexto)
    return doc

# --- Interface do Streamlit ---
st.title("📄 Gerador de Ordens de Serviço (OS)")

# --- Upload de Arquivos ---
st.sidebar.header("1. Carregar Arquivos")
arquivo_funcionarios = st.sidebar.file_uploader("Planilha de Funcionários", type="xlsx", help="Ficheiro .xlsx obrigatório com os dados dos funcionários.")
arquivo_modelo = st.sidebar.file_uploader("Modelo de OS (Word)", type="docx", help="Ficheiro .docx obrigatório que servirá de modelo.")

arquivo_logo = st.sidebar.file_uploader("Logo da Empresa (Opcional)", type=["png", "jpg", "jpeg"])

# --- Carregamento e Processamento dos Dados ---
df_pgr = pd.DataFrame([{'categoria': 'fisico', 'risco': 'Ruído', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'}, {'categoria': 'fisico', 'risco': 'Vibração', 'possiveis_danos': 'Doenças osteomusculares, problemas circulatórios.'}, {'categoria': 'fisico', 'risco': 'Calor', 'possiveis_danos': 'Desidratação, insolação, cãibras, exaustão, intermação.'}, {'categoria': 'fisico', 'risco': 'Frio', 'possiveis_danos': 'Hipotermia, congelamento, doenças respiratórias.'}, {'categoria': 'fisico', 'risco': 'Radiações Ionizantes', 'possiveis_danos': 'Câncer, mutações genéticas, queimaduras.'}, {'categoria': 'fisico', 'risco': 'Radiações Não Ionizantes', 'possiveis_danos': 'Queimaduras, lesões oculares, câncer de pele.'}, {'categoria': 'fisico', 'risco': 'Pressões Anormais', 'possiveis_danos': 'Doença descompressiva, barotrauma.'}, {'categoria': 'quimico', 'risco': 'Poeiras', 'possiveis_danos': 'Pneumoconioses, irritação respiratória, alergias.'}, {'categoria': 'quimico', 'risco': 'Fumos', 'possiveis_danos': 'Doenças respiratórias, intoxicações.'}, {'categoria': 'quimico', 'risco': 'Névoas', 'possiveis_danos': 'Irritação respiratória, dermatites.'}, {'categoria': 'quimico', 'risco': 'Gases', 'possiveis_danos': 'Asfixia, intoxicações, irritação respiratória.'}, {'categoria': 'quimico', 'risco': 'Vapores', 'possiveis_danos': 'Irritação respiratória, intoxicações, dermatites.'}, {'categoria': 'quimico', 'risco': 'Substâncias Químicas (líquidos e sólidos)', 'possiveis_danos': 'Queimaduras, irritações, intoxicações, dermatites, câncer.'}, {'categoria': 'biologico', 'risco': 'Bactérias', 'possiveis_danos': 'Infecções, doenças infecciosas.'}, {'categoria': 'biologico', 'risco': 'Fungos', 'possiveis_danos': 'Micoses, alergias, infecções respiratórias.'}, {'categoria': 'biologico', 'risco': 'Vírus', 'possiveis_danos': 'Doenças virais, infecções.'}, {'categoria': 'biologico', 'risco': 'Parasitas', 'possiveis_danos': 'Doenças parasitárias, infecções.'}, {'categoria': 'biologico', 'risco': 'Protozoários', 'possiveis_danos': 'Doenças parasitárias.'}, {'categoria': 'ergonomico', 'risco': 'Levantamento e Transporte Manual de Peso', 'possiveis_danos': 'Lesões musculoesqueléticas, dores na coluna.'}, {'categoria': 'ergonomico', 'risco': 'Posturas Inadequadas', 'possiveis_danos': 'Dores musculares, lesões na coluna, LER/DORT.'}, {'categoria': 'ergonomico', 'risco': 'Repetitividade', 'possiveis_danos': 'LER/DORT, tendinites, síndrome do túnel do carpo.'}, {'categoria': 'ergonomico', 'risco': 'Jornada de Trabalho Prolongada', 'possiveis_danos': 'Fadiga, estresse, acidentes de trabalho.'}, {'categoria': 'ergonomico', 'risco': 'Monotonia e Ritmo Excessivo', 'possiveis_danos': 'Estresse, fadiga mental, desmotivação.'}, {'categoria': 'ergonomico', 'risco': 'Controle Rígido de Produtividade', 'possiveis_danos': 'Estresse, ansiedade, burnout.'}, {'categoria': 'ergonomico', 'risco': 'Iluminação Inadequada', 'possiveis_danos': 'Fadiga visual, dores de cabeça.'}, {'categoria': 'ergonomico', 'risco': 'Mobiliário Inadequado', 'possiveis_danos': 'Dores musculares, lesões na coluna.'}, {'categoria': 'acidente', 'risco': 'Arranjo Físico Inadequado', 'possiveis_danos': 'Quedas, colisões, esmagamentos.'}, {'categoria': 'acidente', 'risco': 'Máquinas e Equipamentos sem Proteção', 'possiveis_danos': 'Amputações, cortes, esmagamentos, prensamentos.'}, {'categoria': 'acidente', 'risco': 'Ferramentas Inadequadas ou Defeituosas', 'possiveis_danos': 'Cortes, perfurações, fraturas.'}, {'categoria': 'acidente', 'risco': 'Eletricidade', 'possiveis_danos': 'Choque elétrico, queimaduras, fibrilação ventricular.'}, {'categoria': 'acidente', 'risco': 'Incêndio e Explosão', 'possiveis_danos': 'Queimaduras, asfixia, lesões por impacto.'}, {'categoria': 'acidente', 'risco': 'Animais Peçonhentos', 'possiveis_danos': 'Picadas, mordidas, reações alérgicas, envenenamento.'}, {'categoria': 'acidente', 'risco': 'Armazenamento Inadequado', 'possiveis_danos': 'Quedas de materiais, esmagamentos, soterramentos.'}, {'categoria': 'acidente', 'risco': 'Outros (especificar)', 'possiveis_danos': 'Variados, dependendo do risco específico.'}, {'categoria': 'fisico', 'risco': 'Umidade', 'possiveis_danos': 'Doenças respiratórias, dermatites.'}, {'categoria': 'quimico', 'risco': 'Agrotóxicos', 'possiveis_danos': 'Intoxicações, dermatites, câncer.'}, {'categoria': 'biologico', 'risco': 'Parasitas e Protozoários', 'possiveis_danos': 'Doenças parasitárias.'}, {'categoria': 'ergonomico', 'risco': 'Ritmo de Trabalho Excessivo', 'possiveis_danos': 'Estresse, fadiga, LER/DORT.'}, {'categoria': 'acidente', 'risco': 'Trabalho em Altura', 'possiveis_danos': 'Quedas, fraturas, morte.'}, {'categoria': 'acidente', 'risco': 'Espaços Confinados', 'possiveis_danos': 'Asfixia, intoxicações, explosões.'}, {'categoria': 'acidente', 'risco': 'Condução de Veículos', 'possiveis_danos': 'Acidentes de trânsito, lesões diversas.'}, {'categoria': 'acidente', 'risco': 'Exposição a Agentes Biológicos', 'possiveis_danos': 'Infecções, doenças.'}, {'categoria': 'acidente', 'risco': 'Exposição a Agentes Químicos', 'possiveis_danos': 'Intoxicações, queimaduras.'}, {'categoria': 'acidente', 'risco': 'Exposição a Agentes Físicos', 'possiveis_danos': 'Lesões diversas.'}, {'categoria': 'acidente', 'risco': 'Agressão Física', 'possiveis_danos': 'Lesões físicas e psicológicas.'}, {'categoria': 'acidente', 'risco': 'Assalto/Roubo', 'possiveis_danos': 'Lesões físicas e psicológicas.'}, {'categoria': 'acidente', 'risco': 'Atropelamento', 'possiveis_danos': 'Lesões diversas, morte.'}, {'categoria': 'acidente', 'risco': 'Choque contra', 'possiveis_danos': 'Contusões, fraturas.'}, {'categoria': 'acidente', 'risco': 'Contato com superfície, substância ou objeto quente', 'possiveis_danos': 'Queimaduras.'}, {'categoria': 'acidente', 'risco': 'Contato com superfície, substância ou objeto frio', 'possiveis_danos': 'Queimaduras, hipotermia.'}, {'categoria': 'acidente', 'risco': 'Corte/Laceração', 'possiveis_danos': 'Ferimentos, hemorragias.'}, {'categoria': 'acidente', 'risco': 'Esmagamento', 'possiveis_danos': 'Fraturas, lesões internas, amputações.'}, {'categoria': 'acidente', 'risco': 'Exposição a ruído', 'possiveis_danos': 'Perda auditiva.'}, {'categoria': 'acidente', 'risco': 'Exposição a vibração', 'possiveis_danos': 'Doenças osteomusculares.'}, {'categoria': 'acidente', 'risco': 'Exposição a calor', 'possiveis_danos': 'Desidratação, insolação.'}, {'categoria': 'acidente', 'risco': 'Exposição a frio', 'possiveis_danos': 'Hipotermia, congelamento.'}, {'categoria': 'acidente', 'risco': 'Exposição a radiações', 'possiveis_danos': 'Queimaduras, câncer.'}, {'categoria': 'acidente', 'risco': 'Exposição a pressões anormais', 'possiveis_danos': 'Barotrauma.'}, {'categoria': 'acidente', 'risco': 'Exposição a poeiras', 'possiveis_danos': 'Pneumoconioses.'}, {'categoria': 'acidente', 'risco': 'Exposição a fumos', 'possiveis_danos': 'Doenças respiratórias.'}, {'categoria': 'acidente', 'risco': 'Exposição a névoas', 'possiveis_danos': 'Irritação respiratória.'}, {'categoria': 'acidente', 'risco': 'Exposição a gases', 'possiveis_danos': 'Asfixia, intoxicações.'}, {'categoria': 'acidente', 'risco': 'Exposição a vapores', 'possiveis_danos': 'Irritação respiratória.'}, {'categoria': 'acidente', 'risco': 'Exposição a substâncias químicas', 'possiveis_danos': 'Queimaduras, intoxicações.'}, {'categoria': 'acidente', 'risco': 'Exposição a bactérias', 'possiveis_danos': 'Infecções.'}, {'categoria': 'acidente', 'risco': 'Exposição a fungos', 'possiveis_danos': 'Micoses.'}, {'categoria': 'acidente', 'risco': 'Exposição a vírus', 'possiveis_danos': 'Doenças virais.'}, {'categoria': 'acidente', 'risco': 'Exposição a parasitas', 'possiveis_danos': 'Doenças parasitárias.'}, {'categoria': 'acidente', 'risco': 'Exposição a protozoários', 'possiveis_danos': 'Doenças parasitárias.'}, {'categoria': 'acidente', 'risco': 'Esforço físico intenso', 'possiveis_danos': 'Lesões musculoesqueléticas.'}, {'categoria': 'acidente', 'risco': 'Postura inadequada', 'possiveis_danos': 'Dores musculares.'}, {'categoria': 'acidente', 'risco': 'Repetitividade de movimentos', 'possiveis_danos': 'LER/DORT.'}, {'categoria': 'acidente', 'risco': 'Jornada de trabalho prolongada', 'possiveis_danos': 'Fadiga, estresse.'}, {'categoria': 'acidente', 'risco': 'Monotonia', 'possiveis_danos': 'Estresse, desmotivação.'}, {'categoria': 'acidente', 'risco': 'Ritmo excessivo', 'possiveis_danos': 'Estresse, ansiedade.'}, {'categoria': 'acidente', 'risco': 'Iluminação inadequada', 'possiveis_danos': 'Fadiga visual.'}, {'categoria': 'acidente', 'risco': 'Mobiliário inadequado', 'possiveis_danos': 'Dores musculares.'}, {'categoria': 'acidente', 'risco': 'Arranjo físico inadequado', 'possiveis_danos': 'Quedas, colisões.'}, {'categoria': 'acidente', 'risco': 'Máquinas e equipamentos sem proteção', 'possiveis_danos': 'Amputações, cortes.'}, {'categoria': 'acidente', 'risco': 'Ferramentas inadequadas', 'possiveis_danos': 'Cortes, perfurações.'}, {'categoria': 'acidente', 'risco': 'Eletricidade', 'possiveis_danos': 'Choque elétrico.'}, {'categoria': 'acidente', 'risco': 'Incêndio', 'possiveis_danos': 'Queimaduras, asfixia.'}, {'categoria': 'acidente', 'risco': 'Explosão', 'possiveis_danos': 'Queimaduras, lesões por impacto.'}, {'categoria': 'acidente', 'risco': 'Animais peçonhentos', 'possiveis_danos': 'Picadas, mordidas.'}, {'categoria': 'acidente', 'risco': 'Armazenamento inadequado', 'possiveis_danos': 'Quedas de materiais.'}, {'categoria': 'acidente', 'risco': 'Outros', 'possiveis_danos': 'Variados.'}, {'categoria': 'acidente', 'risco': 'Abertura no piso', 'possiveis_danos': 'Queda de pessoa com diferença de nível.'}, {'categoria': 'acidente', 'risco': 'Agressão de animais', 'possiveis_danos': 'Mordidas, arranhões, infecções.'}, {'categoria': 'acidente', 'risco': 'Agressão de pessoas', 'possiveis_danos': 'Lesões físicas, traumas psicológicos.'}, {'categoria': 'acidente', 'risco': 'Armazenamento de produtos químicos', 'possiveis_danos': 'Vazamentos, derramamentos, incêndios, explosões.'}, {'categoria': 'acidente', 'risco': 'Ausência de sinalização', 'possiveis_danos': 'Acidentes diversos.'}, {'categoria': 'acidente', 'risco': 'Caminhão/Carreta', 'possiveis_danos': 'Atropelamento, colisão, esmagamento.'}, {'categoria': 'acidente', 'risco': 'Carga e descarga de materiais', 'possiveis_danos': 'Quedas, esmagamentos, lesões por esforço.'}, {'categoria': 'acidente', 'risco': 'Choque elétrico', 'possiveis_danos': 'Queimaduras, parada cardíaca.'}, {'categoria': 'acidente', 'risco': 'Colisão', 'possiveis_danos': 'Contusões, fraturas, lesões internas.'}, {'categoria': 'acidente', 'risco': 'Contato com máquinas e/ou equipamentos', 'possiveis_danos': 'Prensamento, corte, esmagamento.'}, {'categoria': 'acidente', 'risco': 'Contato com objetos cortantes/perfurocortantes', 'possiveis_danos': 'Cortes, perfurações.'}, {'categoria': 'acidente', 'risco': 'Contato com produtos químicos', 'possiveis_danos': 'Queimaduras, irritações, intoxicações.'}, {'categoria': 'acidente', 'risco': 'Desabamento/Soterramento', 'possiveis_danos': 'Asfixia, esmagamento, morte.'}, {'categoria': 'acidente', 'risco': 'Empilhadeira', 'possiveis_danos': 'Atropelamento, colisão, esmagamento.'}, {'categoria': 'acidente', 'risco': 'Escada (móvel ou fixa)', 'possiveis_danos': 'Quedas, fraturas.'}, {'categoria': 'acidente', 'risco': 'Explosão', 'possiveis_danos': 'Queimaduras, lesões por impacto, morte.'}, {'categoria': 'acidente', 'risco': 'Exposição a agentes biológicos', 'possiveis_danos': 'Infecções, doenças.'}, {'categoria': 'acidente', 'risco': 'Exposição a agentes químicos', 'possiveis_danos': 'Intoxicações, queimaduras.'}, {'categoria': 'acidente', 'risco': 'Exposição a agentes físicos', 'possiveis_danos': 'Lesões diversas.'}, {'categoria': 'acidente', 'risco': 'Faca/Estilete', 'possiveis_danos': 'Cortes, perfurações.'}, {'categoria': 'acidente', 'risco': 'Ferramentas manuais', 'possiveis_danos': 'Cortes, contusões.'}, {'categoria': 'acidente', 'risco': 'Gases inflamáveis', 'possiveis_danos': 'Incêndio, explosão.'}, {'categoria': 'acidente', 'risco': 'Impacto de objetos', 'possiveis_danos': 'Contusões, fraturas.'}, {'categoria': 'acidente', 'risco': 'Incêndio', 'possiveis_danos': 'Queimaduras, asfixia, morte.'}, {'categoria': 'acidente', 'risco': 'Ingestão de substância cáustica, tóxica ou nociva.', 'possiveis_danos': 'Intoxicação, envenenamento, queimadura, irritação ou reação alérgica.'}, {'categoria': 'acidente', 'risco': 'Inalação, ingestão e/ou absorção.', 'possiveis_danos': 'Intoxicação, envenenamento, queimadura, irritação ou reação alérgica.'}, {'categoria': 'acidente', 'risco': 'Incêndio/Explosão', 'possiveis_danos': 'Queimadura de 1º, 2º ou 3º grau, asfixia,  arremessos, cortes, escoriações, luxações, fraturas.'}, {'categoria': 'acidente', 'risco': 'Objetos cortantes/perfurocortantes', 'possiveis_danos': 'Corte, laceração, ferida contusa, punctura (ferida aberta), perfuração.'}, {'categoria': 'acidente', 'risco': 'Pessoas não autorizadas e/ou visitantes no local de trabalho', 'possiveis_danos': 'Escoriação, ferimento, corte, luxação, fratura, entre outros danos devido às características do local e atividades realizadas.'}, {'categoria': 'acidente', 'risco': 'Portas, escotilhas, tampas, "bocas de visita", flanges', 'possiveis_danos': 'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações, exposição à gases tóxicos.'}, {'categoria': 'acidente', 'risco': 'Projeção de Partículas sólidas e/ou líquidas', 'possiveis_danos': 'Ferimento, corte, queimadura, perfuração, intoxicação.'}, {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível de andaime, passarela, plataforma, etc.', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'}, {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível de escada (móvel ou fixa).', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'}, {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível de material empilhado.', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'}, {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível de veículo.', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'}, {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível em poço, escavação, abertura no piso, etc.', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'}, {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível <= 2m', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'}, {'categoria': 'acidente', 'risco': 'Queda de pessoa com diferença de nível > 2m', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas, morte.'}, {'categoria': 'acidente', 'risco': 'Queda de pessoa em mesmo nível', 'possiveis_danos': 'Escoriações, ferimentos, cortes, luxações, fraturas.'}, {'categoria': 'acidente', 'risco': 'Reação do corpo a seus movimentos (escorregão sem queda, etc.)', 'possiveis_danos': 'Torções, distensções, rupturas ou outras lesões musculares internas.'}, {'categoria': 'acidente', 'risco': 'Vidro (recipientes, portas, bancadas, janelas, objetos diversos).', 'possiveis_danos': 'Corte, ferimento, perfuração.'}, {'categoria': 'acidente', 'risco': 'Soterramento', 'possiveis_danos': 'Asfixia, desconforto respiratório, nível de consciência alterado, letargia, palidez, pele azulada, tosse, transtorno neurológico.'}, {'categoria': 'acidente', 'risco': 'Substâncias tóxicas e/ou inflamáveis', 'possiveis_danos': 'Intoxicação, asfixia, queimaduras de  1º, 2º ou 3º grau.'}, {'categoria': 'acidente', 'risco': 'Superfícies, substâncias e/ou objetos aquecidos ', 'possiveis_danos': 'Queimadura de 1º, 2º ou 3º grau.'}, {'categoria': 'acidente', 'risco': 'Superfícies, substâncias e/ou objetos em baixa temperatura ', 'possiveis_danos': 'Queimadura de 1º, 2º ou 3º grau.'}, {'categoria': 'acidente', 'risco': 'Tombamento, quebra e/ou ruptura de estrutura (fixa ou móvel)', 'possiveis_danos': 'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.'}, {'categoria': 'acidente', 'risco': 'Tombamento de máquina/equipamento', 'possiveis_danos': 'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.'}, {'categoria': 'acidente', 'risco': 'Trabalho à céu aberto', 'possiveis_danos': 'Intermação, insolação, cãibra, exaustão, desidratação, resfriados.'}, {'categoria': 'acidente', 'risco': 'Trabalho em espaços confinados', 'possiveis_danos': 'Asfixia, hiperóxia, contaminação por poeiras e/ou gases tóxicos, queimadura de 1º, 2º ou 3º grau, arremessos, cortes, escoriações, luxações, fraturas.'}, {'categoria': 'acidente', 'risco': 'Trabalho com máquinas portáteis rotativas.', 'possiveis_danos': 'Cortes, ferimentos, escoriações, amputações.'}, {'categoria': 'acidente', 'risco': 'Trabalho com máquinas e/ou equipamentos', 'possiveis_danos': 'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações, choque elétrico.'}])
df_funcionarios_raw = carregar_planilha(arquivo_funcionarios)
df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_raw) if df_funcionarios_raw is not None else None

# --- Validação e Entrada de Dados Faltantes ---
colunas_ok = False
if df_funcionarios is not None:
    colunas_necessarias = ['nome_do_funcionario', 'funcao', 'setor', 'data_de_admissao']
    colunas_faltantes = [col for col in colunas_necessarias if col not in df_funcionarios.columns]

    if colunas_faltantes:
        st.error(f"**Erro na Planilha de Funcionários!**\n\n"
                 f"As seguintes colunas obrigatórias não foram reconhecidas: **{', '.join(colunas_faltantes)}**.\n\n"
                 f"Verifique o seu ficheiro .xlsx.")
    else:
        if 'empresa' not in df_funcionarios.columns:
            empresa_input = st.text_input("Nome da Empresa", help="Esta informação não foi encontrada na planilha.")
            if empresa_input: df_funcionarios['empresa'] = empresa_input

        if 'unidade' not in df_funcionarios.columns:
            unidade_input = st.text_input("Nome da Unidade/Filial", help="Esta informação não foi encontrada na planilha.")
            if unidade_input: df_funcionarios['unidade'] = unidade_input

        if 'empresa' in df_funcionarios.columns and 'unidade' in df_funcionarios.columns:
            colunas_ok = True

# --- Interface Principal (só aparece se as colunas estiverem OK) ---
if colunas_ok:
    st.sidebar.header("2. Seleção e Filtros")
    empresa_selecionada = st.sidebar.selectbox("Empresa", df_funcionarios["empresa"].unique())
    unidade_selecionada = st.sidebar.selectbox("Unidade", df_funcionarios[df_funcionarios["empresa"] == empresa_selecionada]["unidade"].unique())
    setor_selecionado = st.sidebar.selectbox("Setor", df_funcionarios[(df_funcionarios["empresa"] == empresa_selecionada) & (df_funcionarios["unidade"] == unidade_selecionada)]["setor"].unique())
    funcao_selecionada = st.sidebar.selectbox("Função", df_funcionarios[(df_funcionarios["empresa"] == empresa_selecionada) & (df_funcionarios["unidade"] == unidade_selecionada) & (df_funcionarios["setor"] == setor_selecionado)]["funcao"].unique())

    # Contagem de funcionários na hierarquia selecionada
    funcionarios_filtrados = df_funcionarios[
        (df_funcionarios["empresa"] == empresa_selecionada) &
        (df_funcionarios["unidade"] == unidade_selecionada) &
        (df_funcionarios["setor"] == setor_selecionado) &
        (df_funcionarios["funcao"] == funcao_selecionada)
    ]
    total_funcionarios = len(funcionarios_filtrados)
    
    # Exibir contagem na sidebar
    st.sidebar.info(f"**Funcionários encontrados:** {total_funcionarios}")
    
    if total_funcionarios == 0:
        st.sidebar.warning("Nenhum funcionário encontrado para esta seleção.")

    st.subheader("Informações da Função")
    descricao_atividades_input = st.text_area(f"Descrição das Atividades da Função: **{funcao_selecionada}**",
                                              value=st.session_state.descricoes.get(funcao_selecionada, ""), height=150)
    st.session_state.descricoes[funcao_selecionada] = descricao_atividades_input
    
    col1, col2 = st.columns(2)
    with col1:
        epis_manuais_input = st.text_input("Adicionar EPIs", help="Adicione EPIs extras separados por vírgula (,).")
    with col2:
        medicoes_manuais_input = st.text_area("Adicionar Medições", height=100, help="Adicione medições extras, uma por linha.")
    
    with st.expander("Adicionar Risco Manualmente"):
        perigo_manual_input = st.text_input("Perigo (Fator de Risco/Agente Nocivo/Situação Perigosa)")
        danos_manuais_input = st.text_input("Possíveis Danos ou Agravos à Saúde")
        categoria_manual_input = st.selectbox("Selecione a Categoria para o Risco Manual",
                                              options=["Acidentes", "Biológicos", "Ergonômicos", "Físicos", "Químicos"])

    st.subheader("Seleção de Riscos")
    categorias = {"Físicos": "fisico", "Químicos": "quimico", "Biológicos": "biologico", "Ergonômicos": "ergonomico", "Acidentes": "acidente"}
    riscos_selecionados_total = []
    
    # Filtra a planilha PGR pela função selecionada
    riscos_da_funcao = df_pgr[df_pgr['funcao'] == funcao_selecionada] if 'funcao' in df_pgr.columns else df_pgr.copy()

    for nome_exibicao, nome_categoria in categorias.items():
        with st.expander(f"Riscos {nome_exibicao}"):
            # Lista apenas os riscos da função selecionada para a categoria atual
            riscos_para_selecao = riscos_da_funcao[riscos_da_funcao['categoria'] == nome_categoria]['risco'].unique().tolist()
            
            # Se não há riscos cadastrados para esta categoria, permite adição manual
            if not riscos_para_selecao:
                st.info(f"Nenhum risco de {nome_exibicao.lower()} cadastrado para esta função. Você pode adicionar riscos manualmente abaixo.")
                
                # Campo para adicionar riscos manualmente
                riscos_manuais_key = f"riscos_manuais_{nome_categoria}"
                if riscos_manuais_key not in st.session_state:
                    st.session_state[riscos_manuais_key] = []
                
                col1, col2 = st.columns([3, 1])
                with col1:
                    novo_risco = st.text_input(f"Adicionar risco {nome_exibicao.lower()}", key=f"input_{nome_categoria}")
                with col2:
                    if st.button("Adicionar", key=f"btn_{nome_categoria}"):
                        if novo_risco and novo_risco not in st.session_state[riscos_manuais_key]:
                            st.session_state[riscos_manuais_key].append(novo_risco)
                            st.rerun()
                
                # Exibe os riscos adicionados manualmente
                if st.session_state[riscos_manuais_key]:
                    st.write(f"Riscos {nome_exibicao.lower()} adicionados:")
                    for i, risco in enumerate(st.session_state[riscos_manuais_key]):
                        col1, col2 = st.columns([4, 1])
                        with col1:
                            st.write(f"• {risco}")
                        with col2:
                            if st.button("Remover", key=f"remove_{nome_categoria}_{i}"):
                                st.session_state[riscos_manuais_key].remove(risco)
                                st.rerun()
                
                # Multiselect com os riscos adicionados manualmente
                selecionados = st.multiselect(f"Selecione os riscos de {nome_exibicao.lower()}",
                                              options=st.session_state[riscos_manuais_key], default=[], key=f"riscos_{nome_categoria}")
            else:
                # Multiselect normal com os riscos da planilha
                selecionados = st.multiselect(f"Selecione os riscos de {nome_exibicao.lower()}",
                                              options=riscos_para_selecao, default=[], key=f"riscos_{nome_categoria}")
            
            riscos_selecionados_total.extend(selecionados)

    st.sidebar.header("3. Selecionar Funcionários")
    nomes_funcionarios = funcionarios_filtrados["nome_do_funcionario"].tolist()
    selecionar_todos = st.sidebar.checkbox("Selecionar Todos os Funcionários")
    if selecionar_todos:
        funcionarios_selecionados = nomes_funcionarios
    else:
        funcionarios_selecionados = st.sidebar.multiselect("Escolha os funcionários", options=nomes_funcionarios, default=[])

    st.header("4. Geração dos Documentos")
    if st.button("🚀 Gerar OS", type="primary"):
        if not all([arquivo_funcionarios, arquivo_modelo]):
            st.error("Erro: Por favor, carregue os ficheiros de Funcionários e Modelo de OS.")
        elif not funcionarios_selecionados:
            st.warning("Atenção: Nenhum funcionário foi selecionado.")
        else:
            with st.spinner("A gerar documentos..."):
                with tempfile.TemporaryDirectory() as tmpdirname:
                    zip_path = os.path.join(tmpdirname, "ordens_servico.zip")
                    with zipfile.ZipFile(zip_path, "w") as zipf:
                        for nome_func in funcionarios_selecionados:
                            dados_funcionario = funcionarios_filtrados[funcionarios_filtrados["nome_do_funcionario"] == nome_func].iloc[0]
                            doc_gerado = gerar_os(
                                dados_funcionario, df_pgr, 
                                arquivo_modelo, descricao_atividades_input, 
                                riscos_selecionados_total, epis_manuais_input, 
                                medicoes_manuais_input, perigo_manual_input, 
                                danos_manuais_input, categoria_manual_input, logo_path=arquivo_logo
                            )
                            
                            nome_arquivo_docx = f"OS_{nome_func.replace(' ', '_')}.docx"
                            caminho_docx = os.path.join(tmpdirname, nome_arquivo_docx)
                            doc_gerado.save(caminho_docx)
                            zipf.write(caminho_docx, os.path.basename(caminho_docx))


                    
                    with open(zip_path, "rb") as f:
                        st.download_button(
                            label="📥 Baixar Ficheiro ZIP",
                            data=f.read(),
                            file_name="Ordens_de_Servico.zip",
                            mime="application/zip"
                        )

elif df_funcionarios is None:
    st.warning("⚠️ **Planilha de Funcionários não carregada!**\n\nPor favor, carregue a planilha de funcionários para continuar.")

else:
    st.info("⏳ **Aguardando informações...**\n\nPreencha os campos obrigatórios para continuar.")

