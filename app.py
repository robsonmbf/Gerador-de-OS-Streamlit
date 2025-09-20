import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import time
import re
import hashlib
import datetime

# --- Constantes e dados de risco aqui (mesmo que já definidos previamente) ---

UNIDADES_DE_MEDIDA = ["dB(A)", "m/s²", "ppm", "mg/m³", "%", "°C", "lx", "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"]

# Definição de riscos Físico, Químico, Biológico, Ergonômico, Acidente
# [Coloque aqui as listas RISCOS_FISICO, RISCOS_QUIMICO, etc.]

AGENTES_POR_CATEGORIA = {
    'fisico': RISCOS_FISICO,
    'quimico': RISCOS_QUIMICO,
    'biologico': RISCOS_BIOLOGICO,
    'ergonomico': RISCOS_ERGONOMICO,
    'acidente': RISCOS_ACIDENTE,
}

CATEGORIAS_RISCO = {
    'fisico': '🔥 Físicos',
    'quimico': '⚗️ Químicos',
    'biologico': '🦠 Biológicos',
    'ergonomico': '🏃 Ergonômicos',
    'acidente': '⚠️ Acidentes'
}

# Funções auxiliares para autenticação, validação de email, criação de amostra, etc. (igual antes)...

def substituir_placeholders_no_documento(doc, dados_funcionario, agentes_risco, epis, medicoes=""):
    try:
        substituicoes = {
            '[NOME EMPRESA]': str(dados_funcionario.get('Empresa', '') or ''),
            '[UNIDADE]': str(dados_funcionario.get('Unidade', '') or ''),
            '[NOME FUNCIONÁRIO]': str(dados_funcionario.get('Nome', '') or ''),
            '[DATA DE ADMISSÃO]': str(dados_funcionario.get('Data de Admissão', '') or ''),
            '[SETOR]': str(dados_funcionario.get('Setor', '') or ''),
            '[FUNÇÃO]': str(dados_funcionario.get('Função', '') or ''),
            '[DESCRIÇÃO DE ATIVIDADES]': str(dados_funcionario.get('Descrição de Atividades', '') or ''),
            '[MEDIÇÕES]': str(medicoes) if medicoes else "Não aplicável para esta função.",
        }
        riscos_texto = {}
        danos_texto = {}

        for categoria in ['fisico', 'quimico', 'biologico', 'ergonomico', 'acidente']:
            if categoria == 'fisico': categoria_nome = 'FÍSICOS'
            elif categoria == 'quimico': categoria_nome = 'QUÍMICOS'
            elif categoria == 'biologico': categoria_nome = 'BIOLÓGICOS'
            elif categoria == 'ergonomico': categoria_nome = 'ERGONÔMICOS'
            else: categoria_nome = 'ACIDENTE'

            if categoria in agentes_risco and agentes_risco[categoria]:
                riscos_lista = []
                for risco in agentes_risco[categoria]:
                    risco_text = str(risco['agente'])
                    if risco.get('intensidade'):
                        risco_text += f": {risco['intensidade']}"
                    if risco.get('unidade') and risco['unidade'] != 'Não aplicável':
                        risco_text += f" {risco['unidade']}"
                    riscos_lista.append(risco_text)
                riscos_texto[f'[RISCOS {categoria_nome}]'] = '; '.join(riscos_lista)

                # Possíveis danos genéricos
                if categoria == 'fisico':
                    danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "Perda auditiva, lesões por vibração, queimaduras, hipotermia, hipertermia"
                elif categoria == 'quimico':
                    danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "Intoxicação, dermatoses, pneumoconioses, alergias respiratórias"
                elif categoria == 'biologico':
                    danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "Infecções, doenças infectocontagiosas, alergias"
                elif categoria == 'ergonomico':
                    danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "LER/DORT, fadiga, estresse, dores musculares"
                else:
                    danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "Fraturas, cortes, contusões, queimaduras, morte"
            else:
                riscos_texto[f'[RISCOS {categoria_nome}]'] = "Ausência de Fator de Risco"
                danos_texto[f'[POSSÍVEIS DANOS RISCOS {categoria_nome}]'] = "Não aplicável"

        substituicoes.update(riscos_texto)
        substituicoes.update(danos_texto)

        substituicoes['[EPIS]'] = '; '.join(map(str, epis)) if epis else "Conforme análise de risco específica da função"

        for paragrafo in doc.paragraphs:
            for placeholder, valor in substituicoes.items():
                if placeholder in paragrafo.text:
                    paragrafo.text = paragrafo.text.replace(placeholder, str(valor))

        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for placeholder, valor in substituicoes.items():
                        if placeholder in celula.text:
                            celula.text = celula.text.replace(placeholder, str(valor))

        return doc
    except Exception as e:
        st.error(f"Erro ao substituir placeholders: {str(e)}")
        return None

def gerar_documento_os(dados_funcionario, agentes_risco, epis, uploaded_template=None):
    try:
        if uploaded_template:
            doc = Document(uploaded_template)
            doc = substituir_placeholders_no_documento(doc, dados_funcionario, agentes_risco, epis)
        else:
            doc = Document()
            titulo = doc.add_heading('ORDEM DE SERVIÇO', 0)
            titulo.alignment = 1
            subtitulo = doc.add_paragraph('Informações sobre Condições de Segurança e Saúde no Trabalho - NR-01')
            subtitulo.alignment = 1

            doc.add_paragraph()
            info_func = doc.add_paragraph()
            info_func.add_run(f"Empresa: {dados_funcionario.get('Empresa', '')}\t\t")
            info_func.add_run(f"Unidade: {dados_funcionario.get('Unidade', '')}")

            info_func2 = doc.add_paragraph()
            info_func2.add_run(f"Nome do Funcionário: {dados_funcionario.get('Nome', '')}")

            info_func3 = doc.add_paragraph()
            info_func3.add_run(f"Data de Admissão: {dados_funcionario.get('Data de Admissão', '')}")

            info_func4 = doc.add_paragraph()
            info_func4.add_run(f"Setor de Trabalho: {dados_funcionario.get('Setor', '')}\t\t")
            info_func4.add_run(f"Função: {dados_funcionario.get('Função', '')}")

            doc.add_paragraph()

            doc.add_heading('TAREFAS DA FUNÇÃO', level=1)
            doc.add_paragraph(dados_funcionario.get('Descrição de Atividades', 'Atividades relacionadas à função exercida.'))

            tem_riscos = any(agentes_risco.get(categoria, []) for categoria in agentes_risco.keys()) if agentes_risco else False
            doc.add_heading('AGENTES DE RISCOS OCUPACIONAIS', level=1)
            if tem_riscos:
                for categoria, riscos in agentes_risco.items():
                    if riscos:
                        categoria_titulo = categoria.replace('_', ' ').title()
                        doc.add_heading(f'Riscos {categoria_titulo}', level=2)
                        for risco in riscos:
                            par = doc.add_paragraph()
                            par.add_run(f"• {risco['agente']}")
                            if risco.get('intensidade'):
                                par.add_run(f": {risco['intensidade']}")
                            if risco.get('unidade'):
                                par.add_run(f" {risco['unidade']}")
            else:
                doc.add_paragraph("Ausência de Fator de Risco")

            if epis:
                doc.add_heading('EQUIPAMENTOS DE PROTEÇÃO INDIVIDUAL (EPIs)', level=1)
                for epi in epis:
                    doc.add_paragraph(f"• {epi}", style='List Bullet')
            else:
                doc.add_heading('EQUIPAMENTOS DE PROTEÇÃO INDIVIDUAL (EPIs)', level=1)
                doc.add_paragraph("Conforme análise de risco específica da função")

            doc.add_paragraph()
            nota_legal = doc.add_paragraph()
            nota_legal.add_run("IMPORTANTE: ").bold = True
            nota_legal.add_run(
                "Conforme Art. 158 da CLT e NR-01, o descumprimento das disposições "
                "sobre segurança e saúde no trabalho sujeita o empregado às penalidades "
                "legais, inclusive demissão por justa causa."
            )

            doc.add_paragraph()
            doc.add_paragraph("_" * 40 + "\t\t" + "_" * 40)
            doc.add_paragraph("Funcionário\t\t\t\t\tResponsável pela Área")
            doc.add_paragraph(f"Data: {datetime.date.today().strftime('%d/%m/%Y')}")

        return doc

    except Exception as e:
        st.error(f"Erro ao gerar documento: {str(e)}")
        return None

# O restante do app segue como anteriormente configurado, incluindo autenticação, upload, filtros e geração.

