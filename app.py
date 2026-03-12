
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import zipfile
from io import BytesIO
import time
import re
import sys
import os

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from database.models import DatabaseManager
from database.auth import AuthManager
from database.user_data import UserDataManager

st.set_page_config(
    page_title="Gerador de Ordens de Serviço (OS)",
    page_icon="📄",
    layout="wide",
)

UNIDADES_DE_MEDIDA = ["dB(A)", "m/s²", "m/s¹⋅⁷⁵", "ppm", "mg/m³", "%", "°C", "lx", "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"]
AGENTES_DE_RISCO = sorted([
    "Ruído (Contínuo ou Intermitente)", "Ruído (Impacto)", "Vibração de Corpo Inteiro", "Vibração de Mãos e Braços",
    "Radiações Ionizantes", "Radiações Não-Ionizantes", "Frio", "Calor", "Pressões Anormais", "Umidade", "Poeiras", 
    "Fumos", "Névoas", "Neblinas", "Gases", "Vapores", "Produtos Químicos em Geral", "Vírus", "Bactérias", 
    "Protozoários", "Fungos", "Parasitas", "Bacilos"
])
CATEGORIAS_RISCO = {'fisico': '🔥 Físicos', 'quimico': '⚗️ Químicos', 'biologico': '🦠 Biológicos', 'ergonomico': '🏃 Ergonômicos', 'acidente': '⚠️ Acidentes'}

RISCOS_PGR_DADOS = {
    'quimico': {
        'riscos': ['Exposição a Produto Químico'],
        'danos': ['Irritação/lesão ocular, na pele e mucosas; Dermatites; Queimadura Química; Intoxicação; Náuseas; Vômitos.']
    },
    'fisico': {
        'riscos': [
            'Ambiente Artificialmente Frio', 'Exposição ao Ruído', 'Vibrações Localizadas (mão/braço)',
            'Vibração de Corpo Inteiro (AREN)', 'Vibração de Corpo Inteiro (VDVR)', 'Exposição à Radiações Ionizantes',
            'Exposição à Radiações Não-ionizantes', 'Exposição à Temperatura Ambiente Elevada',
            'Exposição à Temperatura Ambiente Baixa', 'Pressão Atmosférica Anormal (condições hiperbáricas)', 'Umidade'
        ],
        'danos': [
            'Estresse, desconforto, dormência, rigidez nas partes com maior intensidade de exposição ao frio, redução da destreza, formigamento, redução da sensibilidade dos dedos e flexibilidade das articulações.',
            'Perda Auditiva Induzida pelo Ruído Ocupacional (PAIRO).',
            'Alterações articulares e vasomotoras.',
            'Alterações no sistema digestivo, sistema musculoesquelético, sistema nervoso, alterações na visão, enjoos, náuseas, palidez.',
            'Alterações no sistema digestivo, sistema musculoesquelético, sistema nervoso, alterações na visão, enjoos, náuseas, palidez.',
            'Dano às células do corpo humano, causando doenças graves, inclusive fatais, como câncer.',
            'Depressão imunológica, fotoenvelhecimento, lesões oculares como ceratoconjuntivite, pterígio e catarata; Doenças graves, inclusives fatais, como câncer.',
            'Desidratação, erupções cutâneas, câibras, fadiga física, problemas cardiocirculatórios, distúrbios psicológicos.',
            'Estresse, desconforto, dormência, rigidez nas partes com maior intensidade de exposição ao frio, redução da destreza, formigamento, redução da sensibilidade dos dedos e flexibilidade das articulações.',
            'Barotrauma pulmonar, lesão de tecido pulmonar ou pneumotórax, embolia arterial gasosa, barotrauma de ouvido, barotrauma sinusal, barotrauma dental, barotrauma facial, doença descompressiva.',
            'Doenças do aparelho respiratório, quedas, doenças de pele, doenças circulatórias, entre outras.'
        ]
    },
    'biologico': {
        'riscos': [
            'Água e/ou alimentos contaminados',
            'Contato com Fluido Orgânico (sangue, hemoderivados, secreções, excreções)',
            'Contato com Pessoas Doentes e/ou Material Infectocontagiante',
            'Contaminação pelo Corona Vírus',
            'Exposição à Agentes Microbiológicos (fungos, bactérias, vírus, protozoários, parasitas)'
        ],
        'danos': [
            'Intoxicação, diarreias, infecções intestinais.',
            'Doenças infectocontagiosas.',
            'Doenças infectocontagiosas.',
            'COVID-19, podendo causar gripes, febre, tosse seca, cansaço, dores e desconfortos, dor de garganta, diarreia, perda de paladar ou olfato, dificuldade de respirar ou falta de ar, dor ou pressão no peito, perda de fala ou movimentos.',
            'Doenças infectocontagiosas, dermatites, irritação, desconforto, infecção do sistema respiratório.'
        ]
    },
    'ergonomico': {
        'riscos': [
            'Posturas incômodas/pouco confortáveis por longos períodos', 'Postura sentada por longos períodos',
            'Postura em pé por longos períodos', 'Frequente deslocamento à pé durante à jornada de trabalho',
            'Esforço físico intenso', 'Levantamento e transporte manual de cargas ou volumes',
            'Frequente ação de empurrar/puxar cargas ou volumes', 'Frequente execução de movimentos repetitivos',
            'Manuseio de ferramentas e/ou objetos pesados por longos períodos',
            'Uso frequente de força, pressão, preensão, flexão, extensão ou torção dos segmentos corporais',
            'Compressão de partes do corpo por superfícies rígidas ou com quinas vivas',
            'Flexões da coluna vertebral frequentes', 'Uso frequente de pedais', 'Uso frequente de alavancas',
            'Elevação frequente de membros superiores',
            'Manuseio ou movimentação de cargas e volumes sem pega ou com "pega pobre"',
            'Exposição à vibração de corpo inteiro', 'Exposição à vibrações localizadas (mão, braço)',
            'Uso frequente de escadas', 'Trabalho intensivo com teclado ou outros dispositivos de entrada de dados',
            'Posto de trabalho improvisado/inadequado', 'Mobiliário sem meios de regulagem de ajustes',
            'Equipamentos e/ou máquinas sem meios de regulagem de ajustes ou sem condições de uso',
            'Posto de trabalho não planejado/adaptado para à posição sentada', 'Assento inadequado',
            'Encosto do assento inadequado ou ausente',
            'Mobiliário ou equipamento sem espaço para movimentação de segmentos corporais',
            'Necessidade de alcançar objetos, documentos, controles, etc, além das zonas de alcance ideais',
            'Equipamentos/mobiliário não adaptados à antropometria do trabalhador',
            'Trabalho realizado sem pausas pré-definidas para descanso',
            'Necessidade de manter ritmos intensos de trabalho', 'Trabalho com necessidade de variação de turnos',
            'Monotonia', 'Trabalho noturno', 'Insuficiência de capacitação para à execução da tarefa',
            'Trabalho com utilização rigorosa de metas de produção', 'Trabalho remunerado por produção',
            'Cadência do trabalho imposta por um equipamento',
            'Desequilíbrio entre tempo de trabalho e tempo de repouso',
            'Pressão sonora fora dos parâmetros de conforto', 'Temperatura efetiva fora dos parâmetros de conforto',
            'Velocidade do ar fora dos parâmetros de conforto', 'Umidade do ar fora dos parâmetros de conforto',
            'Iluminação inadequada', 'Reflexos que causem desconforto ou prejudiquem à visão',
            'Piso escorregadio ou irregular', 'Situações de estresse no local de trabalho',
            'Situações de sobrecarga de trabalho mental', 'Exigência de concentração, atenção e memória',
            'Trabalho em condições de difícil comunicação', 'Conflitos hierárquicos no trabalho',
            'Problemas de relacionamento no trabalho', 'Assédio de qualquer natureza no trabalho',
            'Dificuldades para cumprir ordens e determinações da chefia relacionadas ao trabalho',
            'Realização de múltiplas tarefas com alta demanda mental/cognitiva', 'Insatisfação no trabalho',
            'Falta de autonomia para a realização de tarefas no trabalho'
        ],
        'danos': [
            'Distúrbios musculoesqueléticos em músculos e articulações dos membros superiores, inferiores e coluna.',
            'Sobrecarga dos membros superiores e coluna vertebral; Aumento na pressão dos discos intervertebrais; Dor localizada.',
            'Sobrecarga corporal, dores nos membros inferiores e em alguns casos na coluna vertebral e cansaço físico.',
            'Sobrecarga corporal, dores nos membros inferiores e em alguns casos na coluna vertebral e cansaço físico.',
            'Distúrbios musculoesqueléticos; Fadiga, Dor localizada; Redução da produtividade e da percepção de risco.',
            'Distúrbios musculoesqueléticos; Fadiga, Dor localizada; Redução da produtividade e da percepção de risco.',
            'Distúrbios musculoesqueléticos em músculos e articulações dos membros superiores, inferiores e coluna lombar.',
            'Distúrbios osteomusculares em músculos e articulações dos membros utilizados na execução dos movimentos repetitivos.',
            'Fadiga muscular; Dor localizada; Lesões musculares; Redução da produtividade e da percepção de risco.',
            'Sobrecarga muscular, fadiga, dor localizada e perda de produtividade.',
            'Restrição localizada temporária do fluxo cardiovascular.',
            'Tensão na parte inferior das costas (coluna lombar), podendo causar fadiga, dor localizada e/ou lesões musculoesqueléticas.',
            'Distúrbio musculoesqueléticos em músculos e articulações dos membros inferiores.',
            'Distúrbios musculoesqueléticos em músculos e articulações dos membros superiores.',
            'Sobrecarga na região do pescoço, ombros e braços, podendo causar fadiga e/ou dor localizada.',
            'Sobrecarga corporal, aumento da força durante o manuseio, fadiga, dor localizada e perda de produtividade.',
            'Alterações no sistema digestivo, sistema musculoesquelético, sistema nervoso, alterações na visão, enjoos, náuseas, palidez.',
            'Alterações articulares e vasomotoras.',
            'Distúrbios musculoesqueléticos em músculos e articulações dos membros inferiores.',
            'Sobrecarga nas articulações das mãos, punhos e antebraços, podendo causar lesões como artrite e dificuldade de flexão.',
            'Adoção de movimentos e posturas inadequadas; Fadiga muscular; Dor localizada; Distúrbios musculoesqueléticos.',
            'Adoção de movimentos e posturas inadequadas; Fadiga muscular; Dor localizada; Distúrbios musculoesqueléticos.',
            'Adoção de movimentos e posturas inadequadas; Fadiga muscular; Dor localizada; Distúrbios musculoesqueléticos.',
            'Sobrecarga dos membros superiores e coluna vertebral; Aumento na pressão dos discos intervertebrais; Dor localizada.',
            'Sobrecarga corporal e dores nos membros superiores, inferiores e coluna vertebral.',
            'Sobrecarga corporal e dores na região da coluna vertebral.',
            'Adoção de movimentos e posturas inadequadas; Fadiga muscular; Dor localizada; Distúrbios musculoesqueléticos.',
            'Adoção de movimentos e posturas inadequadas; Fadiga muscular; Dor localizada; Distúrbios musculoesqueléticos.',
            'Adoção de movimentos e posturas inadequadas; Fadiga muscular; Dor localizada; Distúrbios musculoesqueléticos.',
            'Alterações psicofisiológicas; Sobrecarga e fadiga física e cognitiva; Perda de Produtividade e Redução da Percepção de Riscos.',
            'Sobrecarga e fadiga física e cognitiva; Redução da Percepção de Risco.',
            'Alterações psicofisiológicas e/ou sociais.',
            'Fadiga cognitiva; Sonolência; Morosidade e Redução da Percepção de Riscos.',
            'Alterações psicofisiológicas e/ou sociais.',
            'Desconhecimento dos riscos aos quais se expõe e consequente redução da percepção de riscos.',
            'Sobrecarga e fadiga física e cognitiva; Redução da Percepção de Risco.',
            'Sobrecarga e fadiga física e cognitiva; Redução da Percepção de Risco.',
            'Fadiga física e cognitiva.',
            'Alterações psicofisiológicas; Sobrecarga e fadiga física e cognitiva; Perda de Produtividade e Redução da Percepção de Riscos.',
            'Irritabilidade, estresse, dores de cabeça, perda de foco no trabalho e redução da produtividade.',
            'Irritabilidade, estresse, dores de cabeça, perda de foco no trabalho e redução da produtividade.',
            'Estresse, desconforto térmico, irritabilidade, dores de cabeça, perda foco no trabalho e redução da produtividade.',
            'Cansaço, estresse, dor de cabeça, alergias, ressecamento da pele, crise de asma, infecções virais ou bacterianas.',
            'Fadiga visual e cognitiva; Desconforto e Redução da Percepção de Riscos.',
            'Fadiga visual e cognitiva; Desconforto; Perda de desempenho e Redução da Percepção de Riscos.',
            'Fadiga muscular; Perda de desempenho; Escoriação; Ferimento; Luxação; Torção.',
            'Alterações psicofisiológicas e sociais; Fadiga cognitiva; Perda de desempenho; Redução da percepção de risco.',
            'Alterações psicofisiológicas, Fadiga cognitiva, Perda de desempenho e Redução da percepção de risco.',
            'Alterações psicofisiológicas, Fadiga cognitiva, Perda de desempenho e Redução da percepção de risco.',
            'Fadiga cognitiva e perda de desempenho.',
            'Alterações psicofisiológicas e sociais; Fadiga cognitiva.',
            'Alterações psicofisiológicas e sociais; Fadiga cognitiva.',
            'Alterações psicofisiológicas e sociais; Fadiga cognitiva; Perda de desempenho; Redução da percepção de risco.',
            'Alterações psicofisiológicas; Desconforto, Fadiga cognitiva, Perda de desempenho e Redução da percepção de risco.',
            'Alterações psicofisiológicas; Desconforto, Fadiga muscular e cognitiva, Perda de desempenho e Redução da percepção de risco.',
            'Alterações psicofisiológicas e sociais; Fadiga cognitiva; Irritabilidade; Perda de desempenho; Redução da percepção de risco.',
            'Alterações psicofisiológicas; Desconforto, Fadiga cognitiva e Perda de desempenho.'
        ]
    },
    'acidente': {
        'riscos': [
            'Absorção (por contato) de substância cáustica, tóxica ou nociva.', 'Afogamento, imersão, engolfamento.',
            'Aprisionamento em, sob ou entre', 'Aprisionamento em, sob ou entre um objeto parado e outro em movimento.',
            'Aprisionamento em, sob ou entre objetos em movimento convergente.',
            'Aprisionamento em, sob ou entre dois ou mais objetos em movimento (sem encaixe).',
            'Aprisionamento em, sob ou entre um objeto parado e outro em movimento.',
            'Aprisionamento em, sob ou entre desabamento ou desmoronamento de edificação, estrutura, barreira, etc.',
            'Arestas cortantes, superfícies com rebarbas, farpas ou elementos de fixação espostos',
            'Ataque de ser vivo por mordedura, picada, chifrada, coice, etc.', 'Ataque de ser vivo com peçonha',
            'Ataque de ser vivo com transmissão de doença', 'Ataque de ser vivo (inclusive humano)',
            'Atrito ou abrasão por encostar em objeto', 'Atrito ou abrasão por manusear objeto',
            'Atrito ou abrasão por corpo estranho no olho', 'Atrito ou abrasão', 'Atropelamento',
            'Batida contra objeto parado ou em movimento', 'Carga Suspensa',
            'Colisão entre veículos e/ou equipamentos autopropelidos',
            'Condições climáticas adversas (sol, chuva, vento, etc.)',
            'Contato com objeto ou substância em movimento',
            'Contato com objeto ou substância a temperatura muito alta',
            'Contato com objeto ou substância a temperatura muito baixa',
            'Desabamento/Desmoronamento de edificação, estrutura e/ou materiais diversos.',
            'Elementos Móveis e/ou Rotativos', 'Emergências na circunvizinhança',
            'Equipamento pressurizado hidráulico ou pressurizado.', 'Exposição à Energia Elétrica',
            'Ferramentas manuais', 'Ferramentas elétricas', 'Gases/vapores/poeiras (tóxicos ou não tóxicos)',
            'Gases/vapores/poeiras inflamáveis', 'Impacto de pessoa contra objeto parado',
            'Impacto de pessoa contra objeto em movimento', 'Impacto sofrido por pessoa.',
            'Impacto sofrido por pessoa, de objeto em movimento', 'Impacto sofrido poe pessoa, de objeto que cai',
            'Impacto sofrido poe pessoa, de objeto projetado', 'Inalação de substância tóxica/nociva.',
            'Ingestão de substância cáustica, tóxica ou nociva.', 'Inalação, ingestão e/ou absorção.',
            'Incêndio/Explosão', 'Objetos cortantes/perfurocortantes',
            'Pessoas não autorizadas e/ou visitantes no local de trabalho',
            'Portas, escotilhas, tampas, "bocas de visita", flanges',
            'Projeção de Partículas sólidas e/ou líquidas',
            'Queda de pessoa com diferença de nível de andaime, passarela, plataforma, etc.',
            'Queda de pessoa com diferença de nível de escada (móvel ou fixa).',
            'Queda de pessoa com diferença de nível de material empilhado.',
            'Queda de pessoa com diferença de nível de veículo.',
            'Queda de pessoa com diferença de nível em poço, escavação, abertura no piso, etc.',
            'Queda de pessoa com diferença de nível ≤ 2m', 'Queda de pessoa com diferença de nível > 2m',
            'Queda de pessoa em mesmo nível', 'Reação do corpo a seus movimentos (escorregão sem queda, etc.)',
            'Vidro (recipientes, portas, bancadas, janelas, objetos diversos).', 'Soterramento',
            'Substâncias tóxicas e/ou inflamáveis', 'Superfícies, substâncias e/ou objetos aquecidos',
            'Superfícies, substâncias e/ou objetos em baixa temperatura',
            'Tombamento, quebra e/ou ruptura de estrutura (fixa ou móvel)', 'Tombamento de máquina/equipamento',
            'Trabalho à céu aberto', 'Trabalho em espaços confinados',
            'Trabalho com máquinas portáteis rotativas.', 'Trabalho com máquinas e/ou equipamentos'
        ],
        'danos': [
            'Intoxicação, envenenamento, queimadura, irritação ou reação alérgica.',
            'Asfixia, desconforto respiratório, nível de consciência alterado, letargia, palidez, pele azulada, tosse, transtorno neurológico.',
            'Compressão/esmagamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Compressão/esmagamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Compressão/esmagamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Compressão/esmagamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Compressão/esmagamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Compressão/esmagamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Corte, laceração, ferida contusa, punctura (ferida aberta).',
            'Perfurações, cortes, arranhões, escoriações, fraturas.',
            'Dor, inchaço, manchas arroxeadas, sangramento, hemorragia em regiões vitais, infecção, necrose, insuficiência renal.',
            'Arranhões, lacerações, infecções bacterianas, raiva, entre outros tipos de doenças.',
            'Ferimentos de diversos tipos, incluindo com uso de armas, cortes, perfurações, luxações, escoriações, fraturas.',
            'Cortes, ferimentos, esfoladura, escoriações, raspagem superficial da pele, mucosas, etc.',
            'Cortes, ferimentos, esfoladura, escoriações, raspagem superficial da pele, mucosas, etc.',
            'Raspagem superficial das córneas.',
            'Cortes, ferimentos, esfoladura, escoriações, raspagem superficial da pele, mucosas, etc.',
            'Compressão/esmagamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Cortes, escoriações, luxações, fraturas, amputações.',
            'Esmagamento, prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Compressão/esmagamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Intermação, insolação, câibra, exaustão, desidratação, resfriados.',
            'Cortes, escoriações, luxações, fraturas, amputações.',
            'Queimadura ou escaldadura.',
            'Congelamento, geladura e outros efeitos da exposição à baixa temperatura.',
            'Compressão e/ou esmagamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Escoriação, ferimento, corte, luxação, fratura, amputação.',
            'Danos materiais, danos pessoais (queimaduras, contusões, asfixia, aprosionamento, fraturas, etc.).',
            'Ferimentos, rompimento do tímpano, deslocamento de retina ocular, projeção de partículas sólidas e liquidas, queimaduras, choque elétrico.',
            'Choque elétrico e eletroplessão (eletrocussão).',
            'Cortes, ferimentos, escoriações.',
            'Cortes, ferimentos, escoriações, choque elétrico.',
            'Irritação os olhos e/ou da pele, dermatites, doenças respiratórias, intoxicação.',
            'Asfixia, queimaduras, morte por explosão.',
            'Cortes, escoriações, luxações, fraturas, amputações.',
            'Cortes, escoriações, luxações, fraturas, amputações.',
            'Cortes, escoriações, luxações, fraturas, amputações.',
            'Cortes, escoriações, luxações, fraturas, amputações.',
            'Esmagamento, prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Escoriação, ferimento, perfuração, corte, luxação, fratura, prensamento.',
            'Intoxicação, envenenamento, queimadura, irritação ou reação alérgica.',
            'Intoxicação, envenenamento, queimadura, irritação ou reação alérgica.',
            'Intoxicação, envenenamento, queimadura, irritação ou reação alérgica.',
            'Queimadura de 1º, 2º ou 3º grau, asfixia,  arremessos, cortes, escoriações, luxações, fraturas.',
            'Corte, laceração, ferida contusa, punctura (ferida aberta), perfuração.',
            'Escoriação, ferimento, corte, luxação, fratura, entre outros danos devido às características do local e atividades realizadas.',
            'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações, exposição à gases tóxicos.',
            'Ferimento, corte, queimadura, perfuração, intoxicação.',
            'Escoriações, ferimentos, cortes, luxações, fraturas, morte.',
            'Escoriações, ferimentos, cortes, luxações, fraturas, morte.',
            'Escoriações, ferimentos, cortes, luxações, fraturas, morte.',
            'Escoriações, ferimentos, cortes, luxações, fraturas, morte.',
            'Escoriações, ferimentos, cortes, luxações, fraturas, morte.',
            'Escoriações, ferimentos, cortes, luxações, fraturas, morte.',
            'Escoriações, ferimentos, cortes, luxações, fraturas, morte.',
            'Escoriações, ferimentos, cortes, luxações, fraturas.',
            'Torções, distensões, rupturas ou outras lesões musculares internas.',
            'Corte, ferimento, perfuração.',
            'Asfixia, desconforto respiratório, nível de consciência alterado, letargia, palidez, pele azulada, tosse, transtorno neurológico.',
            'Intoxicação, asfixia, queimaduras de  1º, 2º ou 3º grau.',
            'Queimadura de 1º, 2º ou 3º grau.',
            'Queimadura de 1º, 2º ou 3º grau.',
            'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Intermação, insolação, câibra, exaustão, desidratação, resfriados.',
            'Asfixia, hiperóxia, contaminação por poeiras e/ou gases tóxicos, queimadura de 1º, 2º ou 3º grau, arremessos, cortes, escoriações, luxações, fraturas.',
            'Cortes, ferimentos, escoriações, amputações.',
            'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações, choque elétrico.'
        ]
    },
}

DATA_MANAGER_API_VERSION = 2

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

@st.cache_resource
def init_managers(api_version=DATA_MANAGER_API_VERSION):
    db_manager = DatabaseManager()
    auth_manager = AuthManager(db_manager)
    user_data_manager = UserDataManager(db_manager)

    # Compatibilidade defensiva para objetos legados em cache
    if not hasattr(user_data_manager, 'get_user_docx_templates'):
        user_data_manager.get_user_docx_templates = lambda _user_id: []
    if not hasattr(user_data_manager, 'add_docx_template'):
        user_data_manager.add_docx_template = lambda *_args, **_kwargs: (False, 'Funcionalidade indisponível nesta sessão', None)
    if not hasattr(user_data_manager, 'remove_docx_template'):
        user_data_manager.remove_docx_template = lambda *_args, **_kwargs: (False, 'Funcionalidade indisponível nesta sessão')

    return db_manager, auth_manager, user_data_manager

db_manager, auth_manager, user_data_manager = init_managers(DATA_MANAGER_API_VERSION)

st.markdown("""
<style>
    [data-testid="stSidebar"] {display: none;}
    .main-header {text-align: center; padding-bottom: 20px;}
    .auth-container {max-width: 400px; margin: 0 auto; padding: 2rem; border: 1px solid #ddd; border-radius: 10px; background-color: #f9f9f9;}
    .user-info {background-color: #262730; color: white; padding: 1rem; border-radius: 5px; margin-bottom: 1rem; border: 1px solid #3DD56D;}
</style>
""", unsafe_allow_html=True)

def show_login_page():
    st.markdown("""<div class="main-header"><h1>🔐 Acesso ao Sistema</h1><p>Faça login ou registre-se para acessar o Gerador de OS</p></div>""", unsafe_allow_html=True)
    tab1, tab2 = st.tabs(["Login", "Registro"])
    with tab1:
        with st.form("login_form"):
            email = st.text_input("Email", placeholder="seu@email.com")
            password = st.text_input("Senha", type="password")
            if st.form_submit_button("Entrar", use_container_width=True):
                if email and password:
                    success, message, session_data = auth_manager.login_user(email, password)
                    if success:
                        st.session_state.authenticated = True
                        st.session_state.user_data = session_data
                        st.session_state.user_data_loaded = False 
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)
                else:
                    st.error("Por favor, preencha todos os campos")
    with tab2:
        with st.form("register_form"):
            reg_email = st.text_input("Email", placeholder="seu@email.com", key="reg_email")
            reg_password = st.text_input("Senha", type="password", key="reg_password")
            reg_password_confirm = st.text_input("Confirmar Senha", type="password")
            if st.form_submit_button("Registrar", use_container_width=True):
                if reg_email and reg_password and reg_password_confirm:
                    if reg_password != reg_password_confirm:
                        st.error("As senhas não coincidem")
                    else:
                        success, message = auth_manager.register_user(reg_email, reg_password)
                        if success:
                            st.success(message)
                            st.info("Agora você pode fazer login com suas credenciais")
                        else:
                            st.error(message)
                else:
                    st.error("Por favor, preencha todos os campos")

def check_authentication():
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'user_data' not in st.session_state:
        st.session_state.user_data = None
    if st.session_state.authenticated and st.session_state.user_data:
        session_token = st.session_state.user_data.get('session_token')
        if session_token:
            is_valid, _ = auth_manager.validate_session(session_token)
            if not is_valid:
                st.session_state.authenticated = False
                st.session_state.user_data = None
                st.rerun()

def logout_user():
    if st.session_state.user_data and st.session_state.user_data.get('session_token'):
        auth_manager.logout_user(st.session_state.user_data['session_token'])
    st.session_state.authenticated = False
    st.session_state.user_data = None
    st.session_state.user_data_loaded = False
    st.rerun()

def show_user_info():
    if st.session_state.get('authenticated'):
        user_email = st.session_state.user_data.get('email', 'N/A')
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown(f'<div class="user-info">👤 <strong>Usuário:</strong> {user_email}</div>', unsafe_allow_html=True)
        with col2:
            if st.button("Sair", type="secondary"):
                logout_user()

def carregar_modelos_docx_usuario(user_id):
    """Carrega modelos DOCX com fallback para ambientes com cache legado."""
    metodo = getattr(user_data_manager, 'get_user_docx_templates', None)
    if not callable(metodo):
        return []
    try:
        return metodo(user_id)
    except Exception:
        return []

def init_user_session_state():
    if st.session_state.get('authenticated') and not st.session_state.get('user_data_loaded'):
        user_id = st.session_state.user_data.get('user_id')
        if user_id:
            st.session_state.medicoes_adicionadas = user_data_manager.get_user_measurements(user_id)
            st.session_state.epis_adicionados = user_data_manager.get_user_epis(user_id)
            st.session_state.riscos_manuais_adicionados = user_data_manager.get_user_manual_risks(user_id)
            st.session_state.modelos_docx_adicionados = carregar_modelos_docx_usuario(user_id)
            st.session_state.user_data_loaded = True
    if 'medicoes_adicionadas' not in st.session_state:
        st.session_state.medicoes_adicionadas = []
    if 'epis_adicionados' not in st.session_state:
        st.session_state.epis_adicionados = []
    if 'riscos_manuais_adicionados' not in st.session_state:
        st.session_state.riscos_manuais_adicionados = []
    if 'modelos_docx_adicionados' not in st.session_state:
        st.session_state.modelos_docx_adicionados = []
    if 'cargos_concluidos' not in st.session_state:
        st.session_state.cargos_concluidos = set()
    if 'os_setor_cargo_manuais' not in st.session_state:
        st.session_state.os_setor_cargo_manuais = []

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
        'descricao_de_atividades': ['descricaodeatividades', 'atividades', 'descricaoatividades', 'descriçãodeatividades', 'tarefas', 'descricaodasTarefas'],
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

@st.cache_data
def obter_dados_pgr():
    data = [
        {'categoria': 'fisico', 'risco': 'Ruído (Contínuo ou Intermitente)', 'possiveis_danos': 'Perda auditiva, zumbido, estresse, irritabilidade.'},
        {'categoria': 'fisico', 'risco': 'Ruído (Impacto)', 'possiveis_danos': 'Perda auditiva, trauma acústico.'},
        {'categoria': 'fisico', 'risco': 'Vibração de Corpo Inteiro', 'possiveis_danos': 'Problemas na coluna, dores lombares.'},
        {'categoria': 'fisico', 'risco': 'Vibração de Mãos e Braços', 'possiveis_danos': 'Doenças osteomusculares, problemas circulatórios.'},
        {'categoria': 'fisico', 'risco': 'Calor', 'possiveis_danos': 'Desidratação, insolação, câibras, exaustão, intermação.'},
        {'categoria': 'fisico', 'risco': 'Frio', 'possiveis_danos': 'Hipotermia, congelamento, doenças respiratórias.'},
        {'categoria': 'fisico', 'risco': 'Radiações Ionizantes', 'possiveis_danos': 'Câncer, mutações genéticas, queimaduras.'},
        {'categoria': 'fisico', 'risco': 'Radiações Não-Ionizantes', 'possiveis_danos': 'Queimaduras, lesões oculares, câncer de pele.'},
        {'categoria': 'fisico', 'risco': 'Pressões Anormais', 'possiveis_danos': 'Doença descompressiva, barotrauma.'},
        {'categoria': 'fisico', 'risco': 'Umidade', 'possiveis_danos': 'Doenças respiratórias, dermatites, micoses.'},
        {'categoria': 'quimico', 'risco': 'Poeiras', 'possiveis_danos': 'Pneumoconioses (silicose, asbestose), irritação respiratória.'},
        {'categoria': 'quimico', 'risco': 'Fumos', 'possiveis_danos': 'Doenças respiratórias (febre dos fumos metálicos), intoxicações.'},
        {'categoria': 'quimico', 'risco': 'Névoas', 'possiveis_danos': 'Irritação respiratória, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Gases', 'possiveis_danos': 'Asfixia, intoxicações, irritação respiratória.'},
        {'categoria': 'quimico', 'risco': 'Vapores', 'possiveis_danos': 'Irritação respiratória, intoxicações, dermatites.'},
        {'categoria': 'quimico', 'risco': 'Produtos Químicos em Geral', 'possiveis_danos': 'Queimaduras, irritações, intoxicações, dermatites, câncer.'},
        {'categoria': 'biologico', 'risco': 'Bactérias', 'possiveis_danos': 'Infecções, doenças infecciosas (tétano, tuberculose).'},
        {'categoria': 'biologico', 'risco': 'Fungos', 'possiveis_danos': 'Micoses, alergias, infecções respiratórias.'},
        {'categoria': 'biologico', 'risco': 'Vírus', 'possiveis_danos': 'Doenças virais (hepatite, HIV), infecções.'},
        {'categoria': 'ergonomico', 'risco': 'Levantamento e Transporte Manual de Peso', 'possiveis_danos': 'Lesões musculoesqueléticas, dores na coluna.'},
        {'categoria': 'ergonomico', 'risco': 'Posturas Inadequadas', 'possiveis_danos': 'Dores musculares, lesões na coluna, LER/DORT.'},
        {'categoria': 'ergonomico', 'risco': 'Repetitividade', 'possiveis_danos': 'LER/DORT, tendinites, síndrome do túnel do carpo.'},
        {'categoria': 'acidente', 'risco': 'Máquinas e Equipamentos sem Proteção', 'possiveis_danos': 'Amputações, cortes, esmagamentos, prensamentos.'},
        {'categoria': 'acidente', 'risco': 'Eletricidade', 'possiveis_danos': 'Choque elétrico, queimaduras, fibrilação ventricular.'},
        {'categoria': 'acidente', 'risco': 'Trabalho em Altura', 'possiveis_danos': 'Quedas, fraturas, morte.'},
        {'categoria': 'acidente', 'risco': 'Projeção de Partículas', 'possiveis_danos': 'Lesões oculares, cortes na pele.'}
    ]
    return pd.DataFrame(data)

def substituir_placeholders(doc, contexto):
    def aplicar_formatacao_padrao(run):
        run.font.name = 'Segoe UI'
        run.font.size = Pt(9)
        return run

    def processar_paragrafo(p):
        texto_original_paragrafo = p.text
        if "[MEDIÇÕES]" in texto_original_paragrafo:
            for run in p.runs:
                run.text = ''
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            medicoes_valor = contexto.get("[MEDIÇÕES]", "Não aplicável")
            if medicoes_valor == "Não aplicável" or not medicoes_valor.strip():
                run = aplicar_formatacao_padrao(p.add_run("Não aplicável"))
                run.font.bold = False
            else:
                linhas = medicoes_valor.split('\n')
                for i, linha in enumerate(linhas):
                    if not linha.strip(): continue
                    if i > 0: p.add_run().add_break()
                    if ":" in linha:
                        partes = linha.split(":", 1)
                        agente_texto = partes[0].strip() + ":"
                        valor_texto = partes[1].strip()
                        run_agente = aplicar_formatacao_padrao(p.add_run(agente_texto + " "))
                        run_agente.font.bold = True
                        run_valor = aplicar_formatacao_padrao(p.add_run(valor_texto))
                        run_valor.font.bold = False
                    else:
                        run_simples = aplicar_formatacao_padrao(p.add_run(linha))
                        run_simples.font.bold = False
            return
        placeholders_no_paragrafo = [key for key in contexto if key in texto_original_paragrafo]
        if not placeholders_no_paragrafo:
            return
        estilo_rotulo = {
            'bold': p.runs[0].bold if p.runs else False,
            'italic': p.runs[0].italic if p.runs else False,
            'underline': p.runs[0].underline if p.runs else False,
        }
        texto_final = texto_original_paragrafo
        for key in placeholders_no_paragrafo:
            texto_final = texto_final.replace(key, str(contexto[key]))
        p.clear()
        texto_restante = texto_final
        for i, key in enumerate(placeholders_no_paragrafo):
            valor_placeholder = str(contexto[key])
            partes = texto_restante.split(valor_placeholder, 1)
            if partes[0]:
                run_rotulo = aplicar_formatacao_padrao(p.add_run(partes[0]))
                run_rotulo.font.bold = estilo_rotulo['bold']
                run_rotulo.font.italic = estilo_rotulo['italic']
                run_rotulo.underline = estilo_rotulo['underline']
            run_valor = aplicar_formatacao_padrao(p.add_run(valor_placeholder))
            run_valor.font.bold = False
            run_valor.font.italic = False
            run_valor.font.underline = False
            texto_restante = partes[1]
        if texto_restante:
            run_final = aplicar_formatacao_padrao(p.add_run(texto_restante))
            run_final.font.bold = estilo_rotulo['bold']
            run_final.font.italic = estilo_rotulo['italic']
            run_final.underline = estilo_rotulo['underline']

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    processar_paragrafo(p)
    for p in doc.paragraphs:
        processar_paragrafo(p)

def gerar_os(funcionario, df_pgr, riscos_selecionados, epis_manuais, medicoes_manuais, riscos_manuais, modelo_doc_bytes):
    doc = Document(BytesIO(modelo_doc_bytes))
    riscos_info = df_pgr[df_pgr['risco'].isin(riscos_selecionados)]
    riscos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}
    danos_por_categoria = {cat: [] for cat in CATEGORIAS_RISCO.keys()}

    for _, risco_row in riscos_info.iterrows():
        categoria = str(risco_row.get("categoria", "")).lower()
        if categoria in riscos_por_categoria:
            riscos_por_categoria[categoria].append(str(risco_row.get("risco", "")))
            danos = risco_row.get("possiveis_danos")
            if pd.notna(danos): 
                danos_por_categoria[categoria].append(str(danos))

    if riscos_manuais:
        map_categorias_rev = {v: k for k, v in CATEGORIAS_RISCO.items()}
        for risco_manual in riscos_manuais:
            categoria_display = risco_manual.get('category')
            categoria_alvo = map_categorias_rev.get(categoria_display)
            if categoria_alvo:
                riscos_por_categoria[categoria_alvo].append(risco_manual.get('risk_name', ''))
                if risco_manual.get('possible_damages'):
                    danos_por_categoria[categoria_alvo].append(risco_manual.get('possible_damages'))

    for cat in danos_por_categoria:
        danos_por_categoria[cat] = sorted(list(set(danos_por_categoria[cat])))

    medicoes_formatadas = []
    for med in medicoes_manuais:
        agente = str(med.get('agent', '')).strip()
        valor = str(med.get('value', '')).strip()
        unidade = str(med.get('unit', '')).strip()
        epi_associado = str(med.get('epi_name', med.get('epi', ''))).strip()
        if agente and agente not in ['', 'N/A', 'nan', 'None'] and valor and valor not in ['', 'N/A', 'nan', 'None']:
            linha = f"{agente}: {valor}"
            if unidade and unidade not in ['', 'N/A', 'nan', 'None']:
                linha += f" {unidade}"
            if epi_associado and epi_associado not in ['', 'N/A', 'nan', 'None']:
                linha += f" | EPI: {epi_associado}"
            medicoes_formatadas.append(linha)
    medicoes_texto = "\n".join(medicoes_formatadas) if medicoes_formatadas else "Não aplicável"

    data_admissao = "Não informado"
    if 'data_de_admissao' in funcionario and pd.notna(funcionario['data_de_admissao']):
        try: 
            data_admissao = pd.to_datetime(funcionario['data_de_admissao']).strftime('%d/%m/%Y')
        except Exception: 
            data_admissao = str(funcionario['data_de_admissao'])
    elif 'Data de Admissão' in funcionario and pd.notna(funcionario['Data de Admissão']):
        try: 
            data_admissao = pd.to_datetime(funcionario['Data de Admissão']).strftime('%d/%m/%Y')
        except Exception: 
            data_admissao = str(funcionario['Data de Admissão'])

    descricao_atividades = "Não informado"
    if 'descricao_de_atividades' in funcionario and pd.notna(funcionario['descricao_de_atividades']):
        descricao_atividades = str(funcionario['descricao_de_atividades']).strip()
    elif 'Descrição de Atividades' in funcionario and pd.notna(funcionario['Descrição de Atividades']):
        descricao_atividades = str(funcionario['Descrição de Atividades']).strip()
    if descricao_atividades == "Não informado" or descricao_atividades == "" or descricao_atividades == "nan":
        funcao = str(funcionario.get('funcao', funcionario.get('Função', 'N/A')))
        setor = str(funcionario.get('setor', funcionario.get('Setor', 'N/A')))
        if funcao != 'N/A' and setor != 'N/A':
            descricao_atividades = f"Atividades relacionadas à função de {funcao} no setor {setor}, incluindo todas as tarefas operacionais, administrativas e de apoio inerentes ao cargo."
        else:
            descricao_atividades = "Atividades operacionais, administrativas e de apoio conforme definido pela chefia imediata."

    def tratar_lista_vazia(lista, separador=", "):
        if not lista or all(not item.strip() for item in lista): 
            return "Não identificado"
        return separador.join(sorted(list(set(item for item in lista if item and item.strip()))))

    contexto = {
        "[NOME EMPRESA]": str(funcionario.get("empresa", funcionario.get("Empresa", "N/A"))), 
        "[UNIDADE]": str(funcionario.get("unidade", funcionario.get("Unidade", "N/A"))),
        "[NOME FUNCIONÁRIO]": str(funcionario.get("nome_do_funcionario", funcionario.get("Nome", "N/A"))), 
        "[DATA DE ADMISSÃO]": data_admissao,
        "[SETOR]": str(funcionario.get("setor", funcionario.get("Setor", "N/A"))), 
        "[FUNÇÃO]": str(funcionario.get("funcao", funcionario.get("Função", "N/A"))),
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
        "[EPIS]": tratar_lista_vazia([epi['epi_name'] for epi in epis_manuais]),
        "[MEDIÇÕES]": medicoes_texto,
    }
    substituir_placeholders(doc, contexto)
    return doc

def main():
    check_authentication()
    init_user_session_state()
    if not st.session_state.get('authenticated'):
        show_login_page()
        return
    user_id = st.session_state.user_data['user_id']
    show_user_info()
    st.markdown("""<div class="main-header"><h1>📄 Gerador de Ordens de Serviço (OS)</h1><p>Gere OS em lote a partir de um modelo Word (.docx), com planilha de funcionários ou por setor/cargo.</p></div>""", unsafe_allow_html=True)

    modo_geracao = st.radio(
        "Como deseja gerar as OS?",
        options=["Com planilha de funcionários", "Sem planilha (por setor e cargo)"],
        horizontal=True
    )

    with st.container(border=True):
        st.markdown("##### 📂 1. Carregue os Documentos")
        col1, col2 = st.columns(2)
        with col1:
            arquivo_funcionarios = st.file_uploader("📄 **Planilha de Funcionários (.xlsx)**", type="xlsx")
            if modo_geracao == "Sem planilha (por setor e cargo)":
                st.caption("Opcional neste modo.")
        with col2:
            arquivo_modelo_os_upload = st.file_uploader("📝 **Upload de Modelo de OS (.docx)**", type="docx")

            with st.expander("💾 Modelos DOCX cadastrados"):
                with st.form("form_modelo_docx", clear_on_submit=True):
                    nome_modelo_docx = st.text_input("Nome do modelo")
                    arquivo_modelo_salvar = st.file_uploader("Arquivo .docx para cadastrar", type="docx", key="upload_modelo_salvar")
                    if st.form_submit_button("Salvar modelo"):
                        if nome_modelo_docx and arquivo_modelo_salvar:
                            add_template = getattr(user_data_manager, 'add_docx_template', None)
                            if not callable(add_template):
                                st.warning('Funcionalidade de modelos DOCX indisponível nesta sessão. Recarregue o app.')
                                sucesso, mensagem, _ = False, '', None
                            else:
                                sucesso, mensagem, _ = add_template(
                                user_id,
                                nome_modelo_docx,
                                arquivo_modelo_salvar.getvalue()
                            )
                            if sucesso:
                                st.success(mensagem)
                                st.session_state.user_data_loaded = False
                                st.rerun()
                            else:
                                st.warning(mensagem)
                        else:
                            st.warning("Informe o nome e selecione um arquivo .docx.")

                if st.session_state.modelos_docx_adicionados:
                    st.write("**Modelos salvos:**")
                    for modelo in st.session_state.modelos_docx_adicionados:
                        col_modelo, col_btn = st.columns([4, 1])
                        col_modelo.markdown(f"- {modelo['template_name']}")
                        if col_btn.button("Remover", key=f"rem_modelo_{modelo['id']}"):
                            remove_template = getattr(user_data_manager, 'remove_docx_template', None)
                            if not callable(remove_template):
                                st.warning('Funcionalidade de modelos DOCX indisponível nesta sessão. Recarregue o app.')
                            else:
                                remove_template(user_id, modelo['id'])
                            st.session_state.user_data_loaded = False
                            st.rerun()
                else:
                    st.caption("Nenhum modelo DOCX salvo ainda.")

            opcoes_modelo_salvo = [m['template_name'] for m in st.session_state.modelos_docx_adicionados]
            modelo_salvo_nome = st.selectbox(
                "Ou selecione um modelo cadastrado",
                options=["-- Nenhum --"] + opcoes_modelo_salvo,
                index=0
            )

    arquivo_modelo_os = arquivo_modelo_os_upload
    if not arquivo_modelo_os and modelo_salvo_nome != "-- Nenhum --":
        modelo_selecionado = next((m for m in st.session_state.modelos_docx_adicionados if m['template_name'] == modelo_salvo_nome), None)
        if modelo_selecionado:
            arquivo_modelo_os = BytesIO(modelo_selecionado['file_content'])

    if not arquivo_modelo_os:
        st.info("📋 Faça upload de um Modelo de OS ou selecione um modelo cadastrado para continuar.")
        return

    if hasattr(arquivo_modelo_os, "getvalue"):
        modelo_doc_bytes = arquivo_modelo_os.getvalue()
    else:
        modelo_doc_bytes = arquivo_modelo_os.read()

    df_funcionarios = pd.DataFrame()
    if modo_geracao == "Com planilha de funcionários":
        if not arquivo_funcionarios:
            st.info("📋 Neste modo, carregue a Planilha de Funcionários para continuar.")
            return

        df_funcionarios_raw = carregar_planilha(arquivo_funcionarios)
        if df_funcionarios_raw is None:
            st.stop()

        df_funcionarios = mapear_e_renomear_colunas_funcionarios(df_funcionarios_raw)
    else:
        with st.container(border=True):
            st.markdown('##### 👥 2. Monte a Lista de OS por Setor e Cargo')
            st.caption("Adicione os setores/cargos para gerar uma OS por combinação.")
            with st.form("form_setor_cargo", clear_on_submit=True):
                col_setor, col_funcao, col_empresa, col_unidade = st.columns(4)
                with col_setor:
                    setor_manual = st.text_input("Setor")
                with col_funcao:
                    funcao_manual = st.text_input("Cargo/Função")
                with col_empresa:
                    empresa_manual = st.text_input("Empresa (opcional)")
                with col_unidade:
                    unidade_manual = st.text_input("Unidade (opcional)")
                descricao_manual = st.text_area("Descrição de atividades (opcional)")
                if st.form_submit_button("Adicionar setor/cargo"):
                    if setor_manual.strip() and funcao_manual.strip():
                        st.session_state.os_setor_cargo_manuais.append({
                            'setor': setor_manual.strip(),
                            'funcao': funcao_manual.strip(),
                            'empresa': empresa_manual.strip(),
                            'unidade': unidade_manual.strip(),
                            'descricao_de_atividades': descricao_manual.strip(),
                            'nome_do_funcionario': f"Colaboradores - {setor_manual.strip()} / {funcao_manual.strip()}",
                            'data_de_admissao': pd.NaT,
                        })
                        st.rerun()
                    else:
                        st.warning("Informe ao menos Setor e Cargo/Função.")

            if st.session_state.os_setor_cargo_manuais:
                st.write("**Combinações adicionadas:**")
                for i, combinacao in enumerate(st.session_state.os_setor_cargo_manuais):
                    col_desc, col_btn = st.columns([5, 1])
                    col_desc.markdown(f"- **{combinacao['setor']}** / **{combinacao['funcao']}**")
                    if col_btn.button("Remover", key=f"rem_setor_cargo_{i}"):
                        st.session_state.os_setor_cargo_manuais.pop(i)
                        st.rerun()

            if st.session_state.os_setor_cargo_manuais:
                df_funcionarios = pd.DataFrame(st.session_state.os_setor_cargo_manuais)
    df_pgr = obter_dados_pgr()

    with st.container(border=True):
        st.markdown('##### 👥 3. Selecione os Registros para Geração')
        setores = sorted(df_funcionarios['setor'].dropna().unique().tolist()) if 'setor' in df_funcionarios.columns else []
        setor_sel = st.multiselect("Filtrar por Setor(es)", setores)
        df_filtrado_setor = df_funcionarios[df_funcionarios['setor'].isin(setor_sel)] if setor_sel else df_funcionarios
        tipo_registro = "funcionário(s)" if modo_geracao == "Com planilha de funcionários" else "registro(s)"
        st.caption(f"{len(df_filtrado_setor)} {tipo_registro} no(s) setor(es) selecionado(s).")
        funcoes_disponiveis = sorted(df_filtrado_setor['funcao'].dropna().unique().tolist()) if 'funcao' in df_filtrado_setor.columns else []
        funcoes_formatadas = []
        if setor_sel:
            for funcao in funcoes_disponiveis:
                concluido = all((s, funcao) in st.session_state.cargos_concluidos for s in setor_sel)
                if concluido:
                    funcoes_formatadas.append(f"{funcao} ✅ Concluído")
                else:
                    funcoes_formatadas.append(funcao)
        else:
            funcoes_formatadas = funcoes_disponiveis
        funcao_sel_formatada = st.multiselect("Filtrar por Função/Cargo(s)", funcoes_formatadas)
        funcao_sel = [f.replace(" ✅ Concluído", "") for f in funcao_sel_formatada]
        df_final_filtrado = df_filtrado_setor[df_filtrado_setor['funcao'].isin(funcao_sel)] if funcao_sel else df_filtrado_setor
        st.success(f"**{len(df_final_filtrado)} {tipo_registro} selecionado(s) para gerar OS.**")
        if {'nome_do_funcionario', 'setor', 'funcao'}.issubset(df_final_filtrado.columns):
            st.dataframe(df_final_filtrado[['nome_do_funcionario', 'setor', 'funcao']])
        elif {'setor', 'funcao'}.issubset(df_final_filtrado.columns):
            st.dataframe(df_final_filtrado[['setor', 'funcao']])

    with st.container(border=True):
        st.markdown('##### ⚠️ 4. Configure os Riscos e Medidas de Controle')
        st.info("Configure os riscos que serão aplicados a TODOS os funcionários selecionados.")

        tab_fisico, tab_quimico, tab_biologico, tab_ergonomico, tab_acidente, tab_manual = st.tabs([
            "🔥 Físicos", "⚗️ Químicos", "🦠 Biológicos", "🏃 Ergonômicos", "⚠️ Acidentes", "➕ Manual"
        ])

        riscos_selecionados_pgr = {}

        with tab_fisico:
            if 'fisico' in RISCOS_PGR_DADOS:
                st.write(f"**Riscos Físicos PGR:** {len(RISCOS_PGR_DADOS['fisico']['riscos'])} opções disponíveis")
                riscos_selecionados_pgr['fisico'] = st.multiselect(
                    "Selecione os Riscos Físicos:",
                    options=RISCOS_PGR_DADOS['fisico']['riscos'],
                    key="riscos_pgr_fisico",
                    help="Riscos físicos da planilha PGR"
                )
                if riscos_selecionados_pgr['fisico']:
                    danos = get_danos_por_riscos_pgr('fisico', riscos_selecionados_pgr['fisico'])
                    if danos:
                        st.info(f"**Possíveis Danos:** {danos}")

        with tab_quimico:
            if 'quimico' in RISCOS_PGR_DADOS:
                st.write(f"**Riscos Químicos PGR:** {len(RISCOS_PGR_DADOS['quimico']['riscos'])} opções disponíveis")
                riscos_selecionados_pgr['quimico'] = st.multiselect(
                    "Selecione os Riscos Químicos:",
                    options=RISCOS_PGR_DADOS['quimico']['riscos'],
                    key="riscos_pgr_quimico",
                    help="Riscos químicos da planilha PGR"
                )
                if riscos_selecionados_pgr['quimico']:
                    danos = get_danos_por_riscos_pgr('quimico', riscos_selecionados_pgr['quimico'])
                    if danos:
                        st.info(f"**Possíveis Danos:** {danos}")

        with tab_biologico:
            if 'biologico' in RISCOS_PGR_DADOS:
                st.write(f"**Riscos Biológicos PGR:** {len(RISCOS_PGR_DADOS['biologico']['riscos'])} opções disponíveis")
                riscos_selecionados_pgr['biologico'] = st.multiselect(
                    "Selecione os Riscos Biológicos:",
                    options=RISCOS_PGR_DADOS['biologico']['riscos'],
                    key="riscos_pgr_biologico",
                    help="Riscos biológicos da planilha PGR"
                )
                if riscos_selecionados_pgr['biologico']:
                    danos = get_danos_por_riscos_pgr('biologico', riscos_selecionados_pgr['biologico'])
                    if danos:
                        st.info(f"**Possíveis Danos:** {danos}")

        with tab_ergonomico:
            if 'ergonomico' in RISCOS_PGR_DADOS:
                st.write(f"**Riscos Ergonômicos PGR:** {len(RISCOS_PGR_DADOS['ergonomico']['riscos'])} opções disponíveis")
                riscos_selecionados_pgr['ergonomico'] = st.multiselect(
                    "Selecione os Riscos Ergonômicos:",
                    options=RISCOS_PGR_DADOS['ergonomico']['riscos'],
                    key="riscos_pgr_ergonomico",
                    help="Riscos ergonômicos da planilha PGR"
                )
                if riscos_selecionados_pgr['ergonomico']:
                    danos = get_danos_por_riscos_pgr('ergonomico', riscos_selecionados_pgr['ergonomico'])
                    if danos:
                        st.info(f"**Possíveis Danos:** {danos}")

        with tab_acidente:
            if 'acidente' in RISCOS_PGR_DADOS:
                st.write(f"**Riscos de Acidente PGR:** {len(RISCOS_PGR_DADOS['acidente']['riscos'])} opções disponíveis")
                riscos_selecionados_pgr['acidente'] = st.multiselect(
                    "Selecione os Riscos de Acidente:",
                    options=RISCOS_PGR_DADOS['acidente']['riscos'],
                    key="riscos_pgr_acidente",
                    help="Riscos de acidente da planilha PGR"
                )
                if riscos_selecionados_pgr['acidente']:
                    danos = get_danos_por_riscos_pgr('acidente', riscos_selecionados_pgr['acidente'])
                    if danos:
                        st.info(f"**Possíveis Danos:** {danos}")

        with tab_manual:
            with st.form("form_risco_manual", clear_on_submit=True):
                st.markdown("###### Adicionar um Risco que não está na lista")
                risco_manual_nome = st.text_input("Descrição do Risco")
                categoria_manual = st.selectbox("Categoria do Risco Manual", list(CATEGORIAS_RISCO.values()))
                danos_manuais = st.text_area("Possíveis Danos (Opcional)")
                if st.form_submit_button("Adicionar Risco Manual"):
                    if risco_manual_nome and categoria_manual:
                        user_data_manager.add_manual_risk(user_id, categoria_manual, risco_manual_nome, danos_manuais)
                        st.session_state.user_data_loaded = False
                        st.rerun()
            if st.session_state.riscos_manuais_adicionados:
                st.write("**Riscos manuais salvos:**")
                for r in st.session_state.riscos_manuais_adicionados:
                    col1, col2 = st.columns([4, 1])
                    col1.markdown(f"- **{r['risk_name']}** ({r['category']})")
                    if col2.button("Remover", key=f"rem_risco_{r['id']}"):
                        user_data_manager.remove_manual_risk(user_id, r['id'])
                        st.session_state.user_data_loaded = False
                        st.rerun()

        # CORREÇÃO PRINCIPAL: Consolidar riscos em formato compatível com df_pgr
        riscos_selecionados_para_df_pgr = []
        for categoria_key in riscos_selecionados_pgr:
            if riscos_selecionados_pgr[categoria_key]:
                for risco_nome in riscos_selecionados_pgr[categoria_key]:
                    danos = get_danos_por_riscos_pgr(categoria_key, [risco_nome])
                    riscos_selecionados_para_df_pgr.append({
                        'categoria': categoria_key,
                        'risco': risco_nome,
                        'possiveis_danos': danos
                    })

        total_riscos = len(riscos_selecionados_para_df_pgr) + len(st.session_state.riscos_manuais_adicionados)
        if total_riscos > 0:
            with st.expander(f"📖 Resumo de Riscos Selecionados ({total_riscos} no total)", expanded=True):
                riscos_para_exibir = {cat: [] for cat in CATEGORIAS_RISCO.values()}
                
                for categoria_key, categoria_display in CATEGORIAS_RISCO.items():
                    if categoria_key in riscos_selecionados_pgr and riscos_selecionados_pgr[categoria_key]:
                        riscos_para_exibir[categoria_display].extend(riscos_selecionados_pgr[categoria_key])
                
                for risco_manual in st.session_state.riscos_manuais_adicionados:
                    riscos_para_exibir[risco_manual['category']].append(risco_manual['risk_name'])
                
                for categoria, lista_riscos in riscos_para_exibir.items():
                    if lista_riscos:
                        st.markdown(f"**{categoria}**")
                        for risco in sorted(list(set(lista_riscos))):
                            st.markdown(f"- {risco}")
        
        st.divider()

        col_exp1, col_exp2 = st.columns(2)
        with col_exp1:
            with st.expander("📊 **Adicionar Medições**"):
                with st.form("form_medicao", clear_on_submit=True):
                    opcoes_agente = ["-- Digite um novo agente abaixo --"] + AGENTES_DE_RISCO
                    agente_selecionado = st.selectbox("Selecione um Agente/Fonte da lista...", options=opcoes_agente)
                    agente_manual = st.text_input("...ou digite um novo aqui:")
                    valor = st.text_input("Valor Medido")
                    unidade = st.selectbox("Unidade", UNIDADES_DE_MEDIDA)
                    epi_med = st.text_input("EPI Associado (Opcional)")
                    if st.form_submit_button("Adicionar Medição"):
                        agente_a_salvar = agente_manual.strip() if agente_manual.strip() else agente_selecionado
                        if agente_a_salvar != "-- Digite um novo agente abaixo --" and valor:
                            user_data_manager.add_measurement(user_id, agente_a_salvar, valor, unidade, epi_med)
                            st.session_state.user_data_loaded = False
                            st.rerun()
                        else:
                            st.warning("Por favor, preencha o Agente e o Valor.")
                if st.session_state.medicoes_adicionadas:
                    st.write("**Medições salvas:**")
                    for med in st.session_state.medicoes_adicionadas:
                        col1, col2 = st.columns([4, 1])
                        col1.markdown(f"- {med['agent']}: {med['value']} {med['unit']}")
                        if col2.button("Remover", key=f"rem_med_{med['id']}"):
                            user_data_manager.remove_measurement(user_id, med['id'])
                            st.session_state.user_data_loaded = False
                            st.rerun()
        with col_exp2:
            with st.expander("🦺 **Adicionar EPIs Gerais**"):
                with st.form("form_epi", clear_on_submit=True):
                    epi_nome = st.text_input("Nome do EPI")
                    if st.form_submit_button("Adicionar EPI"):
                        if epi_nome:
                            user_data_manager.add_epi(user_id, epi_nome)
                            st.session_state.user_data_loaded = False
                            st.rerun()
                if st.session_state.epis_adicionados:
                    st.write("**EPIs salvos:**")
                    for epi in st.session_state.epis_adicionados:
                        col1, col2 = st.columns([4, 1])
                        col1.markdown(f"- {epi['epi_name']}")
                        if col2.button("Remover", key=f"rem_epi_{epi['id']}"):
                            user_data_manager.remove_epi(user_id, epi['id'])
                            st.session_state.user_data_loaded = False
                            st.rerun()

    st.divider()
    if st.button("🚀 Gerar OS para Funcionários Selecionados", type="primary", use_container_width=True, disabled=df_final_filtrado.empty):
        with st.spinner(f"Gerando {len(df_final_filtrado)} documentos..."):
            documentos_gerados = []
            combinacoes_processadas = set()
            
            for _, func in df_final_filtrado.iterrows():
                combinacoes_processadas.add((func['setor'], func['funcao']))
                
                # CORREÇÃO: Criar DataFrame consolidado com riscos selecionados
                df_riscos_consolidado = pd.DataFrame(riscos_selecionados_para_df_pgr) if riscos_selecionados_para_df_pgr else pd.DataFrame(columns=['categoria', 'risco', 'possiveis_danos'])
                nomes_riscos = df_riscos_consolidado['risco'].tolist() if not df_riscos_consolidado.empty else []
                
                doc = gerar_os(
                    func, 
                    df_riscos_consolidado,
                    nomes_riscos,
                    st.session_state.epis_adicionados,
                    st.session_state.medicoes_adicionadas, 
                    st.session_state.riscos_manuais_adicionados, 
                    modelo_doc_bytes
                )
                doc_io = BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)
                nome_limpo = re.sub(r'[^\w\s-]', '', func.get("nome_do_funcionario", "Func_Sem_Nome")).strip().replace(" ", "_")
                caminho_no_zip = f"{func.get('setor', 'SemSetor')}/{func.get('funcao', 'SemFuncao')}/OS_{nome_limpo}.docx"
                documentos_gerados.append((caminho_no_zip, doc_io.getvalue()))
            
            st.session_state.cargos_concluidos.update(combinacoes_processadas)
            
            if documentos_gerados:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for nome_arquivo, conteudo_doc in documentos_gerados:
                        zip_file.writestr(nome_arquivo, conteudo_doc)
                nome_arquivo_zip = f"OS_Geradas_{time.strftime('%Y%m%d')}.zip"
                st.success(f"🎉 **{len(documentos_gerados)} Ordens de Serviço geradas!**")
                st.download_button(
                    label="📥 Baixar Todas as OS (.zip)", 
                    data=zip_buffer.getvalue(), 
                    file_name=nome_arquivo_zip, 
                    mime="application/zip",
                    use_container_width=True
                )

if __name__ == "__main__":
    main()
