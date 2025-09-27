import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import datetime

# Dados dos Riscos baseados na planilha PGR
RISCOS_PGR_DADOS = {
    'Químico': {
        'riscos': [
            'Exposição a Produto Químico',
        ],
        'danos': [
            'Irritação/lesão ocular, na pele e mucosas; Dermatites; Queimadura Química; Intoxicação; Náuseas; Vômitos.',
        ]
    },
    'Físico': {
        'riscos': [
            'Ambiente Artificialmente Frio',
            'Exposição ao Ruído',
            'Vibrações Localizadas (mão/braço)',
            'Vibração de Corpo Inteiro (AREN)',
            'Vibração de Corpo Inteiro (VDVR)',
            'Exposição à Radiações Ionizantes',
            'Exposição à Radiações Não-ionizantes',
            'Exposição à Temperatura Ambiente Elevada',
            'Exposição à Temperatura Ambiente Baixa',
            'Pressão Atmosférica Anormal (condições hiperbáricas)',
            'Umidade',
        ],
        'danos': [
            'Estresse, desconforto, dormência, rigidez nas partes com maior intensidade de exposição ao frio, redução da destreza, formigamento, redução da sensibilidade dos dedos e flexibilidade das articulações.',
            'Perda Auditiva Induzida pelo Ruído Ocupacional (PAIRO).',
            'Alterações articulares e vasomotoras.',
            'Alterações no sistema digestivo, sistema musculoesquelético, sistema nervoso, alterações na visão, enjoos, náuseas, palidez.',
            'Alterações no sistema digestivo, sistema musculoesquelético, sistema nervoso, alterações na visão, enjoos, náuseas, palidez.',
            'Dano às células do corpo humano, causando doenças graves, inclusive fatais, como câncer.',
            'Depressão imunológica, fotoenvelhecimento, lesões oculares como ceratoconjuntivite, pterígio e catarata; Doenças graves, inclusives fatais, como câncer.',
            'Desidratação, erupções cutâneas, cãibras, fadiga física, problemas cardiocirculatórios, distúrbios psicológicos.',
            'Estresse, desconforto, dormência, rigidez nas partes com maior intensidade de exposição ao frio, redução da destreza, formigamento, redução da sensibilidade dos dedos e flexibilidade das articulações.',
            'Barotrauma pulmonar, lesão de tecido pulmonar ou pneumotórax, embolia arterial gasosa, barotrauma de ouvido, barotrauma sinusal, barotrauma dental, barotrauma facial, doença descompressiva.',
            'Doenças do aparelho respiratório, quedas, doenças de pele, doenças circulatórias, entre outras.',
        ]
    },
    'Biológico': {
        'riscos': [
            'Água e/ou alimentos contaminados',
            'Contato com Fluido Orgânico (sangue, hemoderivados, secreções, excreções)',
            'Contato com Pessoas Doentes e/ou Material Infectocontagiante',
            'Contaminação pelo Corona Vírus',
            'Exposição à Agentes Microbiológicos (fungos, bactérias, vírus, protozoários, parasitas)',
        ],
        'danos': [
            'Intoxicação, diarreias, infecções intestinais.',
            'Doenças infectocontagiosas.',
            'Doenças infectocontagiosas.',
            'COVID-19, podendo causar gripes, febre, tosse seca, cansaço, dores e desconfortos, dor de garganta, diarreia, perda de paladar ou olfato, dificuldade de respirar ou falta de ar, dor ou pressão no peito, perda de fala ou movimentos.',
            'Doenças infectocontagiosas, dermatites, irritação, desconforto, infecção do sistema respiratório.',
        ]
    },
    'Ergonômico': {
        'riscos': [
            'Posturas incômodas/pouco confortáveis por longos períodos',
            'Postura sentada por longos períodos',
            'Postura em pé por longos períodos',
            'Frequente deslocamento à pé durante à jornada de trabalho',
            'Esforço físico intenso',
            'Levantamento e transporte manual de cargas ou volumes',
            'Frequente ação de empurrar/puxar cargas ou volumes',
            'Frequente execução de movimentos repetitivos',
            'Manuseio de ferramentas e/ou objetos pesados por longos períodos',
            'Uso frequente de força, pressão, preensão, flexão, extensão ou torção dos segmentos corporais',
            'Compressão de partes do corpo por superfícies rígidas ou com quinas vivas',
            'Flexões da coluna vertebral frequentes',
            'Uso frequente de pedais',
            'Uso frequente de alavancas',
            'Elevação frequente de membros superiores',
            'Manuseio ou movimentação de cargas e volumes sem pega ou com "pega pobre"',
            'Exposição à vibração de corpo inteiro',
            'Exposição à vibrações localizadas (mão, braço)',
            'Uso frequente de escadas',
            'Trabalho intensivo com teclado ou outros dispositivos de entrada de dados',
            'Posto de trabalho improvisado/inadequado',
            'Mobiliário sem meios de regulagem de ajustes',
            'Equipamentos e/ou máquinas sem meios de regulagem de ajustes ou sem condições de uso',
            'Posto de trabalho não planejado/adaptado para à posição sentada',
            'Assento inadequado',
            'Encosto do assento inadequado ou ausente',
            'Mobiliário ou equipamento sem espaço para movimentação de segmentos corporais',
            'Necessidade de alcançar objetos, documentos, controles, etc, além das zonas de alcance ideais',
            'Equipamentos/mobiliário não adaptados à antropometria do trabalhador',
            'Trabalho realizado sem pausas pré-definidas para descanso',
            'Necessidade de manter ritmos intensos de trabalho',
            'Trabalho com necessidade de variação de turnos',
            'Monotonia',
            'Trabalho noturno',
            'Insuficiência de capacitação para à execução da tarefa',
            'Trabalho com utilização rigorosa de metas de produção',
            'Trabalho remunerado por produção',
            'Cadência do trabalho imposta por um equipamento',
            'Desequilíbrio entre tempo de trabalho e tempo de repouso',
            'Pressão sonora fora dos parâmetros de conforto',
            'Temperatura efetiva fora dos parâmetros de conforto',
            'Velocidade do ar fora dos parâmetros de conforto',
            'Umidade do ar fora dos parâmetros de conforto',
            'Iluminação inadequada',
            'Reflexos que causem desconforto ou prejudiquem à visão',
            'Piso escorregadio ou irregular',
            'Situações de estresse no local de trabalho',
            'Situações de sobrecarga de trabalho mental',
            'Exigência de concentração, atenção e memória',
            'Trabalho em condições de difícil comunicação',
            'Conflitos hierárquicos no trabalho',
            'Problemas de relacionamento no trabalho',
            'Assédio de qualquer natureza no trabalho',
            'Dificuldades para cumprir ordens e determinações da chefia relacionadas ao trabalho',
            'Realização de múltiplas tarefas com alta demanda mental/cognitiva',
            'Insatisfação no trabalho',
            'Falta de autonomia para a realização de tarefas no trabalho',
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
            'Alterações psicofisiológicas; Desconforto, Fadiga cognitiva e Perda de desempenho.',
        ]
    },
    'Acidente': {
        'riscos': [
            'Absorção (por contato) de substância cáustica, tóxica ou nociva.',
            'Afogamento, imersão, engolfamento.',
            'Aprisionamento em, sob ou entre',
            'Aprisionamento em, sob ou entre um objeto parado e outro em movimento.',
            'Aprisionamento em, sob ou entre objetos em movimento convergente.',
            'Aprisionamento em, sob ou entre dois ou mais objetos em movimento (sem encaixe).',
            'Aprisionamento em, sob ou entre um objeto parado e outro em movimento.',
            'Aprisionamento em, sob ou entre desabamento ou desmoronamento de edificação, estrutura, barreira, etc.',
            'Arestas cortantes, superfícies com rebarbas, farpas ou elementos de fixação espostos',
            'Ataque de ser vivo por mordedura, picada, chifrada, coice, etc.',
            'Ataque de ser vivo com peçonha',
            'Ataque de ser vivo com transmissão de doença',
            'Ataque de ser vivo (inclusive humano)',
            'Atrito ou abrasão por encostar em objeto',
            'Atrito ou abrasão por manusear objeto',
            'Atrito ou abrasão por corpo estranho no olho',
            'Atrito ou abrasão',
            'Atropelamento',
            'Batida contra objeto parado ou em movimento',
            'Carga Suspensa',
            'Colisão entre veículos e/ou equipamentos autopropelidos',
            'Condições climáticas adversas (sol, chuva, vento, etc.)',
            'Contato com objeto ou substância em movimento',
            'Contato com objeto ou substância a temperatura muito alta',
            'Contato com objeto ou substância a temperatura muito baixa',
            'Desabamento/Desmoronamento de edificação, estrutura e/ou materiais diversos.',
            'Elementos Móveis e/ou Rotativos',
            'Emergências na circunvizinhança',
            'Equipamento pressurizado hidráulico ou pressurizado.',
            'Exposição à Energia Elétrica',
            'Ferramentas manuais',
            'Ferramentas elétricas',
            'Gases/vapores/poeiras (tóxicos ou não tóxicos)',
            'Gases/vapores/poeiras inflamáveis',
            'Impacto de pessoa contra objeto parado',
            'Impacto de pessoa contra objeto em movimento',
            'Impacto sofrido por pessoa.',
            'Impacto sofrido por pessoa, de objeto em movimento',
            'Impacto sofrido poe pessoa, de objeto que cai',
            'Impacto sofrido poe pessoa, de objeto projetado',
            'Inalação de substância tóxica/nociva.',
            'Ingestão de substância cáustica, tóxica ou nociva.',
            'Inalação, ingestão e/ou absorção.',
            'Incêndio/Explosão',
            'Objetos cortantes/perfurocortantes',
            'Pessoas não autorizadas e/ou visitantes no local de trabalho',
            'Portas, escotilhas, tampas, "bocas de visita", flanges',
            'Projeção de Partículas sólidas e/ou líquidas',
            'Queda de pessoa com diferença de nível de andaime, passarela, plataforma, etc.',
            'Queda de pessoa com diferença de nível de escada (móvel ou fixa).',
            'Queda de pessoa com diferença de nível de material empilhado.',
            'Queda de pessoa com diferença de nível de veículo.',
            'Queda de pessoa com diferença de nível em poço, escavação, abertura no piso, etc.',
            'Queda de pessoa com diferença de nível ≤ 2m',
            'Queda de pessoa com diferença de nível > 2m',
            'Queda de pessoa em mesmo nível',
            'Reação do corpo a seus movimentos (escorregão sem queda, etc.)',
            'Vidro (recipientes, portas, bancadas, janelas, objetos diversos).',
            'Soterramento',
            'Substâncias tóxicas e/ou inflamáveis',
            'Superfícies, substâncias e/ou objetos aquecidos',
            'Superfícies, substâncias e/ou objetos em baixa temperatura',
            'Tombamento, quebra e/ou ruptura de estrutura (fixa ou móvel)',
            'Tombamento de máquina/equipamento',
            'Trabalho à céu aberto',
            'Trabalho em espaços confinados',
            'Trabalho com máquinas portáteis rotativas.',
            'Trabalho com máquinas e/ou equipamentos',
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
            'Intermação, insolação, cãibra, exaustão, desidratação, resfriados.',
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
            'Torções, distensções, rupturas ou outras lesões musculares internas.',
            'Corte, ferimento, perfuração.',
            'Asfixia, desconforto respiratório, nível de consciência alterado, letargia, palidez, pele azulada, tosse, transtorno neurológico.',
            'Intoxicação, asfixia, queimaduras de  1º, 2º ou 3º grau.',
            'Queimadura de 1º, 2º ou 3º grau.',
            'Queimadura de 1º, 2º ou 3º grau.',
            'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações.',
            'Intermação, insolação, cãibra, exaustão, desidratação, resfriados.',
            'Asfixia, hiperóxia, contaminação por poeiras e/ou gases tóxicos, queimadura de 1º, 2º ou 3º grau, arremessos, cortes, escoriações, luxações, fraturas.',
            'Cortes, ferimentos, escoriações, amputações.',
            'Prensamento ou aprisionamento de partes do corpo, cortes, escoriações, luxações, fraturas, amputações, choque elétrico.',
        ]
    },
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
