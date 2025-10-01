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
            'Sobrecarga corporal e dores nos
