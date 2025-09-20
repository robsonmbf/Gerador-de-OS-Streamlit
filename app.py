import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import time
import re
import hashlib
import datetime

# Configuração da página
st.set_page_config(page_title="Gerador de Ordens de Serviço (OS)", page_icon="📄", layout="wide", initial_sidebar_state="expanded")

UNIDADES_DE_MEDIDA = ["dB(A)", "m/s²", "ppm", "mg/m³", "%", "°C", "lx", "cal/cm²", "µT", "kV/m", "W/m²", "f/cm³", "Não aplicável"]

# Definição dos riscos (igual sua base expandida, resumida aqui)
RISCOS_FISICO = sorted([...])
RISCOS_QUIMICO = sorted([...])
RISCOS_BIOLOGICO = sorted([...])
RISCOS_ERGONOMICO = sorted([...])
RISCOS_ACIDENTE = sorted([...])

AGENTES_POR_CATEGORIA = {
    'fisico': RISCOS_FISICO,
    'quimico': RISCOS_QUIMICO,
    'biologico': RISCOS_BIOLOGICO,
    'ergonomico': RISCOS_ERGONOMICO,
    'acidente': RISCOS_ACIDENTE,
}

CATEGORIAS_RISCO = {...}

# CSS tema preto total
st.markdown("""
<style>
    .stApp, .main {
        background: #000000 !important;
        color: #ffffff !important;
    }
    header[data-testid="stHeader"] {
        height: 0px; max-height: 0px; overflow: hidden;
    }
    .main .block-container {
        padding-top: 1rem; padding-bottom: 1rem;
    }
    .login-container {
        max-width: 500px;
        margin: 2rem auto;
        padding: 2rem;
        background: #111111;
        border-radius: 15px;
        box-shadow: 0 8px 32px rgba(0,255,0,0.2);
        border: 2px solid #00ff00;
    }
    /* ... estilos para inputs, botões, cards, etc. */
</style>
""", unsafe_allow_html=True)


def is_valid_email(email):
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    valid_domains = ['gmail.com', 'outlook.com', 'hotmail.com', 'yahoo.com', 'yahoo.com.br',
        'uol.com.br', 'terra.com.br', 'bol.com.br', 'ig.com.br', 'globo.com',
        'live.com', 'msn.com', 'icloud.com', 'me.com', 'mac.com',
        'protonmail.com', 'zoho.com', 'yandex.com']
    if re.match(pattern, email):
        domain = email.split('@')[1].lower()
        return domain in valid_domains
    return False


def hash_password(pw):
    return hashlib.sha256(pw.encode()).hexdigest()

def initialize_users():
    if 'users_db' not in st.session_state:
        st.session_state.users_db = {
            'robsonmbf@hotmail.com': {
                'id': 1,
                'password': hash_password('123456'),
                'nome': 'Robson',
                'empresa': 'Minha Empresa',
                'email': 'robsonmbf@hotmail.com',
                'credits': 999999,
                'is_admin': True
            }
        }

def authenticate_user(email, password):
    initialize_users()
    pw_hash = hash_password(password)
    user = st.session_state.users_db.get(email)
    if user and user['password'] == pw_hash:
        return user
    return None

def register_user(email, password, nome, empresa):
    initialize_users()
    if email in st.session_state.users_db:
        return None
    user_id = len(st.session_state.users_db) + 1
    st.session_state.users_db[email] = {
        'id': user_id,
        'password': hash_password(password),
        'nome': nome,
        'empresa': empresa,
        'email': email,
        'credits': 100,
        'is_admin': False
    }
    return user_id

def get_user_credits(user_id):
    initialize_users()
    for user in st.session_state.users_db.values():
        if user['id'] == user_id:
            return "∞" if user.get('is_admin') else user['credits']
    return 0

def debit_credits(user_id, amount):
    initialize_users()
    for user in st.session_state.users_db.values():
        if user['id'] == user_id:
            if user.get('is_admin'):
                return True
            user['credits'] = max(0, user['credits'] - amount)
            return True
    return False

def check_sufficient_credits(user_id, amount):
    initialize_users()
    for user in st.session_state.users_db.values():
        if user['id'] == user_id:
            if user.get('is_admin'):
                return True
            return user['credits'] >= amount
    return False

def create_sample_data():
    return pd.DataFrame({
        'Nome': ['JOÃO SILVA SANTOS', 'MARIA OLIVEIRA COSTA', 'PEDRO ALVES FERREIRA'],
        'Setor': ['PRODUCAO DE LA DE ACO', 'ADMINISTRACAO DE RH', 'MANUTENCAO QUIMICA'],
        'Função': ['OPERADOR PRODUCAO I', 'ANALISTA ADM PESSOAL PL', 'MECANICO MANUT II'],
        'Data de Admissão': ['15/03/2020', '22/08/2019', '10/01/2021'],
        'Empresa': ['SUA EMPRESA', 'SUA EMPRESA', 'SUA EMPRESA'],
        'Unidade': ['Matriz', 'Matriz', 'Matriz'],
        'Descrição de Atividades': [
            'Operar equipamentos de produção nível I ...',
            'Executar atividades de administração de pessoal ...',
            'Executar manutenção preventiva e corretiva ...'
        ]
    })

def validate_excel_structure(df):
    required_cols = ['Nome', 'Setor', 'Função', 'Data de Admissão', 'Empresa', 'Unidade', 'Descrição de Atividades']
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        return False, f"Colunas obrigatórias faltando: {', '.join(missing)}"
    if df.empty:
        return False, "A planilha está vazia"
    return True, "Estrutura válida"

def gerar_documento_os(dados_funcionario, agentes_risco, epis, medidas, observacoes, template_doc=None):
    import copy
    if template_doc:
        doc = copy.deepcopy(template_doc)
        for p in doc.paragraphs:
            p.text = (p.text.replace('{{NOME}}', dados_funcionario.get('Nome', ''))
                          .replace('{{FUNCAO}}', dados_funcionario.get('Função', ''))
                          .replace('{{SETOR}}', dados_funcionario.get('Setor', ''))
                          .replace('{{EMPRESA}}', dados_funcionario.get('Empresa', ''))
                          .replace('{{UNIDADE}}', dados_funcionario.get('Unidade', ''))
                          .replace('{{ATIVIDADES}}', dados_funcionario.get('Descrição de Atividades', '')))
        # Pode colocar o código para adicionar riscos, epis, prevenções se quiser no modelo também
    else:
        doc = Document()
        # Modelo original aqui
    
    return doc

def show_login_page():
    st.markdown('<div class="title-header">🔐 Gerador de Ordens de Serviço (OS)</div>', unsafe_allow_html=True)
    total_riscos = sum(len(r) for r in AGENTES_POR_CATEGORIA.values())
    st.markdown(f"""
    <div class="info-msg">
      <strong>🆕 SISTEMA ATUALIZADO - Base Expandida de Riscos!</strong><br><br>
      ✨ <strong>{total_riscos} opções</strong> de riscos ocupacionais organizados em 5 categorias<br>
      📄 Sistema profissional conforme NR-01 com tema black!
    </div>""", unsafe_allow_html=True)

    login_tab, register_tab = st.tabs(["🔑 Login", "👤 Criar Conta"])
    with login_tab:
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        st.markdown('<div class="login-title">🔑 Faça seu Login</div>', unsafe_allow_html=True)
        with st.form("login_form"):
            email = st.text_input("📧 Email:", placeholder="seu@gmail.com")
            password = st.text_input("🔒 Senha:", type="password", placeholder="Sua senha")
            login_button = st.form_submit_button("🚀 Entrar")
            if login_button:
                if email and password:
                    if is_valid_email(email):
                        user = authenticate_user(email, password)
                        if user:
                            st.session_state.user = user
                            st.session_state.authenticated = True
                            st.markdown('<div class="success-msg">✅ Login realizado com sucesso!</div>', unsafe_allow_html=True)
                            time.sleep(1)
                            st.experimental_rerun()
                        else:
                            st.markdown('<div class="error-msg">❌ Email ou senha incorretos.</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="error-msg">❌ Email deve ser de um provedor válido (Gmail, Outlook, Yahoo, etc.)</div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="error-msg">⚠️ Por favor, preencha todos os campos.</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with register_tab:
        # Similar cadastro
        pass

def show_main_app(user):
    col1, col2, col3 = st.columns([3,1,1])
    with col1:
        st.markdown(f"# 📄 Gerador de OS - Bem-vindo, **{user['nome']}**!")
    with col2:
        credits = get_user_credits(user['id'])
        st.metric("💳 Créditos", credits)
    with col3:
        if st.button("🚪 Logout"):
            st.session_state.authenticated = False
            st.session_state.user = None
            st.experimental_rerun()
            
    st.markdown(f"🏢 **Empresa:** {user['empresa']}")
    if user.get('is_admin', False):
        st.markdown("""
          <div class="warning-msg">
          <strong>👑 CONTA ADMINISTRADOR</strong><br>
          • Créditos ilimitados<br>
          • Não há cobrança de créditos<br>
          • Acesso completo ao sistema
          </div>
          """, unsafe_allow_html=True)

    total_riscos = sum(len(r) for r in AGENTES_POR_CATEGORIA.values())
    st.markdown(f"""
      <div class="info-msg">
      <strong>🚀 SISTEMA ATUALIZADO - Nova Base de Riscos!</strong><br><br>
      📊 <strong>Total:</strong> {total_riscos} opções de riscos ocupacionais organizados em 5 categorias<br>
      </div>
      """, unsafe_allow_html=True)

    # Upload e filtros (setor, função)
    uploaded_excel = st.file_uploader("Selecione a planilha Excel", type=['xlsx'])
    uploaded_template = st.file_uploader("Selecione template Word (opcional)", type=['docx'])
    if uploaded_excel:
        df = pd.read_excel(uploaded_excel)
        is_valid, msg = validate_excel_structure(df)
        if not is_valid:
            st.error(msg)
            return
        st.success(f"Planilha carregada: {len(df)} funcionários")
        setores = sorted(df['Setor'].dropna().unique())
        funcoes = sorted(df['Função'].dropna().unique())
        selected_setores = st.multiselect("Filtrar por Setores:", setores)
        selected_funcoes = st.multiselect("Filtrar por Funções:", funcoes)

        df_filtered = df
        if selected_setores:
            df_filtered = df_filtered[df_filtered['Setor'].isin(selected_setores)]
        if selected_funcoes:
            df_filtered = df_filtered[df_filtered['Função'].isin(selected_funcoes)]

        modo_selecao = st.radio("Modo de Seleção:", ["Individual", "Múltiplos", "Todos Filtrados"])
        funcionarios_selecionados = []
        if modo_selecao == "Individual":
            funcionarios_selecionados = [st.selectbox("Selecione funcionário:", df_filtered['Nome'].tolist())]
        elif modo_selecao == "Múltiplos":
            funcionarios_selecionados = st.multiselect("Selecione funcionários:", df_filtered['Nome'].tolist())
        else:
            funcionarios_selecionados = df_filtered['Nome'].tolist()

        if funcionarios_selecionados:
            st.success(f"{len(funcionarios_selecionados)} funcionários selecionados")
            # Gerar documentos
            if uploaded_template:
                template_doc = Document(uploaded_template)
            else:
                template_doc = None

            if st.button("📄 Gerar OS"):
                documentos = []
                progress = st.progress(0)
                for i, nome_fun in enumerate(funcionarios_selecionados):
                    dados_fun = df_filtered[df_filtered['Nome'] == nome_fun].iloc[0].to_dict()
                    doc = gerar_documento_os(
                        dados_fun, None, None, None, None,
                        template_doc
                    )
                    out_buffer = BytesIO()
                    doc.save(out_buffer)
                    out_buffer.seek(0)
                    documentos.append((nome_fun, out_buffer))
                    progress.progress((i+1)/len(funcionarios_selecionados))

                if len(documentos) == 1:
                    st.download_button("Download OS", data=documentos[0][1].getvalue(),
                                       file_name=f"OS_{documentos[0][0]}.docx",
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                else:
                    zip_buffer = BytesIO()
                    import zipfile
                    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
                        for nome, buffer in documentos:
                            zipf.writestr(f"OS_{nome}.docx", buffer.getvalue())
                    zip_buffer.seek(0)
                    st.download_button("Download ZIP OS", data=zip_buffer.read(),
                                       file_name="OS_Lote.zip",
                                       mime="application/zip")

def main():
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'user' not in st.session_state:
        st.session_state.user = None
    if st.session_state.authenticated and st.session_state.user:
        show_main_app(st.session_state.user)
    else:
        show_login_page()

if __name__=="__main__":
    main()
