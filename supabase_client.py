import streamlit as st
from supabase import create_client, Client
def init_supabase() -> Client:
"""Inicializa cliente Supabase usando secrets do Streamlit"""
try:
# Usar secrets configurados no Streamlit Cloud
 Código para colar (substituir pelas suas informações):
 PASSO 3 Atualizar requirements.txt
 O que fazer:
 Vá para seu repositório GitHub
 Abra o arquivo "requirements.txt"
 Clique "Edit" (ícone de lápis)
 Adicione esta linha:
 Código para adicionar:
 PASSO 4 Criar arquivo supabase_client.py
 O que fazer:
 No seu repositório GitHub
 Clique "Add file" → "Create new file"
 Nome: supabase_client.py
 Cole este código:
 Código completo:
url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]
return create_client(url, key)
except Exception as e:
st.error(f"Erro na conexão Supabase: {e}")
return None
# Cliente global
supabase = init_supabase()
