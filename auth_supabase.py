from supabase_client import supabase
import hashlib
import secrets
import streamlit as st
class SupabaseAuthManager:
def _hash_password(self, password: str) -> str:
"""Cria hash seguro da senha"""
salt = secrets.token_hex(32)
password_hash = hashlib.sha256((password + salt).encode()).hexdigest()
return f"{salt}${password_hash}"
def _verify_password(self, password: str, stored_hash: str) -> bool:
"""Verifica se senha está correta"""
try:
salt, password_hash = stored_hash.split('$')
return hashlib.sha256((password + salt).encode()).hexdigest() == password_has
except:
return False
def register_user(self, username: str, email: str, password: str, full_name: str = ""
"""Cadastrar novo usuário"""
try:
if not supabase:
return {"success": False, "message": "Erro na conexão com banco"}
# Verificar se usuário já existe
existing = supabase.table('users').select('id').or_(f'username.eq.{username},
if existing.data:
 PASSO 5 Criar arquivo auth_supabase.py
 O que fazer:
 Create new file: auth_supabase.py
 Cole este código:
 Código completo:
return {"success": False, "message": "Usuário ou email já cadastrado"}
# Validações
if len(username) < 3:
return {"success": False, "message": "Nome de usuário muito curto"}
if len(password) < 6:
return {"success": False, "message": "Senha deve ter pelo menos 6 caracte
if '@' not in email:
return {"success": False, "message": "Email inválido"}
# Criar hash da senha
password_hash = self._hash_password(password)
# Inserir usuário
result = supabase.table('users').insert({
'username': username,
'email': email,
'password_hash': password_hash,
'full_name': full_name
}).execute()
if not result.data:
return {"success": False, "message": "Erro ao criar usuário"}
user_id = result.data[0]['id']
# Criar créditos iniciais (5 grátis)
supabase.table('user_credits').insert({
'user_id': user_id,
'balance': 5
}).execute()
# Registrar transação de bônus
supabase.table('transactions').insert({
'user_id': user_id,
'type': 'bonus',
'amount': 5,
'description': 'Créditos de boas-vindas'
}).execute()
return {
"success": True,
"message": "Cadastro realizado! Você ganhou 5 créditos gratuitos."
}
except Exception as e:
return {"success": False, "message": f"Erro interno: {str(e)}"}
def login_user(self, username: str, password: str):
"""Fazer login do usuário"""
try:
if not supabase:
return {"success": False, "message": "Erro na conexão com banco"}
# Buscar usuário
result = supabase.table('users').select('*').or_(f'username.eq.{username},ema
if not result.data:
return {"success": False, "message": "Usuário não encontrado"}
user = result.data[0]
# Verificar se conta está ativa
if not user.get('is_active', True):
return {"success": False, "message": "Conta desativada"}
# Verificar senha
if not self._verify_password(password, user['password_hash']):
return {"success": False, "message": "Senha incorreta"}
# Buscar créditos do usuário
credits_result = supabase.table('user_credits').select('*').eq('user_id', use
credits = credits_result.data[0] if credits_result.data else {
'balance': 0, 'total_purchased': 0, 'total_used': 0
}
# Atualizar último login
supabase.table('users').update({
'last_login': 'NOW()'
}).eq('id', user['id']).execute()
# Preparar dados do usuário
user_info = {
'id': user['id'],
'username': user['username'],
'email': user['email'],
'full_name': user['full_name'],
'is_admin': user.get('is_admin', False),
'balance': credits['balance'],
'total_purchased': credits.get('total_purchased', 0),
'total_used': credits.get('total_used', 0)
}
return {"success": True, "user": user_info}
except Exception as e:
return {"success": False, "message": f"Erro interno: {str(e)}"}