from supabase_client import supabase
import streamlit as st
class SupabaseCreditManager:
def get_user_credits(self, user_id: int):
"""Obter créditos do usuário"""
try:
if not supabase:
return {'balance': 0, 'total_purchased': 0, 'total_used': 0}
result = supabase.table('user_credits').select('*').eq('user_id', user_id).ex
if result.data:
return result.data[0]
else:
# Criar registro se não existir
supabase.table('user_credits').insert({
'user_id': user_id,
'balance': 0
}).execute()
return {'balance': 0, 'total_purchased': 0, 'total_used': 0}
except Exception as e:
st.error(f"Erro ao obter créditos: {e}")
return {'balance': 0, 'total_purchased': 0, 'total_used': 0}
def get_credit_packages(self):
"""Obter pacotes disponíveis"""
try:
if not supabase:
return []
result = supabase.table('credit_packages').select('*').eq('is_active', True).
return result.data if result.data else []
except Exception as e:
st.error(f"Erro ao carregar pacotes: {e}")
return []
def use_credits(self, user_id: int, amount: int, description: str):
"""Usar créditos do usuário"""
try:
if not supabase:
return {"success": False, "message": "Erro na conexão"}
# Verificar saldo atual
credits = self.get_user_credits(user_id)
 O que fazer:
 Create new file: credits_supabase.py
 Cole este código:
 Código completo:
if credits['balance'] < amount:
return {"success": False, "message": "Créditos insuficientes"}
# Calcular novo saldo
new_balance = credits['balance'] - amount
new_used = credits.get('total_used', 0) + amount
# Atualizar créditos
supabase.table('user_credits').update({
'balance': new_balance,
'total_used': new_used
}).eq('user_id', user_id).execute()
# Registrar transação
supabase.table('transactions').insert({
'user_id': user_id,
'type': 'usage',
'amount': -amount,
'description': description
}).execute()
return {"success": True, "new_balance": new_balance}
except Exception as e:
return {"success": False, "message": f"Erro: {str(e)}"}
def purchase_credits(self, user_id: int, package_id: int):
"""Comprar pacote de créditos"""
try:
if not supabase:
return {"success": False, "message": "Erro na conexão"}
# Buscar informações do pacote
package_result = supabase.table('credit_packages').select('*').eq('id', packa
if not package_result.data:
return {"success": False, "message": "Pacote não encontrado"}
package = package_result.data[0]
# Simular pagamento aprovado (integre com gateway real aqui)
# Obter créditos atuais
credits = self.get_user_credits(user_id)
# Calcular novos valores
new_balance = credits['balance'] + package['credits']
new_purchased = credits.get('total_purchased', 0) + package['credits']
# Atualizar créditos
supabase.table('user_credits').update({
'balance': new_balance,
'total_purchased': new_purchased,
'last_purchase': 'NOW()'
}).eq('user_id', user_id).execute()
# Registrar transação
supabase.table('transactions').insert({
'user_id': user_id,
'type': 'purchase',
'amount': package['credits'],
'description': f"Compra: {package['name']}",
'metadata': f"package_id:{package_id},price:{package['price']}"
}).execute()
return {
"success": True,
"message": f"Compra realizada! {package['credits']} créditos adicionados.
"new_balance": new_balance,
"package_name": package['name']
}
except Exception as e:
return {"success": False, "message": f"Erro: {str(e)}"}