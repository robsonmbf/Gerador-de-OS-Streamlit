import bcrypt
import uuid
import re
from datetime import datetime, timedelta

def hash_password(password):
    """Gera hash seguro da senha usando bcrypt"""
    salt = bcrypt.gensalt()
    password_hash = bcrypt.hashpw(password.encode('utf-8'), salt)
    return password_hash.decode('utf-8')

def verify_password(password, password_hash):
    """Verifica se a senha corresponde ao hash"""
    return bcrypt.checkpw(password.encode('utf-8'), password_hash.encode('utf-8'))

def generate_session_token():
    """Gera um token único para sessão"""
    return str(uuid.uuid4())

def is_valid_email(email):
    """Valida formato do email"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def is_strong_password(password):
    """Verifica se a senha atende aos critérios de segurança"""
    if len(password) < 8:
        return False, "A senha deve ter pelo menos 8 caracteres"
    
    if not re.search(r'[A-Za-z]', password):
        return False, "A senha deve conter pelo menos uma letra"
    
    if not re.search(r'\d', password):
        return False, "A senha deve conter pelo menos um número"
    
    return True, "Senha válida"

def get_session_expiry(hours=24):
    """Retorna data de expiração da sessão"""
    return datetime.now() + timedelta(hours=hours)

def sanitize_input(text):
    """Sanitiza entrada de texto para prevenir injeções"""
    if not isinstance(text, str):
        return str(text)
    
    # Remove caracteres potencialmente perigosos
    sanitized = re.sub(r'[<>"\';]', '', text)
    return sanitized.strip()

