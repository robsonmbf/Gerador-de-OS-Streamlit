from datetime import datetime
from database.models import DatabaseManager
from utils.security import hash_password, verify_password, generate_session_token, get_session_expiry, is_valid_email, is_strong_password, sanitize_input

class AuthManager:
    def __init__(self, db_manager):
        self.db = db_manager
    
    def register_user(self, email, password):
        """Registra um novo usuário"""
        # Sanitizar entrada
        email = sanitize_input(email).lower()
        
        # Validar email
        if not is_valid_email(email):
            return False, "Email inválido"
        
        # Validar senha
        is_valid, message = is_strong_password(password)
        if not is_valid:
            return False, message
        
        # Verificar se email já existe
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT id FROM users WHERE email = ?', (email,))
        if cursor.fetchone():
            conn.close()
            return False, "Email já está em uso"
        
        # Criar hash da senha
        password_hash = hash_password(password)
        
        # Inserir usuário
        try:
            cursor.execute('''
                INSERT INTO users (email, password_hash)
                VALUES (?, ?)
            ''', (email, password_hash))
            
            user_id = cursor.lastrowid
            conn.commit()
            
            # Log da atividade
            self.db.log_activity(user_id, 'user_register', {'email': email})
            
            conn.close()
            return True, "Usuário registrado com sucesso"
        
        except Exception as e:
            conn.close()
            return False, f"Erro ao registrar usuário: {str(e)}"
    
    def login_user(self, email, password):
        """Autentica um usuário e cria sessão"""
        # Sanitizar entrada
        email = sanitize_input(email).lower()
        
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        # Buscar usuário
        cursor.execute('''
            SELECT id, email, password_hash, is_active
            FROM users
            WHERE email = ?
        ''', (email,))
        
        user = cursor.fetchone()
        if not user:
            conn.close()
            return False, "Email ou senha incorretos", None
        
        if not user['is_active']:
            conn.close()
            return False, "Conta desativada", None
        
        # Verificar senha
        if not verify_password(password, user['password_hash']):
            conn.close()
            return False, "Email ou senha incorretos", None
        
        # Criar sessão
        session_token = generate_session_token()
        expires_at = get_session_expiry()
        
        try:
            cursor.execute('''
                INSERT INTO user_sessions (user_id, session_token, expires_at)
                VALUES (?, ?, ?)
            ''', (user['id'], session_token, expires_at))
            
            # Atualizar último login
            cursor.execute('''
                UPDATE users
                SET last_login = CURRENT_TIMESTAMP
                WHERE id = ?
            ''', (user['id'],))
            
            conn.commit()
            
            # Log da atividade
            self.db.log_activity(user['id'], 'user_login', {'email': email})
            
            conn.close()
            
            return True, "Login realizado com sucesso", {
                'user_id': user['id'],
                'email': user['email'],
                'session_token': session_token,
                'expires_at': expires_at
            }
        
        except Exception as e:
            conn.close()
            return False, f"Erro ao criar sessão: {str(e)}", None
    
    def validate_session(self, session_token):
        """Valida uma sessão ativa"""
        if not session_token:
            return False, None
        
        # Limpar sessões expiradas
        self.db.cleanup_expired_sessions()
        
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT s.user_id, u.email, s.expires_at
            FROM user_sessions s
            JOIN users u ON s.user_id = u.id
            WHERE s.session_token = ? AND s.is_active = TRUE AND u.is_active = TRUE
        ''', (session_token,))
        
        session = cursor.fetchone()
        conn.close()
        
        if not session:
            return False, None
        
        # Verificar se não expirou
        expires_at = datetime.fromisoformat(session['expires_at'])
        if datetime.now() > expires_at:
            self.logout_user(session_token)
            return False, None
        
        return True, {
            'user_id': session['user_id'],
            'email': session['email'],
            'expires_at': session['expires_at']
        }
    
    def logout_user(self, session_token):
        """Faz logout do usuário invalidando a sessão"""
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        # Buscar user_id antes de invalidar
        cursor.execute('''
            SELECT user_id FROM user_sessions
            WHERE session_token = ? AND is_active = TRUE
        ''', (session_token,))
        
        session = cursor.fetchone()
        if session:
            user_id = session['user_id']
            
            # Invalidar sessão
            cursor.execute('''
                UPDATE user_sessions
                SET is_active = FALSE
                WHERE session_token = ?
            ''', (session_token,))
            
            conn.commit()
            
            # Log da atividade
            self.db.log_activity(user_id, 'user_logout')
        
        conn.close()
        return True
    
    def get_user_info(self, user_id):
        """Retorna informações do usuário"""
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, email, created_at, last_login
            FROM users
            WHERE id = ? AND is_active = TRUE
        ''', (user_id,))
        
        user = cursor.fetchone()
        conn.close()
        
        if user:
            return {
                'id': user['id'],
                'email': user['email'],
                'created_at': user['created_at'],
                'last_login': user['last_login']
            }
        
        return None

