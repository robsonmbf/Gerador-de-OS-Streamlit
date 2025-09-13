import sqlite3
import os
from datetime import datetime, timedelta
import json

class DatabaseManager:
    def __init__(self, db_path="os_generator.db"):
        self.db_path = db_path
        self.init_database()
    
    def get_connection(self):
        """Retorna uma conexão com o banco de dados"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row  # Permite acessar colunas por nome
        return conn
    
    def init_database(self):
        """Inicializa o banco de dados com todas as tabelas necessárias"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Tabela de usuários
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_login TIMESTAMP,
                is_active BOOLEAN DEFAULT TRUE
            )
        ''')
        
        # Tabela de sessões de usuário
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                session_token TEXT UNIQUE NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                expires_at TIMESTAMP NOT NULL,
                is_active BOOLEAN DEFAULT TRUE,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')
        
        # Tabela de atividades do usuário
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_activities (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                activity_type TEXT NOT NULL,
                activity_data TEXT,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')
        
        # Tabela de medições do usuário
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_measurements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                agent TEXT NOT NULL,
                value TEXT NOT NULL,
                unit TEXT NOT NULL,
                epi TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                is_active BOOLEAN DEFAULT TRUE,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')
        
        # Tabela de EPIs do usuário
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_epis (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                epi_name TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                is_active BOOLEAN DEFAULT TRUE,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')
        
        # Tabela de riscos manuais do usuário
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS user_manual_risks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                category TEXT NOT NULL,
                risk_name TEXT NOT NULL,
                possible_damages TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                is_active BOOLEAN DEFAULT TRUE,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        ''')
        
        # Criar índices para melhor performance
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_users_email ON users (email)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sessions_token ON user_sessions (session_token)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sessions_user ON user_sessions (user_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_activities_user ON user_activities (user_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_measurements_user ON user_measurements (user_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_epis_user ON user_epis (user_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_risks_user ON user_manual_risks (user_id)')
        
        conn.commit()
        conn.close()
    
    def cleanup_expired_sessions(self):
        """Remove sessões expiradas do banco de dados"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            UPDATE user_sessions 
            SET is_active = FALSE 
            WHERE expires_at < ? AND is_active = TRUE
        ''', (datetime.now(),))
        
        conn.commit()
        conn.close()
    
    def log_activity(self, user_id, activity_type, activity_data=None):
        """Registra uma atividade do usuário"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        activity_data_json = json.dumps(activity_data) if activity_data else None
        
        cursor.execute('''
            INSERT INTO user_activities (user_id, activity_type, activity_data)
            VALUES (?, ?, ?)
        ''', (user_id, activity_type, activity_data_json))
        
        conn.commit()
        conn.close()
    
    def get_user_activities(self, user_id, limit=50):
        """Retorna as atividades recentes de um usuário"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT activity_type, activity_data, timestamp
            FROM user_activities
            WHERE user_id = ?
            ORDER BY timestamp DESC
            LIMIT ?
        ''', (user_id, limit))
        
        activities = []
        for row in cursor.fetchall():
            activity_data = json.loads(row['activity_data']) if row['activity_data'] else {}
            activities.append({
                'type': row['activity_type'],
                'data': activity_data,
                'timestamp': row['timestamp']
            })
        
        conn.close()
        return activities

