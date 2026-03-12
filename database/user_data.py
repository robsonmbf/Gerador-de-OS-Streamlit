from datetime import datetime
from database.models import DatabaseManager
from utils.security import sanitize_input
import json

class UserDataManager:
    def __init__(self, db_manager):
        self.db = db_manager
    
    # ===== GERENCIAMENTO DE MEDIÇÕES =====
    
    def add_measurement(self, user_id, agent, value, unit, epi=None):
        """Adiciona uma medição para o usuário"""
        # Sanitizar entradas
        agent = sanitize_input(agent)
        value = sanitize_input(value)
        unit = sanitize_input(unit)
        epi = sanitize_input(epi) if epi else None

        if not agent or not value or not unit:
            return False, "Agente, valor e unidade são obrigatórios", None

        conn = self.db.get_connection()
        cursor = conn.cursor()

        # Evita duplicidades ativas para o mesmo usuário
        cursor.execute('''
            SELECT id FROM user_measurements
            WHERE user_id = ? AND agent = ? AND value = ? AND unit = ?
              AND COALESCE(epi, '') = COALESCE(?, '') AND is_active = TRUE
        ''', (user_id, agent, value, unit, epi))

        if cursor.fetchone():
            conn.close()
            return False, "Medição já cadastrada", None

        try:
            cursor.execute('''
                INSERT INTO user_measurements (user_id, agent, value, unit, epi)
                VALUES (?, ?, ?, ?, ?)
            ''', (user_id, agent, value, unit, epi))

            measurement_id = cursor.lastrowid
            conn.commit()
            conn.close()

            # Log da atividade
            self.db.log_activity(user_id, 'add_measurement', {
                'agent': agent,
                'value': value,
                'unit': unit,
                'epi': epi
            })

            return True, "Medição adicionada com sucesso", measurement_id

        except Exception as e:
            conn.close()
            return False, f"Erro ao adicionar medição: {str(e)}", None
    
    def get_user_measurements(self, user_id):
        """Retorna todas as medições ativas do usuário"""
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, agent, value, unit, epi, created_at
            FROM user_measurements
            WHERE user_id = ? AND is_active = TRUE
            ORDER BY created_at DESC
        ''', (user_id,))
        
        measurements = []
        for row in cursor.fetchall():
            measurements.append({
                'id': row['id'],
                'agent': row['agent'],
                'value': row['value'],
                'unit': row['unit'],
                'epi': row['epi'],
                'created_at': row['created_at']
            })
        
        conn.close()
        return measurements
    
    def remove_measurement(self, user_id, measurement_id):
        """Remove uma medição do usuário"""
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        # Verificar se a medição pertence ao usuário
        cursor.execute('''
            SELECT agent, value, unit FROM user_measurements
            WHERE id = ? AND user_id = ? AND is_active = TRUE
        ''', (measurement_id, user_id))
        
        measurement = cursor.fetchone()
        if not measurement:
            conn.close()
            return False, "Medição não encontrada"
        
        # Marcar como inativa
        cursor.execute('''
            UPDATE user_measurements
            SET is_active = FALSE
            WHERE id = ? AND user_id = ?
        ''', (measurement_id, user_id))
        
        conn.commit()
        
        # Log da atividade
        self.db.log_activity(user_id, 'remove_measurement', {
            'measurement_id': measurement_id,
            'agent': measurement['agent'],
            'value': measurement['value'],
            'unit': measurement['unit']
        })
        
        conn.close()
        return True, "Medição removida com sucesso"
    
    # ===== GERENCIAMENTO DE EPIs =====
    
    def add_epi(self, user_id, epi_name):
        """Adiciona um EPI para o usuário"""
        epi_name = sanitize_input(epi_name)
        
        if not epi_name:
            return False, "Nome do EPI é obrigatório", None
        
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        # Verificar se EPI já existe para o usuário
        cursor.execute('''
            SELECT id FROM user_epis
            WHERE user_id = ? AND epi_name = ? AND is_active = TRUE
        ''', (user_id, epi_name))
        
        if cursor.fetchone():
            conn.close()
            return False, "EPI já adicionado", None
        
        try:
            cursor.execute('''
                INSERT INTO user_epis (user_id, epi_name)
                VALUES (?, ?)
            ''', (user_id, epi_name))
            
            epi_id = cursor.lastrowid
            conn.commit()
            
            # Log da atividade
            self.db.log_activity(user_id, 'add_epi', {
                'epi_name': epi_name
            })
            
            conn.close()
            return True, "EPI adicionado com sucesso", epi_id
        
        except Exception as e:
            conn.close()
            return False, f"Erro ao adicionar EPI: {str(e)}", None
    
    def get_user_epis(self, user_id):
        """Retorna todos os EPIs ativos do usuário"""
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, epi_name, created_at
            FROM user_epis
            WHERE user_id = ? AND is_active = TRUE
            ORDER BY epi_name
        ''', (user_id,))
        
        epis = []
        for row in cursor.fetchall():
            epis.append({
                'id': row['id'],
                'epi_name': row['epi_name'],
                'created_at': row['created_at']
            })
        
        conn.close()
        return epis
    
    def remove_epi(self, user_id, epi_id):
        """Remove um EPI do usuário"""
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        # Verificar se o EPI pertence ao usuário
        cursor.execute('''
            SELECT epi_name FROM user_epis
            WHERE id = ? AND user_id = ? AND is_active = TRUE
        ''', (epi_id, user_id))
        
        epi = cursor.fetchone()
        if not epi:
            conn.close()
            return False, "EPI não encontrado"
        
        # Marcar como inativo
        cursor.execute('''
            UPDATE user_epis
            SET is_active = FALSE
            WHERE id = ? AND user_id = ?
        ''', (epi_id, user_id))
        
        conn.commit()
        
        # Log da atividade
        self.db.log_activity(user_id, 'remove_epi', {
            'epi_id': epi_id,
            'epi_name': epi['epi_name']
        })
        
        conn.close()
        return True, "EPI removido com sucesso"
    
    # ===== GERENCIAMENTO DE RISCOS MANUAIS =====
    
    def add_manual_risk(self, user_id, category, risk_name, possible_damages=None):
        """Adiciona um risco manual para o usuário"""
        category = sanitize_input(category)
        risk_name = sanitize_input(risk_name)
        possible_damages = sanitize_input(possible_damages) if possible_damages else None

        if not category or not risk_name:
            return False, "Categoria e nome do risco são obrigatórios", None

        conn = self.db.get_connection()
        cursor = conn.cursor()

        # Evita riscos duplicados por categoria para o mesmo usuário
        cursor.execute('''
            SELECT id FROM user_manual_risks
            WHERE user_id = ? AND category = ? AND risk_name = ? AND is_active = TRUE
        ''', (user_id, category, risk_name))

        if cursor.fetchone():
            conn.close()
            return False, "Risco manual já adicionado", None

        try:
            cursor.execute('''
                INSERT INTO user_manual_risks (user_id, category, risk_name, possible_damages)
                VALUES (?, ?, ?, ?)
            ''', (user_id, category, risk_name, possible_damages))

            risk_id = cursor.lastrowid
            conn.commit()
            conn.close()

            # Log da atividade
            self.db.log_activity(user_id, 'add_manual_risk', {
                'category': category,
                'risk_name': risk_name,
                'possible_damages': possible_damages
            })

            return True, "Risco manual adicionado com sucesso", risk_id

        except Exception as e:
            conn.close()
            return False, f"Erro ao adicionar risco manual: {str(e)}", None
    
    def get_user_manual_risks(self, user_id):
        """Retorna todos os riscos manuais ativos do usuário"""
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, category, risk_name, possible_damages, created_at
            FROM user_manual_risks
            WHERE user_id = ? AND is_active = TRUE
            ORDER BY category, risk_name
        ''', (user_id,))
        
        risks = []
        for row in cursor.fetchall():
            risks.append({
                'id': row['id'],
                'category': row['category'],
                'risk_name': row['risk_name'],
                'possible_damages': row['possible_damages'],
                'created_at': row['created_at']
            })
        
        conn.close()
        return risks
    
    def remove_manual_risk(self, user_id, risk_id):
        """Remove um risco manual do usuário"""
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        # Verificar se o risco pertence ao usuário
        cursor.execute('''
            SELECT category, risk_name FROM user_manual_risks
            WHERE id = ? AND user_id = ? AND is_active = TRUE
        ''', (risk_id, user_id))
        
        risk = cursor.fetchone()
        if not risk:
            conn.close()
            return False, "Risco manual não encontrado"
        
        # Marcar como inativo
        cursor.execute('''
            UPDATE user_manual_risks
            SET is_active = FALSE
            WHERE id = ? AND user_id = ?
        ''', (risk_id, user_id))
        
        conn.commit()
        
        # Log da atividade
        self.db.log_activity(user_id, 'remove_manual_risk', {
            'risk_id': risk_id,
            'category': risk['category'],
            'risk_name': risk['risk_name']
        })
        
        conn.close()
        return True, "Risco manual removido com sucesso"
    

    # ===== GERENCIAMENTO DE MODELOS DOCX =====

    def add_docx_template(self, user_id, template_name, file_content):
        """Adiciona um modelo DOCX para o usuário"""
        template_name = sanitize_input(template_name)

        if not template_name:
            return False, "Nome do modelo é obrigatório", None

        if not file_content:
            return False, "Arquivo do modelo é obrigatório", None

        conn = self.db.get_connection()
        cursor = conn.cursor()

        cursor.execute(
            '''
            SELECT id FROM user_docx_templates
            WHERE user_id = ? AND template_name = ? AND is_active = TRUE
            ''',
            (user_id, template_name)
        )

        if cursor.fetchone():
            conn.close()
            return False, "Já existe um modelo ativo com este nome", None

        try:
            cursor.execute(
                '''
                INSERT INTO user_docx_templates (user_id, template_name, file_content)
                VALUES (?, ?, ?)
                ''',
                (user_id, template_name, file_content)
            )

            template_id = cursor.lastrowid
            conn.commit()
            conn.close()

            self.db.log_activity(
                user_id,
                'add_docx_template',
                {'template_name': template_name}
            )

            return True, "Modelo DOCX adicionado com sucesso", template_id

        except Exception as e:
            conn.close()
            return False, f"Erro ao adicionar modelo DOCX: {str(e)}", None

    def get_user_docx_templates(self, user_id):
        """Retorna os modelos DOCX ativos do usuário"""
        conn = self.db.get_connection()
        cursor = conn.cursor()

        cursor.execute(
            '''
            SELECT id, template_name, file_content, created_at
            FROM user_docx_templates
            WHERE user_id = ? AND is_active = TRUE
            ORDER BY template_name
            ''',
            (user_id,)
        )

        templates = []
        for row in cursor.fetchall():
            templates.append({
                'id': row['id'],
                'template_name': row['template_name'],
                'file_content': row['file_content'],
                'created_at': row['created_at']
            })

        conn.close()
        return templates

    def remove_docx_template(self, user_id, template_id):
        """Remove (inativa) um modelo DOCX do usuário"""
        conn = self.db.get_connection()
        cursor = conn.cursor()

        cursor.execute(
            '''
            SELECT template_name FROM user_docx_templates
            WHERE id = ? AND user_id = ? AND is_active = TRUE
            ''',
            (template_id, user_id)
        )

        template = cursor.fetchone()
        if not template:
            conn.close()
            return False, "Modelo DOCX não encontrado"

        cursor.execute(
            '''
            UPDATE user_docx_templates
            SET is_active = FALSE
            WHERE id = ? AND user_id = ?
            ''',
            (template_id, user_id)
        )

        conn.commit()

        self.db.log_activity(
            user_id,
            'remove_docx_template',
            {'template_id': template_id, 'template_name': template['template_name']}
        )

        conn.close()
        return True, "Modelo DOCX removido com sucesso"

    # ===== FUNÇÕES AUXILIARES =====
    
    def get_user_summary(self, user_id):
        """Retorna um resumo dos dados do usuário"""
        measurements = self.get_user_measurements(user_id)
        epis = self.get_user_epis(user_id)
        risks = self.get_user_manual_risks(user_id)
        templates = self.get_user_docx_templates(user_id)
        activities = self.db.get_user_activities(user_id, limit=10)
        
        return {
            'measurements_count': len(measurements),
            'epis_count': len(epis),
            'manual_risks_count': len(risks),
            'docx_templates_count': len(templates),
            'recent_activities': activities,
            'measurements': measurements,
            'epis': epis,
            'manual_risks': risks,
            'docx_templates': templates
        }
    
    def clear_user_data(self, user_id, data_type='all'):
        """Limpa dados específicos do usuário"""
        conn = self.db.get_connection()
        cursor = conn.cursor()
        
        try:
            if data_type in ['all', 'measurements']:
                cursor.execute('''
                    UPDATE user_measurements
                    SET is_active = FALSE
                    WHERE user_id = ?
                ''', (user_id,))
            
            if data_type in ['all', 'epis']:
                cursor.execute('''
                    UPDATE user_epis
                    SET is_active = FALSE
                    WHERE user_id = ?
                ''', (user_id,))
            
            if data_type in ['all', 'risks']:
                cursor.execute('''
                    UPDATE user_manual_risks
                    SET is_active = FALSE
                    WHERE user_id = ?
                ''', (user_id,))

            if data_type in ['all', 'templates']:
                cursor.execute('''
                    UPDATE user_docx_templates
                    SET is_active = FALSE
                    WHERE user_id = ?
                ''', (user_id,))
            
            conn.commit()
            
            # Log da atividade
            self.db.log_activity(user_id, 'clear_user_data', {
                'data_type': data_type
            })
            
            conn.close()
            return True, f"Dados {data_type} limpos com sucesso"
        
        except Exception as e:
            conn.close()
            return False, f"Erro ao limpar dados: {str(e)}"

