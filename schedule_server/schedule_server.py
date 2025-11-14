import socket
import threading
import sqlite3
import json
import base64
import os
from datetime import datetime
from calendar import monthrange

class ScheduleServer:
    def __init__(self, host='127.0.0.1', port=8888, db_path='schedule_server.db'):
        self.host = host
        self.port = port
        self.db_path = db_path
        self.clients = []
        self.lock = threading.Lock()
        self.init_database()
    
    def init_database(self):
        """Инициализация базы данных"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Таблица сотрудников
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS employees (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL,
                    position TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Таблица периодов (месяцев)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS periods (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period TEXT UNIQUE NOT NULL,
                    year INTEGER NOT NULL,
                    month INTEGER NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Таблица сотрудников в периодах
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS period_employees (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period_id INTEGER NOT NULL,
                    employee_id INTEGER NOT NULL,
                    FOREIGN KEY (period_id) REFERENCES periods (id),
                    FOREIGN KEY (employee_id) REFERENCES employees (id),
                    UNIQUE(period_id, employee_id)
                )
            ''')
            
            # Таблица расписания
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS schedule (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period_id INTEGER NOT NULL,
                    employee_id INTEGER NOT NULL,
                    day INTEGER NOT NULL,
                    status INTEGER NOT NULL,
                    FOREIGN KEY (period_id) REFERENCES periods (id),
                    FOREIGN KEY (employee_id) REFERENCES employees (id),
                    UNIQUE(period_id, employee_id, day)
                )
            ''')
            
            # Таблица примечаний
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS notes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period_id INTEGER NOT NULL,
                    employee_id INTEGER NOT NULL,
                    day INTEGER NOT NULL,
                    end_time TEXT,
                    worked_hours REAL,
                    FOREIGN KEY (period_id) REFERENCES periods (id),
                    FOREIGN KEY (employee_id) REFERENCES employees (id),
                    UNIQUE(period_id, employee_id, day)
                )
            ''')
            
            # Новая таблица для Excel файлов
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS excel_files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_name TEXT NOT NULL,
                    file_data BLOB NOT NULL,
                    file_size INTEGER NOT NULL,
                    created_by TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            conn.commit()
    
    def handle_client(self, client_socket, address):
        """Обработка клиентского соединения"""
        print(f"Новое подключение: {address}")
        
        try:
            while True:
                # Получаем данные от клиента
                data = client_socket.recv(10240).decode('utf-8')
                if not data:
                    break
                
                try:
                    request = json.loads(data)
                    response = self.process_request(request)
                    client_socket.send(json.dumps(response).encode('utf-8'))
                except json.JSONDecodeError:
                    response = {"status": "error", "message": "Invalid JSON"}
                    client_socket.send(json.dumps(response).encode('utf-8'))
                
        except Exception as e:
            print(f"Ошибка с клиентом {address}: {e}")
        finally:
            client_socket.close()
            print(f"Отключение: {address}")
    
    def process_request(self, request):
        """Обработка запросов от клиента"""
        action = request.get('action')
        
        with self.lock:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                
                try:
                    if action == 'get_periods':
                        return self.get_periods(cursor)
                    elif action == 'get_schedule':
                        return self.get_schedule(cursor, request)
                    elif action == 'save_schedule':
                        return self.save_schedule(cursor, conn, request)
                    elif action == 'get_employees':
                        return self.get_employees(cursor)
                    elif action == 'add_employee':
                        return self.add_employee(cursor, conn, request)
                    elif action == 'create_period':
                        return self.create_period(cursor, conn, request)
                    elif action == 'add_employees_to_period':
                        return self.add_employees_to_period(cursor, conn, request)
                    
                    # НОВЫЕ МЕТОДЫ ДЛЯ EXCEL ФАЙЛОВ
                    elif action == 'save_excel_file':
                        return self.save_excel_file(cursor, conn, request)
                    elif action == 'get_excel_file':
                        return self.get_excel_file(cursor, request)
                    elif action == 'get_excel_files_list':
                        return self.get_excel_files_list(cursor)
                    elif action == 'delete_excel_file':
                        return self.delete_excel_file(cursor, conn, request)
                    else:
                        return {"status": "error", "message": "Unknown action"}
                except Exception as e:
                    return {"status": "error", "message": str(e)}
    
    def get_periods(self, cursor):
        """Получить список периодов"""
        cursor.execute('''
            SELECT period, year, month 
            FROM periods 
            ORDER BY year DESC, month DESC
        ''')
        periods = [dict(row) for row in cursor.fetchall()]
        return {"status": "success", "data": periods}
    
    def get_employees(self, cursor):
        """Получить список всех сотрудников"""
        cursor.execute('SELECT id, name, position FROM employees ORDER BY name')
        employees = [dict(row) for row in cursor.fetchall()]
        return {"status": "success", "data": employees}
    
    def add_employee(self, cursor, conn, request):
        """Добавить сотрудника"""
        name = request.get('name')
        if not name:
            return {"status": "error", "message": "Name is required"}
        
        try:
            cursor.execute('INSERT INTO employees (name, position) VALUES (?, ?)', 
                         (name, request.get('position', '')))
            conn.commit()
            return {"status": "success", "message": "Employee added"}
        except sqlite3.IntegrityError:
            return {"status": "error", "message": "Employee already exists"}
    
    def create_period(self, cursor, conn, request):
        """Создать новый период"""
        period = request.get('period')
        employee_ids = request.get('employee_ids', [])
        
        try:
            year, month = map(int, period.split('-'))
        except:
            return {"status": "error", "message": "Invalid period format"}
        
        try:
            # Создаем период
            cursor.execute('INSERT INTO periods (period, year, month) VALUES (?, ?, ?)', 
                         (period, year, month))
            period_id = cursor.lastrowid
            
            # Добавляем сотрудников в период
            for emp_id in employee_ids:
                cursor.execute('INSERT INTO period_employees (period_id, employee_id) VALUES (?, ?)', 
                             (period_id, emp_id))
            
            # Инициализируем пустое расписание
            days_in_month = self.get_days_in_month(year, month)
            for emp_id in employee_ids:
                for day in range(days_in_month):
                    cursor.execute('''
                        INSERT INTO schedule (period_id, employee_id, day, status) 
                        VALUES (?, ?, ?, ?)
                    ''', (period_id, emp_id, day, 4))  # 4 = пусто
            
            conn.commit()
            return {"status": "success", "message": "Period created"}
        except sqlite3.IntegrityError:
            return {"status": "error", "message": "Period already exists"}
    
    def add_employees_to_period(self, cursor, conn, request):
        """Добавить сотрудников в существующий период"""
        period = request.get('period')
        employee_ids = request.get('employee_ids', [])
        
        if not employee_ids:
            return {"status": "error", "message": "No employees selected"}
        
        # Получаем ID периода
        cursor.execute('SELECT id, year, month FROM periods WHERE period = ?', (period,))
        period_row = cursor.fetchone()
        if not period_row:
            return {"status": "error", "message": "Period not found"}
        
        period_id = period_row['id']
        year, month = period_row['year'], period_row['month']
        days_in_month = self.get_days_in_month(year, month)
        
        # Добавляем сотрудников в период
        for emp_id in employee_ids:
            try:
                cursor.execute('INSERT INTO period_employees (period_id, employee_id) VALUES (?, ?)', 
                             (period_id, emp_id))
                
                # Инициализируем пустое расписание для новых сотрудников
                for day in range(days_in_month):
                    cursor.execute('''
                        INSERT INTO schedule (period_id, employee_id, day, status) 
                        VALUES (?, ?, ?, ?)
                    ''', (period_id, emp_id, day, 4))
            except sqlite3.IntegrityError:
                continue  # Уже существует
        
        conn.commit()
        return {"status": "success", "message": f"Added {len(employee_ids)} employees"}
    
    def get_schedule(self, cursor, request):
        """Получить расписание для периода"""
        period = request.get('period')
        
        # Получаем ID периода
        cursor.execute('SELECT id, year, month FROM periods WHERE period = ?', (period,))
        period_row = cursor.fetchone()
        if not period_row:
            return {"status": "error", "message": "Period not found"}
        
        period_id = period_row['id']
        year, month = period_row['year'], period_row['month']
        days_in_month = self.get_days_in_month(year, month)
        
        # Получаем сотрудников в периоде
        cursor.execute('''
            SELECT e.id, e.name, e.position 
            FROM employees e
            JOIN period_employees pe ON e.id = pe.employee_id
            WHERE pe.period_id = ?
            ORDER BY e.name
        ''', (period_id,))
        employees = [dict(row) for row in cursor.fetchall()]
        
        # Получаем расписание
        cursor.execute('''
            SELECT s.employee_id, s.day, s.status 
            FROM schedule s
            WHERE s.period_id = ?
        ''', (period_id,))
        schedule_rows = cursor.fetchall()
        
        # Получаем примечания
        cursor.execute('''
            SELECT n.employee_id, n.day, n.end_time, n.worked_hours 
            FROM notes n
            WHERE n.period_id = ?
        ''', (period_id,))
        note_rows = cursor.fetchall()
        
        # Формируем данные
        schedule_data = {}
        for emp in employees:
            emp_id = emp['id']
            schedule_data[emp['name']] = [4] * days_in_month  # По умолчанию пусто
        
        for row in schedule_rows:
            emp_id = row['employee_id']
            day = row['day']
            status = row['status']
            
            # Находим имя сотрудника
            emp_name = next((emp['name'] for emp in employees if emp['id'] == emp_id), None)
            if emp_name and day < days_in_month:
                schedule_data[emp_name][day] = status
        
        # Формируем примечания
        notes_data = {}
        for row in note_rows:
            emp_id = row['employee_id']
            day = row['day']
            
            emp_name = next((emp['name'] for emp in employees if emp['id'] == emp_id), None)
            if emp_name:
                if emp_name not in notes_data:
                    notes_data[emp_name] = {}
                notes_data[emp_name][str(day)] = {
                    'end_time': row['end_time'],
                    'worked_hours': row['worked_hours']
                }
        
        result = {
            "employees": employees,
            "schedule": schedule_data,
            "notes": notes_data
        }
        
        return {"status": "success", "data": result}
    
    def save_schedule(self, cursor, conn, request):
        """Сохранить изменения в расписании"""
        period = request.get('period')
        schedule_data = request.get('schedule_data', {})
        
        # Получаем ID периода
        cursor.execute('SELECT id FROM periods WHERE period = ?', (period,))
        period_row = cursor.fetchone()
        if not period_row:
            return {"status": "error", "message": "Period not found"}
        
        period_id = period_row['id']
        
        # Получаем всех сотрудников
        cursor.execute('SELECT id, name FROM employees')
        employees = {row['name']: row['id'] for row in cursor.fetchall()}
        
        # Сохраняем расписание
        for emp_name, schedule in schedule_data.get('schedule', {}).items():
            emp_id = employees.get(emp_name)
            if not emp_id:
                continue
            
            for day, status in enumerate(schedule):
                cursor.execute('''
                    INSERT OR REPLACE INTO schedule (period_id, employee_id, day, status)
                    VALUES (?, ?, ?, ?)
                ''', (period_id, emp_id, day, status))
        
        # Сохраняем примечания
        notes_data = schedule_data.get('notes', {})
        for emp_name, notes in notes_data.items():
            emp_id = employees.get(emp_name)
            if not emp_id:
                continue
            
            for day_str, note in notes.items():
                try:
                    day = int(day_str)
                    cursor.execute('''
                        INSERT OR REPLACE INTO notes (period_id, employee_id, day, end_time, worked_hours)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (period_id, emp_id, day, note.get('end_time'), note.get('worked_hours')))
                except ValueError:
                    continue
        
        conn.commit()
        return {"status": "success", "message": "Schedule saved"}
    
    # НОВЫЕ МЕТОДЫ ДЛЯ EXCEL ФАЙЛОВ
    
    def save_excel_file(self, cursor, conn, request):
        """Сохранить Excel файл на сервере"""
        file_name = request.get('file_name')
        file_data_base64 = request.get('file_data')
        created_by = request.get('created_by', 'Unknown')
        
        if not file_name or not file_data_base64:
            return {"status": "error", "message": "File name and data are required"}
        
        try:
            # Декодируем base64
            file_data = base64.b64decode(file_data_base64)
            file_size = len(file_data)
            
            cursor.execute('''
                INSERT INTO excel_files (file_name, file_data, file_size, created_by)
                VALUES (?, ?, ?, ?)
            ''', (file_name, file_data, file_size, created_by))
            
            conn.commit()
            return {"status": "success", "message": f"File {file_name} saved successfully"}
        except Exception as e:
            return {"status": "error", "message": f"Error saving file: {str(e)}"}
    
    def get_excel_file(self, cursor, request):
        """Получить Excel файл с сервера"""
        file_name = request.get('file_name')
        file_id = request.get('file_id')
        
        if file_id:
            cursor.execute('''
                SELECT file_name, file_data, created_at, created_by 
                FROM excel_files WHERE id = ?
            ''', (file_id,))
        else:
            cursor.execute('''
                SELECT file_name, file_data, created_at, created_by 
                FROM excel_files WHERE file_name = ? 
                ORDER BY created_at DESC LIMIT 1
            ''', (file_name,))
        
        row = cursor.fetchone()
        
        if row:
            file_data_base64 = base64.b64encode(row['file_data']).decode('utf-8')
            return {
                "status": "success", 
                "file_name": row['file_name'],
                "file_data": file_data_base64,
                "created_at": row['created_at'],
                "created_by": row['created_by']
            }
        else:
            return {"status": "error", "message": "File not found"}
    
    def get_excel_files_list(self, cursor):
        """Получить список доступных Excel файлов на сервере"""
        cursor.execute('''
            SELECT id, file_name, file_size, created_by, created_at 
            FROM excel_files 
            ORDER BY created_at DESC
        ''')
        files = []
        for row in cursor.fetchall():
            files.append({
                'id': row['id'],
                'file_name': row['file_name'],
                'file_size': row['file_size'],
                'created_by': row['created_by'],
                'created_at': row['created_at']
            })
        return {"status": "success", "data": files}
    
    def delete_excel_file(self, cursor, conn, request):
        """Удалить Excel файл с сервера"""
        file_id = request.get('file_id')
        
        if not file_id:
            return {"status": "error", "message": "File ID is required"}
        
        try:
            cursor.execute('DELETE FROM excel_files WHERE id = ?', (file_id,))
            conn.commit()
            
            if cursor.rowcount > 0:
                return {"status": "success", "message": "File deleted successfully"}
            else:
                return {"status": "error", "message": "File not found"}
        except Exception as e:
            return {"status": "error", "message": f"Error deleting file: {str(e)}"}
    
    def get_days_in_month(self, year, month):
        """Получить количество дней в месяце"""
        return monthrange(year, month)[1]
    
    def start(self):
        """Запуск сервера"""
        server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        
        try:
            server_socket.bind((self.host, self.port))
            server_socket.listen(5)
            print(f"Сервер запущен на {self.host}:{self.port}")
            
            while True:
                client_socket, address = server_socket.accept()
                client_thread = threading.Thread(
                    target=self.handle_client,
                    args=(client_socket, address)
                )
                client_thread.daemon = True
                client_thread.start()
                
        except KeyboardInterrupt:
            print("Остановка сервера...")
        finally:
            server_socket.close()

if __name__ == "__main__":
    server = ScheduleServer()
    server.start()