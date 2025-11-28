import sys
import json
import socket
import threading
import os
import sqlite3
from datetime import datetime
from collections import defaultdict
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
                             QPushButton, QVBoxLayout, QWidget, QHBoxLayout, QMenu,
                             QLabel, QMessageBox, QFileDialog, QHeaderView,
                             QAction, QComboBox, QInputDialog, QDialog, 
                             QVBoxLayout, QCheckBox, QScrollArea, QDialogButtonBox,
                             QShortcut, QTimeEdit, QFormLayout, QGridLayout, QSplitter,
                             QLineEdit, QMenuBar)
from PyQt5.QtCore import Qt, QSettings, QThread, pyqtSignal, QTime, QDate, QTimer
from PyQt5.QtGui import QColor, QKeySequence, QFont, QPainter, QIcon
from calendar import monthrange
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter


class NetworkManager:
    def __init__(self):
        self.server_host = 'localhost'
        self.server_port = 8888
        self.connected = False
        self.socket = None
        self.local_db = "local_schedule.db"
        self.init_pending_changes_table()
    
    def init_pending_changes_table(self):
        """Инициализация таблицы ожидающих изменений"""
        conn = sqlite3.connect(self.local_db)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pending_changes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                action TEXT NOT NULL,
                data TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def connect(self, host, port):
        """Подключение к серверу"""
        try:
            self.server_host = host
            self.server_port = port
            
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.settimeout(5)
            self.socket.connect((host, port))
            self.connected = True
            
            # Загружаем ожидающие изменения
            self.load_pending_changes()
            return True
        except Exception as e:
            print(f"Ошибка подключения: {e}")
            self.connected = False
            return False
    
    def disconnect(self):
        """Отключение от сервера"""
        if self.socket:
            self.socket.close()
        self.connected = False
    
    def send_request(self, request):
        """Отправка запроса на сервер"""
        if not self.connected:
            # Сохраняем запрос в SQLite
            self.save_pending_change(request)
            return {'status': 'offline', 'message': 'Сохранено локально'}
        
        try:
            self.socket.settimeout(10)
            self.socket.send(json.dumps(request).encode('utf-8'))
            
            response = self.socket.recv(10240).decode('utf-8')
            return json.loads(response)
        except Exception as e:
            self.connected = False
            self.save_pending_change(request)
            return {'status': 'error', 'message': str(e)}
    
    def save_pending_change(self, request):
        """Сохранение ожидающих изменений в SQLite"""
        conn = sqlite3.connect(self.local_db)
        cursor = conn.cursor()
        
        cursor.execute(
            "INSERT INTO pending_changes (action, data) VALUES (?, ?)",
            (request.get('action'), json.dumps(request))
        )
        
        conn.commit()
        conn.close()
    
    def load_pending_changes(self):
        """Загрузка и отправка ожидающих изменений из SQLite"""
        if not self.connected:
            return
        
        conn = sqlite3.connect(self.local_db)
        cursor = conn.cursor()
        
        cursor.execute("SELECT id, action, data FROM pending_changes ORDER BY created_at")
        changes = cursor.fetchall()
        
        successful_changes = []
        
        for change_id, action, data in changes:
            request = json.loads(data)
            response = self.send_request(request)
            
            if response.get('status') == 'success':
                successful_changes.append(change_id)
        
        # Удаляем успешно отправленные изменения
        if successful_changes:
            placeholders = ','.join('?' * len(successful_changes))
            cursor.execute(f"DELETE FROM pending_changes WHERE id IN ({placeholders})", successful_changes)
        
        conn.commit()
        conn.close()


class ScheduleManager:
    def __init__(self, network_manager):
        self.network_manager = network_manager
        self.local_db = "local_schedule.db"
        self.init_local_db()
        self._cache = {}
        self._employees = []  # Кэш сотрудников
        self._periods = []    # Кэш периодов
    
    @property
    def employees(self):
        """Свойство для совместимости со старым кодом"""
        if not self._employees:
            self._employees = self.load_employees()
        return self._employees
    
    @property
    def periods(self):
        """Свойство для периодов"""
        if not self._periods:
            self._periods = self.get_periods()
        return self._periods
    
    def init_local_db(self):
        """Инициализация локальной SQLite базы данных"""
        conn = sqlite3.connect(self.local_db)
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
        
        # Таблица периодов
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS periods (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                period TEXT UNIQUE NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Таблица расписаний
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS schedules (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                period_id INTEGER,
                employee_name TEXT NOT NULL,
                day_index INTEGER NOT NULL,
                status INTEGER DEFAULT 4,
                end_time TEXT,
                worked_hours REAL,
                FOREIGN KEY (period_id) REFERENCES periods (id),
                UNIQUE(period_id, employee_name, day_index)
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def load_employees(self):
        """Загрузка сотрудников с сервера или из локальной БД"""
        # Пытаемся загрузить с сервера
        response = self.network_manager.send_request({'action': 'get_employees'})
        
        if response.get('status') == 'success':
            # Сохраняем в локальную БД
            self._save_employees_to_local(response['data'])
            self._employees = response['data']  # Обновляем кэш
            return response['data']
        else:
            # Загрузка из локальной БД
            employees = self._load_employees_from_local()
            self._employees = employees  # Обновляем кэш
            return employees
    
    def _save_employees_to_local(self, employees):
        """Сохранение сотрудников в локальную БД"""
        conn = sqlite3.connect(self.local_db)
        cursor = conn.cursor()
        
        try:
            for emp in employees:
                cursor.execute('''
                    INSERT OR REPLACE INTO employees (name, position) 
                    VALUES (?, ?)
                ''', (emp['name'], emp.get('position', '')))
            
            conn.commit()
        except Exception as e:
            print(f"Ошибка сохранения сотрудников: {e}")
        finally:
            conn.close()
    
    def _load_employees_from_local(self):
        """Загрузка сотрудников из локальной БД"""
        conn = sqlite3.connect(self.local_db)
        cursor = conn.cursor()
        
        try:
            cursor.execute("SELECT name, position FROM employees")
            employees = [{'name': row[0], 'position': row[1]} for row in cursor.fetchall()]
            return employees
        except Exception as e:
            print(f"Ошибка загрузки сотрудников: {e}")
            return []
        finally:
            conn.close()
    
    def get_periods(self):
        """Получение периодов с сервера или из локальной БД"""
        response = self.network_manager.send_request({'action': 'get_periods'})
        
        if response.get('status') == 'success':
            # Сохраняем в локальную БД
            self._save_periods_to_local(response['data'])
            self._periods = response['data']  # Обновляем кэш
            return response['data']
        else:
            # Загрузка из локальной БД
            periods = self._load_periods_from_local()
            self._periods = periods  # Обновляем кэш
            return periods
    
    def _save_periods_to_local(self, periods):
        """Сохранение периодов в локальную БД"""
        conn = sqlite3.connect(self.local_db)
        cursor = conn.cursor()
        
        try:
            for period in periods:
                cursor.execute('''
                    INSERT OR IGNORE INTO periods (period) VALUES (?)
                ''', (period,))
            
            conn.commit()
        except Exception as e:
            print(f"Ошибка сохранения периодов: {e}")
        finally:
            conn.close()
    
    def _load_periods_from_local(self):
        """Загрузка периодов из локальной БД"""
        conn = sqlite3.connect(self.local_db)
        cursor = conn.cursor()
        
        try:
            cursor.execute("SELECT period FROM periods ORDER BY period")
            periods = [row[0] for row in cursor.fetchall()]
            return periods
        except Exception as e:
            print(f"Ошибка загрузки периодов: {e}")
            return []
        finally:
            conn.close()
    
    def load_schedule(self, period):
        """Загрузка расписания с сервера или из локальной БД"""
        if period in self._cache:
            return self._cache[period]
        
        response = self.network_manager.send_request({
            'action': 'get_schedule', 
            'period': period
        })
        
        if response.get('status') == 'success':
            data = response['data']
            if 'notes' not in data:
                data['notes'] = {}
            
            # Сохраняем в локальную БД
            self._save_schedule_to_local(period, data)
            self._cache[period] = data
            return data
        else:
            # Загрузка из локальной БД
            data = self._load_schedule_from_local(period)
            if data:
                self._cache[period] = data
            return data
    
    def _save_schedule_to_local(self, period, data):
        """Сохранение расписания в локальную БД"""
        conn = sqlite3.connect(self.local_db)
        cursor = conn.cursor()
        
        try:
            # Получаем или создаем период
            cursor.execute("SELECT id FROM periods WHERE period = ?", (period,))
            result = cursor.fetchone()
            
            if result:
                period_id = result[0]
            else:
                cursor.execute("INSERT INTO periods (period) VALUES (?)", (period,))
                period_id = cursor.lastrowid
            
            # Сохраняем расписание
            for emp_name, schedule in data['schedule'].items():
                for day_index, status in enumerate(schedule):
                    notes_data = data['notes'].get(emp_name, {}).get(str(day_index), {})
                    
                    cursor.execute('''
                        INSERT OR REPLACE INTO schedules 
                        (period_id, employee_name, day_index, status, end_time, worked_hours)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (
                        period_id, emp_name, day_index, status,
                        notes_data.get('end_time'), notes_data.get('worked_hours')
                    ))
            
            conn.commit()
        except Exception as e:
            print(f"Ошибка сохранения расписания: {e}")
            conn.rollback()
        finally:
            conn.close()
    
    def _load_schedule_from_local(self, period):
        """Загрузка расписания из локальной БД"""
        conn = sqlite3.connect(self.local_db)
        cursor = conn.cursor()
        
        try:
            # Получаем ID периода
            cursor.execute("SELECT id FROM periods WHERE period = ?", (period,))
            result = cursor.fetchone()
            
            if not result:
                return None
            
            period_id = result[0]
            
            # Получаем сотрудников периода
            cursor.execute('''
                SELECT DISTINCT employee_name FROM schedules 
                WHERE period_id = ? ORDER BY employee_name
            ''', (period_id,))
            employees = [{'name': row[0]} for row in cursor.fetchall()]
            
            if not employees:
                return None
            
            # Получаем расписание
            cursor.execute('''
                SELECT employee_name, day_index, status, end_time, worked_hours 
                FROM schedules WHERE period_id = ?
            ''', (period_id,))
            
            schedule = {}
            notes = {}
            
            for row in cursor.fetchall():
                emp_name, day_index, status, end_time, worked_hours = row
                
                if emp_name not in schedule:
                    schedule[emp_name] = []
                
                # Заполняем массив до нужного размера
                while len(schedule[emp_name]) <= day_index:
                    schedule[emp_name].append(4)  # Статус по умолчанию
                
                schedule[emp_name][day_index] = status
                
                # Обработка примечаний
                if end_time or worked_hours:
                    if emp_name not in notes:
                        notes[emp_name] = {}
                    notes[emp_name][str(day_index)] = {
                        'end_time': end_time,
                        'worked_hours': worked_hours
                    }
            
            return {
                'employees': employees,
                'schedule': schedule,
                'notes': notes
            }
            
        except Exception as e:
            print(f"Ошибка загрузки расписания: {e}")
            return None
        finally:
            conn.close()
    
    def save_schedule(self, period, data):
        """Сохранение расписания на сервер и в локальную БД"""
        # Всегда сохраняем в локальную БД
        self._save_schedule_to_local(period, data)
        self._cache[period] = data
        
        # Отправляем на сервер
        response = self.network_manager.send_request({
            'action': 'save_schedule',
            'period': period,
            'data': data
        })
        
        return response.get('status') == 'success' or response.get('status') == 'offline'
    
    def add_employee(self, name):
        """Добавление сотрудника"""
        # Проверяем локально
        employees = self._load_employees_from_local()
        if any(emp["name"].lower() == name.lower() for emp in employees):
            return False
        
        # Добавляем в локальную БД
        conn = sqlite3.connect(self.local_db)
        cursor = conn.cursor()
        
        try:
            cursor.execute("INSERT INTO employees (name) VALUES (?)", (name,))
            conn.commit()
            # Обновляем кэш
            self._employees = self._load_employees_from_local()
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()
        
        # Отправляем на сервер
        response = self.network_manager.send_request({
            'action': 'add_employee',
            'name': name
        })
        
        return True
    
    def create_period(self, period, employees):
        """Создание периода"""
        # Создаем локальную структуру данных
        try:
            year, month = map(int, period.split('-'))
            days_in_month = monthrange(year, month)[1]
        except:
            days_in_month = 31
        
        schedule_data = {
            "employees": [{"name": name} for name in employees],
            "schedule": {name: [4] * days_in_month for name in employees},
            "notes": {}
        }
        
        # Сохраняем в локальную БД
        self._save_schedule_to_local(period, schedule_data)
        self._cache[period] = schedule_data
        
        # Обновляем кэш периодов
        if period not in self._periods:
            self._periods.append(period)
            self._periods.sort()
        
        # Отправляем на сервер
        response = self.network_manager.send_request({
            'action': 'create_period',
            'period': period,
            'employees': employees
        })
        
        return response.get('status') == 'success' or response.get('status') == 'offline'
    
    def add_employees_to_period(self, period, employees):
        """Добавление сотрудников в период"""
        # Загружаем текущее расписание
        current_data = self.load_schedule(period)
        if not current_data:
            return False
        
        # Добавляем новых сотрудников
        for emp_name in employees:
            if emp_name not in [emp["name"] for emp in current_data["employees"]]:
                current_data["employees"].append({"name": emp_name})
                current_data["schedule"][emp_name] = [4] * 31  # 31 день по умолчанию
        
        # Сохраняем в локальную БД
        self._save_schedule_to_local(period, current_data)
        self._cache[period] = current_data
        
        # Отправляем на сервер
        response = self.network_manager.send_request({
            'action': 'add_employees_to_period',
            'period': period,
            'employees': employees
        })
        
        return response.get('status') == 'success' or response.get('status') == 'offline'
    
    def sync_with_server(self):
        """Синхронизация данных с сервером"""
        if not self.network_manager.connected:
            return False
        
        # Собираем все локальные данные для синхронизации
        sync_data = {
            'employees': self._load_employees_from_local(),
            'periods': self._load_periods_from_local(),
            'schedules': []
        }
        
        # Добавляем расписания всех периодов
        for period in sync_data['periods']:
            schedule_data = self.load_schedule(period)
            if schedule_data:
                sync_data['schedules'].append({
                    'period': period,
                    'schedule': schedule_data['schedule'],
                    'notes': schedule_data.get('notes', {})
                })
        
        # Отправляем на сервер
        response = self.network_manager.send_request({
            'action': 'sync_data',
            'data': sync_data
        })
        
        return response.get('status') == 'success'


class ConnectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройка подключения")
        self.setModal(True)
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        form_layout = QFormLayout()
        
        self.host_edit = QLineEdit()
        self.host_edit.setText("localhost")
        form_layout.addRow("Адрес сервера:", self.host_edit)
        
        self.port_edit = QLineEdit()
        self.port_edit.setText("8888")
        form_layout.addRow("Порт:", self.port_edit)
        
        self.status_label = QLabel("Не подключено")
        self.status_label.setStyleSheet("color: red; font-weight: bold;")
        form_layout.addRow("Статус:", self.status_label)
        
        layout.addLayout(form_layout)
        
        buttons_layout = QHBoxLayout()
        self.test_btn = QPushButton("Тест подключения")
        self.test_btn.clicked.connect(self.test_connection)
        buttons_layout.addWidget(self.test_btn)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        buttons_layout.addWidget(button_box)
        
        layout.addLayout(buttons_layout)
    
    def test_connection(self):
        """Тестирование подключения"""
        host = self.host_edit.text()
        try:
            port = int(self.port_edit.text())
        except ValueError:
            self.status_label.setText("Неверный порт")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
            return
        
        try:
            socket.create_connection((host, port), timeout=5)
            self.status_label.setText("Подключено")
            self.status_label.setStyleSheet("color: green; font-weight: bold;")
        except:
            self.status_label.setText("Ошибка подключения")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
    
    def get_connection_params(self):
        return self.host_edit.text(), int(self.port_edit.text())


class MonthSelectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent) 
        self.setWindowTitle("Выбор месяца и года")
        self.setModal(True)
        self.setMinimumWidth(300)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        grid_layout = QGridLayout()
        
        # Текущая дата для значений по умолчанию
        current_date = QDate.currentDate()
        
        # Год
        grid_layout.addWidget(QLabel("Год:"), 0, 0)
        self.year_combo = QComboBox()
        current_year = current_date.year()
        for year in range(current_year - 5, current_year + 6):
            self.year_combo.addItem(str(year), year)
        self.year_combo.setCurrentText(str(current_year))
        grid_layout.addWidget(self.year_combo, 0, 1)
        
        # Месяц
        grid_layout.addWidget(QLabel("Месяц:"), 1, 0)
        self.month_combo = QComboBox()
        month_names = {
            1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
            5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
            9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
        }
        for month_num, month_name in month_names.items():
            self.month_combo.addItem(month_name, month_num)
        self.month_combo.setCurrentIndex(current_date.month() - 1)
        grid_layout.addWidget(self.month_combo, 1, 1)
        
        layout.addLayout(grid_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def get_selected_period(self):
        year = self.year_combo.currentData()
        month = self.month_combo.currentData()
        return f"{year}-{month:02d}"


class NoteDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Примечание к рабочему дню")
        self.setModal(True)
        self.setMinimumWidth(300)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        form_layout = QFormLayout()
        
        self.end_time_edit = QTimeEdit()
        self.end_time_edit.setTime(QTime(20, 0))
        self.end_time_edit.setDisplayFormat("HH:mm")
        
        form_layout.addRow("Время окончания:", self.end_time_edit)
        
        layout.addLayout(form_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def get_end_time(self):
        return self.end_time_edit.time().toString("HH:mm")


class NoteTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_widget = parent  # Сохраняем ссылку на родительский виджет
        self.is_loading = False
    
    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self.viewport())
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor(255, 0, 0))
        painter.end()
    
    def setItem(self, row, column, item):
        super().setItem(row, column, item)
        if not self.is_loading and column < self.columnCount() - 4:
            self.parent_widget.update_counters()
            # Исправляем обращение к главному окну
            if hasattr(self.parent_widget, 'parent') and hasattr(self.parent_widget.parent, 'set_unsaved_changes'):
                self.parent_widget.parent.set_unsaved_changes(True)


class NoteItem(QTableWidgetItem):
    def __init__(self, text="", has_note=False, note_data=None):
        super().__init__(text)
        self.has_note = has_note
        self.note_data = note_data or {}
    
    def clone(self):
        return NoteItem(self.text(), self.has_note, self.note_data.copy())


class MonthWidget(QWidget):
    def __init__(self, parent=None, db=None):
        super().__init__(parent)
        self.parent_app = parent  # Сохраняем ссылку на главное приложение
        self.db = db
        self.current_period = None
        self.initUI()
    
    def initUI(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        self.setLayout(layout)
        
        # Верхняя панель с выбором месяца
        top_layout = QHBoxLayout()
        top_layout.setContentsMargins(5, 2, 5, 2)
        
        self.period_combo = QComboBox()
        
        # УВЕЛИЧИВАЕМ ШИРИНУ COMBOBOX
        self.period_combo.setMinimumWidth(300)
        self.period_combo.setFixedHeight(28)
        
        # Стили для комбобокса
        self.period_combo.setStyleSheet("""
            QComboBox {
                min-width: 300px;
                height: 28px;
                font-size: 12px;
                padding: 4px;
                border: 1px solid #ccc;
                border-radius: 4px;
            }
            QComboBox::drop-down {
                width: 25px;
                border: none;
            }
            QComboBox QAbstractItemView {
                min-width: 350px;
                font-size: 12px;
            }
        """)
        
        self.period_combo.currentIndexChanged.connect(self.period_changed)
        
        month_label = QLabel("Месяц:")
        month_label_font = month_label.font()
        month_label_font.setPointSize(11)
        month_label.setFont(month_label_font)
        
        top_layout.addWidget(month_label)
        top_layout.addWidget(self.period_combo)
        
        top_layout.addStretch()
        layout.addLayout(top_layout)
        
        # Таблица
        self.table = NoteTableWidget(self)  # Передаем self как parent
        self.setup_table(self.table)
        layout.addWidget(self.table)
    
    def setup_table(self, table):
        table.setSelectionMode(QTableWidget.ExtendedSelection)
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(self.show_context_menu)
        
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        table.verticalHeader().setDefaultSectionSize(20)
        
        font = table.horizontalHeader().font()
        font.setBold(True)
        font.setPointSize(12)
        table.horizontalHeader().setFont(font)
        
        font = table.verticalHeader().font()
        font.setBold(True)
        font.setPointSize(12)
        table.verticalHeader().setFont(font)
    
    def load_periods(self):
        periods = self.db.periods
        self.period_combo.clear()
        
        if periods:
            for period in periods:
                try:
                    year, month = map(int, period.split('-'))
                    month_name = f"{self.parent_app.month_names[month]} {year}"
                    self.period_combo.addItem(month_name, period)
                except:
                    self.period_combo.addItem(period)
            
            # Автоматически выбираем первый период, если он не выбран
            if self.period_combo.count() > 0 and not self.current_period:
                self.period_combo.setCurrentIndex(0)
    
    def period_changed(self, index):
        if index >= 0:
            period = self.period_combo.itemData(index)
            self.current_period = period
            self.load_data(period)
    
    def get_working_days_count(self, year, month):
        """Возвращает количество рабочих дней в месяце (без воскресений)"""
        days_in_month = monthrange(year, month)[1]
        working_days = 0
        
        for day in range(1, days_in_month + 1):
            if datetime(year, month, day).weekday() != 6:
                working_days += 1
                
        return working_days
    
    def get_day_mapping(self, year, month):
        """Создает маппинг между номером столбца и фактическим днем месяца (исключая воскресенья)"""
        days_in_month = monthrange(year, month)[1]
        day_mapping = []
        
        for day in range(1, days_in_month + 1):
            if datetime(year, month, day).weekday() != 6:
                day_mapping.append(day)
                
        return day_mapping
    
    def load_data(self, period):
        if not period:
            return
            
        self.table.is_loading = True
        
        self.table.clear()
        self.current_period = period
            
        schedule_data = self.db.load_schedule(period)
        
        if not schedule_data:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self.table.is_loading = False
            return
        
        employees_in_period = schedule_data.get("employees", [])
        if not employees_in_period:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self.table.is_loading = False
            return
        
        try:
            year, month = map(int, period.split('-'))
            working_days_count = self.get_working_days_count(year, month)
            day_mapping = self.get_day_mapping(year, month)
            days_in_month = monthrange(year, month)[1]
        except:
            days_in_month = 31
            working_days_count = days_in_month
            day_mapping = list(range(1, days_in_month + 1))
        
        self.table.setRowCount(len(employees_in_period))
        self.table.setColumnCount(working_days_count + 4)
        
        self.table.setVerticalHeaderLabels([emp["name"] for emp in employees_in_period])
        
        notes_data = schedule_data.get('notes', {})
        
        for row, emp in enumerate(employees_in_period):
            emp_name = emp["name"]
            emp_schedule = schedule_data['schedule'].get(emp_name, [4]*days_in_month)  
            emp_notes = notes_data.get(emp_name, {})
            
            for col, actual_day in enumerate(day_mapping):
                day_index = actual_day - 1
                day_status = emp_schedule[day_index]
                icon, _, bg_color, _ = self.parent_app.status_mapping.get(day_status, ("", "Пусто", QColor(Qt.white), 0))
                
                note_key = str(day_index)
                has_note = note_key in emp_notes
                note_data = emp_notes.get(note_key, {})
                
                item = NoteItem(icon, has_note, note_data)
                item.setTextAlignment(Qt.AlignCenter)
                item.setBackground(bg_color)
                item.setForeground(Qt.black)
                
                if has_note:
                    end_time = note_data.get('end_time', '20:00')
                    item.setText(end_time)
                    item.setData(Qt.UserRole, True)
                    item.setToolTip(f"Отработано: {note_data.get('worked_hours', 12)} часов")
                else:
                    item.setText("")
                
                self.table.setItem(row, col, item)
        
        self.table.is_loading = False
        self.update_counters()
        
        fixed_day_width = 25
        
        for col in range(working_days_count):
            self.table.setColumnWidth(col, fixed_day_width)
        
        self.table.setColumnWidth(working_days_count, 70)
        self.table.setColumnWidth(working_days_count + 1, 70)
        self.table.setColumnWidth(working_days_count + 2, 70)
        self.table.setColumnWidth(working_days_count + 3, 80)
        
        for col in range(working_days_count + 4):
            self.table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Fixed)
        
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.viewport().update()

    def hours_to_hours_minutes(self, total_hours):
        """Конвертирует дробное количество часов в часы и минуты"""
        hours = int(total_hours)
        minutes = int(round((total_hours - hours) * 60))
        
        if minutes == 60:
            hours += 1
            minutes = 0
            
        return hours, minutes

    def update_counters(self):
        if not self.current_period or self.table.is_loading:
            return
        
        try:
            year, month = map(int, self.current_period.split('-'))
            working_days_count = self.get_working_days_count(year, month)
            day_mapping = self.get_day_mapping(year, month)
            days_in_month = monthrange(year, month)[1]
        except:
            working_days_count = self.table.columnCount() - 4
            day_mapping = list(range(1, working_days_count + 1))
            days_in_month = working_days_count
        
        working_counts = [0] * working_days_count
        call_center_counts = [0] * working_days_count
        registry_counts = [0] * working_days_count
        
        for row in range(self.table.rowCount()):
            total_shifts = 0
            call_center_days = 0
            registry_days = 0
            total_hours = 0.0
            
            for col in range(working_days_count):
                item = self.table.item(row, col)
                if item:
                    bg_color = item.background().color()
                    
                    if bg_color == self.parent_app.status_mapping[0][2]:
                        call_center_days += 1
                        total_shifts += 1
                        working_counts[col] += 1
                        call_center_counts[col] += 1
                        
                        if isinstance(item, NoteItem) and item.has_note:
                            total_hours += item.note_data.get('worked_hours', 12)
                        else:
                            total_hours += self.parent_app.status_mapping[0][3]
                            
                    elif bg_color == self.parent_app.status_mapping[1][2]:
                        registry_days += 1
                        total_shifts += 1
                        working_counts[col] += 1
                        registry_counts[col] += 1
                        
                        if isinstance(item, NoteItem) and item.has_note:
                            total_hours += item.note_data.get('worked_hours', 12)
                        else:
                            total_hours += self.parent_app.status_mapping[1][3]
            
            hours, minutes = self.hours_to_hours_minutes(total_hours)
            
            if working_days_count < self.table.columnCount():
                total_shifts_item = self.table.item(row, working_days_count)
                if not total_shifts_item:
                    total_shifts_item = QTableWidgetItem()
                    total_shifts_item.setTextAlignment(Qt.AlignCenter)
                    total_shifts_item.setFlags(total_shifts_item.flags() ^ Qt.ItemIsEditable)
                    total_shifts_item.setFont(QFont("Arial", 11, QFont.Bold))
                    self.table.setItem(row, working_days_count, total_shifts_item)
                total_shifts_item.setText(f"{total_shifts}")
            
            if working_days_count + 1 < self.table.columnCount():
                registry_item = self.table.item(row, working_days_count + 1)
                if not registry_item:
                    registry_item = QTableWidgetItem()
                    registry_item.setTextAlignment(Qt.AlignCenter)
                    registry_item.setFlags(registry_item.flags() ^ Qt.ItemIsEditable)
                    registry_item.setFont(QFont("Arial", 11, QFont.Bold))
                    self.table.setItem(row, working_days_count + 1, registry_item)
                registry_item.setText(f"{registry_days}")
            
            if working_days_count + 2 < self.table.columnCount():
                call_center_item = self.table.item(row, working_days_count + 2)
                if not call_center_item:
                    call_center_item = QTableWidgetItem()
                    call_center_item.setTextAlignment(Qt.AlignCenter)
                    call_center_item.setFlags(call_center_item.flags() ^ Qt.ItemIsEditable)
                    call_center_item.setFont(QFont("Arial", 11, QFont.Bold))
                    self.table.setItem(row, working_days_count + 2, call_center_item)
                call_center_item.setText(f"{call_center_days}")
            
            if working_days_count + 3 < self.table.columnCount():
                hours_item = self.table.item(row, working_days_count + 3)
                if not hours_item:
                    hours_item = QTableWidgetItem()
                    hours_item.setTextAlignment(Qt.AlignCenter)
                    hours_item.setFlags(hours_item.flags() ^ Qt.ItemIsEditable)
                    hours_item.setFont(QFont("Arial", 10, QFont.Bold))
                    self.table.setItem(row, working_days_count + 3, hours_item)
                
                hours_text = f"{hours}ч {minutes}м"
                hours_item.setText(hours_text)
        
        day_headers = []
        for col, actual_day in enumerate(day_mapping):
            try:
                day_of_week = (datetime(year, month, actual_day).weekday())
                day_name = self.parent_app.day_names[day_of_week]
                count = working_counts[col]
                day_headers.append(f"{actual_day}\n{day_name}\n({count})")
            except:
                day_headers.append(str(actual_day))
        
        headers = day_headers + ["Смены", "Рег", "КЦ", "Часы"]
        self.table.setHorizontalHeaderLabels(headers)

    def show_context_menu(self, pos):
        table = self.table
        selected = table.selectedIndexes()
        if not selected:
            return
        
        menu = QMenu()
        
        copy_action = QAction("Копировать", self)
        copy_action.triggered.connect(self.copy_selected)
        menu.addAction(copy_action)
        
        paste_action = QAction("Вставить", self)
        paste_action.triggered.connect(self.paste_selected)
        menu.addAction(paste_action)
        
        menu.addSeparator()
        
        if len(selected) == 1:
            index = selected[0]
            if index.column() < table.columnCount() - 4:
                item = table.item(index.row(), index.column())
                if item and item.background().color() in [self.parent_app.status_mapping[0][2], self.parent_app.status_mapping[1][2]]:
                    if isinstance(item, NoteItem) and item.has_note:
                        remove_note_action = QAction("Удалить примечание", self)
                        remove_note_action.triggered.connect(lambda: self.remove_note(index))
                        menu.addAction(remove_note_action)
                        menu.addSeparator()
                    
                    note_action = QAction("Примечание", self)
                    note_action.triggered.connect(lambda: self.add_note(index))
                    menu.addAction(note_action)
                    menu.addSeparator()
        
        for status, (icon, text, _, _) in self.parent_app.status_mapping.items():
            action = QAction(f"■ {text}", self)
            action.triggered.connect(lambda _, s=status: self.update_selected_status(s))
            menu.addAction(action)
        
        menu.exec_(table.viewport().mapToGlobal(pos))
    
    def add_note(self, index):
        item = self.table.item(index.row(), index.column())
        if not item or item.background().color() not in [self.parent_app.status_mapping[0][2], self.parent_app.status_mapping[1][2]]:
            return
        
        dialog = NoteDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            end_time = dialog.get_end_time()
            
            end_hour = int(end_time.split(':')[0])
            minutes = int(end_time.split(':')[1])
            worked_hours = end_hour - 8 + (minutes / 60.0)
            
            if isinstance(item, NoteItem):
                item.has_note = True
                item.note_data = {
                    'end_time': end_time,
                    'worked_hours': round(worked_hours, 2)
                }
                item.setText(end_time)
                item.setData(Qt.UserRole, True)
                item.setToolTip(f"Отработано: {round(worked_hours, 2)} часов")
                self.table.viewport().update()
                self.update_counters()
                # Исправляем обращение к главному окну
                if hasattr(self, 'parent_app') and hasattr(self.parent_app, 'set_unsaved_changes'):
                    self.parent_app.set_unsaved_changes(True)
    
    def remove_note(self, index):
        item = self.table.item(index.row(), index.column())
        if item and isinstance(item, NoteItem):
            item.has_note = False
            item.note_data = {}
            item.setText("")
            item.setData(Qt.UserRole, False)
            item.setToolTip("")
            self.table.viewport().update()
            self.update_counters()
            # Исправляем обращение к главному окну
            if hasattr(self, 'parent_app') and hasattr(self.parent_app, 'set_unsaved_changes'):
                self.parent_app.set_unsaved_changes(True)
    
    def copy_selected(self):
        selected = self.table.selectedIndexes()
        if not selected:
            return
            
        rows = defaultdict(list)
        for index in selected:
            rows[index.row()].append(index.column())
        
        self.copied_data = []
        for row, cols in rows.items():
            row_data = []
            for col in cols:
                item = self.table.item(row, col)
                bg_color = item.background().color() if item else QColor(Qt.white)
                status = next((k for k, v in self.parent_app.status_mapping.items() if v[2] == bg_color), 4)
                row_data.append(status)
            self.copied_data.append(row_data)
    
    def paste_selected(self):
        if not hasattr(self, 'copied_data') or not self.copied_data:
            return
            
        selected = self.table.selectedIndexes()
        if not selected:
            return
            
        min_row = min(index.row() for index in selected)
        min_col = min(index.column() for index in selected)
        
        for row_offset, row_data in enumerate(self.copied_data):
            for col_offset, status in enumerate(row_data):
                target_row = min_row + row_offset
                target_col = min_col + col_offset
                
                if target_row >= self.table.rowCount() or target_col >= self.table.columnCount() - 4:
                    continue
                
                item = self.table.item(target_row, target_col)
                if item and target_col < self.table.columnCount() - 4:
                    icon, _, bg_color, _ = self.parent_app.status_mapping.get(status, ("", "Пусто", QColor(Qt.white), 0))
                    
                    old_has_note = False
                    old_note_data = {}
                    if isinstance(item, NoteItem):
                        old_has_note = item.has_note
                        old_note_data = item.note_data
                    
                    new_item = NoteItem("", old_has_note, old_note_data)
                    new_item.setTextAlignment(Qt.AlignCenter)
                    new_item.setBackground(bg_color)
                    new_item.setForeground(Qt.black)
                    
                    if old_has_note:
                        new_item.setText(old_note_data.get('end_time', '20:00'))
                        new_item.setData(Qt.UserRole, True)
                        new_item.setToolTip(f"Отработано: {old_note_data.get('worked_hours', 12)} часов")
                    else:
                        new_item.setText("")
                    
                    self.table.setItem(target_row, target_col, new_item)
        
        self.update_counters()
        # Исправляем обращение к главному окну
        if hasattr(self, 'parent_app') and hasattr(self.parent_app, 'set_unsaved_changes'):
            self.parent_app.set_unsaved_changes(True)
    
    def update_selected_status(self, status):
        selected = self.table.selectedIndexes()
        if not selected:
            return
        
        icon, _, bg_color, _ = self.parent_app.status_mapping[status]
        
        for index in selected:
            if index.column() >= self.table.columnCount() - 4:
                continue
            
            item = self.table.item(index.row(), index.column())
            if item:
                old_has_note = False
                old_note_data = {}
                if isinstance(item, NoteItem):
                    old_has_note = item.has_note
                    old_note_data = item.note_data
                
                new_item = NoteItem("", old_has_note, old_note_data)
                new_item.setTextAlignment(Qt.AlignCenter)
                new_item.setBackground(bg_color)
                new_item.setForeground(Qt.black)
                
                if old_has_note:
                    new_item.setText(old_note_data.get('end_time', '20:00'))
                    new_item.setData(Qt.UserRole, True)
                    new_item.setToolTip(f"Отработано: {old_note_data.get('worked_hours', 12)} часов")
                else:
                    new_item.setText("")
                
                self.table.setItem(index.row(), index.column(), new_item)
        
        self.update_counters()
        # Исправляем обращение к главному окну
        if hasattr(self, 'parent_app') and hasattr(self.parent_app, 'set_unsaved_changes'):
            self.parent_app.set_unsaved_changes(True)


class EmployeeSelectionDialog(QDialog):
    def __init__(self, employees, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Выбор сотрудников")
        self.setModal(True)
        self.setMinimumWidth(350)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        if not employees:
            layout.addWidget(QLabel("Нет доступных сотрудников"))
            button_box = QDialogButtonBox(QDialogButtonBox.Ok)
            button_box.accepted.connect(self.reject)
            layout.addWidget(button_box)
            return
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.content = QWidget()
        self.scroll_layout = QVBoxLayout(self.content)
        
        self.checkboxes = []
        for emp in employees:
            cb = QCheckBox(emp["name"])
            cb.setChecked(True)
            self.checkboxes.append((emp["name"], cb))
            self.scroll_layout.addWidget(cb)
        
        scroll.setWidget(self.content)
        layout.addWidget(scroll)
        
        btn_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("Выбрать все")
        self.select_all_btn.clicked.connect(lambda: self.set_all_checkboxes(True))
        btn_layout.addWidget(self.select_all_btn)
        
        self.deselect_all_btn = QPushButton("Снять все")
        self.deselect_all_btn.clicked.connect(lambda: self.set_all_checkboxes(False))
        btn_layout.addWidget(self.deselect_all_btn)
        
        layout.addLayout(btn_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def set_all_checkboxes(self, state):
        for _, cb in self.checkboxes:
            cb.setChecked(state)
    
    def get_selected_employees(self):
        return [name for name, cb in self.checkboxes if cb.isChecked()]


class AddEmployeeToPeriodDialog(QDialog):
    def __init__(self, all_employees, current_employees, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Добавить сотрудников в месяц")
        self.setModal(True)
        self.setMinimumWidth(350)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        current_employee_names = [emp["name"] for emp in current_employees]
        available_employees = [emp for emp in all_employees if emp["name"] not in current_employee_names]
        
        if not available_employees:
            layout.addWidget(QLabel("Нет доступных сотрудников для добавления"))
            button_box = QDialogButtonBox(QDialogButtonBox.Ok)
            button_box.accepted.connect(self.reject)
            layout.addWidget(button_box)
            return
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.content = QWidget()
        self.scroll_layout = QVBoxLayout(self.content)
        
        self.checkboxes = []
        for emp in available_employees:
            cb = QCheckBox(emp["name"])
            cb.setChecked(True)
            self.checkboxes.append((emp["name"], cb))
            self.scroll_layout.addWidget(cb)
        
        scroll.setWidget(self.content)
        layout.addWidget(scroll)
        
        btn_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("Выбрать все")
        self.select_all_btn.clicked.connect(lambda: self.set_all_checkboxes(True))
        btn_layout.addWidget(self.select_all_btn)
        
        self.deselect_all_btn = QPushButton("Снять все")
        self.deselect_all_btn.clicked.connect(lambda: self.set_all_checkboxes(False))
        btn_layout.addWidget(self.deselect_all_btn)
        
        layout.addLayout(btn_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def set_all_checkboxes(self, state):
        for _, cb in self.checkboxes:
            cb.setChecked(state)
    
    def get_selected_employees(self):
        return [name for name, cb in self.checkboxes if cb.isChecked()]


class ScheduleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("MyCompany", "WorkSchedule")
        
        # Инициализация сетевого менеджера
        self.network_manager = NetworkManager()
        self.db = ScheduleManager(self.network_manager)
        
        self.export_folder = self.settings.value("export_folder", "")
        
        # Настройки подключения
        self.server_host = self.settings.value("server_host", "localhost")
        self.server_port = int(self.settings.value("server_port", "8888"))
        
        # Таймер для автосохранения
        self.auto_save_timer = QTimer()
        self.auto_save_timer.setSingleShot(True)
        self.auto_save_timer.timeout.connect(self.auto_save)
        self.last_change_time = None
        
        # Статус несохраненных изменений
        self.unsaved_changes = False
        
        self.status_mapping = {
            0: ("", "Колл-центр", QColor("#77FF77"), 12),
            1: ("", "Регистратура", QColor("#7777FF"), 12),
            2: ("", "Не работает", QColor("#FF7777"), 0),
            3: ("", "Отпуск", QColor("#FFFF77"), 0),
            4: ("", "Пусто", QColor("#FFFFFF"), 0)
        }
        
        self.month_names = {
            1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
            5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
            9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
        }
        
        self.day_names = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
        
        self.initUI()
        self.init_shortcuts()
        
        # Автоподключение при запуске
        QTimer.singleShot(100, self.auto_connect)
    
    def initUI(self):
        self.setWindowTitle("Рабочее расписание")
        self.setGeometry(100, 100, 1600, 800)
        self.setWindowIcon(QIcon('schedule.ico'))
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(5, 5, 5, 5)
        self.layout.setSpacing(5)
        self.central_widget.setLayout(self.layout)
        
        # Добавляем меню подключения
        menubar = self.menuBar()
        connection_menu = menubar.addMenu('Подключение')
        
        connection_action = QAction('Настройки подключения', self)
        connection_action.triggered.connect(self.show_connection_dialog)
        connection_menu.addAction(connection_action)
        
        # Статус подключения в тулбаре
        self.connection_status = QLabel("Не подключено")
        self.connection_status.setStyleSheet("color: red; font-weight: bold; padding: 2px;")
        
        # Статус сохранения в тулбаре
        self.save_status = QLabel("Сохранено")
        self.save_status.setStyleSheet("color: green; font-weight: bold; padding: 2px;")
        
        # Добавляем виджеты в статусбар
        self.statusBar().addPermanentWidget(self.connection_status)
        self.statusBar().addPermanentWidget(self.save_status)
        
        # Верхняя панель с кнопками
        self.buttons_layout = QHBoxLayout()
        self.buttons_layout.setContentsMargins(0, 0, 0, 0)
        self.buttons_layout.setSpacing(5)
        
        # Левая часть панели - основные кнопки
        left_buttons_layout = QHBoxLayout()
        left_buttons_layout.setSpacing(5)
        
        # Кнопка Новый месяц
        self.new_month_btn = QPushButton("Новый месяц")
        self.new_month_btn.clicked.connect(self.select_month_and_year)
        self.new_month_btn.setFixedHeight(35)
        self.new_month_btn.setFixedWidth(120)
        self.new_month_btn.setStyleSheet("""
            QPushButton {
                background: #06B6D4;
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 400;
            }
            QPushButton:hover {
                background: #0891B2;
            }
            QPushButton:pressed {
                background-color: #0E7490;
            }
        """)
        left_buttons_layout.addWidget(self.new_month_btn)
        
        # Кнопка Новый сотрудник
        self.add_employee_btn = QPushButton("Новый сотрудник")
        self.add_employee_btn.clicked.connect(self.add_employee_dialog)
        self.add_employee_btn.setFixedHeight(30)
        self.add_employee_btn.setFixedWidth(140)
        self.add_employee_btn.setStyleSheet("""
            QPushButton {
                background-color: #808080;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 400;
            }
            QPushButton:hover {
                background-color: #696969;
            }
            QPushButton:pressed {
                background-color: #505050;
            }
        """)
        left_buttons_layout.addWidget(self.add_employee_btn)
        
        # Кнопка для добавления сотрудников в текущий период
        self.add_to_period_btn = QPushButton("Добавить в месяц")
        self.add_to_period_btn.clicked.connect(self.add_employees_to_period)
        self.add_to_period_btn.setFixedHeight(30)
        self.add_to_period_btn.setFixedWidth(140)
        self.add_to_period_btn.setStyleSheet("""
            QPushButton {
                background-color: #808080;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 400;
            }
            QPushButton:hover {
                background-color: #696969;
            }
            QPushButton:pressed {
                background-color: #505050;
            }
        """)
        left_buttons_layout.addWidget(self.add_to_period_btn)
        
        # Правая часть панели - кнопка Сохранить и другие
        right_buttons_layout = QHBoxLayout()
        right_buttons_layout.setSpacing(5)
        
        # Кнопка Открыть документ
        self.open_doc_btn = QPushButton("Открыть документ")
        self.open_doc_btn.clicked.connect(self.open_document)
        self.open_doc_btn.setFixedHeight(30)
        self.open_doc_btn.setFixedWidth(140)
        self.open_doc_btn.setStyleSheet("""
            QPushButton {
                background-color: #808080;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 400;
            }
            QPushButton:hover {
                background-color: #696969;
            }
            QPushButton:pressed {
                background-color: #505050;
            }
        """)
        right_buttons_layout.addWidget(self.open_doc_btn)
        
        # Кнопка Выбрать папку
        self.select_folder_btn = QPushButton("Выбрать папку")
        self.select_folder_btn.clicked.connect(self.select_export_folder)
        self.select_folder_btn.setFixedHeight(30)
        self.select_folder_btn.setFixedWidth(120)
        self.select_folder_btn.setStyleSheet("""
            QPushButton {
                background-color: #808080;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 400;
            }
            QPushButton:hover {
                background-color: #696969;
            }
            QPushButton:pressed {
                background-color: #505050;
            }
        """)
        right_buttons_layout.addWidget(self.select_folder_btn)
        
        # Кнопка Сохранить
        self.save_btn = QPushButton("СОХРАНИТЬ")
        self.save_btn.clicked.connect(self.save_data)
        self.save_btn.setFixedHeight(35)
        self.save_btn.setFixedWidth(120)
        self.save_btn.setStyleSheet("""
            QPushButton {
                background: #28A745;
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 400;
            }
            QPushButton:hover {
                background-color: #218838;
            }
            QPushButton:pressed {
                background-color: #1E7E34;
            }
        """)
        right_buttons_layout.addWidget(self.save_btn)
        
        # Объединяем левую и правую части
        self.buttons_layout.addLayout(left_buttons_layout)
        self.buttons_layout.addStretch()
        self.buttons_layout.addLayout(right_buttons_layout)
        
        self.layout.addLayout(self.buttons_layout)
        
        # Создаем сплиттер для вертикального расположения месяцев
        self.splitter = QSplitter(Qt.Vertical)
        self.splitter.setHandleWidth(1)
        self.layout.addWidget(self.splitter)
        
        # Создаем два виджета месяцев
        self.month_widget1 = MonthWidget(self, self.db)
        self.month_widget2 = MonthWidget(self, self.db)
        
        # Добавляем виджеты в сплиттер
        self.splitter.addWidget(self.month_widget1)
        self.splitter.addWidget(self.month_widget2)
        
        # Устанавливаем компактные размеры
        self.splitter.setSizes([350, 350])
        
        # Легенда статусов
        self.legend_layout = QHBoxLayout()
        self.legend_layout.setContentsMargins(0, 2, 0, 2)
        self.legend_layout.setSpacing(10)
        
        for status, (icon, text, color, _) in self.status_mapping.items():
            legend_item = QHBoxLayout()
            legend_item.setSpacing(3)
            color_label = QLabel("")
            color_label.setStyleSheet(f"font-size: 14px; background: {color.name()}; padding: 0px;")
            legend_item.addWidget(color_label)
            
            text_label = QLabel(text)
            text_label.setStyleSheet("padding: 0px; font-size: 10px;")
            legend_item.addWidget(text_label)
            self.legend_layout.addLayout(legend_item)
        
        self.legend_layout.addStretch()
        
        # Кнопка обновления в правом нижнем углу
        self.refresh_btn = QPushButton("Обновить")
        self.refresh_btn.setFixedHeight(25)
        self.refresh_btn.setFixedWidth(80)
        self.refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #6c757d;
                color: white;
                border: none;f;
                border-radius: 4px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #5a6268;
            }
            QPushButton:pressed {
                background-color: #545b62;
            }
        """)
        self.refresh_btn.clicked.connect(self.refresh_data)
        self.legend_layout.addWidget(self.refresh_btn)
        
        self.layout.addLayout(self.legend_layout)
        
        self.statusBar().showMessage("Готово")
    
    def set_unsaved_changes(self, has_changes):
        """Установка статуса несохраненных изменений"""
        self.unsaved_changes = has_changes
        if has_changes:
            self.save_status.setText("● Не сохранено")
            self.save_status.setStyleSheet("color: orange; font-weight: bold; padding: 2px;")
            # Запускаем таймер автосохранения
            self.last_change_time = datetime.now()
            self.auto_save_timer.start(5 * 60 * 1000)  # 5 минут - 300000 мс
        else:
            self.save_status.setText("Сохранено")
            self.save_status.setStyleSheet("color: green; font-weight: bold; padding: 2px;")
    
    def init_shortcuts(self):
        for status, (icon, text, _, _) in self.status_mapping.items():
            shortcut = QShortcut(QKeySequence(f"Ctrl+{status}"), self)
            shortcut.activated.connect(lambda _, s=status: self.update_selected_status(s))
    
    def auto_connect(self):
        """Автоматическое подключение при запуске"""
        if self.settings.value("auto_connect", True, type=bool):
            self.connect_to_server()
        else:
            self.show_connection_dialog()
    
    def connect_to_server(self):
        """Подключение к серверу"""
        if self.network_manager.connect(self.server_host, self.server_port):
            self.connection_status.setText("Подключено")
            self.connection_status.setStyleSheet("color: green; font-weight: bold; padding: 2px;")
            self.statusBar().showMessage("Успешное подключение к серверу", 3000)
            
            # Обновляем данные
            self.load_initial_data()
        else:
            self.connection_status.setText("Не подключено")
            self.connection_status.setStyleSheet("color: red; font-weight: bold; padding: 2px;")
            self.statusBar().showMessage("Ошибка подключения к серверу", 3000)
            
            # Все равно загружаем локальные данные
            self.load_initial_data()
    
    def show_connection_dialog(self):
        """Показать диалог подключения"""
        dialog = ConnectionDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.server_host, self.server_port = dialog.get_connection_params()
            
            # Сохраняем настройки
            self.settings.setValue("server_host", self.server_host)
            self.settings.setValue("server_port", self.server_port)
            
            self.connect_to_server()
    
    def load_initial_data(self):
        """Загружаем периоды в оба виджета"""
        self.month_widget1.load_periods()
        self.month_widget2.load_periods()
        
        periods = self.db.periods
        if periods:
            self.month_widget1.period_combo.setCurrentIndex(0)
            if len(periods) > 1:
                self.month_widget2.period_combo.setCurrentIndex(len(periods) - 1)
            else:
                self.month_widget2.period_combo.setCurrentIndex(0)
    def refresh_data(self):
        """Обновление данных с сервера"""
        if self.network_manager.connected:
            # Загружаем свежие данные с сервера
            try:
                # Принудительно обновляем кэши
                self.db._employees = []  # Сбрасываем кэш сотрудников
                self.db._periods = []    # Сбрасываем кэш периодов
                self.db._cache = {}      # Сбрасываем кэш расписаний
                
                # Загружаем обновленные данные
                employees = self.db.load_employees()
                periods = self.db.get_periods()
                
                # Обновляем отображение
                self.load_initial_data()
                
                self.statusBar().showMessage("Данные успешно обновлены с сервера", 3000)
            except Exception as e:
                self.statusBar().showMessage(f"Ошибка обновления: {str(e)}", 3000)
        else:
            self.statusBar().showMessage("Сервер недоступен, используются локальные данные", 3000)
    
    def select_month_and_year(self):
        """Открывает диалог выбора месяца и года"""
        dialog = MonthSelectionDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            selected_period = dialog.get_selected_period()
            
            # Проверяем, существует ли уже такой период
            existing_periods = [self.month_widget1.period_combo.itemData(i) for i in range(self.month_widget1.period_combo.count())]
            if selected_period in existing_periods:
                self.statusBar().showMessage("Этот месяц уже существует", 3000)
                return
            
            self.create_new_period(selected_period)

    def create_new_period(self, period):
        """Создает новый период с выбранными сотрудниками"""
        dialog = EmployeeSelectionDialog(self.db.employees, self)
        if dialog.exec_() == QDialog.Accepted:
            selected_names = dialog.get_selected_employees()
            if not selected_names:
                self.statusBar().showMessage("Не выбрано ни одного сотрудника", 3000)
                return
            
            # Создаем период
            if self.db.create_period(period, selected_names):
                self.statusBar().showMessage(f"Период {period} создан и синхронизирован с сервером", 3000)
            else:
                self.statusBar().showMessage(f"Период {period} создан локально, будет синхронизирован при подключении", 3000)
            
            # Обновляем комбо-боксы в обоих виджетах
            self.month_widget1.load_periods()
            self.month_widget2.load_periods()
            
            # Устанавливаем новый период в первый виджет
            try:
                year, month = map(int, period.split('-'))
                month_name = f"{self.month_names[month]} {year}"
                index = self.month_widget1.period_combo.findText(month_name)
                if index >= 0:
                    self.month_widget1.period_combo.setCurrentIndex(index)
            except:
                pass
    
    def add_employees_to_period(self):
        """Добавляет сотрудников в уже созданный месяц"""
        current_widget = self.get_active_month_widget()
        if not current_widget or not current_widget.current_period:
            self.statusBar().showMessage("Сначала выберите месяц в одном из окон", 3000)
            return
        
        schedule_data = self.db.load_schedule(current_widget.current_period)
        if not schedule_data:
            self.statusBar().showMessage("Не удалось загрузить данные выбранного месяца", 3000)
            return
        
        current_employees = schedule_data.get("employees", [])
        
        dialog = AddEmployeeToPeriodDialog(self.db.employees, current_employees, self)
        if dialog.exec_() == QDialog.Accepted:
            selected_names = dialog.get_selected_employees()
            if not selected_names:
                return
            
            # Добавляем сотрудников в период
            if self.db.add_employees_to_period(current_widget.current_period, selected_names):
                self.statusBar().showMessage(f"Добавлено сотрудников: {len(selected_names)}", 3000)
            else:
                self.statusBar().showMessage(f"Добавлено сотрудников локально, будет синхронизировано при подключении", 3000)
            
            # Перезагружаем данные
            current_widget.load_data(current_widget.current_period)
    
    def get_active_month_widget(self):
        """Определяет, какой виджет месяца активен (имеет фокус)"""
        if self.month_widget1.table.hasFocus():
            return self.month_widget1
        elif self.month_widget2.table.hasFocus():
            return self.month_widget2
        else:
            return self.month_widget1
    
    def save_data(self):
        """Сохранение данных с отправкой на сервер"""
        try:
            # Сохраняем оба месяца в локальную БД и отправляем на сервер
            success_count = 0
            total_count = 0
            saved_periods = []
            
            for widget in [self.month_widget1, self.month_widget2]:
                if widget.current_period:
                    total_count += 1
                    schedule_data = self.prepare_save_data(widget)
                    if schedule_data:
                        if self.db.save_schedule(widget.current_period, schedule_data):
                            success_count += 1
                            saved_periods.append(widget.current_period)
            
            # Показываем результат в статусбаре
            if success_count > 0:
                if self.network_manager.connected:
                    message = f"Успешно сохранено {success_count} из {total_count} месяцев на сервер"
                else:
                    message = f"Успешно сохранено {success_count} из {total_count} месяцев локально (сервер недоступен)"
                
                self.statusBar().showMessage(message, 5000)
                # Сбрасываем статус несохраненных изменений
                self.set_unsaved_changes(False)
            else:
                self.statusBar().showMessage("Не удалось сохранить данные", 5000)
            
            # Всегда экспортируем в Excel
            self.export_to_excel()
            
        except Exception as e:
            self.statusBar().showMessage(f"Ошибка сохранения: {str(e)}", 5000)
    
    def prepare_save_data(self, widget):
        schedule_data = self.db.load_schedule(widget.current_period)
        if not schedule_data:
            return None
            
        try:
            year, month = map(int, widget.current_period.split('-'))
            days_in_month = monthrange(year, month)[1]
            working_days_count = widget.get_working_days_count(year, month)
            day_mapping = widget.get_day_mapping(year, month)
        except:
            days_in_month = widget.table.columnCount() - 4
            working_days_count = days_in_month
            day_mapping = list(range(1, days_in_month + 1))
        
        schedule_data["schedule"] = {}
        schedule_data["notes"] = {}
        
        for row in range(widget.table.rowCount()):
            emp_name = widget.table.verticalHeaderItem(row).text()
            full_schedule = [4] * days_in_month
            notes = {}
            
            for col, actual_day in enumerate(day_mapping):
                day_index = actual_day - 1
                item = widget.table.item(row, col)
                if item and isinstance(item, NoteItem):
                    bg_color = item.background().color()
                    if bg_color == QColor(Qt.white):
                        status = 4
                    else:
                        status = next((k for k, v in self.status_mapping.items() if v[2] == bg_color), 4)
                    full_schedule[day_index] = status
                    
                    if item.has_note:
                        notes[str(day_index)] = item.note_data
            
            schedule_data["schedule"][emp_name] = full_schedule
            if notes:
                schedule_data["notes"][emp_name] = notes
        
        return schedule_data
    
    def update_selected_status(self, status):
        """Обновляет статус для активного виджета"""
        active_widget = self.get_active_month_widget()
        if active_widget:
            active_widget.update_selected_status(status)
    
    def auto_save(self):
        """Автоматическое сохранение через 5 минут после последнего изменения"""
        if self.unsaved_changes and self.last_change_time and (datetime.now() - self.last_change_time).total_seconds() >= 60:
            self.statusBar().showMessage("Автоматическое сохранение...", 2000)
            self.save_data()
    
    def add_employee_dialog(self):
        name, ok = QInputDialog.getText(self, "Добавить сотрудника", "ФИО сотрудника:")
        if ok and name:
            name = name.strip()
            if not name:
                self.statusBar().showMessage("Имя сотрудника не может быть пустым", 3000)
                return
                
            if self.db.add_employee(name):
                self.statusBar().showMessage("Сотрудник добавлен", 3000)
            else:
                self.statusBar().showMessage("Сотрудник с таким именем уже существует", 3000)
    
    def open_document(self):
        """Открывает сохраненную таблицу Excel"""
        if not self.export_folder:
            self.statusBar().showMessage("Сначала выберите папку для экспорта", 3000)
            return
        
        file_path = os.path.join(self.export_folder, "Расписание_все_месяцы.xlsx")
        
        if not os.path.exists(file_path):
            self.statusBar().showMessage(f"Файл не найден: {file_path}", 3000)
            return
        
        try:
            if sys.platform == "win32":
                os.startfile(file_path)
            elif sys.platform == "darwin":
                os.system(f'open "{file_path}"')
            else:
                os.system(f'xdg-open "{file_path}"')
            
            self.statusBar().showMessage(f"Открыт файл: {file_path}", 3000)
            
        except Exception as e:
            self.statusBar().showMessage(f"Не удалось открыть файл: {str(e)}", 3000)
        
    def select_export_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения")
        if folder:
            self.export_folder = folder
            self.settings.setValue("export_folder", folder)
            self.statusBar().showMessage(f"Папка для экспорта установлена: {folder}", 3000)
    
    def export_to_excel(self):
        """Экспорт всех данных в Excel"""
        try:
            if not self.export_folder:
                self.select_export_folder()
                if not self.export_folder:
                    self.statusBar().showMessage("Экспорт в Excel отменен - не выбрана папка", 3000)
                    return

            file_name = "Расписание_все_месяцы.xlsx"
            file_path = os.path.join(self.export_folder, file_name)

            # Создаем папку если она не существует
            os.makedirs(self.export_folder, exist_ok=True)

            wb = Workbook()
            ws = wb.active
            ws.title = "Расписание"

            header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            bold_font = Font(bold=True)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                            top=Side(style='thin'), bottom=Side(style='thin'))
            
            status_styles = {
                0: PatternFill(start_color="7FFF7F", end_color="7FFF7F", fill_type="solid"),
                1: PatternFill(start_color="8080FF", end_color="8080FF", fill_type="solid"),
                2: PatternFill(start_color="FF7777", end_color="FF7777", fill_type="solid"),
                3: PatternFill(start_color="FFFF77", end_color="FFFF77", fill_type="solid"),
                4: PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            }

            total_shifts_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
            registry_fill = PatternFill(start_color="8080FF", end_color="8080FF", fill_type="solid")
            call_center_fill = PatternFill(start_color="7FFF7F", end_color="7FFF7F", fill_type="solid")
            hours_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")

            max_working_days = 0
            periods = self.db.periods
            
            # Проверяем есть ли данные для экспорта
            if not periods:
                self.statusBar().showMessage("Нет данных для экспорта", 3000)
                return
            
            has_data = False
            for period in periods:
                schedule_data = self.db.load_schedule(period)
                if schedule_data and schedule_data.get('schedule'):
                    has_data = True
                    try:
                        year, month = map(int, period.split('-'))
                        working_days_count = self.month_widget1.get_working_days_count(year, month)
                        max_working_days = max(max_working_days, working_days_count)
                    except:
                        working_days_count = len(next(iter(schedule_data['schedule'].values())))
                        max_working_days = max(max_working_days, working_days_count)
            
            if not has_data:
                self.statusBar().showMessage("Нет данных для экспорта", 3000)
                return
            
            if max_working_days == 0:
                max_working_days = 31

            RESULT_COLUMNS_START = max_working_days + 2
            RESULT_COLUMNS = {
                "Смены": RESULT_COLUMNS_START,
                "Рег": RESULT_COLUMNS_START + 1,
                "КЦ": RESULT_COLUMNS_START + 2,
                "Часы": RESULT_COLUMNS_START + 3
            }

            start_row = 1

            for period in periods:
                schedule_data = self.db.load_schedule(period)
                if not schedule_data or not schedule_data.get('schedule'):
                    continue

                try:
                    year, month = map(int, period.split('-'))
                    month_name = f"{self.month_names[month]} {year}"
                    working_days_count = self.month_widget1.get_working_days_count(year, month)
                    day_mapping = self.month_widget1.get_day_mapping(year, month)
                    days_in_month = monthrange(year, month)[1]
                except:
                    month_name = period
                    working_days_count = len(next(iter(schedule_data['schedule'].values())))
                    day_mapping = list(range(1, working_days_count + 1))
                    days_in_month = working_days_count

                employees = schedule_data.get('employees', [])
                if not employees:
                    continue

                notes_data = schedule_data.get('notes', {})

                working_counts = [0] * working_days_count
                for emp in employees:
                    emp_schedule = schedule_data['schedule'].get(emp["name"], [2]*days_in_month)
                    for col, actual_day in enumerate(day_mapping):
                        day_index = actual_day - 1
                        if emp_schedule[day_index] in [0, 1]:
                            working_counts[col] += 1

                total_columns = max_working_days + 5
                ws.cell(row=start_row, column=1, value=month_name).font = Font(bold=True, size=12)
                ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=total_columns)
                start_row += 1

                # Заголовки столбцов
                ws.cell(row=start_row, column=1, value="Сотрудник").fill = header_fill
                ws.cell(row=start_row, column=1).alignment = center_alignment
                ws.cell(row=start_row, column=1).font = bold_font
                ws.cell(row=start_row, column=1).border = thin_border

                for col, actual_day in enumerate(day_mapping):
                    try:
                        day_of_week = (datetime(year, month, actual_day).weekday())
                        day_name = self.day_names[day_of_week]
                        count = working_counts[col]
                        header_text = f"{actual_day}\n{day_name}\n({count})"
                    except:
                        header_text = str(actual_day)
                    
                    cell = ws.cell(row=start_row, column=col+2, value=header_text)
                    cell.fill = header_fill
                    cell.alignment = center_alignment
                    cell.font = bold_font
                    cell.border = thin_border

                # Заголовки результирующих столбцов
                ws.cell(row=start_row, column=RESULT_COLUMNS["Смены"], value="Смены").fill = total_shifts_fill
                ws.cell(row=start_row, column=RESULT_COLUMNS["Смены"]).alignment = center_alignment
                ws.cell(row=start_row, column=RESULT_COLUMNS["Смены"]).font = bold_font
                ws.cell(row=start_row, column=RESULT_COLUMNS["Смены"]).border = thin_border

                ws.cell(row=start_row, column=RESULT_COLUMNS["Рег"], value="Рег").fill = registry_fill
                ws.cell(row=start_row, column=RESULT_COLUMNS["Рег"]).alignment = center_alignment
                ws.cell(row=start_row, column=RESULT_COLUMNS["Рег"]).font = bold_font
                ws.cell(row=start_row, column=RESULT_COLUMNS["Рег"]).border = thin_border

                ws.cell(row=start_row, column=RESULT_COLUMNS["КЦ"], value="КЦ").fill = call_center_fill
                ws.cell(row=start_row, column=RESULT_COLUMNS["КЦ"]).alignment = center_alignment
                ws.cell(row=start_row, column=RESULT_COLUMNS["КЦ"]).font = bold_font
                ws.cell(row=start_row, column=RESULT_COLUMNS["КЦ"]).border = thin_border

                ws.cell(row=start_row, column=RESULT_COLUMNS["Часы"], value="Часы").fill = hours_fill
                ws.cell(row=start_row, column=RESULT_COLUMNS["Часы"]).alignment = center_alignment
                ws.cell(row=start_row, column=RESULT_COLUMNS["Часы"]).font = bold_font
                ws.cell(row=start_row, column=RESULT_COLUMNS["Часы"]).border = thin_border

                start_row += 1

                ws.column_dimensions['A'].width = 30
                
                for col in range(2, max_working_days + 2):
                    ws.column_dimensions[get_column_letter(col)].width = 4
                
                ws.column_dimensions[get_column_letter(RESULT_COLUMNS["Смены"])].width = 8
                ws.column_dimensions[get_column_letter(RESULT_COLUMNS["Рег"])].width = 8
                ws.column_dimensions[get_column_letter(RESULT_COLUMNS["КЦ"])].width = 8
                ws.column_dimensions[get_column_letter(RESULT_COLUMNS["Часы"])].width = 8

                for emp in employees:
                    emp_name = emp["name"]
                    ws.cell(row=start_row, column=1, value=emp_name).font = bold_font
                    
                    schedule = schedule_data['schedule'].get(emp["name"], [2]*days_in_month)
                    emp_notes = notes_data.get(emp["name"], {})
                    total_shifts = 0
                    total_hours = 0.0
                    call_center_days = 0
                    registry_days = 0
                    
                    for col, actual_day in enumerate(day_mapping):
                        day_index = actual_day - 1
                        status = schedule[day_index]
                        
                        if status == 0:
                            call_center_days += 1
                            total_shifts += 1
                            if str(day_index) in emp_notes:
                                total_hours += emp_notes[str(day_index)].get('worked_hours', 12)
                            else:
                                total_hours += self.status_mapping[0][3]
                        elif status == 1:
                            registry_days += 1
                            total_shifts += 1
                            if str(day_index) in emp_notes:
                                total_hours += emp_notes[str(day_index)].get('worked_hours', 12)
                            else:
                                total_hours += self.status_mapping[1][3]
                        
                        cell = ws.cell(row=start_row, column=col+2)
                        
                        if str(day_index) in emp_notes:
                            end_time = emp_notes[str(day_index)].get('end_time', '20:00')
                            cell.value = end_time
                        else:
                            cell.value = ""
                        
                        cell.alignment = center_alignment
                        cell.border = thin_border
                        
                        if status in status_styles:
                            cell.fill = status_styles[status]
                    
                    hours, minutes = self.hours_to_hours_minutes(total_hours)
                    hours_text = f"{hours}ч {minutes}м"
                    
                    cell = ws.cell(row=start_row, column=RESULT_COLUMNS["Смены"], value=total_shifts)
                    cell.fill = total_shifts_fill
                    cell.alignment = center_alignment
                    cell.font = bold_font
                    cell.border = thin_border
                    
                    cell = ws.cell(row=start_row, column=RESULT_COLUMNS["Рег"], value=registry_days)
                    cell.fill = registry_fill
                    cell.alignment = center_alignment
                    cell.font = bold_font
                    cell.border = thin_border
                    
                    cell = ws.cell(row=start_row, column=RESULT_COLUMNS["КЦ"], value=call_center_days)
                    cell.fill = call_center_fill
                    cell.alignment = center_alignment
                    cell.font = bold_font
                    cell.border = thin_border
                    
                    cell = ws.cell(row=start_row, column=RESULT_COLUMNS["Часы"], value=hours_text)
                    cell.fill = hours_fill
                    cell.alignment = center_alignment
                    cell.font = bold_font
                    cell.border = thin_border
                    
                    start_row += 1

                start_row += 2

            # Сохраняем файл
            wb.save(file_path)
            self.statusBar().showMessage(f"Файл успешно экспортирован: {file_path}", 5000)

        except PermissionError:
            self.statusBar().showMessage(f"Ошибка: Нет прав для записи в папку {self.export_folder}", 5000)
        except Exception as e:
            self.statusBar().showMessage(f"Ошибка экспорта: {str(e)}", 5000)

    def hours_to_hours_minutes(self, total_hours):
        """Конвертирует дробное количество часов в часы и минуты"""
        hours = int(total_hours)
        minutes = int(round((total_hours - hours) * 60))
        
        if minutes == 60:
            hours += 1
            minutes = 0
            
        return hours, minutes


if __name__ == "__main__":
    app = QApplication(sys.argv)
    font = app.font()
    font.setFamily("Arial")
    font.setPointSize(10)
    app.setFont(font)
    window = ScheduleApp()
    window.showMaximized()
    sys.exit(app.exec_())