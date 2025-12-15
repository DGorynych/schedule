import sys
import sqlite3
import socket
import threading
import json
import os
import winreg
from calendar import monthrange
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QTextEdit, QPushButton, QCheckBox, 
                             QLabel, QSystemTrayIcon, QMenu, QAction, QMessageBox,
                             QGroupBox)
from PyQt5.QtCore import QTimer, pyqtSignal, QObject, QSettings, Qt
from PyQt5.QtGui import QIcon, QFont

class ProcessLock:
    """Класс для блокировки процесса и предотвращения повторного запуска"""
    
    def __init__(self, lock_name="schedule_server.lock"):
        self.lock_name = lock_name
        self.lock_file = self.get_temp_path(lock_name)
        self.lock_handle = None
    
    def get_temp_path(self, filename):
        """Получает путь к временной директории"""
        try:
            # Пробуем стандартный метод
            if hasattr(os, 'gettempdir'):
                return os.path.join(os.gettempdir(), filename)
            # Для Windows
            elif 'TEMP' in os.environ:
                return os.path.join(os.environ['TEMP'], filename)
            elif 'TMP' in os.environ:
                return os.path.join(os.environ['TMP'], filename)
            # Для Unix-like
            elif 'TMPDIR' in os.environ:
                return os.path.join(os.environ['TMPDIR'], filename)
            else:
                # Используем текущую директорию как запасной вариант
                return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        except:
            # Если все методы не сработали, используем текущую директорию
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    
    def is_already_running(self):
        """Проверяет, запущен ли уже процесс"""
        # Проверяем по файлу блокировки
        if os.path.exists(self.lock_file):
            try:
                # Читаем PID из файла
                with open(self.lock_file, 'r') as f:
                    pid_line = f.read().strip()
                    if ':' in pid_line:
                        pid, port = pid_line.split(':')
                        pid = int(pid)
                    else:
                        pid = int(pid_line)
                
                # Простая проверка существования процесса (кроссплатформенная)
                if self.is_process_running(pid):
                    return True
                else:
                    # Если процесс не существует, удаляем старый файл блокировки
                    try:
                        os.remove(self.lock_file)
                    except:
                        pass
            except (ValueError, IOError):
                # Если файл поврежден, удаляем его
                try:
                    os.remove(self.lock_file)
                except:
                    pass
        
        return False
    
    def is_process_running(self, pid):
        """Проверяет, запущен ли процесс с указанным PID"""
        try:
            if os.name == 'nt':  # Windows
                import ctypes
                kernel32 = ctypes.windll.kernel32
                handle = kernel32.OpenProcess(1, False, pid)
                if handle:
                    kernel32.CloseHandle(handle)
                    return True
                return False
            else:  # Unix-like
                import errno
                try:
                    os.kill(pid, 0)
                    return True
                except OSError as e:
                    return e.errno == errno.EPERM
        except:
            return False
    
    def create_lock(self):
        """Создает файл блокировки"""
        try:
            with open(self.lock_file, 'w') as f:
                f.write(f"{os.getpid()}:48888")
            return True
        except:
            return False
    
    def release_lock(self):
        """Удаляет файл блокировки"""
        try:
            if os.path.exists(self.lock_file):
                os.remove(self.lock_file)
            return True
        except:
            return False


class SingleInstance:
    """Класс для обеспечения единственного экземпляра приложения"""
    
    def __init__(self, port=48888):
        self.port = port
        self.socket = None
        self.listener_thread = None
        self.is_primary = False
        self.process_lock = ProcessLock()
    
    def is_already_running(self):
        """Проверяет, запущен ли уже экземпляр приложения"""
        # Сначала проверяем по файлу блокировки
        if self.process_lock.is_already_running():
            return True
        
        # Затем пробуем подключиться к сокету
        return self.can_connect_to_socket()
    
    def can_connect_to_socket(self):
        """Пытается подключиться к сокету другого экземпляра"""
        try:
            test_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            test_socket.settimeout(1)
            result = test_socket.connect_ex(('localhost', self.port))
            test_socket.close()
            return result == 0
        except:
            return False
    
    def send_message_to_running_instance(self, message="show"):
        """Отправляет сообщение уже запущенному экземпляру"""
        try:
            client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            client_socket.settimeout(2)
            client_socket.connect(('localhost', self.port))
            client_socket.send(message.encode('utf-8'))
            client_socket.close()
            return True
        except:
            return False
    
    def start_listener(self, callback=None):
        """Запускает прослушивание сообщений от других экземпляров"""
        try:
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.socket.bind(('localhost', self.port))
            self.socket.listen(1)
            self.socket.settimeout(1)
            self.is_primary = True
            
            self.listener_thread = threading.Thread(
                target=self._listen_for_messages, 
                args=(callback,),
                daemon=True
            )
            self.listener_thread.start()
            
            # Создаем файл блокировки
            self.process_lock.create_lock()
            
            return True
        except:
            return False
    
    def _listen_for_messages(self, callback):
        """Прослушивает входящие сообщения"""
        while self.is_primary and self.socket:
            try:
                client, addr = self.socket.accept()
                data = client.recv(1024).decode('utf-8')
                if callback and data:
                    # Вызываем callback в главном потоке
                    QTimer.singleShot(0, lambda: callback(data))
                client.close()
            except socket.timeout:
                continue
            except:
                break
    
    def cleanup(self):
        """Очистка ресурсов"""
        self.is_primary = False
        if self.socket:
            try:
                self.socket.close()
            except:
                pass
            self.socket = None
        
        self.process_lock.release_lock()


def check_already_running():
    """Проверяет, не запущен ли уже экземпляр приложения"""
    single_instance = SingleInstance()
    
    if single_instance.is_already_running():
        # Если приложение уже запущено, отправляем сообщение и закрываемся
        if single_instance.send_message_to_running_instance("show"):
            QMessageBox.information(None, "Schedule Server", 
                                  "Приложение уже запущено. Окно будет показано.")
        else:
            QMessageBox.warning(None, "Schedule Server", 
                              "Приложение уже запущено, но не отвечает.\n\n"
                              "Если вы уверены, что приложение не запущено, "
                              "удалите файл блокировки и попробуйте снова.")
        return True
    return False


class ServerSignals(QObject):
    log_message = pyqtSignal(str)
    update_clients = pyqtSignal(int)

class ScheduleServer:
    def __init__(self, host='0.0.0.0', port=8888):
        self.host = host
        self.port = port
        self.running = False
        self.server_socket = None
        self.clients = []
        self.signals = ServerSignals()
        
        # Инициализация базы данных
        self.init_database()
    
    def init_database(self):
        """Инициализация SQLite базы данных"""
        self.conn = sqlite3.connect('schedule_server.db', check_same_thread=False)
        self.cursor = self.conn.cursor()
        
        # Таблица сотрудников
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                position TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Таблица периодов
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS periods (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                period TEXT UNIQUE NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Таблица расписаний
        self.cursor.execute('''
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
        
        self.conn.commit()
    
    def handle_client(self, client_socket, address):
        """Обработка клиентских подключений"""
        self.signals.log_message.emit(f"Клиент подключен: {address}")
        self.clients.append((client_socket, address))
        self.signals.update_clients.emit(len(self.clients))
        
        try:
            while self.running:
                # Получение данных от клиента
                data = client_socket.recv(10240).decode('utf-8')
                if not data:
                    break
                
                try:
                    request = json.loads(data)
                    response = self.process_request(request)
                    client_socket.send(json.dumps(response).encode('utf-8'))
                except json.JSONDecodeError:
                    self.signals.log_message.emit(f"Ошибка декодирования JSON от {address}")
                
        except Exception as e:
            self.signals.log_message.emit(f"Ошибка с клиентом {address}: {str(e)}")
        finally:
            if (client_socket, address) in self.clients:
                self.clients.remove((client_socket, address))
            self.signals.update_clients.emit(len(self.clients))
            client_socket.close()
            self.signals.log_message.emit(f"Клиент отключен: {address}")
    
    def process_request(self, request):
        """Обработка запросов от клиентов"""
        action = request.get('action')
        
        try:
            if action == 'get_employees':
                return self.get_employees()
            elif action == 'get_periods':
                return self.get_periods()
            elif action == 'get_schedule':
                return self.get_schedule(request['period'])
            elif action == 'save_schedule':
                return self.save_schedule(request['period'], request['data'])
            elif action == 'add_employee':
                return self.add_employee(request['name'])
            elif action == 'create_period':
                return self.create_period(request['period'], request['employees'])
            elif action == 'add_employees_to_period':
                return self.add_employees_to_period(request['period'], request['employees'])
            elif action == 'sync_data':
                return self.sync_data(request['data'])
            else:
                return {'status': 'error', 'message': 'Неизвестное действие'}
        except Exception as e:
            return {'status': 'error', 'message': str(e)}
    
    def get_employees(self):
        """Получение списка сотрудников"""
        self.cursor.execute("SELECT name, position FROM employees")
        employees = [{'name': row[0], 'position': row[1]} for row in self.cursor.fetchall()]
        return {'status': 'success', 'data': employees}
    
    def get_periods(self):
        """Получение списка периодов"""
        self.cursor.execute("SELECT period FROM periods ORDER BY period")
        periods = [row[0] for row in self.cursor.fetchall()]
        return {'status': 'success', 'data': periods}
    
    def get_schedule(self, period):
        """Получение расписания для периода"""
        # Получаем или создаем период
        period_id = self.get_or_create_period(period)
        
        # Получаем сотрудников периода
        self.cursor.execute('''
            SELECT DISTINCT employee_name FROM schedules 
            WHERE period_id = ? ORDER BY employee_name
        ''', (period_id,))
        employees = [{'name': row[0]} for row in self.cursor.fetchall()]
        
        # Получаем расписание
        self.cursor.execute('''
            SELECT employee_name, day_index, status, end_time, worked_hours 
            FROM schedules WHERE period_id = ?
        ''', (period_id,))
        
        schedule = {}
        notes = {}
        
        for row in self.cursor.fetchall():
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
            'status': 'success', 
            'data': {
                'employees': employees,
                'schedule': schedule,
                'notes': notes
            }
        }
    
    def save_schedule(self, period, data):
        """Сохранение расписания"""
        period_id = self.get_or_create_period(period)
        
        try:
            for emp_name, schedule in data['schedule'].items():
                for day_index, status in enumerate(schedule):
                    # Проверяем существование записи
                    self.cursor.execute('''
                        SELECT id FROM schedules 
                        WHERE period_id = ? AND employee_name = ? AND day_index = ?
                    ''', (period_id, emp_name, day_index))
                    
                    result = self.cursor.fetchone()
                    
                    notes_data = data['notes'].get(emp_name, {}).get(str(day_index), {})
                    
                    if result:
                        # Обновляем существующую запись
                        self.cursor.execute('''
                            UPDATE schedules SET status = ?, end_time = ?, worked_hours = ?
                            WHERE id = ?
                        ''', (status, notes_data.get('end_time'), notes_data.get('worked_hours'), result[0]))
                    else:
                        # Создаем новую запись
                        self.cursor.execute('''
                            INSERT INTO schedules (period_id, employee_name, day_index, status, end_time, worked_hours)
                            VALUES (?, ?, ?, ?, ?, ?)
                        ''', (period_id, emp_name, day_index, status, 
                              notes_data.get('end_time'), notes_data.get('worked_hours')))
            
            self.conn.commit()
            return {'status': 'success', 'message': 'Расписание сохранено'}
        except Exception as e:
            self.conn.rollback()
            return {'status': 'error', 'message': str(e)}
    
    def get_or_create_period(self, period):
        """Получение или создание периода"""
        self.cursor.execute("SELECT id FROM periods WHERE period = ?", (period,))
        result = self.cursor.fetchone()
        
        if result:
            return result[0]
        else:
            self.cursor.execute("INSERT INTO periods (period) VALUES (?)", (period,))
            self.conn.commit()
            return self.cursor.lastrowid
    
    def add_employee(self, name):
        """Добавление сотрудника"""
        try:
            self.cursor.execute("INSERT INTO employees (name) VALUES (?)", (name,))
            self.conn.commit()
            return {'status': 'success', 'message': 'Сотрудник добавлен'}
        except sqlite3.IntegrityError:
            return {'status': 'error', 'message': 'Сотрудник уже существует'}
    
    def create_period(self, period, employees):
        """Создание нового периода"""
        period_id = self.get_or_create_period(period)
        
        try:
            # Определяем количество дней в месяце
            try:
                year, month = map(int, period.split('-'))
                days_in_month = monthrange(year, month)[1]
            except:
                days_in_month = 31
            
            # Добавляем сотрудников в период
            for emp_name in employees:
                # Создаем пустое расписание на нужное количество дней
                for day_index in range(days_in_month):
                    self.cursor.execute('''
                        INSERT OR IGNORE INTO schedules (period_id, employee_name, day_index, status)
                        VALUES (?, ?, ?, ?)
                    ''', (period_id, emp_name, day_index, 4))
            
            self.conn.commit()
            return {'status': 'success', 'message': 'Период создан'}
        except Exception as e:
            self.conn.rollback()
            return {'status': 'error', 'message': str(e)}
    
    def add_employees_to_period(self, period, employees):
        """Добавление сотрудников в существующий период"""
        period_id = self.get_or_create_period(period)
        
        try:
            for emp_name in employees:
                # Создаем пустое расписание на 31 день
                for day_index in range(31):
                    self.cursor.execute('''
                        INSERT OR IGNORE INTO schedules (period_id, employee_name, day_index, status)
                        VALUES (?, ?, ?, ?)
                    ''', (period_id, emp_name, day_index, 4))
            
            self.conn.commit()
            return {'status': 'success', 'message': 'Сотрудники добавлены'}
        except Exception as e:
            self.conn.rollback()
            return {'status': 'error', 'message': str(e)}
    
    def sync_data(self, data):
        """Синхронизация данных с клиента"""
        try:
            # Синхронизация сотрудников
            for emp in data.get('employees', []):
                try:
                    self.cursor.execute("INSERT OR IGNORE INTO employees (name, position) VALUES (?, ?)", 
                                      (emp['name'], emp.get('position', '')))
                except:
                    pass
            
            # Синхронизация периодов
            for period in data.get('periods', []):
                try:
                    self.cursor.execute("INSERT OR IGNORE INTO periods (period) VALUES (?)", (period,))
                except:
                    pass
            
            # Синхронизация расписаний
            for schedule_data in data.get('schedules', []):
                period = schedule_data['period']
                period_id = self.get_or_create_period(period)
                
                for emp_name, schedule in schedule_data['schedule'].items():
                    for day_index, status in enumerate(schedule):
                        notes_data = schedule_data['notes'].get(emp_name, {}).get(str(day_index), {})
                        
                        self.cursor.execute('''
                            INSERT OR REPLACE INTO schedules 
                            (period_id, employee_name, day_index, status, end_time, worked_hours)
                            VALUES (?, ?, ?, ?, ?, ?)
                        ''', (period_id, emp_name, day_index, status,
                              notes_data.get('end_time'), notes_data.get('worked_hours')))
            
            self.conn.commit()
            return {'status': 'success', 'message': 'Данные синхронизированы'}
        except Exception as e:
            self.conn.rollback()
            return {'status': 'error', 'message': str(e)}
    
    def start_server(self):
        """Запуск сервера"""
        try:
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.server_socket.bind((self.host, self.port))
            self.server_socket.listen(5)
            self.running = True
            
            self.signals.log_message.emit(f"Сервер запущен на {self.host}:{self.port}")
            
            while self.running:
                try:
                    client_socket, address = self.server_socket.accept()
                    client_thread = threading.Thread(
                        target=self.handle_client, 
                        args=(client_socket, address),
                        daemon=True
                    )
                    client_thread.start()
                except socket.error:
                    if self.running:
                        self.signals.log_message.emit("Ошибка при принятии подключения")
        except Exception as e:
            self.signals.log_message.emit(f"Ошибка запуска сервера: {str(e)}")
    
    def stop_server(self):
        """Остановка сервера"""
        self.running = False
        if self.server_socket:
            self.server_socket.close()
        
        for client_socket, _ in self.clients:
            try:
                client_socket.close()
            except:
                pass
        self.clients.clear()
        
        if hasattr(self, 'conn'):
            self.conn.close()
        
        self.signals.log_message.emit("Сервер остановлен")


class AutoStartManager:
    """Менеджер автозапуска для Windows"""
    
    @staticmethod
    def get_registry_path():
        """Получает путь в реестре для автозапуска"""
        return r"Software\Microsoft\Windows\CurrentVersion\Run"
    
    @staticmethod
    def get_app_name():
        """Получает имя приложения для реестра"""
        return "ScheduleServer"
    
    @staticmethod
    def get_executable_path():
        """Получает путь к исполняемому файлу"""
        if getattr(sys, 'frozen', False):
            # Если приложение собрано в exe
            return sys.executable
        else:
            # Если запуск из Python
            return os.path.abspath(sys.argv[0])
    
    @staticmethod
    def is_auto_start_enabled():
        """Проверяет, включен ли автозапуск"""
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AutoStartManager.get_registry_path()) as key:
                try:
                    winreg.QueryValueEx(key, AutoStartManager.get_app_name())
                    return True
                except FileNotFoundError:
                    return False
        except Exception:
            return False
    
    @staticmethod
    def enable_auto_start():
        """Включает автозапуск"""
        try:
            executable_path = AutoStartManager.get_executable_path()
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AutoStartManager.get_registry_path(), 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, AutoStartManager.get_app_name(), 0, winreg.REG_SZ, f'"{executable_path}" --minimized')
            return True
        except Exception as e:
            print(f"Ошибка включения автозапуска: {e}")
            return False
    
    @staticmethod
    def disable_auto_start():
        """Выключает автозапуск"""
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AutoStartManager.get_registry_path(), 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, AutoStartManager.get_app_name())
            return True
        except Exception as e:
            print(f"Ошибка выключения автозапуска: {e}")
            return False


class ServerWindow(QMainWindow):
    def __init__(self, single_instance):
        super().__init__()
        self.single_instance = single_instance
        self.server = None
        self.server_thread = None
        self.tray_icon = None
        self.settings = QSettings("MyCompany", "ScheduleServer")
        self.init_ui()
        self.init_tray()
        
        # Проверяем аргументы командной строки
        if '--minimized' in sys.argv:
            self.hide()
            if self.tray_icon:
                self.tray_icon.showMessage(
                    "Schedule Server",
                    "Сервер запущен в фоновом режиме",
                    QSystemTrayIcon.Information,
                    2000
                )
    
    def handle_instance_message(self, message):
        """Обработка сообщений от других экземпляров приложения"""
        if message == "show":
            self.show_window()
            self.activateWindow()
            self.raise_()
    
    def init_ui(self):
        self.setWindowTitle("Schedule Server")
        self.setGeometry(300, 300, 800, 600)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Группа настроек
        settings_group = QGroupBox("Настройки сервера")
        settings_layout = QVBoxLayout(settings_group)
        
        # Статус сервера
        status_layout = QHBoxLayout()
        status_layout.addWidget(QLabel("Статус сервера:"))
        self.status_label = QLabel("Остановлен")
        self.status_label.setStyleSheet("color: red; font-weight: bold;")
        status_layout.addWidget(self.status_label)
        status_layout.addStretch()
        settings_layout.addLayout(status_layout)
        
        # Подключенные клиенты
        clients_layout = QHBoxLayout()
        clients_layout.addWidget(QLabel("Подключенные клиенты:"))
        self.clients_label = QLabel("0")
        clients_layout.addWidget(self.clients_label)
        clients_layout.addStretch()
        settings_layout.addLayout(clients_layout)
        
        # Автозапуск при запуске программы
        self.auto_start_cb = QCheckBox("Автозапуск сервера при запуске программы")
        self.auto_start_cb.setChecked(self.settings.value("auto_start", True, type=bool))
        self.auto_start_cb.stateChanged.connect(self.toggle_auto_start)
        settings_layout.addWidget(self.auto_start_cb)
        
        # Автозапуск при запуске системы
        self.system_auto_start_cb = QCheckBox("Автозапуск сервера при запуске системы")
        self.system_auto_start_cb.setChecked(AutoStartManager.is_auto_start_enabled())
        self.system_auto_start_cb.stateChanged.connect(self.toggle_system_auto_start)
        settings_layout.addWidget(self.system_auto_start_cb)
        
        # Кнопки управления
        buttons_layout = QHBoxLayout()
        self.start_btn = QPushButton("Запустить сервер")
        self.start_btn.clicked.connect(self.start_server)
        buttons_layout.addWidget(self.start_btn)
        
        self.stop_btn = QPushButton("Остановить сервер")
        self.stop_btn.clicked.connect(self.stop_server)
        self.stop_btn.setEnabled(False)
        buttons_layout.addWidget(self.stop_btn)
        
        settings_layout.addLayout(buttons_layout)
        layout.addWidget(settings_group)
        
        # Логи
        log_group = QGroupBox("Логи сервера")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        font = QFont("Courier New", 9)
        self.log_text.setFont(font)
        log_layout.addWidget(self.log_text)
        layout.addWidget(log_group)
        
        # Кнопка очистки логов
        clear_btn = QPushButton("Очистить логи")
        clear_btn.clicked.connect(self.clear_logs)
        layout.addWidget(clear_btn)
    
    def init_tray(self):
        """Инициализация системного трея"""
        if QSystemTrayIcon.isSystemTrayAvailable():
            self.tray_icon = QSystemTrayIcon(self)
            # Создаем простую иконку если файл не существует
            if not os.path.exists('server.ico'):
                # Можно создать простую иконку программно или использовать стандартную
                self.tray_icon.setIcon(self.style().standardIcon(self.style().SP_ComputerIcon))
            else:
                self.tray_icon.setIcon(QIcon('server.ico'))
            
            tray_menu = QMenu()
            
            show_action = QAction("Показать", self)
            show_action.triggered.connect(self.show_window)
            tray_menu.addAction(show_action)
            
            start_action = QAction("Запустить сервер", self)
            start_action.triggered.connect(self.start_server)
            tray_menu.addAction(start_action)
            
            stop_action = QAction("Остановить сервер", self)
            stop_action.triggered.connect(self.stop_server)
            stop_action.setEnabled(False)
            tray_menu.addAction(stop_action)
            
            tray_menu.addSeparator()
            
            exit_action = QAction("Выход", self)
            exit_action.triggered.connect(self.quit_application)
            tray_menu.addAction(exit_action)
            
            self.tray_icon.setContextMenu(tray_menu)
            self.tray_icon.activated.connect(self.tray_icon_activated)
            self.tray_icon.show()
    
    def show_window(self):
        """Показать окно"""
        self.show()
        self.activateWindow()
        self.raise_()
    
    def tray_icon_activated(self, reason):
        """Обработка активации иконки в трее"""
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_window()
    
    def toggle_auto_start(self, state):
        """Включение/выключение автозапуска при запуске программы"""
        self.settings.setValue("auto_start", state == Qt.Checked)
    
    def toggle_system_auto_start(self, state):
        """Включение/выключение автозапуска при запуске системы"""
        if state == Qt.Checked:
            if AutoStartManager.enable_auto_start():
                self.add_log("Автозапуск при запуске системы включен")
            else:
                self.system_auto_start_cb.setChecked(False)
                self.add_log("Ошибка включения автозапуска при запуске системы")
        else:
            if AutoStartManager.disable_auto_start():
                self.add_log("Автозапуск при запуске системы выключен")
            else:
                self.system_auto_start_cb.setChecked(True)
                self.add_log("Ошибка выключения автозапуска при запуске системы")
    
    def start_server(self):
        """Запуск сервера"""
        self.server = ScheduleServer()
        self.server.signals.log_message.connect(self.add_log)
        self.server.signals.update_clients.connect(self.update_clients_count)
        
        self.server_thread = threading.Thread(target=self.server.start_server, daemon=True)
        self.server_thread.start()
        
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.status_label.setText("Запущен")
        self.status_label.setStyleSheet("color: green; font-weight: bold;")
        
        self.add_log("Сервер запущен")
        
        # Обновляем меню трея
        if self.tray_icon:
            tray_menu = self.tray_icon.contextMenu()
            tray_menu.actions()[1].setEnabled(False)  # Запустить сервер
            tray_menu.actions()[2].setEnabled(True)   # Остановить сервер
    def stop_server(self):
        """Остановка сервера"""
        if self.server:
            self.server.stop_server()
            self.server = None
        
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.status_label.setText("Остановлен")
        self.status_label.setStyleSheet("color: red; font-weight: bold;")
        self.clients_label.setText("0")
        
        self.add_log("Сервер остановлен")
        
        # Обновляем меню трея
        if self.tray_icon:
            tray_menu = self.tray_icon.contextMenu()
            tray_menu.actions()[1].setEnabled(True)   # Запустить сервер
            tray_menu.actions()[2].setEnabled(False)  # Остановить сервер
    
    def update_clients_count(self, count):
        """Обновление счетчика клиентов"""
        self.clients_label.setText(str(count))
    
    def add_log(self, message):
        """Добавление сообщения в лог"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def clear_logs(self):
        """Очистка логов"""
        self.log_text.clear()
    
    def closeEvent(self, event):
        """Обработка закрытия окна"""
        if self.tray_icon and self.tray_icon.isVisible():
            self.hide()
            self.tray_icon.showMessage(
                "Schedule Server",
                "Сервер продолжает работу в системном трее",
                QSystemTrayIcon.Information,
                2000
            )
            event.ignore()
        else:
            self.quit_application()
    
    def quit_application(self):
        """Полный выход из приложения"""
        self.stop_server()
        if self.tray_icon:
            self.tray_icon.hide()
        self.single_instance.cleanup()
        QApplication.quit()


def main():
    # Проверяем, не запущен ли уже экземпляр приложения ДО создания QApplication
    if check_already_running():
        sys.exit(1)
    
    # Создаем QApplication только если проверка пройдена
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    
    # Создаем экземпляр SingleInstance для этого приложения
    single_instance = SingleInstance()
    
    # Запускаем прослушивание сообщений
    if not single_instance.start_listener():
        QMessageBox.warning(None, "Schedule Server", 
                          "Не удалось запустить слушатель сообщений.")
        sys.exit(1)
    
    window = ServerWindow(single_instance)
    single_instance.start_listener(window.handle_instance_message)
    
    # Автозапуск сервера при запуске программы
    if window.settings.value("auto_start", True, type=bool):
        QTimer.singleShot(100, window.start_server)
    
    # Показываем окно только если не запущен с параметром --minimized
    if '--minimized' not in sys.argv:
        window.show()
    
    try:
        result = app.exec_()
        single_instance.cleanup()
        return result
    except Exception as e:
        single_instance.cleanup()
        raise e


if __name__ == "__main__":
    main()