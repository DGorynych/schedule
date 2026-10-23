import logging
import os
import socket
import sqlite3
import sys
import threading
from datetime import datetime

from PyQt5.QtCore import QObject, QSettings, Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QCheckBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSystemTrayIcon,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from schedule_common import (
    STATUS_EMPTY,
    ProtocolError,
    days_in_month,
    local_data_file,
    migrate_period_names,
    normalize_period_name,
    recv_message,
    resource_path,
    send_message,
    setup_logging,
    sort_periods,
)

logger = logging.getLogger(__name__)

try:
    import winreg  # type: ignore
except ImportError:  # не Windows
    winreg = None


class ProcessLock:
    """Файл-блокировка с PID, чтобы не запускать второй экземпляр."""

    def __init__(self, lock_name="schedule_server.lock", port=48888):
        self.lock_name = lock_name
        self.port = port
        self.lock_file = self.get_temp_path(lock_name)

    @staticmethod
    def get_temp_path(filename):
        try:
            return os.path.join(os.gettempdir(), filename)
        except Exception:
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

    def is_already_running(self):
        if not os.path.exists(self.lock_file):
            return False
        try:
            with open(self.lock_file, "r", encoding="utf-8") as handle:
                line = handle.read().strip()
            pid_text = line.split(":", 1)[0]
            pid = int(pid_text)
        except (OSError, ValueError):
            self._remove_lock()
            return False

        if self.is_process_running(pid):
            return True
        self._remove_lock()
        return False

    def _remove_lock(self):
        try:
            os.remove(self.lock_file)
        except OSError:
            pass

    @staticmethod
    def is_process_running(pid):
        if pid <= 0:
            return False
        try:
            if os.name == "nt":
                import ctypes

                kernel32 = ctypes.windll.kernel32
                process_query_limited_information = 0x1000
                handle = kernel32.OpenProcess(
                    process_query_limited_information, False, pid
                )
                if not handle:
                    # Процесс мог быть открыт с другими правами — проверяем факт наличия.
                    handle = kernel32.OpenProcess(0x0400, False, pid)
                if not handle:
                    return False
                try:
                    exit_code = ctypes.c_ulong()
                    if kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                        return exit_code.value == 259  # STILL_ACTIVE
                    return True
                finally:
                    kernel32.CloseHandle(handle)
            else:
                os.kill(pid, 0)
                return True
        except (OSError, ValueError):
            return False
        except Exception:
            return False

    def create_lock(self):
        """Атомарно создаёт файл блокировки. False — файл уже существует."""
        try:
            fd = os.open(self.lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            return False
        except OSError:
            return False
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as handle:
                handle.write(f"{os.getpid()}:{self.port}")
            return True
        except OSError:
            self._remove_lock()
            return False

    def release_lock(self):
        self._remove_lock()


class _MessageBridge(QObject):
    """Передаёт сообщение из сетевого потока в GUI-поток."""

    message = pyqtSignal(str)


class SingleInstance:
    """Один экземпляр сервера + приём команды «показать окно»."""

    def __init__(self, port=48888):
        self.port = port
        self.socket = None
        self.listener_thread = None
        self.is_primary = False
        self.process_lock = ProcessLock(port=port)
        self._bridge = _MessageBridge()
        self._callback_connected = False

    def is_already_running(self):
        if self.process_lock.is_already_running():
            return True
        return self.can_connect_to_socket()

    def can_connect_to_socket(self):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as test_socket:
                test_socket.settimeout(1)
                return test_socket.connect_ex(("localhost", self.port)) == 0
        except OSError:
            return False

    def send_message_to_running_instance(self, message="show"):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as client_socket:
                client_socket.settimeout(2)
                client_socket.connect(("localhost", self.port))
                client_socket.sendall(message.encode("utf-8"))
            return True
        except OSError:
            return False

    def start_listener(self, callback=None):
        if self.is_primary:
            return True
        try:
            self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.socket.bind(("localhost", self.port))
            self.socket.listen(1)
            self.socket.settimeout(1)
        except OSError as exc:
            logger.exception("Не удалось запустить слушатель: %s", exc)
            self._close_socket()
            return False

        if callback and not self._callback_connected:
            self._bridge.message.connect(callback)
            self._callback_connected = True

        self.is_primary = True
        self.listener_thread = threading.Thread(
            target=self._listen_for_messages, name="instance-listener", daemon=True
        )
        self.listener_thread.start()

        if not self.process_lock.create_lock():
            logger.warning("Файл блокировки уже существует: %s", self.process_lock.lock_file)

        return True

    def _listen_for_messages(self):
        while self.is_primary and self.socket:
            try:
                client, _addr = self.socket.accept()
                with client:
                    data = client.recv(1024).decode("utf-8", errors="replace").strip()
                if data:
                    # Сигнал доставит сообщение в главный поток.
                    self._bridge.message.emit(data)
            except socket.timeout:
                continue
            except OSError:
                break

    def _close_socket(self):
        if self.socket:
            try:
                self.socket.close()
            except OSError:
                pass
            self.socket = None

    def cleanup(self):
        self.is_primary = False
        self._close_socket()
        self.process_lock.release_lock()


def check_already_running():
    """Проверяет второй экземпляр. Вызывать только после создания QApplication."""
    single_instance = SingleInstance()
    if not single_instance.is_already_running():
        return False

    if single_instance.send_message_to_running_instance("show"):
        QMessageBox.information(
            None, "Schedule Server", "Приложение уже запущено. Окно будет показано."
        )
    else:
        QMessageBox.warning(
            None,
            "Schedule Server",
            "Приложение уже запущено, но не отвечает.\n\n"
            "Если вы уверены, что приложение не запущено, удалите файл блокировки "
            f"и попробуйте снова:\n{single_instance.process_lock.lock_file}",
        )
    return True


class ServerSignals(QObject):
    log_message = pyqtSignal(str)
    update_clients = pyqtSignal(int)
    server_status = pyqtSignal(bool, str)


class ScheduleServer:
    """Сетевой сервер и слой доступа к SQLite."""

    def __init__(self, host="0.0.0.0", port=8888, db_path=None):
        self.host = host
        self.port = port
        self.running = False
        self.server_socket = None
        self.clients = []
        self._clients_lock = threading.Lock()
        self._client_threads = []
        self._db_lock = threading.RLock()
        self.signals = ServerSignals()
        self.db_path = db_path or local_data_file("schedule_server.db")
        self.conn = sqlite3.connect(self.db_path, check_same_thread=False, timeout=10)
        self.conn.row_factory = sqlite3.Row
        self._configure_database()
        self.init_database()

    # ------------------------------------------------------------------
    # База данных
    # ------------------------------------------------------------------

    def _configure_database(self):
        with self._db_lock:
            self.conn.execute("PRAGMA journal_mode=WAL")
            self.conn.execute("PRAGMA synchronous=NORMAL")
            self.conn.execute("PRAGMA busy_timeout=5000")
            self.conn.execute("PRAGMA foreign_keys=ON")

    def init_database(self):
        with self._db_lock:
            cursor = self.conn.cursor()
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS employees (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL,
                    position TEXT,
                    is_deleted INTEGER DEFAULT 0,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS periods (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period TEXT UNIQUE NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS schedules (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period_id INTEGER,
                    employee_name TEXT NOT NULL,
                    day_index INTEGER NOT NULL,
                    status INTEGER DEFAULT 4,
                    end_time TEXT,
                    worked_hours REAL,
                    note_type TEXT DEFAULT '',
                    FOREIGN KEY (period_id) REFERENCES periods (id),
                    UNIQUE (period_id, employee_name, day_index)
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS employee_order (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period TEXT NOT NULL,
                    employee_name TEXT NOT NULL,
                    sort_order INTEGER NOT NULL,
                    UNIQUE (period, employee_name)
                )
                """
            )
            cursor.execute(
                "CREATE INDEX IF NOT EXISTS idx_schedules_period ON schedules(period_id)"
            )

            self._ensure_column(cursor, "schedules", "note_type", "TEXT DEFAULT ''")
            self._ensure_column(cursor, "employees", "is_deleted", "INTEGER DEFAULT 0")
            self._ensure_column(cursor, "employees", "position", "TEXT DEFAULT ''")
            self.conn.commit()

            migrate_period_names(self.conn)

        logger.info("База данных сервера: %s", self.db_path)

    @staticmethod
    def _ensure_column(cursor, table, column, definition):
        cursor.execute(f"PRAGMA table_info({table})")
        columns = {row[1] for row in cursor.fetchall()}
        if column not in columns:
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")

    def _period_id(self, period, create=True):
        canonical = normalize_period_name(period) or (period or "").strip()
        if not canonical:
            raise ValueError("Пустой период")
        cursor = self.conn.cursor()
        cursor.execute("SELECT id FROM periods WHERE period = ?", (canonical,))
        row = cursor.fetchone()
        if row:
            return row[0], canonical
        if not create:
            return None, canonical
        cursor.execute("INSERT INTO periods (period) VALUES (?)", (canonical,))
        return cursor.lastrowid, canonical

    def _ensure_employee_order(self, period, names):
        """Добавляет отсутствующих сотрудников в конец порядка периода."""
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT employee_name, sort_order FROM employee_order WHERE period = ?",
            (period,),
        )
        existing = {row[0] for row in cursor.fetchall()}
        cursor.execute(
            "SELECT MAX(sort_order) FROM employee_order WHERE period = ?",
            (period,),
        )
        max_order = cursor.fetchone()[0]
        next_order = (max_order if max_order is not None else -1) + 1
        rows = []
        for name in names:
            if name and name not in existing:
                rows.append((period, name, next_order))
                existing.add(name)
                next_order += 1
        if rows:
            cursor.executemany(
                "INSERT OR IGNORE INTO employee_order (period, employee_name, sort_order) "
                "VALUES (?, ?, ?)",
                rows,
            )

    # ------------------------------------------------------------------
    # Сеть
    # ------------------------------------------------------------------

    def start_server(self):
        try:
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.server_socket.bind((self.host, self.port))
            self.server_socket.listen(10)
            # Таймаут нужен, чтобы цикл accept замечал остановку сервера.
            self.server_socket.settimeout(1.0)
        except OSError as exc:
            message = f"Не удалось запустить сервер на {self.host}:{self.port}: {exc}"
            logger.exception(message)
            self.running = False
            self._close_socket()
            self._close_database()
            self.signals.server_status.emit(False, message)
            return

        self.running = True
        message = f"Сервер запущен на {self.host}:{self.port}"
        logger.info(message)
        self.signals.log_message.emit(message)
        self.signals.server_status.emit(True, message)

        while self.running:
            try:
                client_socket, address = self.server_socket.accept()
            except socket.timeout:
                continue
            except OSError as exc:
                if self.running:
                    logger.exception("Ошибка accept: %s", exc)
                break
            thread = threading.Thread(
                target=self.handle_client,
                args=(client_socket, address),
                name=f"client-{address[0]}:{address[1]}",
                daemon=True,
            )
            with self._clients_lock:
                self._client_threads = [
                    t for t in self._client_threads if t.is_alive()
                ]
                self._client_threads.append(thread)
            thread.start()

        self._close_socket()

    def stop_server(self):
        self.running = False
        self._close_socket()

        with self._clients_lock:
            clients = list(self.clients)
        for client_socket, _ in clients:
            try:
                client_socket.shutdown(socket.SHUT_RDWR)
            except OSError:
                pass
            try:
                client_socket.close()
            except OSError:
                pass
        with self._clients_lock:
            self.clients.clear()
            threads = list(self._client_threads)
        self.signals.update_clients.emit(0)

        for thread in threads:
            if thread.is_alive():
                thread.join(timeout=1.0)

        self._close_database()
        message = "Сервер остановлен"
        logger.info(message)
        self.signals.log_message.emit(message)
        self.signals.server_status.emit(False, message)

    def _close_socket(self):
        if self.server_socket:
            try:
                self.server_socket.close()
            except OSError:
                pass
            self.server_socket = None

    def _close_database(self):
        with self._db_lock:
            if self.conn is not None:
                try:
                    self.conn.close()
                except sqlite3.Error:
                    pass
                self.conn = None

    def handle_client(self, client_socket, address):
        logger.info("Клиент подключен: %s", address)
        with self._clients_lock:
            self.clients.append((client_socket, address))
            count = len(self.clients)
        self.signals.update_clients.emit(count)
        self.signals.log_message.emit(f"Клиент подключен: {address}")

        try:
            client_socket.settimeout(60)
            while self.running:
                request = recv_message(client_socket)
                if request is None:
                    break
                response = self.process_request(request)
                send_message(client_socket, response)
        except ProtocolError as exc:
            logger.warning("Ошибка протокола с %s: %s", address, exc)
            self.signals.log_message.emit(f"Ошибка протокола с {address}: {exc}")
        except (OSError, socket.timeout) as exc:
            logger.info("Соединение с %s прервано: %s", address, exc)
        except Exception as exc:
            logger.exception("Ошибка с клиентом %s", address)
            self.signals.log_message.emit(f"Ошибка с клиентом {address}: {exc}")
        finally:
            with self._clients_lock:
                if (client_socket, address) in self.clients:
                    self.clients.remove((client_socket, address))
                count = len(self.clients)
            try:
                client_socket.close()
            except OSError:
                pass
            self.signals.update_clients.emit(count)
            self.signals.log_message.emit(f"Клиент отключен: {address}")

    # ------------------------------------------------------------------
    # Обработка запросов
    # ------------------------------------------------------------------

    def process_request(self, request):
        action = request.get("action")
        handlers = {
            "ping": lambda _r: {"status": "success", "message": "pong"},
            "get_employees": lambda _r: self.get_employees(),
            "get_periods": lambda _r: self.get_periods(),
            "get_schedule": lambda r: self.get_schedule(r.get("period")),
            "save_schedule": lambda r: self.save_schedule(r.get("period"), r.get("data") or {}),
            "add_employee": lambda r: self.add_employee(r.get("name"), r.get("position", "")),
            "update_employee": lambda r: self.update_employee(
                r.get("old_name"), r.get("new_name"), r.get("position", "")
            ),
            "delete_employee": lambda r: self.set_employee_deleted(r.get("name"), True),
            "restore_employee": lambda r: self.set_employee_deleted(r.get("name"), False),
            "set_employee_order": lambda r: self.set_employee_order(
                r.get("period"), r.get("employees") or []
            ),
            "create_period": lambda r: self.create_period(
                r.get("period"), r.get("employees") or []
            ),
            "add_employees_to_period": lambda r: self.add_employees_to_period(
                r.get("period"), r.get("employees") or []
            ),
            "delete_period": lambda r: self.delete_period(r.get("period")),
            "rename_period": lambda r: self.rename_period(r.get("old_period"), r.get("period")),
            "sync_data": lambda r: self.sync_data(r.get("data") or {}),
        }

        handler = handlers.get(action)
        if handler is None:
            return {"status": "error", "message": "Неизвестное действие"}

        try:
            with self._db_lock:
                if self.conn is None:
                    return {"status": "error", "message": "База данных закрыта"}
                return handler(request)
        except Exception as exc:
            logger.exception("Ошибка обработки действия %s", action)
            try:
                self.conn.rollback()
            except sqlite3.Error:
                pass
            return {"status": "error", "message": str(exc)}

    # ------------------------------------------------------------------
    # Сотрудники
    # ------------------------------------------------------------------

    def get_employees(self):
        cursor = self.conn.cursor()
        cursor.execute(
            """
            SELECT name, COALESCE(position, ''), COALESCE(is_deleted, 0)
            FROM employees ORDER BY name COLLATE NOCASE
            """
        )
        data = [
            {"name": row[0], "position": row[1], "is_deleted": row[2]}
            for row in cursor.fetchall()
        ]
        return {"status": "success", "data": data}

    def add_employee(self, name, position=""):
        name = (name or "").strip()
        if not name:
            return {"status": "error", "message": "Пустое имя сотрудника"}
        cursor = self.conn.cursor()
        # SQLite NOCASE не понимает кириллицу, поэтому сравниваем в Python.
        cursor.execute("SELECT id, name, is_deleted FROM employees")
        rows = cursor.fetchall()
        for employee_id, existing_name, is_deleted in rows:
            if (existing_name or "").casefold() == name.casefold():
                if is_deleted:
                    cursor.execute(
                        "UPDATE employees SET is_deleted = 0 WHERE id = ?",
                        (employee_id,),
                    )
                    self.conn.commit()
                    return {"status": "success", "message": "Сотрудник восстановлен"}
                return {"status": "success", "message": "Сотрудник уже существует"}

        cursor.execute(
            "INSERT INTO employees (name, position, is_deleted) VALUES (?, ?, 0)",
            (name, (position or "").strip()),
        )
        self.conn.commit()
        return {"status": "success", "message": "Сотрудник добавлен"}

    def update_employee(self, old_name, new_name, position=""):
        old_name = (old_name or "").strip()
        new_name = (new_name or "").strip()
        if not old_name or not new_name:
            return {"status": "error", "message": "Пустое имя сотрудника"}

        cursor = self.conn.cursor()
        cursor.execute("SELECT id, name FROM employees")
        rows = cursor.fetchall()
        old_row = next(
            ((employee_id, existing_name) for employee_id, existing_name in rows
             if (existing_name or "").casefold() == old_name.casefold()),
            None,
        )
        if not old_row:
            return {"status": "error", "message": "Сотрудник не найден"}
        old_id, actual_old_name = old_row

        if new_name.casefold() != old_name.casefold():
            for employee_id, existing_name in rows:
                if employee_id != old_id and (existing_name or "").casefold() == new_name.casefold():
                    return {"status": "error", "message": "Сотрудник с таким именем уже существует"}

        cursor.execute(
            "UPDATE employees SET name = ?, position = ? WHERE id = ?",
            (new_name, (position or "").strip(), old_id),
        )
        cursor.execute(
            "UPDATE schedules SET employee_name = ? WHERE employee_name = ?",
            (new_name, actual_old_name),
        )
        cursor.execute(
            "UPDATE employee_order SET employee_name = ? WHERE employee_name = ?",
            (new_name, actual_old_name),
        )
        self.conn.commit()
        return {"status": "success", "message": "Данные сотрудника обновлены"}

    def set_employee_deleted(self, name, deleted):
        name = (name or "").strip()
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, name FROM employees")
        row = next(
            ((employee_id, existing_name) for employee_id, existing_name in cursor.fetchall()
             if (existing_name or "").casefold() == name.casefold()),
            None,
        )
        if not row:
            return {"status": "error", "message": "Сотрудник не найден"}
        employee_id, _actual_name = row
        cursor.execute(
            "UPDATE employees SET is_deleted = ? WHERE id = ?",
            (1 if deleted else 0, employee_id),
        )
        self.conn.commit()
        return {
            "status": "success",
            "message": "Сотрудник помечен как удаленный" if deleted else "Сотрудник восстановлен",
        }

    def set_employee_order(self, period, employees):
        canonical = normalize_period_name(period) or (period or "").strip()
        if not canonical:
            return {"status": "error", "message": "Пустой период"}
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM employee_order WHERE period = ?", (canonical,))
        cursor.executemany(
            "INSERT INTO employee_order (period, employee_name, sort_order) VALUES (?, ?, ?)",
            [(canonical, name, index) for index, name in enumerate(employees) if name],
        )
        self.conn.commit()
        return {"status": "success", "message": "Порядок сотрудников сохранен"}

    # ------------------------------------------------------------------
    # Периоды и расписания
    # ------------------------------------------------------------------

    def get_periods(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT period FROM periods")
        periods = []
        seen = set()
        for (period,) in cursor.fetchall():
            canonical = normalize_period_name(period) or period
            if canonical not in seen:
                seen.add(canonical)
                periods.append(canonical)
        return {"status": "success", "data": sort_periods(periods)}

    def get_schedule(self, period):
        canonical = normalize_period_name(period) or (period or "").strip()
        if not canonical:
            return {"status": "error", "message": "Пустой период"}

        period_id, _ = self._period_id(canonical, create=False)
        if period_id is None:
            return {
                "status": "success",
                "data": {"employees": [], "schedule": {}, "notes": {}},
            }

        cursor = self.conn.cursor()
        cursor.execute(
            """
            SELECT eo.employee_name, COALESCE(e.position, '')
            FROM employee_order AS eo
            LEFT JOIN employees AS e ON e.name = eo.employee_name
            WHERE eo.period = ? AND COALESCE(e.is_deleted, 0) = 0
            ORDER BY eo.sort_order
            """,
            (canonical,),
        )
        employees = [{"name": row[0], "position": row[1]} for row in cursor.fetchall()]
        known_names = {employee["name"] for employee in employees}

        # Добавляем сотрудников, у которых есть строки расписания, но нет
        # записи в порядке (данные старых версий или частичная синхронизация).
        cursor.execute(
            """
            SELECT DISTINCT s.employee_name, COALESCE(e.position, '')
            FROM schedules AS s
            LEFT JOIN employees AS e ON e.name = s.employee_name
            WHERE s.period_id = ? AND COALESCE(e.is_deleted, 0) = 0
            ORDER BY s.employee_name COLLATE NOCASE
            """,
            (period_id,),
        )
        for name, position in cursor.fetchall():
            if name not in known_names:
                employees.append({"name": name, "position": position})
                known_names.add(name)

        active_names = {employee["name"] for employee in employees}
        length = days_in_month(canonical)
        schedule = {name: [STATUS_EMPTY] * length for name in active_names}
        notes = {}

        cursor.execute(
            """
            SELECT employee_name, day_index, status, end_time, worked_hours, note_type
            FROM schedules WHERE period_id = ?
            """,
            (period_id,),
        )
        for emp_name, day_index, status, end_time, worked_hours, note_type in cursor.fetchall():
            if emp_name not in active_names or not 0 <= day_index < length:
                continue
            schedule[emp_name][day_index] = status
            if end_time or worked_hours or note_type:
                notes.setdefault(emp_name, {})[str(day_index)] = {
                    "end_time": end_time,
                    "worked_hours": worked_hours,
                    "note_type": note_type or "",
                }

        return {
            "status": "success",
            "data": {"employees": employees, "schedule": schedule, "notes": notes},
        }

    def _upsert_schedule_rows(self, period_id, employee_name, schedule, notes):
        rows = []
        notes = notes or {}
        for day_index, status in enumerate(schedule):
            note = notes.get(str(day_index), {}) or {}
            try:
                status_value = int(status)
            except (TypeError, ValueError):
                status_value = STATUS_EMPTY
            if status_value not in (0, 1, 2, 3, 4):
                status_value = STATUS_EMPTY
            rows.append((
                period_id,
                employee_name,
                day_index,
                status_value,
                note.get("end_time"),
                note.get("worked_hours"),
                note.get("note_type", ""),
            ))
        return rows

    def save_schedule(self, period, data):
        canonical = normalize_period_name(period) or (period or "").strip()
        if not canonical:
            return {"status": "error", "message": "Пустой период"}

        period_id, canonical = self._period_id(canonical)
        employees = data.get("employees") or []
        names = [employee.get("name") for employee in employees if employee.get("name")]
        if names:
            self._ensure_employee_order(canonical, names)

        rows = []
        length = days_in_month(canonical)
        for emp_name, schedule in (data.get("schedule") or {}).items():
            values = list(schedule or [])[:length]
            values.extend([STATUS_EMPTY] * (length - len(values)))
            rows.extend(
                self._upsert_schedule_rows(
                    period_id, emp_name, values, (data.get("notes") or {}).get(emp_name)
                )
            )

        if rows:
            self.conn.executemany(
                """
                INSERT INTO schedules
                    (period_id, employee_name, day_index, status, end_time, worked_hours, note_type)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT (period_id, employee_name, day_index) DO UPDATE SET
                    status = excluded.status,
                    end_time = excluded.end_time,
                    worked_hours = excluded.worked_hours,
                    note_type = excluded.note_type
                """,
                rows,
            )
        self.conn.commit()
        logger.info("Сохранено расписание %s: %s строк", canonical, len(rows))
        return {"status": "success", "message": "Расписание сохранено"}

    def create_period(self, period, employees):
        canonical = normalize_period_name(period) or (period or "").strip()
        if not canonical:
            return {"status": "error", "message": "Пустой период"}

        period_id, canonical = self._period_id(canonical)
        names = [name for name in employees if name]
        self._ensure_employee_order(canonical, names)

        days = days_in_month(canonical)
        rows = [
            (period_id, name, day_index, STATUS_EMPTY)
            for name in names
            for day_index in range(days)
        ]
        if rows:
            self.conn.executemany(
                "INSERT OR IGNORE INTO schedules "
                "(period_id, employee_name, day_index, status) VALUES (?, ?, ?, ?)",
                rows,
            )
        self.conn.commit()
        return {"status": "success", "message": "Период создан"}

    def add_employees_to_period(self, period, employees):
        canonical = normalize_period_name(period) or (period or "").strip()
        if not canonical:
            return {"status": "error", "message": "Пустой период"}

        period_id, canonical = self._period_id(canonical)
        names = [name for name in employees if name]
        self._ensure_employee_order(canonical, names)

        days = days_in_month(canonical)
        rows = [
            (period_id, name, day_index, STATUS_EMPTY)
            for name in names
            for day_index in range(days)
        ]
        if rows:
            self.conn.executemany(
                "INSERT OR IGNORE INTO schedules "
                "(period_id, employee_name, day_index, status) VALUES (?, ?, ?, ?)",
                rows,
            )
        self.conn.commit()
        return {"status": "success", "message": "Сотрудники добавлены"}

    def delete_period(self, period):
        canonical = normalize_period_name(period) or (period or "").strip()
        period_id, canonical = self._period_id(canonical, create=False)
        if period_id is None:
            # Идемпотентно: повторное удаление не должно «застревать» в очереди.
            return {"status": "success", "message": "Период не найден"}

        self.conn.execute("DELETE FROM schedules WHERE period_id = ?", (period_id,))
        self.conn.execute("DELETE FROM employee_order WHERE period = ?", (canonical,))
        self.conn.execute("DELETE FROM periods WHERE id = ?", (period_id,))
        self.conn.commit()
        return {"status": "success", "message": "Период удален"}

    def rename_period(self, old_period, new_period):
        old_canonical = normalize_period_name(old_period) or (old_period or "").strip()
        new_canonical = normalize_period_name(new_period) or (new_period or "").strip()
        if not old_canonical or not new_canonical:
            return {"status": "error", "message": "Пустой период"}

        cursor = self.conn.cursor()
        cursor.execute("SELECT id FROM periods WHERE period = ?", (old_canonical,))
        if not cursor.fetchone():
            return {"status": "error", "message": "Период не найден"}
        cursor.execute("SELECT id FROM periods WHERE period = ?", (new_canonical,))
        if cursor.fetchone():
            return {"status": "error", "message": "Период с таким названием уже существует"}

        cursor.execute(
            "UPDATE periods SET period = ? WHERE period = ?", (new_canonical, old_canonical)
        )
        cursor.execute(
            "UPDATE employee_order SET period = ? WHERE period = ?",
            (new_canonical, old_canonical),
        )
        self.conn.commit()
        return {"status": "success", "message": "Период переименован"}

    def sync_data(self, data):
        for employee in data.get("employees") or []:
            name = (employee.get("name") or "").strip()
            if not name:
                continue
            self.conn.execute(
                """
                INSERT INTO employees (name, position, is_deleted)
                VALUES (?, ?, ?)
                ON CONFLICT (name) DO UPDATE SET
                    position = excluded.position,
                    is_deleted = excluded.is_deleted
                """,
                (
                    name,
                    employee.get("position", ""),
                    1 if employee.get("is_deleted") else 0,
                ),
            )

        for period in data.get("periods") or []:
            canonical = normalize_period_name(period) or (period or "").strip()
            if canonical:
                self.conn.execute(
                    "INSERT OR IGNORE INTO periods (period) VALUES (?)", (canonical,)
                )

        for schedule_data in data.get("schedules") or []:
            period = schedule_data.get("period")
            if not period:
                continue
            period_id, canonical = self._period_id(period)
            length = days_in_month(canonical)
            names = list((schedule_data.get("schedule") or {}).keys())
            self._ensure_employee_order(canonical, names)
            for emp_name, schedule in (schedule_data.get("schedule") or {}).items():
                values = list(schedule or [])[:length]
                values.extend([STATUS_EMPTY] * (length - len(values)))
                rows = self._upsert_schedule_rows(
                    period_id,
                    emp_name,
                    values,
                    (schedule_data.get("notes") or {}).get(emp_name),
                )
                if rows:
                    self.conn.executemany(
                        """
                        INSERT INTO schedules
                            (period_id, employee_name, day_index, status, end_time,
                             worked_hours, note_type)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT (period_id, employee_name, day_index) DO UPDATE SET
                            status = excluded.status,
                            end_time = excluded.end_time,
                            worked_hours = excluded.worked_hours,
                            note_type = excluded.note_type
                        """,
                        rows,
                    )

        self.conn.commit()
        return {"status": "success", "message": "Данные синхронизированы"}


class AutoStartManager:
    """Автозапуск через HKEY_CURRENT_USER\\...\\Run (Windows)."""

    @staticmethod
    def get_registry_path():
        return r"Software\Microsoft\Windows\CurrentVersion\Run"

    @staticmethod
    def get_app_name():
        return "ScheduleServer"

    @staticmethod
    def get_command():
        if getattr(sys, "frozen", False):
            return f'"{sys.executable}" --minimized'
        python = sys.executable
        pythonw = os.path.join(os.path.dirname(python), "pythonw.exe")
        if not os.path.exists(pythonw):
            pythonw = python
        script = os.path.abspath(__file__)
        return f'"{pythonw}" "{script}" --minimized'

    @staticmethod
    def is_auto_start_enabled():
        if winreg is None:
            return False
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AutoStartManager.get_registry_path()) as key:
                try:
                    winreg.QueryValueEx(key, AutoStartManager.get_app_name())
                    return True
                except FileNotFoundError:
                    return False
        except OSError:
            return False

    @staticmethod
    def enable_auto_start():
        if winreg is None:
            return False
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                AutoStartManager.get_registry_path(),
                0,
                winreg.KEY_SET_VALUE,
            ) as key:
                winreg.SetValueEx(
                    key,
                    AutoStartManager.get_app_name(),
                    0,
                    winreg.REG_SZ,
                    AutoStartManager.get_command(),
                )
            return True
        except OSError:
            logger.exception("Не удалось включить автозапуск")
            return False

    @staticmethod
    def disable_auto_start():
        if winreg is None:
            return False
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                AutoStartManager.get_registry_path(),
                0,
                winreg.KEY_SET_VALUE,
            ) as key:
                winreg.DeleteValue(key, AutoStartManager.get_app_name())
            return True
        except OSError:
            logger.exception("Не удалось выключить автозапуск")
            return False


class ServerWindow(QMainWindow):
    def __init__(self, single_instance):
        super().__init__()
        self.single_instance = single_instance
        self.server = None
        self.server_thread = None
        self.tray_icon = None
        self.settings = QSettings("MyCompany", "ScheduleServer")
        self._stopping = False
        self.init_ui()
        self.init_tray()

        if "--minimized" in sys.argv:
            self.hide()
            if self.tray_icon:
                self.tray_icon.showMessage(
                    "Schedule Server",
                    "Сервер запущен в фоновом режиме",
                    QSystemTrayIcon.Information,
                    2000,
                )

    def handle_instance_message(self, message):
        if message == "show":
            self.show_window()

    def init_ui(self):
        self.setWindowTitle("Schedule Server")
        self.setGeometry(300, 300, 800, 600)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        settings_group = QGroupBox("Настройки сервера")
        settings_layout = QVBoxLayout(settings_group)

        status_layout = QHBoxLayout()
        status_layout.addWidget(QLabel("Статус сервера:"))
        self.status_label = QLabel("Остановлен")
        self.status_label.setStyleSheet("color: red; font-weight: bold;")
        status_layout.addWidget(self.status_label)
        status_layout.addStretch()
        settings_layout.addLayout(status_layout)

        clients_layout = QHBoxLayout()
        clients_layout.addWidget(QLabel("Подключенные клиенты:"))
        self.clients_label = QLabel("0")
        clients_layout.addWidget(self.clients_label)
        clients_layout.addStretch()
        settings_layout.addLayout(clients_layout)

        self.auto_start_cb = QCheckBox("Автозапуск сервера при запуске программы")
        self.auto_start_cb.setChecked(self.settings.value("auto_start", True, type=bool))
        self.auto_start_cb.stateChanged.connect(self.toggle_auto_start)
        settings_layout.addWidget(self.auto_start_cb)

        self.system_auto_start_cb = QCheckBox("Автозапуск сервера при запуске системы")
        self.system_auto_start_cb.setChecked(AutoStartManager.is_auto_start_enabled())
        self.system_auto_start_cb.stateChanged.connect(self.toggle_system_auto_start)
        settings_layout.addWidget(self.system_auto_start_cb)

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

        log_group = QGroupBox("Логи сервера")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.document().setMaximumBlockCount(5000)
        self.log_text.setFont(QFont("Courier New", 9))
        log_layout.addWidget(self.log_text)
        layout.addWidget(log_group)

        clear_btn = QPushButton("Очистить логи")
        clear_btn.clicked.connect(self.clear_logs)
        layout.addWidget(clear_btn)

    def init_tray(self):
        if not QSystemTrayIcon.isSystemTrayAvailable():
            return
        self.tray_icon = QSystemTrayIcon(self)
        icon = QIcon(resource_path("schedule.ico"))
        if icon.isNull():
            icon = self.style().standardIcon(self.style().SP_ComputerIcon)
        self.tray_icon.setIcon(icon)

        tray_menu = QMenu()
        show_action = QAction("Показать", self)
        show_action.triggered.connect(self.show_window)
        tray_menu.addAction(show_action)

        self.tray_start_action = QAction("Запустить сервер", self)
        self.tray_start_action.triggered.connect(self.start_server)
        tray_menu.addAction(self.tray_start_action)

        self.tray_stop_action = QAction("Остановить сервер", self)
        self.tray_stop_action.triggered.connect(self.stop_server)
        self.tray_stop_action.setEnabled(False)
        tray_menu.addAction(self.tray_stop_action)

        tray_menu.addSeparator()
        exit_action = QAction("Выход", self)
        exit_action.triggered.connect(self.quit_application)
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()

    def show_window(self):
        self.show()
        self.activateWindow()
        self.raise_()

    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_window()

    def toggle_auto_start(self, state):
        self.settings.setValue("auto_start", state == Qt.Checked)

    def toggle_system_auto_start(self, state):
        if state == Qt.Checked:
            if AutoStartManager.enable_auto_start():
                logger.info("Автозапуск при запуске системы включен")
                self.add_log("Автозапуск при запуске системы включен")
            else:
                self.system_auto_start_cb.setChecked(False)
                self.add_log("Ошибка включения автозапуска при запуске системы")
        else:
            if AutoStartManager.disable_auto_start():
                logger.info("Автозапуск при запуске системы выключен")
                self.add_log("Автозапуск при запуске системы выключен")
            else:
                self.system_auto_start_cb.setChecked(True)
                self.add_log("Ошибка выключения автозапуска при запуске системы")

    def start_server(self):
        if self.server is not None and self.server.running:
            return
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.status_label.setText("Запуск...")
        self.status_label.setStyleSheet("color: orange; font-weight: bold;")

        self.server = None
        try:
            self.server = ScheduleServer(db_path=local_data_file("schedule_server.db"))
        except Exception as exc:
            logger.exception("Не удалось открыть базу данных сервера")
            self.status_label.setText("Ошибка запуска")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
            self.add_log(f"Не удалось открыть базу данных сервера: {exc}")
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            return
        self.server.signals.log_message.connect(self.add_log)
        self.server.signals.update_clients.connect(self.update_clients_count)
        self.server.signals.server_status.connect(self.on_server_status)

        self.server_thread = threading.Thread(
            target=self.server.start_server, name="schedule-server", daemon=True
        )
        self.server_thread.start()

    def stop_server(self):
        self._stopping = True
        try:
            if self.server is not None:
                self.server.stop_server()
                self.server = None
        finally:
            self._stopping = False
        self._set_running_ui(False)

    def on_server_status(self, running, message):
        if running:
            self.add_log(message)
            self._set_running_ui(True)
        else:
            self.add_log(message)
            self._set_running_ui(False)
            if not self._stopping and self.server is not None and not self.server.running:
                # Старт не удался: освобождаем объект, чтобы следующий старт был чистым.
                self.server = None

    def _set_running_ui(self, running):
        if running:
            self.status_label.setText("Запущен")
            self.status_label.setStyleSheet("color: green; font-weight: bold;")
            self.start_btn.setEnabled(False)
            self.stop_btn.setEnabled(True)
        else:
            self.status_label.setText("Остановлен")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.clients_label.setText("0")
        if self.tray_icon:
            self.tray_start_action.setEnabled(not running)
            self.tray_stop_action.setEnabled(running)

    def update_clients_count(self, count):
        self.clients_label.setText(str(count))

    def add_log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")

    def clear_logs(self):
        self.log_text.clear()

    def closeEvent(self, event):
        if self.tray_icon and self.tray_icon.isVisible():
            self.hide()
            self.tray_icon.showMessage(
                "Schedule Server",
                "Сервер продолжает работу в системном трее",
                QSystemTrayIcon.Information,
                2000,
            )
            event.ignore()
        else:
            self.quit_application()

    def quit_application(self):
        self.stop_server()
        if self.tray_icon:
            self.tray_icon.hide()
        self.single_instance.cleanup()
        QApplication.quit()


def main():
    setup_logging("schedule_server")
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    # QMessageBox требует QApplication, поэтому проверка — после его создания.
    if check_already_running():
        sys.exit(1)

    single_instance = SingleInstance()
    window = ServerWindow(single_instance)
    if not single_instance.start_listener(window.handle_instance_message):
        QMessageBox.warning(
            None, "Schedule Server", "Не удалось запустить слушатель сообщений."
        )
        sys.exit(1)

    if window.settings.value("auto_start", True, type=bool):
        QTimer.singleShot(100, window.start_server)

    if "--minimized" not in sys.argv:
        window.show()

    try:
        return app.exec_()
    finally:
        single_instance.cleanup()


if __name__ == "__main__":
    sys.exit(main())
