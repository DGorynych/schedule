import json
import logging
import os
import socket
import sqlite3
import sys
from collections import defaultdict
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PyQt5.QtCore import QDate, QSettings, QSignalBlocker, Qt, QTime, QTimer
from PyQt5.QtGui import QColor, QFont, QIcon, QKeySequence, QPainter, QPen, QPixmap
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QShortcut,
    QSplitter,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
)

from schedule_common import (
    DAY_NAMES,
    DEFAULT_END_TIME,
    DEFAULT_WORK_HOURS,
    MONTH_NAMES,
    STATUS_DAY_OFF,
    STATUS_EMPTY,
    WORK_STATUSES,
    ProtocolError,
    base_period_name,
    compute_worked_hours,
    days_in_month,
    days_in_month as month_length,
    format_note_text,
    has_meaningful_note,
    is_sunday,
    local_data_file,
    migrate_period_names,
    normalize_period_name,
    note_tooltip,
    parse_period,
    period_filename,
    period_sort_key,
    recv_message,
    resource_path,
    send_message,
    setup_logging,
    sort_periods,
    sunday_flags,
)
from ui_theme import apply_theme, set_pill_state, set_variant

logger = logging.getLogger(__name__)

# Действия, которые изменяют данные и могут быть отложены до подключения.
WRITE_ACTIONS = {
    "save_schedule",
    "add_employee",
    "update_employee",
    "delete_employee",
    "restore_employee",
    "set_employee_order",
    "create_period",
    "add_employees_to_period",
    "delete_period",
    "rename_period",
    "sync_data",
}


class NetworkManager:
    """Сетевое взаимодействие с сервером и офлайн-очередь изменений."""

    CONNECT_TIMEOUT = 3
    REQUEST_TIMEOUT = 15

    def __init__(self):
        self.server_host = "localhost"
        self.server_port = 8888
        self.connected = False
        self.socket = None
        self.local_db = local_data_file("local_schedule.db")
        self.on_connection_change = None
        self.init_pending_changes_table()

    def init_pending_changes_table(self):
        with sqlite3.connect(self.local_db, timeout=10) as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS pending_changes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    action TEXT NOT NULL,
                    data TEXT NOT NULL,
                    dedupe_key TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            cursor.execute("PRAGMA table_info(pending_changes)")
            columns = {row[1] for row in cursor.fetchall()}
            if "dedupe_key" not in columns:
                cursor.execute("ALTER TABLE pending_changes ADD COLUMN dedupe_key TEXT")
            self._compact_pending_changes(conn)
            conn.commit()

    def _compact_pending_changes(self, conn):
        """Чистит очередь от старых read-запросов и дублей.

        Прежние версии клиента складывали в очередь вообще все запросы
        (включая get_schedule/get_periods) и не схлопывали повторы. Здесь
        остаётся только последнее изменение на каждый объект — это резко
        уменьшает объём синхронизации при подключении.
        """
        rows = conn.execute(
            "SELECT id, action, data FROM pending_changes ORDER BY id"
        ).fetchall()
        latest = {}
        remove_ids = []
        for change_id, action, payload in rows:
            if action not in WRITE_ACTIONS:
                remove_ids.append(change_id)
                continue
            try:
                request = json.loads(payload)
            except (TypeError, ValueError, json.JSONDecodeError):
                request = {}
            if not isinstance(request, dict):
                request = {}
            request.setdefault("action", action)
            key = self._dedupe_key(request)
            previous = latest.get(key)
            if previous is not None:
                remove_ids.append(previous)
            latest[key] = change_id

        if remove_ids:
            conn.executemany(
                "DELETE FROM pending_changes WHERE id = ?",
                [(change_id,) for change_id in remove_ids],
            )
        for key, change_id in latest.items():
            conn.execute(
                "UPDATE pending_changes SET dedupe_key = ? WHERE id = ?",
                (key, change_id),
            )
        if remove_ids:
            logger.info("Очередь синхронизации сокращена на %s записей", len(remove_ids))

    def _set_connected(self, value):
        if self.connected == value:
            return
        self.connected = value
        callback = self.on_connection_change
        if callback:
            try:
                callback(value)
            except Exception:
                logger.exception("Ошибка обработчика состояния подключения")

    def _close_socket(self):
        if self.socket is not None:
            try:
                self.socket.close()
            except OSError:
                pass
            self.socket = None

    def connect(self, host, port):
        self.server_host = host
        self.server_port = port
        self._close_socket()
        try:
            self.socket = socket.create_connection(
                (host, port), timeout=self.CONNECT_TIMEOUT
            )
            self.socket.settimeout(self.REQUEST_TIMEOUT)
            self._set_connected(True)
            self.load_pending_changes()
            return self.connected
        except OSError as exc:
            logger.warning("Ошибка подключения к %s:%s: %s", host, port, exc)
            self._close_socket()
            self._set_connected(False)
            return False

    def disconnect(self):
        self._close_socket()
        self._set_connected(False)

    def _send_raw(self, request):
        if self.socket is None:
            raise ConnectionError("Нет соединения с сервером")
        send_message(self.socket, request)
        response = recv_message(self.socket)
        if response is None:
            raise ConnectionError("Сервер закрыл соединение")
        return response

    def send_request(self, request):
        """Отправляет запрос. Read-запросы в офлайн-очередь не попадают."""
        if not self.connected:
            if request.get("action") in WRITE_ACTIONS:
                self.save_pending_change(request)
                return {"status": "offline", "message": "Сохранено локально"}
            return {"status": "offline", "message": "Нет соединения с сервером"}

        try:
            return self._send_raw(request)
        except (OSError, ProtocolError, ConnectionError) as exc:
            logger.warning("Ошибка запроса %s: %s", request.get("action"), exc)
            self._close_socket()
            self._set_connected(False)
            if request.get("action") in WRITE_ACTIONS:
                self.save_pending_change(request)
            return {"status": "error", "message": str(exc)}

    @staticmethod
    def _dedupe_key(request):
        action = request.get("action", "")
        period = (
            normalize_period_name(request.get("period"))
            or normalize_period_name(request.get("old_period"))
            or request.get("period")
            or request.get("old_period")
            or ""
        )
        if action in {"save_schedule", "create_period", "add_employees_to_period",
                      "delete_period", "set_employee_order", "rename_period"}:
            return f"{action}:{period}"
        if action == "update_employee":
            return f"{action}:{request.get('old_name') or ''}"
        if action in {"add_employee", "delete_employee", "restore_employee"}:
            return f"{action}:{(request.get('name') or '').casefold()}"
        return f"{action}:{period}"

    def save_pending_change(self, request):
        """Сохраняет изменение, заменяя предыдущее изменение того же объекта."""
        key = self._dedupe_key(request)
        payload = json.dumps(request, ensure_ascii=False)
        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM pending_changes WHERE dedupe_key = ?", (key,))
                if request.get("action") == "delete_period":
                    # Удаление периода отменяет ещё не отправленное сохранение.
                    cursor.execute(
                        "DELETE FROM pending_changes WHERE dedupe_key = ?",
                        (
                            self._dedupe_key(
                                {
                                    "action": "save_schedule",
                                    "period": request.get("period"),
                                }
                            ),
                        ),
                    )
                cursor.execute(
                    "INSERT INTO pending_changes (action, data, dedupe_key) VALUES (?, ?, ?)",
                    (request.get("action"), payload, key),
                )
                conn.commit()
        except sqlite3.Error:
            logger.exception("Не удалось сохранить отложенное изменение")

    def load_pending_changes(self):
        """Отправляет накопленные изменения по порядку, без повторной постановки в очередь."""
        if not self.connected:
            return

        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT id, data FROM pending_changes ORDER BY id"
                )
                changes = cursor.fetchall()
        except sqlite3.Error:
            logger.exception("Не удалось прочитать отложенные изменения")
            return

        sent_ids = []
        for change_id, data in changes:
            try:
                request = json.loads(data)
            except (TypeError, ValueError, json.JSONDecodeError):
                logger.warning("Повреждённое отложенное изменение %s пропущено", change_id)
                sent_ids.append(change_id)
                continue
            if not isinstance(request, dict):
                logger.warning("Некорректное отложенное изменение %s пропущено", change_id)
                sent_ids.append(change_id)
                continue
            try:
                response = self._send_raw(request)
            except (OSError, ProtocolError, ConnectionError) as exc:
                logger.warning("Не удалось отправить отложенное изменение %s: %s", change_id, exc)
                self._close_socket()
                self._set_connected(False)
                break
            if response.get("status") == "success":
                sent_ids.append(change_id)
            else:
                logger.warning(
                    "Сервер отклонил отложенное изменение %s: %s",
                    change_id,
                    response.get("message"),
                )

        if sent_ids:
            try:
                with sqlite3.connect(self.local_db, timeout=10) as conn:
                    placeholders = ",".join("?" * len(sent_ids))
                    conn.execute(
                        f"DELETE FROM pending_changes WHERE id IN ({placeholders})",
                        sent_ids,
                    )
                    conn.commit()
            except sqlite3.Error:
                logger.exception("Не удалось очистить отправленные изменения")


class ScheduleManager:
    def __init__(self, network_manager):
        self.network_manager = network_manager
        self.local_db = local_data_file("local_schedule.db")
        self.init_local_db()
        self._cache = {}
        self._employees = []  # Кэш сотрудников (только активные)
        self._all_employees = []  # Кэш всех сотрудников (активные + удаленные)
        self._periods_cache = None
    
    @property
    def periods(self):
        """Свойство для периодов"""
        if self._periods_cache is None:
            self._periods_cache = self.get_periods()
        return self._periods_cache
    
    @property
    def employees(self):
        """Свойство для совместимости со старым кодом - возвращает только активных сотрудников"""
        if not self._employees:
            self._employees = self.load_employees()
        return self._employees
    
    @property
    def all_employees(self):
        """Свойство для получения всех сотрудников (активных + удаленных)"""
        if not self._all_employees:
            self._all_employees = self.load_all_employees()
        return self._all_employees
    
    def init_local_db(self):
        """Инициализация локальной SQLite базы данных."""
        with sqlite3.connect(self.local_db, timeout=10) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA busy_timeout=5000")
            cursor = conn.cursor()

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
            self._ensure_column(cursor, "employees", "is_deleted", "INTEGER DEFAULT 0")
            self._ensure_column(cursor, "schedules", "note_type", "TEXT DEFAULT ''")
            conn.commit()
            migrate_period_names(conn)

    @staticmethod
    def _ensure_column(cursor, table, column, definition):
        cursor.execute(f"PRAGMA table_info({table})")
        columns = {row[1] for row in cursor.fetchall()}
        if column not in columns:
            cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")

    def load_employees(self, include_deleted=False):
        """Загрузка сотрудников с сервера или из локальной БД"""
        # Пытаемся загрузить с сервера
        response = self.network_manager.send_request({'action': 'get_employees'})
        
        if response.get('status') == 'success':
            employees = response.get('data') or []
            # Сохраняем в локальную БД (даже если запись не удалась,
            # интерфейсу отдаём актуальные данные сервера).
            self._save_employees_to_local(employees)
            self._update_employee_caches()
            if include_deleted:
                return employees
            return [
                {"name": emp.get("name"), "position": emp.get("position") or ""}
                for emp in employees
                if not emp.get("is_deleted")
            ]
        # Загрузка из локальной БД
        self._update_employee_caches()
        return self._get_employees_from_cache(include_deleted)
    
    def _save_employees_to_local(self, employees):
        """UPSERT сотрудников: не теряет position и is_deleted."""
        rows = [
            (
                (emp.get("name") or "").strip(),
                emp.get("position") or "",
                1 if emp.get("is_deleted") else 0,
            )
            for emp in employees
            if (emp.get("name") or "").strip()
        ]
        if not rows:
            return
        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                conn.executemany(
                    """
                    INSERT INTO employees (name, position, is_deleted)
                    VALUES (?, ?, ?)
                    ON CONFLICT (name) DO UPDATE SET
                        position = excluded.position,
                        is_deleted = excluded.is_deleted
                    """,
                    rows,
                )
                conn.commit()
        except sqlite3.Error:
            logger.exception("Не удалось сохранить сотрудников")

    def _update_employee_caches(self):
        """Обновление кэшей сотрудников одним проходом по БД."""
        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    SELECT name, COALESCE(position, ''), COALESCE(is_deleted, 0)
                    FROM employees ORDER BY name COLLATE NOCASE
                    """
                )
                all_employees = [
                    {"name": row[0], "position": row[1], "is_deleted": row[2]}
                    for row in cursor.fetchall()
                ]
            self._all_employees = all_employees
            self._employees = [
                {"name": emp["name"], "position": emp["position"]}
                for emp in all_employees
                if emp["is_deleted"] == 0
            ]
        except sqlite3.Error:
            logger.exception("Не удалось обновить кэш сотрудников")

    def _get_employees_from_cache(self, include_deleted=False):
        """Получение сотрудников из кэша"""
        if include_deleted:
            return self._all_employees
        else:
            return self._employees
    
    def load_all_employees(self):
        """Загрузка всех сотрудников (включая удаленных)"""
        return self.load_employees(include_deleted=True)
    
    def get_periods(self):
        """Периоды с сервера (или из локальной БД при отсутствии связи)."""
        response = self.network_manager.send_request({"action": "get_periods"})
        if response.get("status") == "success":
            periods = sort_periods(
                [normalize_period_name(period) or period for period in response.get("data", [])]
            )
            self._save_periods_to_local(periods)
        else:
            periods = self._load_periods_from_local()
        self._periods_cache = periods
        return periods

    def _save_periods_to_local(self, periods):
        rows = [(period,) for period in periods if period]
        if not rows:
            return
        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                conn.executemany(
                    "INSERT OR IGNORE INTO periods (period) VALUES (?)", rows
                )
                conn.commit()
        except sqlite3.Error:
            logger.exception("Не удалось сохранить периоды")

    def _load_periods_from_local(self):
        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                cursor = conn.execute("SELECT period FROM periods")
                periods = [row[0] for row in cursor.fetchall()]
            return sort_periods(periods)
        except sqlite3.Error:
            logger.exception("Не удалось загрузить периоды")
            return []

    def load_schedule(self, period):
        """Расписание периода с сервера, при ошибке — из локальной БД."""
        period = normalize_period_name(period) or period
        if period in self._cache:
            return self._cache[period]

        response = self.network_manager.send_request(
            {"action": "get_schedule", "period": period}
        )
        if response.get("status") == "success":
            data = response.get("data") or {}
            data.setdefault("employees", [])
            data.setdefault("schedule", {})
            data.setdefault("notes", {})
            data = self._normalize_schedule_data(period, data)
            self._save_schedule_to_local(period, data)
        else:
            data = self._load_schedule_from_local(period)
            if data is None:
                return None
            data = self._normalize_schedule_data(period, data)
        self._cache[period] = data
        return data

    @staticmethod
    def _normalize_schedule_data(period, data):
        """Приводит расписание к длине месяца и убирает лишние примечания."""
        days = days_in_month(period)
        normalized = dict(data or {})
        schedule = {}
        for name, values in (normalized.get("schedule") or {}).items():
            values = list(values or [])[:days]
            values.extend([STATUS_EMPTY] * (days - len(values)))
            schedule[name] = values
        normalized["schedule"] = schedule

        # Отработанные часы — производное значение: пересчитываем их
        # от начала смены (07:30), чтобы старые записи не расходились
        # с новым стандартным рабочим днём.
        notes = {}
        for name, emp_notes in (normalized.get("notes") or {}).items():
            employee_schedule = schedule.get(name, [])
            employee_notes = {}
            for day, note in (emp_notes or {}).items():
                if not str(day).isdigit() or not isinstance(note, dict):
                    continue
                day_index = int(day)
                if not 0 <= day_index < days:
                    continue
                status = (
                    employee_schedule[day_index]
                    if day_index < len(employee_schedule)
                    else STATUS_EMPTY
                )
                note = dict(note)
                note["worked_hours"] = compute_worked_hours(
                    status, note.get("end_time")
                )
                employee_notes[str(day_index)] = note
            notes[name] = employee_notes
        normalized["notes"] = notes
        return normalized

    def _save_schedule_to_local(self, period, data):
        """Пакетное сохранение расписания (executemany вместо запроса на ячейку)."""
        period = normalize_period_name(period) or period
        days = days_in_month(period)
        notes_by_employee = data.get("notes") or {}
        rows = []
        for emp_name, schedule in (data.get("schedule") or {}).items():
            notes = notes_by_employee.get(emp_name) or {}
            for day_index, status in enumerate(list(schedule or [])[:days]):
                note = notes.get(str(day_index)) or {}
                rows.append(
                    (
                        emp_name,
                        day_index,
                        status,
                        note.get("end_time"),
                        note.get("worked_hours"),
                        note.get("note_type", ""),
                    )
                )

        order_rows = [
            (period, emp.get("name"), index)
            for index, emp in enumerate(data.get("employees") or [])
            if emp.get("name")
        ]

        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM periods WHERE period = ?", (period,))
                row = cursor.fetchone()
                if row:
                    period_id = row[0]
                else:
                    cursor.execute("INSERT INTO periods (period) VALUES (?)", (period,))
                    period_id = cursor.lastrowid

                if rows:
                    cursor.executemany(
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
                        [
                            (period_id, emp, day, status, end, hours, note_type)
                            for emp, day, status, end, hours, note_type in rows
                        ],
                    )
                if order_rows:
                    cursor.executemany(
                        """
                        INSERT INTO employee_order (period, employee_name, sort_order)
                        VALUES (?, ?, ?)
                        ON CONFLICT (period, employee_name) DO UPDATE SET
                            sort_order = excluded.sort_order
                        """,
                        order_rows,
                    )
                conn.commit()
        except sqlite3.Error:
            logger.exception("Не удалось сохранить расписание %s", period)
            return False
        return True

    def _load_schedule_from_local(self, period):
        period = normalize_period_name(period) or period
        days = days_in_month(period)
        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM periods WHERE period = ?", (period,))
                row = cursor.fetchone()
                if not row:
                    return None
                period_id = row[0]

                cursor.execute(
                    """
                    SELECT eo.employee_name, COALESCE(e.position, '')
                    FROM employee_order AS eo
                    LEFT JOIN employees AS e
                        ON eo.employee_name = e.name
                    WHERE eo.period = ? AND COALESCE(e.is_deleted, 0) = 0
                    ORDER BY eo.sort_order
                    """,
                    (period,),
                )
                employees = [
                    {"name": name, "position": position}
                    for name, position in cursor.fetchall()
                ]
                known_names = {employee["name"] for employee in employees}
                cursor.execute(
                    """
                    SELECT DISTINCT s.employee_name, COALESCE(e.position, '')
                    FROM schedules AS s
                    LEFT JOIN employees AS e ON s.employee_name = e.name
                    WHERE s.period_id = ? AND COALESCE(e.is_deleted, 0) = 0
                    ORDER BY s.employee_name COLLATE NOCASE
                    """,
                    (period_id,),
                )
                for name, position in cursor.fetchall():
                    if name not in known_names:
                        employees.append({"name": name, "position": position})
                        known_names.add(name)

                if not employees:
                    return None

                schedule = {
                    employee["name"]: [STATUS_EMPTY] * days for employee in employees
                }
                notes = {}
                cursor.execute(
                    """
                    SELECT employee_name, day_index, status, end_time,
                           worked_hours, note_type
                    FROM schedules WHERE period_id = ?
                    """,
                    (period_id,),
                )
                active = {employee["name"] for employee in employees}
                for emp_name, day_index, status, end_time, worked_hours, note_type in cursor.fetchall():
                    if emp_name not in active or not 0 <= day_index < days:
                        continue
                    try:
                        status = int(status)
                    except (TypeError, ValueError):
                        status = STATUS_EMPTY
                    if status not in (0, 1, 2, 3, 4):
                        status = STATUS_EMPTY
                    schedule[emp_name][day_index] = status
                    if end_time or worked_hours or note_type:
                        notes.setdefault(emp_name, {})[str(day_index)] = {
                            "end_time": end_time,
                            "worked_hours": worked_hours,
                            "note_type": note_type or "",
                        }
                return {"employees": employees, "schedule": schedule, "notes": notes}
        except sqlite3.Error:
            logger.exception("Не удалось загрузить расписание %s", period)
            return None

    def save_schedule(self, period, data):
        """Локальное сохранение + отправка на сервер (с офлайн-очередью)."""
        period = normalize_period_name(period) or period
        data = self._normalize_schedule_data(period, data)
        saved = self._save_schedule_to_local(period, data)
        self._cache[period] = data
        response = self.network_manager.send_request(
            {"action": "save_schedule", "period": period, "data": data}
        )
        if response.get("status") == "error":
            if saved:
                logger.warning(
                    "Расписание сохранено локально, но не отправлено: %s",
                    response.get("message"),
                )
            else:
                logger.error(
                    "Не удалось сохранить расписание %s: %s",
                    period,
                    response.get("message"),
                )
        return saved

    def update_employee_order(self, period, employees):
        period = normalize_period_name(period) or period
        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM employee_order WHERE period = ?", (period,))
                cursor.executemany(
                    "INSERT INTO employee_order (period, employee_name, sort_order) "
                    "VALUES (?, ?, ?)",
                    [(period, name, index) for index, name in enumerate(employees)],
                )
                conn.commit()
        except sqlite3.Error:
            logger.exception("Не удалось обновить порядок сотрудников")
            return False

        if period in self._cache:
            current = self._cache[period]
            by_name = {emp["name"]: emp for emp in current.get("employees", [])}
            current["employees"] = [by_name[name] for name in employees if name in by_name]
            self._cache[period] = current

        self.network_manager.send_request(
            {"action": "set_employee_order", "period": period, "employees": employees}
        )
        return True

    def add_employee(self, name):
        name = (name or "").strip()
        if not name:
            return False
        for emp in self._load_employees_from_local(include_deleted=True):
            if emp["name"].casefold() == name.casefold():
                if emp.get("is_deleted", 0) == 1:
                    return self.restore_employee(emp["name"])
                return False

        response = self.network_manager.send_request(
            {"action": "add_employee", "name": name}
        )
        if response.get("status") == "error":
            logger.warning("Сервер не добавил сотрудника: %s", response.get("message"))
            self.load_employees()
            return False

        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                conn.execute(
                    "INSERT INTO employees (name, is_deleted) VALUES (?, 0)", (name,)
                )
                conn.commit()
            self._update_employee_caches()
        except sqlite3.IntegrityError:
            return False
        return True

    def update_employee(self, old_name, new_name, new_position):
        old_name = (old_name or "").strip()
        new_name = (new_name or "").strip()
        if not old_name or not new_name:
            return False, "Пустое имя сотрудника"
        if new_name.casefold() != old_name.casefold():
            for emp in self._load_employees_from_local(include_deleted=True):
                if emp["name"].casefold() == new_name.casefold():
                    return False, "Сотрудник с таким именем уже существует"

        response = self.network_manager.send_request(
            {
                "action": "update_employee",
                "old_name": old_name,
                "new_name": new_name,
                "position": new_position or "",
            }
        )
        if response.get("status") == "error":
            return False, response.get("message", "Ошибка обновления")

        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "UPDATE employees SET name = ?, position = ? WHERE name = ?",
                    (new_name, new_position or "", old_name),
                )
                cursor.execute(
                    "UPDATE schedules SET employee_name = ? WHERE employee_name = ?",
                    (new_name, old_name),
                )
                cursor.execute(
                    "UPDATE employee_order SET employee_name = ? WHERE employee_name = ?",
                    (new_name, old_name),
                )
                conn.commit()
            self._update_employee_caches()
            self._cache.clear()
            return True, "Данные сотрудника обновлены"
        except sqlite3.Error as exc:
            logger.exception("Ошибка локального обновления сотрудника")
            return False, f"Ошибка обновления: {exc}"

    def _set_employee_deleted(self, name, deleted):
        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                conn.execute(
                    "UPDATE employees SET is_deleted = ? WHERE name = ?",
                    (1 if deleted else 0, name),
                )
                conn.commit()
            self._update_employee_caches()
        except sqlite3.Error as exc:
            logger.exception("Ошибка изменения статуса сотрудника")
            return False, f"Ошибка: {exc}"

        action = "delete_employee" if deleted else "restore_employee"
        self.network_manager.send_request({"action": action, "name": name})
        return True, ("Сотрудник помечен как удаленный" if deleted else "Сотрудник восстановлен")

    def delete_employee(self, name):
        return self._set_employee_deleted(name, True)

    def restore_employee(self, name):
        return self._set_employee_deleted(name, False)

    def create_period(self, period, employees, copy_from=None):
        """Создаёт период в каноническом формате «Месяц Год»."""
        period = normalize_period_name(period) or period
        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                if conn.execute(
                    "SELECT 1 FROM periods WHERE period = ?", (period,)
                ).fetchone():
                    logger.warning("Период %s уже существует", period)
                    return False
        except sqlite3.Error:
            logger.exception("Не удалось проверить период %s", period)
            return False

        days = days_in_month(period)
        employees = list(dict.fromkeys(name for name in employees if name))

        schedule = {name: [STATUS_EMPTY] * days for name in employees}
        notes = {}
        if copy_from:
            source = self.load_schedule(copy_from)
            if source:
                source_names = [emp["name"] for emp in source["employees"]]
                ordered = [name for name in source_names if name in employees]
                ordered += [name for name in employees if name not in ordered]
                employees = ordered
                for name in employees:
                    values = list(source["schedule"].get(name, []))
                    values = (values + [STATUS_EMPTY] * days)[:days]
                    schedule[name] = values
                    source_notes = source.get("notes", {}).get(name)
                    if source_notes:
                        notes[name] = {
                            key: value
                            for key, value in source_notes.items()
                            if str(key).isdigit() and int(key) < days
                        }

        schedule_data = {
            "employees": [{"name": name} for name in employees],
            "schedule": schedule,
            "notes": notes,
        }
        saved = self._save_schedule_to_local(period, schedule_data)
        self._cache[period] = schedule_data
        if not saved:
            return False
        if self._periods_cache is None:
            self._periods_cache = []
        if period not in self._periods_cache:
            self._periods_cache.append(period)
        self._periods_cache = sort_periods(self._periods_cache)

        response = self.network_manager.send_request(
            {"action": "create_period", "period": period, "employees": employees}
        )
        return response.get("status") != "error"

    def delete_period(self, period):
        period = normalize_period_name(period) or period
        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM periods WHERE period = ?", (period,))
                row = cursor.fetchone()
                if not row:
                    return False
                period_id = row[0]
                cursor.execute("DELETE FROM schedules WHERE period_id = ?", (period_id,))
                cursor.execute("DELETE FROM employee_order WHERE period = ?", (period,))
                cursor.execute("DELETE FROM periods WHERE id = ?", (period_id,))
                conn.commit()
        except sqlite3.Error:
            logger.exception("Не удалось удалить период %s", period)
            return False

        if self._periods_cache is not None and period in self._periods_cache:
            self._periods_cache.remove(period)
        self._cache.pop(period, None)
        self.network_manager.send_request({"action": "delete_period", "period": period})
        return True

    def add_employees_to_period(self, period, employees):
        period = normalize_period_name(period) or period
        current_data = self.load_schedule(period)
        if not current_data:
            return False

        existing = {emp["name"] for emp in current_data["employees"]}
        days = days_in_month(period)
        for name in employees:
            if name not in existing:
                current_data["employees"].append({"name": name, "position": ""})
                current_data["schedule"][name] = [STATUS_EMPTY] * days
                existing.add(name)

        saved = self._save_schedule_to_local(period, current_data)
        self._cache[period] = current_data
        if not saved:
            return False
        response = self.network_manager.send_request(
            {"action": "add_employees_to_period", "period": period, "employees": employees}
        )
        return response.get("status") != "error"

    def sync_with_server(self):
        if not self.network_manager.connected:
            return False
        sync_data = {
            "employees": self._load_employees_from_local(include_deleted=True),
            "periods": self._load_periods_from_local(),
            "schedules": [],
        }
        for period in sync_data["periods"]:
            schedule_data = self.load_schedule(period)
            if schedule_data:
                sync_data["schedules"].append(
                    {
                        "period": period,
                        "schedule": schedule_data["schedule"],
                        "notes": schedule_data.get("notes", {}),
                    }
                )
        response = self.network_manager.send_request(
            {"action": "sync_data", "data": sync_data}
        )
        return response.get("status") == "success"

    def _load_employees_from_local(self, include_deleted=False):
        try:
            with sqlite3.connect(self.local_db, timeout=10) as conn:
                cursor = conn.cursor()
                if include_deleted:
                    cursor.execute(
                        "SELECT name, COALESCE(position, ''), COALESCE(is_deleted, 0) "
                        "FROM employees ORDER BY name COLLATE NOCASE"
                    )
                    return [
                        {"name": row[0], "position": row[1], "is_deleted": row[2]}
                        for row in cursor.fetchall()
                    ]
                cursor.execute(
                    "SELECT name, COALESCE(position, '') FROM employees "
                    "WHERE COALESCE(is_deleted, 0) = 0 ORDER BY name COLLATE NOCASE"
                )
                return [
                    {"name": row[0], "position": row[1]} for row in cursor.fetchall()
                ]
        except sqlite3.Error:
            logger.exception("Не удалось загрузить сотрудников")
            return []


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
        self.status_label.setObjectName("statusPill")
        set_pill_state(self.status_label, "offline")
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
            set_pill_state(self.status_label, "error")
            return
        
        try:
            with socket.create_connection((host, port), timeout=5):
                pass
            self.status_label.setText("Подключено")
            set_pill_state(self.status_label, "connected")
        except OSError:
            self.status_label.setText("Ошибка подключения")
            set_pill_state(self.status_label, "error")
    
    def get_connection_params(self):
        return self.host_edit.text(), int(self.port_edit.text())

    def accept(self):
        """Не закрываем диалог с некорректным адресом/портом."""
        host = self.host_edit.text().strip()
        try:
            port = int(self.port_edit.text())
        except ValueError:
            port = 0
        if not host or not 0 < port < 65536:
            self.status_label.setText("Неверный адрес или порт")
            set_pill_state(self.status_label, "error")
            return
        super().accept()


class MonthSelectionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent) 
        self.setWindowTitle("Выбор месяца и года")
        self.setModal(True)
        self.setMinimumWidth(350)
        
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
        for month_num, month_name in MONTH_NAMES.items():
            self.month_combo.addItem(month_name, month_num)
        self.month_combo.setCurrentIndex(current_date.month() - 1)
        grid_layout.addWidget(self.month_combo, 1, 1)
        
        # Поле для ввода названия месяца (для вариантов)
        grid_layout.addWidget(QLabel("Название:"), 2, 0)
        self.name_edit = QLineEdit()
        grid_layout.addWidget(self.name_edit, 2, 1)
        
        layout.addLayout(grid_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def get_selected_period(self):
        year = self.year_combo.currentData()
        month = self.month_combo.currentData()
        return f"{year}-{month:02d}"
    
    def get_period_name(self):
        """Получение полного названия периода (с вариантом)"""
        year = self.year_combo.currentData()
        month_name = self.month_combo.currentText()
        custom_name = self.name_edit.text().strip()
        
        base_name = f"{month_name} {year}"
        if custom_name:
            return f"{base_name} ({custom_name})"
        return base_name


class MonthExistsDialog(QDialog):
    def __init__(self, period_name, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Месяц уже существует")
        self.setModal(True)
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Сообщение
        message_label = QLabel(
            f"Месяц '{period_name}' уже существует.\n\n"
            "Хотите создать дополнительный вариант?"
        )
        message_label.setWordWrap(True)
        layout.addWidget(message_label)
        
        # Поле для названия варианта
        form_layout = QFormLayout()
        self.variant_name_edit = QLineEdit()
        self.variant_name_edit.setPlaceholderText("Например: Вариант 2, Резервная копия и т.д.")
        form_layout.addRow("Название варианта:", self.variant_name_edit)
        layout.addLayout(form_layout)
        
        # Кнопки
        button_box = QDialogButtonBox(
            QDialogButtonBox.Yes | QDialogButtonBox.No | QDialogButtonBox.Cancel
        )
        button_box.button(QDialogButtonBox.Yes).setText("Создать вариант")
        button_box.button(QDialogButtonBox.No).setText("Перейти к существующему")
        button_box.button(QDialogButtonBox.Cancel).setText("Отмена")

        self.action = "cancel"
        button_box.button(QDialogButtonBox.Yes).clicked.connect(
            lambda: self.accept_with_action("variant")
        )
        button_box.button(QDialogButtonBox.No).clicked.connect(
            lambda: self.accept_with_action("existing")
        )
        button_box.button(QDialogButtonBox.Cancel).clicked.connect(self.reject)

        layout.addWidget(button_box)
    
    def accept_with_action(self, action):
        self.action = action
        super().accept()
    
    def get_variant_name(self):
        return self.variant_name_edit.text().strip()
    
    def get_action(self):
        return getattr(self, 'action', 'cancel')


class NoteDialog(QDialog):
    def __init__(self, parent=None, current_end_time=DEFAULT_END_TIME, current_note_type=""):
        super().__init__(parent)
        self.setWindowTitle("Примечание к рабочему дню")
        self.setModal(True)
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        form_layout = QFormLayout()
        
        # Время окончания
        self.end_time_edit = QTimeEdit()
        try:
            if current_end_time:
                hours, minutes = map(int, current_end_time.split(':'))
                self.end_time_edit.setTime(QTime(hours, minutes))
            else:
                self.end_time_edit.setTime(QTime(20, 0))
        except (TypeError, ValueError):
            self.end_time_edit.setTime(QTime(20, 0))
        self.end_time_edit.setDisplayFormat("HH:mm")
        form_layout.addRow("Время окончания:", self.end_time_edit)
        
        # Тип примечания
        self.note_type_group = QButtonGroup(self)
        self.note_type_group.setExclusive(True)
        
        note_type_layout = QHBoxLayout()
        
        self.no_note_radio = QRadioButton("Без примечания")
        self.no_note_radio.setChecked(True)
        self.note_type_group.addButton(self.no_note_radio, 0)
        note_type_layout.addWidget(self.no_note_radio)
        
        self.responsible_radio = QRadioButton("Ответственный")
        self.note_type_group.addButton(self.responsible_radio, 1)
        note_type_layout.addWidget(self.responsible_radio)
        
        self.dms_radio = QRadioButton("ДМС")
        self.note_type_group.addButton(self.dms_radio, 2)
        note_type_layout.addWidget(self.dms_radio)
        
        # Кнопка для "Нежелание работать"
        self.no_desire_radio = QRadioButton("Нежелание работать")
        self.note_type_group.addButton(self.no_desire_radio, 3)
        note_type_layout.addWidget(self.no_desire_radio)
        
        # Кнопка для "Архив"
        self.archive_radio = QRadioButton("Архив")
        self.note_type_group.addButton(self.archive_radio, 4)
        note_type_layout.addWidget(self.archive_radio)

        form_layout.addRow("Тип примечания:", note_type_layout)
        
        # Устанавливаем текущее значение
        if current_note_type == 'о':
            self.responsible_radio.setChecked(True)
        elif current_note_type == 'д':
            self.dms_radio.setChecked(True)
        elif current_note_type == '*':
            self.no_desire_radio.setChecked(True)
        elif current_note_type == 'а':
            self.archive_radio.setChecked(True)
        else:
            self.no_note_radio.setChecked(True)
        
        layout.addLayout(form_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def get_end_time(self):
        return self.end_time_edit.time().toString("HH:mm")
    
    def get_note_type(self):
        if self.responsible_radio.isChecked():
            return 'о'
        elif self.dms_radio.isChecked():
            return 'д'
        elif self.no_desire_radio.isChecked():
            return '*'
        elif self.archive_radio.isChecked():
            return 'а'
        else:
            return ''


class ScheduleItemDelegate(QStyledItemDelegate):
    """Сохраняет цвет статуса и рисует синюю рамку выбранной ячейки."""

    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        opt.state &= ~QStyle.State_Selected
        opt.state &= ~QStyle.State_HasFocus
        super().paint(painter, opt, index)

        if option.state & QStyle.State_Selected:
            painter.save()
            painter.setPen(QPen(QColor("#2563EB"), 2))
            painter.setBrush(Qt.NoBrush)
            painter.drawRect(option.rect.adjusted(1, 1, -1, -1))
            painter.restore()


class NoteTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_widget = parent
        self.is_loading = False
        self._pending_updates = False
        self.setItemDelegate(ScheduleItemDelegate(self))

        # Debounce для пересчёта итогов: при массовой загрузке таймер не нужен.
        self._update_timer = QTimer(self)
        self._update_timer.setSingleShot(True)
        self._update_timer.timeout.connect(self._delayed_update_counters)

    def setItem(self, row, column, item):
        super().setItem(row, column, item)
        if self.is_loading or column >= self.columnCount() - 4:
            return
        self._pending_updates = True
        self._update_timer.start(300)
        app = getattr(self.parent_widget, "parent_app", None)
        if app is not None:
            app.set_unsaved_changes(True)

    def _delayed_update_counters(self):
        if self._pending_updates and not self.is_loading:
            self.parent_widget.update_counters()
            self._pending_updates = False


class NoteItem(QTableWidgetItem):
    def __init__(self, text="", has_note=False, note_data=None, status=STATUS_EMPTY):
        super().__init__(text)
        self.has_note = has_note
        self.note_data = note_data or {}
        self.status = status
    
    def clone(self):
        return NoteItem(
            self.text(), self.has_note, self.note_data.copy(), getattr(self, "status", STATUS_EMPTY)
        )


class EmployeeManagerDialog(QDialog):
    """Диалог управления сотрудниками"""
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Управление сотрудниками")
        self.setModal(True)
        self.setMinimumSize(600, 500)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Создаем вкладки
        self.tab_widget = QTabWidget()
        
        # Вкладка активных сотрудников
        self.active_tab = QWidget()
        self.active_layout = QVBoxLayout(self.active_tab)
        self.setup_active_tab()
        
        # Вкладка удаленных сотрудников
        self.deleted_tab = QWidget()
        self.deleted_layout = QVBoxLayout(self.deleted_tab)
        self.setup_deleted_tab()
        
        self.tab_widget.addTab(self.active_tab, "Активные сотрудники")
        self.tab_widget.addTab(self.deleted_tab, "Удаленные сотрудники")
        
        layout.addWidget(self.tab_widget)
        
        # Кнопки внизу
        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        # Загружаем данные
        self.load_employees()
    
    def setup_active_tab(self):
        """Настройка вкладки активных сотрудников"""
        self.active_layout.setContentsMargins(12, 12, 12, 12)
        self.active_layout.setSpacing(10)

        # Панель кнопок
        button_layout = QHBoxLayout()
        button_layout.setSpacing(8)

        self.add_btn = QPushButton("Добавить")
        self.add_btn.setFixedHeight(34)
        self.add_btn.setMinimumWidth(128)
        set_variant(self.add_btn, "primary")
        self.add_btn.clicked.connect(self.add_employee)
        button_layout.addWidget(self.add_btn)

        self.edit_btn = QPushButton("Редактировать")
        self.edit_btn.setFixedHeight(34)
        self.edit_btn.setMinimumWidth(140)
        self.edit_btn.clicked.connect(self.edit_employee)
        button_layout.addWidget(self.edit_btn)

        self.delete_btn = QPushButton("Удалить")
        self.delete_btn.setFixedHeight(34)
        self.delete_btn.setMinimumWidth(116)
        set_variant(self.delete_btn, "danger")
        self.delete_btn.clicked.connect(self.delete_employee)
        button_layout.addWidget(self.delete_btn)

        button_layout.addStretch()
        self.active_layout.addLayout(button_layout)

        # Список сотрудников
        self.active_list = QListWidget()
        self.active_list.itemDoubleClicked.connect(self.edit_employee)
        self.active_layout.addWidget(self.active_list)

    def setup_deleted_tab(self):
        """Настройка вкладки удаленных сотрудников"""
        self.deleted_layout.setContentsMargins(12, 12, 12, 12)
        self.deleted_layout.setSpacing(10)

        # Панель кнопок
        button_layout = QHBoxLayout()
        button_layout.setSpacing(8)

        self.restore_btn = QPushButton("Восстановить")
        self.restore_btn.setFixedHeight(34)
        self.restore_btn.setMinimumWidth(140)
        self.restore_btn.clicked.connect(self.restore_employee)
        button_layout.addWidget(self.restore_btn)

        button_layout.addStretch()
        self.deleted_layout.addLayout(button_layout)

        # Список удаленных сотрудников
        self.deleted_list = QListWidget()
        self.deleted_layout.addWidget(self.deleted_list)
    
    def load_employees(self):
        """Загрузка сотрудников (один запрос к серверу/БД)."""
        self.active_list.clear()
        self.deleted_list.clear()
        all_employees = self.db.load_employees(include_deleted=True)
        for emp in all_employees:
            item_text = emp['name']
            if emp.get('position'):
                item_text += f" ({emp['position']})"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, emp)
            if emp.get('is_deleted', 0) == 1:
                item.setForeground(Qt.gray)
                self.deleted_list.addItem(item)
            else:
                self.active_list.addItem(item)
    
    def add_employee(self):
        """Добавление нового сотрудника"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Добавить сотрудника")
        dialog.setModal(True)
        
        layout = QVBoxLayout(dialog)
        
        form_layout = QFormLayout()
        
        name_edit = QLineEdit()
        form_layout.addRow("ФИО сотрудника:", name_edit)
        
        layout.addLayout(form_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        if dialog.exec_() == QDialog.Accepted:
            name = name_edit.text().strip()
            
            if not name:
                QMessageBox.warning(self, "Ошибка", "ФИО сотрудника не может быть пустым")
                return
            
            if self.db.add_employee(name):
                QMessageBox.information(self, "Успех", "Сотрудник добавлен")
                self.load_employees()
                # Обновляем данные в главном окне
                if self.parent():
                    self.parent().load_initial_data()
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось добавить сотрудника")
    
    def edit_employee(self):
        """Редактирование сотрудника"""
        current_item = self.active_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Ошибка", "Выберите сотрудника для редактирования")
            return
        
        emp_data = current_item.data(Qt.UserRole)
        old_name = emp_data['name']
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Редактировать сотрудника")
        dialog.setModal(True)
        
        layout = QVBoxLayout(dialog)
        
        form_layout = QFormLayout()
        
        name_edit = QLineEdit(old_name)
        form_layout.addRow("ФИО сотрудника:", name_edit)
        
        layout.addLayout(form_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        if dialog.exec_() == QDialog.Accepted:
            new_name = name_edit.text().strip()
            
            if not new_name:
                QMessageBox.warning(self, "Ошибка", "ФИО сотрудника не может быть пустым")
                return
            
            if new_name == old_name:
                return  # Ничего не изменилось
            
            # Вызываем метод update_employee, сохраняя существующую должность
            success, message = self.db.update_employee(
                old_name, new_name, emp_data.get("position", "")
            )
            if success:
                # Не показываем сообщение об успехе
                self.load_employees()
                # Обновляем данные в главном окне
                if self.parent():
                    self.parent().load_initial_data()
            else:
                QMessageBox.warning(self, "Ошибка", message)
                    
    def delete_employee(self):
        """Удаление сотрудника"""
        current_item = self.active_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Ошибка", "Выберите сотрудника для удаления")
            return
        
        emp_data = current_item.data(Qt.UserRole)
        name = emp_data['name']
        
        reply = QMessageBox.question(self, "Подтверждение", 
                                   f"Вы уверены, что хотите удалить сотрудника '{name}'?\n\n"
                                   "Примечание: Сотрудник не будет удален из существующих таблиц расписания, "
                                   "но будет скрыт из списка активных сотрудников.",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            success, message = self.db.delete_employee(name)
            if success:
                QMessageBox.information(self, "Успех", message)
                self.load_employees()
                # Обновляем данные в главном окне
                if self.parent():
                    self.parent().load_initial_data()
            else:
                QMessageBox.warning(self, "Ошибка", message)
    
    def restore_employee(self):
        """Восстановление удаленного сотрудника"""
        current_item = self.deleted_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "Ошибка", "Выберите сотрудника для восстановления")
            return
        
        emp_data = current_item.data(Qt.UserRole)
        name = emp_data['name']
        
        success, message = self.db.restore_employee(name)
        if success:
            QMessageBox.information(self, "Успех", message)
            self.load_employees()
            # Обновляем данные в главном окне
            if self.parent():
                self.parent().load_initial_data()
        else:
            QMessageBox.warning(self, "Ошибка", message)


class NewPeriodDialog(QDialog):
    """Диалог создания нового периода с опцией копирования сотрудников"""
    def __init__(self, parent=None, db=None, selected_period=None):
        super().__init__(parent)
        self.db = db
        self.selected_period = selected_period
        self.setWindowTitle("Создание нового месяца")
        self.setModal(True)
        self.setMinimumWidth(450)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Информация о выбранном периоде
        parsed = parse_period(selected_period)
        if parsed:
            period_info = QLabel(f"Создается месяц: {MONTH_NAMES[parsed[1]]} {parsed[0]}")
            period_info.setStyleSheet(
                "font-weight: bold; font-size: 12px; padding: 5px;"
            )
            layout.addWidget(period_info)
        
        # Опция копирования сотрудников
        self.copy_checkbox = QCheckBox("Скопировать сотрудников из...")
        self.copy_checkbox.setChecked(True)
        self.copy_checkbox.stateChanged.connect(self.on_copy_changed)
        layout.addWidget(self.copy_checkbox)
        
        # Выбор источника для копирования
        copy_options_layout = QHBoxLayout()
        copy_options_layout.setContentsMargins(20, 0, 0, 0)  # Отступ для вложенности
        
        self.prev_month_radio = QRadioButton("предыдущего месяца")
        self.prev_month_radio.setChecked(True)
        copy_options_layout.addWidget(self.prev_month_radio)
        
        self.specific_month_radio = QRadioButton("определённого месяца:")
        copy_options_layout.addWidget(self.specific_month_radio)
        
        self.source_combo = QComboBox()
        self.source_combo.setEnabled(False)
        
        # Заполняем список периодов
        periods = self.db.periods if self.db else []
        for period in periods:
            self.source_combo.addItem(period, period)
        if not periods:
            # Копировать неоткуда — показываем выбор сотрудников вручную.
            # Блокируем сигнал: employee_frame ещё не создан.
            with QSignalBlocker(self.copy_checkbox):
                self.copy_checkbox.setChecked(False)
            self.copy_checkbox.setEnabled(False)

        copy_options_layout.addWidget(self.source_combo)
        copy_options_layout.addStretch()
        
        layout.addLayout(copy_options_layout)
        
        # Связываем радио-кнопки
        self.prev_month_radio.toggled.connect(self.on_source_changed)
        self.specific_month_radio.toggled.connect(self.on_source_changed)
        
        layout.addSpacing(10)
        
        # Область выбора сотрудников (изначально отключена)
        self.employee_frame = QGroupBox("Выбор сотрудников")
        self.employee_frame.setEnabled(False)
        employee_layout = QVBoxLayout(self.employee_frame)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.content = QWidget()
        self.scroll_layout = QVBoxLayout(self.content)
        
        self.checkboxes = []
        employees = self.db.employees if self.db else []
        for emp in employees:
            cb = QCheckBox(emp["name"])
            cb.setChecked(True)
            self.checkboxes.append((emp["name"], cb))
            self.scroll_layout.addWidget(cb)
        
        scroll.setWidget(self.content)
        employee_layout.addWidget(scroll)
        
        btn_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("Выбрать все")
        self.select_all_btn.clicked.connect(lambda: self.set_all_checkboxes(True))
        btn_layout.addWidget(self.select_all_btn)
        
        self.deselect_all_btn = QPushButton("Снять все")
        self.deselect_all_btn.clicked.connect(lambda: self.set_all_checkboxes(False))
        btn_layout.addWidget(self.deselect_all_btn)
        
        employee_layout.addLayout(btn_layout)
        
        layout.addWidget(self.employee_frame)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        # Применяем начальное состояние
        self.on_copy_changed()
    
    def on_copy_changed(self):
        """Обработка изменения состояния чекбокса копирования"""
        if self.copy_checkbox.isChecked():
            self.employee_frame.setEnabled(False)
            self.prev_month_radio.setEnabled(True)
            self.specific_month_radio.setEnabled(True)
            self.source_combo.setEnabled(self.specific_month_radio.isChecked())
        else:
            self.employee_frame.setEnabled(True)
            self.prev_month_radio.setEnabled(False)
            self.specific_month_radio.setEnabled(False)
            self.source_combo.setEnabled(False)
    
    def on_source_changed(self):
        """Обработка изменения источника копирования"""
        self.source_combo.setEnabled(self.specific_month_radio.isChecked())
    
    def set_all_checkboxes(self, state):
        """Установка состояния всех чекбоксов"""
        for _, cb in self.checkboxes:
            cb.setChecked(state)
    
    def get_selected_employees(self):
        """Получение выбранных сотрудников"""
        if self.copy_checkbox.isChecked():
            # Если копирование включено, возвращаем None
            return None
        else:
            # Иначе возвращаем выбранных сотрудников
            return [name for name, cb in self.checkboxes if cb.isChecked()]
    
    def get_copy_source(self):
        """Получение источника для копирования"""
        if not self.copy_checkbox.isChecked():
            return None
        
        if self.prev_month_radio.isChecked():
            # Берём хронологически предыдущий период относительно создаваемого.
            periods = sort_periods(self.db.periods) if self.db else []
            if not periods:
                return None
            target_key = period_sort_key(self.selected_period)
            previous = [
                period for period in periods if period_sort_key(period) < target_key
            ]
            return previous[-1] if previous else periods[-1]
        else:
            # Копирование из выбранного месяца
            return self.source_combo.currentData()


class MonthWidget(QWidget):
    def __init__(self, parent=None, db=None):
        super().__init__(parent)
        self.parent_app = parent  # Сохраняем ссылку на главное приложение
        self.db = db
        self.current_period = None
        self.sunday_columns = []  # Список индексов колонок воскресений
        self.copied_data = None  # Для хранения скопированных данных
        self.initUI()
    
    @staticmethod
    def create_trash_icon():
        """Создает контурную минималистичную иконку корзины"""
        pixmap = QPixmap(18, 18)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Контурная кисть и перо (белый цвет для контраста на красном фоне)
        pen = QPen(QColor(185, 28, 28), 1.8)
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        
        # Минималистичная контурная корзина
        # Ручка (верхняя часть) - простая горизонтальная линия с вертикальными опорами
        painter.drawLine(5, 4, 5, 6)  # Левая опора
        painter.drawLine(13, 4, 13, 6)  # Правая опора
        painter.drawLine(5, 4, 13, 4)  # Верхняя перекладина
        
        # Корпус корзины - простой контур
        painter.drawLine(4, 6, 4, 13)  # Левая сторона
        painter.drawLine(14, 6, 14, 13)  # Правая сторона
        painter.drawLine(4, 6, 14, 6)  # Верх
        
        # Низ корзины - простая линия
        painter.drawLine(4, 13, 14, 13)
        
        painter.end()
        return QIcon(pixmap)
    
    def initUI(self):
        self.setObjectName("monthCard")
        self.setAttribute(Qt.WA_StyledBackground, True)

        layout = QVBoxLayout()
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(10)
        self.setLayout(layout)

        # Верхняя панель карточки месяца
        top_layout = QHBoxLayout()
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(8)

        month_label = QLabel("ПЕРИОД")
        month_label.setObjectName("sectionCaption")
        top_layout.addWidget(month_label)

        self.period_combo = QComboBox()
        self.period_combo.setMinimumWidth(210)
        self.period_combo.setFixedHeight(36)
        self.period_combo.currentIndexChanged.connect(self.period_changed)
        top_layout.addWidget(self.period_combo)
        top_layout.addStretch()

        # Кнопка для изменения порядка сотрудников
        self.order_btn = QPushButton("↑↓")
        self.order_btn.setFixedSize(36, 36)
        self.order_btn.setToolTip("Изменить порядок сотрудников")
        set_variant(self.order_btn, "icon")
        self.order_btn.clicked.connect(self.change_employee_order)
        top_layout.addWidget(self.order_btn)

        # Кнопка экспорта этого месяца в Excel
        self.export_month_btn = QPushButton("Экспорт")
        self.export_month_btn.setFixedHeight(36)
        self.export_month_btn.setMinimumWidth(92)
        self.export_month_btn.setToolTip("Экспорт месяца в Excel")
        self.export_month_btn.clicked.connect(self._export_this_month)
        top_layout.addWidget(self.export_month_btn)

        # Кнопка удаления этого месяца
        self.delete_month_btn = QPushButton()
        self.delete_month_btn.setIcon(self.create_trash_icon())
        self.delete_month_btn.setFixedSize(36, 36)
        self.delete_month_btn.setToolTip("Удалить месяц")
        set_variant(self.delete_month_btn, "dangerIcon")
        self.delete_month_btn.clicked.connect(self._delete_this_month)
        top_layout.addWidget(self.delete_month_btn)

        layout.addLayout(top_layout)

        # Таблица
        self.table = NoteTableWidget(self)
        self.setup_table(self.table)
        layout.addWidget(self.table)
    
    def setup_table(self, table):
        table.setObjectName("scheduleTable")
        table.setSelectionMode(QTableWidget.ExtendedSelection)
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(self.show_context_menu)

        table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        table.horizontalHeader().setMinimumSectionSize(28)
        table.verticalHeader().setMinimumSectionSize(28)
        table.horizontalHeader().setHighlightSections(False)
        table.verticalHeader().setHighlightSections(False)
        table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        table.verticalHeader().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        table.verticalHeader().setFixedWidth(172)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setWordWrap(False)
        table.setShowGrid(True)
        table.cellClicked.connect(
            lambda *_: setattr(self.parent_app, "last_active_widget", self)
        )

        table.verticalHeader().setDefaultSectionSize(30)

        font = table.horizontalHeader().font()
        font.setBold(True)
        font.setPointSizeF(9.0)
        table.horizontalHeader().setFont(font)

        font = table.verticalHeader().font()
        font.setBold(True)
        font.setPointSizeF(9.0)
        table.verticalHeader().setFont(font)
    
    def _export_this_month(self):
        """Экспорт таблицы этого месяца в Excel (вызов главного окна)."""
        if self.parent_app and hasattr(self.parent_app, 'export_month_to_excel'):
            self.parent_app.export_month_to_excel(self)
    
    def _delete_this_month(self):
        """Удаление этого месяца с подтверждением."""
        if not self.current_period:
            return
        
        if self.parent_app and hasattr(self.parent_app, 'delete_period'):
            self.parent_app.delete_period(self.current_period, self)
    
    def change_employee_order(self):
        """Изменение порядка сотрудников в таблице"""
        if not self.current_period:
            return
        if not self.parent_app.resolve_unsaved_changes():
            return
        
        # Получаем текущий список сотрудников
        current_employees = []
        for row in range(self.table.rowCount()):
            item = self.table.verticalHeaderItem(row)
            if item:
                current_employees.append(item.text())
        
        if not current_employees:
            return
        
        # Создаем диалог для изменения порядка
        dialog = QDialog(self)
        dialog.setWindowTitle("Изменение порядка сотрудников")
        dialog.setModal(True)
        dialog.setMinimumSize(400, 500)
        
        layout = QVBoxLayout(dialog)
        
        # Создаем список с возможностью перетаскивания
        list_widget = QListWidget()

        # Устанавливаем режим перетаскивания, который позволяет только перемещать между позициями
        list_widget.setDragDropMode(QListWidget.InternalMove)
        
        # Отключаем возможность выбора нескольких элементов
        list_widget.setSelectionMode(QListWidget.SingleSelection)
        
        for emp in current_employees:
            item = QListWidgetItem(emp)
            # Устанавливаем флаги только для перетаскивания, без возможности наложения
            item.setFlags(item.flags() | Qt.ItemIsDragEnabled)
            list_widget.addItem(item)
        
        layout.addWidget(list_widget)
        
        # Добавляем инструкцию
        instruction_label = QLabel("Перетащите сотрудника в нужную позицию между другими сотрудниками")
        instruction_label.setWordWrap(True)
        instruction_label.setObjectName("hintLabel")
        layout.addWidget(instruction_label)
        
        # Кнопки для ручного перемещения (дополнительная опция)
        move_buttons_layout = QHBoxLayout()
        
        move_up_btn = QPushButton("↑")
        move_up_btn.clicked.connect(lambda: self.move_employee_up(list_widget))
        move_buttons_layout.addWidget(move_up_btn)
        
        move_down_btn = QPushButton("↓")
        move_down_btn.clicked.connect(lambda: self.move_employee_down(list_widget))
        move_buttons_layout.addWidget(move_down_btn)
        
        layout.addLayout(move_buttons_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        if dialog.exec_() == QDialog.Accepted:
            # Получаем новый порядок сотрудников
            new_order = []
            for i in range(list_widget.count()):
                item = list_widget.item(i)
                new_order.append(item.text())
            
            # Обновляем порядок в базе данных
            if self.db.update_employee_order(self.current_period, new_order):
                # Порядок уже сохранён в БД; просто перерисовываем таблицу.
                self.load_data(self.current_period)

    def move_employee_up(self, list_widget):
        """Переместить выбранного сотрудника вверх"""
        current_row = list_widget.currentRow()
        if current_row > 0:
            current_item = list_widget.takeItem(current_row)
            list_widget.insertItem(current_row - 1, current_item)
            list_widget.setCurrentRow(current_row - 1)

    def move_employee_down(self, list_widget):
        """Переместить выбранного сотрудника вниз"""
        current_row = list_widget.currentRow()
        if current_row >= 0 and current_row < list_widget.count() - 1:
            current_item = list_widget.takeItem(current_row)
            list_widget.insertItem(current_row + 1, current_item)
            list_widget.setCurrentRow(current_row + 1)
    
    def load_periods(self):
        periods = sort_periods(self.db.periods)
        current = self.current_period
        with QSignalBlocker(self.period_combo):
            self.period_combo.clear()
            for period in periods:
                self.period_combo.addItem(period, period)
            if periods:
                index = self.period_combo.findData(current) if current else 0
                self.period_combo.setCurrentIndex(max(0, index))
        if periods:
            self.current_period = self.period_combo.currentData()
            self.load_data(self.current_period)
        else:
            self.current_period = None
            self.table.clear()
            self.table.setRowCount(0)
            self.table.setColumnCount(0)

    def period_changed(self, index):
        if index < 0:
            return
        period = self.period_combo.itemData(index)
        if not period or period == self.current_period:
            return

        if self.current_period and self.parent_app.unsaved_changes:
            choice = self.parent_app.confirm_unsaved_changes()
            if choice == "cancel":
                with QSignalBlocker(self.period_combo):
                    previous = self.period_combo.findData(self.current_period)
                    self.period_combo.setCurrentIndex(previous if previous >= 0 else 0)
                return
            if choice == "discard":
                self.parent_app.set_unsaved_changes(False)

        self.current_period = period
        self.parent_app.last_active_widget = self
        self.load_data(period)

    def _apply_item_visual(self, item, status=None):
        """Единая отрисовка ячейки: цвет статуса, текст и подсказка примечания."""
        if status is not None:
            item.status = status
        status = getattr(item, "status", STATUS_EMPTY)
        _, _, background, _ = self.parent_app.status_mapping.get(
            status, ("", "Пусто", QColor(Qt.white), 0)
        )
        item.setTextAlignment(Qt.AlignCenter)
        item.setBackground(background)
        item.setForeground(Qt.black)
        if item.has_note:
            item.setText(
                format_note_text(
                    item.note_data.get("end_time"), item.note_data.get("note_type")
                )
            )
            item.setToolTip(note_tooltip(item.note_data))
        else:
            item.setText("")
            item.setToolTip("")
        item.setData(Qt.UserRole, item.has_note)

    def _make_item(self, status, has_note=False, note_data=None):
        item = NoteItem("", has_note, note_data, status)
        self._apply_item_visual(item)
        return item

    def load_data(self, period):
        if not period:
            return

        self.table.is_loading = True
        self.table.setUpdatesEnabled(False)
        try:
            self.table.clear()
            self.current_period = period
            days = days_in_month(period)
            sundays = sunday_flags(period)
            self.sunday_columns = [index for index, sunday in enumerate(sundays) if sunday]

            schedule_data = self.db.load_schedule(period)
            employees_in_period = (schedule_data or {}).get("employees", [])
            if not employees_in_period:
                self.table.setRowCount(0)
                self.table.setColumnCount(0)
                return

            self.table.setRowCount(len(employees_in_period))
            self.table.setColumnCount(days + 4)
            for row, employee in enumerate(employees_in_period):
                self.table.setVerticalHeaderItem(row, QTableWidgetItem(employee["name"]))

            notes_data = schedule_data.get("notes", {})
            for row, employee in enumerate(employees_in_period):
                employee_name = employee["name"]
                employee_schedule = schedule_data["schedule"].get(
                    employee_name, [STATUS_EMPTY] * days
                )
                employee_notes = notes_data.get(employee_name, {})
                for day_index in range(days):
                    sunday = sundays[day_index]
                    note_data = employee_notes.get(str(day_index), {}) if not sunday else {}
                    has_note = bool(note_data)
                    status = (
                        STATUS_DAY_OFF
                        if sunday
                        else (
                            employee_schedule[day_index]
                            if day_index < len(employee_schedule)
                            else STATUS_EMPTY
                        )
                    )
                    item = NoteItem("", has_note, note_data, status)
                    self._apply_item_visual(item)
                    self.table.setItem(row, day_index, item)

            self.table.is_loading = False
            self.update_counters()
            self.apply_scale(self.parent_app.current_scale)
            self.update_header_style()
        except Exception:
            logger.exception("Ошибка загрузки данных периода %s", period)
        finally:
            self.table.is_loading = False
            self.table.setUpdatesEnabled(True)
            self.table.viewport().update()
    
    def update_header_style(self):
        """Обновляет стиль заголовков таблицы"""
        if not self.sunday_columns:
            # Если нет воскресных колонок, сбрасываем стиль
            self.table.horizontalHeader().setStyleSheet("")
            return
        
        # Создаем CSS стиль для воскресных заголовков
        style_sheet = ""
        for col in range(self.table.columnCount()):
            if col in self.sunday_columns:
                style_sheet += f"""
                    QHeaderView::section:nth-child({col + 1}) {{
                        background-color: #FFF1F2;
                        color: #B91C1C;
                        font-weight: 600;
                        padding: 6px 2px;
                        border-right: 1px solid #EEF2F7;
                        border-bottom: 1px solid #E2E8F0;
                    }}
                """
            else:
                style_sheet += f"""
                    QHeaderView::section:nth-child({col + 1}) {{
                        background-color: #F8FAFC;
                        color: #475569;
                        font-weight: 600;
                        padding: 6px 2px;
                        border-right: 1px solid #EEF2F7;
                        border-bottom: 1px solid #E2E8F0;
                    }}
                """
        
        # Применяем стиль к заголовку таблицы
        self.table.horizontalHeader().setStyleSheet(style_sheet)
    
    def update_counters(self):
        if not self.current_period or self.table.is_loading:
            return

        days = days_in_month(self.current_period)
        sundays = sunday_flags(self.current_period)
        self.sunday_columns = [index for index, sunday in enumerate(sundays) if sunday]
        working_counts = [0] * days

        for row in range(self.table.rowCount()):
            total_shifts = 0
            call_center_days = 0
            registry_days = 0
            total_hours = 0.0

            for col in range(days):
                if sundays[col]:
                    continue
                item = self.table.item(row, col)
                if not item:
                    continue

                status = getattr(item, "status", STATUS_EMPTY)
                if status in WORK_STATUSES:
                    total_shifts += 1
                    working_counts[col] += 1
                    if status == 0:
                        call_center_days += 1
                    elif status == 1:
                        registry_days += 1

                    worked = None
                    if isinstance(item, NoteItem) and item.has_note:
                        worked = item.note_data.get("worked_hours")
                    if worked is None:
                        worked = self.parent_app.status_mapping.get(status, (None, None, None, 0))[3]
                    total_hours += worked or 0
                elif isinstance(item, NoteItem) and item.has_note:
                    total_hours += item.note_data.get("worked_hours") or 0

            totals = [
                (days, f"{total_shifts}", 9),
                (days + 1, f"{registry_days}", 9),
                (days + 2, f"{call_center_days}", 9),
                (days + 3, f"{total_hours:.1f}".replace(".", ","), 8),
            ]
            for column, text, font_size in totals:
                if column >= self.table.columnCount():
                    continue
                cell = self.table.item(row, column)
                if not cell:
                    cell = QTableWidgetItem()
                    cell.setTextAlignment(Qt.AlignCenter)
                    cell.setFlags(cell.flags() ^ Qt.ItemIsEditable)
                    cell_font = QFont(self.table.font())
                    cell_font.setBold(True)
                    cell_font.setPointSizeF(max(8.0, float(font_size)))
                    cell.setFont(cell_font)
                    self.table.setItem(row, column, cell)
                cell.setText(text)
                cell.setBackground(QColor("#F8FAFC"))
                cell.setForeground(QColor("#334155"))

        parsed = parse_period(self.current_period)
        year, month = (parsed[0], parsed[1]) if parsed else (2000, 1)
        headers = []
        for col in range(days):
            weekday = datetime(year, month, col + 1).weekday()
            day_name = DAY_NAMES[weekday]
            count = 0 if sundays[col] else working_counts[col]
            headers.append(f"{col + 1}\n{day_name}\n({count})")
        headers += ["Смены", "Рег", "КЦ", "Часы"]
        self.table.setHorizontalHeaderLabels(headers)
        self.update_header_style()
    
    def show_context_menu(self, pos):
        self.parent_app.last_active_widget = self
        table = self.table
        selected = table.selectedIndexes()
        if not selected:
            return
        
        menu = QMenu()
        
        copy_action = QAction("Копировать", self)
        copy_action.triggered.connect(self.copy_selected)
        menu.addAction(copy_action)
        
        cut_action = QAction("Вырезать", self)
        cut_action.triggered.connect(self.cut_selected)
        menu.addAction(cut_action)
        
        paste_action = QAction("Вставить", self)
        paste_action.triggered.connect(self.paste_selected)
        menu.addAction(paste_action)
        
        menu.addSeparator()
        
        if len(selected) == 1:
            index = selected[0]
            if index.column() < table.columnCount() - 4:
                item = table.item(index.row(), index.column())
                
                sunday_cell = is_sunday(self.current_period, index.column())
                # Не позволяем добавлять примечания к воскресеньям
                if not sunday_cell and item:
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
        if is_sunday(self.current_period, index.column()):
            QMessageBox.warning(self, "Ошибка", "Нельзя добавлять примечания к воскресеньям")
            return
        if not isinstance(item, NoteItem):
            return

        note_data = item.note_data if item.has_note else {}
        dialog = NoteDialog(
            self, note_data.get("end_time", ""), note_data.get("note_type", "")
        )
        if dialog.exec_() != QDialog.Accepted:
            return

        end_time = dialog.get_end_time()
        note_type = dialog.get_note_type()
        has_note = has_meaningful_note(
            {"end_time": end_time, "note_type": note_type}
        )
        if has_note:
            worked_hours = compute_worked_hours(
                item.status, end_time, self.parent_app.status_mapping[item.status][3]
            )
            item.has_note = True
            item.note_data = {
                "end_time": end_time,
                "worked_hours": worked_hours,
                "note_type": note_type,
            }
        else:
            item.has_note = False
            item.note_data = {}

        self._apply_item_visual(item)
        self.table.viewport().update()
        self.update_counters()
        self.parent_app.set_unsaved_changes(True)
    
    def remove_note(self, index):
        item = self.table.item(index.row(), index.column())
        if isinstance(item, NoteItem):
            item.has_note = False
            item.note_data = {}
            self._apply_item_visual(item)
            self.table.viewport().update()
            self.update_counters()
            self.parent_app.set_unsaved_changes(True)
    
    def copy_selected(self):
        """Копирование выделенных ячеек"""
        selected = self.table.selectedIndexes()
        if not selected:
            return
        
        rows = defaultdict(list)
        for index in selected:
            rows[index.row()].append(index.column())
        
        self.copied_data = []
        for row, cols in rows.items():
            row_data = []
            for col in sorted(cols):
                if col >= self.table.columnCount() - 4:
                    continue  # Пропускаем итоговые столбцы
                
                item = self.table.item(row, col)
                status = getattr(item, "status", STATUS_EMPTY)
                note_data = {}
                if isinstance(item, NoteItem) and item.has_note:
                    note_data = item.note_data.copy()
                row_data.append((status, note_data))

            if row_data:
                self.copied_data.append(row_data)
    
    def cut_selected(self):
        """Вырезание выделенных ячеек"""
        selected = self.table.selectedIndexes()
        if not selected:
            return
        
        # Группируем ячейки по строкам
        rows = defaultdict(list)
        for index in selected:
            rows[index.row()].append(index.column())
        
        # Сохраняем данные
        self.copied_data = []
        for row, cols in rows.items():
            row_data = []
            for col in sorted(cols):
                if col >= self.table.columnCount() - 4:
                    continue  # Пропускаем итоговые столбцы
                
                item = self.table.item(row, col)
                status = getattr(item, "status", STATUS_EMPTY)
                note_data = {}
                if isinstance(item, NoteItem) and item.has_note:
                    note_data = item.note_data.copy()
                row_data.append((status, note_data))

                # Очищаем ячейку, кроме воскресений (они всегда нерабочие).
                if item and not is_sunday(self.current_period, col):
                    self.table.setItem(row, col, self._make_item(STATUS_EMPTY))
            
            if row_data:
                self.copied_data.append(row_data)
        
        self.update_counters()
        if hasattr(self, 'parent_app') and hasattr(self.parent_app, 'set_unsaved_changes'):
            self.parent_app.set_unsaved_changes(True)
    
    def paste_selected(self):
        """Вставка скопированных ячеек (воскресенья не изменяются)."""
        if not self.copied_data:
            return
        selected = self.table.selectedIndexes()
        if not selected:
            return

        min_row = min(index.row() for index in selected)
        min_col = min(index.column() for index in selected)
        for row_offset, row_data in enumerate(self.copied_data):
            for col_offset, (status, note_data) in enumerate(row_data):
                target_row = min_row + row_offset
                target_col = min_col + col_offset
                if target_row >= self.table.rowCount():
                    continue
                if target_col >= self.table.columnCount() - 4:
                    continue
                if is_sunday(self.current_period, target_col):
                    continue

                has_note = bool(note_data) and has_meaningful_note(note_data)
                item = self._make_item(status, has_note, note_data.copy() if has_note else {})
                self.table.setItem(target_row, target_col, item)

        self.update_counters()
        self.parent_app.set_unsaved_changes(True)
    
    def update_selected_status(self, status):
        selected = self.table.selectedIndexes()
        if not selected:
            return

        for index in selected:
            if index.column() >= self.table.columnCount() - 4:
                continue
            if is_sunday(self.current_period, index.column()):
                continue
            item = self.table.item(index.row(), index.column())
            if isinstance(item, NoteItem):
                item.status = status
                if item.has_note:
                    default_hours = self.parent_app.status_mapping.get(
                        status, (None, None, None, 0)
                    )[3]
                    item.note_data["worked_hours"] = compute_worked_hours(
                        status, item.note_data.get("end_time"), default_hours
                    )
                self._apply_item_visual(item)
            else:
                self.table.setItem(
                    index.row(), index.column(), self._make_item(status)
                )

        self.update_counters()
        self.parent_app.set_unsaved_changes(True)
    
    def apply_scale(self, scale_percent):
        """Применение масштаба к таблице"""
        if not self.table:
            return
        
        # Базовые значения (при масштабе 100%)
        base_row_height = 30
        base_font_size = 9.5
        base_header_font_size = 9.0
        base_day_width = 34
        base_result_width = 64
        
        # Рассчитываем новые значения
        scale_factor = scale_percent / 100.0
        
        # Устанавливаем высоту строк
        new_row_height = int(base_row_height * scale_factor)
        self.table.verticalHeader().setDefaultSectionSize(new_row_height)
        
        # Устанавливаем ширину столбцов дней
        days_in_month = self.table.columnCount() - 4
        for col in range(days_in_month):
            new_width = int(base_day_width * scale_factor)
            self.table.setColumnWidth(col, max(new_width, 10))  # Минимум 10 пикселей
        
        # Устанавливаем ширину итоговых столбцов
        if days_in_month < self.table.columnCount():
            new_result_width = int(base_result_width * scale_factor)
            self.table.setColumnWidth(days_in_month, max(new_result_width, 30))  # Смены
        if days_in_month + 1 < self.table.columnCount():
            self.table.setColumnWidth(days_in_month + 1, max(new_result_width, 30))  # Рег
        if days_in_month + 2 < self.table.columnCount():
            self.table.setColumnWidth(days_in_month + 2, max(new_result_width, 30))  # КЦ
        if days_in_month + 3 < self.table.columnCount():
            self.table.setColumnWidth(days_in_month + 3, max(new_result_width, 30))  # Часы
        
        # Обновляем шрифты
        font = self.table.font()
        font.setPointSizeF(base_font_size * scale_factor)
        self.table.setFont(font)
        
        # Обновляем шрифты заголовков
        header_font = self.table.horizontalHeader().font()
        header_font.setPointSizeF(base_header_font_size * scale_factor)
        self.table.horizontalHeader().setFont(header_font)
        self.table.verticalHeader().setFont(header_font)
        
        # Обновляем отображение
        self.table.viewport().update()
        
        # Обновляем стиль заголовков после изменения масштаба
        self.update_header_style()


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


class ExportSettingsDialog(QDialog):
    """Окно экспорта: папка для всего расписания и папка для активного месяца."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Экспорт")
        self.setModal(True)
        self.setMinimumWidth(500)
        
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # 1. Папка для сохранения всего расписания
        layout.addWidget(QLabel("Папка для сохранения всего расписания:"))
        row1 = QHBoxLayout()
        self.path_all_edit = QLineEdit()
        self.path_all_edit.setReadOnly(True)
        self.path_all_edit.setPlaceholderText("Не выбрана")
        row1.addWidget(self.path_all_edit)
        self.btn_change_all = QPushButton("Изменить")
        self.btn_change_all.clicked.connect(self._on_change_folder_all)
        row1.addWidget(self.btn_change_all)
        layout.addLayout(row1)
        
        layout.addSpacing(15)
        
        # 2. Папка для сохранения активного месяца
        layout.addWidget(QLabel("Папка для сохранения активного месяца:"))
        row2 = QHBoxLayout()
        self.path_month_edit = QLineEdit()
        self.path_month_edit.setReadOnly(True)
        self.path_month_edit.setPlaceholderText("Не выбрана")
        row2.addWidget(self.path_month_edit)
        self.btn_change_month = QPushButton("Изменить")
        self.btn_change_month.clicked.connect(self._on_change_folder_month)
        row2.addWidget(self.btn_change_month)
        layout.addLayout(row2)
        
        layout.addStretch()
        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        button_box.rejected.connect(self.accept)
        layout.addWidget(button_box)
    
    def showEvent(self, event):
        super().showEvent(event)
        self._refresh_paths()
    
    def _refresh_paths(self):
        if self.parent():
            self.path_all_edit.setText(self.parent().export_folder or "")
            self.path_month_edit.setText(self.parent().export_folder_single or "")
    
    def _on_change_folder_all(self):
        if self.parent():
            self.parent().select_export_folder()
            self._refresh_paths()
    
    def _on_change_folder_month(self):
        if self.parent():
            self.parent().select_export_folder_single()
            self._refresh_paths()


class ScheduleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("MyCompany", "WorkSchedule")
        
        # Инициализация сетевого менеджера
        self.network_manager = NetworkManager()
        self.db = ScheduleManager(self.network_manager)

        # Добавляем переменную для времени последнего сохранения
        self.last_save_time = None
        self.last_active_widget = None
        
        # Добавляем переменную для текущего масштаба
        self.current_scale = self.settings.value("scale", 100, type=int)
        
        # Загружаем последние открытые месяцы
        self.last_opened_periods = self.settings.value("last_opened_periods", [], type=list)
        
        self.export_folder = self.settings.value("export_folder", "")
        self.export_folder_single = self.settings.value("export_folder_single", "")
        
        # Настройки подключения
        self.server_host = self.settings.value("server_host", "localhost")
        try:
            self.server_port = int(self.settings.value("server_port", "8888"))
        except (TypeError, ValueError):
            self.server_port = 8888
        
        # Таймер для автосохранения
        self.auto_save_timer = QTimer()
        self.auto_save_timer.setSingleShot(True)
        self.auto_save_timer.timeout.connect(self.auto_save)
        self.last_change_time = None
        
        # Статус несохраненных изменений
        self.unsaved_changes = False

        # Периодическая попытка переподключения, если сервер был недоступен.
        self.reconnect_timer = QTimer(self)
        self.reconnect_timer.setInterval(60000)
        self.reconnect_timer.timeout.connect(self.try_reconnect)

        # Статус 3 — архив (рабочий день). Стандартная смена: 07:30–20:00.
        self.status_mapping = {
            0: ("", "Колл-центр", QColor("#30D5C8"), DEFAULT_WORK_HOURS),
            1: ("", "Регистратура", QColor("#7777FF"), DEFAULT_WORK_HOURS),
            2: ("", "Не работает", QColor("#FF7777"), 0),
            3: ("", "Архив", QColor("#FFB347"), DEFAULT_WORK_HOURS),
            4: ("", "Пусто", QColor("#FFFFFF"), 0)
        }

        apply_theme(QApplication.instance())
        self.initUI()
        self.init_shortcuts()
        self.network_manager.on_connection_change = self.on_connection_change

        # Автоподключение при запуске
        QTimer.singleShot(100, self.auto_connect)
        
        # Загружаем время последнего сохранения из настроек при старте
        self.load_last_save_time()
        
        # Таймер для обновления времени каждую минуту
        self.time_update_timer = QTimer()
        self.time_update_timer.timeout.connect(self.update_time_display)
        self.time_update_timer.start(60000)  # Каждую минуту
    
    def initUI(self):
        self.setWindowTitle("Рабочее расписание")
        self.setWindowIcon(QIcon(resource_path("schedule.ico")))
        
        self.central_widget = QWidget()
        self.central_widget.setObjectName("appRoot")
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(20, 16, 20, 16)
        self.layout.setSpacing(14)
        self.central_widget.setLayout(self.layout)
        self.setMinimumSize(1100, 720)
        
        # Создаем меню
        menubar = self.menuBar()
        
        # Меню "Настройки"
        settings_menu = menubar.addMenu('Настройки')
        connection_action = QAction('Настройки подключения', self)
        connection_action.triggered.connect(self.show_connection_dialog)
        settings_menu.addAction(connection_action)
        export_action = QAction('Экспорт', self)
        export_action.triggered.connect(self.show_export_settings_dialog)
        settings_menu.addAction(export_action)
        
        # Меню "Сотрудники"
        employees_menu = menubar.addMenu('Сотрудники')
        
        # Пункт "Управление сотрудниками"
        manage_employees_action = QAction('Управление сотрудники', self)
        manage_employees_action.triggered.connect(self.show_employee_manager)
        employees_menu.addAction(manage_employees_action)
        
        employees_menu.addSeparator()
        
        # Пункт "Добавить сотрудника" (для быстрого доступа)
        add_employee_action = QAction('Добавить сотрудника', self)
        add_employee_action.triggered.connect(self.add_employee_dialog)
        employees_menu.addAction(add_employee_action)
        
        # Меню "Масштаб"
        scale_menu = menubar.addMenu('Масштаб')
        self.scale_actions = []
        
        # Создаем варианты масштаба от 10% до 100% с шагом 10%
        for scale_value in range(10, 101, 10):
            scale_action = QAction(f'{scale_value}%', self)
            scale_action.triggered.connect(lambda checked, sv=scale_value: self.set_scale_by_menu(sv))
            scale_menu.addAction(scale_action)
            self.scale_actions.append(scale_action)
            
            # Отмечаем текущий масштаб
            if scale_value == self.current_scale:
                scale_action.setChecked(True)
        
        # Статусы соединения и сохранения — компактные плашки в шапке.
        self.connection_status = QLabel("Не подключено")
        self.connection_status.setObjectName("statusPill")
        set_pill_state(self.connection_status, "offline")

        self.save_status = QLabel("Сохранено")
        self.save_status.setObjectName("statusPill")
        set_pill_state(self.save_status, "saved")

        # Шапка приложения
        header = QFrame()
        header.setObjectName("appHeader")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(20, 16, 20, 16)
        header_layout.setSpacing(16)

        title_block = QVBoxLayout()
        title_block.setSpacing(2)
        title_label = QLabel("Рабочее расписание")
        title_label.setObjectName("appTitle")
        subtitle_label = QLabel("Планирование смен и контроль рабочего времени")
        subtitle_label.setObjectName("appSubtitle")
        title_block.addWidget(title_label)
        title_block.addWidget(subtitle_label)

        header_layout.addLayout(title_block)
        header_layout.addStretch()
        header_layout.addWidget(self.connection_status)
        header_layout.addWidget(self.save_status)
        self.layout.addWidget(header)

        # Панель действий
        toolbar = QFrame()
        toolbar.setObjectName("toolbarCard")
        self.buttons_layout = QHBoxLayout(toolbar)
        self.buttons_layout.setContentsMargins(12, 10, 12, 10)
        self.buttons_layout.setSpacing(8)

        self.new_month_btn = QPushButton("Новый месяц")
        self.new_month_btn.setFixedHeight(36)
        self.new_month_btn.setMinimumWidth(132)
        set_variant(self.new_month_btn, "primary")
        self.new_month_btn.clicked.connect(self.select_month_and_year)
        self.buttons_layout.addWidget(self.new_month_btn)

        self.add_employee_btn = QPushButton("Новый сотрудник")
        self.add_employee_btn.setFixedHeight(36)
        self.add_employee_btn.setMinimumWidth(146)
        self.add_employee_btn.clicked.connect(self.add_employee_dialog)
        self.buttons_layout.addWidget(self.add_employee_btn)

        self.add_to_period_btn = QPushButton("Добавить в месяц")
        self.add_to_period_btn.setFixedHeight(36)
        self.add_to_period_btn.setMinimumWidth(154)
        self.add_to_period_btn.clicked.connect(self.add_employees_to_period)
        self.buttons_layout.addWidget(self.add_to_period_btn)

        self.open_doc_btn = QPushButton("Открыть документ")
        self.open_doc_btn.setFixedHeight(36)
        self.open_doc_btn.setMinimumWidth(154)
        set_variant(self.open_doc_btn, "ghost")
        self.open_doc_btn.clicked.connect(self.open_document)
        self.buttons_layout.addWidget(self.open_doc_btn)

        self.buttons_layout.addStretch()

        self.save_btn = QPushButton("Сохранить")
        self.save_btn.setFixedHeight(36)
        self.save_btn.setMinimumWidth(132)
        set_variant(self.save_btn, "success")
        self.save_btn.clicked.connect(self.save_data)
        self.buttons_layout.addWidget(self.save_btn)

        self.layout.addWidget(toolbar)
        
        # Создаем сплиттер для вертикального расположения месяцев
        self.splitter = QSplitter(Qt.Vertical)
        self.splitter.setObjectName("mainSplitter")
        self.splitter.setHandleWidth(8)
        self.layout.addWidget(self.splitter)
        
        # Создаем два виджета месяцев
        self.month_widget1 = MonthWidget(self, self.db)
        self.month_widget2 = MonthWidget(self, self.db)
        
        # Добавляем виджеты в сплиттер
        self.splitter.addWidget(self.month_widget1)
        self.splitter.addWidget(self.month_widget2)
        
        # Устанавливаем размеры сплиттера (с восстановлением сохранённых)
        saved_sizes = self.settings.value("splitter_sizes", [300, 300], type=list)
        try:
            sizes = [int(value) for value in saved_sizes]
        except (TypeError, ValueError):
            sizes = [300, 300]
        self.splitter.setSizes(sizes if len(sizes) == 2 else [300, 300])
        
        # Легенда статусов и действия внизу
        legend_card = QFrame()
        legend_card.setObjectName("legendCard")
        self.legend_layout = QHBoxLayout(legend_card)
        self.legend_layout.setContentsMargins(14, 8, 14, 8)
        self.legend_layout.setSpacing(10)

        legend_caption = QLabel("СТАТУСЫ")
        legend_caption.setObjectName("sectionCaption")
        self.legend_layout.addWidget(legend_caption)
        self.legend_layout.addSpacing(4)

        for status, (icon, text, color, _) in self.status_mapping.items():
            legend_item = QHBoxLayout()
            legend_item.setSpacing(5)
            color_label = QLabel("")
            color_label.setFixedSize(10, 10)
            color_label.setStyleSheet(
                f"background: {color.name()}; border: 1px solid #CBD5E1; border-radius: 5px;"
            )
            legend_item.addWidget(color_label)

            text_label = QLabel(text)
            text_label.setObjectName("legendLabel")
            legend_item.addWidget(text_label)
            self.legend_layout.addLayout(legend_item)

        self.legend_layout.addStretch()

        # Метка для отображения времени последнего сохранения
        self.last_update_label = QLabel("Еще не сохранено")
        self.last_update_label.setObjectName("mutedLabel")
        self.legend_layout.addWidget(self.last_update_label)

        # Кнопка обновления в правом нижнем углу
        self.refresh_btn = QPushButton("Обновить")
        self.refresh_btn.setFixedHeight(32)
        self.refresh_btn.setMinimumWidth(96)
        set_variant(self.refresh_btn, "ghost")
        self.refresh_btn.clicked.connect(self.refresh_data)
        self.legend_layout.addWidget(self.refresh_btn)

        self.layout.addWidget(legend_card)

        # Сразу показываем локальные данные; после подключения они обновятся.
        self.load_initial_data()

        self.statusBar().showMessage("Готово")
    
    def show_employee_manager(self):
        """Показать диалог управления сотрудниками"""
        if not self.resolve_unsaved_changes():
            return
        dialog = EmployeeManagerDialog(self.db, self)
        dialog.exec_()
    
    def set_unsaved_changes(self, has_changes):
        """Установка статуса несохраненных изменений"""
        self.unsaved_changes = has_changes
        if has_changes:
            self.save_status.setText("Не сохранено")
            set_pill_state(self.save_status, "unsaved")
            # Запускаем таймер автосохранения
            self.last_change_time = datetime.now()
            self.auto_save_timer.start(5 * 60 * 1000)  # 5 минут - 300000 мс
        else:
            self.save_status.setText("Сохранено")
            set_pill_state(self.save_status, "saved")
            self.auto_save_timer.stop()  # Останавливаем таймер, если изменения сохранены
    
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
        """Подключение к серверу. При несохранённых данных сначала спрашиваем."""
        if self.unsaved_changes:
            choice = self.confirm_unsaved_changes()
            if choice == "cancel":
                return

        if self.network_manager.connect(self.server_host, self.server_port):
            self.connection_status.setText("Подключено")
            set_pill_state(self.connection_status, "connected")
            self.statusBar().showMessage("Успешное подключение к серверу", 3000)
            self.reconnect_timer.stop()
            # Очередь уже отправлена; забираем актуальные данные сервера.
            self.refresh_data(silent=True)
        else:
            self.connection_status.setText("Не подключено")
            set_pill_state(self.connection_status, "offline")
            self.statusBar().showMessage("Ошибка подключения к серверу", 3000)
            self.reconnect_timer.start()
            self.load_initial_data()

    def on_connection_change(self, connected):
        """Реакция на потерю/восстановление соединения."""
        if connected:
            self.connection_status.setText("Подключено")
            set_pill_state(self.connection_status, "connected")
            self.reconnect_timer.stop()
        else:
            self.connection_status.setText("Не подключено")
            set_pill_state(self.connection_status, "offline")
            self.reconnect_timer.start()

    def try_reconnect(self):
        """Фоновая (по таймеру) попытка переподключения без потери правок."""
        if self.network_manager.connected:
            self.reconnect_timer.stop()
            return
        if not self.settings.value("auto_connect", True, type=bool):
            return
        if self.network_manager.connect(self.server_host, self.server_port):
            self.statusBar().showMessage("Соединение с сервером восстановлено", 3000)
            if not self.unsaved_changes:
                self.refresh_data(silent=True)

    def confirm_unsaved_changes(self):
        """Спрашивает, что делать с несохранёнными изменениями.

        Возвращает "save", "discard" или "cancel".
        """
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Question)
        box.setWindowTitle("Несохранённые изменения")
        box.setText("Есть несохранённые изменения. Сохранить их?")
        save_button = box.addButton("Сохранить", QMessageBox.AcceptRole)
        discard_button = box.addButton("Не сохранять", QMessageBox.DestructiveRole)
        box.addButton("Отмена", QMessageBox.RejectRole)
        box.setDefaultButton(save_button)
        box.exec_()

        clicked = box.clickedButton()
        if clicked is save_button:
            return "save" if self.save_data() else "cancel"
        if clicked is discard_button:
            self.set_unsaved_changes(False)
            return "discard"
        return "cancel"

    def resolve_unsaved_changes(self):
        """True, если можно продолжать: изменения сохранены или отброшены."""
        if not self.unsaved_changes:
            return True
        return self.confirm_unsaved_changes() != "cancel"
    
    def show_connection_dialog(self):
        """Показать диалог подключения"""
        dialog = ConnectionDialog(self)
        dialog.host_edit.setText(self.server_host)
        dialog.port_edit.setText(str(self.server_port))
        if dialog.exec_() == QDialog.Accepted:
            self.server_host, self.server_port = dialog.get_connection_params()
            
            # Сохраняем настройки
            self.settings.setValue("server_host", self.server_host)
            self.settings.setValue("server_port", self.server_port)
            
            self.connect_to_server()
    
    def load_initial_data(self):
        """Загружает периоды и восстанавливает последние открытые месяцы."""
        self.month_widget1.load_periods()
        self.month_widget2.load_periods()
        periods = self.db.periods
        if not periods:
            self.apply_scale(self.current_scale)
            return

        if self.last_opened_periods and len(self.last_opened_periods) >= 2:
            period1 = self.last_opened_periods[0]
            period2 = self.last_opened_periods[1]
        else:
            period1 = periods[0]
            period2 = periods[-1] if len(periods) > 1 else periods[0]

        for widget, period in ((self.month_widget1, period1), (self.month_widget2, period2)):
            index = widget.period_combo.findData(period)
            if index < 0:
                index = 0
            with QSignalBlocker(widget.period_combo):
                widget.period_combo.setCurrentIndex(index)
            widget.current_period = widget.period_combo.itemData(index)
            widget.load_data(widget.current_period)

        self.apply_scale(self.current_scale)

    def refresh_data(self, silent=False):
        """Обновляет данные с сервера, сохраняя выбранные периоды."""
        if not silent and self.unsaved_changes:
            choice = self.confirm_unsaved_changes()
            if choice == "cancel":
                return

        if not self.network_manager.connected:
            self.statusBar().showMessage(
                "Сервер недоступен, используются локальные данные", 3000
            )
            return

        current_period1 = self.month_widget1.current_period
        current_period2 = self.month_widget2.current_period
        try:
            self.db._employees = []
            self.db._all_employees = []
            self.db._periods_cache = None
            self.db._cache = {}

            self.db.load_employees()
            self.db.get_periods()
            self.month_widget1.load_periods()
            self.month_widget2.load_periods()

            for widget, period in (
                (self.month_widget1, current_period1),
                (self.month_widget2, current_period2),
            ):
                if not period:
                    continue
                index = widget.period_combo.findData(period)
                if index >= 0:
                    with QSignalBlocker(widget.period_combo):
                        widget.period_combo.setCurrentIndex(index)
                widget.current_period = period
                widget.load_data(period)

            if self.last_save_time:
                total_minutes = (datetime.now() - self.last_save_time).seconds // 60
                diff = (
                    f"{total_minutes // 60} ч. {total_minutes % 60} мин. назад"
                    if total_minutes >= 60
                    else f"{total_minutes} мин. назад"
                )
                message = f"Данные обновлены с сервера. Последнее сохранение {diff}"
            else:
                message = f"Данные обновлены с сервера в {datetime.now():%H:%M}"
            self.statusBar().showMessage(message, 5000)
        except Exception as exc:
            logger.exception("Ошибка обновления данных")
            self.statusBar().showMessage(f"Ошибка обновления: {exc}", 5000)
    
    def select_month_and_year(self):
        """Открывает диалог выбора месяца и года."""
        dialog = MonthSelectionDialog(self)
        if dialog.exec_() != QDialog.Accepted:
            return

        period_name = normalize_period_name(dialog.get_period_name()) or dialog.get_period_name()
        base = base_period_name(period_name)
        existing = [
            self.month_widget1.period_combo.itemData(i)
            for i in range(self.month_widget1.period_combo.count())
        ]
        if any(base_period_name(period) == base for period in existing):
            exists_dialog = MonthExistsDialog(period_name, self)
            if exists_dialog.exec_() != QDialog.Accepted:
                return
            action = exists_dialog.get_action()
            if action == "existing":
                self.select_existing_period(base)
                return
            if action != "variant":
                return

            variant = exists_dialog.get_variant_name() or "Вариант 2"
            period_name = f"{base} ({variant})"

        self.create_period_with_dialog(period_name)

    def select_existing_period(self, period_name):
        """Переход к существующему периоду."""
        index = self.month_widget1.period_combo.findData(period_name)
        if index < 0:
            base = base_period_name(period_name)
            for i in range(self.month_widget1.period_combo.count()):
                if base_period_name(self.month_widget1.period_combo.itemData(i)) == base:
                    index = i
                    break
        if index < 0:
            self.statusBar().showMessage("Не удалось найти существующий месяц", 3000)
            return

        self.month_widget1.period_combo.setCurrentIndex(index)
        self.statusBar().showMessage(
            f"Перешли к существующему месяцу: {self.month_widget1.period_combo.itemData(index)}",
            3000,
        )

    def create_period_with_dialog(self, period_name):
        """Создание периода с диалогом выбора/копирования сотрудников."""
        if not self.resolve_unsaved_changes():
            return
        canonical = normalize_period_name(period_name) or period_name
        if canonical in self.db.periods:
            QMessageBox.warning(
                self, "Ошибка", f"Период '{canonical}' уже существует"
            )
            return
        dialog = NewPeriodDialog(self, self.db, period_name)
        if dialog.exec_() != QDialog.Accepted:
            return

        selected_names = dialog.get_selected_employees()
        copy_source = dialog.get_copy_source()
        if copy_source:
            source_data = self.db.load_schedule(copy_source)
            if not source_data:
                self.statusBar().showMessage(
                    f"Не удалось загрузить данные из периода {copy_source}", 3000
                )
                return
            selected_names = [emp["name"] for emp in source_data["employees"]]
        elif not selected_names:
            self.statusBar().showMessage("Не выбрано ни одного сотрудника", 3000)
            return

        if self.db.create_period(period_name, selected_names, copy_source):
            self.statusBar().showMessage(f"Период '{period_name}' создан", 3000)
        else:
            self.statusBar().showMessage(
                f"Период '{period_name}' создан локально, синхронизация повторится позже",
                5000,
            )

        self.month_widget1.load_periods()
        self.month_widget2.load_periods()
        index = self.month_widget1.period_combo.findData(period_name)
        if index >= 0:
            self.month_widget1.period_combo.setCurrentIndex(index)
    
    def delete_period(self, period, widget=None):
        """Удаление периода с подтверждением"""
        if not period:
            return
        if not self.resolve_unsaved_changes():
            return
        
        # Подтверждение удаления
        reply = QMessageBox.question(
            self,
            "Подтверждение удаления",
            f"Вы уверены, что хотите удалить период '{period}'?\n\nЭто действие нельзя отменить.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            if self.db.delete_period(period):
                self.month_widget1.load_periods()
                self.month_widget2.load_periods()
                self.statusBar().showMessage(f"Период '{period}' удален", 3000)
            else:
                self.statusBar().showMessage(f"Ошибка удаления периода '{period}'", 5000)
    
    def add_employees_to_period(self):
        """Добавляет сотрудников в уже созданный месяц"""
        current_widget = self.get_active_month_widget()
        if not current_widget or not current_widget.current_period:
            self.statusBar().showMessage("Сначала выберите месяц в одном из окон", 3000)
            return
        if not self.resolve_unsaved_changes():
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
                self.statusBar().showMessage("Добавлено сотрудников локально, будет синхронизировано при подключении", 3000)
            
            # Перезагружаем данные
            current_widget.load_data(current_widget.current_period)
    
    def get_active_month_widget(self):
        """Определяет активный виджет месяца (фокус, затем последний активный)."""
        focus = QApplication.focusWidget()
        while focus is not None:
            if focus is self.month_widget1:
                self.last_active_widget = self.month_widget1
                return self.month_widget1
            if focus is self.month_widget2:
                self.last_active_widget = self.month_widget2
                return self.month_widget2
            focus = focus.parent()
        return self.last_active_widget or self.month_widget1

    def save_data(self):
        """Сохраняет оба месяца. Возвращает True, если локальное сохранение удалось."""
        try:
            success_count = 0
            total_count = 0
            for widget in (self.month_widget1, self.month_widget2):
                if not widget.current_period:
                    continue
                total_count += 1
                schedule_data = self.prepare_save_data(widget)
                if schedule_data and self.db.save_schedule(
                    widget.current_period, schedule_data
                ):
                    success_count += 1

            if success_count == 0 and total_count > 0:
                self.statusBar().showMessage("Не удалось сохранить данные", 5000)
                return False
            if total_count == 0:
                return True

            self.update_last_save_time()
            if self.network_manager.connected:
                self.statusBar().showMessage(
                    f"Сохранено {success_count} из {total_count} месяцев на сервер", 3000
                )
                self.show_refresh_status()
            else:
                self.statusBar().showMessage(
                    f"Сохранено {success_count} из {total_count} месяцев локально "
                    "(сервер недоступен)",
                    5000,
                )

            all_saved = success_count == total_count
            if all_saved:
                self.set_unsaved_changes(False)
                if self.export_folder:
                    self.export_to_excel()
            return all_saved
        except Exception as exc:
            logger.exception("Ошибка сохранения")
            self.statusBar().showMessage(f"Ошибка сохранения: {exc}", 5000)
            return False

    def auto_save(self):
        """Автосохранение через 5 минут после последнего изменения."""
        if (
            self.unsaved_changes
            and self.last_change_time
            and (datetime.now() - self.last_change_time).total_seconds() >= 300
        ):
            self.statusBar().showMessage("Автоматическое сохранение...", 2000)
            self.save_data()
    
    def show_refresh_status(self):
        """Показать статус обновления с сервера."""
        self.refresh_data(silent=True)

    def prepare_save_data(self, widget):
        period = widget.current_period
        if not period:
            return None
        schedule_data = self.db.load_schedule(period)
        if not schedule_data:
            return None

        days = days_in_month(period)
        sundays = sunday_flags(period)
        positions = {
            employee["name"]: employee.get("position", "")
            for employee in schedule_data.get("employees", [])
        }
        schedule_data["schedule"] = {}
        schedule_data["notes"] = {}

        for row in range(widget.table.rowCount()):
            header = widget.table.verticalHeaderItem(row)
            if not header:
                continue
            emp_name = header.text()
            full_schedule = [STATUS_EMPTY] * days
            notes = {}
            for day_index in range(days):
                item = widget.table.item(row, day_index)
                if not item:
                    continue
                if sundays[day_index]:
                    full_schedule[day_index] = STATUS_DAY_OFF
                    continue
                full_schedule[day_index] = getattr(item, "status", STATUS_EMPTY)
                if isinstance(item, NoteItem) and item.has_note:
                    notes[str(day_index)] = item.note_data
            schedule_data["schedule"][emp_name] = full_schedule
            if notes:
                schedule_data["notes"][emp_name] = notes

        # Порядок сотрудников в таблице — источник истины для периода.
        schedule_data["employees"] = [
            {
                "name": widget.table.verticalHeaderItem(row).text(),
                "position": positions.get(
                    widget.table.verticalHeaderItem(row).text(), ""
                ),
            }
            for row in range(widget.table.rowCount())
            if widget.table.verticalHeaderItem(row)
        ]
        return schedule_data
    
    def update_selected_status(self, status):
        """Обновляет статус для активного виджета"""
        active_widget = self.get_active_month_widget()
        if active_widget:
            active_widget.update_selected_status(status)
    
    def add_employee_dialog(self):
        """Быстрое добавление сотрудника через диалог"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Добавить сотрудника")
        dialog.setModal(True)
        
        layout = QVBoxLayout(dialog)
        
        form_layout = QFormLayout()
        
        name_edit = QLineEdit()
        form_layout.addRow("ФИО сотрудника:", name_edit)
        
        layout.addLayout(form_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        if dialog.exec_() == QDialog.Accepted:
            name = name_edit.text().strip()
            
            if not name:
                QMessageBox.warning(self, "Ошибка", "ФИО сотрудника не может быть пустым")
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
        
    def show_export_settings_dialog(self):
        """Показать окно настроек экспорта."""
        dialog = ExportSettingsDialog(self)
        dialog.exec_()
    
    def select_export_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения всего расписания")
        if folder:
            self.export_folder = folder
            self.settings.setValue("export_folder", folder)
            self.statusBar().showMessage(f"Папка для всего расписания: {folder}", 3000)
    
    def select_export_folder_single(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения активного месяца")
        if folder:
            self.export_folder_single = folder
            self.settings.setValue("export_folder_single", folder)
            self.statusBar().showMessage(f"Папка для активного месяца: {folder}", 3000)
    
    def _period_to_filename(self, period):
        """Формирует безопасное имя файла из названия периода."""
        return period_filename(period)
    
    def export_month_to_excel(self, widget):
        """Экспорт таблицы указанного виджета месяца в отдельный файл (перезапись при повторном экспорте)."""
        if not widget or not widget.current_period:
            self.statusBar().showMessage("Нет месяца для экспорта", 3000)
            return
        if not self.resolve_unsaved_changes():
            return
        self._do_export_month_to_excel(widget.current_period)
    
    def _do_export_month_to_excel(self, period):
        """Общая логика экспорта одного периода в Excel."""
        try:
            folder = self.export_folder_single or self.export_folder
            if not folder:
                self.select_export_folder_single()
                folder = self.export_folder_single or self.export_folder
                if not folder:
                    self.statusBar().showMessage("Экспорт отменён — не выбрана папка", 3000)
                    return
            
            schedule_data = self.db.load_schedule(period)
            if not schedule_data or not schedule_data.get('schedule'):
                self.statusBar().showMessage("Нет данных по выбранному месяцу для экспорта", 3000)
                return
            
            file_name = self._period_to_filename(period)
            file_path = os.path.join(folder, file_name)
            os.makedirs(folder, exist_ok=True)
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Расписание"
            
            header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            sunday_header_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
            center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            bold_font = Font(bold=True)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                top=Side(style='thin'), bottom=Side(style='thin'))
            
            status_styles = {
                0: PatternFill(start_color="30D5C8", end_color="30D5C8", fill_type="solid"),
                1: PatternFill(start_color="8080FF", end_color="8080FF", fill_type="solid"),
                2: PatternFill(start_color="FF7777", end_color="FF7777", fill_type="solid"),
                3: PatternFill(start_color="FFB347", end_color="FFB347", fill_type="solid"),  # Архив
                4: PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            }
            total_shifts_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
            registry_fill = PatternFill(start_color="8080FF", end_color="8080FF", fill_type="solid")
            call_center_fill = PatternFill(start_color="30D5C8", end_color="30D5C8", fill_type="solid")
            hours_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
            
            employees = schedule_data.get('employees', [])
            notes_data = schedule_data.get('notes', {})
            
            parsed = parse_period(period)
            if parsed:
                year, month_num = parsed[0], parsed[1]
                days_in_month = month_length(period)
                sundays = sunday_flags(period)
            else:
                days_in_month = len(next(iter(schedule_data['schedule'].values())))
                year = datetime.now().year
                month_num = datetime.now().month
                sundays = [False] * days_in_month
            day_mapping = list(range(1, days_in_month + 1))

            working_counts = [0] * days_in_month
            for emp in employees:
                emp_schedule = schedule_data['schedule'].get(emp["name"], [2] * days_in_month)
                for col, actual_day in enumerate(day_mapping):
                    day_index = actual_day - 1
                    if emp_schedule[day_index] in [0, 1, 3]:  # Добавлен статус 3 (Архив)
                        working_counts[col] += 1
            
            RESULT_COLUMNS_START = days_in_month + 2
            RESULT_COLUMNS = {
                "Смены": RESULT_COLUMNS_START,
                "Рег": RESULT_COLUMNS_START + 1,
                "КЦ": RESULT_COLUMNS_START + 2,
                "Часы": RESULT_COLUMNS_START + 3
            }
            total_columns = days_in_month + 5
            start_row = 1
            
            ws.cell(row=start_row, column=1, value=period).font = Font(bold=True, size=16)
            ws.cell(row=start_row, column=1, value=period).fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=total_columns)
            start_row += 1
            
            ws.cell(row=start_row, column=1, value="Сотрудник").fill = header_fill
            ws.cell(row=start_row, column=1).alignment = center_alignment
            ws.cell(row=start_row, column=1).font = bold_font
            ws.cell(row=start_row, column=1).border = thin_border
            
            for col, actual_day in enumerate(day_mapping):
                day_of_week = datetime(year, month_num, actual_day).weekday()
                day_name = DAY_NAMES[day_of_week]
                if sundays[col]:
                    header_fill_type = sunday_header_fill
                    header_text = f"{actual_day}\n{day_name}\n(0)"
                else:
                    header_fill_type = header_fill
                    header_text = f"{actual_day}\n{day_name}\n({working_counts[col]})"
                cell = ws.cell(row=start_row, column=col + 2, value=header_text)
                cell.fill = header_fill_type
                cell.alignment = center_alignment
                cell.font = bold_font
                cell.border = thin_border
            
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
            
            ws.column_dimensions['A'].width = 35
            for col in range(2, days_in_month + 2):
                ws.column_dimensions[get_column_letter(col)].width = 6
            ws.column_dimensions[get_column_letter(RESULT_COLUMNS["Смены"])].width = 8
            ws.column_dimensions[get_column_letter(RESULT_COLUMNS["Рег"])].width = 8
            ws.column_dimensions[get_column_letter(RESULT_COLUMNS["КЦ"])].width = 8
            ws.column_dimensions[get_column_letter(RESULT_COLUMNS["Часы"])].width = 8
            
            for emp in employees:
                emp_name = emp["name"]
                ws.cell(row=start_row, column=1, value=emp_name).font = bold_font
                schedule = schedule_data['schedule'].get(emp["name"], [2] * days_in_month)
                emp_notes = notes_data.get(emp["name"], {})
                total_shifts = 0
                total_hours = 0.0
                call_center_days = 0
                registry_days = 0
                archive_days = 0
                
                for col, actual_day in enumerate(day_mapping):
                    day_index = actual_day - 1
                    status = schedule[day_index]
                    if status == 0:
                        call_center_days += 1
                        total_shifts += 1
                        total_hours += emp_notes.get(str(day_index), {}).get('worked_hours', self.status_mapping[0][3])
                    elif status == 1:
                        registry_days += 1
                        total_shifts += 1
                        total_hours += emp_notes.get(str(day_index), {}).get('worked_hours', self.status_mapping[1][3])
                    elif status == 3:  # Архив
                        archive_days += 1
                        total_shifts += 1
                        total_hours += emp_notes.get(str(day_index), {}).get('worked_hours', self.status_mapping[3][3])
                    
                    cell = ws.cell(row=start_row, column=col + 2)
                    note_data = emp_notes.get(str(day_index))
                    cell.value = (
                        format_note_text(
                            note_data.get("end_time"), note_data.get("note_type")
                        )
                        if note_data and not sundays[day_index]
                        else ""
                    )
                    cell.alignment = center_alignment
                    cell.border = thin_border
                    if status in status_styles:
                        cell.fill = status_styles[status]
                
                hours_value = round(total_hours, 1)
                ws.cell(row=start_row, column=RESULT_COLUMNS["Смены"], value=total_shifts).fill = total_shifts_fill
                ws.cell(row=start_row, column=RESULT_COLUMNS["Смены"]).alignment = center_alignment
                ws.cell(row=start_row, column=RESULT_COLUMNS["Смены"]).font = bold_font
                ws.cell(row=start_row, column=RESULT_COLUMNS["Смены"]).border = thin_border
                ws.cell(row=start_row, column=RESULT_COLUMNS["Рег"], value=registry_days).fill = registry_fill
                ws.cell(row=start_row, column=RESULT_COLUMNS["Рег"]).alignment = center_alignment
                ws.cell(row=start_row, column=RESULT_COLUMNS["Рег"]).font = bold_font
                ws.cell(row=start_row, column=RESULT_COLUMNS["Рег"]).border = thin_border
                ws.cell(row=start_row, column=RESULT_COLUMNS["КЦ"], value=call_center_days).fill = call_center_fill
                ws.cell(row=start_row, column=RESULT_COLUMNS["КЦ"]).alignment = center_alignment
                ws.cell(row=start_row, column=RESULT_COLUMNS["КЦ"]).font = bold_font
                ws.cell(row=start_row, column=RESULT_COLUMNS["КЦ"]).border = thin_border
                ws.cell(row=start_row, column=RESULT_COLUMNS["Часы"], value=hours_value).fill = hours_fill
                ws.cell(row=start_row, column=RESULT_COLUMNS["Часы"]).alignment = center_alignment
                ws.cell(row=start_row, column=RESULT_COLUMNS["Часы"]).font = bold_font
                ws.cell(row=start_row, column=RESULT_COLUMNS["Часы"]).border = thin_border
                start_row += 1
            
            wb.save(file_path)
            self.statusBar().showMessage(f"Месяц экспортирован: {file_path}", 5000)
        except PermissionError:
            self.statusBar().showMessage("Ошибка: нет прав для записи в папку", 5000)
        except Exception as e:
            self.statusBar().showMessage(f"Ошибка экспорта месяца: {str(e)}", 5000)
    
    def export_to_excel(self):
        """Экспорт всех данных в Excel (без диалогов; папка задаётся в настройках)."""
        try:
            if not self.export_folder:
                return

            file_name = "Расписание_все_месяцы.xlsx"
            file_path = os.path.join(self.export_folder, file_name)

            # Создаем папку если она не существует
            os.makedirs(self.export_folder, exist_ok=True)

            wb = Workbook()
            ws = wb.active
            ws.title = "Расписание"

            header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            sunday_header_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")  # Красный для воскресений
            center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            bold_font = Font(bold=True)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                            top=Side(style='thin'), bottom=Side(style='thin'))
            
            status_styles = {
                0: PatternFill(start_color="30D5C8", end_color="30D5C8", fill_type="solid"),
                1: PatternFill(start_color="8080FF", end_color="8080FF", fill_type="solid"),
                2: PatternFill(start_color="FF7777", end_color="FF7777", fill_type="solid"),
                3: PatternFill(start_color="FFB347", end_color="FFB347", fill_type="solid"),  # Архив
                4: PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            }

            total_shifts_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
            registry_fill = PatternFill(start_color="8080FF", end_color="8080FF", fill_type="solid")
            call_center_fill = PatternFill(start_color="30D5C8", end_color="30D5C8", fill_type="solid")
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
                    days_in_month = (
                        month_length(period)
                        if parse_period(period)
                        else len(next(iter(schedule_data["schedule"].values())))
                    )
                    max_working_days = max(max_working_days, days_in_month)
            
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

            sorted_periods = sorted(periods, key=period_sort_key)

            for period in sorted_periods:
                schedule_data = self.db.load_schedule(period)
                if not schedule_data or not schedule_data.get('schedule'):
                    continue

                # Используем полное название периода
                period_name = period
                
                employees = schedule_data.get('employees', [])
                if not employees:
                    continue

                notes_data = schedule_data.get('notes', {})

                parsed = parse_period(period)
                if parsed:
                    year, month_num = parsed[0], parsed[1]
                    days_in_month = month_length(period)
                    sundays = sunday_flags(period)
                else:
                    days_in_month = len(next(iter(schedule_data['schedule'].values())))
                    year, month_num = datetime.now().year, datetime.now().month
                    sundays = [False] * days_in_month
                day_mapping = list(range(1, days_in_month + 1))

                working_counts = [0] * days_in_month
                for emp in employees:
                    emp_schedule = schedule_data['schedule'].get(emp["name"], [2]*days_in_month)
                    for col, actual_day in enumerate(day_mapping):
                        day_index = actual_day - 1
                        if emp_schedule[day_index] in [0, 1, 3]:  # Добавлен статус 3 (Архив)
                            working_counts[col] += 1

                total_columns = max_working_days + 5
                # Настройки теста EXCEL строки с названием месяца
                ws.cell(row=start_row, column=1, value=period_name).font = Font(bold=True, size=16)
                ws.cell(row=start_row, column=1, value=period_name).fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
                ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=total_columns)
                start_row += 1

                # Заголовки столбцов
                ws.cell(row=start_row, column=1, value="Сотрудник").fill = header_fill
                ws.cell(row=start_row, column=1).alignment = center_alignment
                ws.cell(row=start_row, column=1).font = bold_font
                ws.cell(row=start_row, column=1).border = thin_border

                for col, actual_day in enumerate(day_mapping):
                    day_of_week = datetime(year, month_num, actual_day).weekday()
                    day_name = DAY_NAMES[day_of_week]
                    if sundays[col]:
                        count = 0
                        header_fill_type = sunday_header_fill
                    else:
                        count = working_counts[col]
                        header_fill_type = header_fill
                    header_text = f"{actual_day}\n{day_name}\n({count})"
                    
                    cell = ws.cell(row=start_row, column=col+2, value=header_text)
                    cell.fill = header_fill_type
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
                
                # Ширина столбцов EXCEL ФИО 
                ws.column_dimensions['A'].width = 25
                
                # Ширина столбцов EXCEL дней
                for col in range(2, max_working_days + 2):
                    ws.column_dimensions[get_column_letter(col)].width = 6

                # Устанавливаем высоту строк для данных
                for row in range(start_row, start_row + len(employees)):
                    ws.row_dimensions[row].height = 28  # Высота строки в пунктах (1 пункт ≈ 1/72 дюйма)
                
                # Ширина столбцов EXCEL Смены\Рег\КЦ\Часы
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
                    archive_days = 0
                    
                    for col, actual_day in enumerate(day_mapping):
                        day_index = actual_day - 1
                        status = schedule[day_index]
                        
                        if status == 0:
                            call_center_days += 1
                            total_shifts += 1
                            if str(day_index) in emp_notes:
                                total_hours += emp_notes[str(day_index)].get('worked_hours', DEFAULT_WORK_HOURS)
                            else:
                                total_hours += self.status_mapping[0][3]
                        elif status == 1:
                            registry_days += 1
                            total_shifts += 1
                            if str(day_index) in emp_notes:
                                total_hours += emp_notes[str(day_index)].get('worked_hours', DEFAULT_WORK_HOURS)
                            else:
                                total_hours += self.status_mapping[1][3]
                        elif status == 3:  # Архив
                            archive_days += 1
                            total_shifts += 1
                            if str(day_index) in emp_notes:
                                total_hours += emp_notes[str(day_index)].get('worked_hours', DEFAULT_WORK_HOURS)
                            else:
                                total_hours += self.status_mapping[3][3]
                        
                        cell = ws.cell(row=start_row, column=col+2)
                        is_sunday_day = sundays[day_index]

                        note_data = emp_notes.get(str(day_index))
                        cell.value = (
                            format_note_text(
                                note_data.get("end_time"), note_data.get("note_type")
                            )
                            if note_data and not is_sunday_day
                            else ""
                        )
                        
                        cell.alignment = center_alignment
                        cell.border = thin_border
                        
                        if status in status_styles:
                            cell.fill = status_styles[status]
                    
                    hours_value = round(total_hours, 1)
                    
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
                    
                    cell = ws.cell(row=start_row, column=RESULT_COLUMNS["Часы"], value=hours_value)
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

    # ========== МЕТОДЫ ДЛЯ ОТОБРАЖЕНИЯ ВРЕМЕНИ И МАСШТАБА ==========
    
    def update_last_save_time(self):
        """Обновление времени последнего сохранения"""
        current_time = datetime.now()
        self.last_save_time = current_time
        
        # Форматируем время для отображения
        hours = current_time.hour
        minutes = current_time.minute
        
        # Обновляем текст метки
        self.last_update_label.setText(f"Сохранено {hours:02d}:{minutes:02d}")
        
        # Сохраняем время в настройки
        self.save_last_save_time()
    
    def save_last_save_time(self):
        """Сохранение времени последнего сохранения в настройки"""
        if self.last_save_time:
            time_str = self.last_save_time.strftime("%Y-%m-%d %H:%M:%S")
            self.settings.setValue("last_save_time", time_str)
    
    def load_last_save_time(self):
        """Загрузка времени последнего сохранения из настроек"""
        time_str = self.settings.value("last_save_time")
        if time_str:
            try:
                self.last_save_time = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
                hours = self.last_save_time.hour
                minutes = self.last_save_time.minute
                self.last_update_label.setText(f"Сохранено {hours:02d}:{minutes:02d}")
            except ValueError:
                self.last_update_label.setText("Еще не сохранено")
        else:
            self.last_update_label.setText("Еще не сохранено")
    
    def update_time_display(self):
        """Обновление отображения времени последнего сохранения"""
        if self.last_save_time:
            current_time = datetime.now()
            
            # Рассчитываем разницу во времени
            time_diff = current_time - self.last_save_time
            
            if time_diff.days > 0:
                # Если прошло больше суток, показываем дату
                if time_diff.days == 1:
                    self.last_update_label.setText(f"Сохранено вчера в {self.last_save_time.hour:02d}:{self.last_save_time.minute:02d}")
                else:
                    self.last_update_label.setText(f"Сохранено {self.last_save_time.day:02d}.{self.last_save_time.month:02d} в {self.last_save_time.hour:02d}:{self.last_save_time.minute:02d}")
            else:
                # Рассчитываем часы и минуты разницы
                total_minutes = time_diff.seconds // 60
                hours_diff = total_minutes // 60
                minutes_diff = total_minutes % 60
                
                if hours_diff > 0:
                    self.last_update_label.setText(f"Сохранено {hours_diff} ч. {minutes_diff} мин. назад")
                else:
                    self.last_update_label.setText(f"Сохранено {minutes_diff} мин. назад")
    
    def set_scale_by_menu(self, scale_percent):
        """Изменение масштаба через меню"""
        self.current_scale = scale_percent
        self.apply_scale(scale_percent)
        
        # Сохраняем настройки масштаба
        self.settings.setValue("scale", scale_percent)
        
        # Показываем сообщение
        self.statusBar().showMessage(f"Масштаб изменен на {scale_percent}%", 2000)
    
    def apply_scale(self, scale_percent):
        """Применение масштаба к обеим таблицам"""
        # Применяем масштаб к обоим виджетам месяцев
        if self.month_widget1:
            self.month_widget1.apply_scale(scale_percent)
        if self.month_widget2:
            self.month_widget2.apply_scale(scale_percent)
    
    def save_last_opened_periods(self):
        """Сохранение последних открытых месяцев"""
        current_periods = []
        if self.month_widget1.current_period:
            current_periods.append(self.month_widget1.current_period)
        if self.month_widget2.current_period:
            current_periods.append(self.month_widget2.current_period)
        
        if current_periods:
            self.settings.setValue("last_opened_periods", current_periods)
    
    def closeEvent(self, event):
        """Закрытие окна с проверкой несохранённых изменений."""
        if self.unsaved_changes:
            choice = self.confirm_unsaved_changes()
            if choice == "cancel":
                event.ignore()
                return

        self.save_last_save_time()
        self.save_last_opened_periods()
        self.settings.setValue("splitter_sizes", self.splitter.sizes())
        self.auto_save_timer.stop()
        self.reconnect_timer.stop()
        self.network_manager.disconnect()
        event.accept()


if __name__ == "__main__":
    setup_logging("schedule_client")
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    apply_theme(app)
    window = ScheduleApp()
    window.showMaximized()
    sys.exit(app.exec_())