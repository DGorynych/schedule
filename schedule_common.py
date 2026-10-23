"""Общие помощники для клиента и сервера расписания.

Здесь собрано то, что раньше дублировалось в schedule_client.py и
schedule_server.py: разбор названий периодов, работа с календарём,
форматирование примечаний, пути к данным/ресурсам, логирование и
протокол обмена сообщениями.
"""

from __future__ import annotations

import calendar
import json
import logging
import os
import re
import socket
import sqlite3
import struct
import sys
from datetime import datetime
from logging.handlers import RotatingFileHandler
from typing import Any

# ---------------------------------------------------------------------------
# Константы предметной области
# ---------------------------------------------------------------------------

MONTH_NAMES = {
    1: "Январь",
    2: "Февраль",
    3: "Март",
    4: "Апрель",
    5: "Май",
    6: "Июнь",
    7: "Июль",
    8: "Август",
    9: "Сентябрь",
    10: "Октябрь",
    11: "Ноябрь",
    12: "Декабрь",
}
MONTH_NAMES_CASEFOLD = {name.casefold(): num for num, name in MONTH_NAMES.items()}

DAY_NAMES = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]

STATUS_CALL_CENTER = 0
STATUS_REGISTRY = 1
STATUS_DAY_OFF = 2
STATUS_ARCHIVE = 3
STATUS_EMPTY = 4

WORK_STATUSES = (STATUS_CALL_CENTER, STATUS_REGISTRY, STATUS_ARCHIVE)
NOTE_TYPES = {
    "о": "Ответственный",
    "д": "ДМС",
    "*": "Нежелание работать",
    "а": "Архив",
}

DEFAULT_END_TIME = "20:00"
WORKDAY_START_HOUR = 7
WORKDAY_START_MINUTE = 30
WORKDAY_START_MINUTES = WORKDAY_START_HOUR * 60 + WORKDAY_START_MINUTE
DEFAULT_WORK_HOURS = 12.5

# ---------------------------------------------------------------------------
# Периоды
# ---------------------------------------------------------------------------

_DISPLAY_PERIOD_RE = re.compile(
    r"^\s*(?P<month>[^\W\d_]+)\s+(?P<year>\d{4})\s*(?:\((?P<variant>[^)]*)\))?\s*$",
    re.UNICODE,
)
_NUMERIC_PERIOD_RE = re.compile(
    r"^\s*(?P<year>\d{4})-(?P<month>\d{1,2})\s*(?:\((?P<variant>[^)]*)\))?\s*$"
)


def parse_period(name: str | None) -> tuple[int, int, str] | None:
    """Разбирает период в (год, месяц, вариант).

    Поддерживаются оба исторических формата: «Январь 2026» и «2026-01».
    Возвращает None, если название не похоже на период.
    """
    if not name:
        return None

    match = _DISPLAY_PERIOD_RE.match(name)
    if match:
        month = MONTH_NAMES_CASEFOLD.get(match.group("month").casefold())
        if month:
            return int(match.group("year")), month, (match.group("variant") or "").strip()

    match = _NUMERIC_PERIOD_RE.match(name)
    if match:
        month = int(match.group("month"))
        if 1 <= month <= 12:
            return int(match.group("year")), month, (match.group("variant") or "").strip()

    return None


def format_period(year: int, month: int, variant: str = "") -> str:
    """Собирает каноническое отображаемое имя периода."""
    base = f"{MONTH_NAMES.get(month, str(month))} {year}"
    variant = (variant or "").strip()
    return f"{base} ({variant})" if variant else base


def normalize_period_name(name: str | None) -> str | None:
    """Приводит период к каноническому виду «Месяц Год (вариант)»."""
    parsed = parse_period(name)
    if not parsed:
        return None
    return format_period(*parsed)


def base_period_name(name: str) -> str:
    """Название периода без варианта."""
    parsed = parse_period(name)
    if parsed:
        return format_period(parsed[0], parsed[1])
    return (name or "").split(" (")[0].strip()


def period_sort_key(name: str) -> tuple[int, int, str]:
    """Ключ сортировки периодов: год, месяц, вариант."""
    parsed = parse_period(name)
    if parsed:
        return parsed[0], parsed[1], parsed[2].casefold()
    return 0, 0, (name or "").casefold()


def sort_periods(periods: list[str]) -> list[str]:
    return sorted(periods, key=period_sort_key)


def days_in_month(period: str, default: int = 31) -> int:
    parsed = parse_period(period)
    if not parsed:
        return default
    return calendar.monthrange(parsed[0], parsed[1])[1]


def is_sunday(period: str, day_index: int) -> bool:
    parsed = parse_period(period)
    if not parsed or not 0 <= day_index < 31:
        return False
    year, month, _ = parsed
    if not 1 <= day_index + 1 <= calendar.monthrange(year, month)[1]:
        return False
    return datetime(year, month, day_index + 1).weekday() == 6


def sunday_flags(period: str) -> list[bool]:
    parsed = parse_period(period)
    if not parsed:
        return [False] * 31
    year, month, _ = parsed
    return [datetime(year, month, day).weekday() == 6
            for day in range(1, days_in_month(period) + 1)]


# ---------------------------------------------------------------------------
# Примечания и часы
# ---------------------------------------------------------------------------


def format_note_text(end_time: str | None, note_type: str | None) -> str:
    """Короткий текст для ячейки: «15-о», «д», «15» или пусто."""
    end_time = (end_time or "").strip()
    note_type = (note_type or "").strip()
    hour_part = end_time.split(":")[0] if ":" in end_time else end_time

    if note_type:
        if end_time and end_time != DEFAULT_END_TIME:
            return f"{hour_part}-{note_type}"
        return note_type
    if end_time and end_time != DEFAULT_END_TIME:
        return hour_part
    return ""


def note_tooltip(note_data: dict[str, Any] | None) -> str:
    """Подсказка для ячейки с примечанием."""
    if not note_data:
        return ""
    parts: list[str] = []
    worked_hours = note_data.get("worked_hours")
    if worked_hours:
        parts.append(f"Отработано: {worked_hours:g} часов")
    label = NOTE_TYPES.get((note_data.get("note_type") or "").strip())
    if label:
        parts.append(label)
    return "\n".join(parts)


def has_meaningful_note(note_data: dict[str, Any] | None) -> bool:
    """Есть ли у примечания смысл (тип или нестандартное время)."""
    if not note_data:
        return False
    end_time = (note_data.get("end_time") or "").strip()
    note_type = (note_data.get("note_type") or "").strip()
    return bool(note_type) or (bool(end_time) and end_time != DEFAULT_END_TIME)


def compute_worked_hours(
    status: int,
    end_time: str | None,
    default_hours: float = DEFAULT_WORK_HOURS,
) -> float:
    """Считает отработанные часы по статусу и времени окончания.

    Для нерабочих статусов — 0. Для рабочих: время окончания минус начало
    смены (07:30). Результат ограничен диапазоном 0..24, чтобы опечатка
    вроде «07:00» не давала отрицательные часы.
    """
    if status not in WORK_STATUSES:
        return 0.0
    if not end_time:
        return float(default_hours)
    try:
        hours, minutes = (int(part) for part in end_time.split(":")[:2])
    except (TypeError, ValueError):
        return float(default_hours)
    total_minutes = hours * 60 + minutes - WORKDAY_START_MINUTES
    return round(min(24.0, max(0.0, total_minutes / 60.0)), 2)


# ---------------------------------------------------------------------------
# Файлы и ресурсы
# ---------------------------------------------------------------------------

_UNSAFE_FILENAME_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sanitize_filename(name: str, max_length: int = 120) -> str:
    """Безопасное имя файла без разделителей пути."""
    cleaned = _UNSAFE_FILENAME_RE.sub("_", name or "")
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned[:max_length] or "period"


def period_filename(period: str) -> str:
    safe = sanitize_filename(period.replace(" (", "_").replace(")", ""))
    return f"Расписание_{safe}.xlsx"


def resource_path(relative_path: str) -> str:
    """Путь к ресурсу и в исходниках, и внутри PyInstaller onefile."""
    candidates = []
    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        candidates.append(os.path.join(bundle_dir, relative_path))
    if getattr(sys, "frozen", False):
        candidates.append(os.path.join(os.path.dirname(sys.executable), relative_path))
    candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path))
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    return candidates[0] if candidates else relative_path


def app_data_dir() -> str:
    """Каталог для данных приложения (не зависит от рабочей директории)."""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "WorkSchedule")
    os.makedirs(path, exist_ok=True)
    return path


def local_data_file(filename: str) -> str:
    """Файл данных: рядом с приложением (совместимость), иначе в APPDATA."""
    if getattr(sys, "frozen", False):
        app_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
    local = os.path.join(app_dir, filename)
    if os.path.exists(local):
        return local
    return os.path.join(app_data_dir(), filename)


def setup_logging(name: str, level: int = logging.INFO) -> logging.Logger:
    """Настраивает файловый лог с ротацией (один раз на процесс)."""
    logger = logging.getLogger()
    if not getattr(logger, "_schedule_configured", False):
        os.makedirs(os.path.join(app_data_dir(), "logs"), exist_ok=True)
        log_path = os.path.join(app_data_dir(), "logs", f"{name}.log")
        handler = RotatingFileHandler(
            log_path, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
        )
        handler.setFormatter(logging.Formatter(
            "%(asctime)s %(levelname)s [%(threadName)s] %(name)s: %(message)s"
        ))
        logger.setLevel(level)
        logger.addHandler(handler)
        logger._schedule_configured = True  # type: ignore[attr-defined]
    return logging.getLogger(name)


# ---------------------------------------------------------------------------
# Протокол: сообщение = 4 байта длины + UTF-8 JSON
# ---------------------------------------------------------------------------

MAX_MESSAGE_SIZE = 16 * 1024 * 1024
_HEADER = struct.Struct("!I")


class ProtocolError(Exception):
    """Ошибка формата/размера сообщения."""


def send_message(sock: socket.socket, payload: dict[str, Any]) -> None:
    """Отправляет одно сообщение целиком (sendall, не send)."""
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    if len(data) > MAX_MESSAGE_SIZE:
        raise ProtocolError(f"Сообщение слишком большое: {len(data)} байт")
    sock.sendall(_HEADER.pack(len(data)) + data)


def _recv_exact(sock: socket.socket, size: int) -> bytes | None:
    chunks: list[bytes] = []
    remaining = size
    while remaining > 0:
        chunk = sock.recv(min(remaining, 65536))
        if not chunk:
            return None
        chunks.append(chunk)
        remaining -= len(chunk)
    return b"".join(chunks)


def recv_message(sock: socket.socket) -> dict[str, Any] | None:
    """Читает одно сообщение целиком. None — если соединение закрыто."""
    header = _recv_exact(sock, _HEADER.size)
    if header is None:
        return None
    (size,) = _HEADER.unpack(header)
    if size <= 0 or size > MAX_MESSAGE_SIZE:
        raise ProtocolError(f"Недопустимый размер сообщения: {size}")
    body = _recv_exact(sock, size)
    if body is None:
        raise ProtocolError("Соединение закрыто до конца сообщения")
    try:
        message = json.loads(body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"Некорректный JSON: {exc}") from exc
    if not isinstance(message, dict):
        raise ProtocolError("Ожидался JSON-объект")
    return message


# ---------------------------------------------------------------------------
# Миграция названий периодов
# ---------------------------------------------------------------------------


def migrate_period_names(conn: sqlite3.Connection) -> None:
    """Приводит старые числовые периоды («2026-01») к «Январь 2026».

    Если оба варианта уже существуют, данные объединяются, дубли удаляются.
    """
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id, period FROM periods")
        rows = cursor.fetchall()
    except sqlite3.OperationalError:
        return

    for period_id, name in rows:
        canonical = normalize_period_name(name)
        if not canonical or canonical == name:
            continue

        cursor.execute("SELECT id FROM periods WHERE period = ?", (canonical,))
        target = cursor.fetchone()
        if target:
            target_id = target[0]
            cursor.execute(
                """
                DELETE FROM schedules
                WHERE period_id = ?
                  AND EXISTS (
                      SELECT 1 FROM schedules AS t
                      WHERE t.period_id = ?
                        AND t.employee_name = schedules.employee_name
                        AND t.day_index = schedules.day_index
                  )
                """,
                (period_id, target_id),
            )
            cursor.execute(
                "UPDATE schedules SET period_id = ? WHERE period_id = ?",
                (target_id, period_id),
            )
            try:
                cursor.execute(
                    """
                    DELETE FROM employee_order
                    WHERE period = ?
                      AND employee_name IN (
                          SELECT employee_name FROM employee_order WHERE period = ?
                      )
                    """,
                    (name, canonical),
                )
                cursor.execute(
                    "UPDATE employee_order SET period = ? WHERE period = ?",
                    (canonical, name),
                )
            except sqlite3.OperationalError:
                pass
            cursor.execute("DELETE FROM periods WHERE id = ?", (period_id,))
        else:
            cursor.execute("UPDATE periods SET period = ? WHERE id = ?", (canonical, period_id))
            try:
                cursor.execute(
                    "UPDATE employee_order SET period = ? WHERE period = ?",
                    (canonical, name),
                )
            except sqlite3.OperationalError:
                pass

    conn.commit()
