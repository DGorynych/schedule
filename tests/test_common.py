"""Тесты общих функций, протокола и миграции периодов.

Запуск: python -m unittest discover -s tests -v
"""

import os
import socket
import sqlite3
import tempfile
import unittest

from schedule_common import (
    ProtocolError,
    compute_worked_hours,
    days_in_month,
    format_note_text,
    format_period,
    has_meaningful_note,
    migrate_period_names,
    normalize_period_name,
    parse_period,
    period_filename,
    recv_message,
    sanitize_filename,
    send_message,
    sort_periods,
    sunday_flags,
)


class PeriodTests(unittest.TestCase):
    def test_parse_display_and_numeric(self):
        self.assertEqual(parse_period("Январь 2026"), (2026, 1, ""))
        self.assertEqual(parse_period("2026-01"), (2026, 1, ""))
        self.assertEqual(
            parse_period("Март 2026 (Вариант 2)"), (2026, 3, "Вариант 2")
        )
        self.assertIsNone(parse_period("не период"))

    def test_normalize_and_format(self):
        self.assertEqual(normalize_period_name("2026-01"), "Январь 2026")
        self.assertEqual(format_period(2024, 2), "Февраль 2024")
        self.assertEqual(
            format_period(2024, 2, "Вариант 2"), "Февраль 2024 (Вариант 2)"
        )

    def test_calendar(self):
        self.assertEqual(days_in_month("Февраль 2024"), 29)
        self.assertEqual(days_in_month("Февраль 2025"), 28)
        self.assertTrue(sunday_flags("Март 2026")[0])
        self.assertEqual(
            sort_periods(["Февраль 2026", "Январь 2026", "Декабрь 2025"]),
            ["Декабрь 2025", "Январь 2026", "Февраль 2026"],
        )

    def test_filenames(self):
        self.assertEqual(sanitize_filename('a/b:c*?'), "a_b_c__")
        self.assertTrue(period_filename("Январь 2026").endswith(".xlsx"))
        self.assertNotIn("/", period_filename("Январь 2026"))


class NoteTests(unittest.TestCase):
    def test_worked_hours(self):
        self.assertEqual(compute_worked_hours(0, "20:00"), 12.5)
        self.assertEqual(compute_worked_hours(0, "17:30"), 10.0)
        self.assertEqual(compute_worked_hours(0, "07:30"), 0.0)
        self.assertEqual(compute_worked_hours(0, "07:00"), 0.0)
        self.assertEqual(compute_worked_hours(0, None), 12.5)
        self.assertEqual(compute_worked_hours(2, "20:00"), 0.0)

    def test_note_helpers(self):
        self.assertEqual(format_note_text("15:00", "о"), "15-о")
        self.assertEqual(format_note_text("20:00", "д"), "д")
        self.assertEqual(format_note_text("20:00", ""), "")
        self.assertFalse(has_meaningful_note({"end_time": "20:00"}))
        self.assertTrue(has_meaningful_note({"note_type": "д"}))


class ProtocolTests(unittest.TestCase):
    def test_roundtrip(self):
        left, right = socket.socketpair()
        try:
            payload = {"action": "ping", "text": "Кириллица", "number": 42}
            send_message(left, payload)
            self.assertEqual(recv_message(right), payload)
        finally:
            left.close()
            right.close()

    def test_rejects_invalid_frame(self):
        left, right = socket.socketpair()
        try:
            left.sendall(b"\x00\x00\x00\x00")
            with self.assertRaises(ProtocolError):
                recv_message(right)
        finally:
            left.close()
            right.close()


class MigrationTests(unittest.TestCase):
    def test_numeric_periods_are_merged(self):
        fd, path = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        try:
            conn = sqlite3.connect(path)
            conn.execute(
                "CREATE TABLE periods (id INTEGER PRIMARY KEY, period TEXT UNIQUE)"
            )
            conn.execute(
                """
                CREATE TABLE schedules (
                    id INTEGER PRIMARY KEY,
                    period_id INTEGER,
                    employee_name TEXT,
                    day_index INTEGER,
                    status INTEGER,
                    end_time TEXT,
                    worked_hours REAL,
                    note_type TEXT,
                    UNIQUE(period_id, employee_name, day_index)
                )
                """
            )
            conn.execute(
                "CREATE TABLE employee_order ("
                "id INTEGER PRIMARY KEY, period TEXT, "
                "employee_name TEXT, sort_order INTEGER)"
            )
            conn.execute("INSERT INTO periods (id, period) VALUES (1, '2026-01')")
            conn.execute(
                "INSERT INTO schedules (period_id, employee_name, day_index, status) "
                "VALUES (1, 'Иванов', 0, 0)"
            )
            conn.execute(
                "INSERT INTO employee_order (period, employee_name, sort_order) "
                "VALUES ('2026-01', 'Иванов', 0)"
            )
            conn.commit()

            migrate_period_names(conn)

            self.assertEqual(
                conn.execute("SELECT period FROM periods").fetchall(),
                [("Январь 2026",)],
            )
            self.assertEqual(
                conn.execute("SELECT period FROM employee_order").fetchone()[0],
                "Январь 2026",
            )
            self.assertEqual(
                conn.execute("SELECT COUNT(*) FROM schedules").fetchone()[0], 1
            )
            conn.close()
        finally:
            try:
                os.remove(path)
            except OSError:
                pass


if __name__ == "__main__":
    unittest.main()
