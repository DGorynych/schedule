"""Интеграционный тест клиента и сервера на localhost.

Проверяет реальный TCP-протокол, работу с SQLite и основные операции
с периодами, сотрудниками и расписанием. Использует временные базы данных,
поэтому рабочие данные не затрагиваются.
"""

import os
import sqlite3
import tempfile
import threading
import time
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt5.QtWidgets import QApplication

import schedule_client as sc
import schedule_server as ss


class ServerIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        cls.tmpdir = tempfile.mkdtemp(prefix="schedule_it_")
        cls.local_db = os.path.join(cls.tmpdir, "local_schedule.db")
        cls.server_db = os.path.join(cls.tmpdir, "schedule_server.db")

        cls.original_local_data_file = sc.local_data_file
        sc.local_data_file = lambda filename: cls.local_db

        cls.server = ss.ScheduleServer(
            host="127.0.0.1", port=0, db_path=cls.server_db
        )
        cls.server_thread = threading.Thread(
            target=cls.server.start_server, name="test-schedule-server", daemon=True
        )
        cls.server_thread.start()

        cls.port = None
        for _ in range(200):
            server_socket = cls.server.server_socket
            if server_socket is not None:
                try:
                    cls.port = server_socket.getsockname()[1]
                except OSError:
                    cls.port = None
            if cls.port and cls.server.running:
                break
            time.sleep(0.01)
        if not cls.port:
            raise RuntimeError("Сервер не запустился на localhost")

        cls.network = sc.NetworkManager()
        connected = False
        for _ in range(100):
            if cls.network.connect("127.0.0.1", cls.port):
                connected = True
                break
            time.sleep(0.02)
        if not connected:
            raise RuntimeError("Клиент не подключился к тестовому серверу")

        cls.manager = sc.ScheduleManager(cls.network)

    @classmethod
    def tearDownClass(cls):
        try:
            if cls.network is not None:
                cls.network.disconnect()
        except Exception:
            pass
        try:
            if cls.server is not None:
                cls.server.stop_server()
        except Exception:
            pass
        if cls.server_thread is not None:
            cls.server_thread.join(timeout=3)
            if cls.server_thread.is_alive():
                raise RuntimeError("Поток тестового сервера не остановился")

        sc.local_data_file = cls.original_local_data_file

        for suffix in ("", "-wal", "-shm"):
            for base in (cls.local_db, cls.server_db):
                try:
                    os.remove(base + suffix)
                except OSError:
                    pass
        try:
            os.rmdir(cls.tmpdir)
        except OSError:
            pass

    def test_client_server_roundtrip(self):
        self.assertEqual(
            self.network.send_request({"action": "ping"}),
            {"status": "success", "message": "pong"},
        )

        self.assertTrue(self.manager.add_employee("Иванов"))
        self.assertTrue(self.manager.add_employee("Петров"))
        self.assertTrue(self.manager.add_employee("Сидоров"))
        employees = self.manager.load_employees()
        self.assertEqual(
            {employee["name"] for employee in employees},
            {"Иванов", "Петров", "Сидоров"},
        )

        self.assertTrue(self.manager.create_period("Апрель 2026", ["Иванов", "Петров"]))
        data = {
            "employees": [{"name": "Иванов"}, {"name": "Петров"}],
            "schedule": {
                "Иванов": [0, 1, 2, 3, 4] + [4] * 25,
                "Петров": [4] * 30,
            },
            "notes": {
                "Иванов": {
                    "0": {"end_time": "17:30", "worked_hours": 9.5, "note_type": "о"},
                    "1": {"end_time": "20:00", "worked_hours": 12.5, "note_type": ""},
                }
            },
        }
        self.assertTrue(self.manager.save_schedule("Апрель 2026", data))

        periods = self.network.send_request({"action": "get_periods"})
        self.assertEqual(periods, {"status": "success", "data": ["Апрель 2026"]})

        response = self.network.send_request(
            {"action": "get_schedule", "period": "Апрель 2026"}
        )
        self.assertEqual(response["status"], "success")
        payload = response["data"]
        self.assertEqual(
            [employee["name"] for employee in payload["employees"]],
            ["Иванов", "Петров"],
        )
        self.assertEqual(payload["schedule"]["Иванов"][:5], [0, 1, 2, 3, 4])
        # Клиент пересчитывает часы от стандартного начала смены 07:30.
        self.assertEqual(payload["notes"]["Иванов"]["0"]["worked_hours"], 10.0)
        self.assertEqual(payload["notes"]["Иванов"]["1"]["worked_hours"], 12.5)

        self.assertTrue(
            self.manager.update_employee_order("Апрель 2026", ["Петров", "Иванов"])
        )
        ordered = self.network.send_request(
            {"action": "get_schedule", "period": "Апрель 2026"}
        )
        self.assertEqual(
            [employee["name"] for employee in ordered["data"]["employees"]],
            ["Петров", "Иванов"],
        )

        ok, message = self.manager.update_employee("Петров", "Петров П.", "Инженер")
        self.assertTrue(ok, message)
        renamed = self.network.send_request(
            {"action": "get_schedule", "period": "Апрель 2026"}
        )
        self.assertEqual(
            [employee["name"] for employee in renamed["data"]["employees"]],
            ["Петров П.", "Иванов"],
        )
        self.assertEqual(renamed["data"]["employees"][0]["position"], "Инженер")
        self.assertEqual(renamed["data"]["schedule"]["Петров П."], [4] * 30)

        self.assertTrue(self.manager.add_employees_to_period("Апрель 2026", ["Сидоров"]))
        added = self.network.send_request(
            {"action": "get_schedule", "period": "Апрель 2026"}
        )
        self.assertEqual(
            [employee["name"] for employee in added["data"]["employees"]],
            ["Петров П.", "Иванов", "Сидоров"],
        )
        self.assertEqual(added["data"]["schedule"]["Сидоров"], [4] * 30)

        rename = self.network.send_request(
            {
                "action": "rename_period",
                "old_period": "Апрель 2026",
                "period": "Май 2026",
            }
        )
        self.assertEqual(rename["status"], "success")
        self.assertEqual(
            self.network.send_request({"action": "get_periods"}),
            {"status": "success", "data": ["Май 2026"]},
        )

        self.manager._cache.clear()
        self.manager._periods_cache = None
        renamed_schedule = self.manager.load_schedule("Май 2026")
        self.assertIsNotNone(renamed_schedule)
        self.assertEqual(
            [employee["name"] for employee in renamed_schedule["employees"]],
            ["Петров П.", "Иванов", "Сидоров"],
        )

        delete = self.network.send_request(
            {"action": "delete_period", "period": "Май 2026"}
        )
        self.assertEqual(delete["status"], "success")
        self.assertEqual(
            self.network.send_request({"action": "get_periods"}),
            {"status": "success", "data": []},
        )

        self.network.disconnect()
        self.server.stop_server()
        self.server_thread.join(timeout=3)
        self.assertFalse(self.server_thread.is_alive())

        with sqlite3.connect(self.server_db) as conn:
            tables = {
                row[0]
                for row in conn.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                ).fetchall()
            }
            self.assertTrue(
                {"employees", "periods", "schedules", "employee_order"} <= tables
            )
            self.assertEqual(
                conn.execute("SELECT COUNT(*) FROM periods").fetchone()[0], 0
            )


if __name__ == "__main__":
    unittest.main()
