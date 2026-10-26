"""Безопасный запуск клиента для просмотра нового оформления.

Использует отдельную копию local_schedule.db рядом с этим файлом и не
подключается к серверу, чтобы изменения из предпросмотра не попали в рабочие
данные.
"""

import os
import sys

from PyQt5.QtCore import QSettings as QtSettings
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication

import schedule_client as sc
from ui_theme import apply_theme, set_pill_state


APP_DIR = os.path.dirname(os.path.abspath(__file__))
PREVIEW_DB = os.path.join(APP_DIR, "local_schedule.db")


class PreviewSettings(QtSettings):
    """Отдельные настройки, чтобы предпросмотр не менял настройки клиента."""

    def __init__(self, organization, application):
        super().__init__(organization, f"{application}-Preview")


def main():
    sc.QSettings = PreviewSettings
    sc.local_data_file = lambda filename: PREVIEW_DB
    sc.NetworkManager.connect = lambda self, host, port: False
    sc.ScheduleApp.auto_connect = lambda self: None
    sc.ScheduleApp.try_reconnect = lambda self: None

    sc.setup_logging("schedule_preview")
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    apply_theme(app)

    window = sc.ScheduleApp()
    window.setWindowTitle(window.windowTitle() + " — предпросмотр")
    window.connection_status.setText("Локальный режим")
    set_pill_state(window.connection_status, "local")
    window.showMaximized()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
