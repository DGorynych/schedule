@echo off
chcp 65001
echo Компиляция серверной части...
pyinstaller --onefile --windowed --icon=schedule.ico --name="ScheduleServer" schedule_server.py
echo Готово!
pause