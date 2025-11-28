@echo off
chcp 65001
echo Компиляция клиентской части...
pyinstaller --onefile --windowed --icon=schedule.ico --name="ScheduleClient" schedule_client.py
echo Готово!
pause