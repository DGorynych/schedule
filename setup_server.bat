@echo off
echo Компиляция сервера...

REM Создаем виртуальное окружение, если его нет
if not exist "venv_server" (
    echo Создание виртуального окружения...
    python -m venv venv_server
)

REM Активируем виртуальное окружение
call venv_server\Scripts\activate.bat

REM Обновляем pip
python -m pip install --upgrade pip

REM Устанавливаем зависимости
echo Установка зависимостей...
pip install pyinstaller PyQt5

REM Компилируем сервер
echo Компиляция schedule_server.py в исполняемый файл...
pyinstaller ^
    --onefile ^
    --windowed ^
    --icon=schedule.ico ^
    --name="WorkScheduleServer" ^
    --add-data="schedule.ico;." ^
    --hidden-import=sqlite3 ^
    --hidden-import=winreg ^
    --hidden-import=calendar ^
    --clean ^
    schedule_server.py

echo Готово! Исполняемый файл находится в папке dist\WorkScheduleServer.exe
pause