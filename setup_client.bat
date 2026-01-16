@echo off
echo Компиляция клиента...

REM Создаем виртуальное окружение, если его нет
if not exist "venv_client" (
    echo Создание виртуального окружения...
    python -m venv venv_client
)

REM Активируем виртуальное окружение
call venv_client\Scripts\activate.bat

REM Обновляем pip
python -m pip install --upgrade pip

REM Устанавливаем зависимости
echo Установка зависимостей...
pip install pyinstaller PyQt5 openpyxl

REM Компилируем клиента
echo Компиляция schedule_client.py в исполняемый файл...
pyinstaller ^
    --onefile ^
    --windowed ^
    --icon=schedule.ico ^
    --name="WorkScheduleClient" ^
    --add-data="schedule.ico;." ^
    --hidden-import=sqlite3 ^
    --hidden-import=openpyxl ^
    --hidden-import=calendar ^
    --clean ^
    schedule_client.py

echo Готово! Исполняемый файл находится в папке dist\WorkScheduleClient.exe
pause