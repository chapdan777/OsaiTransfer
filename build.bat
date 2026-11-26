@echo off
chcp 65001 >nul
title Сборка OsaiTransfer.exe

echo ========================================================
echo   Сборка переносимого приложения OsaiTransfer (ЧПУ OSAI)
echo ========================================================
echo.

REM Проверка наличия Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не найден в системе. Установите Python 3.10+ с галочкой 'Add Python to PATH'.
    pause
    exit /b 1
)

echo [1/4] Создание изолированного виртуального окружения...
if not exist "venv" (
    python -m venv venv
)
call venv\Scripts\activate.bat

echo [2/4] Установка необходимых библиотек (pysmb, pyinstaller)...
python -m pip install --upgrade pip
pip install -r requirements.txt

echo [3/4] Сборка автономного исполняемого файла OsaiTransfer.exe...
pyinstaller --noconfirm --onefile --windowed ^
    --name "OsaiTransfer" ^
    --hidden-import "smb" ^
    --hidden-import "smb.SMBConnection" ^
    --hidden-import "smb.base" ^
    --hidden-import "win32com" ^
    --hidden-import "win32com.client" ^
    --hidden-import "win32com.shell" ^
    --hidden-import "win32com.shell.shell" ^
    --hidden-import "win32com.shell.shellcon" ^
    --hidden-import "win32gui" ^
    --hidden-import "win32gui_struct" ^
    --hidden-import "win32con" ^
    --hidden-import "pythoncom" ^
    --hidden-import "pywintypes" ^
    --collect-data "tkinterdnd2" ^
    app.py

if errorlevel 1 (
    echo.
    echo [ОШИБКА] Сборка завершилась с ошибкой.
    pause
    exit /b 1
)

echo [4/4] Подготовка портативного каталога...
if not exist "dist\config.ini" (
    copy /Y "config.ini" "dist\config.ini"
)

echo.
echo ========================================================
echo   СБОРКА УСПЕШНО ЗАВЕРШЕНА!
echo   Исполняемый файл готов: dist\OsaiTransfer.exe
echo   Рядом скопирован файл: dist\config.ini
echo ========================================================
echo.
echo Для работы на станке скопируйте папку 'dist' (или файлы OsaiTransfer.exe и config.ini)
echo на Рабочий стол оператора.
echo.
pause
