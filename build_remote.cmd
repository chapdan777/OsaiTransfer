@echo off
chcp 65001 >nul
setlocal

set PY_EXE=C:\Temp\py311\python.exe
set BUILD_DIR=C:\Temp\osai_build
set TARGET_DIR=C:\Users\OEM\Desktop\OsaiTransfer

if not exist "%PY_EXE%" (
    echo [ERROR] Python not found at %PY_EXE%
    exit /b 1
)

if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
mkdir "%BUILD_DIR%"
cd /d "%BUILD_DIR%"

echo [1/4] Copying sources to build directory...
copy /Y "%~dp0app.py" "%BUILD_DIR%\app.py" >nul
copy /Y "%~dp0smb_client.py" "%BUILD_DIR%\smb_client.py" >nul
copy /Y "%~dp0config.ini" "%BUILD_DIR%\config.ini" >nul

echo [2/4] Compiling OsaiTransfer.exe with PyInstaller...
"%PY_EXE%" -m PyInstaller --noconfirm --onefile --windowed ^
    --name OsaiTransfer ^
    --hidden-import smb ^
    --hidden-import smb.SMBConnection ^
    --hidden-import smb.base ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --hidden-import win32com.shell ^
    --hidden-import win32com.shell.shell ^
    --hidden-import win32com.shell.shellcon ^
    --hidden-import win32gui ^
    --hidden-import win32gui_struct ^
    --hidden-import win32con ^
    --hidden-import pythoncom ^
    --hidden-import pywintypes ^
    --collect-data tkinterdnd2 ^
    app.py

if not exist "%BUILD_DIR%\dist\OsaiTransfer.exe" (
    echo [ERROR] Build failed! OsaiTransfer.exe not found.
    exit /b 1
)

echo [3/4] Deploying to %TARGET_DIR%...
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"
copy /Y "%BUILD_DIR%\dist\OsaiTransfer.exe" "%TARGET_DIR%\OsaiTransfer.exe" >nul
if not exist "%TARGET_DIR%\config.ini" copy /Y "%~dp0config.ini" "%TARGET_DIR%\config.ini" >nul

echo [4/4] Cleaning up temporary build directory...
cd /d C:\Temp
rmdir /s /q "%BUILD_DIR%"

echo ========================================================
echo SUCCESS: OsaiTransfer deployed to %TARGET_DIR%
echo ========================================================
