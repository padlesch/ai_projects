@echo off
rem Собрать ReportMaker.exe (один файл, Python на других компьютерах не нужен).
rem Нужен установленный Python 3.10+; первый раз скачает PyInstaller и библиотеки.
chcp 65001 >nul
cd /d "%~dp0"
set "PY=python"
where py >nul 2>nul && set "PY=py -3"
%PY% -m pip install -r requirements.txt pyinstaller || goto :fail
%PY% -m PyInstaller --noconfirm --onefile --name ReportMaker --console --hidden-import tkinter --hidden-import tkinter.filedialog --exclude-module pyarrow report_app.py || goto :fail
echo.
echo Готово: %~dp0dist\ReportMaker.exe
pause
exit /b 0
:fail
echo Сборка не удалась — см. сообщения выше.
pause
exit /b 1
