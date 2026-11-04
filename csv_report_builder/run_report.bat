@echo off
rem Запуск: перетащите папку с CSV на этот файл или запустите и введите путь.
chcp 65001 >nul
set "SRC=%~1"
if "%SRC%"=="" set /p "SRC=CSV folder: "
where py >nul 2>nul && (py -3 "%~dp0report_builder.py" "%SRC%") || (python "%~dp0report_builder.py" "%SRC%")
echo.
pause
