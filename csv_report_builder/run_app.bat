@echo off
rem Веб-приложение: откроется браузер. Окно не закрывайте, пока работаете.
chcp 65001 >nul
where py >nul 2>nul && (py -3 "%~dp0report_app.py") || (python "%~dp0report_app.py")
pause
