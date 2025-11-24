@echo off
chcp 65001 >nul
echo Установка задачи в планировщике Windows...
echo.

REM Получаем путь к текущей директории
set SCRIPT_DIR=%~dp0
set PYTHON_SCRIPT=%SCRIPT_DIR%update_ostanki.py
set SCHEDULER_SCRIPT=%SCRIPT_DIR%scheduler.py

REM Создаем задачу в планировщике Windows
schtasks /create /tn "Обновление остатков P7" /tr "python \"%SCHEDULER_SCRIPT%\"" /sc daily /st 19:20 /ru SYSTEM /f

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Задача успешно создана в планировщике Windows!
    echo Задача будет запускаться каждый день в 19:20
    echo.
    echo Для просмотра задачи: schtasks /query /tn "Обновление остатков P7"
    echo Для удаления задачи: schtasks /delete /tn "Обновление остатков P7" /f
) else (
    echo.
    echo Ошибка при создании задачи. Возможно, нужны права администратора.
    echo Запустите этот файл от имени администратора.
)

pause

