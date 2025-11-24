@echo off
chcp 65001 >nul
echo Запуск планировщика обновления остатков...
echo Запуск каждый день в 19:20
echo.
python scheduler.py
pause

