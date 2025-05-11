@echo off
echo Запуск программы Генератор договорів...
python fill_contract_gui.py
if %ERRORLEVEL% NEQ 0 (
    echo Программа завершилась с ошибкой!
    echo Детали ошибки можно найти в файле error.txt
    pause
)