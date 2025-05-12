@echo off
chcp 65001 > nul
echo.
echo =====================================================
echo           ОНОВЛЕННЯ ДОДАТКУ - НЕ ЗАКРИВАЙТЕ ЦЕ ВІКНО!
echo =====================================================
echo.
echo Закриття попередньої версії додатку...
taskkill /IM "python.exe" /F > nul 2>&1
timeout /t 3 /nobreak > nul

echo.
echo Копіювання нових файлів...
if exist "E:\2. Басай\00. Заходи 2025\main prog\2.8.1" ( rmdir /S /Q "E:\2. Басай\00. Заходи 2025\main prog\2.8.1" )
xcopy "update_extracted_temp\sportfprall-main\2.8.1" "E:\2. Басай\00. Заходи 2025\main prog\2.8.1\" /E /I /Y /Q
if exist "E:\2. Басай\00. Заходи 2025\main prog\version 2.8" ( rmdir /S /Q "E:\2. Басай\00. Заходи 2025\main prog\version 2.8" )
xcopy "update_extracted_temp\sportfprall-main\version 2.8" "E:\2. Басай\00. Заходи 2025\main prog\version 2.8\" /E /I /Y /Q
if exist "E:\2. Басай\00. Заходи 2025\main prog\version.txt" ( del /F /Q "E:\2. Басай\00. Заходи 2025\main prog\version.txt" )
move /Y "update_extracted_temp\sportfprall-main\version.txt" "E:\2. Басай\00. Заходи 2025\main prog\version.txt"

echo.
echo Очищення тимчасових файлів...
rmdir /S /Q "update_extracted_temp" > nul 2>&1

echo.
echo Оновлення завершено! Запуск нової версії...
start "" "C:\Program Files\Python313\python.exe" "E:\2. Басай\00. Заходи 2025\main prog\sport.py"

echo.
echo Скрипт оновлення завершить роботу за 5 секунд...
timeout /t 5 /nobreak > nul
del "%~f0"
exit
