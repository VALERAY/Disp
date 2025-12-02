@echo off
echo ========================================
echo Сборка приложения Disp с инсталлером
echo ========================================

echo.
echo Шаг 1: Очистка предыдущих сборок...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "installer" rmdir /s /q "installer"

echo.
echo Шаг 2: Сборка исполняемого файла с PyInstaller...
pyinstaller --clean --noconfirm .\Disp.spec

if %ERRORLEVEL% neq 0 (
    echo ОШИБКА: Не удалось собрать исполняемый файл!
    pause
    exit /b 1
)

echo.
echo Шаг 3: Создание папки для инсталлера...
if not exist "installer" mkdir "installer"

echo.
echo Шаг 4: Сборка инсталлера с Inno Setup...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "Disp_Setup.iss"

if %ERRORLEVEL% neq 0 (
    echo ОШИБКА: Не удалось создать инсталлер!
    echo Убедитесь, что Inno Setup установлен по пути:
    echo "C:\Program Files (x86)\Inno Setup 6\"
    pause
    exit /b 1
)

echo.
echo ========================================
echo СБОРКА ЗАВЕРШЕНА УСПЕШНО!
echo ========================================
echo.
echo Исполняемый файл: dist\Disp.exe
echo Инсталлер: installer\Disp_Setup.exe
echo.
pause

