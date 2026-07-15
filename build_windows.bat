@echo off
setlocal
chcp 65001 >nul

set "SCRIPT_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%build_windows.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
    echo Сборка завершилась с ошибкой.
    pause
    exit /b %EXIT_CODE%
)

echo Сборка завершена. Можно закрыть это окно.
pause >nul
