@echo off
cd /d "%~dp0"

python --version >nul 2>&1
if %errorlevel% == 0 (
    start "Threat Detection Server" python -m http.server 8080
) else (
    py --version >nul 2>&1
    if %errorlevel% == 0 (
        start "Threat Detection Server" py -m http.server 8080
    ) else (
        echo Python not found. Please install Python.
        pause
        exit /b 1
    )
)

timeout /t 2 /nobreak >nul
start http://localhost:8080
pause
