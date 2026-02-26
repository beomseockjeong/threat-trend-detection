@echo off
:: 위협동향 탐지 서비스 실행 스크립트 (Windows)
:: data.xlsx를 이 파일과 같은 폴더에 두고 실행하세요.

cd /d "%~dp0"

echo 서버 시작 중... (http://localhost:8080)
start /b python -m http.server 8080

timeout /t 1 /nobreak >nul
start http://localhost:8080

echo.
echo 서버가 실행 중입니다. 이 창을 닫으면 서버가 종료됩니다.
pause
