#!/bin/bash
# 위협동향 탐지 서비스 실행 스크립트 (macOS)
# data.xlsx를 이 스크립트와 같은 폴더에 두고 실행하세요.

cd "$(dirname "$0")"

# 포트 8080 사용 중이면 종료
lsof -ti:8080 | xargs kill -9 2>/dev/null

echo "서버 시작 중... (http://localhost:8080)"
python3 -m http.server 8080 &
SERVER_PID=$!

sleep 1
open http://localhost:8080

echo "종료하려면 이 창에서 Ctrl+C를 누르세요."
wait $SERVER_PID
