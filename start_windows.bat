@echo off
echo ============================================
echo    파일 번역기 - File Translator v1.0
echo    한국어 ↔ 중국어 번역 도구
echo ============================================
echo.

:: Python 설치 확인
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo Python 3.8 이상을 설치해주세요: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 의존성 설치 확인
echo [1/3] 필요한 패키지 확인 중...
pip show flask >nul 2>&1
if %errorlevel% neq 0 (
    echo 필요한 패키지를 설치합니다...
    pip install -r requirements.txt
)

:: 번역 서버 시작
echo [2/3] 번역 서버 시작 중...
start /B python translate_server.py

:: 3초 대기
timeout /t 3 /nobreak >nul

:: 웹 서버 시작
echo [3/3] 웹 서버 시작 중...
start /B python -m http.server 8000

:: 2초 대기
timeout /t 2 /nobreak >nul

:: 브라우저 열기
echo.
echo ============================================
echo    서비스가 시작되었습니다!
echo    브라우저를 여는 중...
echo ============================================
echo.
echo    주소: http://localhost:8000
echo.
echo    종료하려면 이 창을 닫으세요.
echo ============================================
echo.

start http://localhost:8000

:: 서버 유지
:loop
timeout /t 60 /nobreak >nul
goto loop