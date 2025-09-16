@echo off
echo ============================================
echo    FileTranslator.exe 빌드 스크립트
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

:: PyInstaller 설치 확인
echo [1/4] PyInstaller 확인 중...
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller 설치 중...
    pip install pyinstaller
)

:: 의존성 설치 확인
echo [2/4] 필요한 패키지 확인 중...
pip show flask >nul 2>&1
if %errorlevel% neq 0 (
    echo 필요한 패키지를 설치합니다...
    pip install -r requirements.txt
)

:: 기존 빌드 정리
echo [3/4] 기존 빌드 파일 정리 중...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del *.spec

:: EXE 빌드 (모든 파일 포함)
echo [4/4] EXE 파일 생성 중... (3-5분 소요)
echo.

pyinstaller --onefile ^
    --name "FileTranslator" ^
    --windowed ^
    --add-data "index.html;." ^
    --add-data "style.css;." ^
    --add-data "script.js;." ^
    --add-data "excel_translator_template.py;." ^
    --add-data "translate_server.py;." ^
    --add-data "requirements.txt;." ^
    --hidden-import=flask ^
    --hidden-import=flask_cors ^
    --hidden-import=openpyxl ^
    --hidden-import=googletrans ^
    --hidden-import=deep_translator ^
    --hidden-import=requests ^
    --hidden-import=werkzeug ^
    --hidden-import=jinja2 ^
    --hidden-import=markupsafe ^
    --hidden-import=itsdangerous ^
    --hidden-import=click ^
    --hidden-import=blinker ^
    app.py

echo.
if exist dist\FileTranslator.exe (
    echo ============================================
    echo    빌드 완료!
    echo.
    echo    생성된 파일:
    echo    dist\FileTranslator.exe
    echo.
    echo    이 EXE 파일만 배포하면 됩니다!
    echo    파일 크기: 
    for %%F in (dist\FileTranslator.exe) do echo    %%~zF bytes
    echo ============================================
) else (
    echo ============================================
    echo    빌드 실패!
    echo    오류를 확인하고 다시 시도해주세요.
    echo ============================================
)
echo.
pause