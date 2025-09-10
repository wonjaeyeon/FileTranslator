@echo off
echo ============================================
echo    FileTranslator.exe 빌드 스크립트
echo ============================================
echo.

:: PyInstaller 설치 확인
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller 설치 중...
    pip install pyinstaller
)

:: EXE 빌드 (모든 파일 포함)
echo EXE 파일 생성 중... (3-5분 소요)
echo.

pyinstaller --onefile ^
    --name "FileTranslator" ^
    --windowed ^
    --add-data "index.html;." ^
    --add-data "style.css;." ^
    --add-data "script.js;." ^
    --add-data "excel_translator_template.py;." ^
    --add-data "translate_server.py;." ^
    --add-data "icons;icons" ^
    --hidden-import=flask ^
    --hidden-import=flask_cors ^
    --hidden-import=openpyxl ^
    --hidden-import=googletrans ^
    --hidden-import=deep_translator ^
    --hidden-import=requests ^
    app.py

echo.
echo ============================================
echo    빌드 완료!
echo.
echo    생성된 파일:
echo    dist\FileTranslator.exe
echo.
echo    이 EXE 파일만 배포하면 됩니다!
echo ============================================
echo.
pause