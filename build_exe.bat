@echo off
echo ============================================
echo    EXE 파일 생성 스크립트
echo ============================================
echo.

:: PyInstaller 설치
echo PyInstaller 설치 중...
pip install pyinstaller

:: 메인 실행 파일 생성
echo.
echo 메인 앱 파일 생성 중...

:: app.py 생성 (통합 실행 파일)
echo import threading > app.py
echo import time >> app.py
echo import webbrowser >> app.py
echo import os >> app.py
echo import sys >> app.py
echo from translate_server import app >> app.py
echo import http.server >> app.py
echo import socketserver >> app.py
echo. >> app.py
echo def run_translation_server(): >> app.py
echo     app.run(host='0.0.0.0', port=5001, debug=False) >> app.py
echo. >> app.py
echo def run_web_server(): >> app.py
echo     os.chdir(os.path.dirname(os.path.abspath(__file__))) >> app.py
echo     Handler = http.server.SimpleHTTPRequestHandler >> app.py
echo     with socketserver.TCPServer(("", 8000), Handler) as httpd: >> app.py
echo         httpd.serve_forever() >> app.py
echo. >> app.py
echo if __name__ == '__main__': >> app.py
echo     print('파일 번역기 시작 중...') >> app.py
echo     t1 = threading.Thread(target=run_translation_server) >> app.py
echo     t1.daemon = True >> app.py
echo     t1.start() >> app.py
echo     t2 = threading.Thread(target=run_web_server) >> app.py
echo     t2.daemon = True >> app.py
echo     t2.start() >> app.py
echo     time.sleep(3) >> app.py
echo     webbrowser.open('http://localhost:8000') >> app.py
echo     print('브라우저에서 http://localhost:8000 을 열어주세요') >> app.py
echo     print('종료하려면 Ctrl+C를 누르세요') >> app.py
echo     try: >> app.py
echo         while True: >> app.py
echo             time.sleep(1) >> app.py
echo     except KeyboardInterrupt: >> app.py
echo         print('프로그램을 종료합니다...') >> app.py
echo         sys.exit(0) >> app.py

:: PyInstaller로 빌드
echo.
echo EXE 파일 빌드 중... (시간이 걸릴 수 있습니다)
pyinstaller --onefile ^
    --name "FileTranslator" ^
    --icon "icons/excel_new.png" ^
    --add-data "index.html;." ^
    --add-data "style.css;." ^
    --add-data "script.js;." ^
    --add-data "excel_translator_template.py;." ^
    --add-data "icons;icons" ^
    --hidden-import flask ^
    --hidden-import flask_cors ^
    --hidden-import openpyxl ^
    --hidden-import googletrans ^
    --hidden-import deep_translator ^
    app.py

echo.
echo ============================================
echo    빌드 완료!
echo    dist/FileTranslator.exe 파일이 생성되었습니다.
echo ============================================
pause