#!/usr/bin/env python3
"""
파일 번역기 통합 실행 파일
"""
import threading
import time
import webbrowser
import os
import sys
from translate_server import app
import http.server
import socketserver

def run_translation_server():
    """번역 서버 실행"""
    app.run(host='0.0.0.0', port=5001, debug=False, use_reloader=False)

def run_web_server():
    """웹 서버 실행"""
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    Handler = http.server.SimpleHTTPRequestHandler
    with socketserver.TCPServer(("", 8000), Handler) as httpd:
        httpd.serve_forever()

if __name__ == '__main__':
    print('=' * 50)
    print('   파일 번역기 v1.0')
    print('   한국어 ↔ 중국어 번역 도구')
    print('=' * 50)
    print('서버 시작 중...')
    
    # 번역 서버 시작
    t1 = threading.Thread(target=run_translation_server)
    t1.daemon = True
    t1.start()
    
    # 웹 서버 시작
    t2 = threading.Thread(target=run_web_server)
    t2.daemon = True
    t2.start()
    
    # 잠시 대기
    time.sleep(3)
    
    # 브라우저 열기
    print('\n브라우저를 여는 중...')
    webbrowser.open('http://localhost:8000')
    
    print('\n' + '=' * 50)
    print('서비스가 실행 중입니다!')
    print('주소: http://localhost:8000')
    print('종료하려면 Ctrl+C를 누르세요')
    print('=' * 50 + '\n')
    
    # 프로그램 유지
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print('\n프로그램을 종료합니다...')
        sys.exit(0)