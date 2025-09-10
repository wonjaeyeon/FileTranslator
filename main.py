#!/usr/bin/env python3
"""
파일 번역기 v2.0 - 리팩토링된 통합 실행 파일
"""
import threading
import time
import webbrowser
import os
import sys
import http.server
import socketserver

# 프로젝트 루트를 Python 경로에 추가
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

def run_translation_server():
    """번역 API 서버 실행"""
    from src.api.server import app
    app.run(host='0.0.0.0', port=5001, debug=False, use_reloader=False)

def run_web_server():
    """웹 서버 실행 (정적 파일 서빙)"""
    web_dir = os.path.join(project_root, 'src', 'web')
    os.chdir(web_dir)
    
    Handler = http.server.SimpleHTTPRequestHandler
    with socketserver.TCPServer(("", 9000), Handler) as httpd:
        httpd.serve_forever()

def main():
    print('=' * 60)
    print('   파일 번역기 v2.0 (리팩토링 버전)')
    print('   한국어 ↔ 중국어 번역 도구')
    print('   - 개선된 아키텍처')
    print('   - 외부 설정 파일 지원')
    print('   - 더 나은 코드 구조')
    print('=' * 60)
    print('서버 시작 중...')
    
    # 번역 API 서버 시작
    api_thread = threading.Thread(target=run_translation_server)
    api_thread.daemon = True
    api_thread.start()
    
    # 웹 서버 시작  
    web_thread = threading.Thread(target=run_web_server)
    web_thread.daemon = True
    web_thread.start()
    
    # 잠시 대기
    time.sleep(3)
    
    # 브라우저 열기
    print('\\n브라우저를 여는 중...')
    webbrowser.open('http://localhost:9000')
    
    print('\\n' + '=' * 60)
    print('서비스가 실행 중입니다!')
    print('주소: http://localhost:9000')
    print('API: http://localhost:5001')
    print('종료하려면 Ctrl+C를 누르세요')
    print('=' * 60 + '\\n')
    
    # 프로그램 유지
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print('\\n프로그램을 종료합니다...')
        sys.exit(0)

if __name__ == '__main__':
    main()