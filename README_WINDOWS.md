# 파일 번역기 - Windows 설치 가이드

## 📋 시스템 요구사항
- Windows 10 이상
- Python 3.8 이상 (설치 필요)
- 인터넷 연결 (번역 API 사용)

## 🚀 빠른 시작

### 1단계: Python 설치 (이미 설치되어 있으면 건너뛰기)
1. [Python 공식 사이트](https://www.python.org/downloads/)에서 Python 다운로드
2. 설치 시 **"Add Python to PATH"** 체크박스 반드시 선택
3. 설치 완료 후 명령 프롬프트에서 `python --version` 입력하여 확인

### 2단계: 프로그램 실행
1. FileTranslator 폴더로 이동
2. **`start_windows.bat`** 파일 더블클릭
3. 자동으로 브라우저가 열리며 프로그램이 시작됩니다

## 📁 폴더 구조
```
FileTranslator/
├── start_windows.bat        # 실행 파일
├── index.html              # 웹 인터페이스
├── style.css               # 스타일시트
├── script.js               # JavaScript
├── translate-server.py     # 번역 서버
├── excel_translator_template.py  # Excel 번역 모듈
├── requirements.txt        # Python 패키지 목록
└── icons/                  # 아이콘 폴더
    ├── excel_new.png
    ├── word_new.png
    └── pdf_new.png
```

## 💡 사용 방법

### Excel 파일 번역
1. 프로그램 실행 후 **Excel** 버튼 클릭
2. 번역 방향 선택 (한국어→중국어 또는 중국어→한국어)
3. Excel 파일 업로드
4. **번역 제외 기능**: 미리보기에서 번역하지 않을 셀 클릭 (빨간색 표시)
5. "번역 시작" 버튼 클릭
6. 번역 완료 후 다운로드

### 주요 기능
- ✅ Excel 서식 100% 보존
- ✅ 영문 내용 자동 유지
- ✅ 새 시트에 번역본 추가 옵션
- ✅ 특정 셀 번역 제외
- ✅ 실시간 진행률 표시

## 🔧 문제 해결

### "Python이 설치되어 있지 않습니다" 오류
- Python 3.8 이상 설치 필요
- 설치 시 PATH 추가 옵션 확인

### "패키지를 찾을 수 없습니다" 오류
명령 프롬프트에서 다음 실행:
```bash
pip install -r requirements.txt
```

### 포트 충돌 오류
다른 프로그램이 8000, 5001 포트를 사용 중일 수 있습니다.
작업 관리자에서 Python 프로세스를 종료하고 재시작하세요.

## 📞 지원
- 이메일: hgyjg@teclast.co.kr
- 회사: (주)테클라스트코리아

## 🔄 업데이트 내역
- v1.0 (2025.09): 초기 버전 출시
  - Excel 번역 기능
  - 셀 제외 기능
  - 한국어 ↔ 중국어 번역

---
© 2025 Teclast Korea. All rights reserved.