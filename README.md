# 파일 번역기 (File Translator)

한국어 ↔ 중국어 Excel 파일 자동 번역 도구

## 🌟 주요 기능

- ✅ **Excel 서식 100% 보존** - 셀 병합, 테두리, 색상, 폰트 모두 유지
- ✅ **영문 자동 보존** - 영어 텍스트는 번역하지 않음
- ✅ **선택적 번역** - 특정 셀을 클릭하여 번역에서 제외
- ✅ **실시간 미리보기** - Excel 파일을 웹에서 바로 확인
- ✅ **진행률 표시** - 번역 진행 상황 실시간 확인

## 🚀 빠른 시작

### Windows 사용자 (간편 실행)

1. `start_windows.bat` 더블클릭
2. 자동으로 브라우저가 열림
3. Excel 파일 업로드 → 번역 → 다운로드

### 개발자 실행

```bash
# 의존성 설치
pip install -r requirements.txt

# 통합 실행 (권장)
python app.py

# 또는 개별 실행
python translate-server.py  # 터미널 1
python -m http.server 8000  # 터미널 2
```

## 💡 사용 방법

1. **파일 선택**: Excel 버튼 클릭
2. **번역 설정**: 
   - 한국어 → 중국어 또는 중국어 → 한국어 선택
   - 영문 유지 옵션 (기본 활성화)
   - 새 시트 추가 옵션 (기본 활성화)
3. **제외 설정**: 미리보기에서 번역하지 않을 셀 클릭 (빨간색 표시)
4. **번역 실행**: "번역 시작" 버튼 클릭
5. **다운로드**: 번역 완료 후 파일 다운로드

## 🛠 기술 스택

- **Frontend**: HTML5, CSS3, JavaScript (Vanilla)
- **Backend**: Python Flask
- **번역 API**: Google Translate
- **Excel 처리**: openpyxl (Python), XLSX.js (JavaScript)

## 📦 Windows EXE 빌드

```bash
# Windows에서 실행
build_exe_simple.bat
# dist/FileTranslator.exe 생성됨
```

## 📂 프로젝트 구조

```
FileTranslator/
├── app.py                      # 통합 실행 파일
├── translate_server.py         # 번역 서버
├── excel_translator_template.py # Excel 번역 모듈
├── index.html                  # 웹 인터페이스
├── style.css                   # 스타일시트
├── script.js                   # 프론트엔드 로직
├── requirements.txt            # Python 패키지
├── start_windows.bat           # Windows 실행 스크립트
├── build_exe_simple.bat        # EXE 빌드 스크립트
└── icons/                      # 아이콘 파일들
```

## 📝 라이선스

MIT License

## 📞 문의

- 이메일: hgyjg@teclast.co.kr
- 회사: (주)테클라스트코리아

---
© 2025 Teclast Korea. All rights reserved.