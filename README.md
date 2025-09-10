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

### 개발자 실행 (v2.0 리팩토링 버전)

```bash
# 의존성 설치
pip install -r requirements.txt

# v2.0 실행 (권장 - 리팩토링된 버전)
python main.py

# 또는 기존 버전 실행
python app.py
```

### v2.0 주요 개선사항

- ✅ **모듈화된 아키텍처** - 체계적인 코드 구조
- ✅ **외부 설정 파일** - 하드코딩된 사전을 JSON 파일로 분리
- ✅ **개선된 번역 엔진** - 사전 우선, Google Translate 보조
- ✅ **더 나은 코드 품질** - 클린 코드 원칙 적용

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
├── main.py                     # v2.0 통합 실행 파일 (권장)
├── app.py                      # v1.0 통합 실행 파일
├── src/                        # 소스 코드 모듈
│   ├── core/                   # 핵심 번역 엔진
│   │   ├── dictionary_manager.py  # 사전 관리자
│   │   ├── translator.py       # 번역 엔진
│   │   └── excel_translator.py # Excel 번역기
│   ├── api/                    # API 서버
│   │   └── server.py           # Flask 서버
│   └── web/                    # 웹 인터페이스
│       ├── index.html          # 메인 페이지
│       ├── style.css           # 스타일시트
│       ├── script.js           # 프론트엔드 로직
│       └── icons/              # 아이콘 파일들
├── config/                     # 설정 파일
│   ├── translation_dict.json   # 기본 번역 사전
│   └── custom_dict.json        # 회사별 커스텀 사전
├── requirements.txt            # Python 패키지
├── start_windows.bat           # Windows 실행 스크립트
└── build_exe_simple.bat        # EXE 빌드 스크립트
```

### v2.0 아키텍처 특징

- **분리된 관심사**: 번역 로직, API, 웹 UI가 각각 분리
- **설정 기반**: 하드코딩된 사전을 외부 JSON 파일로 관리
- **모듈화**: 재사용 가능한 컴포넌트 구조
- **확장성**: 새로운 번역 엔진이나 파일 형식 추가 용이

## 📝 라이선스

본 소프트웨어의 사용을 위해서는 반드시 개발자 원재연의 사전 허락이 필요합니다.
무단 사용, 복제, 배포, 수정은 엄격히 금지됩니다.

## 📞 문의

- 이메일: woncow977@gmail.com
- 개발자: 원재연

---
© 2025 원재연(Won Jae Yeon). All rights reserved.