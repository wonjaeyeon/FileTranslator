# 파일 번역기 (File Translator)

한국어 ↔ 중국어 Excel 파일 번역 도구 with GPT 워크플로우

## 🌟 주요 기능

- ✅ **Excel 서식 100% 보존** - 셀 병합, 테두리, 색상, 폰트, 이미지, 도형 모두 유지
- ✅ **GPT 번역 워크플로우** - 3단계 GPT 기반 번역 (추출 → GPT 번역 → 적용)
- ✅ **자동 번역 지원** - Google Translate 기반 기본 번역
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
python translate_server.py  # 터미널 1
python -m http.server 8000  # 터미널 2
```

## 💡 사용 방법

### 📝 GPT 번역 워크플로우 (권장)

1. **파일 선택**: Excel 버튼 클릭하여 파일 업로드
2. **번역 방법 선택**: "GPT 번역 워크플로우" 선택
3. **번역 설정**:
   - 한국어 → 중국어 또는 중국어 → 한국어 선택
   - 영문 유지 옵션 (기본 활성화)
   - 새 시트 추가 옵션 (기본 활성화)
4. **1단계 - 단어 추출**: "단어 추출" 버튼 클릭
5. **2단계 - GPT 번역**: 생성된 프롬프트를 GPT에 입력하고 결과를 텍스트 영역에 붙여넣기
6. **3단계 - 번역 적용**: "번역 적용" 버튼 클릭
7. **다운로드**: 번역 완료 후 파일 다운로드

### 🔄 자동 번역 (기본 방식)

1. **파일 선택**: Excel 버튼 클릭하여 파일 업로드
2. **번역 방법 선택**: "자동 번역" 선택
3. **번역 설정**: 방향, 영문 유지, 새 시트 추가 옵션 설정
4. **제외 설정**: 미리보기에서 번역하지 않을 셀 클릭 (빨간색 표시)
5. **번역 실행**: "번역 시작" 버튼 클릭
6. **다운로드**: 번역 완료 후 파일 다운로드

## 🏗️ 시스템 아키텍처

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                          🌐 CLIENT LAYER (Browser)                              │
├─────────────────────────────────────────────────────────────────────────────────┤
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────────────────────┐  │
│  │   Web Interface │  │  Excel Preview  │  │       File Upload               │  │
│  │  HTML5/CSS3/JS  │  │    XLSX.js      │  │    Base64 Encoding              │  │
│  └─────────────────┘  └─────────────────┘  └─────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────────┘
                                    │ HTTP Requests
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────────┐
│                      ⚙️  APPLICATION LAYER (Flask Server)                       │
├─────────────────────────────────────────────────────────────────────────────────┤
│  ┌─────────────────┐                    ┌─────────────────────────────────────┐  │
│  │  Flask Router   │◄──────────────────►│         File Manager                │  │
│  │  CORS Handler   │                    │                                     │  │
│  └─────────────────┘                    └─────────────────────────────────────┘  │
│           │                                              │                      │
│           ▼                                              ▼                      │
│  ┌───────────────────────────────────┐  ┌─────────────────────────────────────┐  │
│  │     🔄 TRANSLATION SERVICES       │  │      📊 EXCEL PROCESSING            │  │
│  │                                   │  │                                     │  │
│  │  ┌─────────────────────────────┐  │  │  ┌─────────────────────────────────┐│  │
│  │  │    GPT Workflow Service     │◄─┼──┼─►│        Excel Parser         ││  │
│  │  │   /extract-words            │  │  │  │         openpyxl            ││  │
│  │  │   /process-gpt-translation  │  │  │  └─────────────────────────────────┘│  │
│  │  └─────────────────────────────┘  │  │  ┌─────────────────────────────────┐│  │
│  │                                   │  │  │      Cell Text Extractor     ││  │
│  │  ┌─────────────────────────────┐  │  │  └─────────────────────────────────┘│  │
│  │  │   Auto Translation Service  │◄─┼──┼─►│      Format Preserver       ││  │
│  │  │         /translate          │  │  │  │   Images/Shapes/Styles      ││  │
│  │  └─────────────────────────────┘  │  │  └─────────────────────────────────┘│  │
│  └───────────────────────────────────┘  └─────────────────────────────────────┘  │
│           │                                              │                      │
│           ▼                                              ▼                      │
│  ┌─────────────────────────────────────────────────────────────────────────────┐  │
│  │                       💾 STORAGE LAYER                                      │  │
│  │  ┌─────────────────┐  ┌─────────────────┐  ┌───────────────────────────┐   │  │
│  │  │ Job Status Cache│  │ Temporary Storage│  │     File System           │   │  │
│  │  │translation_jobs{}│  │  tmp/ directory │  │input/backup/translated   │   │  │
│  │  │   (In-Memory)   │  │                 │  │     (Auto cleanup)        │   │  │
│  │  └─────────────────┘  └─────────────────┘  └───────────────────────────┘   │  │
│  └─────────────────────────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────────┘
                    │
                    ▼ API Calls
┌─────────────────────────────────────────────────────────────────────────────────┐
│                    🔗 EXTERNAL SERVICES                                         │
├─────────────────────────────────────────────────────────────────────────────────┤
│  ┌─────────────────────────────────┐    ┌─────────────────────────────────────┐  │
│  │         GPT Service             │    │      Google Translate API          │  │
│  │      User Interaction           │    │         googletrans                │  │
│  └─────────────────────────────────┘    └─────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────────────────────┘

📊 실제 데이터 플로우:
Client ──HTTP──► Flask Router ──► File Manager ──► Storage Layer
                      │                               ▲
                      ▼                               │
              Translation Services ◄──────────────────┘
                      │
                      ▼
                Excel Processing ──► Format Preserver ──► Storage Layer
                      │
                      ▼
               External Services (Google API)
```

### 🔄 시스템 컴포넌트 상세

| 레이어 | 컴포넌트 | 역할 | 기술 스택 |
|--------|----------|------|-----------|
| **Client** | Web Interface | 사용자 인터페이스 | HTML5, CSS3, Vanilla JS |
| **Client** | Excel Preview | 파일 미리보기 | XLSX.js |
| **Client** | File Upload | 파일 업로드 처리 | Base64 Encoding |
| **Application** | Flask Router | HTTP 요청 라우팅 | Flask, CORS |
| **Application** | File Manager | 파일 시스템 관리 | Python OS, shutil |
| **Application** | GPT Workflow | GPT 번역 워크플로우 | Custom Logic |
| **Application** | Auto Translate | 자동 번역 서비스 | googletrans |
| **Application** | Excel Parser | 엑셀 파일 파싱 | openpyxl |
| **Application** | Format Preserver | 서식 보존 처리 | openpyxl + Custom |
| **External** | GPT Service | 외부 GPT 서비스 | User Interaction |
| **External** | Google API | 구글 번역 API | googletrans |
| **Storage** | Job Cache | 작업 상태 저장 | Python Dict (Memory) |
| **Storage** | Temp Storage | 임시 파일 저장 | File System (tmp/) |
| **Storage** | File System | 결과 파일 저장 | File System (Root) |

## 📊 데이터 플로우

### GPT 번역 워크플로우
```
1. File Upload → Base64 Encoding → Flask Server
2. Excel Parse → Cell Text Extraction → <CELL, TEXT> Format
3. GPT Prompt Generation → User GPT Interaction
4. GPT Response → Response Parsing → Translation Map
5. Translation Application → Format Preservation → File Download
```

### 자동 번역 플로우
```
1. File Upload → Excel Parse → Cell Detection
2. Text Filtering → Google Translate API → Translation
3. Cell Mapping → Format Preservation → File Download
```

## 🛠 기술 스택

- **Frontend**: HTML5, CSS3, JavaScript (Vanilla)
- **Backend**: Python Flask
- **번역 엔진**:
  - GPT 번역 워크플로우 (사용자 GPT 활용)
  - Google Translate API (자동 번역)
- **Excel 처리**: openpyxl (Python), XLSX.js (JavaScript)
- **파일 처리**: Base64 인코딩, 임시 파일 시스템

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
├── translate_server.py         # Flask 번역 서버 (GPT 워크플로우 + 자동 번역)
├── excel_translator_template.py # Excel 번역 모듈
├── index.html                  # 웹 인터페이스 (GPT 워크플로우 UI 포함)
├── style.css                   # 스타일시트 (GPT 워크플로우 스타일 포함)
├── script.js                   # 프론트엔드 로직 (GPT 워크플로우 구현)
├── requirements.txt            # Python 패키지
├── start_windows.bat           # Windows 실행 스크립트
├── build_exe.bat               # EXE 빌드 스크립트
├── tmp/                        # 임시 파일 저장 (Git 제외)
└── icons/                      # 아이콘 파일들
```

## 🆕 v1.1.0 업데이트 내용

- ✨ **GPT 번역 워크플로우 추가**: 3단계 GPT 기반 번역 지원
- 🔧 **API 엔드포인트 추가**: `/extract-words`, `/process-gpt-translation`
- 🎨 **UI/UX 개선**: 번역 방법 선택 및 GPT 워크플로우 인터페이스
- 🛡️ **파일 보안**: 이미지, 도형 등 모든 시각적 요소 완벽 보존
- 📁 **임시 파일 관리**: tmp 폴더 자동 생성 및 정리

## 📝 라이선스

MIT License

## 📞 문의

- 이메일: hgyjg@teclast.co.kr
- 회사: (주)테클라스트코리아

---
© 2025 Teclast Korea. All rights reserved.