class FileTranslator {
    constructor() {
        this.currentFile = null;
        this.workbook = null;
        this.translatedWorkbook = null;
        this.excludedCells = new Set();  // 제외할 셀 목록
        this.excludedSheets = new Set();  // 제외할 시트 목록
        this.excludedPatterns = new Set();  // 제외할 패턴 목록
        this.rangeStartCell = null;  // 범위 선택 시작 셀
        this.currentSheet = null;  // 현재 선택된 시트
        this.init();
    }

    init() {
        this.setupEventListeners();
    }

    setupEventListeners() {
        console.log('Setting up event listeners...');
        
        const backBtn = document.getElementById('backBtn');
        const selectFileBtn = document.getElementById('selectFileBtn');
        const fileInput = document.getElementById('fileInput');
        const uploadZone = document.getElementById('uploadZone');
        const translateBtn = document.getElementById('translateBtn');
        const downloadBtn = document.getElementById('downloadBtn');
        
        // File type cards click handlers
        const excelCard = document.querySelector('.file-type-card[data-type="excel"]');
        console.log('Excel card found:', excelCard);
        
        if (excelCard) {
            excelCard.addEventListener('click', (event) => {
                console.log('Excel card clicked!');
                event.preventDefault();
                event.stopPropagation();
                
                // Remove selected class from all cards
                document.querySelectorAll('.file-type-card').forEach(c => c.classList.remove('selected'));
                // Add selected class to clicked card
                excelCard.classList.add('selected');
                // Show Excel translator
                this.showExcelTranslator();
            });
        }

        backBtn.addEventListener('click', () => this.showFileTypeSelection());
        selectFileBtn.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e.target.files[0]));
        translateBtn.addEventListener('click', () => this.startTranslation());

        // GPT Workflow event listeners
        const extractWordsBtn = document.getElementById('extractWordsBtn');
        const copyPromptBtn = document.getElementById('copyPromptBtn');
        const applyTranslationBtn = document.getElementById('applyTranslationBtn');

        if (extractWordsBtn) {
            extractWordsBtn.addEventListener('click', () => this.extractWords());
        }
        if (copyPromptBtn) {
            copyPromptBtn.addEventListener('click', () => this.copyPromptToClipboard());
        }
        if (applyTranslationBtn) {
            applyTranslationBtn.addEventListener('click', () => this.applyGptTranslation());
        }

        // Translation method change listener
        const translationMethodRadios = document.querySelectorAll('input[name="translationMethod"]');
        translationMethodRadios.forEach(radio => {
            radio.addEventListener('change', () => this.onTranslationMethodChange());
        });
        downloadBtn.addEventListener('click', () => this.downloadTranslatedFile());

        uploadZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadZone.classList.add('dragover');
        });

        uploadZone.addEventListener('dragleave', () => {
            uploadZone.classList.remove('dragover');
        });

        uploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadZone.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.handleFileSelect(files[0]);
            }
        });

        uploadZone.addEventListener('click', () => fileInput.click());
    }

    showExcelTranslator() {
        document.getElementById('fileTypeSelection').classList.add('hidden');
        document.getElementById('excelTranslator').classList.remove('hidden');
    }

    showFileTypeSelection() {
        document.getElementById('excelTranslator').classList.add('hidden');
        document.getElementById('fileTypeSelection').classList.remove('hidden');
        this.resetTranslator();
    }

    resetTranslator() {
        this.currentFile = null;
        this.workbook = null;
        this.translatedWorkbook = null;
        this.translatedFileUrl = null;
        this.translatedFileName = null;
        
        document.getElementById('fileInfo').classList.add('hidden');
        document.getElementById('translationControls').classList.add('hidden');
        document.getElementById('progressArea').classList.add('hidden');
        document.getElementById('downloadArea').classList.add('hidden');
        
        document.getElementById('fileInput').value = '';
    }

    async handleFileSelect(file) {
        if (!file) return;

        if (!this.isValidExcelFile(file)) {
            alert('지원되지 않는 파일 형식입니다. .xlsx 또는 .xls 파일을 선택해주세요.');
            return;
        }

        this.currentFile = file;
        
        try {
            const data = await this.readFileAsArrayBuffer(file);
            this.workbook = XLSX.read(data, { type: 'array' });
            
            this.displayFileInfo(file);
            this.showTranslationControls();
            this.onTranslationMethodChange(); // 초기 번역 방법에 따른 UI 설정
        } catch (error) {
            console.error('파일 읽기 오류:', error);
            alert('파일을 읽는 중 오류가 발생했습니다.');
        }
    }

    isValidExcelFile(file) {
        const validExtensions = ['.xlsx', '.xls'];
        return validExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
    }

    readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    displayFileInfo(file) {
        document.getElementById('fileName').textContent = file.name;
        document.getElementById('fileSize').textContent = this.formatFileSize(file.size);
        document.getElementById('sheetCount').textContent = this.workbook.SheetNames.length;
        
        document.getElementById('fileInfo').classList.remove('hidden');
        
        // 엑셀 미리보기 표시
        this.showExcelPreview();
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    showTranslationControls() {
        document.getElementById('translationControls').classList.remove('hidden');
    }

    async startTranslation() {
        const direction = document.querySelector('input[name="direction"]:checked').value;
        const preserveEnglish = document.getElementById('preserveEnglish').checked;
        const addToNewSheet = document.getElementById('addToNewSheet').checked;

        document.getElementById('translateBtn').disabled = true;
        document.getElementById('progressArea').classList.remove('hidden');
        
        try {
            await this.translateWithPython(direction, preserveEnglish, addToNewSheet);
        } catch (error) {
            console.error('번역 오류:', error);
            alert('번역 중 오류가 발생했습니다.');
        } finally {
            document.getElementById('translateBtn').disabled = false;
        }
    }

    async translateWithPython(direction, preserveEnglish, addToNewSheet) {
        const formData = new FormData();
        formData.append('file', this.currentFile);
        formData.append('direction', direction === 'ko-to-zh' ? 'ko-zh' : 'zh-ko');
        formData.append('preserve_english', preserveEnglish ? 'true' : 'false');
        formData.append('add_new_sheet', addToNewSheet ? 'true' : 'false');
        
        // 번역 제외 설정 추가 (Set을 문자열로 변환)
        const excludeSheets = Array.from(this.excludedSheets).join(',');
        const excludeCells = Array.from(this.excludedCells).join(',');
        const excludePatterns = Array.from(this.excludedPatterns).join(',');
        
        if (excludeCells) {
            console.log(`제외할 셀 ${this.excludedCells.size}개:`, excludeCells);
        }
        
        formData.append('exclude_sheets', excludeSheets);
        formData.append('exclude_cells', excludeCells);
        formData.append('exclude_patterns', excludePatterns);

        this.updateProgress('Python 번역기로 파일 업로드 중...', 5);

        try {
            const response = await fetch('http://localhost:5001/translate-excel', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || '서버 오류');
            }

            const result = await response.json();
            
            if (result.success && result.job_id) {
                // 작업 ID를 받았으면 진행률을 주기적으로 확인
                await this.monitorTranslationProgress(result.job_id);
            } else {
                throw new Error(result.error || '번역 시작 실패');
            }

        } catch (error) {
            console.error('Python 번역 오류:', error);
            this.updateProgress('Python 번역 실패, JavaScript 번역기 사용...', 50);
            
            // 백업으로 기존 JavaScript 번역기 사용
            await this.translateWorkbook(direction, preserveEnglish, addToNewSheet);
            this.showDownloadArea();
        }
    }

    async monitorTranslationProgress(jobId) {
        const checkStatus = async () => {
            try {
                const response = await fetch(`http://localhost:5001/translation-status/${jobId}`);
                if (!response.ok) {
                    throw new Error('상태 조회 실패');
                }
                
                const status = await response.json();
                
                // 진행률 업데이트
                this.updateProgress(status.message || '번역 중...', status.progress || 0);
                
                if (status.status === 'completed') {
                    // 번역 완료
                    this.translatedFileUrl = `http://localhost:5001/download/${status.file_id}/${status.download_filename}`;
                    this.translatedFileName = status.download_filename.replace(`translated_${status.file_id}_`, '');
                    this.showDownloadArea();
                    return;
                } else if (status.status === 'error') {
                    // 오류 발생
                    throw new Error(status.error || '번역 중 오류 발생');
                } else {
                    // 아직 진행 중이면 1초 후 다시 확인
                    setTimeout(checkStatus, 1000);
                }
                
            } catch (error) {
                console.error('진행률 조회 오류:', error);
                // 백업으로 JavaScript 번역기 사용
                this.updateProgress('진행률 조회 실패, JavaScript 번역기 사용...', 50);
                const direction = document.querySelector('input[name="direction"]:checked').value;
                const preserveEnglish = document.getElementById('preserveEnglish').checked;
                const addToNewSheet = document.getElementById('addToNewSheet').checked;
                
                await this.translateWorkbook(direction, preserveEnglish, addToNewSheet);
                this.showDownloadArea();
            }
        };
        
        // 초기 상태 확인
        checkStatus();
    }

    async translateWorkbook(direction, preserveEnglish, addToNewSheet) {
        this.translatedWorkbook = XLSX.utils.book_new();
        const totalSheets = this.workbook.SheetNames.length;
        
        for (let i = 0; i < totalSheets; i++) {
            const sheetName = this.workbook.SheetNames[i];
            this.updateProgress(`시트 "${sheetName}" 번역 중...`, (i / totalSheets) * 100);
            
            const worksheet = this.workbook.Sheets[sheetName];
            const translatedSheet = await this.translateWorksheet(worksheet, direction, preserveEnglish);
            
            if (addToNewSheet) {
                XLSX.utils.book_append_sheet(this.translatedWorkbook, worksheet, sheetName);
                const translatedSheetName = direction === 'ko-to-zh' ? `${sheetName}_中文` : `${sheetName}_한국어`;
                XLSX.utils.book_append_sheet(this.translatedWorkbook, translatedSheet, translatedSheetName);
            } else {
                XLSX.utils.book_append_sheet(this.translatedWorkbook, translatedSheet, sheetName);
            }
        }
        
        this.updateProgress('번역 완료!', 100);
    }

    async translateWorksheet(worksheet, direction, preserveEnglish) {
        const translatedSheet = this.deepCopyWorksheet(worksheet);
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = worksheet[cellAddress];
                
                if (cell && cell.v && typeof cell.v === 'string' && cell.v.trim() !== '') {
                    let translatedValue = cell.v;
                    
                    if (preserveEnglish && this.isEnglish(cell.v)) {
                        translatedValue = cell.v;
                    } else {
                        translatedValue = await this.translateText(cell.v, direction);
                    }
                    
                    translatedSheet[cellAddress].v = translatedValue;
                    if (translatedSheet[cellAddress].w) {
                        translatedSheet[cellAddress].w = translatedValue;
                    }
                }
            }
        }
        
        return translatedSheet;
    }

    deepCopyWorksheet(worksheet) {
        const copied = {};
        
        for (const key in worksheet) {
            if (worksheet.hasOwnProperty(key)) {
                if (key.startsWith('!')) {
                    if (key === '!merges' && Array.isArray(worksheet[key])) {
                        copied[key] = worksheet[key].map(merge => ({ ...merge }));
                    } else if (key === '!cols' && Array.isArray(worksheet[key])) {
                        copied[key] = worksheet[key].map(col => ({ ...col }));
                    } else if (key === '!rows' && Array.isArray(worksheet[key])) {
                        copied[key] = worksheet[key].map(row => ({ ...row }));
                    } else if (typeof worksheet[key] === 'object') {
                        copied[key] = { ...worksheet[key] };
                    } else {
                        copied[key] = worksheet[key];
                    }
                } else {
                    const cell = worksheet[key];
                    if (cell && typeof cell === 'object') {
                        copied[key] = {
                            v: cell.v,
                            t: cell.t,
                            w: cell.w,
                            f: cell.f,
                            s: cell.s,
                            z: cell.z,
                            l: cell.l,
                            c: cell.c,
                            h: cell.h,
                            r: cell.r
                        };
                        
                        if (cell.s) {
                            copied[key].s = { ...cell.s };
                        }
                    }
                }
            }
        }
        
        return copied;
    }

    isEnglish(text) {
        const englishPattern = /^[a-zA-Z0-9\s\.,\-\(\)\[\]\{\}]+$/;
        return englishPattern.test(text.trim());
    }

    async translateText(text, direction) {
        try {
            const isKorean = /[가-힣]/.test(text);
            const isChinese = /[\u4e00-\u9fff]/.test(text);
            
            if ((direction === 'ko-to-zh' && !isKorean) || (direction === 'zh-to-ko' && !isChinese)) {
                return text;
            }
            
            const translationDirection = direction === 'ko-to-zh' ? 'ko-zh' : 'zh-ko';
            
            try {
                const response = await fetch('http://localhost:5001/translate', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({
                        text: text,
                        direction: translationDirection
                    })
                });
                
                if (response.ok) {
                    const data = await response.json();
                    if (data.translatedText) {
                        await new Promise(resolve => setTimeout(resolve, 100));
                        return data.translatedText;
                    }
                }
            } catch (serverError) {
                console.warn('번역 서버 연결 실패, 백업 번역 사용');
            }
            
            return this.fallbackTranslation(text, direction);
            
        } catch (error) {
            console.error('번역 오류:', error);
            return this.fallbackTranslation(text, direction);
        }
    }

    fallbackTranslation(text, direction) {
        if (direction === 'ko-to-zh') {
            return this.mockTranslateKoreanToChinese(text);
        } else {
            return this.mockTranslateChineseToKorean(text);
        }
    }

    mockTranslateKoreanToChinese(text) {
        const translations = {
            '발주서': '订单书',
            '수주처': '接单处',
            '상호': '商号',
            '대표': '代表',
            '발주일': '订单日期',
            '이메일': '邮箱',
            '연락처': '联系方式',
            '주소': '地址',
            '납기일자': '交货日期',
            '발송정보': '配送信息',
            '발송일': '发货日',
            '품목': '品目',
            '단위': '单位',
            '수량': '数量',
            '구분': '区分',
            '비고': '备注',
            '합계': '合计',
            '요구사항': '要求事项',
            '확인': '确认',
            '아래와 같이 발주합니다': '订单如下',
            '주식회사': '股份有限公司',
            '테클라스트코리아': '泰克拉斯特韩国',
            '이상모': '李相模',
            '등록번호': '注册号码',
            '경기도': '京畿道',
            '광명시': '光明市',
            '하안로': '下安路',
            '광명테크노파크': '光明科技园',
            '서비스': '服务',
            '도소매': '批发零售',
            '종목': '种目',
            '태블릿PC': '平板电脑',
            '유재건부장': '刘在建部长',
            '심대용과장': '沈大龙科长',
            '반입분': '入库分',
            '남품장소': '南品场所',
            '남품일정': '南品日程'
        };
        
        let translated = text;
        for (const [korean, chinese] of Object.entries(translations)) {
            translated = translated.replace(new RegExp(korean, 'g'), chinese);
        }
        
        return translated;
    }

    mockTranslateChineseToKorean(text) {
        const translations = {
            '订单书': '발주서',
            '接单处': '수주처',
            '商号': '상호',
            '代表': '대표',
            '订单日期': '발주일',
            '邮箱': '이메일',
            '联系方式': '연락처',
            '地址': '주소',
            '交货日期': '납기일자',
            '配送信息': '발송정보',
            '发货日': '발송일',
            '品目': '품목',
            '单位': '단위',
            '数量': '수량',
            '区分': '구분',
            '备注': '비고',
            '合计': '합계',
            '要求事项': '요구사항',
            '确认': '확인',
            '订单如下': '아래와 같이 발주합니다',
            '股份有限公司': '주식회사',
            '泰克拉斯特韩国': '테클라스트코리아',
            '李相模': '이상모',
            '注册号码': '등록번호',
            '京畿道': '경기도',
            '光明市': '광명시',
            '下安路': '하안로',
            '光明科技园': '광명테크노파크',
            '服务': '서비스',
            '批发零售': '도소매',
            '种目': '종목',
            '平板电脑': '태블릿PC',
            '刘在建部长': '유재건부장',
            '沈大龙科长': '심대용과장',
            '入库分': '반입분',
            '南品场所': '남품장소',
            '南品日程': '남품일정'
        };
        
        let translated = text;
        for (const [chinese, korean] of Object.entries(translations)) {
            translated = translated.replace(new RegExp(chinese, 'g'), korean);
        }
        
        return translated;
    }

    updateProgress(text, percentage) {
        document.getElementById('progressText').textContent = text;
        document.getElementById('progressPercent').textContent = `${Math.round(percentage)}%`;
        document.getElementById('progressFill').style.width = `${percentage}%`;
    }

    showDownloadArea() {
        document.getElementById('progressArea').classList.add('hidden');
        document.getElementById('downloadArea').classList.remove('hidden');
    }

    showExcelPreview() {
        document.getElementById('excelPreview').classList.remove('hidden');
        
        // 시트 선택 드롭다운 채우기
        const sheetSelector = document.getElementById('sheetSelector');
        sheetSelector.innerHTML = '';
        
        this.workbook.SheetNames.forEach((sheetName, index) => {
            const option = document.createElement('option');
            option.value = sheetName;
            option.textContent = sheetName;
            sheetSelector.appendChild(option);
        });
        
        // 첫 번째 시트 표시
        if (this.workbook.SheetNames.length > 0) {
            this.currentSheet = this.workbook.SheetNames[0];
            this.renderSheet(this.currentSheet);
        }
        
        // 이벤트 리스너 설정
        this.setupPreviewEventListeners();
    }
    
    setupPreviewEventListeners() {
        const sheetSelector = document.getElementById('sheetSelector');
        const clearBtn = document.getElementById('clearSelectionBtn');
        const excludeSheetBtn = document.getElementById('excludeSheetBtn');
        
        sheetSelector.addEventListener('change', (e) => {
            this.currentSheet = e.target.value;
            this.renderSheet(this.currentSheet);
        });
        
        clearBtn.addEventListener('click', () => {
            this.excludedCells.clear();
            this.excludedSheets.clear();
            this.excludedPatterns.clear();
            this.renderSheet(this.currentSheet);
            this.updateExcludedSummary();
        });
        
        excludeSheetBtn.addEventListener('click', () => {
            if (this.currentSheet) {
                if (this.excludedSheets.has(this.currentSheet)) {
                    this.excludedSheets.delete(this.currentSheet);
                    excludeSheetBtn.textContent = '이 시트 전체 제외';
                } else {
                    this.excludedSheets.add(this.currentSheet);
                    excludeSheetBtn.textContent = '시트 제외 취소';
                }
                this.renderSheet(this.currentSheet);
                this.updateExcludedSummary();
            }
        });
    }
    
    renderSheet(sheetName) {
        const sheet = this.workbook.Sheets[sheetName];
        const container = document.getElementById('tableContainer');
        const excludeSheetBtn = document.getElementById('excludeSheetBtn');
        
        // 시트 제외 버튼 텍스트 업데이트
        if (this.excludedSheets.has(sheetName)) {
            excludeSheetBtn.textContent = '시트 제외 취소';
        } else {
            excludeSheetBtn.textContent = '이 시트 전체 제외';
        }
        
        // HTML 테이블 생성
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
        let html = '<table class="excel-table">';
        
        // 헤더 행 (열 번호)
        html += '<thead><tr><th></th>';
        for (let C = range.s.c; C <= range.e.c && C < 20; ++C) {  // 최대 20열까지만 표시
            html += `<th>${XLSX.utils.encode_col(C)}</th>`;
        }
        html += '</tr></thead><tbody>';
        
        // 데이터 행
        for (let R = range.s.r; R <= range.e.r && R < 50; ++R) {  // 최대 50행까지만 표시
            html += '<tr>';
            html += `<td class="row-header">${R + 1}</td>`;  // 행 번호
            
            for (let C = range.s.c; C <= range.e.c && C < 20; ++C) {
                const cellAddress = XLSX.utils.encode_cell({r: R, c: C});
                const cell = sheet[cellAddress];
                const value = cell ? (cell.w || cell.v || '') : '';
                const cellRef = `${sheetName}!${cellAddress}`;
                
                let className = '';
                if (this.excludedSheets.has(sheetName) || this.excludedCells.has(cellRef)) {
                    className = 'excluded';
                }
                
                html += `<td class="${className}" data-cell="${cellRef}" data-value="${this.escapeHtml(value)}">${this.escapeHtml(value)}</td>`;
            }
            html += '</tr>';
        }
        
        html += '</tbody></table>';
        container.innerHTML = html;
        
        // 셀 클릭 이벤트 설정
        this.setupCellClickEvents();
    }
    
    setupCellClickEvents() {
        const cells = document.querySelectorAll('.excel-table td:not(.row-header)');
        
        cells.forEach(cell => {
            cell.addEventListener('click', (e) => {
                const cellRef = cell.dataset.cell;
                const isShiftPressed = e.shiftKey;
                
                if (!cellRef) return;
                
                if (isShiftPressed && this.rangeStartCell) {
                    // 범위 선택
                    this.selectCellRange(this.rangeStartCell, cellRef);
                } else {
                    // 단일 셀 선택
                    if (this.excludedCells.has(cellRef)) {
                        this.excludedCells.delete(cellRef);
                        cell.classList.remove('excluded');
                    } else {
                        this.excludedCells.add(cellRef);
                        cell.classList.add('excluded');
                    }
                    this.rangeStartCell = cellRef;
                }
                
                this.updateExcludedSummary();
            });
            
            // 호버 효과로 범위 선택 미리보기
            cell.addEventListener('mouseenter', (e) => {
                if (e.shiftKey && this.rangeStartCell) {
                    this.previewCellRange(this.rangeStartCell, cell.dataset.cell);
                }
            });
            
            cell.addEventListener('mouseleave', () => {
                this.clearRangePreview();
            });
        });
    }
    
    selectCellRange(startCell, endCell) {
        const [startSheet, startAddr] = startCell.split('!');
        const [endSheet, endAddr] = endCell.split('!');
        
        if (startSheet !== endSheet) return;
        
        const start = XLSX.utils.decode_cell(startAddr);
        const end = XLSX.utils.decode_cell(endAddr);
        
        const minRow = Math.min(start.r, end.r);
        const maxRow = Math.max(start.r, end.r);
        const minCol = Math.min(start.c, end.c);
        const maxCol = Math.max(start.c, end.c);
        
        // 범위 내의 모든 셀 선택/해제
        for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
                const cellAddr = XLSX.utils.encode_cell({r, c});
                const cellRef = `${startSheet}!${cellAddr}`;
                
                // 토글
                if (this.excludedCells.has(cellRef)) {
                    this.excludedCells.delete(cellRef);
                } else {
                    this.excludedCells.add(cellRef);
                }
            }
        }
        
        // 화면 업데이트
        this.renderSheet(this.currentSheet);
        this.updateExcludedSummary();
    }
    
    previewCellRange(startCell, endCell) {
        if (!startCell || !endCell) return;
        
        const [startSheet, startAddr] = startCell.split('!');
        const [endSheet, endAddr] = endCell.split('!');
        
        if (startSheet !== endSheet) return;
        
        const start = XLSX.utils.decode_cell(startAddr);
        const end = XLSX.utils.decode_cell(endAddr);
        
        const minRow = Math.min(start.r, end.r);
        const maxRow = Math.max(start.r, end.r);
        const minCol = Math.min(start.c, end.c);
        const maxCol = Math.max(start.c, end.c);
        
        // 미리보기 스타일 적용
        document.querySelectorAll('.excel-table td').forEach(cell => {
            cell.classList.remove('range-selecting');
        });
        
        for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
                const cellAddr = XLSX.utils.encode_cell({r, c});
                const cellRef = `${startSheet}!${cellAddr}`;
                const cell = document.querySelector(`[data-cell="${cellRef}"]`);
                if (cell) {
                    cell.classList.add('range-selecting');
                }
            }
        }
    }
    
    clearRangePreview() {
        document.querySelectorAll('.range-selecting').forEach(cell => {
            cell.classList.remove('range-selecting');
        });
    }
    
    updateExcludedSummary() {
        const summary = document.getElementById('excludedSummary');
        let summaryText = '';
        
        if (this.excludedSheets.size > 0) {
            summaryText += `제외된 시트: ${Array.from(this.excludedSheets).join(', ')}<br>`;
        }
        
        if (this.excludedCells.size > 0) {
            // 셀을 시트별로 그룹화
            const cellsBySheet = {};
            this.excludedCells.forEach(cellRef => {
                const [sheet, cell] = cellRef.split('!');
                if (!cellsBySheet[sheet]) cellsBySheet[sheet] = [];
                cellsBySheet[sheet].push(cell);
            });
            
            for (const sheet in cellsBySheet) {
                summaryText += `${sheet} 시트의 제외 셀: ${cellsBySheet[sheet].join(', ')}<br>`;
            }
        }
        
        if (this.excludedPatterns.size > 0) {
            summaryText += `제외 패턴: ${Array.from(this.excludedPatterns).join(', ')}<br>`;
        }
        
        summary.innerHTML = summaryText || '없음';
    }
    
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    downloadTranslatedFile() {
        if (this.translatedFileUrl) {
            // Python 번역 결과 다운로드
            const link = document.createElement('a');
            link.href = this.translatedFileUrl;
            link.download = this.translatedFileName || 'translated_file.xlsx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        } else if (this.translatedWorkbook) {
            // JavaScript 번역 결과 다운로드
            const direction = document.querySelector('input[name="direction"]:checked').value;
            const originalName = this.currentFile.name.replace(/\.[^/.]+$/, "");
            const suffix = direction === 'ko-to-zh' ? '_중문번역' : '_한국어번역';
            const fileName = `${originalName}${suffix}.xlsx`;

            XLSX.writeFile(this.translatedWorkbook, fileName);
        } else {
            alert('번역된 파일이 없습니다.');
        }
    }

    // GPT Workflow Methods
    onTranslationMethodChange() {
        const method = document.querySelector('input[name="translationMethod"]:checked').value;
        const gptWorkflowArea = document.getElementById('gptWorkflowArea');
        const translationControls = document.getElementById('translationControls');

        if (method === 'gpt') {
            // GPT 워크플로우 표시
            gptWorkflowArea.classList.remove('hidden');
            translationControls.classList.add('hidden');
        } else {
            // 기존 자동 번역 방식
            gptWorkflowArea.classList.add('hidden');
            translationControls.classList.remove('hidden');
        }
    }

    async extractWords() {
        if (!this.currentFile) {
            alert('먼저 파일을 선택해주세요.');
            return;
        }

        const extractWordsBtn = document.getElementById('extractWordsBtn');
        const step1Result = document.getElementById('step1Result');

        try {
            extractWordsBtn.disabled = true;
            extractWordsBtn.textContent = '단어 추출 중...';

            const formData = new FormData();
            formData.append('file', this.currentFile);
            formData.append('direction', this.getDirection());

            const response = await fetch('/extract-words', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                // 결과 표시
                document.getElementById('wordCount').textContent = result.word_count;
                document.getElementById('wordList').textContent = result.word_list.join('\n');
                document.getElementById('gptPrompt').value = result.gpt_prompt;

                step1Result.classList.remove('hidden');

                // 2단계 활성화
                document.getElementById('step2').style.opacity = '1';
                document.getElementById('copyPromptBtn').disabled = false;

                // 추출 결과 저장
                this.extractedWords = result.word_list;
                this.gptPrompt = result.gpt_prompt;

                extractWordsBtn.textContent = '단어 추출 완료';
            } else {
                alert(`단어 추출 실패: ${result.error}`);
                extractWordsBtn.disabled = false;
                extractWordsBtn.textContent = '단어 추출하기';
            }
        } catch (error) {
            console.error('단어 추출 오류:', error);
            alert('단어 추출 중 오류가 발생했습니다.');
            extractWordsBtn.disabled = false;
            extractWordsBtn.textContent = '단어 추출하기';
        }
    }

    async copyPromptToClipboard() {
        const gptPrompt = document.getElementById('gptPrompt');
        const copyBtn = document.getElementById('copyPromptBtn');

        try {
            await navigator.clipboard.writeText(gptPrompt.value);
            const originalText = copyBtn.textContent;
            copyBtn.textContent = '복사됨!';
            copyBtn.style.background = '#059669';

            // 3단계 활성화
            document.getElementById('step3').style.opacity = '1';
            document.getElementById('applyTranslationBtn').disabled = false;

            setTimeout(() => {
                copyBtn.textContent = originalText;
                copyBtn.style.background = '#10b981';
            }, 2000);
        } catch (error) {
            console.error('클립보드 복사 실패:', error);
            alert('클립보드 복사에 실패했습니다.');
        }
    }

    async applyGptTranslation() {
        const gptResponse = document.getElementById('gptResponse').value.trim();
        const applyBtn = document.getElementById('applyTranslationBtn');

        if (!gptResponse) {
            alert('GPT 번역 결과를 입력해주세요.');
            return;
        }

        if (!this.currentFile) {
            alert('원본 파일이 없습니다.');
            return;
        }

        try {
            applyBtn.disabled = true;
            applyBtn.textContent = '번역 적용 중...';

            // 파일을 base64로 변환
            const fileBase64 = await this.fileToBase64(this.currentFile);

            const requestData = {
                gpt_response: gptResponse,
                original_file: fileBase64,
                direction: this.getDirection(),
                preserve_english: document.getElementById('preserveEnglish').checked,
                add_new_sheet: document.getElementById('addToNewSheet').checked
            };

            const response = await fetch('/process-gpt-translation', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(requestData)
            });

            const result = await response.json();

            if (result.success) {
                applyBtn.textContent = '번역 적용 완료';

                // 다운로드 영역 표시
                this.translatedFileUrl = `/download/${result.job_id}/${encodeURIComponent('translated_' + result.job_id + '.xlsx')}`;
                this.translatedFileName = `translated_${this.currentFile.name}`;

                document.getElementById('downloadArea').classList.remove('hidden');
                alert(`번역이 완료되었습니다! ${result.translations_applied}개 항목이 번역되었습니다.`);
            } else {
                alert(`번역 적용 실패: ${result.error}`);
                applyBtn.disabled = false;
                applyBtn.textContent = '번역 적용하기';
            }
        } catch (error) {
            console.error('번역 적용 오류:', error);
            alert('번역 적용 중 오류가 발생했습니다.');
            applyBtn.disabled = false;
            applyBtn.textContent = '번역 적용하기';
        }
    }

    fileToBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => {
                // "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," 부분 제거
                const base64 = reader.result.split(',')[1];
                resolve(base64);
            };
            reader.onerror = error => reject(error);
        });
    }

    getDirection() {
        const direction = document.querySelector('input[name="direction"]:checked').value;
        return direction === 'ko-to-zh' ? 'ko-zh' : 'zh-ko';
    }
}

document.addEventListener('DOMContentLoaded', () => {
    window.fileTranslator = new FileTranslator();
});