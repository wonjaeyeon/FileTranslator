#!/usr/bin/env python3
"""
간단한 번역 프록시 서버
CORS 문제를 해결하고 안정적인 번역을 위해 사용
"""

from flask import Flask, request, jsonify
from flask_cors import CORS
import requests
import json
import time
import threading
import uuid

app = Flask(__name__)
CORS(app)

# 번역 작업 상태 저장
translation_jobs = {}

# LibreTranslate 공개 인스턴스들
LIBRETRANSLATE_URLS = [
    "https://libretranslate.de/translate",
    "https://translate.argosopentech.com/translate",
    "https://libretranslate.com/translate"
]

def translate_with_libretranslate(text, source_lang, target_lang):
    """LibreTranslate API를 사용한 번역"""
    for url in LIBRETRANSLATE_URLS:
        try:
            data = {
                "q": text,
                "source": source_lang,
                "target": target_lang,
                "format": "text"
            }
            
            response = requests.post(url, json=data, timeout=8)
            if response.status_code == 200:
                result = response.json()
                if 'translatedText' in result and result['translatedText'].strip():
                    return result['translatedText']
        except Exception:
            continue
    return None

def translate_with_google(text, direction):
    """Google Translate 무료 API로 번역"""
    try:
        from googletrans import Translator
        translator = Translator()
        
        if direction == 'ko-zh':
            source, target = 'ko', 'zh-CN'
        else:
            source, target = 'zh-CN', 'ko'
        
        result = translator.translate(text, src=source, dest=target)
        if result and result.text and result.text.strip() != text:
            return result.text.strip()
    except Exception:
        pass
    return None

def translate_with_deep_translator(text, direction):
    """Deep Translator로 번역 (Google, Bing, DeepL 등 지원)"""
    try:
        from deep_translator import GoogleTranslator, BingTranslator
        
        if direction == 'ko-zh':
            source, target = 'ko', 'zh-CN'
        else:
            source, target = 'zh-CN', 'ko'
        
        # Google Translator 시도
        try:
            translator = GoogleTranslator(source=source, target=target)
            result = translator.translate(text)
            if result and result.strip() != text:
                return result.strip()
        except:
            pass
        
        # Bing Translator 시도
        try:
            translator = BingTranslator(source=source, target=target)
            result = translator.translate(text)
            if result and result.strip() != text:
                return result.strip()
        except:
            pass
            
    except Exception:
        pass
    return None

def translate_with_ollama(text, direction):
    """Ollama 로컬 LLM으로 번역"""
    try:
        if direction == 'ko-zh':
            prompt = f"다음 한국어 텍스트를 중국어로 정확하게 번역해주세요. 영어는 그대로 유지하세요. 번역 결과만 답해주세요:\n\n{text}"
        else:
            prompt = f"다음 중국어 텍스트를 한국어로 정확하게 번역해주세요. 영어는 그대로 유지하세요. 번역 결과만 답해주세요:\n\n{text}"
        
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={
                "model": "llama3.2:3b",  # 가벼운 모델
                "prompt": prompt,
                "stream": False
            },
            timeout=15
        )
        
        if response.status_code == 200:
            result = response.json()
            translation = result.get('response', '').strip()
            if translation and translation != text:
                return translation
    except Exception:
        pass
    return None

def is_korean(text):
    import re
    return bool(re.search('[가-힣]', str(text)))

def is_chinese(text):
    import re
    return bool(re.search('[\u4e00-\u9fff]', str(text)))

def is_english(text):
    import re
    return bool(re.match(r'^[a-zA-Z0-9\s\.\,\-\(\)\[\]\{\}@:\/]+$', str(text).strip()))

def fallback_translation(text, direction):
    """백업 번역 사전"""
    if direction == 'ko-zh':
        translations = {
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
            '남품일정': '南品日程',
            # 추가 번역 용어들
            '프로젝트': '项目',
            '업무': '业务',
            '관리자': '管理员',
            '완료': '完成',
            '고객': '客户',
            '회사': '公司',
            '부서': '部门',
            '담당자': '负责人',
            '직원': '职员',
            '팀장': '组长',
            '과장': '科长',
            '부장': '部长',
            '사장': '社长',
            '작업': '工作',
            '계획': '计划',
            '일정': '日程',
            '진행': '进行',
            '시작': '开始',
            '종료': '结束',
            '검토': '审查',
            '승인': '批准',
            '변경': '变更',
            '수정': '修改',
            '업데이트': '更新',
            '품질': '质量',
            '성능': '性能',
            '정보': '信息',
            '데이터': '数据',
            '문서': '文件',
            '보고서': '报告书',
            '제안서': '提案书',
            '계획서': '计划书',
            '내용': '内容',
            '항목': '项',
            '목록': '列表',
            '설정': '设置',
            '상태': '状态',
            '결과': '结果',
            '목적': '目的',
            '목표': '目标',
            '방법': '方法',
            '절차': '程序',
            '과정': '过程',
            '단계': '阶段',
            '관리': '管理',
            '점검': '检查',
            '테스트': '测试',
            '확인': '确认',
            '평가': '评价',
            '분석': '分析',
            '개발': '开发',
            '설계': '设计',
            '생산': '生产',
            '제조': '制造',
            '설치': '安装',
            '배송': '配送',
            '시스템': '系统',
            '장비': '设备',
            '제품': '产品',
            '모델': '型号',
            '종류': '种类',
            '크기': '尺寸',
            '시간': '时间',
            '기간': '期间',
            '가격': '价格',
            '비용': '费用',
            '금액': '金额',
            '총계': '总计',
            '기본': '基本',
            '표준': '标准',
            '특별': '特别',
            '현재': '现在',
            '새로운': '新的',
            '최신': '最新'
        }
    else:  # zh-ko
        translations = {
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
            '南品日程': '남품일정',
            # 추가 번역 용어들 (중국어 -> 한국어)
            '项目': '프로젝트',
            '业务': '업무',
            '管理员': '관리자',
            '完成': '완료',
            '客户': '고객',
            '公司': '회사',
            '部门': '부서',
            '负责人': '담당자',
            '职员': '직원',
            '组长': '팀장',
            '科长': '과장',
            '部长': '부장',
            '社长': '사장',
            '工作': '작업',
            '计划': '계획',
            '日程': '일정',
            '进行': '진행',
            '开始': '시작',
            '结束': '종료',
            '审查': '검토',
            '批准': '승인',
            '变更': '변경',
            '修改': '수정',
            '更新': '업데이트',
            '质量': '품질',
            '性能': '성능',
            '信息': '정보',
            '数据': '데이터',
            '文件': '문서',
            '报告书': '보고서',
            '提案书': '제안서',
            '计划书': '계획서',
            '内容': '내용',
            '项': '항목',
            '列表': '목록',
            '设置': '설정',
            '状态': '상태',
            '结果': '결과',
            '目的': '목적',
            '目标': '목표',
            '方法': '방법',
            '程序': '절차',
            '过程': '과정',
            '阶段': '단계',
            '管理': '관리',
            '检查': '점검',
            '测试': '테스트',
            '确认': '확인',
            '评价': '평가',
            '分析': '분석',
            '开发': '개발',
            '设计': '설계',
            '生产': '생산',
            '制造': '제조',
            '安装': '설치',
            '配送': '배송',
            '系统': '시스템',
            '设备': '장비',
            '产品': '제품',
            '型号': '모델',
            '种类': '종류',
            '尺寸': '크기',
            '时间': '시간',
            '期间': '기간',
            '价格': '가격',
            '费用': '비용',
            '金额': '금액',
            '总计': '총계',
            '基本': '기본',
            '标准': '표준',
            '特别': '특별',
            '现在': '현재',
            '新的': '새로운',
            '最新': '최신'
        }
    
    translated = text
    for original, translation in translations.items():
        translated = translated.replace(original, translation)
    
    return translated

def hybrid_translate_text(text, direction, preserve_english=True):
    """하이브리드 번역 시스템 (사전 -> API -> LLM)"""
    if not text or not isinstance(text, str) or not str(text).strip():
        return text
    
    text_str = str(text).strip()
    
    # 영문만으로 구성된 텍스트는 유지
    if preserve_english and is_english(text_str):
        return text
    
    # 혼합 텍스트 처리 (영문 + 한국어/중국어)
    if direction == 'ko-zh':
        if not is_korean(text_str):
            return text  # 한국어가 없으면 번역하지 않음
    elif direction == 'zh-ko':
        if not is_chinese(text_str):
            return text  # 중국어가 없으면 번역하지 않음
    
    # 1단계: 번역 사전 먼저 시도 (빠르고 정확)
    dict_translated = fallback_translation(text_str, direction)
    has_dict_translation = dict_translated != text_str
    
    # 사전 번역 후에도 원본 언어가 남아있는지 확인
    still_has_original_lang = False
    if direction == 'ko-zh' and is_korean(dict_translated):
        still_has_original_lang = True
    elif direction == 'zh-ko' and is_chinese(dict_translated):
        still_has_original_lang = True
    
    # 사전만으로 완전 번역되면 사전 결과 반환
    if has_dict_translation and not still_has_original_lang:
        return dict_translated
    
    # 2단계: Google Translate 시도 (전체 문장 번역)
    google_translation = translate_with_google(text_str, direction)
    if google_translation:
        return google_translation
    
    # 3단계: Deep Translator (Google/Bing) 시도
    deep_translation = translate_with_deep_translator(text_str, direction)
    if deep_translation:
        return deep_translation
    
    # 4단계: LibreTranslate API 시도
    source_lang = 'ko' if direction == 'ko-zh' else 'zh'
    target_lang = 'zh' if direction == 'ko-zh' else 'ko'
    
    libre_translation = translate_with_libretranslate(text_str, source_lang, target_lang)
    if libre_translation:
        return libre_translation
    
    # 5단계: Ollama 로컬 LLM 시도
    ollama_translation = translate_with_ollama(text_str, direction)
    if ollama_translation:
        return ollama_translation
    
    # 6단계: 모든 방법 실패시 사전 번역 결과라도 반환
    return dict_translated

@app.route('/translate', methods=['POST'])
def translate():
    try:
        data = request.json
        text = data.get('text', '')
        direction = data.get('direction', 'ko-zh')
        preserve_english = data.get('preserve_english', True)
        
        if not text.strip():
            return jsonify({'translatedText': text})
        
        # 하이브리드 번역 시스템 사용
        translated = hybrid_translate_text(text, direction, preserve_english)
        
        # 어떤 방법이 사용되었는지 확인
        dict_only = fallback_translation(text, direction)
        if translated == text:
            method = 'unchanged'
        elif translated == dict_only:
            # 사전만으로 번역된 경우라도, 원본 언어가 남아있으면 API 필요
            if direction == 'ko-zh' and is_korean(translated):
                method = 'partial_dictionary'
            elif direction == 'zh-ko' and is_chinese(translated):
                method = 'partial_dictionary'
            else:
                method = 'dictionary'
        else:
            method = 'api_or_llm'
        
        return jsonify({
            'translatedText': translated, 
            'method': method,
            'original': text
        })
            
    except Exception as e:
        print(f"번역 오류: {e}")
        return jsonify({'error': str(e)}), 500

def run_translation(job_id, input_path, output_path, direction, preserve_english, add_new_sheet, exclude_sheets=None, exclude_cells=None, exclude_patterns=None):
    """백그라운드에서 번역 실행"""
    try:
        from excel_translator_template import ExcelTranslatorTemplate
        
        def progress_callback(message, percentage):
            translation_jobs[job_id]['progress'] = percentage
            translation_jobs[job_id]['message'] = message
            print(f"Job {job_id}: {message} ({percentage}%)")
        
        translation_jobs[job_id]['status'] = 'running'
        translation_jobs[job_id]['progress'] = 0
        translation_jobs[job_id]['message'] = '번역 시작'
        
        translator = ExcelTranslatorTemplate(progress_callback)
        result_path = translator.translate_excel_file(
            input_path=input_path,
            output_path=output_path,
            direction=direction,
            preserve_english=preserve_english,
            add_new_sheet=add_new_sheet,
            exclude_sheets=exclude_sheets,
            exclude_cells=exclude_cells,
            exclude_patterns=exclude_patterns
        )
        
        translation_jobs[job_id]['status'] = 'completed'
        translation_jobs[job_id]['progress'] = 100
        translation_jobs[job_id]['message'] = '번역 완료'
        translation_jobs[job_id]['result_path'] = result_path
        
        # 임시 파일 정리
        import os
        if os.path.exists(input_path):
            os.remove(input_path)
            
    except Exception as e:
        translation_jobs[job_id]['status'] = 'error'
        translation_jobs[job_id]['error'] = str(e)
        print(f"번역 오류 (Job {job_id}): {e}")

@app.route('/translate-excel', methods=['POST'])
def translate_excel():
    try:
        if 'file' not in request.files:
            return jsonify({'error': '파일이 업로드되지 않았습니다.'}), 400
        
        file = request.files['file']
        direction = request.form.get('direction', 'ko-zh')
        preserve_english = request.form.get('preserve_english', 'true').lower() == 'true'
        add_new_sheet = request.form.get('add_new_sheet', 'true').lower() == 'true'
        
        # 번역 제외 설정 파싱
        exclude_sheets_str = request.form.get('exclude_sheets', '')
        exclude_cells_str = request.form.get('exclude_cells', '')
        exclude_patterns_str = request.form.get('exclude_patterns', '')
        
        exclude_sheets = [s.strip() for s in exclude_sheets_str.split(',') if s.strip()] if exclude_sheets_str else None
        exclude_cells = [c.strip() for c in exclude_cells_str.split(',') if c.strip()] if exclude_cells_str else None
        exclude_patterns = [p.strip() for p in exclude_patterns_str.split(',') if p.strip()] if exclude_patterns_str else None
        
        if exclude_cells:
            print(f"제외할 셀: {len(exclude_cells)}개")
        if exclude_sheets:
            print(f"제외할 시트: {exclude_sheets}")
        if exclude_patterns:
            print(f"제외할 패턴: {exclude_patterns}")
        
        if file.filename == '':
            return jsonify({'error': '파일이 선택되지 않았습니다.'}), 400
        
        # 파일 저장
        import os
        
        unique_id = str(uuid.uuid4())[:8]
        input_filename = f"input_{unique_id}_{file.filename}"
        output_filename = f"translated_{unique_id}_{file.filename}"
        
        input_path = os.path.join(os.getcwd(), input_filename)
        output_path = os.path.join(os.getcwd(), output_filename)
        
        file.save(input_path)
        
        # 번역 작업 정보 저장
        job_id = unique_id
        translation_jobs[job_id] = {
            'status': 'queued',
            'progress': 0,
            'message': '번역 준비 중...',
            'output_filename': output_filename,
            'original_filename': file.filename
        }
        
        # 백그라운드에서 번역 실행
        thread = threading.Thread(
            target=run_translation,
            args=(job_id, input_path, output_path, direction, preserve_english, add_new_sheet, exclude_sheets, exclude_cells, exclude_patterns)
        )
        thread.daemon = True
        thread.start()
        
        return jsonify({
            'success': True,
            'job_id': job_id,
            'message': '번역이 시작되었습니다.'
        })
        
    except Exception as e:
        print(f"엑셀 번역 오류: {e}")
        return jsonify({'error': f'번역 중 오류가 발생했습니다: {str(e)}'}), 500

@app.route('/translation-status/<job_id>', methods=['GET'])
def get_translation_status(job_id):
    """번역 상태 조회"""
    if job_id not in translation_jobs:
        return jsonify({'error': '작업을 찾을 수 없습니다.'}), 404
    
    job = translation_jobs[job_id]
    
    if job['status'] == 'completed':
        return jsonify({
            'status': job['status'],
            'progress': job['progress'],
            'message': job['message'],
            'download_filename': job['output_filename'],
            'file_id': job_id
        })
    elif job['status'] == 'error':
        return jsonify({
            'status': job['status'],
            'error': job.get('error', '알 수 없는 오류')
        })
    else:
        return jsonify({
            'status': job['status'],
            'progress': job['progress'],
            'message': job['message']
        })

@app.route('/download/<file_id>/<filename>', methods=['GET'])
def download_file(file_id, filename):
    try:
        from flask import send_file
        import os
        
        file_path = os.path.join(os.getcwd(), filename)
        
        if not os.path.exists(file_path):
            return jsonify({'error': '파일을 찾을 수 없습니다.'}), 404
        
        def remove_file():
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
            except:
                pass
        
        response = send_file(file_path, as_attachment=True, download_name=filename.replace(f'translated_{file_id}_', ''))
        
        # 파일 전송 후 정리 (1초 후)
        import threading
        timer = threading.Timer(1.0, remove_file)
        timer.start()
        
        return response
        
    except Exception as e:
        print(f"파일 다운로드 오류: {e}")
        return jsonify({'error': f'파일 다운로드 중 오류가 발생했습니다: {str(e)}'}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'message': '번역 서버가 정상 작동 중입니다.'})

if __name__ == '__main__':
    print("번역 서버 시작 중...")
    print("URL: http://localhost:5001")
    print("중지하려면 Ctrl+C를 누르세요.")
    app.run(host='0.0.0.0', port=5001, debug=True)