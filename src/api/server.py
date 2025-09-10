#!/usr/bin/env python3
"""
번역 API 서버 - 리팩토링된 버전
"""

import os
import sys
import tempfile
import uuid
import threading
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

# 프로젝트 루트를 Python 경로에 추가
project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, project_root)

from src.core.excel_translator import ExcelTranslator

app = Flask(__name__)
CORS(app)

# 전역 변수
translation_jobs = {}
translator = ExcelTranslator()

class TranslationJob:
    def __init__(self, job_id: str):
        self.job_id = job_id
        self.status = "pending"
        self.progress = 0
        self.message = "대기 중..."
        self.result_file = None
        self.error = None
    
    def update_progress(self, message: str, progress: int):
        self.message = message
        self.progress = progress
        if progress >= 100:
            self.status = "completed"
        elif progress > 0:
            self.status = "processing"

@app.route('/api/health', methods=['GET'])
def health_check():
    """서버 상태 확인"""
    return jsonify({"status": "ok", "message": "Translation server is running"})

@app.route('/api/translate', methods=['POST'])
def translate_excel():
    """Excel 파일 번역 시작"""
    try:
        # 파일 확인
        if 'file' not in request.files:
            return jsonify({"error": "파일이 없습니다"}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "파일이 선택되지 않았습니다"}), 400
        
        # 번역 옵션 가져오기
        direction = request.form.get('direction', 'ko_to_zh')
        preserve_english = request.form.get('preserve_english', 'true').lower() == 'true'
        add_new_sheet = request.form.get('add_new_sheet', 'true').lower() == 'true'
        
        # 제외 옵션
        exclude_cells = set(request.form.get('exclude_cells', '').split(',')) if request.form.get('exclude_cells') else set()
        exclude_sheets = set(request.form.get('exclude_sheets', '').split(',')) if request.form.get('exclude_sheets') else set()
        exclude_patterns = set(request.form.get('exclude_patterns', '').split(',')) if request.form.get('exclude_patterns') else set()
        
        # 빈 문자열 제거
        exclude_cells = {cell.strip() for cell in exclude_cells if cell.strip()}
        exclude_sheets = {sheet.strip() for sheet in exclude_sheets if sheet.strip()}
        exclude_patterns = {pattern.strip() for pattern in exclude_patterns if pattern.strip()}
        
        # 작업 ID 생성
        job_id = str(uuid.uuid4())
        job = TranslationJob(job_id)
        translation_jobs[job_id] = job
        
        # 임시 파일 생성
        temp_dir = tempfile.gettempdir()
        input_filename = secure_filename(file.filename)
        input_path = os.path.join(temp_dir, f"{job_id}_input_{input_filename}")
        output_path = os.path.join(temp_dir, f"{job_id}_output_{input_filename}")
        
        # 입력 파일 저장
        file.save(input_path)
        
        # 백그라운드에서 번역 실행
        thread = threading.Thread(
            target=_perform_translation,
            args=(job, input_path, output_path, direction, preserve_english, 
                  add_new_sheet, exclude_cells, exclude_sheets, exclude_patterns)
        )
        thread.daemon = True
        thread.start()
        
        return jsonify({
            "job_id": job_id,
            "status": "started",
            "message": "번역이 시작되었습니다"
        })
        
    except Exception as e:
        return jsonify({"error": f"번역 시작 오류: {str(e)}"}), 500

def _perform_translation(job, input_path, output_path, direction, preserve_english, 
                        add_new_sheet, exclude_cells, exclude_sheets, exclude_patterns):
    """백그라운드에서 실제 번역 수행"""
    try:
        # 진행률 콜백 함수
        def progress_callback(message, progress):
            job.update_progress(message, progress)
        
        # 번역기 생성 및 실행
        translator = ExcelTranslator(progress_callback)
        success = translator.translate_file(
            input_path=input_path,
            output_path=output_path,
            direction=direction,
            preserve_english=preserve_english,
            add_new_sheet=add_new_sheet,
            exclude_cells=exclude_cells,
            exclude_sheets=exclude_sheets,
            exclude_patterns=exclude_patterns
        )
        
        if success:
            job.result_file = output_path
            job.status = "completed"
            job.progress = 100
            job.message = "번역이 완료되었습니다!"
        else:
            job.status = "failed"
            job.error = "번역 중 오류가 발생했습니다"
        
        # 입력 파일 정리
        if os.path.exists(input_path):
            os.remove(input_path)
            
    except Exception as e:
        job.status = "failed"
        job.error = str(e)
        job.message = f"번역 실패: {str(e)}"

@app.route('/api/progress/<job_id>', methods=['GET'])
def get_progress(job_id):
    """번역 진행률 조회"""
    if job_id not in translation_jobs:
        return jsonify({"error": "작업을 찾을 수 없습니다"}), 404
    
    job = translation_jobs[job_id]
    return jsonify({
        "job_id": job_id,
        "status": job.status,
        "progress": job.progress,
        "message": job.message,
        "error": job.error
    })

@app.route('/api/download/<job_id>', methods=['GET'])
def download_result(job_id):
    """번역 결과 파일 다운로드"""
    if job_id not in translation_jobs:
        return jsonify({"error": "작업을 찾을 수 없습니다"}), 404
    
    job = translation_jobs[job_id]
    if job.status != "completed" or not job.result_file:
        return jsonify({"error": "번역이 완료되지 않았습니다"}), 400
    
    if not os.path.exists(job.result_file):
        return jsonify({"error": "결과 파일을 찾을 수 없습니다"}), 404
    
    # 다운로드 후 임시 파일 정리를 위한 콜백
    def cleanup_file():
        try:
            if os.path.exists(job.result_file):
                os.remove(job.result_file)
            del translation_jobs[job_id]
        except:
            pass
    
    # 5분 후 정리
    cleanup_timer = threading.Timer(300, cleanup_file)
    cleanup_timer.start()
    
    return send_file(
        job.result_file,
        as_attachment=True,
        download_name=f"translated_{os.path.basename(job.result_file)}"
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=False)