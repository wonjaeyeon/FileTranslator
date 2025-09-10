#!/usr/bin/env python3
"""
엑셀 파일 번역기 - xlwings 사용
실제 Excel 애플리케이션을 사용하여 100% 서식 보존
"""

import xlwings as xw
import os
import re
import requests
import time

class ExcelTranslatorXlwings:
    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback or (lambda msg, pct: None)
        
        # 발주서 전용 번역 사전
        self.ko_to_zh_dict = {
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
        }
        
        # 중국어 → 한국어 사전
        self.zh_to_ko_dict = {v: k for k, v in self.ko_to_zh_dict.items()}

    def is_korean(self, text):
        return bool(re.search('[가-힣]', str(text)))

    def is_chinese(self, text):
        return bool(re.search('[\u4e00-\u9fff]', str(text)))

    def is_english(self, text):
        return bool(re.match(r'^[a-zA-Z0-9\s\.\,\-\(\)\[\]\{\}@:\/]+$', str(text).strip()))

    def translate_text(self, text, direction, preserve_english=True):
        if not text or not isinstance(text, str) or not str(text).strip():
            return text
        
        text_str = str(text).strip()
        
        if preserve_english and self.is_english(text_str):
            return text
        
        if direction == 'ko-zh' and not self.is_korean(text_str):
            return text
        elif direction == 'zh-ko' and not self.is_chinese(text_str):
            return text
        
        # 번역 사전 사용
        dictionary = self.ko_to_zh_dict if direction == 'ko-zh' else self.zh_to_ko_dict
        
        translated = text_str
        for original, translation in dictionary.items():
            translated = translated.replace(original, translation)
        
        return translated

    def translate_excel_file(self, input_path, output_path, direction='ko-zh', preserve_english=True, add_new_sheet=True):
        """xlwings를 사용한 엑셀 파일 번역 (100% 서식 보존)"""
        try:
            self.progress_callback("Excel 애플리케이션 시작 중...", 0)
            
            # Excel 애플리케이션 시작 (보이지 않게)
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            
            self.progress_callback("원본 파일 열기 중...", 5)
            
            # 절대 경로로 변환
            input_path = os.path.abspath(input_path)
            output_path = os.path.abspath(output_path)
            
            # 워크북 열기
            wb = app.books.open(input_path)
            
            total_sheets = len(wb.sheets)
            
            for sheet_idx, sheet in enumerate(wb.sheets):
                sheet_name = sheet.name
                sheet_progress = (sheet_idx / total_sheets) * 80 + 10
                
                if add_new_sheet:
                    # 새 시트 생성 (원본 시트 복사)
                    translated_sheet_name = f"{sheet_name}_中文" if direction == 'ko-zh' else f"{sheet_name}_한국어"
                    
                    self.progress_callback(f"시트 '{sheet_name}' 복사 중...", sheet_progress)
                    
                    # 시트 전체 복사
                    sheet.copy(after=sheet)
                    copied_sheet = wb.sheets[wb.sheets.count - 1]  # 마지막 시트가 복사된 시트
                    copied_sheet.name = translated_sheet_name
                    
                    # 복사된 시트에서 번역 작업
                    self.translate_worksheet_xlwings(copied_sheet, direction, preserve_english, sheet_idx, total_sheets)
                else:
                    # 원본 시트에서 직접 번역
                    self.translate_worksheet_xlwings(sheet, direction, preserve_english, sheet_idx, total_sheets)
            
            self.progress_callback("파일 저장 중...", 90)
            
            # 파일 저장
            wb.save(output_path)
            wb.close()
            
            self.progress_callback("Excel 종료 중...", 95)
            app.quit()
            
            self.progress_callback("번역 완료!", 100)
            return output_path
            
        except Exception as e:
            # Excel 정리
            try:
                if 'wb' in locals():
                    wb.close()
                if 'app' in locals():
                    app.quit()
            except:
                pass
            raise e

    def translate_worksheet_xlwings(self, sheet, direction, preserve_english, sheet_idx, total_sheets):
        """xlwings 워크시트 번역"""
        try:
            # 사용된 범위 가져오기
            used_range = sheet.used_range
            if not used_range:
                return
            
            # 모든 셀의 값을 한번에 가져오기 (성능 최적화)
            values = used_range.value
            
            if values is None:
                return
            
            # 단일 셀인 경우 리스트로 변환
            if not isinstance(values, list):
                values = [[values]]
            elif isinstance(values[0], (str, int, float)) or values[0] is None:
                values = [values]
            
            total_cells = len(values) * len(values[0]) if values else 0
            current_cell = 0
            
            # 번역된 값들을 저장할 배열
            translated_values = []
            
            for row_idx, row in enumerate(values):
                translated_row = []
                
                for col_idx, cell_value in enumerate(row):
                    current_cell += 1
                    
                    # 20셀마다 진행률 업데이트
                    if current_cell % 20 == 0 or current_cell == total_cells:
                        sheet_base = (sheet_idx / total_sheets) * 70 + 10
                        cell_progress = (current_cell / total_cells) * (70 / total_sheets)
                        total_progress = min(85, sheet_base + cell_progress)
                        self.progress_callback(f"시트 번역 중... ({current_cell}/{total_cells} 셀)", total_progress)
                    
                    # 텍스트인 경우에만 번역
                    if cell_value and isinstance(cell_value, str):
                        translated_value = self.translate_text(cell_value, direction, preserve_english)
                        translated_row.append(translated_value)
                    else:
                        translated_row.append(cell_value)
                
                translated_values.append(translated_row)
            
            # 번역된 값들을 한번에 시트에 적용 (성능 최적화)
            if translated_values:
                used_range.value = translated_values
                
        except Exception as e:
            print(f"워크시트 번역 오류: {e}")
            raise e

if __name__ == "__main__":
    def test_progress(msg, pct):
        print(f"[{pct:5.1f}%] {msg}")
    
    translator = ExcelTranslatorXlwings(test_progress)
    
    # 테스트
    input_file = "sample-purchase-order.xlsx"
    output_file = "sample-purchase-order_xlwings.xlsx"
    
    if os.path.exists(input_file):
        try:
            result = translator.translate_excel_file(
                input_path=input_file,
                output_path=output_file,
                direction='ko-zh',
                preserve_english=True,
                add_new_sheet=True
            )
            print(f"번역 완료: {result}")
        except Exception as e:
            print(f"번역 오류: {e}")
            print("Excel이 설치되지 않았거나 xlwings 설정에 문제가 있을 수 있습니다.")
    else:
        print(f"파일을 찾을 수 없습니다: {input_file}")