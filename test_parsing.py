#!/usr/bin/env python3

# GPT 응답 파싱 테스트
gpt_response = """<Sheet1!A1, Payhere 软件基本需求列表> -> Payhere 소프트웨어 기본 요구사항 목록
<Sheet1!A2, 操作系统(안드로이드버전)> -> 운영체제(안드로이드버전)
<Sheet1!C2, 默认语言(기본언어)> -> 기본 언어(기본언어)"""

translation_map = {}
lines = gpt_response.strip().split('\n')

print(f"GPT 응답 파싱 시작. 총 {len(lines)}줄")
print(f"GPT 응답 샘플: {gpt_response[:200]}...")

for i, line in enumerate(lines):
    line = line.strip()
    if not line:
        continue

    print(f"라인 {i+1}: {line}")

    # 다양한 형식 지원
    if ' -> ' in line or '→' in line or ':' in line:
        try:
            # 다양한 구분자 시도
            separator = ' -> '
            if '→' in line:
                separator = '→'
            elif ' -> ' not in line and ':' in line:
                separator = ':'

            parts = line.split(separator, 1)
            if len(parts) != 2:
                continue

            left_part = parts[0].strip()
            translated_text = parts[1].strip()

            # <셀주소, 원본텍스트> 형식
            if left_part.startswith('<') and left_part.endswith('>'):
                inner = left_part[1:-1]
                if ', ' in inner:
                    cell_address, original_text = inner.split(', ', 1)
                    translation_map[cell_address] = translated_text
                    print(f"  → 번역 매핑: {cell_address} = {translated_text}")

            # 단순한 형식도 지원: 셀주소, 원본 -> 번역
            elif ', ' in left_part and not left_part.startswith('<'):
                cell_address, original_text = left_part.split(', ', 1)
                translation_map[cell_address] = translated_text
                print(f"  → 번역 매핑 (단순): {cell_address} = {translated_text}")

            # 셀주소만 있는 경우: Sheet1!A1 -> 번역
            elif '!' in left_part and ', ' not in left_part:
                cell_address = left_part
                translation_map[cell_address] = translated_text
                print(f"  → 번역 매핑 (셀주소만): {cell_address} = {translated_text}")

        except Exception as e:
            print(f"라인 파싱 오류: {line} - {e}")
            continue

print(f"\n파싱된 번역 맵: {len(translation_map)}개 항목")
for addr, trans in translation_map.items():
    print(f"  {addr}: {trans}")