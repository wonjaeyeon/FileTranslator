#!/usr/bin/env python3
"""
번역 엔진 - Google Translate API와 사전 기반 번역을 조합
"""

import re
from typing import Optional, Callable
from googletrans import Translator
from .dictionary_manager import DictionaryManager

class TranslationEngine:
    def __init__(self, progress_callback: Optional[Callable] = None):
        self.progress_callback = progress_callback or (lambda msg, pct: None)
        self.google_translator = Translator()
        self.dict_manager = DictionaryManager()
        
        # 영문 패턴 (영어는 번역하지 않음)
        self.english_pattern = re.compile(r'^[a-zA-Z0-9\s\.,\-_/\\()&%$#@!+=]*$')
    
    def is_english_text(self, text: str) -> bool:
        """텍스트가 영문인지 확인"""
        if not text or not text.strip():
            return False
        return bool(self.english_pattern.match(text.strip()))
    
    def translate_with_dict(self, text: str, direction: str) -> Optional[str]:
        """사전 기반 번역 시도"""
        if not text or not text.strip():
            return None
        
        # 정확한 매칭 시도
        translation_dict = self.dict_manager.get_translation_dict(direction)
        exact_match = translation_dict.get(text.strip())
        if exact_match:
            return exact_match
        
        # 부분 매칭 시도 (더 긴 키워드부터)
        sorted_keys = sorted(translation_dict.keys(), key=len, reverse=True)
        for key in sorted_keys:
            if key in text:
                return text.replace(key, translation_dict[key])
        
        return None
    
    def translate_with_google(self, text: str, direction: str) -> Optional[str]:
        """Google Translate API 사용"""
        if not text or not text.strip():
            return None
        
        try:
            if direction == "ko_to_zh":
                result = self.google_translator.translate(text, src='ko', dest='zh-cn')
            elif direction == "zh_to_ko":
                result = self.google_translator.translate(text, src='zh-cn', dest='ko')
            else:
                return None
            
            return result.text if result else None
        except Exception as e:
            print(f"Google 번역 오류: {e}")
            return None
    
    def translate_text(self, text: str, direction: str, preserve_english: bool = True) -> str:
        """텍스트 번역 (사전 우선, Google Translate 보조)"""
        if not text or not text.strip():
            return text
        
        # 영문 보존 옵션이 켜져있고 영문인 경우
        if preserve_english and self.is_english_text(text):
            return text
        
        # 1. 사전 기반 번역 시도
        dict_result = self.translate_with_dict(text, direction)
        if dict_result:
            return dict_result
        
        # 2. Google Translate 사용
        google_result = self.translate_with_google(text, direction)
        if google_result:
            return google_result
        
        # 3. 번역 실패시 원본 반환
        return text