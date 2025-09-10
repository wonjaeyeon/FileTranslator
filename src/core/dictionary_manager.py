#!/usr/bin/env python3
"""
사전 관리자 - 번역 사전을 로드하고 관리하는 모듈
"""

import json
import os
from typing import Dict, Any

class DictionaryManager:
    def __init__(self, config_dir: str = "config"):
        self.config_dir = config_dir
        self.translation_dict = {}
        self.custom_dict = {}
        self._load_dictionaries()
    
    def _load_dictionaries(self):
        """번역 사전들을 로드"""
        try:
            # 기본 번역 사전 로드
            main_dict_path = os.path.join(self.config_dir, "translation_dict.json")
            if os.path.exists(main_dict_path):
                with open(main_dict_path, 'r', encoding='utf-8') as f:
                    self.translation_dict = json.load(f)
            
            # 커스텀 사전 로드
            custom_dict_path = os.path.join(self.config_dir, "custom_dict.json")
            if os.path.exists(custom_dict_path):
                with open(custom_dict_path, 'r', encoding='utf-8') as f:
                    self.custom_dict = json.load(f)
        except Exception as e:
            print(f"사전 로드 오류: {e}")
    
    def get_translation_dict(self, direction: str) -> Dict[str, str]:
        """번역 방향에 따른 통합 사전 반환"""
        combined_dict = {}
        
        # 기본 사전에서 모든 카테고리의 단어들을 병합
        if direction in self.translation_dict:
            main_dict = self.translation_dict[direction]
            if isinstance(main_dict, dict):
                for category, words in main_dict.items():
                    if isinstance(words, dict):
                        combined_dict.update(words)
        
        # 커스텀 사전 병합 (우선순위 높음)
        if direction in self.custom_dict:
            custom_dict = self.custom_dict[direction]
            if isinstance(custom_dict, dict):
                for category, words in custom_dict.items():
                    if isinstance(words, dict):
                        combined_dict.update(words)
        
        return combined_dict
    
    def add_custom_translation(self, source: str, target: str, direction: str, category: str = "사용자_추가"):
        """커스텀 번역 추가"""
        if direction not in self.custom_dict:
            self.custom_dict[direction] = {}
        
        if category not in self.custom_dict[direction]:
            self.custom_dict[direction][category] = {}
        
        self.custom_dict[direction][category][source] = target
        
        # 파일에 저장
        self._save_custom_dict()
    
    def _save_custom_dict(self):
        """커스텀 사전을 파일에 저장"""
        try:
            custom_dict_path = os.path.join(self.config_dir, "custom_dict.json")
            with open(custom_dict_path, 'w', encoding='utf-8') as f:
                json.dump(self.custom_dict, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"커스텀 사전 저장 오류: {e}")