"""
config_manager.py
외부 config.json 파일에서 설정을 로드하고 관리하는 모듈
"""
import json
import os
import sys


def get_base_path():
    """PyInstaller로 빌드된 exe에서도 올바른 경로를 반환"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_config(config_path=None):
    """config.json 파일을 로드하여 딕셔너리로 반환"""
    if config_path is None:
        config_path = os.path.join(get_base_path(), "config.json")

    if not os.path.exists(config_path):
        raise FileNotFoundError(f"설정 파일을 찾을 수 없습니다: {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    return config


def save_config(config, config_path=None):
    """설정을 config.json 파일로 저장"""
    if config_path is None:
        config_path = os.path.join(get_base_path(), "config.json")

    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


def validate_config(config):
    """설정 파일의 필수 키를 검증"""
    required_keys = ["data_source", "survey_config", "ppt_template", "course_info", "email"]
    missing = [key for key in required_keys if key not in config]
    if missing:
        raise ValueError(f"설정 파일에 필수 키가 누락되었습니다: {', '.join(missing)}")
    return True


def resolve_path(relative_path):
    """상대 경로를 절대 경로로 변환 (exe 환경 대응)"""
    base = get_base_path()
    return os.path.join(base, relative_path)
