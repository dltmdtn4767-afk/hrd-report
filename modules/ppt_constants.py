"""
ppt_constants.py
Master styling constants for PPT generation.
Syncs with EXC internal report standards.
"""
from pptx.util import Cm
from pptx.dml.color import RGBColor

# 폰트 설정 (Pretendard -> 나눔바른고딕 Light로 강제 고정)
FONT_NAME = "나눔바른고딕 Light"

# 브랜드 컬러 세팅 (Hex to RGB)
BRAND_COLORS = {
    "BLUE": RGBColor(37, 99, 235),     # 2563EB
    "SUCCESS": RGBColor(16, 185, 129), # 10B981
    "DANGER": RGBColor(239, 68, 68),   # EF4444
    "GRID": RGBColor(226, 232, 240)    # E2E8F0
}

# 엑스퍼트컨설팅 제안/보고서용 픽스 사이즈
CHART_SIZE = {"width": Cm(25.43), "height": Cm(9.25)}
TABLE_SIZE = {"width": Cm(25.43), "height": Cm(2.99)}

# 차트 세부 설정
CHART_CONFIG = {
    "gap_width": 60,
    "label_format": "0.00" # 만족도 5.00 표시용
}
