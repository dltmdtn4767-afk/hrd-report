"""
builder_generator.py
빌더 구성 → 실제 템플릿 기반 PPT 생성
"""
from pathlib import Path
from datetime import datetime
import copy

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE

BASE_DIR = Path(__file__).parent.parent
TEMPLATE_PATH = BASE_DIR / "templates" / "template.pptx"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# 템플릿 슬라이드 인덱스 (0-based)
TPL_COVER    = 0
TPL_TOC      = 1
TPL_SEC1     = 2   # Ⅰ 과정 개요
TPL_OVERVIEW = 3   # 과정 개요 표
TPL_SCHEDULE = 4   # 교육 일정표
TPL_SEC2     = 7   # Ⅱ 만족도 결과
TPL_SUMMARY  = 8   # Executive Summary
TPL_QUANT    = 9   # 정량 평가 결과 (복사용)
TPL_QUAL     = 10  # 정성 평가 결과
TPL_BACK     = 11  # 뒤표지


# ────────────────────────────────────────────
# 유틸
# ────────────────────────────────────────────
def _set_text(shape, text: str, bold=None, size=None, color=None, align=None):
    """텍스트박스/플레이스홀더 텍스트 교체"""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    # 전체 텍스트 첫 단락에 설정
    for i, para in enumerate(tf.paragraphs):
        for run in para.runs:
            run.text = text if i == 0 else ''
            if bold is not None: run.font.bold = bold
            if size: run.font.size = Pt(size)
            if color: run.font.color.rgb = RGBColor(*color)
            break
        if i == 0 and not para.runs:
            run = para.add_run()
            run.text = text
            if bold is not None: run.font.bold = bold
            if size: run.font.size = Pt(size)
            if color: run.font.color.rgb = RGBColor(*color)
        if align and i == 0:
            para.alignment = align
        if i > 0:
            for run in para.runs:
                run.text = ''


def _find_shape(slide, name_contains: str):
    for sh in slide.shapes:
        if name_contains.lower() in sh.name.lower():
            return sh
    return None


def _set_table_cell(table, row, col, text, bold=False, bg=None):
    try:
        cell = table.cell(row, col)
        tf = cell.text_frame
        for p in tf.paragraphs:
            for r in p.runs:
                r.text = ''
        if tf.paragraphs:
            run = tf.paragraphs[0].add_run()
            run.text = str(text)
            run.font.bold = bold
            if bg:
                from pptx.oxml.ns import qn
                from lxml import etree
                # fill solid
                tc = cell._tc
                tcPr = tc.find(qn('a:tcPr')) or etree.SubElement(tc, qn('a:tcPr'))
                solidFill = etree.SubElement(tcPr, qn('a:solidFill'))
                srgb = etree.SubElement(solidFill, qn('a:srgbClr'))
                srgb.set('val', bg.lstrip('#'))
    except Exception:
        pass


def _clone_slide(prs: Presentation, src_idx: int) -> object:
    """슬라이드 복제 (XML 복사)"""
    import copy
    from pptx.oxml.ns import qn
    from lxml import etree

    src_slide = prs.slides[src_idx]
    blank_layout = src_slide.slide_layout

    new_slide = prs.slides.add_slide(blank_layout)

    # XML 내용 복사
    new_slide.shapes._spTree.clear()
    for elem in copy.deepcopy(src_slide.shapes._spTree):
        new_slide.shapes._spTree.append(elem)

    return new_slide


# ────────────────────────────────────────────
# 슬라이드 타입별 채우기 함수
# ────────────────────────────────────────────
def _fill_cover(slide, data: dict):
    """앞표지: 고객사명 + 과정명"""
    for sh in slide.shapes:
        if sh.has_text_frame:
            txt = sh.text_frame.text
            if '고객사명' in txt:
                _set_text(sh, data.get('company', ''))
            elif '과정명' in txt:
                _set_text(sh, data.get('course', ''))


def _fill_overview(slide, data: dict):
    """과정 개요 표"""
    rows = data.get('rows', [])
    tbl_shape = _find_shape(slide, '표')
    if not tbl_shape or not tbl_shape.has_table:
        return
    tbl = tbl_shape.table
    for i, row in enumerate(rows):
        try:
            if i < tbl.rows.__len__():
                _set_table_cell(tbl, i, 0, row.get('key', ''), bold=True)
                _set_table_cell(tbl, i, 1, row.get('val', ''))
        except Exception:
            pass


def _fill_schedule(slide, data: dict):
    """교육 일정표"""
    rows = data.get('rows', [])
    tbl_shape = _find_shape(slide, '표')
    if not tbl_shape or not tbl_shape.has_table:
        return
    tbl = tbl_shape.table
    for i, row in enumerate(rows):
        try:
            r_idx = i + 1  # 헤더 제외
            if r_idx < tbl.rows.__len__():
                _set_table_cell(tbl, r_idx, 0, row.get('label', ''))
                _set_table_cell(tbl, r_idx, 1, row.get('date', ''))
                _set_table_cell(tbl, r_idx, 2, row.get('place', ''))
                _set_table_cell(tbl, r_idx, 3, row.get('count', ''))
        except Exception:
            pass


def _fill_summary(slide, data: dict):
    """Executive Summary — 차트 업데이트 + 표"""
    cats = data.get('categories', [])

    # 차트 업데이트
    for sh in slide.shapes:
        try:
            chart = sh.chart
            cd = ChartData()
            cd.categories = [c.get('name', '') for c in cats]
            cd.add_series('평균', [round(c.get('avg', 0), 2) for c in cats])
            chart.replace_data(cd)
            break
        except Exception:
            pass

    # 표 (전체 평균 / 응답 인원)
    tbl_shape = _find_shape(slide, '표')
    if tbl_shape and tbl_shape.has_table:
        tbl = tbl_shape.table
        try:
            overall = data.get('overall', 0)
            resp    = data.get('response_count', 0)
            total_q = sum(len(c.get('questions', [])) for c in cats)
            _set_table_cell(tbl, 1, 1, f"{overall:.2f}점")
            _set_table_cell(tbl, 1, 2, f"{resp}명")
            _set_table_cell(tbl, 1, 3, f"{total_q}개")
        except Exception:
            pass


def _fill_quant_chart(slide, group: dict):
    """정량 차트 슬라이드 (슬라이드 10 복제본)"""
    questions = group.get('questions', [])
    title     = group.get('title', '정량 평가 결과')

    # 슬라이드 제목
    title_sh = _find_shape(slide, '제목') or _find_shape(slide, 'TextBox 224')
    if title_sh:
        _set_text(title_sh, title, bold=True)

    # 표 채우기
    tbl_shape = _find_shape(slide, '표')
    if tbl_shape and tbl_shape.has_table:
        tbl = tbl_shape.table
        _set_table_cell(tbl, 0, 0, '번호', bold=True)
        _set_table_cell(tbl, 0, 1, '문항', bold=True)
        _set_table_cell(tbl, 0, 2, '평균', bold=True)
        for i, q in enumerate(questions):
            r = i + 1
            try:
                if r < tbl.rows.__len__():
                    _set_table_cell(tbl, r, 0, q.get('id', str(i+1)))
                    _set_table_cell(tbl, r, 1, q.get('label', ''))
                    _set_table_cell(tbl, r, 2, f"{q.get('avg',0):.2f}")
            except Exception:
                pass

    # 차트 (있을 경우)
    for sh in slide.shapes:
        try:
            chart = sh.chart
            cd = ChartData()
            cd.categories = [q.get('id', str(i+1)) for i, q in enumerate(questions)]
            cd.add_series('평균', [round(q.get('avg', 0), 2) for q in questions])
            chart.replace_data(cd)
            break
        except Exception:
            pass


def _fill_qual(slide, qual_data: list):
    """정성 평가 슬라이드 — 텍스트 그룹 채우기"""
    if not qual_data:
        return
    # 각 Q별 공통응답을 텍스트로 합침
    lines = []
    for oe in qual_data:
        qid   = oe.get('id', '')
        label = oe.get('label', '')
        lines.append(f"■ {qid} {label}")
        for g in oe.get('groups', []):
            lines.append(f"  · {g.get('label','')} ({g.get('count',0)}건)")
        lines.append('')

    full_text = '\n'.join(lines)

    # 텍스트박스에 삽입 (첫 번째 큰 텍스트박스)
    for sh in slide.shapes:
        if sh.has_text_frame and sh.width > Inches(4):
            _set_text(sh, full_text)
            break


# ────────────────────────────────────────────
# 메인 빌드 함수
# ────────────────────────────────────────────
def build_custom_ppt(slides: list, quant_groups: list, qual_data: list,
                     data: dict, config: dict, template_path: str = None) -> str:
    tpl = Path(template_path) if template_path and Path(template_path).exists() else TEMPLATE_PATH
    if not tpl.exists():
        raise FileNotFoundError(f"템플릿 없음: {tpl}")
    prs = Presentation(str(tpl))

    # ── 1. 템플릿 슬라이드 조작 ──────────────────
    # 슬라이드 순서를 빌더 구성대로 재구성
    # 전략: 템플릿을 기반으로 각 슬라이드 타입의 내용만 채움

    # 표지 채우기
    cover_data = next((s['data'] for s in slides if s['type'] == 'cover'), {})
    _fill_cover(prs.slides[TPL_COVER], cover_data)

    # 과정 개요 표
    ov_data = next((s['data'] for s in slides if s['type'] == 'overview'), {})
    _fill_overview(prs.slides[TPL_OVERVIEW], ov_data)

    # 교육 일정표
    sch_data = next((s['data'] for s in slides if s['type'] == 'schedule'), {})
    _fill_schedule(prs.slides[TPL_SCHEDULE], sch_data)

    # Executive Summary
    summary_data = next((s['data'] for s in slides if s['type'] == 'summary'), {})
    if not summary_data.get('categories'):
        summary_data['categories'] = data.get('categories', [])
    _fill_summary(prs.slides[TPL_SUMMARY], summary_data)

    # 정성 평가
    _fill_qual(prs.slides[TPL_QUAL], qual_data)

    # ── 2. 정량 그룹 슬라이드 (슬라이드 10 복제) ──
    # 슬라이드 10 (TPL_QUANT) 을 각 그룹마다 복제 삽입
    # 기존 슬라이드 10은 첫 그룹으로 사용, 나머지는 새로 삽입

    if quant_groups:
        # 첫 그룹 → 기존 슬라이드 10 채우기
        _fill_quant_chart(prs.slides[TPL_QUANT], quant_groups[0])

        # 추가 그룹 → 슬라이드 복제 후 TPL_QUAL 앞에 삽입
        for g in quant_groups[1:]:
            new_slide = _clone_slide(prs, TPL_QUANT)
            _fill_quant_chart(new_slide, g)
            # 슬라이드를 TPL_QUAL 위치로 이동
            xml_slides = prs.slides._sldIdLst
            # 마지막에 추가된 슬라이드를 정성 슬라이드 앞으로 이동
            # (현재 마지막 = 방금 추가한 슬라이드)
            last = xml_slides[-1]
            qual_xml_idx = len(prs.slides) - 2  # 정성은 뒤에서 2번째
            xml_slides.remove(last)
            xml_slides.insert(qual_xml_idx, last)

    # ── 3. 파일명 & 저장 ─────────────────────────
    company  = cover_data.get('company', '')
    course   = cover_data.get('course', '')
    today    = datetime.now().strftime('%Y%m%d')
    filename = f"[결과보고서] {company}_{course}_{today}.pptx"
    out_path = OUTPUT_DIR / filename
    prs.save(str(out_path))
    return str(out_path)
