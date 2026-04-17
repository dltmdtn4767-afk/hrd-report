from pptx import Presentation
from pptx.util import Pt
import json

prs = Presentation('templates/[결과보고서] 템플렛.pptx')
slides = []
for i, slide in enumerate(prs.slides):
    shapes = []
    for sh in slide.shapes:
        info = {'name': sh.name, 'type': str(sh.shape_type)}
        if sh.has_text_frame:
            txt = ' | '.join(p.text for p in sh.text_frame.paragraphs if p.text.strip())
            info['text'] = txt[:100]
        try:
            info['chart_type'] = str(sh.chart.chart_type)
        except Exception:
            pass
        shapes.append(info)
    slides.append({'slide': i+1, 'layout': slide.slide_layout.name[:40], 'shapes': shapes})

with open('template_analysis.json', 'w', encoding='utf-8') as f:
    json.dump(slides, f, ensure_ascii=False, indent=2)
print(f'슬라이드 수: {len(slides)}')
for s in slides:
    print(f'\n[슬라이드 {s["slide"]}] 레이아웃: {s["layout"]}')
    for sh in s['shapes']:
        t = sh.get('text','')[:60]
        c = sh.get('chart_type','')
        print(f'  - {sh["name"]} ({sh["type"]}) {t} {c}')
