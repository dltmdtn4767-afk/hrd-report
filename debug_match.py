"""매칭 디버그 — 새 알고리즘"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
from modules.config_manager import load_config
from modules.sample_analyzer import SampleAnalyzer
config = load_config()
sa = SampleAnalyzer(config)

summary = {"categories": 8, "num_instructors": 2, "has_modules": True}
best = sa.find_best_match(summary)
print(f"Best: {best['name']}")
print()

for p in sa.patterns:
    f = p["features"]
    c = p.get("counts", {})
    score = 0
    target_cats = 8
    cat_diff = abs(f.get("categories", 0) - target_cats)
    score += max(0, 15 - cat_diff * 3)
    sample_instr = min(f.get("instructors", 0), 5)
    instr_diff = abs(sample_instr - 2)
    score += max(0, 8 - instr_diff * 3)
    if f.get("has_modules") == True:
        score += 8
    if c.get("exec_slides", 0) > 1:
        score += 4
    if c.get("quant_slides", 0) > 1 and True:
        score += 4
    if c.get("qual_slides", 0) >= 1:
        score += 2
    total_diff = cat_diff + instr_diff
    score += max(0, 5 - total_diff)
    print(f"{score:2d}pt | cats={f.get('categories'):2d} i={f.get('instructors')} mod={f.get('has_modules')} e={c.get('exec_slides')} q={c.get('quant_slides')} | {p['name'][:50]}")
