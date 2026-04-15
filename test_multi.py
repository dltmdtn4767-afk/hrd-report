"""멀티시트 + 공통응답 그룹핑 테스트"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
from modules.data_loader import load_all_sheets, get_available_sheets
from modules.response_grouper import group_responses_rule_based
import glob, os

# 테스트 파일 찾기
files = glob.glob('uploads/**/*.xlsx', recursive=True)
unique = {}
for f in files:
    bn = os.path.basename(f)
    if bn not in unique:
        unique[bn] = f

print(f"파일 {len(unique)}개 테스트\n")

for bn, path in list(unique.items())[:3]:
    print(f"=== {bn[:60]} ===")
    sheets = get_available_sheets(path)
    print(f" 시트: {sheets}")
    
    result = load_all_sheets(path)
    print(f" 멀티세션: {result['multi_session']}, 세션 수: {len(result['sessions'])}")
    
    for s in result['sessions']:
        print(f"  [{s.get('session_label','')}] 응답:{s['response_count']}명 / 평균:{s['overall_average']} / 문항:{len(s['questions'])}개")
    
    combined = result['combined']
    print(f" 종합: 총 {combined['response_count']}명, 평균 {combined['overall_average']}")
    
    # 공통응답 테스트
    for oe in combined.get('open_ended', []):
        if oe.get('answers'):
            groups = group_responses_rule_based(oe['answers'])
            print(f" 주관식[{oe['id']}] {len(oe['answers'])}건 → {len(groups)}그룹:")
            for g in groups[:3]:
                print(f"   [{g['common_id']}] {g['label'][:40]} ({g['count']}건)")
    print()
