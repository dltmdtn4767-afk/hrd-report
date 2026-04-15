"""다양한 엑셀 파일 테스트"""
import sys, glob, os
sys.stdout.reconfigure(encoding='utf-8')
from modules.data_loader import load_from_excel

# uploads 폴더의 고유 파일들 찾기
seen = set()
test_files = []
for d in sorted(glob.glob('uploads/*/')):
    files = glob.glob(os.path.join(d, '*.xlsx'))
    if not files: continue
    bn = os.path.basename(files[0])
    if bn not in seen:
        seen.add(bn)
        test_files.append(files[0])

# 직접 올린 파일도
for f in glob.glob('uploads/*.xlsx'):
    bn = os.path.basename(f)
    if bn not in seen:
        seen.add(bn)
        test_files.append(f)

print(f"테스트 파일: {len(test_files)}개\n")

for f in test_files:
    bn = os.path.basename(f)
    try:
        data = load_from_excel(f)
        ci = data["course_info"]
        print(f"OK | {bn[:50]}")
        print(f"   고객: {ci['company']} | 과정: {ci['course_name']}")
        print(f"   문항: {len(data['questions'])}개 | 카테고리: {len(data['categories'])}개 | 응답: {data['response_count']}명 | 평균: {data['overall_average']}")
        print(f"   주관식: {len(data['open_ended'])}개")
    except Exception as e:
        print(f"FAIL | {bn[:50]}")
        print(f"   에러: {e}")
    print()
