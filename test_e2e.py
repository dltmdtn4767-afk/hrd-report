"""E2E 테스트"""
import sys, glob, json
sys.stdout.reconfigure(encoding='utf-8')
import requests

print("=" * 50)

# 1. 서버
r = requests.get('http://localhost:8080')
print(f"1. 서버: {r.status_code}")

# 2. 샘플
r = requests.get('http://localhost:8080/api/samples')
patterns = r.json()
print(f"2. 샘플 패턴: {len(patterns)}개")
for p in patterns[:3]:
    print(f"   - {p['name']} ({p['slide_count']}슬라이드)")

# 3. 업로드
xlsx = glob.glob('uploads/*.xlsx')
print(f"\n3. 엑셀 파일: {len(xlsx)}개")
if xlsx:
    fname = xlsx[0]
    print(f"   테스트: {fname.split(chr(92))[-1]}")
    with open(fname, 'rb') as f:
        r = requests.post('http://localhost:8080/api/upload',
                          files={'file': (fname.split(chr(92))[-1], f)})
    data = r.json()
    print(f"   상태: {r.status_code}")
    
    if 'session_id' in data:
        sid = data['session_id']
        s = data['summary']
        ai = data['ai_result']
        print(f"   세션: {sid}")
        print(f"   고객사: {s['company']}")
        print(f"   과정: {s['course_name']}")
        print(f"   문항: {s['total_questions']}개, 카테고리: {s['categories']}개")
        print(f"   AI 패턴: {ai['matched_pattern']}")
        print(f"   권장 슬라이드: {len(ai['recommended_slides'])}개")
        for sl in ai['recommended_slides']:
            print(f"     - {sl.get('title', sl.get('type'))}")
        
        # 4. 생성
        print(f"\n4. 보고서 생성...")
        r2 = requests.post(f'http://localhost:8080/api/generate/{sid}')
        gen = r2.json()
        print(f"   결과: {'성공' if gen.get('success') else '실패'}")
        if gen.get('success'):
            print(f"   슬라이드: {gen['slide_count']}개")
            print(f"   검증 점수: {gen['review']['score']}점")
            for c in gen['review']['checks']:
                icon = "PASS" if c['status'] == 'pass' else "WARN"
                print(f"   [{icon}] {c['detail']}")
            
            # 5. 다운로드 테스트
            r3 = requests.get(f'http://localhost:8080/api/download/{sid}')
            print(f"\n5. 다운로드: {r3.status_code} ({len(r3.content)//1024}KB)")
        else:
            print(f"   에러: {gen.get('error')}")
    else:
        print(f"   에러: {json.dumps(data, ensure_ascii=False)}")

print("\n" + "=" * 50)
print("E2E 테스트 완료!")
