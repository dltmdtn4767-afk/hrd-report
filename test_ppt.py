import requests, sys, glob
sys.stdout.reconfigure(encoding='utf-8')

files = glob.glob('uploads/**/*.xlsx', recursive=True)
if not files:
    print('업로드 파일 없음')
    exit()

# 업로드
with open(files[0], 'rb') as f:
    r = requests.post('http://localhost:8080/api/upload',
                      files={'file': (files[0].split('\\')[-1], f)})
sid = r.json().get('session_id', '')
print(f'업로드 완료: {sid}')

# PPT 생성
r2 = requests.post(f'http://localhost:8080/api/generate/{sid}')
d = r2.json()
print(f'상태코드: {r2.status_code}')
print(f'성공: {d.get("success")}')
print(f'파일: {d.get("output_file", "없음")}')
print(f'슬라이드: {d.get("slide_count", 0)}건')
if not d.get("success"):
    print(f'오류: {d.get("error", "없음")}')
