import requests, glob, os, zipfile

BASE = "https://hrd-report.onrender.com"

r = requests.get(f"{BASE}/api/status")
print("Status:", r.json())

xlsx = min(glob.glob(r"C:\Users\User\Desktop\업무폴더\*.xlsx"), key=os.path.getsize)
print(f"Uploading: {os.path.basename(xlsx)}")
with open(xlsx, "rb") as f:
    r = requests.post(f"{BASE}/api/upload", files={"file": f}, timeout=120)
print(f"Upload: {r.status_code}")
if r.status_code != 200:
    print(f"ERROR: {r.text[:300]}")
    exit(1)
d = r.json()
sid = d["session_id"]
print(f"Session: {sid}")

payload = {
    "slides": [
        {"type":"cover","data":{"company":"test","course":"test"}},
        {"type":"summary","data":{}},
        {"type":"qual_text","data":{}},
        {"type":"back_cover","data":{}}
    ],
    "quant_groups": [],
    "qual_data": []
}
r = requests.post(f"{BASE}/api/build_ppt/{sid}", json=payload, timeout=60)
print(f"PPT: {r.status_code} type={r.headers.get('content-type','?')} size={len(r.content)}")
if r.status_code == 200:
    with open("render_test.pptx", "wb") as f:
        f.write(r.content)
    print(f"Valid PPTX: {zipfile.is_zipfile('render_test.pptx')}")
else:
    print(f"ERROR: {r.text[:500]}")
