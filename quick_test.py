import requests, os, glob
BASE = "http://localhost:8080"
xlsx = min(glob.glob(r"C:\Users\User\Desktop\업무폴더\*.xlsx"), key=os.path.getsize)
with open(xlsx, "rb") as f:
    r = requests.post(f"{BASE}/api/upload", files={"file": f})
d = r.json()
sid = d["session_id"]
s = d["summary"]
print("Upload OK:", sid, s.get("course_name"), "Q:", s.get("total_questions"))
payload = {"slides": [{"type":"cover","data":{}},{"type":"summary","data":{}},{"type":"back_cover","data":{}}], "quant_groups": [], "qual_data": []}
r = requests.post(f"{BASE}/api/build_ppt/{sid}", json=payload)
print("PPT:", r.status_code, "size:", len(r.content))
for f in ["custom_slides.js","custom_slides.css","builder.js","app.js"]:
    r = requests.get(f"{BASE}/static/{f}")
    print(f"{f}: {r.status_code} ({len(r.text)} chars)")
print("ALL OK")
