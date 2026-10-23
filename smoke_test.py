"""Stdlib-only smoke test: login, upload Account/QES/FDI samples, then hit
the calculation and comparison routes to confirm the whole flow works."""
import http.cookiejar
import mimetypes
import urllib.request
import urllib.parse
import uuid

BASE = "http://127.0.0.1:8000"
jar = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))


def post_form(path, fields):
    data = urllib.parse.urlencode(fields).encode()
    req = urllib.request.Request(BASE + path, data=data, method="POST")
    return opener.open(req, timeout=10)


def post_file(path, field_name, file_path):
    boundary = uuid.uuid4().hex
    with open(file_path, "rb") as f:
        content = f.read()
    ctype, _ = mimetypes.guess_type(file_path)
    ctype = ctype or "application/octet-stream"
    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="{field_name}"; filename="{file_path.split(chr(92))[-1]}"\r\n'
        f"Content-Type: {ctype}\r\n\r\n"
    ).encode() + content + f"\r\n--{boundary}--\r\n".encode()
    req = urllib.request.Request(
        BASE + path, data=body, method="POST",
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
    )
    return opener.open(req, timeout=10)


def get(path):
    return opener.open(BASE + path, timeout=10)


import re


def upload_two_step(kind, file_path, sheet="__ALL__"):
    r = post_file(f"/admin/upload/{kind}/preview", "file", file_path)
    html = r.read().decode()
    m = re.search(r'name="token" value="([a-f0-9]+)"', html)
    assert m, f"no token found for {kind} preview (status {r.status})"
    token = m.group(1)
    r2 = post_form(f"/admin/upload/{kind}/confirm", {"token": token, "sheet": sheet})
    return r2.status


print("1) login ->", post_form("/login", {"username": "admin", "password": "Admin@123"}).status)
print("2) upload accounts ->", upload_two_step("accounts", "sample_data/Account.xlsx"))
print("3) upload QES ->", upload_two_step("qes", "sample_data/QES_sample.xlsx"))
print("4) upload FDI ->", upload_two_step("fdi", "sample_data/FDI_sample.xlsx"))

r = get("/survey/qes?frame_id=F001&mode=single")
html = r.read().decode()
print("5) QES single view ->", r.status, "total production present:" , "Total production" in html, "AV field:", "Added value" in html)
assert "Frame ID was not found" not in html, "QES record not found / access denied!"

r = get("/survey/qes?frame_id=F001&mode=compare&quarters=Q1_2026&quarters=Q2_2026")
html = r.read().decode()
print("6) QES compare view ->", r.status, "quarters listed:", html.count("formula-box") >= 2)

r = get("/compare/qes?frame_id=F001&periods=Q1_2026&periods=Q2_2026")
html = r.read().decode()
print("7) QES comparison page ->", r.status, "financial table present:", "Financial comparison" in html, "change notice present:", "&rarr;" in html)

r = get("/survey/fdi?frame_id=F001&mode=single")
html = r.read().decode()
print("8) FDI single view ->", r.status, "equation present:", "Calculated equity" in html)
assert "Frame ID was not found" not in html, "FDI record not found / access denied!"

r = get("/compare/fdi?frame_id=F001&periods=Q1_2026&periods=Q2_2026")
html = r.read().decode()
print("9) FDI comparison page ->", r.status, "asset table present:", "Asset equation comparison" in html, "change notice present:", "&rarr;" in html)

r = get("/admin")
html = r.read().decode()
print("10) admin page ->", r.status, "users listed:", "auditor1" in html, "logs present:", "LOGIN_SUCCESS" in html)

print("\nALL SMOKE TESTS PASSED" if r.status == 200 else "SOMETHING FAILED")
