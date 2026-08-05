import http.cookiejar, urllib.request, urllib.parse, mimetypes, uuid, re

BASE = 'http://127.0.0.1:8000'
jar = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))


def post_form(path, fields):
    data = urllib.parse.urlencode(fields).encode()
    return opener.open(urllib.request.Request(BASE + path, data=data, method='POST'), timeout=10)


def post_file(path, field_name, file_path):
    boundary = uuid.uuid4().hex
    with open(file_path, 'rb') as f:
        content = f.read()
    ctype, _ = mimetypes.guess_type(file_path)
    ctype = ctype or 'application/octet-stream'
    body = (f'--{boundary}\r\nContent-Disposition: form-data; name="{field_name}"; filename="{file_path}"\r\nContent-Type: {ctype}\r\n\r\n').encode() + content + f'\r\n--{boundary}--\r\n'.encode()
    req = urllib.request.Request(BASE + path, data=body, method='POST', headers={'Content-Type': f'multipart/form-data; boundary={boundary}'})
    return opener.open(req, timeout=10)


def upload_two_step(kind, path, sheet='__ALL__'):
    r = post_file(f'/admin/upload/{kind}/preview', 'file', path)
    html = r.read().decode()
    m = re.search(r'name="token" value="([a-f0-9]+)"', html)
    token = m.group(1)
    return post_form(f'/admin/upload/{kind}/confirm', {'token': token, 'sheet': sheet}).status


post_form('/login', {'username': 'admin', 'password': 'Admin@123'})
print('upload QES dummy ->', upload_two_step('qes', 'sample_data/QES_dummy_20x4Q.xlsx'))
print('upload FISQ dummy ->', upload_two_step('fisq', 'sample_data/FISQ_dummy_20x4Q.xlsx'))

r = opener.open(BASE + '/survey/qes?frame_id=F106', timeout=10)
html = r.read().decode()
print('F106 4-quarter calc, formula boxes:', html.count('formula-box'))

r = opener.open(BASE + '/compare/qes?frame_id=F106', timeout=10)
html = r.read().decode()
print('F106 compare rows:', html.count('<tr>'), '| trend colors:', html.count('trend-up') + html.count('trend-down'))
print('inventory mismatch red shown (F106 idx5):', 'Mismatch' in html)

r = opener.open(BASE + '/compare/fisq?frame_id=F104', timeout=10)
html = r.read().decode()
print('F104 FISQ needs review shown:', 'Needs review' in html)
