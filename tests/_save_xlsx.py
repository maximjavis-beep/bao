import base64, json, urllib.request
pdf_path = '/Users/streiten/Downloads/QXB20260327-1.pdf'
with open(pdf_path, 'rb') as f: b64 = base64.b64encode(f.read()).decode()
body = json.dumps({'files': [{'name': 'test.pdf', 'data': b64}]}).encode()
req = urllib.request.Request('http://127.0.0.1:8888/api/customs-summary', data=body, headers={'Content-Type': 'application/json'})
data = json.loads(urllib.request.urlopen(req).read().decode())
with open('/Users/streiten/customs/bao/tests/_verify.xlsx', 'wb') as f:
    f.write(base64.b64decode(data['file_b64']))
print('OK, rows:', len(data.get('items',[])))
