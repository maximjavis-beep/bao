"""测试 /api/customs-summary"""
import base64, json, urllib.request, sys

pdf_path = '/Users/streiten/Downloads/QXB20260327-1.pdf'
with open(pdf_path, 'rb') as f:
    b64 = base64.b64encode(f.read()).decode()

body = json.dumps({'files': [{'name': 'QXB20260327-1.pdf', 'data': b64}]}).encode()
req = urllib.request.Request('http://127.0.0.1:8888/api/customs-summary', data=body,
    headers={'Content-Type': 'application/json'})

try:
    resp = urllib.request.urlopen(req, timeout=10)
    data = json.loads(resp.read().decode())
    print('success:', data.get('success'))
    print('ok_count:', data.get('ok_count'))
    print('stats:', json.dumps(data.get('stats'), indent=2, ensure_ascii=False))
    items = data.get('items', [])
    for it in items:
        print('---')
        for k, v in it.items():
            print(f'  {k}: {v}')
    print('file_b64 length:', len(data.get('file_b64', '')))
except Exception as e:
    print(f'ERROR: {e}', file=sys.stderr)
    sys.exit(1)
