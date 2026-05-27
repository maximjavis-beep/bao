import base64, json, os, urllib.request
os.chdir('/Users/streiten/customs/bao')
inv_b64 = base64.b64encode(open('data/示例发票.xlsx','rb').read()).decode()
d = json.loads(urllib.request.urlopen(urllib.request.Request('http://127.0.0.1:7777/api/parse', data=json.dumps({'invoice': inv_b64}).encode(), headers={'Content-Type': 'application/json'})).read())
print('parse:', d['success'])
d2 = json.loads(urllib.request.urlopen(urllib.request.Request('http://127.0.0.1:7777/api/weave', data=json.dumps({'invoice': d['invoice']['path'], 'shipper': 'x'}).encode(), headers={'Content-Type': 'application/json'})).read())
print('weave:', d2['success'])
from urllib.parse import quote
resp = urllib.request.urlopen('http://127.0.0.1:7777/api/download?path=' + quote(d2['download_path']))
data = resp.read()
print('dl:', len(data), 'OK' if len(data) > 100 else 'FAIL')
