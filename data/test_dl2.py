import base64, json, os, urllib.request
from urllib.parse import quote
os.chdir('/Users/streiten/customs/bao')
inv_b64 = base64.b64encode(open('data/示例发票.xlsx','rb').read()).decode()
req = urllib.request.Request('http://127.0.0.1:7777/api/parse', data=json.dumps({'invoice': inv_b64}).encode(), headers={'Content-Type':'application/json'})
d = json.loads(urllib.request.urlopen(req).read())
req2 = urllib.request.Request('http://127.0.0.1:7777/api/weave', data=json.dumps({'invoice': d['invoice']['path'], 'shipper': '测试'}).encode(), headers={'Content-Type':'application/json'})
d2 = json.loads(urllib.request.urlopen(req2).read())
url = 'http://127.0.0.1:7777/api/download?path=' + quote(d2['download_path'])
resp = urllib.request.urlopen(url)
data = resp.read()
print('dl size', len(data), 'bytes OK' if len(data)>1000 else 'FAIL')
open('/tmp/test_dl.xlsx','wb').write(data)
print('saved')
