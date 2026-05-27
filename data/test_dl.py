import base64, json, os, urllib.request
from urllib.parse import quote
os.chdir('/Users/streiten/customs/bao')
inv_b64 = base64.b64encode(open('data/示例发票.xlsx','rb').read()).decode()
req = urllib.request.Request('http://127.0.0.1:7777/api/parse',
    data=json.dumps({'invoice': inv_b64}).encode(),
    headers={'Content-Type': 'application/json'})
d = json.loads(urllib.request.urlopen(req).read())
inv_path = d['invoice']['path']
req2 = urllib.request.Request('http://127.0.0.1:7777/api/weave',
    data=json.dumps({'invoice': inv_path, 'shipper': '测试'}).encode(),
    headers={'Content-Type': 'application/json'})
d2 = json.loads(urllib.request.urlopen(req2).read())
dp = d2['download_path']
print('path:', dp)
print('exists:', os.path.exists(dp))
print('size:', os.path.getsize(dp) if os.path.exists(dp) else 0)
url = 'http://127.0.0.1:7777/api/download?path=' + quote(dp)
print('url:', url)
resp = urllib.request.urlopen(url)
data = resp.read()
print('dl:', len(data), 'bytes', 'OK' if len(data)>1000 else 'FAIL')
