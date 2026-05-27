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
print('weave OK, path:', dp)

# 测试新 download API
url = 'http://127.0.0.1:7777/api/download?path=' + quote(dp)
resp = urllib.request.urlopen(url)
d3 = json.loads(resp.read())
print('download success:', d3['success'])
print('filename:', d3.get('filename', '?'))
print('size:', d3.get('size', 0))
print('base64 len:', len(d3.get('data', '')))
# 验证 base64 数据可以解码
decoded = base64.b64decode(d3['data'])
print('decoded bytes:', len(decoded), 'OK' if len(decoded) == d3['size'] else 'MISMATCH')
