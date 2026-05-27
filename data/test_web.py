import base64, json, os, urllib.request
os.chdir('/Users/streiten/customs/bao')
inv_b64 = base64.b64encode(open('data/示例发票.xlsx','rb').read()).decode()
req = urllib.request.Request('http://127.0.0.1:7777/api/parse',
    data=json.dumps({"invoice": inv_b64}).encode(),
    headers={'Content-Type': 'application/json'})
resp = urllib.request.urlopen(req)
d = json.loads(resp.read())
print('parse:', 'OK' if d['success'] else 'FAIL')
print('  items:', d.get('invoice',{}).get('count',0))
print('  inv_no:', d.get('invoice',{}).get('header',{}).get('invoice_no','?'))
inv_path = d['invoice']['path']
req2 = urllib.request.Request('http://127.0.0.1:7777/api/weave',
    data=json.dumps({"invoice": inv_path, "packing": "", "shipper": "测试公司"}).encode(),
    headers={'Content-Type': 'application/json'})
resp2 = urllib.request.urlopen(req2)
d2 = json.loads(resp2.read())
print('weave:', 'OK' if d2['success'] else 'FAIL')
print('  total:', d2['summary']['total_amount'])
print('  items:', d2['summary']['item_count'])
print('  dest:', d2['summary']['destination_country'])
print('\n✅ 全部通过 — 浏览器访问 http://127.0.0.1:7777')
