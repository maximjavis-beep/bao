import base64, json, os, urllib.request
os.chdir('/Users/streiten/customs/bao')
pkg_b64 = base64.b64encode(open('data/示例装箱单.xlsx','rb').read()).decode()
req = urllib.request.Request('http://127.0.0.1:7777/api/parse',
    data=json.dumps({'packing': pkg_b64}).encode(),
    headers={'Content-Type': 'application/json'})
d = json.loads(urllib.request.urlopen(req).read())
print('packing:', 'OK' if d['success'] else 'FAIL, err='+d.get('error',''))
print('  count:', d.get('packing',{}).get('count',0))
inv_b64 = base64.b64encode(open('data/示例发票.xlsx','rb').read()).decode()
req2 = urllib.request.Request('http://127.0.0.1:7777/api/parse',
    data=json.dumps({'invoice': inv_b64}).encode(),
    headers={'Content-Type': 'application/json'})
d2 = json.loads(urllib.request.urlopen(req2).read())
print('invoice:', 'OK' if d2['success'] else 'FAIL')
print('  items:', d2.get('invoice',{}).get('count',0))
req3 = urllib.request.Request('http://127.0.0.1:7777/api/weave',
    data=json.dumps({'invoice': d2['invoice']['path'], 'packing': d['packing']['path'], 'shipper': '测试'}).encode(),
    headers={'Content-Type': 'application/json'})
d3 = json.loads(urllib.request.urlopen(req3).read())
print('weave:', 'OK' if d3['success'] else 'FAIL')
print('  total:', d3['summary']['total_amount'], 'pkg:', d3['summary']['package_count'])
print('done')
