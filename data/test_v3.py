import base64, json, os, sys, threading, time, urllib.request
from http.server import HTTPServer
sys.path.insert(0, '/Users/streiten/customs/bao')
from bao.web.server import BaoHandler
port = 9091
server = HTTPServer(("", port), BaoHandler)
t = threading.Thread(target=server.serve_forever, daemon=True)
t.start()
time.sleep(1)
os.chdir('/Users/streiten/customs/bao')
inv_b64 = base64.b64encode(open('data/示例发票.xlsx','rb').read()).decode()
d = json.loads(urllib.request.urlopen(urllib.request.Request(f'http://127.0.0.1:{port}/api/parse', data=json.dumps({'invoice': inv_b64}).encode(), headers={'Content-Type':'application/json'})).read())
print('parse:', d['success'])
d2 = json.loads(urllib.request.urlopen(urllib.request.Request(f'http://127.0.0.1:{port}/api/weave', data=json.dumps({'invoice': d['invoice']['path'], 'shipper':'x'}).encode(), headers={'Content-Type':'application/json'})).read())
print('weave:', d2['success'])
print('download_url:', d2['download_url'])
resp = urllib.request.urlopen(f"http://127.0.0.1:{port}{d2['download_url']}")
data = resp.read()
print('dl:', len(data), 'bytes', 'OK' if len(data)>100 else 'FAIL')
server.shutdown()
