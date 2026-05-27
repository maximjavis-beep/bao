import sys, os, threading, time, urllib.request, json
sys.path.insert(0, '/Users/streiten/customs/bao')
for m in list(sys.modules.keys()):
    if 'bao' in m: del sys.modules[m]
from bao.web.server import BaoHandler, HTTPServer
class T(BaoHandler):
    def do_GET(self):
        print('PATH:', repr(self.path), flush=True)
        super().do_GET()
s = HTTPServer(('', 9999), T)
t = threading.Thread(target=s.serve_forever, daemon=True)
t.start()
time.sleep(1)
r = urllib.request.urlopen('http://127.0.0.1:9999/api/archive')
print('STATUS:', r.status)
print('BODY:', r.read().decode()[:200])
s.shutdown()
