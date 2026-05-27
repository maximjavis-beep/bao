"""直接启动 Web 面板"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
# 强制清除模块缓存
for m in list(sys.modules.keys()):
    if 'bao' in m: del sys.modules[m]
from bao.web.server import run_server
port = int(sys.argv[1]) if len(sys.argv) > 1 else 7777
run_server(port)
