#!/bin/bash
# bao 服务重启
cd /Users/streiten/customs/bao || exit 1
source venv/bin/activate
pkill -f start_web.py 2>/dev/null
for port in 7777 8888 7000; do kill -9 $(lsof -ti :$port) 2>/dev/null; done
find . -name "*.pyc" -delete 2>/dev/null
find . -name "__pycache__" -type d -exec rm -rf {} + 2>/dev/null
sleep 1
nohup /Users/streiten/customs/bao/venv/bin/python3 -B start_web.py 8888 > /tmp/bao.log 2>&1 &
sleep 2
echo "bao → http://127.0.0.1:8888"
