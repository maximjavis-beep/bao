#!/bin/bash
PORT=8888
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"
PID=$(lsof -ti :$PORT 2>/dev/null)
if [ -n "$PID" ]; then
    kill $PID 2>/dev/null || kill -9 $PID 2>/dev/null
    sleep 1
fi
echo "bao panel -> http://127.0.0.1:$PORT"
python3 -c "from bao.web.server import run_server; run_server(port=$PORT)"
