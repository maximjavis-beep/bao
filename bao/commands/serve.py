"""serve 命令 — 启动 Web 面板"""

import socket
import webbrowser
from threading import Timer

import typer

from ..web.server import run_server

app = typer.Typer(help="Web 面板", no_args_is_help=True)


def _check_port(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            s.bind(("", port))
            return True
        except OSError:
            return False


def _open_browser(port: int):
    webbrowser.open(f"http://127.0.0.1:{port}")


@app.command("start")
def serve_start(
    port: int = typer.Option(8888, "--port", "-p", help="监听端口"),
    no_browser: bool = typer.Option(False, "--no-browser", help="不自动打开浏览器"),
):
    """启动本地 Web 面板

    上传发票和装箱单，实时预览编织结果，下载报关单 Excel。
    启动后用浏览器访问 http://127.0.0.1:8888
    """
    if not _check_port(port):
        from rich.console import Console
        Console().print(f"[red]端口 {port} 被占用，换一个: bao serve start -p {port+1}[/red]")
        raise typer.Exit(1)

    if not no_browser:
        Timer(0.8, _open_browser, args=(port,)).start()

    run_server(port=port)
