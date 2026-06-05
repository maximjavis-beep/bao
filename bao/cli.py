"""CLI 入口 — Typer + Rich"""
import typer
from rich.console import Console
from .commands import build, serve
from .commands.check import app as check_app
from .commands.archive import app as archive_app

app = typer.Typer(name="bao", help="FBA 装箱单生成助手", no_args_is_help=True)
console = Console()

app.add_typer(build.app, name="build", help="FBA 装箱单生成")
app.add_typer(serve.app, name="serve", help="启动 Web 面板")
app.add_typer(check_app, name="check", help="校验 FBA 装箱单数据")
app.add_typer(archive_app, name="archive", help="历史归档检索")


@app.command("version")
def version():
    """显示版本"""
    console.print("[bold cyan]bao[/bold cyan] v0.6.5 — FBA 装箱单生成助手")


if __name__ == "__main__":
    app()
