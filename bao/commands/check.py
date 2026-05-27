"""check 命令 — FBA 装箱单数据校验"""
import typer
from rich.console import Console
from rich.table import Table

from ..core.validator import check_fba_data, Severity
from ..parsers.fba_parser import FBAParser

app = typer.Typer(help="校验 FBA 装箱单数据", no_args_is_help=True)
console = Console()


@app.command("fba")
def check_fba(
    input_file: str = typer.Option(
        ..., "--input", "-i", help="FBA 货件 Excel 路径"
    ),
):
    """校验 FBA 货件数据"""
    parser = FBAParser()
    data = parser.parse(input_file)

    meta = data.get("meta", {})
    console.print(f"\n货件编号: [cyan]{meta.get('shipment_id', '?')}[/cyan]")

    report = check_fba_data(data)
    console.print(f"\n[bold]校验报告 — 共计 {report.total} 条[/bold]")

    if report.total == 0:
        console.print("[bold green]✓ 所有校验通过[/bold green]")
        return

    tbl = Table(show_header=True, box=None)
    tbl.add_column("级别", width=8)
    tbl.add_column("规则", width=12)
    tbl.add_column("消息")

    icon = {
        Severity.ERROR: "[red]✗ 阻断[/red]",
        Severity.WARNING: "[yellow]⚠ 警告[/yellow]",
        Severity.INFO: "[blue]ℹ 提示[/blue]",
    }

    for r in report.results:
        tbl.add_row(icon.get(r.severity, ""), r.rule, r.message)

    console.print(tbl)
    console.print()
    console.print(
        f"[red]{report.error_count} 阻断[/red]  "
        f"[yellow]{report.warning_count} 警告[/yellow]  "
        f"[blue]{report.info_count} 提示[/blue]"
    )

    if report.is_valid:
        console.print("\n[green]✓ 无阻断性错误[/green]")
    else:
        console.print(
            f"\n[red]✗ 存在 {report.error_count} 个阻断性错误[/red]"
        )
