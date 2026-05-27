"""archive 命令 — 报关单归档检索"""
import typer
from rich.console import Console
from rich.table import Table
from ..db.connection import Database
from ..db.repository import DeclarationRepository

app = typer.Typer(help="报关单归档检索", no_args_is_help=True)
console = Console()

@app.command("list")
def archive_list(limit: int = typer.Option(20, "--limit", "-l")):
    """列出最近归档的报关单"""
    db = Database()
    repo = DeclarationRepository(db)
    rows = repo.list_all(limit=limit)
    db.close()
    if not rows:
        console.print("[dim]暂无归档记录[/dim]")
        return
    tbl = Table(title=f"最近 {len(rows)} 条归档记录")
    tbl.add_column("日期"); tbl.add_column("报关单号"); tbl.add_column("发货人"); tbl.add_column("收货人"); tbl.add_column("金额")
    for r in rows:
        tbl.add_row(r.get("declaration_date",""), r.get("declaration_id",""), r.get("domestic_shipper",""), r.get("overseas_consignee",""), f"USD {r.get('total_amount',0):,.2f}")
    console.print(tbl)

@app.command("search")
def archive_search(keyword: str = typer.Argument(..., help="搜索关键词")):
    """搜索报关单（按发货人/收货人/合同号/发票号）"""
    db = Database()
    repo = DeclarationRepository(db)
    rows = repo.search(keyword)
    db.close()
    console.print(f"搜索 \"{keyword}\" — {len(rows)} 条结果")
    for r in rows:
        console.print(f"  [cyan]{r['declaration_id']}[/cyan] {r['domestic_shipper']} → {r['overseas_consignee']}  USD {r['total_amount']:,.2f}")

@app.command("stats")
def archive_stats():
    """按月统计报关量和金额"""
    db = Database()
    repo = DeclarationRepository(db)
    rows = repo.stats_by_month()
    db.close()
    if not rows:
        console.print("[dim]暂无统计数据[/dim]")
        return
    tbl = Table(title="月度统计")
    tbl.add_column("月份"); tbl.add_column("票数"); tbl.add_column("总金额"); tbl.add_column("FOB总计"); tbl.add_column("退税额")
    for r in rows:
        tbl.add_row(r["month"], str(r["count"]), f"{r['total_amount']:,.2f}", f"{r['total_fob']:,.2f}", f"{r['total_rebate']:,.2f}")
    console.print(tbl)
