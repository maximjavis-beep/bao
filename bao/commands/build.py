"""build 命令 — FBA 装箱单生成"""

from datetime import date
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.table import Table

from ..core.exporter import FBAExporter
from ..core.weaver import weave_fba
from ..parsers.fba_parser import FBAParser

app = typer.Typer(help="FBA 装箱单生成", no_args_is_help=True)
console = Console()


@app.command("from-fba")
def build_from_fba(
    input_file: str = typer.Option(
        ..., "--input", "-i", help="FBA 货件 Excel 文件路径"
    ),
    output: Optional[str] = typer.Option(
        None, "--output", "-o", help="输出装箱单路径（默认：{货件编号}-装箱单.xlsx）"
    ),
    hs_code: Optional[str] = typer.Option(
        None, "--hs-code", help="海关编码覆盖（默认自动匹配）"
    ),
):
    """从 FBA 货件数据生成清关装箱单

    示例:
        bao build from-fba -i shuru/FBA15LRB2LTZ-20260509201827344.xlsx
        bao build from-fba -i 货件数据.xlsx -o 定制输出.xlsx --hs-code 3406000090
    """
    # ── 1. 解析 ────────────────────────────────────────
    console.print("[bold]📦 解析 FBA 货件数据…[/bold]")
    parser = FBAParser()

    try:
        data = parser.parse(input_file)
    except FileNotFoundError as e:
        console.print(f"[red]✗[/red] {e}")
        raise typer.Exit(1)

    meta = data.get("meta", {})
    items = data.get("items", [])

    if not items:
        console.print("[red]✗ 未解析到 SKU 数据[/red]")
        raise typer.Exit(1)

    sid = meta.get("shipment_id", "?")
    console.print(f"  ✓ 货件编号: [cyan]{sid}[/cyan]")
    console.print(f"  ✓ 总箱数: [cyan]{meta.get('total_boxes', '?')}[/cyan]")
    console.print(f"  ✓ SKU 数: [cyan]{len(items)}[/cyan]")

    # ── 2. 编织 ────────────────────────────────────────
    console.print("[bold]🧵 编织装箱单…[/bold]")
    woven = weave_fba(data, hs_code_override=hs_code)

    console.print(f"  ✓ 总重量: [cyan]{woven['total_weight']} kg[/cyan]")
    console.print(f"  ✓ 总体积: [cyan]{woven['total_cbm']} CBM[/cyan]")
    console.print(f"  ✓ 数据行: [cyan]{len(woven['rows'])}[/cyan]")

    # ── 3. 导出 ────────────────────────────────────────
    if output is None:
        output = f"{sid}-装箱单.xlsx" if sid else "装箱单.xlsx"

    console.print("[bold]💾 导出装箱单…[/bold]")
    exporter = FBAExporter()
    out_path = exporter.export(woven, output)
    console.print(f"  ✓ 已保存到: [bold cyan]{out_path}[/bold cyan]")

    # ── 4. 预览 ────────────────────────────────────────
    console.print()
    if woven["rows"]:
        tbl = Table(title="装箱单明细预览（前 10 行）")
        tbl.add_column("箱号段")
        tbl.add_column("件数")
        tbl.add_column("SKU", max_width=30)
        tbl.add_column("数量")
        tbl.add_column("重量")
        tbl.add_column("体积(cm)")

        for row in woven["rows"][:10]:
            dims = f"{row['长']}×{row['宽']}×{row['高']}"
            tbl.add_row(
                row["箱号段"],
                str(row["总件数"]),
                row["SKU"][:28],
                str(row["总数量"]),
                str(row["单箱重量"]),
                dims,
            )
        console.print(tbl)

    console.print(f"\n[bold green]✨ 完成！[/bold green]")


@app.command("preview")
def preview(
    input_file: str = typer.Option(
        ..., "--input", "-i", help="FBA 货件 Excel 文件路径"
    ),
):
    """仅解析和预览 FBA 数据内容（不生成装箱单）"""
    parser = FBAParser()

    console.print("[bold]📄 FBA 货件内容:[/bold]")
    try:
        data = parser.parse(input_file)
    except FileNotFoundError as e:
        console.print(f"[red]✗ {e}[/red]")
        raise typer.Exit(1)

    meta = data.get("meta", {})
    console.print(f"  货件编号: [cyan]{meta.get('shipment_id', '?')}[/cyan]")
    console.print(f"  箱子数量: [cyan]{meta.get('total_boxes', '?')}[/cyan]")
    console.print(f"  SKU 数量: [cyan]{meta.get('sku_count', '?')}[/cyan]")
    console.print(f"  商品数量: [cyan]{meta.get('item_count', '?')}[/cyan]")

    items = data.get("items", [])
    if items:
        tbl = Table(title=f"SKU 明细（共 {len(items)} 项）")
        tbl.add_column("MSKU", max_width=30)
        tbl.add_column("箱数")
        tbl.add_column("申报量")
        tbl.add_column("箱号段")
        tbl.add_column("重量/kg")
        tbl.add_column("长×宽×高")

        for item in items:
            dims = f"{item.get('length', 0)}×{item.get('width', 0)}×{item.get('height', 0)}"
            tbl.add_row(
                item.get("msku", "")[:28],
                str(item.get("box_count", 0)),
                str(int(item.get("declared_qty", 0))),
                str(item.get("box_range", "")),
                str(item.get("weight", 0)),
                dims,
            )
        console.print(tbl)
