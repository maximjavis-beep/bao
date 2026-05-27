"""FBA 装箱单编织引擎

将 FBA 原始货件数据转换为物流商模板行格式。
"""
from pathlib import Path
from typing import Optional

import openpyxl


# HS 编码参考表路径
_HS_PATH = Path(__file__).parent.parent.parent / "data" / "hs_code.xlsx"

# 默认 HS 编码（蜡烛类）
_DEFAULT_HS_CODE = "3406000090"


def load_hs_map() -> dict:
    """加载 HS 编码对照表 {品类关键词: HS编码}"""
    hs_map = {}
    if not _HS_PATH.exists():
        return hs_map

    wb = openpyxl.load_workbook(_HS_PATH, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        keyword = str(row[0]).strip() if row[0] else ""
        code = str(row[1]).strip() if row[1] else ""
        if keyword and code:
            hs_map[keyword] = code
    wb.close()
    return hs_map


def _box_range_normalize(box_range: str) -> str:
    """标准化箱号段：~ → -，单整数 → N-N"""
    if not box_range:
        return ""
    val = box_range.strip().replace("~", "-")
    # 纯整数（无分隔符）→ 补全为 N-N
    if val.isdigit():
        val = f"{val}-{val}"
    return val


def weave_fba(data: dict, hs_code_override: Optional[str] = None) -> dict:
    """将 FBA 解析结果编织为模板数据

    Args:
        data: FBAParser.parse() 返回的字典
        hs_code_override: 强制覆盖 HS 编码

    Returns:
        {
            "shipment_id": "FBA15LRB2LTZ",
            "total_boxes": 45,
            "total_weight": 816.9,
            "total_cbm": 3.191,
            "rows": [...]
        }
    """
    meta = data.get("meta", {})
    items = data.get("items", [])

    shipment_id = meta.get("shipment_id", "")
    total_boxes = meta.get("total_boxes", 0)

    hs_map = load_hs_map()
    rows = []
    total_weight = 0.0
    total_cbm = 0.0

    for item in items:
        box_range_raw = item.get("box_range", "")
        box_range = _box_range_normalize(box_range_raw)
        box_count = item.get("box_count", 0)
        weight = item.get("weight", 0)
        length = item.get("length", 0)
        width = item.get("width", 0)
        height = item.get("height", 0)
        declared_qty = int(item.get("declared_qty", 0))

        # HS 编码
        hs_code = hs_code_override or _DEFAULT_HS_CODE
        if not hs_code_override:
            msku = item.get("msku", "")
            for keyword, code in hs_map.items():
                if keyword in msku:
                    hs_code = code
                    break

        rows.append({
            "箱号段": box_range,
            "总件数": box_count,
            "SKU": item.get("msku", ""),
            "ASIN": item.get("asin", ""),
            "海关编码": hs_code,
            "总数量": declared_qty,
            "单箱重量": weight,
            "长": length,
            "宽": width,
            "高": height,
        })

        total_weight += weight * box_count
        # 先累加未舍入的体积，最后统一 round
        total_cbm += length * width * height * box_count / 1_000_000

    if total_boxes <= 0:
        total_boxes = sum(r["总件数"] for r in rows)

    return {
        "shipment_id": shipment_id,
        "total_boxes": total_boxes,
        "total_weight": round(total_weight, 2),
        "total_cbm": round(total_cbm, 3),
        "rows": rows,
    }
