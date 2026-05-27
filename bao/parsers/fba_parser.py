"""FBA 原始装箱数据解析器

解析亚马逊 FBA 货件导出 Excel 的 "packing" sheet，
提取货件元信息和 SKU 装箱明细。
"""

from pathlib import Path

import pandas as pd


class FBAParser:
    """FBA 货件数据解析器

    输入格式（packing sheet）:
      Row 1-7: 元数据行（工作流程名称、货件编号、箱子数量等）
      Row 9:   列标题
      Row 10+: 数据行

    输出字典结构:
      {
        "meta": {
            "shipment_id": "FBA15LRB2LTZ",
            "total_boxes": 45,
            "sku_count": 9,
            "item_count": 720,
        },
        "items": [
            {
                "msku": "MSTC09EUS01 Citronella",
                "declared_qty": 270,
                "box_count": 15,
                "length": 50.2, "width": 45.7, "height": 33.4,
                "weight": 14.5,
                "box_range": "31~45",
            },
            ...
        ]
      }
    """

    COLUMN_MAP = {
        "msku": ["MSKU"],
        "asin": ["ASIN"],
        "declared_qty": ["商品申报量"],
        "boxed_qty": ["商品装箱量"],
        "box_count": ["箱数"],
        "length": ["箱子长（cm）"],
        "width": ["箱子宽（cm）"],
        "height": ["箱子高（cm）"],
        "weight": ["箱子重量（kg）"],
        "box_range": ["货件箱子编号"],
    }

    def parse(self, file_path: str) -> dict:
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")

        df = pd.read_excel(file_path, sheet_name="packing", header=None, dtype=str)
        if df.empty:
            return {"meta": {}, "items": []}

        meta = self._extract_meta(df)
        header_row = self._find_header_row(df)
        if header_row < 0:
            return {"meta": meta, "items": []}

        col_map = self._build_column_map(df, header_row)
        items = []

        for row_idx in range(header_row + 1, len(df)):
            msku = self._cell(df, row_idx, col_map.get("msku"))
            if not msku or msku.strip() == "":
                continue

            item = {
                "msku": msku.strip(),
                "asin": self._cell(df, row_idx, col_map.get("asin")),
                "declared_qty": self._as_float(df, row_idx, col_map.get("declared_qty")),
                "boxed_qty": self._as_float(df, row_idx, col_map.get("boxed_qty")),
                "box_count": self._as_int(df, row_idx, col_map.get("box_count")),
                "length": self._as_float(df, row_idx, col_map.get("length")),
                "width": self._as_float(df, row_idx, col_map.get("width")),
                "height": self._as_float(df, row_idx, col_map.get("height")),
                "weight": self._as_float(df, row_idx, col_map.get("weight")),
                "box_range": self._cell(df, row_idx, col_map.get("box_range")),
            }
            items.append(item)

        return {"meta": meta, "items": items}

    def _extract_meta(self, df: pd.DataFrame) -> dict:
        meta = {}
        for row_idx in range(min(8, len(df))):
            label = self._cell(df, row_idx, 0)
            value = self._cell(df, row_idx, 1)
            if not label:
                continue
            label_lower = label.strip()
            if "货件编号" in label_lower:
                meta["shipment_id"] = value.strip()
            elif "箱子数量" in label_lower:
                meta["total_boxes"] = self._as_int(df, row_idx, 1)
            elif "SKU 数量" in label_lower or "SKU数量" in label_lower:
                meta["sku_count"] = self._as_int(df, row_idx, 1)
            elif "商品数量" in label_lower:
                meta["item_count"] = self._as_int(df, row_idx, 1)
        return meta

    def _find_header_row(self, df: pd.DataFrame) -> int:
        for row_idx in range(min(20, len(df))):
            for col_idx in range(min(20, len(df.columns))):
                val = self._cell(df, row_idx, col_idx)
                if val and "MSKU" in val:
                    return row_idx
        return -1

    def _build_column_map(self, df: pd.DataFrame, header_row: int) -> dict:
        col_map = {}
        for col_idx in range(min(30, len(df.columns))):
            val = self._cell(df, header_row, col_idx)
            if not val:
                continue
            for field, keywords in self.COLUMN_MAP.items():
                for kw in keywords:
                    if kw in val:
                        col_map[field] = col_idx
                        break
        return col_map

    @staticmethod
    def _cell(df: pd.DataFrame, row: int, col) -> str:
        if col is None:
            return ""
        try:
            val = df.iat[row, col]
            if pd.isna(val):
                return ""
            return str(val)
        except (IndexError, ValueError):
            return ""

    @staticmethod
    def _as_float(df: pd.DataFrame, row: int, col) -> float:
        val = FBAParser._cell(df, row, col)
        if not val:
            return 0.0
        try:
            return float(val)
        except (ValueError, TypeError):
            return 0.0

    @staticmethod
    def _as_int(df: pd.DataFrame, row: int, col) -> int:
        val = FBAParser._cell(df, row, col)
        if not val:
            return 0
        try:
            return int(float(val))
        except (ValueError, TypeError):
            return 0
