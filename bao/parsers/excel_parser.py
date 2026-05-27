"""Excel 发票/装箱单解析器

从业务 Excel 文件中提取结构化数据。
自动检测列标题行位置，通过列名映射适配不同格式。
"""

from pathlib import Path
from typing import Optional

import pandas as pd


class ExcelParser:
    """Excel 单据解析器

    通过列名映射表（column_map）适配不同客户的 Excel 格式。
    自动扫描前 30 行找到列标题行，无需严格格式。
    """

    # 发票默认列名映射
    INVOICE_DEFAULT_MAP = {
        "seq": ["序号", "项号", "No.", "Item", "行号"],
        "name_zh": ["品名", "商品名称", "货名", "Description"],
        "name_en": ["英文品名", "Description(EN)", "英文名称"],
        "specs": ["规格型号", "规格", "Specification", "型号"],
        "hs_code": ["HS编码", "商品编号", "HS Code", "HS"],
        "quantity": ["数量", "Quantity", "Qty"],
        "unit": ["单位", "Unit", "计量单位"],
        "unit_price": ["单价", "Unit Price", "Price"],
        "total_price": ["总价", "金额", "Total", "Amount"],
        "net_weight": ["净重", "Net Weight", "N.W."],
        "gross_weight": ["毛重", "Gross Weight", "G.W."],
        "origin": ["原产国", "Origin", "原产地"],
    }

    PACKING_DEFAULT_MAP = {
        "seq": ["箱号", "序号", "Carton No.", "C/No."],
        "description": ["品名", "货名", "Description"],
        "quantity": ["数量", "Qty", "Quantity"],
        "unit": ["单位", "Unit"],
        "net_weight": ["净重(KG)", "N.W.", "Net Weight"],
        "gross_weight": ["毛重(KG)", "G.W.", "Gross Weight"],
        "dimensions": ["尺寸", "Dimensions", "Size"],
    }

    HEADER_FIELDS = {
        "invoice_no": ["发票号", "Invoice No.", "INV No."],
        "invoice_date": ["发票日期", "Invoice Date", "Date"],
        "seller": ["卖方", "Seller", "Shipper", "出口商"],
        "buyer": ["买方", "Buyer", "Consignee", "进口商", "收货人"],
        "contract_no": ["合同号", "Contract No.", "S/C No."],
        "vessel_name": ["船名", "Vessel"],
        "port_of_loading": ["起运港", "Port of Loading", "装货港"],
        "port_of_discharge": ["目的港", "Port of Discharge", "卸货港"],
        "incoterm": ["贸易术语", "Terms", "Incoterm"],
        "currency": ["币制", "Currency"],
        "package_count": ["件数", "总件数", "Packages", "Total Packages"],
        "package_type": ["包装", "包装种类", "Packing"],
        # 装箱单额外字段
        "packing_no": ["装箱单号", "Packing No."],
        "shipper": ["发货人", "Shipper", "发货方"],
        "consignee": ["收货人", "Consignee"],
        "marks": ["唛头", "Marks", "Shipping Marks"],
    }

    def __init__(self, column_map: Optional[dict] = None):
        self._inv_map = dict(self.INVOICE_DEFAULT_MAP)
        self._pkg_map = dict(self.PACKING_DEFAULT_MAP)
        self._hdr_map = dict(self.HEADER_FIELDS)
        if column_map:
            for key, names in column_map.items():
                if key in self._inv_map:
                    self._inv_map[key] = names
                elif key in self._pkg_map:
                    self._pkg_map[key] = names
                elif key in self._hdr_map:
                    self._hdr_map[key] = names

    # ── 列标题检测 ──────────────────────────────────────────

    def _detect_header_row(self, df: pd.DataFrame, field_map: dict) -> int:
        """扫描前 30 行，找到匹配候选列名最多的行号

        Returns:
            列标题行的 0-based 行号。未找到返回 -1。
        """
        max_rows = min(30, len(df))
        all_candidates = set()
        for names in field_map.values():
            all_candidates.update(n.lower() for n in names)

        best_row = -1
        best_score = 0
        for row_idx in range(max_rows):
            score = 0
            for col_idx in range(min(len(df.columns), 30)):
                val = df.iat[row_idx, col_idx]
                if pd.notna(val) and str(val).strip().lower() in all_candidates:
                    score += 1
            if score > best_score:
                best_score = score
                best_row = row_idx
        # 需要至少匹配 2 个列名
        return best_row if best_score >= 2 else -1

    def _build_column_map_from_row(
        self, df: pd.DataFrame, header_row: int, field_map: dict
    ) -> dict[str, Optional[int]]:
        """从列标题行构建 {字段名: 列索引} 映射"""
        result: dict[str, Optional[int]] = {}
        row_vals = {}
        for col_idx in range(len(df.columns)):
            val = df.iat[header_row, col_idx]
            if pd.notna(val):
                row_vals[str(val).strip().lower()] = col_idx

        for field, candidates in field_map.items():
            found = None
            for cand in candidates:
                if cand.lower() in row_vals:
                    found = row_vals[cand.lower()]
                    break
            result[field] = found
        return result

    def detect_columns(self, file_path: str) -> dict:
        df = pd.read_excel(file_path, header=None, dtype=str)
        header_row = self._detect_header_row(df, self._inv_map)
        rows = []
        if header_row >= 0:
            for col_idx in range(min(len(df.columns), 50)):
                val = df.iat[header_row, col_idx]
                if pd.notna(val):
                    rows.append(str(val).strip())
        col_map = self._build_column_map_from_row(df, header_row, self._inv_map) if header_row >= 0 else {}
        matched = {}
        matched_cols = set()
        for field, col_idx in col_map.items():
            if col_idx is not None and col_idx >= 0:
                cname = str(df.iat[header_row, col_idx]).strip() if pd.notna(df.iat[header_row, col_idx]) else ""
                if cname:
                    matched[field] = cname
                    matched_cols.add(cname)
        unmatched = [c for c in rows if c not in matched_cols]
        return {"columns": rows, "matched": matched, "unmatched": unmatched}

    # ── 表头信息提取（label-value 对） ────────────────────────

    def _extract_header_info(
        self, df: pd.DataFrame, end_row: int
    ) -> dict:
        """在列标题行之前的区域扫描 label-value 对

        Args:
            df: 原始 DataFrame
            end_row: 列标题行之前的最大行号（0-based，不含）
        """
        header = {}
        for field, keywords in self._hdr_map.items():
            for kw in keywords:
                for row_idx in range(min(end_row, len(df))):
                    for col_idx in range(min(10, len(df.columns))):
                        val = df.iat[row_idx, col_idx]
                        if pd.notna(val) and str(val).strip() == kw:
                            # 尝试右侧相邻单元格
                            for offset in range(1, 6):
                                try:
                                    v = df.iat[row_idx, col_idx + offset]
                                    if pd.notna(v):
                                        header[field] = str(v).strip()
                                        break
                                except (IndexError, ValueError):
                                    pass
                            if field in header:
                                break
                    if field in header:
                        break
                if field in header:
                    break
        return header

    # ── 发票解析 ────────────────────────────────────────────

    def parse_invoice(self, file_path: str) -> dict:
        """解析商业发票 Excel 文件"""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"发票文件不存在: {file_path}")

        df = pd.read_excel(file_path, header=None, dtype=str)

        # 检测列标题行
        header_row = self._detect_header_row(df, self._inv_map)
        if header_row < 0:
            return {"header": {}, "items": []}

        # 提取表头信息（列标题行之前）
        inv_header = self._extract_header_info(df, header_row)

        # 构建列映射
        col_map = self._build_column_map_from_row(df, header_row, self._inv_map)

        # 提取数据（列标题行之后的非空行）
        items = []
        for row_idx in range(header_row + 1, len(df)):
            seq_val = df.iat[row_idx, col_map.get("seq", -1)] if col_map.get("seq") is not None else None
            if seq_val is None or pd.isna(seq_val) or str(seq_val).strip() == "":
                continue
            # 跳过汇总行
            try:
                int(float(str(seq_val)))
            except (ValueError, TypeError):
                continue

            item = {}
            for field, col_idx in col_map.items():
                if col_idx is None or col_idx < 0:
                    continue
                val = df.iat[row_idx, col_idx]
                item[field] = str(val).strip() if pd.notna(val) else ""

            for nf in ("seq", "quantity", "unit_price", "total_price",
                        "net_weight", "gross_weight"):
                if nf in item and item[nf]:
                    try:
                        item[nf] = float(item[nf])
                    except (ValueError, TypeError):
                        item[nf] = 0.0
            if "seq" in item:
                try:
                    item["seq"] = int(item["seq"])
                except (ValueError, TypeError):
                    pass

            items.append(item)

        return {"header": inv_header, "items": items}

    # ── 装箱单解析 ────────────────────────────────────────────

    def parse_packing(self, file_path: str) -> dict:
        """解析装箱单 Excel 文件"""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"装箱单文件不存在: {file_path}")

        df = pd.read_excel(file_path, header=None, dtype=str)

        header_row = self._detect_header_row(df, self._pkg_map)
        if header_row < 0:
            return {"header": {}, "packages": []}

        pkg_header = self._extract_header_info(df, header_row)

        col_map = self._build_column_map_from_row(df, header_row, self._pkg_map)

        packages = []
        for row_idx in range(header_row + 1, len(df)):
            seq_val = df.iat[row_idx, col_map.get("seq", -1)] if col_map.get("seq") is not None else None
            if seq_val is None or pd.isna(seq_val) or str(seq_val).strip() == "":
                continue
            try:
                int(float(str(seq_val)))
            except (ValueError, TypeError):
                continue

            pkg = {}
            for field, col_idx in col_map.items():
                if col_idx is None or col_idx < 0:
                    continue
                val = df.iat[row_idx, col_idx]
                pkg[field] = str(val).strip() if pd.notna(val) else ""

            for nf in ("seq", "quantity", "net_weight", "gross_weight"):
                if nf in pkg and pkg[nf]:
                    try:
                        pkg[nf] = float(pkg[nf])
                    except (ValueError, TypeError):
                        pkg[nf] = 0.0
            if "seq" in pkg:
                try:
                    pkg["seq"] = int(pkg["seq"])
                except (ValueError, TypeError):
                    pass

            packages.append(pkg)

        return {"header": pkg_header, "packages": packages}
