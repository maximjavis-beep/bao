"""FBA 装箱单 Excel 导出器

基于 moban/模板.xlsx 生成物流商所需格式的清关装箱单。
在保留模板样式（颜色、合并单元格、Sheet 结构）的前提下填充数据。
"""
import copy
import shutil
import uuid
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font

# 模板路径
_TEMPLATE_PATH = Path(__file__).parent.parent.parent / "templates" / "模板.xlsx"

# 列索引映射 (1-based)
COL = {
    "A": 1, "B": 2, "C": 3, "D": 4, "E": 5, "F": 6, "G": 7, "H": 8,
    "I": 9, "J": 10, "K": 11, "L": 12, "M": 13, "N": 14,
    "O": 15, "P": 16, "Q": 17, "R": 18, "S": 19, "T": 20,
    "U": 21, "V": 22, "W": 23, "X": 24,
}

# 数据起始行（1-based，模板 Row 4 是第一条数据）
DATA_START_ROW = 4

# 白色填充
WHITE_FILL = openpyxl.styles.PatternFill(
    start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
)
# 黄色填充 (FF FF FF 00 去掉 alpha → FFFFFF00)
YELLOW_FILL = openpyxl.styles.PatternFill(
    start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid"
)
# 灰色填充
GRAY_FILL = openpyxl.styles.PatternFill(
    start_color="FFD8D8D8", end_color="FFD8D8D8", fill_type="solid"
)

# 黄色列 (A,B,C,D,E,O,S,U,V,W,X)
YELLOW_COLS = {COL[c] for c in "ABCDEOSUVWX"}
# 灰色列 (F,G,H,I,J,K,L,M,N,P,Q,R,T)
GRAY_COLS = {COL[c] for c in "FGHIJKLMNPQRT"}

# 申报币种说明文字
DECLARATION_NOTE = (
    "申报单价按照实际填写，一般建议不低于销售链接的30%（S列），\n"
    "特殊情况例如：新品单价不一致、断货没有售价等情况请提前与客服沟通。\n \n"
    "请注意申报币种：\n"
    "1,英国申报币种为GBP/USD，\n"
    "2,【勾选-已选择九方进口商】，加拿大&澳洲申报币种为USD\n"
    "3,海运渠道名含【欧盟递延】/空运含【比利时】申报币种为EUR"
)


class FBAExporter:
    """FBA 装箱单导出器"""

    def __init__(self):
        if not _TEMPLATE_PATH.exists():
            raise FileNotFoundError(f"模板文件不存在: {_TEMPLATE_PATH}")

    def export(self, woven: dict, output_path: str) -> str:
        """将编织结果导出为装箱单 Excel

        Args:
            woven: weave_fba() 返回的字典
            output_path: 输出文件路径

        Returns:
            输出文件路径
        """
        # 复制模板作为基础（保留双 Sheet 和样式）
        shutil.copy2(_TEMPLATE_PATH, output_path)
        wb = openpyxl.load_workbook(output_path)
        ws = wb["下单模板"]

        rows = woven.get("rows", [])
        row_count = len(rows)

        # ── 1. 清除模板示例数据 ──────────────────────────
        # 模板数据行从 Row 4 到 Row 44，全部清除（含图片对象）
        for r in range(DATA_START_ROW, 100):
            for c in range(1, 25):
                ws.cell(row=r, column=c).value = None
        # 清除模板中残留的图片（避免 T 列等示例图片被带到输出文件）
        for img in list(ws._images):
            ws._images.remove(img)

        # ── 2. 填充标题行 (Row 2) ─────────────────────────
        ws.cell(row=2, column=COL["F"], value=woven.get("total_boxes", 0))
        ws.cell(row=2, column=COL["H"], value=woven.get("total_cbm", 0))
        ws.cell(row=2, column=COL["J"], value=woven.get("total_weight", 0))
        ws.cell(row=2, column=COL["O"], value=DECLARATION_NOTE)
        ws.cell(row=2, column=COL["S"], value="申报币种*")
        # T 列（图片）不自动填入

        # ── 3. 居中样式 ──────────────────────────────────
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        body_font = Font(name="微软雅黑", size=10)

        # ── 4. 填充数据行 ───────────────────────────────
        for i, row_data in enumerate(rows):
            r = DATA_START_ROW + i

            # A: Shipment ID
            self._set_cell(ws, r, "A", woven.get("shipment_id", ""),
                           center, body_font, YELLOW_FILL)
            # B: Reference ID（留空）
            # C: 箱号段
            self._set_cell(ws, r, "C", row_data.get("箱号段", ""),
                           center, body_font, YELLOW_FILL)
            # D: 总件数
            self._set_cell(ws, r, "D", int(row_data.get("总件数", 0)),
                           center, body_font, YELLOW_FILL)
            # E: SKU
            self._set_cell(ws, r, "E", row_data.get("SKU", ""),
                           center, body_font, YELLOW_FILL)
            # F: 英文品名（留空）
            # G: 中文品名（留空）
            # H: Brand（留空）
            # I: 材质（留空）
            # J: 用途（留空）
            # K: 海关编码
            self._set_cell(ws, r, "K", str(row_data.get("海关编码", "")),
                           center, body_font, GRAY_FILL)
            # L: ASIN
            self._set_cell(ws, r, "L", str(row_data.get("ASIN", "")),
                           center, body_font, GRAY_FILL)
            # M: 是否带电（留空）
            # N: 型号（留空）
            # O: 总数量
            self._set_cell(ws, r, "O", int(row_data.get("总数量", 0)),
                           center, body_font, YELLOW_FILL)
            # P: 单位（留空）
            # Q: 每套个数（留空）
            # R: 采购单价（留空）
            # S: 申报单价（留空）
            # T: 图片（留空）
            # U: 单箱重量
            self._set_cell(ws, r, "U", row_data.get("单箱重量", 0),
                           center, body_font, YELLOW_FILL)
            # V: 长
            self._set_cell(ws, r, "V", row_data.get("长", 0),
                           center, body_font, YELLOW_FILL)
            # W: 宽
            self._set_cell(ws, r, "W", row_data.get("宽", 0),
                           center, body_font, YELLOW_FILL)
            # X: 高
            self._set_cell(ws, r, "X", row_data.get("高", 0),
                           center, body_font, YELLOW_FILL)

        # ── 5. 设置标题行样式 ──────────────────────────
        header_font = Font(name="微软雅黑", bold=True, size=11)
        for c in range(1, 25):
            cell = ws.cell(row=2, column=c)
            if cell.value is not None:
                cell.font = header_font

        # Row 2 标号列样式
        for r in [2]:
            for c in [COL["A"], COL["F"], COL["G"], COL["I"], COL["J"],
                       COL["S"], COL["T"]]:
                cell = ws.cell(row=r, column=c)
                if cell.value is not None:
                    cell.font = header_font

        # ── 6. 保存 ─────────────────────────────────────
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        wb.close()
        return output_path

    def _set_cell(self, ws, row, col_letter, value, alignment=None,
                  font=None, fill=None):
        """设置单元格值和样式"""
        cell = ws.cell(row=row, column=COL[col_letter])
        cell.value = value
        if alignment:
            cell.alignment = alignment
        if font:
            cell.font = font
        if fill:
            cell.fill = fill
