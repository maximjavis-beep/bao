"""FBA 装箱单 Excel 导出器 — 基于蜡烛-模版.xlsx"""
import io
import shutil
import tempfile
from pathlib import Path
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font

_TEMPLATE_PATH = Path(__file__).parent.parent.parent / "templates" / "蜡烛-模版.xlsx"

COL = {
    "A": 1, "B": 2, "C": 3, "D": 4, "E": 5, "F": 6, "G": 7, "H": 8,
    "I": 9, "J": 10, "K": 11, "L": 12, "M": 13, "N": 14,
    "O": 15, "P": 16, "Q": 17, "R": 18, "S": 19, "T": 20,
    "U": 21, "V": 22, "W": 23, "X": 24,
}

DATA_START_ROW = 4
YELLOW_FILL = openpyxl.styles.PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
GRAY_FILL = openpyxl.styles.PatternFill(start_color="FFD8D8D8", end_color="FFD8D8D8", fill_type="solid")
FIXED_COLS = ["B", "F", "G", "H", "I", "J", "K", "M", "N", "P", "Q", "T"]

DECLARATION_NOTE = (
    "申报单价按照实际填写，一般建议不低于销售链接的30%（S列），\n"
    "特殊情况例如：新品单价不一致、断货没有售价等情况请提前与客服沟通。\n \n"
    "请注意申报币种：\n"
    "1,英国申报币种为GBP/USD，\n"
    "2,【勾选-已选择九方进口商】，加拿大&澳洲申报币种为USD\n"
    "3,海运渠道名含【欧盟递延】/空运含【比利时】申报币种为EUR"
)


class FBAExporter:
    def __init__(self):
        if not _TEMPLATE_PATH.exists():
            raise FileNotFoundError(f"模板不存在: {_TEMPLATE_PATH}")

    def export(self, woven: dict, output_path: str) -> str:
        shutil.copy2(_TEMPLATE_PATH, output_path)
        wb = openpyxl.load_workbook(output_path)
        ws = wb["下单模板"]
        rows = woven.get("rows", [])

        # 保存模板中 T 列图片数据 + 尺寸
        t_img_bytes = None
        t_w, t_h = 0, 0
        for img in ws._images:
            try:
                t_img_bytes = img._data()
                t_w, t_h = img.width, img.height
                break
            except Exception:
                pass

        # 读取模板 Row 4 固定值
        fixed = {}
        for col in FIXED_COLS:
            v = ws.cell(row=DATA_START_ROW, column=COL[col]).value
            fixed[col] = v

        # 清除数据行值 + 移除所有图片
        for r in range(DATA_START_ROW, 200):
            for c in range(1, 25):
                cell = ws.cell(row=r, column=c)
                if not isinstance(cell, MergedCell):
                    cell.value = None
        ws._images.clear()

        # Row 2 标题
        ws.cell(row=2, column=COL["F"], value=woven.get("total_boxes", 0))
        ws.cell(row=2, column=COL["H"], value=woven.get("total_cbm", 0))
        ws.cell(row=2, column=COL["J"], value=woven.get("total_weight", 0))
        ws.cell(row=2, column=COL["O"], value=DECLARATION_NOTE)
        ws.cell(row=2, column=COL["S"], value="申报币种*")

        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        bf = Font(name="微软雅黑", size=10)

        # 保留所有 temp 文件路径，save 后再清理
        _temp_files = []

        # 数据行
        for i, rd in enumerate(rows):
            r = DATA_START_ROW + i
            # 固定值
            for col, val in fixed.items():
                if val is not None:
                    self._set(ws, r, col, str(val), center, bf, GRAY_FILL)
            # 动态列
            self._set(ws, r, "A", woven.get("shipment_id", ""), center, bf, YELLOW_FILL)
            self._set(ws, r, "C", rd.get("箱号段", ""), center, bf, YELLOW_FILL)
            self._set(ws, r, "D", int(rd.get("总件数", 0)), center, bf, YELLOW_FILL)
            self._set(ws, r, "E", rd.get("SKU", ""), center, bf, YELLOW_FILL)
            self._set(ws, r, "L", str(rd.get("ASIN", "")), center, bf, GRAY_FILL)
            self._set(ws, r, "O", int(rd.get("总数量", 0)), center, bf, YELLOW_FILL)
            self._set(ws, r, "U", rd.get("单箱重量", 0), center, bf, YELLOW_FILL)
            self._set(ws, r, "V", rd.get("长", 0), center, bf, YELLOW_FILL)
            self._set(ws, r, "W", rd.get("宽", 0), center, bf, YELLOW_FILL)
            self._set(ws, r, "X", rd.get("高", 0), center, bf, YELLOW_FILL)

            # T 列图片：为每行创建副本
            if t_img_bytes:
                tf = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                tf.write(t_img_bytes)
                tf.close()
                _temp_files.append(tf.name)
                img = XLImage(tf.name)
                img.width, img.height = t_w, t_h
                img.anchor = f"T{r}"
                ws.add_image(img)

        # 标题样式
        hf = Font(name="微软雅黑", bold=True, size=11)
        for c in range(1, 25):
            cell = ws.cell(row=2, column=c)
            if cell.value is not None:
                cell.font = hf

        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        wb.close()

        # save 后清理 temp 文件
        for tf in _temp_files:
            try:
                Path(tf).unlink()
            except Exception:
                pass

        return output_path

    def _set(self, ws, row, col, value, align=None, font=None, fill=None):
        c = ws.cell(row=row, column=COL[col])
        if isinstance(c, MergedCell):
            return
        c.value = value
        if align: c.alignment = align
        if font: c.font = font
        if fill: c.fill = fill
