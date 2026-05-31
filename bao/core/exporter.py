"""FBA 装箱单 Excel 导出器 — 模板颜色分灰/黄 + DISPIMG 支持"""
import shutil
import tempfile
from pathlib import Path
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker
from openpyxl.styles import Alignment, Font

_DEFAULT_TEMPLATE = Path(__file__).parent.parent.parent / "templates" / "蜡烛-模版.xlsx"

PATTERNS = [
    ("shipment_id",  ["shipment id", "shipmentid"]),
    ("reference_id", ["reference id", "referenceid"]),
    ("box_range",    ["箱号段", "箱号"]),
    ("total_boxes",  ["总件数"]),
    ("sku",          ["sku"]),
    ("en_name",      ["英文品名"]),
    ("cn_name",      ["中文品名"]),
    ("brand",        ["brand", "品牌"]),
    ("material",     ["材质"]),
    ("usage",        ["用途"]),
    ("hs_code",      ["海关编码", "hs", "hscode", "进口海关编码"]),
    ("asin_url",     ["asin", "销售链接"]),
    ("battery",      ["带电"]),
    ("model",        ["型号"]),
    ("total_qty",    ["总数量", "每箱数量"]),
    ("unit",         ["单位"]),
    ("per_set",      ["每套"]),
    ("cost",         ["采购单价", "采购", "投保单价"]),
    ("declare_price",["申报单价"]),
    ("image",        ["图片"]),
    ("box_weight",   ["单箱重量", "重量", "单箱重"]),
    ("length",       ["长", "单箱长"]),
    ("width",        ["宽", "单箱宽"]),
    ("height",       ["高", "单箱高"]),
]

YELLOW_FILL = openpyxl.styles.PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
GRAY_FILL = openpyxl.styles.PatternFill(start_color="FFD8D8D8", end_color="FFD8D8D8", fill_type="solid")


class FBAExporter:
    def __init__(self, template_path: str = None, tracking_map: dict = None):
        self._template_path = Path(template_path) if template_path else _DEFAULT_TEMPLATE
        self._tracking_map = tracking_map or {}
        if not self._template_path.exists():
            raise FileNotFoundError(f"模板不存在: {self._template_path}")

    def export(self, woven: dict, output_path: str) -> str:
        shutil.copy2(self._template_path, output_path)
        wb = openpyxl.load_workbook(output_path)
        ws_name = "下单模板" if "下单模板" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[ws_name]
        rows = woven.get("rows", [])

        header_row, col_map = self._detect_template(ws)
        data_start = header_row + 1

        def col(field: str) -> int:
            return col_map.get(field, 0)

        # ── 灰/黄分类 ──────────────────────────────────
        grey_fields = set()
        for c in range(1, min(30, ws.max_column or 0) + 1):
            cell = ws.cell(row=header_row, column=c)
            try:
                fg = cell.fill.fgColor
                is_grey = (fg.type == 'theme' and fg.theme is not None)
            except Exception:
                is_grey = False
            if not is_grey:
                continue
            val = str(cell.value or "").strip().lower()
            for field, kws in PATTERNS:
                for kw in kws:
                    if kw in val and field not in grey_fields:
                        grey_fields.add(field)
                        break

        # ── 提取 cellimages (DISPIMG) ──────────────────
        dispimg_info = self._extract_cell_images(str(output_path))

        # ── 提取 openpyxl 嵌入图（cellimages 不存在时）───
        t_img_bytes, t_w, t_h = None, 54, 66
        if not dispimg_info:
            for img in ws._images:
                try:
                    t_img_bytes = img._data()
                    t_w = img.width or 54
                    t_h = img.height or 66
                    break
                except Exception:
                    continue
        if not t_img_bytes and not dispimg_info and _DEFAULT_TEMPLATE.exists():
            try:
                dwb = openpyxl.load_workbook(_DEFAULT_TEMPLATE)
                dsn = "下单模板" if "下单模板" in dwb.sheetnames else dwb.sheetnames[0]
                for img in dwb[dsn]._images:
                    try:
                        t_img_bytes = img._data()
                        t_w = img.width or 54
                        t_h = img.height or 66
                        break
                    except Exception:
                        continue
                dwb.close()
            except Exception:
                pass

        # ── HS 品类映射 + DISPIMG 关联 ──────────────────
        template_categories = {}
        if col("hs_code") > 0:
            for sr in range(data_start, min(data_start + 10, ws.max_row + 1)):
                hs = ws.cell(row=sr, column=col("hs_code")).value
                if not hs:
                    continue
                hs_key = str(hs).strip()
                if hs_key in template_categories:
                    continue
                cat_data = {}
                for fld in grey_fields:
                    c = col(fld)
                    if c > 0:
                        v = ws.cell(row=sr, column=c).value
                        if v is not None:
                            cat_data[fld] = str(v)
                for drow, (dname, _ib, _w, _h) in dispimg_info.items():
                    if drow == sr:
                        cat_data["_dispimg"] = dname
                        break
                template_categories[hs_key] = cat_data

        # ── 清除数据行 ─────────────────────────────────
        max_c = col_map.get("_max_col", 24)
        for r in range(data_start, data_start + 200):
            for c in range(1, max_c + 1):
                cell = ws.cell(row=r, column=c)
                if not isinstance(cell, MergedCell):
                    cell.value = None
        ws._images.clear()

        r2 = {"total_boxes": woven.get("total_boxes", 0),
              "total_cbm": woven.get("total_cbm", 0),
              "total_weight": woven.get("total_weight", 0)}
        for fld, val in r2.items():
            c = col(fld)
            if c > 0:
                ws.cell(row=2, column=c, value=val)

        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        bf = Font(name="微软雅黑", size=10)
        _temp_files = []

        for i, rd in enumerate(rows):
            r = data_start + i
            ws.row_dimensions[r].height = 40

            row_hs = str(rd.get("进口海关编码") or rd.get("海关编码", "")).strip()
            tmpl_cat = template_categories.get(row_hs) if row_hs else None

            for fld in grey_fields:
                c = col(fld)
                if c <= 0:
                    continue
                val = tmpl_cat.get(fld) if tmpl_cat else None
                if val is not None:
                    self._set_col(ws, r, c, str(val), center, bf, GRAY_FILL)

            # 货件追踪码 lookup（Reference ID）
            if col("reference_id") > 0 and self._tracking_map:
                sid = woven.get("shipment_id", "")
                track_code = self._tracking_map.get(sid, "")
                if track_code:
                    self._set_col(ws, r, col("reference_id"), track_code, center, bf, YELLOW_FILL)

            yw = {
                "shipment_id": woven.get("shipment_id", ""),
                "box_range":   rd.get("箱号段", ""),
                "total_boxes": int(rd.get("总件数", 0)),
                "sku":         rd.get("SKU", ""),
                "total_qty":   int(rd.get("总数量", 0)),
                "box_weight":  rd.get("单箱重量", 0),
                "length":      rd.get("长", 0),
                "width":       rd.get("宽", 0),
                "height":      rd.get("高", 0),
                "en_name":     rd.get("英文品名", ""),
                "cn_name":     rd.get("中文品名", ""),
                "material":    rd.get("材质", ""),
                "usage":       rd.get("用途", ""),
                "hs_code":     rd.get("海关编码", ""),
            }
            for fld, val in yw.items():
                if fld in grey_fields or val is None:
                    continue
                c = col(fld)
                if c > 0:
                    self._set_col(ws, r, c, val, center, bf, YELLOW_FILL)

            if col("asin_url") > 0 and "asin_url" not in grey_fields:
                self._set_col(ws, r, col("asin_url"),
                              str(rd.get("ASIN", "")), center, bf, GRAY_FILL)

            # 图片：DISPIMG 公式 or TwoCellAnchor
            if col("image") > 0:
                if tmpl_cat and "_dispimg" in tmpl_cat:
                    ws.cell(row=r, column=col("image")).value = (
                        f'=_xlfn.DISPIMG("{tmpl_cat["_dispimg"]}",1)')
                elif t_img_bytes:
                    tf = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                    tf.write(t_img_bytes)
                    tf.close()
                    _temp_files.append(tf.name)
                    img = XLImage(tf.name)
                    img.width, img.height = max(t_w, 54), max(t_h, 66)
                    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor
                    img.anchor = TwoCellAnchor(
                        _from=AnchorMarker(col=col("image") - 1, row=r - 1,
                                           colOff=0, rowOff=0),
                        to=AnchorMarker(col=col("image"), row=r, colOff=0, rowOff=0),
                        editAs="oneCell")
                    ws.add_image(img)

        hf = Font(name="微软雅黑", bold=True, size=11)
        for c in range(1, max_c + 1):
            cell = ws.cell(row=2, column=c)
            if cell.value is not None:
                cell.font = hf

        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
        wb.close()

        for tf in _temp_files:
            try:
                Path(tf).unlink()
            except Exception:
                pass

        if dispimg_info:
            self._inject_cellimages(output_path)

        return output_path

    def _detect_template(self, ws) -> tuple:
        max_col = min(30, ws.max_column or 0)
        best_row, best_score = 0, 0
        for r in range(1, min(31, ws.max_row + 1)):
            score = 0
            for c in range(1, max_col + 1):
                val = str(ws.cell(row=r, column=c).value or "").strip().lower()
                for _, kws in PATTERNS:
                    for kw in kws:
                        if kw in val:
                            score += 1
                            break
            if score > best_score:
                best_score, best_row = score, r
        if best_row == 0:
            best_row = 3
        col_map = {"_max_col": max_col}
        for c in range(1, max_col + 1):
            val = str(ws.cell(row=best_row, column=c).value or "").strip().lower()
            if not val:
                continue
            for field, kws in PATTERNS:
                for kw in kws:
                    if kw in val:
                        if field not in col_map:
                            col_map[field] = c
                        break
        return best_row, col_map

    @staticmethod
    def _set_col(ws, row, col_idx, value, align=None, font=None, fill=None):
        c = ws.cell(row=row, column=col_idx)
        if isinstance(c, MergedCell):
            return
        c.value = value
        if align:
            c.alignment = align
        if font:
            c.font = font
        if fill:
            c.fill = fill

    def _extract_cell_images(self, wb_path: str) -> dict:
        """{row: (dispimg_name, img_bytes, w_emu, h_emu)}"""
        import zipfile as _zf, xml.etree.ElementTree as _ET, re as _re
        result = {}
        try:
            zf = _zf.ZipFile(wb_path)
            if "xl/cellimages.xml" not in zf.namelist():
                zf.close(); return result
            ci_xml = zf.read("xl/cellimages.xml")
            rels_xml = zf.read("xl/_rels/cellimages.xml.rels")
        except Exception:
            return result
        try:
            rels_root = _ET.fromstring(rels_xml)
            rid_to_media = {rel.get("Id"): rel.get("Target") for rel in rels_root}
            ns_xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
            ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            ci_root = _ET.fromstring(ci_xml)
            name_info = {}
            for ci in ci_root:
                cNvPr = ci.find(f'.//{{{ns_xdr}}}cNvPr')
                if cNvPr is None: continue
                name = cNvPr.get("name", "")
                blip = ci.find(f'.//{{{ns_a}}}blip')
                rid = blip.get(f'{{{ns_r}}}embed') if blip is not None else None
                ext = ci.find(f'.//{{{ns_a}}}ext')
                cx = int(ext.get("cx")) if ext is not None else 0
                cy = int(ext.get("cy")) if ext is not None else 0
                name_info[name] = {"rid": rid, "cx": cx, "cy": cy}
            sheet_xml = zf.read("xl/worksheets/sheet1.xml")
            sheet_root = _ET.fromstring(sheet_xml)
            ns_s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            for cell in sheet_root.iter(f"{{{ns_s}}}c"):
                ref = cell.get("r", "")
                m = _re.match(r"[A-Z]+(\d+)", ref)
                if not m: continue
                row = int(m.group(1))
                for f in cell.iter(f"{{{ns_s}}}f"):
                    if f.text and "DISPIMG" in f.text:
                        img_match = _re.search(r'"([^"]+)"', f.text)
                        if img_match:
                            ni = name_info.get(img_match.group(1))
                            if ni and ni["rid"] in rid_to_media:
                                img_bytes = zf.read("xl/" + rid_to_media[ni["rid"]])
                                result[row] = (img_match.group(1), img_bytes, ni["cx"], ni["cy"])
        finally:
            zf.close()
        return result

    def _inject_cellimages(self, output_path: str):
        import zipfile as _zf, os as _os, shutil as _sh
        tmp_path = output_path + ".tmp"
        WPS_CELLIMAGE_TYPE = "http://www.wps.cn/officeDocument/2020/cellImage"
        try:
            with _zf.ZipFile(output_path, 'r') as zin,                  _zf.ZipFile(tmp_path, 'w', _zf.ZIP_DEFLATED) as zout:
                # Find next available rId from workbook.xml.rels
                wb_rels = zin.read("xl/_rels/workbook.xml.rels").decode()
                max_rid = 0
                import re as _re
                for m in _re.finditer(r'Id="rId(\d+)"', wb_rels):
                    max_rid = max(max_rid, int(m.group(1)))
                next_rid = max_rid + 1
                wb_rels_patched = wb_rels.replace(
                    "</Relationships>",
                    f'<Relationship Id="rId{next_rid}" Type="{WPS_CELLIMAGE_TYPE}" '
                    f'Target="cellimages.xml"/></Relationships>')
                for item in zin.infolist():
                    if item.filename == "[Content_Types].xml":
                        ct_xml = zin.read(item.filename).decode()
                        if "/xl/cellimages.xml" not in ct_xml:
                            ct_xml = ct_xml.replace(
                                "</Types>",
                                '<Override PartName="/xl/cellimages.xml" '
                                'ContentType="application/vnd.ms-excel.cellimages+xml"/></Types>')
                        zout.writestr(item, ct_xml)
                    elif item.filename == "xl/_rels/workbook.xml.rels":
                        zout.writestr(item, wb_rels_patched)
                    else:
                        zout.writestr(item, zin.read(item.filename))
                # Inject cellimages + media from template
                with _zf.ZipFile(str(self._template_path), 'r') as tpl:
                    for name in tpl.namelist():
                        if 'cellimage' in name.lower():
                            zout.writestr(name, tpl.read(name))
                    for name in tpl.namelist():
                        if name.startswith('xl/media/') and tpl.getinfo(name).file_size > 0:
                            if name not in {i.filename for i in zin.infolist()}:
                                zout.writestr(name, tpl.read(name))
            _sh.move(tmp_path, output_path)
        except Exception:
            if _os.path.exists(tmp_path):
                _os.unlink(tmp_path)
