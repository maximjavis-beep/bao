"""Web 面板 — FBA 装箱单生成 HTTP 服务"""
import base64
import io
import json
import os
import shutil
import tempfile
import uuid
import zipfile
from datetime import date
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import unquote

import openpyxl
from openpyxl.styles import PatternFill

from ..core.exporter import FBAExporter
from ..core.weaver import weave_fba
from ..parsers.fba_parser import FBAParser

WEB_DIR = Path(__file__).parent
DOWNLOADS_DIR = WEB_DIR / "downloads"
UPLOAD_DIR = Path(tempfile.gettempdir()) / "bao_uploads"


class BaoHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        p = self.path.split("?")[0]
        if p in ("/", "/index.html"):
            self._send_file("index.html", "text/html; charset=utf-8")
        elif p.startswith("/downloads/"):
            fp = DOWNLOADS_DIR / Path(unquote(p)).name
            if fp.exists():
                data = fp.read_bytes()
                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            else:
                self.send_error(404)
        elif p == "/api/download-zip":
            self._handle_download_zip()
        else:
            self.send_error(404)

    def do_POST(self):
        p = self.path.split("?")[0]
        if p == "/api/parse":
            self._handle_parse()
        elif p == "/api/weave":
            self._handle_weave()
        elif p == "/api/weave-batch":
            self._handle_weave_batch()
        elif p == "/api/merge-plan":
            self._handle_merge_plan()
        else:
            self.send_error(404)

    def _read(self):
        return self.rfile.read(int(self.headers.get("Content-Length", 0)))

    def _json(self, data, status=200):
        b = json.dumps(data, ensure_ascii=False, default=str).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def _send_file(self, path, ct="application/octet-stream"):
        fp = WEB_DIR / path
        if not fp.exists():
            self.send_error(404)
            return
        data = fp.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", ct)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _save(self, b64, prefix):
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        fp = UPLOAD_DIR / f"{prefix}_{uuid.uuid4().hex[:6]}.xlsx"
        fp.write_bytes(base64.b64decode(b64))
        return str(fp)

    def _handle_parse(self):
        """解析上传的 FBA 货件文件，返回预览数据"""
        try:
            body = json.loads(self._read())
            parser = FBAParser()

            b64 = body.get("file", "")
            if not b64:
                self._json({"success": False, "error": "未提供文件"}, 400)
                return

            path = self._save(b64, "fba")
            data = parser.parse(path)
            meta = data.get("meta", {})
            items = data.get("items", [])

            self._json({
                "success": True,
                "meta": meta,
                "items": items,
                "item_count": len(items),
                "path": path,
            })
        except Exception as e:
            self._json({"success": False, "error": str(e)}, 400)

    def _handle_weave(self):
        """编织生成 FBA 装箱单并返回下载链接"""
        try:
            body = json.loads(self._read())
            parser = FBAParser()

            b64 = body.get("file", "")
            hs_code = body.get("hs_code", None)

            if not b64:
                self._json({"success": False, "error": "未提供文件"}, 400)
                return

            path = self._save(b64, "fba")
            data = parser.parse(path)
            meta = data.get("meta", {})
            sid = meta.get("shipment_id", "UNKNOWN")

            woven = weave_fba(data, hs_code_override=hs_code)

            fname = f"{sid}-装箱单_{uuid.uuid4().hex[:6]}.xlsx"
            out = UPLOAD_DIR / fname
            exporter = FBAExporter()
            exporter.export(woven, str(out))

            DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy2(out, DOWNLOADS_DIR / fname)

            self._json({
                "success": True,
                "download_url": f"/downloads/{fname}",
                "summary": {
                    "shipment_id": woven["shipment_id"],
                    "total_boxes": woven["total_boxes"],
                    "total_weight": woven["total_weight"],
                    "total_cbm": woven["total_cbm"],
                    "row_count": len(woven["rows"]),
                },
                "rows": woven["rows"],
            })
        except Exception as e:
            self._json({"success": False, "error": str(e)}, 400)

    def _handle_weave_batch(self):
        """批量上传多个 FBA 货件，一次性编织并返回下载链接 + zip base64"""
        import subprocess as _sp, sys as _sys
        try:
            body = json.loads(self._read())
            files = body.get("files", [])
            hs_code = body.get("hs_code", None)
            if not files:
                self._json({"success": False, "error": "未提供文件"}, 400)
                return
            parser = FBAParser()
            exporter = FBAExporter()
            DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
            UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
            results = []
            for f in files:
                b64 = f.get("data", "")
                fname_orig = f.get("name", "未知")
                if not b64:
                    results.append({"filename": fname_orig, "success": False, "error": "文件数据为空"})
                    continue
                try:
                    path = self._save(b64, "fba")
                    data = parser.parse(path)
                    meta = data.get("meta", {})
                    sid = meta.get("shipment_id", "UNKNOWN")
                    woven = weave_fba(data, hs_code_override=hs_code)
                    out_fname = f"{sid}-装箱单_{uuid.uuid4().hex[:6]}.xlsx"
                    out = UPLOAD_DIR / out_fname
                    exporter.export(woven, str(out))
                    shutil.copy2(out, DOWNLOADS_DIR / out_fname)
                    results.append({
                        "filename": fname_orig, "success": True,
                        "shipment_id": sid,
                        "download_url": f"/downloads/{out_fname}",
                        "summary": {
                            "shipment_id": woven["shipment_id"],
                            "total_boxes": woven["total_boxes"],
                            "total_weight": woven["total_weight"],
                            "total_cbm": woven["total_cbm"],
                            "row_count": len(woven["rows"]),
                        },
                    })
                except Exception as e:
                    results.append({"filename": fname_orig, "success": False, "error": str(e)})
            # 生成 zip base64（内嵌在响应中，前端直接解码下载）
            zip_b64 = ""
            ok = [r for r in results if r.get("success")]
            if ok:
                try:
                    bid = uuid.uuid4().hex[:6]
                    td = UPLOAD_DIR / f"batch_{bid}"
                    td.mkdir(parents=True, exist_ok=True)
                    for r in ok:
                        src = DOWNLOADS_DIR / Path(r["download_url"]).name
                        if src.exists():
                            shutil.copy2(src, td / src.name)
                    zd = str(UPLOAD_DIR / f"batch_{bid}.zip")
                    try:
                        _sp.run(["zip", "-j", zd] + sorted([f.name for f in td.iterdir()]),
                                 cwd=str(td), check=True, capture_output=True, timeout=30)
                    except (_sp.CalledProcessError, FileNotFoundError):
                        with zipfile.ZipFile(zd, "w") as zf:
                            for fp in sorted(td.iterdir()):
                                zf.write(str(fp), fp.name)
                    zip_b64 = base64.b64encode(Path(zd).read_bytes()).decode()
                    shutil.rmtree(td, ignore_errors=True)
                    Path(zd).unlink(missing_ok=True)
                except Exception:
                    pass
            self._json({"success": True, "results": results, "zip_b64": zip_b64})
        except Exception as e:
            self._json({"success": False, "error": str(e)}, 400)

    def _handle_merge_plan(self):
        """上传调整明细 Excel，按规则映射生成发货计划"""
        try:
            body = json.loads(self._read())
            b64 = body.get("file", "")
            if not b64:
                self._json({"success": False, "error": "未提供文件"}, 400)
                return

            # 保存上传的调整明细
            UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
            adj_path = UPLOAD_DIR / f"adj_{uuid.uuid4().hex[:6]}.xlsx"
            adj_path.write_bytes(base64.b64decode(b64))

            # 解析调整明细
            wb_adj = openpyxl.load_workbook(str(adj_path), data_only=True)
            ws_adj = wb_adj.active
            adj_rows = []
            for r in range(2, ws_adj.max_row + 1):
                sku = ws_adj.cell(row=r, column=5).value
                code = ws_adj.cell(row=r, column=6).value
                if code is None:
                    continue
                adj_rows.append({
                    "识别码": str(code).strip(),
                    "店铺": str(ws_adj.cell(row=r, column=7).value or "").strip(),
                    "SKU": str(sku or "").strip(),
                    "调整FNSKU": str(ws_adj.cell(row=r, column=12).value or "").strip(),
                    "调整量": ws_adj.cell(row=r, column=10).value,
                    "调整店铺": str(ws_adj.cell(row=r, column=11).value or "").strip(),
                })
            wb_adj.close()

            if not adj_rows:
                self._json({"success": False, "error": "调整明细无有效数据"}, 400)
                return

            # 加载发货计划模板
            template_dir = Path(__file__).parent.parent.parent / "templates"
            template_path = template_dir / "发货计划-模板.xlsx"
            if not template_path.exists():
                template_path = template_dir / "发货计划-模板 .xlsx"
            if not template_path.exists():
                self._json({"success": False, "error": f"找不到模板文件: {template_path}"}, 500)
                return

            wb_plan = openpyxl.load_workbook(str(template_path))
            ws_plan = wb_plan.active

            # 找到最后一个有内容的行
            last_row = 1
            for r in range(2, ws_plan.max_row + 1):
                has_data = False
                for c in range(1, ws_plan.max_column + 1):
                    if ws_plan.cell(row=r, column=c).value is not None:
                        has_data = True
                        break
                if has_data:
                    last_row = r

            # 黄色填充
            yellow = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

            # 按规则映射写入
            start_row = last_row + 1
            for i, adj in enumerate(adj_rows):
                tr = start_row + i
                d_val = adj["店铺"]  # D列来源

                # C列: 识别码
                ws_plan.cell(row=tr, column=3).value = adj["识别码"]
                # D列: 店铺-国家 = 调整明细 店铺
                ws_plan.cell(row=tr, column=4).value = d_val
                # E列: 国家 = D列中"-"后面的代码
                country = d_val.split("-")[-1] if "-" in d_val else ""
                ws_plan.cell(row=tr, column=5).value = country
                # F列: SKU
                ws_plan.cell(row=tr, column=6).value = adj["SKU"]
                # G列: FNSKU = 调整明细 调整FNSKU
                ws_plan.cell(row=tr, column=7).value = adj["调整FNSKU"]
                # J列: 计划发货量 = 调整量
                qty = int(adj["调整量"]) if adj["调整量"] is not None else None
                ws_plan.cell(row=tr, column=10).value = qty
                # K列: 库存数 = 调整量
                ws_plan.cell(row=tr, column=11).value = qty
                # P列: 店铺 = 调整明细 调整店铺
                ws_plan.cell(row=tr, column=16).value = adj["调整店铺"]

                # 整行黄色标注
                for c in range(1, ws_plan.max_column + 1):
                    ws_plan.cell(row=tr, column=c).fill = yellow

            # 保存输出
            out_name = f"发货计划_{uuid.uuid4().hex[:6]}.xlsx"
            out_path = UPLOAD_DIR / out_name
            wb_plan.save(str(out_path))
            wb_plan.close()

            DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy2(out_path, DOWNLOADS_DIR / out_name)

            self._json({
                "success": True,
                "download_url": f"/downloads/{out_name}",
                "row_count": len(adj_rows),
            })
        except Exception as e:
            self._json({"success": False, "error": str(e)}, 400)

    def _handle_download_zip(self):
        """将所有已生成的装箱单打包为 zip 下载"""
        import subprocess, sys, traceback
        try:
            DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
            xlsx_files = sorted(DOWNLOADS_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
            if not xlsx_files:
                self.send_error(404)
                return

            UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
            batch_id = uuid.uuid4().hex[:6]
            tmp_dir = UPLOAD_DIR / f"batch_{batch_id}"
            tmp_dir.mkdir(parents=True, exist_ok=True)
            for fp in xlsx_files:
                shutil.copy2(fp, tmp_dir / fp.name)

            zip_dest = str(UPLOAD_DIR / f"batch_{batch_id}.zip")
            try:
                subprocess.run(
                    ["zip", "-j", zip_dest] + sorted([f.name for f in tmp_dir.iterdir()]),
                    cwd=str(tmp_dir), check=True, capture_output=True, timeout=30
                )
            except (subprocess.CalledProcessError, FileNotFoundError):
                with zipfile.ZipFile(zip_dest, "w") as zf:
                    for fp in sorted(tmp_dir.iterdir()):
                        zf.write(str(fp), fp.name)

            data = Path(zip_dest).read_bytes()
            shutil.rmtree(tmp_dir, ignore_errors=True)
            Path(zip_dest).unlink(missing_ok=True)

            self.send_response(200)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Disposition", "attachment; filename=装箱单_批量.zip")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as e:
            traceback.print_exc(file=sys.stderr)
            self.send_error(500)


def run_server(port: int = 8888):
    import socketserver
    class ReusableServer(HTTPServer):
        allow_reuse_address = True
    server = ReusableServer(("0.0.0.0", port), BaoHandler)
    print(f"🚀 bao Web 面板 — http://127.0.0.1:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n🛑 服务已停止")
        server.server_close()
