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
    server = HTTPServer(("0.0.0.0", port), BaoHandler)
    print(f"🚀 bao Web 面板 — http://127.0.0.1:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n🛑 服务已停止")
        server.server_close()
