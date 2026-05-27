"""bao Web API — FastAPI for Vercel Serverless"""
import base64
import os
import shutil
import sys
import uuid
import zipfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import JSONResponse, HTMLResponse, Response
from fastapi.middleware.cors import CORSMiddleware

from bao.core.exporter import FBAExporter
from bao.core.weaver import weave_fba
from bao.parsers.fba_parser import FBAParser

app = FastAPI(title="bao", version="0.2.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

ROOT_DIR = Path(__file__).parent.parent
DOWNLOADS_DIR = Path("/tmp/bao_downloads")
UPLOAD_DIR = Path("/tmp/bao_uploads")


def _read_html(path: str) -> str:
    fp = ROOT_DIR / "bao" / "web" / path
    return fp.read_text(encoding="utf-8") if fp.exists() else "<h1>404</h1>"


@app.get("/", response_class=HTMLResponse)
async def index():
    return _read_html("index_vercel.html")


@app.post("/api/parse")
async def api_parse(file: UploadFile = File(...)):
    try:
        parser = FBAParser()
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        tmp = UPLOAD_DIR / f"fba_{uuid.uuid4().hex[:6]}.xlsx"
        tmp.write_bytes(await file.read())
        data = parser.parse(str(tmp))
        return JSONResponse({
            "success": True,
            "meta": data.get("meta", {}),
            "items": data.get("items", []),
            "item_count": len(data.get("items", [])),
        })
    except Exception as e:
        return JSONResponse({"success": False, "error": str(e)}, 400)


@app.post("/api/weave")
async def api_weave(file: UploadFile = File(...), hs_code: str = Form(None)):
    try:
        parser = FBAParser()
        exporter = FBAExporter()
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        tmp = UPLOAD_DIR / f"fba_{uuid.uuid4().hex[:6]}.xlsx"
        tmp.write_bytes(await file.read())
        data = parser.parse(str(tmp))
        meta = data.get("meta", {})
        sid = meta.get("shipment_id", "UNKNOWN")
        woven = weave_fba(data, hs_code_override=hs_code or None)
        fname = f"{sid}-装箱单_{uuid.uuid4().hex[:6]}.xlsx"
        out = UPLOAD_DIR / fname
        exporter.export(woven, str(out))
        DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
        shutil.copy2(out, DOWNLOADS_DIR / fname)
        return JSONResponse({
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
        return JSONResponse({"success": False, "error": str(e)}, 400)


@app.post("/api/weave-batch")
async def api_weave_batch(files: list[UploadFile] = File(...), hs_code: str = Form(None)):
    try:
        parser = FBAParser()
        exporter = FBAExporter()
        DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        results = []
        for f in files:
            try:
                tmp = UPLOAD_DIR / f"fba_{uuid.uuid4().hex[:6]}.xlsx"
                tmp.write_bytes(await f.read())
                d = parser.parse(str(tmp))
                meta = d.get("meta", {})
                sid = meta.get("shipment_id", "UNKNOWN")
                woven = weave_fba(d, hs_code_override=hs_code or None)
                out_fname = f"{sid}-装箱单_{uuid.uuid4().hex[:6]}.xlsx"
                out = UPLOAD_DIR / out_fname
                exporter.export(woven, str(out))
                shutil.copy2(out, DOWNLOADS_DIR / out_fname)
                results.append({
                    "filename": f.filename or "未知", "success": True,
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
                results.append({"filename": f.filename or "未知", "success": False, "error": str(e)})

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
                    import subprocess as sp
                    sp.run(["zip", "-j", zd] + sorted([fl.name for fl in td.iterdir()]),
                           cwd=str(td), check=True, capture_output=True, timeout=30)
                except (sp.CalledProcessError, FileNotFoundError):
                    with zipfile.ZipFile(zd, "w") as zf:
                        for fl in sorted(td.iterdir()):
                            zf.write(str(fl), fl.name)
                zip_b64 = base64.b64encode(Path(zd).read_bytes()).decode()
                shutil.rmtree(td, ignore_errors=True)
                Path(zd).unlink(missing_ok=True)
            except Exception:
                pass

        return JSONResponse({"success": True, "results": results, "zip_b64": zip_b64})
    except Exception as e:
        return JSONResponse({"success": False, "error": str(e)}, 400)


@app.get("/downloads/{filename}")
async def download_file(filename: str):
    fp = DOWNLOADS_DIR / Path(filename).name
    if not fp.exists():
        return Response("Not Found", status_code=404)
    return Response(
        content=fp.read_bytes(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
