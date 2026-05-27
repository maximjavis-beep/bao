"""bao Web API — FastAPI for Vercel Serverless (secure: uploads deleted after use, downloads via one-shot tokens)"""
import base64, os, shutil, sys, uuid, zipfile
from pathlib import Path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse, HTMLResponse, Response
from fastapi.middleware.cors import CORSMiddleware
from bao.core.exporter import FBAExporter
from bao.core.weaver import weave_fba
from bao.parsers.fba_parser import FBAParser

app = FastAPI(title="bao", version="0.2.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])
ROOT_DIR = Path(__file__).parent.parent
UPLOAD_DIR = Path("/tmp/bao_uploads")
_DOWNLOAD_SLOTS = {}

def _read_html(path): fp = ROOT_DIR / "bao" / "web" / path; return fp.read_text(encoding="utf-8") if fp.exists() else "<h1>404</h1>"
def _make_token(fp):
    if not fp.exists(): return ""
    b64 = base64.b64encode(fp.read_bytes()).decode()
    token = uuid.uuid4().hex; _DOWNLOAD_SLOTS[token] = b64
    fp.unlink(missing_ok=True); return token

@app.get("/", response_class=HTMLResponse)
async def index(): return _read_html("index_vercel.html")

@app.post("/api/parse")
async def api_parse(file: UploadFile = File(...)):
    try:
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        tmp = UPLOAD_DIR / f"fba_{uuid.uuid4().hex[:6]}.xlsx"
        tmp.write_bytes(await file.read())
        data = FBAParser().parse(str(tmp)); tmp.unlink(missing_ok=True)
        items = data.get("items", [])
        return JSONResponse({"success": True, "meta": data.get("meta", {}), "items": items, "item_count": len(items)})
    except Exception as e:
        return JSONResponse({"success": False, "error": str(e)}, 400)

@app.post("/api/weave")
async def api_weave(file: UploadFile = File(...), hs_code: str = Form(None)):
    try:
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        tmp = UPLOAD_DIR / f"fba_{uuid.uuid4().hex[:6]}.xlsx"
        tmp.write_bytes(await file.read())
        data = FBAParser().parse(str(tmp)); tmp.unlink(missing_ok=True)
        meta = data.get("meta", {})
        woven = weave_fba(data, hs_code_override=hs_code or None)
        out = UPLOAD_DIR / f"{meta.get('shipment_id','UNKNOWN')}-装箱单_{uuid.uuid4().hex[:6]}.xlsx"
        FBAExporter().export(woven, str(out))
        token = _make_token(out)
        return JSONResponse({"success": True, "download_token": token, "summary": {
            "shipment_id": woven["shipment_id"], "total_boxes": woven["total_boxes"],
            "total_weight": woven["total_weight"], "total_cbm": woven["total_cbm"],
            "row_count": len(woven["rows"])}, "rows": woven["rows"]})
    except Exception as e:
        return JSONResponse({"success": False, "error": str(e)}, 400)

@app.post("/api/weave-batch")
async def api_weave_batch(files: list[UploadFile] = File(...), hs_code: str = Form(None)):
    try:
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        parser, exporter, results = FBAParser(), FBAExporter(), []
        for f in files:
            try:
                tmp = UPLOAD_DIR / f"fba_{uuid.uuid4().hex[:6]}.xlsx"
                tmp.write_bytes(await f.read())
                d = parser.parse(str(tmp)); tmp.unlink(missing_ok=True)
                meta = d.get("meta", {})
                woven = weave_fba(d, hs_code_override=hs_code or None)
                out = UPLOAD_DIR / f"{meta.get('shipment_id','UNKNOWN')}-装箱单_{uuid.uuid4().hex[:6]}.xlsx"
                exporter.export(woven, str(out))
                token = _make_token(out)
                results.append({"filename": f.filename or "未知", "success": True,
                    "shipment_id": meta.get("shipment_id",""), "download_token": token,
                    "summary": {"shipment_id": woven["shipment_id"], "total_boxes": woven["total_boxes"],
                    "total_weight": woven["total_weight"], "total_cbm": woven["total_cbm"], "row_count": len(woven["rows"])}})
            except Exception as e:
                results.append({"filename": f.filename or "未知", "success": False, "error": str(e)})
        zip_b64 = ""
        ok = [r for r in results if r.get("success")]
        if ok:
            try:
                bid = uuid.uuid4().hex[:6]; td = UPLOAD_DIR / f"batch_{bid}"; td.mkdir(parents=True, exist_ok=True)
                for r in ok:
                    t = r.get("download_token","")
                    if t in _DOWNLOAD_SLOTS: (td / f"{r['shipment_id']}.xlsx").write_bytes(base64.b64decode(_DOWNLOAD_SLOTS[t]))
                zd = str(UPLOAD_DIR / f"batch_{bid}.zip")
                try:
                    import subprocess as sp
                    sp.run(["zip","-j",zd]+sorted([fl.name for fl in td.iterdir()]), cwd=str(td), check=True, capture_output=True, timeout=30)
                except (sp.CalledProcessError, FileNotFoundError):
                    with zipfile.ZipFile(zd,"w") as zf:
                        for fl in sorted(td.iterdir()): zf.write(str(fl), fl.name)
                zip_b64 = base64.b64encode(Path(zd).read_bytes()).decode()
            except Exception:
                pass
            finally:
                shutil.rmtree(td, ignore_errors=True); Path(zd).unlink(missing_ok=True)
        return JSONResponse({"success": True, "results": results, "zip_b64": zip_b64})
    except Exception as e:
        return JSONResponse({"success": False, "error": str(e)}, 400)

@app.get("/api/download/{token}")
async def dl_token(token: str):
    b64 = _DOWNLOAD_SLOTS.pop(token, None)
    if not b64: return Response("Not Found or expired", status_code=404)
    return Response(content=base64.b64decode(b64), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": "attachment; filename=装箱单.xlsx"})
