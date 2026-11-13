import sys
import os
import io

# Safeguard sys.stdout and sys.stderr in --noconsole windowed mode
class SafeWriter:
    def write(self, s): pass
    def flush(self): pass
    def isatty(self): return False

if sys.stdout is None:
    sys.stdout = SafeWriter()
elif not hasattr(sys.stdout, "isatty"):
    sys.stdout.isatty = lambda: False

if sys.stderr is None:
    sys.stderr = SafeWriter()
elif not hasattr(sys.stderr, "isatty"):
    sys.stderr.isatty = lambda: False

import time
import datetime
from typing import Optional, Dict, Any, List
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, StreamingResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import pandas as pd

from engine.comparator import InventoryComparator
from engine.report_generator import ReportGenerator
from logger_util import get_system_logger, PRIMARY_LOG_FILE

logger = get_system_logger()

app = FastAPI(title="FTI IT Asset Inventory Compare System", version="2.0.0")

# Request Logging Middleware (Full Debug)
@app.middleware("http")
async def log_requests_middleware(request: Request, call_next):
    start_time = time.time()
    path = request.url.path
    is_heartbeat = path in ("/api/heartbeat", "/api/client_unloading")
    client_ip = request.client.host if request.client else "127.0.0.1"

    try:
        response = await call_next(request)
        elapsed_ms = (time.time() - start_time) * 1000
        if not is_heartbeat:
            logger.info(f"{request.method:4} {path:22} -> {response.status_code} [{elapsed_ms:.1f}ms] (Client: {client_ip})", component="HTTP")
        return response
    except Exception as e:
        elapsed_ms = (time.time() - start_time) * 1000
        logger.error(f"{request.method:4} {path:22} -> FAILED: {e} [{elapsed_ms:.1f}ms]", component="HTTP", exc_info=True)
        raise

# Multi-path search for base data directory
candidate_dirs = [
    os.getcwd(),
    os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__)),
    r"E:\WEB\FTI_Inventory Compare Automations",
    os.path.dirname(os.path.abspath(__file__))
]

BASE_DIR = os.getcwd()
for cdir in candidate_dirs:
    if cdir and os.path.exists(os.path.join(cdir, "0-FTI_asset_HWSWOth.xlsx")):
        BASE_DIR = cdir
        break

if getattr(sys, 'frozen', False):
    EXE_DIR = os.path.dirname(sys.executable)
    BUNDLE_DIR = getattr(sys, '_MEIPASS', EXE_DIR)
    STATIC_DIR = os.path.join(BUNDLE_DIR, "static") if os.path.exists(os.path.join(BUNDLE_DIR, "static")) else os.path.join(EXE_DIR, "static")
else:
    STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")

# Mount static files
if not os.path.exists(STATIC_DIR):
    os.makedirs(STATIC_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Global singleton comparator instance
comparator = InventoryComparator()

# Autosave session configuration
AUTOSAVE_DIR = os.path.join(os.environ.get("LOCALAPPDATA", BASE_DIR), "FTI_Inventory_Compare")
os.makedirs(AUTOSAVE_DIR, exist_ok=True)
SOURCES_DIR = os.path.join(AUTOSAVE_DIR, "sources")
os.makedirs(SOURCES_DIR, exist_ok=True)
AUTOSAVE_FILE = os.path.join(AUTOSAVE_DIR, "session_state.json")
LEGACY_AUTOSAVE_FILE = os.path.join(AUTOSAVE_DIR, "session_autosave.json")

comparator.autosave_path = AUTOSAVE_FILE
comparator.sources_dir = SOURCES_DIR

def init_app_session():
    """Load session or defaults. Note: Files are loaded into 'loaded' state WITHOUT auto-comparing."""
    try:
        # 1. Clean legacy
        if os.path.exists(LEGACY_AUTOSAVE_FILE):
            try: os.remove(LEGACY_AUTOSAVE_FILE)
            except Exception: pass

        # 2. Try loading 2.0 autosave
        restored = comparator.load_autosave(AUTOSAVE_FILE)
        if restored:
            logger.info("Restored session from v2 autosave file.", component="AUTOSAVE")
            return

        # 3. First time fallback: load default base Excel files without comparing
        def find_base_file(fname):
            candidates = [
                os.path.join(BASE_DIR, fname),
                os.path.join(os.path.dirname(BASE_DIR), fname),
                os.path.join(os.path.dirname(BASE_DIR), "Inventory_Compare_data", fname),
                os.path.join(BASE_DIR, "Inventory_Compare_data", fname),
                os.path.join(os.getcwd(), fname),
            ]
            for c in candidates:
                if os.path.exists(c):
                    return c
            return None

        base_master = find_base_file("0-FTI_asset_HWSWOth.xlsx")
        base_hw = find_base_file("1-it_asset_HW.xlsx")
        base_sw = find_base_file("2-it_asset_SW.xlsx")
        base_ups = find_base_file("3-it_asset_UPS_Monitor.xlsx")

        if base_master and (base_hw or base_sw or base_ups):
            logger.info("First-time session: Loading default base Excel files into loaded state (ready for manual compare)...", component="SESSION")
            import shutil
            if base_master:
                comparator.master_filename = "0-FTI_asset_HWSWOth.xlsx"
                comparator.load_master(base_master)
                shutil.copy2(base_master, os.path.join(SOURCES_DIR, "source_master.xlsx"))
            if base_hw:
                comparator.load_category("hw", base_hw, filename="1-it_asset_HW.xlsx")
                shutil.copy2(base_hw, os.path.join(SOURCES_DIR, "source_hw.xlsx"))
            if base_sw:
                comparator.load_category("sw", base_sw, filename="2-it_asset_SW.xlsx")
                shutil.copy2(base_sw, os.path.join(SOURCES_DIR, "source_sw.xlsx"))
            if base_ups:
                comparator.load_category("ups", base_ups, filename="3-it_asset_UPS_Monitor.xlsx")
                shutil.copy2(base_ups, os.path.join(SOURCES_DIR, "source_ups.xlsx"))

            comparator.autosave()
            logger.info("Default files loaded in 'loaded' state.", component="SESSION")
    except Exception as e:
        logger.error(f"Error initializing session: {e}", component="SESSION", exc_info=True)

init_app_session()

# Pydantic models
class CellEditRequest(BaseModel):
    uid: str
    column: str
    value: Any
    user: Optional[str] = "IT Officer"

class RowExclusionRequest(BaseModel):
    uid: str
    exclude: bool
    user: Optional[str] = "IT Officer"

class RollbackRequest(BaseModel):
    index: int

last_heartbeat_time = datetime.datetime.now().timestamp()
unloading_time: Optional[float] = None
server_ref = None

def get_last_heartbeat() -> float:
    global last_heartbeat_time
    return last_heartbeat_time

def get_unloading_time() -> Optional[float]:
    global unloading_time
    return unloading_time

def set_server_instance(srv):
    global server_ref
    server_ref = srv

@app.get("/api/heartbeat")
async def get_heartbeat():
    global last_heartbeat_time, unloading_time
    last_heartbeat_time = datetime.datetime.now().timestamp()
    unloading_time = None
    return {"status": "ok", "time": last_heartbeat_time}

@app.post("/api/client_unloading")
async def notify_client_unloading():
    global unloading_time
    unloading_time = datetime.datetime.now().timestamp()
    try:
        comparator.flush_autosave()
    except Exception:
        pass
    return {"status": "acknowledged", "time": unloading_time}

@app.post("/api/shutdown")
async def request_shutdown(force: bool = False):
    global server_ref
    try:
        comparator.flush_autosave()
    except Exception:
        pass
    if force and server_ref is not None:
        logger.info("Explicit shutdown requested by user.", component="SYSTEM")
        server_ref.should_exit = True
        return {"status": "shutting down"}
    return {"status": "ignored"}

@app.get("/api/logs")
async def get_system_logs(lines: int = Query(default=150, ge=1, le=2000)):
    if not os.path.exists(PRIMARY_LOG_FILE):
        return JSONResponse({"log_file": PRIMARY_LOG_FILE, "lines": []})
    try:
        with open(PRIMARY_LOG_FILE, "r", encoding="utf-8", errors="replace") as f:
            all_lines = f.readlines()
            return JSONResponse({
                "log_file": PRIMARY_LOG_FILE,
                "total_lines": len(all_lines),
                "lines": [l.rstrip("\r\n") for l in all_lines[-lines:]]
            })
    except Exception as e:
        return JSONResponse({"log_file": PRIMARY_LOG_FILE, "error": str(e), "lines": []})

@app.get("/api/logs/download")
async def download_system_log():
    if os.path.exists(PRIMARY_LOG_FILE):
        return FileResponse(PRIMARY_LOG_FILE, filename="fti_debug.log", media_type="text/plain")
    raise HTTPException(status_code=404, detail="Log file not found")

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse("<h1>FTI IT Asset Inventory Compare System</h1><p>UI loading...</p>")

@app.get("/api/version")
async def get_system_version():
    """Retrieve application version and build metadata."""
    v_path = os.path.join(BASE_DIR, "version.json")
    if not os.path.exists(v_path):
        v_path = os.path.join(STATIC_DIR, "version.json")
    if os.path.exists(v_path):
        try:
            with open(v_path, "r", encoding="utf-8") as f:
                return JSONResponse(json.load(f))
        except Exception:
            pass
    return JSONResponse({
        "version": "2.1.0",
        "build_number": 1,
        "build_date": "2026.11.01",
        "compiled_at": "2026-11-01 14:35:00",
        "display_version": "v2.1.0 (Build 2026.11.01-01)"
    })

# =========================================================================
# V2 System State & Comparison API
# =========================================================================

@app.get("/api/state")
async def get_system_state():
    """Return unified application state: file status, category states, summary, and undo/redo capabilities."""
    files_status = {
        "master": {
            "loaded": len(comparator.master_dict) > 0 or comparator.master_df is not None,
            "rows": len(comparator.master_dict) if len(comparator.master_dict) > 0 else (len(comparator.master_df) if comparator.master_df is not None else 0),
            "filename": getattr(comparator, "master_filename", "Master ERP AX Data" if len(comparator.master_dict) > 0 else None)
        },
        "hw": {
            "loaded": len(comparator.records.get("hw", [])) > 0,
            "rows": len(comparator.records.get("hw", [])),
            "filename": getattr(comparator, "hw_filename", "1-it_asset_HW.xlsx" if len(comparator.records.get("hw", [])) > 0 else None),
            "state": comparator.category_states.get("hw", "empty")
        },
        "sw": {
            "loaded": len(comparator.records.get("sw", [])) > 0,
            "rows": len(comparator.records.get("sw", [])),
            "filename": getattr(comparator, "sw_filename", "2-it_asset_SW.xlsx" if len(comparator.records.get("sw", [])) > 0 else None),
            "state": comparator.category_states.get("sw", "empty")
        },
        "ups": {
            "loaded": len(comparator.records.get("ups", [])) > 0,
            "rows": len(comparator.records.get("ups", [])),
            "filename": getattr(comparator, "ups_filename", "3-it_asset_UPS_Monitor.xlsx" if len(comparator.records.get("ups", [])) > 0 else None),
            "state": comparator.category_states.get("ups", "empty")
        }
    }

    return JSONResponse({
        "files": files_status,
        "category_states": comparator.category_states,
        "summary": comparator.summary,
        "can_undo": len(comparator.undo_stack) > 0,
        "can_redo": len(comparator.redo_stack) > 0,
        "history_count": len(comparator.undo_stack)
    })

@app.get("/api/summary")
async def get_summary():
    """Return summary for backward compatibility."""
    return JSONResponse(comparator.summary)

@app.post("/api/compare/{category}")
async def compare_single_category(category: str, reset_to_orig: bool = Query(False, description="Reset automatic cells to original file values before comparing")):
    """Trigger explicit comparison for a specific category."""
    if category not in ("hw", "sw", "ups"):
        raise HTTPException(status_code=400, detail="Category must be hw, sw, or ups")

    if len(comparator.master_dict) == 0:
        raise HTTPException(status_code=400, detail="กรุณาอัปโหลดไฟล์หลัก ERP (Master) ก่อนทำการเปรียบเทียบ")

    if len(comparator.records.get(category, [])) == 0:
        raise HTTPException(status_code=400, detail=f"ยังไม่มีข้อมูลไฟล์ {category.upper()} สำหรับเปรียบเทียบ")

    cat_summary = comparator.compare_category(category, reset_to_orig=reset_to_orig)
    logger.info(f"Explicit comparison completed for {category.upper()}: {cat_summary.get('match_rate')}% match rate (reset_to_orig={reset_to_orig})", component="COMPARE")

    return JSONResponse({
        "status": "success",
        "category": category,
        "category_state": comparator.category_states.get(category),
        "summary": comparator.summary,
        "category_summary": cat_summary
    })

@app.post("/api/compare-all")
async def compare_all_categories(reset_to_orig: bool = Query(True, description="Reset automatic cells to original file values to get baseline comparison")):
    """Trigger comparison across all 3 categories (HW, SW, UPS). Default reset_to_orig=True for baseline comparison."""
    if len(comparator.master_dict) == 0:
        raise HTTPException(status_code=400, detail="กรุณาอัปโหลดไฟล์หลัก ERP (Master) ก่อนทำการเปรียบเทียบ")

    results = {}
    for cat in ("hw", "sw", "ups"):
        if len(comparator.records.get(cat, [])) > 0:
            results[cat] = comparator.compare_category(cat, reset_to_orig=reset_to_orig)

    logger.info(f"Explicit Compare All completed across all categories (reset_to_orig={reset_to_orig}): overall rate {comparator.summary.get('overall', {}).get('match_rate')}%", component="COMPARE")

    return JSONResponse({
        "status": "success",
        "category_states": comparator.category_states,
        "summary": comparator.summary,
        "by_category": results
    })

@app.get("/api/grid/{category}")
async def get_category_grid(
    category: str,
    filter_type: str = Query("all", description="all, mismatch, match, excluded, user_edited"),
    search: str = Query("", description="search query"),
    page: int = Query(1, ge=1),
    page_size: int = Query(50, ge=1, le=100000)
):
    """Retrieve full spreadsheet grid for category: column schema and paginated rows with full cells and ERP comparison."""
    if category not in ("hw", "sw", "ups"):
        raise HTTPException(status_code=400, detail="Category must be hw, sw, or ups")

    all_recs = comparator.records.get(category, [])
    cols = comparator.columns.get(category, [])
    if not cols and all_recs:
        cols = list(all_recs[0]["cells"].keys())

    # Filter logic
    filtered = all_recs
    if filter_type == "mismatch":
        mismatches = [r for r in filtered if r.get("is_match") is False or r.get("is_excluded") is True]
        if not mismatches:
            # If all active have been reconciled, show records with diffs/user edits so user can review/edit
            mismatches = [r for r in filtered if r.get("diff_fields") or r.get("user_edited_cols") or r.get("is_excluded")]
        filtered = mismatches
    elif filter_type == "match":
        filtered = [r for r in filtered if not r.get("is_excluded") and r.get("is_match") is True]
    elif filter_type == "excluded":
        filtered = [r for r in filtered if r.get("is_excluded") is True]
    elif filter_type == "user_edited":
        filtered = [r for r in filtered if len(r.get("user_edited_cols", [])) > 0]

    # Search query across all cells and metadata
    if search:
        q = search.lower().strip()
        def matches_rec(r):
            # check cells
            for val in r.get("cells", {}).values():
                if q in str(val).lower():
                    return True
            # check ax_data
            ax = r.get("ax_data", {})
            for val in ax.values():
                if q in str(val).lower():
                    return True
            # check mismatch/exclude reasons
            if q in str(r.get("mismatch_type", "")).lower() or q in str(r.get("exclude_reason", "")).lower():
                return True
            return False
        filtered = [r for r in filtered if matches_rec(r)]

    total_count = len(filtered)
    offset = (page - 1) * page_size
    paginated_recs = filtered[offset: offset + page_size]

    for r in paginated_recs:
        aid, emp_id, emp_name, dept, sn = comparator._get_key_fields(category, r.get("cells", {}))
        r["asset_id"] = aid
        r["it_emp_id"] = emp_id
        r["it_emp_name"] = emp_name
        r["it_dept"] = dept
        r["it_sn"] = sn

    return JSONResponse({
        "category": category,
        "category_state": comparator.category_states.get(category, "empty"),
        "columns": cols,
        "total": total_count,
        "total_count": total_count,
        "page": page,
        "page_size": page_size,
        "total_pages": (total_count + page_size - 1) // page_size if total_count > 0 else 1,
        "records": paginated_recs
    })

@app.patch("/api/cell")
async def update_cell_value(req: CellEditRequest):
    """Update a single cell in spreadsheet mode."""
    updated = comparator.update_cell(req.uid, req.column, req.value, user=req.user or "IT Officer")
    if not updated:
        raise HTTPException(status_code=404, detail=f"Record {req.uid} not found")

    cat = req.uid.split("_")[0]
    return JSONResponse({
        "status": "success",
        "record": updated,
        "summary": comparator.summary,
        "category_summary": comparator.summary.get("by_category", {}).get(cat, {}),
        "can_undo": len(comparator.undo_stack) > 0,
        "can_redo": len(comparator.redo_stack) > 0
    })

@app.post("/api/row/exclusion")
async def toggle_row_exclusion(req: RowExclusionRequest):
    """Manually move row between active (Tab 3) and excluded (Tab 4)."""
    updated = comparator.set_row_exclusion(req.uid, req.exclude, user=req.user or "IT Officer")
    if not updated:
        raise HTTPException(status_code=404, detail=f"Record {req.uid} not found")

    cat = req.uid.split("_")[0]
    return JSONResponse({
        "status": "success",
        "record": updated,
        "summary": comparator.summary,
        "category_summary": comparator.summary.get("by_category", {}).get(cat, {}),
        "can_undo": len(comparator.undo_stack) > 0,
        "can_redo": len(comparator.redo_stack) > 0
    })

@app.post("/api/auto-fix/{category}")
async def auto_fix_category(category: str):
    """Run Auto-Fix on a specific category (or all) respecting user-edited cells."""
    if category not in ("hw", "sw", "ups", "all"):
        raise HTTPException(status_code=400, detail="Category must be hw, sw, ups, or all")

    res = comparator.batch_quick_fix(category=category, user="IT Officer")

    # Persist updated 4-tab Excel file
    try:
        out_dir = os.path.join(BASE_DIR, "output")
        ReportGenerator.save_all_category_files(comparator, output_dir=out_dir)
    except Exception as e:
        logger.warn(f"Notice saving Excel after auto-fix: {e}", component="EXPORT")

    return JSONResponse({
        "status": "success",
        "result": res,
        "summary": comparator.summary,
        "can_undo": len(comparator.undo_stack) > 0,
        "can_redo": len(comparator.redo_stack) > 0
    })

# =========================================================================
# File Upload & Ingestion (Manual Compare: Stale/Loaded States)
# =========================================================================

@app.post("/api/upload")
async def upload_files(
    file_type: str = Form(..., description="master, hw, sw, ups"),
    file: UploadFile = File(...)
):
    """Upload new Excel file dynamically into 'loaded' state (does NOT auto-compare)."""
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Only .xlsx or .xls files are supported")

    content = await file.read()
    try:
        df = pd.read_excel(io.BytesIO(content))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to parse Excel file: {str(e)}")

    try:
        src_target = os.path.join(SOURCES_DIR, f"source_{file_type}.xlsx")
        with open(src_target, "wb") as f_src:
            f_src.write(content)
    except Exception as e_save:
        logger.warn(f"Could not persist uploaded file to sources: {e_save}", component="UPLOAD")

    if file_type == "master":
        comparator.master_filename = file.filename
        comparator.load_master(df)
        msg = f"โหลดไฟล์หลัก ERP สำเร็จ ({len(df):,} รายการ) — ทุกหมวดพร้อมสำหรับเปรียบเทียบใหม่"
    elif file_type in ("hw", "sw", "ups"):
        comparator.load_category(file_type, df, filename=file.filename)
        msg = f"อัปโหลดไฟล์ {file_type.upper()} สำเร็จ ({len(df):,} รายการ) — พร้อมให้กดเปรียบเทียบ"
    else:
        raise HTTPException(status_code=400, detail="Invalid file_type")

    return JSONResponse({
        "status": "success",
        "message": msg,
        "filename": file.filename,
        "category_states": comparator.category_states,
        "summary": comparator.summary
    })

# =========================================================================
# Google Workspace-style Undo, Redo, Version History & Session Reset
# =========================================================================
@app.post("/api/undo")
async def undo_action():
    res = comparator.undo()
    return JSONResponse(res)

@app.post("/api/redo")
async def redo_action():
    res = comparator.redo()
    return JSONResponse(res)

@app.get("/api/history")
async def get_history_timeline():
    return JSONResponse({
        "history": comparator.get_version_history(),
        "can_undo": len(comparator.undo_stack) > 0,
        "can_redo": len(comparator.redo_stack) > 0
    })

@app.post("/api/rollback")
async def rollback_to_version(req: RollbackRequest):
    res = comparator.rollback_to_index(req.index)
    return JSONResponse(res)

@app.post("/api/session/reset")
async def reset_current_session():
    comparator.reset_session()
    return JSONResponse({
        "success": True,
        "status": "success",
        "message": "ล้างข้อมูลทั้งหมดเรียบร้อยแล้ว พร้อมเริ่มเซสชันใหม่",
        "summary": comparator.summary
    })

@app.get("/api/session/status")
async def get_session_status():
    total = comparator.summary.get('overall', {}).get('total_records', 0) if comparator.summary else 0
    initial_total = comparator.summary.get('overall', {}).get('total_initial_records', 0) if comparator.summary else 0
    matched = comparator.summary.get('overall', {}).get('matched_records', 0) if comparator.summary else 0
    
    path = comparator.autosave_path or ""
    size_kb = 0.0
    last_mod = "-"
    if path and os.path.exists(path):
        try:
            size_kb = round(os.path.getsize(path) / 1024, 1)
            last_mod = pd.Timestamp.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            pass

    import platform
    return JSONResponse({
        "has_data": total > 0 or len(comparator.master_dict) > 0,
        "total_records": total,
        "total_initial_records": initial_total,
        "matched_records": matched,
        "can_undo": len(comparator.undo_stack) > 0,
        "can_redo": len(comparator.redo_stack) > 0,
        "history_count": len(comparator.undo_stack),
        "autosave_path": path,
        "autosave_dir": os.path.dirname(path) if path else "",
        "autosave_size_kb": size_kb,
        "autosave_last_saved": last_mod,
        "python_version": sys.version.split()[0],
        "platform_system": f"{platform.system()} {platform.release()} ({platform.machine()})"
    })

# =========================================================================
# Excel & ZIP Exports (Reflecting 4-Tabs with exact cell edits)
# =========================================================================

@app.get("/api/export/inventory/{category}")
async def export_inventory_file(category: str):
    """Download 4-tab reconciled inventory Excel file."""
    if category not in ("hw", "sw", "ups"):
        raise HTTPException(status_code=400, detail="Category must be hw, sw, or ups")

    if comparator.category_states.get(category) != "compared":
        raise HTTPException(status_code=409, detail=f"กรุณากดเปรียบเทียบหมวด {category.upper()} ก่อนทำการดาวน์โหลดไฟล์")

    try:
        buf = ReportGenerator.export_corrected_inventory_bytes(comparator, category)
        name_map = {
            "hw": "1-it_asset_HW.xlsx",
            "sw": "2-it_asset_SW.xlsx",
            "ups": "3-it_asset_UPS_Monitor.xlsx"
        }
        filename = name_map[category]

        try:
            out_dir = os.path.join(BASE_DIR, "output")
            os.makedirs(out_dir, exist_ok=True)
            with open(os.path.join(out_dir, filename), "wb") as f_out:
                f_out.write(buf.getvalue())
            buf.seek(0)
        except Exception:
            buf.seek(0)

        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        logger.error(f"Error exporting corrected inventory ({category}): {e}", component="EXPORT", exc_info=True)
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการสร้างไฟล์ Inventory: {str(e)}")

@app.get("/api/export/inventory-all")
async def export_all_inventories_zip():
    """Download all compared 4-tab inventory files in a ZIP archive."""
    try:
        any_compared = any(state == "compared" for state in comparator.category_states.values())
        if not any_compared:
            raise HTTPException(status_code=409, detail="ยังไม่มีหมวดหมู่ที่เปรียบเทียบแล้ว กรุณากดเปรียบเทียบอย่างน้อย 1 หมวดหมู่")

        try:
            out_dir = os.path.join(BASE_DIR, "output")
            ReportGenerator.save_all_category_files(comparator, output_dir=out_dir)
        except Exception:
            pass

        zip_buf = ReportGenerator.export_all_inventories_zip_bytes(comparator)
        return StreamingResponse(
            zip_buf,
            media_type="application/zip",
            headers={"Content-Disposition": "attachment; filename=FTI_IT_Asset_100Percent_All_3_Files.zip"}
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error exporting all inventories zip: {e}", component="EXPORT", exc_info=True)
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการสร้างไฟล์ ZIP: {str(e)}")

@app.get("/api/export/mismatch/{category}")
async def export_mismatch_report_file(category: str):
    """Download dedicated Mismatch Report Excel file (e.g. HW_Mismatch_Report.xlsx)."""
    if category not in ("hw", "sw", "ups"):
        raise HTTPException(status_code=400, detail="Category must be hw, sw, or ups")

    if comparator.category_states.get(category) != "compared":
        raise HTTPException(status_code=409, detail=f"กรุณากดเปรียบเทียบหมวด {category.upper()} ก่อนทำการดาวน์โหลดไฟล์ Mismatch Report")

    try:
        buf = ReportGenerator.generate_mismatch_report_bytes(comparator, category)
        name_map = {
            "hw": "HW_Mismatch_Report.xlsx",
            "sw": "SW_Mismatch_Report.xlsx",
            "ups": "UPS_Mismatch_Report.xlsx"
        }
        filename = name_map[category]
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        logger.error(f"Error exporting mismatch report ({category}): {e}", component="EXPORT", exc_info=True)
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการสร้างไฟล์ Mismatch Report: {str(e)}")

@app.get("/api/export/mismatch-all")
async def export_all_mismatch_report_file():
    """Download combined Mismatch Report Excel file across all categories."""
    any_compared = any(state == "compared" for state in comparator.category_states.values())
    if not any_compared:
        raise HTTPException(status_code=409, detail="ยังไม่มีหมวดหมู่ที่เปรียบเทียบแล้ว กรุณากดเปรียบเทียบอย่างน้อย 1 หมวดหมู่")

    try:
        buf = ReportGenerator.generate_mismatch_report_bytes(comparator, "all")
        filename = "All_Mismatch_Report.xlsx"
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        logger.error(f"Error exporting all mismatch report: {e}", component="EXPORT", exc_info=True)
        raise HTTPException(status_code=500, detail=f"เกิดข้อผิดพลาดในการสร้างไฟล์ Mismatch Report รวม: {str(e)}")
