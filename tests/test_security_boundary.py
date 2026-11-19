import sys
import os
import io
import json
import time
import pytest
import openpyxl
from starlette.testclient import TestClient

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from app import app, comparator, SOURCES_DIR
from engine.comparator import InventoryComparator
from engine.report_generator import sanitize_excel_value

def ensure_test_files_loaded():
    def find_file(fname):
        candidates = [
            os.path.join(BASE_DIR, fname),
            os.path.join(os.path.dirname(BASE_DIR), fname),
            os.path.join(os.path.dirname(BASE_DIR), "Inventory_Compare_data", fname),
            os.path.join(BASE_DIR, "Inventory_Compare_data", fname),
        ]
        for c in candidates:
            if os.path.exists(c):
                return c
        return None

    base_master = find_file("0-FTI_asset_HWSWOth.xlsx")
    base_hw = find_file("1-it_asset_HW.xlsx")

    if len(comparator.master_dict) == 0 and base_master:
        comparator.load_master(base_master)
    if len(comparator.records.get("hw", [])) == 0 and base_hw:
        comparator.load_category("hw", base_hw, filename="1-it_asset_HW.xlsx")

@pytest.fixture(scope="module")
def client():
    ensure_test_files_loaded()
    return TestClient(app)

def test_path_traversal_protection(client):
    """Verify that path traversal in export or file parameters is rejected."""
    res = client.get("/api/export/inventory/..%2f..%2fwindows%2fwin.ini")
    assert res.status_code in (400, 404)

    res2 = client.get("/api/export/mismatch/..%2f..%2fboot.ini")
    assert res2.status_code in (400, 404)

    res3 = client.get("/api/grid/invalid_category_name")
    assert res3.status_code == 400

def test_cell_injection_and_unicode_handling(client):
    """Verify that Thai Unicode, special chars, and script injection strings are stored safely without code execution."""
    ensure_test_files_loaded()
    client.post("/api/compare/hw")
    res_grid = client.get("/api/grid/hw?page=1&page_size=5")
    assert res_grid.status_code == 200
    records = res_grid.json().get("records", [])
    assert len(records) > 0, "No records found in hw grid"
    
    uid = records[0]["uid"]
    payload = "<script>alert('XSS')</script> ภาษาไทยทดสอบ ' OR '1'='1"

    # Patch cell
    res_patch = client.patch("/api/cell", json={
        "uid": uid,
        "column": "PC_Name",
        "value": payload,
        "user": "SecurityTester"
    })
    assert res_patch.status_code == 200
    assert res_patch.json()["record"]["cells"]["PC_Name"] == payload

    # Read back from grid
    res_verify = client.get("/api/grid/hw?page=1&page_size=5")
    assert res_verify.status_code == 200
    rec_found = next(r for r in res_verify.json()["records"] if r["uid"] == uid)
    assert rec_found["cells"]["PC_Name"] == payload

    # Revert via undo
    res_undo = client.post("/api/undo")
    assert res_undo.status_code == 200


def test_upload_path_traversal_rejected(client):
    """a) POST /api/upload with traversal file_type returns HTTP 400 and writes nothing outside SOURCES_DIR."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ColA"])
    ws.append(["Val1"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    evil_target = os.path.abspath(os.path.join(SOURCES_DIR, "source_../../evil.xlsx"))
    if os.path.exists(evil_target):
        os.remove(evil_target)

    res = client.post(
        "/api/upload",
        data={"file_type": "../../evil"},
        files={"file": ("test.xlsx", buf.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert res.status_code == 400
    assert not os.path.exists(evil_target)

    # Also verify invalid signature on .xlsx is rejected
    res_bad_sig = client.post(
        "/api/upload",
        data={"file_type": "hw"},
        files={"file": ("fake.xlsx", b"NOT_A_ZIP_FILE", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert res_bad_sig.status_code == 400


def test_version_endpoint_reads_real_version_json(client):
    """b) GET /api/version returns actual version.json metadata instead of hardcoded fallback."""
    with open(os.path.join(BASE_DIR, "version.json"), "r", encoding="utf-8") as f:
        expected = json.load(f)

    res = client.get("/api/version")
    assert res.status_code == 200
    data = res.json()
    assert data["build_number"] == expected["build_number"]
    assert data["display_version"] == expected["display_version"]
    assert "Build 2026.11.01-01" not in data["display_version"]


def test_csrf_cross_origin_rejected(client):
    """c) Cross-origin POST /api/shutdown and /api/session/reset with evil Origin return HTTP 403."""
    headers = {"Origin": "https://evil.example"}

    res_shutdown = client.post("/api/shutdown?force=true", headers=headers)
    assert res_shutdown.status_code == 403

    res_reset = client.post("/api/session/reset", headers=headers)
    assert res_reset.status_code == 403

    # Also verify disallowed Referer and null Origin are rejected
    res_ref = client.post("/api/session/reset", headers={"Referer": "https://evil.example/attack"})
    assert res_ref.status_code == 403

    res_null = client.post("/api/session/reset", headers={"Origin": "null"})
    assert res_null.status_code == 403


def test_sanitize_excel_value_formula_injection():
    """d) sanitize_excel_value neutralizes formula payloads while preserving normal text and numbers."""
    assert sanitize_excel_value("=CMD|' /C calc'!A0") == "'=CMD|' /C calc'!A0"
    assert sanitize_excel_value("+1+1") == "'+1+1"
    assert sanitize_excel_value("@SUM(A1)") == "'@SUM(A1)"
    assert sanitize_excel_value("-cmd") == "'-cmd"
    assert sanitize_excel_value("\t=1+1") == "'\t=1+1"
    assert sanitize_excel_value("\r=1+1") == "'\r=1+1"

    # Preserves normal text, ints, floats, and numeric strings
    assert sanitize_excel_value("Normal Asset Name") == "Normal Asset Name"
    assert sanitize_excel_value("อุปกรณ์ IT 01") == "อุปกรณ์ IT 01"
    assert sanitize_excel_value(12345) == 12345
    assert sanitize_excel_value(-45.67) == -45.67
    assert sanitize_excel_value("-123.45") == "-123.45"
    assert sanitize_excel_value("") == ""
    assert sanitize_excel_value(None) is None


def test_reset_session_cancels_pending_autosave_timer(tmp_path):
    """e) reset_session() cancels pending debounced autosave timer so file is not recreated."""
    comp = InventoryComparator()
    autosave_file = tmp_path / "pending_session.json"
    comp.autosave_path = str(autosave_file)

    # Schedule a debounced autosave 0.25s in the future
    comp.autosave(delay=0.25)
    assert comp._autosave_timer is not None

    # Reset immediately before timer fires
    comp.reset_session()
    assert comp._autosave_timer is None
    assert not autosave_file.exists()

    # Wait past the original delay and verify file was NOT written
    time.sleep(0.4)
    assert not autosave_file.exists()
