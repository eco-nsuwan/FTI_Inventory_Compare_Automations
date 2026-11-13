import sys
import os
import pytest
from starlette.testclient import TestClient

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from app import app, comparator

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
