import sys
import os
import io
import json
import openpyxl
import pandas as pd
from starlette.testclient import TestClient

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Add root directory to sys.path
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from app import app, comparator, AUTOSAVE_FILE, SOURCES_DIR

def test_manual_compare_and_spreadsheet_workflow():
    print("[1] Initializing TestClient...")
    client = TestClient(app)

    # 1. Start fresh and load master + categories
    comparator.reset_session()
    def find_file(fname):
        p1 = os.path.join(BASE_DIR, fname)
        if os.path.exists(p1): return p1
        p2 = os.path.join(os.path.dirname(BASE_DIR), fname)
        if os.path.exists(p2): return p2
        p3 = os.path.join(os.path.dirname(BASE_DIR), "Inventory_Compare_data", fname)
        if os.path.exists(p3): return p3
        return p1

    base_master = find_file("0-FTI_asset_HWSWOth.xlsx")
    base_hw = find_file("1-it_asset_HW.xlsx")
    base_sw = find_file("2-it_asset_SW.xlsx")
    base_ups = find_file("3-it_asset_UPS_Monitor.xlsx")

    print("[2] Loading files without auto-comparing...")
    comparator.load_master(base_master)
    comparator.load_category("hw", base_hw, filename="1-it_asset_HW.xlsx")
    comparator.load_category("sw", base_sw, filename="2-it_asset_SW.xlsx")
    comparator.load_category("ups", base_ups, filename="3-it_asset_UPS_Monitor.xlsx")

    # Verify initial states
    res_st = client.get("/api/state")
    assert res_st.status_code == 200
    st_data = res_st.json()
    assert st_data["category_states"]["hw"] == "loaded"
    assert st_data["category_states"]["sw"] == "loaded"
    assert st_data["category_states"]["ups"] == "loaded"
    # Overall summary should be 0 matched because none is compared yet
    assert st_data["summary"]["overall"]["matched_records"] == 0
    print("    Verified: files loaded in 'loaded' state with 0 auto-comparisons executed.")

    # 2. Test manual comparison for HW only
    print("[3] Testing explicit POST /api/compare/hw...")
    res_comp_hw = client.post("/api/compare/hw")
    assert res_comp_hw.status_code == 200
    hw_comp_data = res_comp_hw.json()
    assert hw_comp_data["category_state"] == "compared"
    # HW has 235 total -> initial matched 81, mismatched 12, excluded 142 before auto-fix
    cat_hw_sum = hw_comp_data["category_summary"]
    assert cat_hw_sum["initial_total"] == 235
    assert cat_hw_sum["total"] == 93
    assert cat_hw_sum["matched"] == 81
    assert cat_hw_sum["excluded"] == 142
    print(f"    HW Compared explicitly: matched={cat_hw_sum['matched']}, excluded={cat_hw_sum['excluded']}")

    # Check that SW and UPS are STILL in 'loaded' state (isolated lifecycle)
    res_st2 = client.get("/api/state")
    st2 = res_st2.json()
    assert st2["category_states"]["hw"] == "compared"
    assert st2["category_states"]["sw"] == "loaded"
    assert st2["category_states"]["ups"] == "loaded"
    print("    Verified: SW and UPS remain untouched in 'loaded' state.")

    # 3. Test Spreadsheet Grid API (/api/grid/hw)
    print("[4] Testing GET /api/grid/hw...")
    res_grid = client.get("/api/grid/hw?page=1&page_size=10")
    assert res_grid.status_code == 200
    grid_data = res_grid.json()
    assert len(grid_data["columns"]) == 25, f"Expected 25 HW columns, got {len(grid_data['columns'])}"
    assert len(grid_data["records"]) == 10
    first_rec = next(r for r in grid_data["records"] if not r.get("is_excluded"))
    assert "cells" in first_rec
    assert "user_edited_cols" in first_rec
    assert "ax_data" in first_rec
    print(f"    Grid retrieved: {len(grid_data['columns'])} full columns preserved per row.")

    # 4. Test In-Cell Editing (/api/cell)
    print("[5] Testing PATCH /api/cell...")
    target_uid = first_rec["uid"]
    target_col = "PC_Name"
    new_val = "CUSTOM_TEST_PC_999"

    res_cell = client.patch("/api/cell", json={
        "uid": target_uid,
        "column": target_col,
        "value": new_val,
        "user": "Tester"
    })
    assert res_cell.status_code == 200
    cell_data = res_cell.json()
    updated_rec = cell_data["record"]
    assert updated_rec["cells"][target_col] == new_val
    assert target_col in updated_rec["user_edited_cols"]
    print(f"    Cell edit successful: {target_col} = {new_val}, marked in user_edited_cols.")

    # 5. Test Undo / Redo for cell editing
    print("[6] Testing Undo / Redo on cell edit...")
    res_undo = client.post("/api/undo")
    assert res_undo.status_code == 200
    # verify undone in grid
    res_grid_after_undo = client.get("/api/grid/hw?page=1&page_size=10")
    rec_after_undo = next(r for r in res_grid_after_undo.json()["records"] if r["uid"] == target_uid)
    assert rec_after_undo["cells"][target_col] != new_val
    assert target_col not in rec_after_undo["user_edited_cols"]
    print("    Undo verified: value rolled back to original and removed from user_edited_cols.")

    res_redo = client.post("/api/redo")
    assert res_redo.status_code == 200
    res_grid_after_redo = client.get("/api/grid/hw?page=1&page_size=10")
    rec_after_redo = next(r for r in res_grid_after_redo.json()["records"] if r["uid"] == target_uid)
    assert rec_after_redo["cells"][target_col] == new_val
    print("    Redo verified: value re-applied.")

    # 6. Test Manual Exclusion Toggle (/api/row/exclusion)
    print("[7] Testing Manual Exclusion toggle...")
    res_excl = client.post("/api/row/exclusion", json={"uid": target_uid, "exclude": True, "user": "Tester"})
    assert res_excl.status_code == 200
    assert res_excl.json()["record"]["is_excluded"] is True
    assert res_excl.json()["record"]["manual_exclude"] is True
    print("    Manual exclusion verified: row moved to excluded (Tab 4).")

    # Undo exclusion
    client.post("/api/undo")
    res_grid_revert_excl = client.get("/api/grid/hw?page=1&page_size=10")
    rec_reverted = next(r for r in res_grid_revert_excl.json()["records"] if r["uid"] == target_uid)
    assert rec_reverted["manual_exclude"] is None
    print("    Undone manual exclusion.")

    # 7. Compare all and run Auto-Fix on all
    print("[8] Comparing SW and UPS, and running Auto-Fix...")
    client.post("/api/compare/sw")
    client.post("/api/compare/ups")

    res_fix_all = client.post("/api/auto-fix/all")
    assert res_fix_all.status_code == 200
    fix_data = res_fix_all.json()
    ov = fix_data["summary"]["overall"]
    assert ov["matched_records"] == 775, f"Expected 775 matched, got {ov['matched_records']}"
    assert ov["excluded_records"] == 741, f"Expected 741 excluded, got {ov['excluded_records']}"
    assert ov["match_rate"] == 100.0
    print(f"    Auto-Fix complete: 100.0% match rate across active records (775 matched, 741 excluded).")

    # Verify custom edited cell was preserved!
    res_grid_hw_final = client.get("/api/grid/hw?page=1&page_size=10")
    rec_final = next(r for r in res_grid_hw_final.json()["records"] if r["uid"] == target_uid)
    assert rec_final["cells"][target_col] == new_val, f"User-edited cell {target_col} was overwritten!"
    print(f"    User-edited precedence verified: {target_col} remained '{new_val}' after Auto-Fix.")

    # 8. Test 4-Tab Excel Exports
    print("[9] Testing 4-Tab Excel Exports...")
    for cat in ("hw", "sw", "ups"):
        res_exp = client.get(f"/api/export/inventory/{cat}")
        assert res_exp.status_code == 200
        wb = openpyxl.load_workbook(io.BytesIO(res_exp.content))
        assert wb.sheetnames == ["1_ข้อมูลเดิม", "2_เปรียบเทียบชี้จุดต่าง", "3_ข้อมูลแก้ไขแล้ว_100%", "4_ข้อมูลที่ถูกตัดออก"]
        # In HW Tab 3, verify custom edited cell is present
        if cat == "hw":
            ws3 = wb["3_ข้อมูลแก้ไขแล้ว_100%"]
            # find column index of PC_Name
            col_idx = None
            for c in range(1, ws3.max_column + 1):
                if ws3.cell(1, c).value == "PC_Name":
                    col_idx = c
                    break
            assert col_idx is not None
            assert ws3.cell(2, col_idx).value == new_val, f"Expected custom val in Tab 3 Excel cell, got {ws3.cell(2, col_idx).value}"
            print("    Tab 3 Excel export verified: contains user-edited cell value.")

    # 8.1 Test Dedicated Mismatch Report Exports
    print("[9.1] Testing Dedicated Mismatch Report Exports...")
    for cat in ("hw", "sw", "ups"):
        res_mism = client.get(f"/api/export/mismatch/{cat}")
        assert res_mism.status_code == 200
        wb_mism = openpyxl.load_workbook(io.BytesIO(res_mism.content))
        assert "สรุป" in wb_mism.sheetnames
        print(f"    Mismatch report for {cat.upper()} verified: sheets={wb_mism.sheetnames}")

    res_mism_all = client.get("/api/export/mismatch-all")
    assert res_mism_all.status_code == 200
    wb_mism_all = openpyxl.load_workbook(io.BytesIO(res_mism_all.content))
    assert "สรุป" in wb_mism_all.sheetnames
    print("    All Mismatch Report verified successfully.")

    # 9. Test Stale State on Master Reload
    print("[10] Testing Stale state when Master ERP is reloaded...")
    client.post("/api/upload", data={"file_type": "master"}, files={"file": ("0-FTI_asset_HWSWOth.xlsx", open(base_master, "rb"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    res_st_after_master = client.get("/api/state")
    st_after = res_st_after_master.json()
    assert st_after["category_states"]["hw"] == "stale"
    assert st_after["category_states"]["sw"] == "stale"
    assert st_after["category_states"]["ups"] == "stale"
    print("    Verified: all categories transitioned to 'stale' when Master ERP reloaded.")

    # 10. Test Compare All with reset_to_orig
    print("[11] Testing Compare All with reset_to_orig=True...")
    res_cmp_all = client.post("/api/compare-all?reset_to_orig=true")
    assert res_cmp_all.status_code == 200
    ov_baseline = res_cmp_all.json()["summary"]["overall"]
    assert ov_baseline["match_rate"] == 51.4, f"Expected 51.4%, got {ov_baseline['match_rate']}"
    assert ov_baseline["mismatched_records"] == 377, f"Expected 377 mismatches, got {ov_baseline['mismatched_records']}"
    print("    Verified: Compare All correctly establishes baseline (51.4% match rate, 377 mismatches).")

    # Clean up test session state so production workspace starts in pristine 'loaded' state
    client.post("/api/session/reset")
    from app import init_app_session
    init_app_session()

    print("\n[SUCCESS] All Manual Compare, Spreadsheet Inspector, and 4-Tab Export tests PASSED successfully!")

def test_debounced_autosave():
    import time
    test_path = os.path.join(BASE_DIR, "output", "test_autosave.json")
    if os.path.exists(test_path):
        try:
            os.remove(test_path)
        except Exception:
            pass

    comparator.autosave(path=test_path, delay=0.2)
    # Debounce delay is 0.2s; immediately checking shouldn't have file written yet
    assert not os.path.exists(test_path)

    time.sleep(0.3)
    assert os.path.exists(test_path)

    comparator.flush_autosave()
    assert os.path.exists(test_path)

    if os.path.exists(test_path):
        try:
            os.remove(test_path)
        except Exception:
            pass
    print("\n[SUCCESS] Debounced autosave verified successfully!")

if __name__ == "__main__":
    test_manual_compare_and_spreadsheet_workflow()
    test_debounced_autosave()
