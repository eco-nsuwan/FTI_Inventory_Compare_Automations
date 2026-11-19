import os
import io
import copy
import math
import zipfile
import tempfile
from typing import List, Dict, Any, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

_EXCEL_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")


def sanitize_excel_value(val: Any) -> Any:
    """
    Neutralize Excel formula injection payloads on string values by prefixing with `'`.
    Preserves None, booleans, ints, floats, and plain finite numeric strings like '-123.45'.
    """
    if val is None or isinstance(val, (bool, int, float)):
        return val
    if isinstance(val, str):
        if not val:
            return val
        if val[0] in _EXCEL_FORMULA_PREFIXES:
            if val[0] in ("-", "+"):
                try:
                    num = float(val)
                    if math.isfinite(num):
                        return val
                except ValueError:
                    pass
            return "'" + val
    return val

# Color codes matching HW_Mismatch_Report.xlsx
FILL_BOTH_MISMATCH = PatternFill(start_color="FFFFCCCC", end_color="FFFFCCCC", fill_type="solid")  # Soft red
FILL_EMP_MISMATCH = PatternFill(start_color="FFFFFFCC", end_color="FFFFFFCC", fill_type="solid")   # Soft yellow
FILL_DEPT_MISMATCH = PatternFill(start_color="FFD6EAF8", end_color="FFD6EAF8", fill_type="solid") # Soft blue
FILL_SCRAPPED = PatternFill(start_color="FFFDEBD0", end_color="FFFDEBD0", fill_type="solid")      # Soft orange
FILL_MISSING_ERP = PatternFill(start_color="FFEAECEE", end_color="FFEAECEE", fill_type="solid")   # Soft gray
FILL_MATCH_GREEN = PatternFill(start_color="FFD1FAE5", end_color="FFD1FAE5", fill_type="solid")   # Soft emerald green
FILL_USER_EDITED = PatternFill(start_color="FFE0F2FE", end_color="FFE0F2FE", fill_type="solid")   # Soft sky blue
FILL_HEADER = PatternFill(start_color="FF1E293B", end_color="FF1E293B", fill_type="solid")        # Dark slate
FILL_HEADER_LIGHT = PatternFill(start_color="FF334155", end_color="FF334155", fill_type="solid")  # Slate
FILL_HEADER_GREEN = PatternFill(start_color="FF15803D", end_color="FF15803D", fill_type="solid")  # Forest green
FILL_HEADER_RED = PatternFill(start_color="FFB91C1C", end_color="FFB91C1C", fill_type="solid")    # Deep red

FONT_HEADER = Font(name="TH Sarabun New", size=13, bold=True, color="FFFFFFFF")
FONT_REGULAR = Font(name="TH Sarabun New", size=11, bold=False, color="FF000000")
FONT_BOLD = Font(name="TH Sarabun New", size=11, bold=True, color="FF000000")
FONT_TITLE = Font(name="TH Sarabun New", size=16, bold=True, color="FF0F172A")
FONT_SUBTITLE = Font(name="TH Sarabun New", size=13, bold=True, color="FF1E293B")
FONT_RED = Font(name="TH Sarabun New", size=11, bold=True, color="FFFF0000")
FONT_ORANGE = Font(name="TH Sarabun New", size=11, bold=True, color="FFD35400")
FONT_GREEN = Font(name="TH Sarabun New", size=11, bold=True, color="FF15803D")

THIN_BORDER = Border(
    left=Side(style='thin', color='FFCBD5E1'),
    right=Side(style='thin', color='FFCBD5E1'),
    top=Side(style='thin', color='FFCBD5E1'),
    bottom=Side(style='thin', color='FFCBD5E1')
)

def get_row_fill_and_font(mismatch_type: str) -> tuple:
    """Return background fill and font based on mismatch type."""
    if not mismatch_type or mismatch_type == "ตรงกันสมบูรณ์":
        return FILL_MATCH_GREEN, FONT_GREEN
    elif "รหัสพนักงาน + แผนก" in mismatch_type:
        return FILL_BOTH_MISMATCH, FONT_RED
    elif "รหัสพนักงาน" in mismatch_type:
        return FILL_EMP_MISMATCH, FONT_RED
    elif "แผนก" in mismatch_type:
        return FILL_DEPT_MISMATCH, FONT_ORANGE
    elif "ตัดจำหน่าย" in mismatch_type:
        return FILL_SCRAPPED, FONT_BOLD
    elif "ไม่พบ" in mismatch_type:
        return FILL_MISSING_ERP, FONT_REGULAR
    return None, FONT_REGULAR


class ReportGenerator:
    """Generates 4-Tab category Excel files and ZIP bundles reflecting exact cell state."""

    @staticmethod
    def _write_dataframe_to_sheet(ws, df: pd.DataFrame, header_fill=FILL_HEADER, highlight_col_name: Optional[str] = None):
        """Write a pandas DataFrame into an openpyxl worksheet with clean styling."""
        ws.views.sheetView[0].showGridLines = True
        cols = list(df.columns)

        # Write headers
        for col_idx, col_name in enumerate(cols, start=1):
            c = ws.cell(row=1, column=col_idx, value=sanitize_excel_value(str(col_name)))
            c.font = FONT_HEADER
            c.fill = header_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = THIN_BORDER
        ws.row_dimensions[1].height = 26

        highlight_col_idx = (cols.index(highlight_col_name) + 1) if (highlight_col_name and highlight_col_name in cols) else None

        # Write rows
        for r_pos, (_, row) in enumerate(df.iterrows(), start=2):
            for c_idx, col_name in enumerate(cols, start=1):
                val = row[col_name]
                if pd.isna(val):
                    val = ""
                cell = ws.cell(row=r_pos, column=c_idx, value=sanitize_excel_value(val))
                cell.font = FONT_REGULAR
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="center")

                if highlight_col_idx and c_idx == highlight_col_idx:
                    s_val = str(val)
                    fill, f_color = get_row_fill_and_font(s_val)
                    if fill:
                        cell.fill = fill
                    if f_color:
                        cell.font = f_color
            ws.row_dimensions[r_pos].height = 20

        # Auto-fit column widths
        for col in ws.columns:
            max_len = max((len(str(cell.value or '')) for cell in col[:250]), default=10)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 48)

    @classmethod
    def build_mismatch_workbook(cls, comparator, category: Optional[str] = None) -> openpyxl.Workbook:
        """
        Create and populate openpyxl Workbook with summary and mismatch records.
        Generates dedicated Mismatch Reports (e.g. HW_Mismatch_Report.xlsx or All_Mismatch_Report.xlsx).
        """
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Remove default sheet

        categories = [category] if category and category != "all" else ["hw", "sw", "ups"]
        all_mismatches = []
        for c in categories:
            all_mismatches.extend([r for r in comparator.records.get(c, []) if not r.get("is_match") or r.get("is_excluded")])

        summary_data = comparator.summary

        # -------------------------------------------------------------
        # Sheet 1: สรุป
        # -------------------------------------------------------------
        ws_summary = wb.create_sheet(title="สรุป")
        ws_summary.views.sheetView[0].showGridLines = True

        title_suffix = f" ({category.upper()})" if category and category != "all" else " (รวมทุกประเภท)"
        ws_summary.cell(row=1, column=1, value=f"รายงาน Mismatch IT Asset Inventory{title_suffix}").font = FONT_TITLE
        ws_summary.cell(row=2, column=1, value=f"เวลาประมวลผล: {pd.Timestamp.now().strftime('%d/%m/%Y %H:%M:%S')}").font = FONT_REGULAR

        # Summary Table by Mismatch Type
        headers_mtype = ["ประเภท Mismatch", "จำนวน (รายการ)", "สัดส่วน (%)"]
        start_row = 4
        for col_idx, h in enumerate(headers_mtype, start=1):
            c = ws_summary.cell(row=start_row, column=col_idx, value=h)
            c.font = FONT_HEADER
            c.fill = FILL_HEADER
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = THIN_BORDER

        curr_row = start_row + 1
        m_breakdown = summary_data.get("mismatch_breakdown", {})
        total_mismatches = len(all_mismatches)

        # Count mismatches for this specific subset
        type_counts = {}
        for r in all_mismatches:
            m = r.get("mismatch_type") or "ไม่ระบุ"
            type_counts[m] = type_counts.get(m, 0) + 1

        for m_name, count in sorted(type_counts.items(), key=lambda x: x[1], reverse=True):
            pct = f"{(count / total_mismatches * 100):.1f}%" if total_mismatches > 0 else "0.0%"
            c1 = ws_summary.cell(row=curr_row, column=1, value=sanitize_excel_value(m_name))
            c2 = ws_summary.cell(row=curr_row, column=2, value=sanitize_excel_value(count))
            c3 = ws_summary.cell(row=curr_row, column=3, value=sanitize_excel_value(pct))

            c2.alignment = Alignment(horizontal="right")
            c3.alignment = Alignment(horizontal="right")

            fill, _ = get_row_fill_and_font(m_name)
            for cell_item in (c1, c2, c3):
                cell_item.font = FONT_REGULAR
                cell_item.border = THIN_BORDER
                if fill:
                    cell_item.fill = fill
            curr_row += 1

        # Total Row
        ws_summary.cell(row=curr_row, column=1, value="รวมจุดต่างทั้งหมด").font = FONT_BOLD
        ws_summary.cell(row=curr_row, column=2, value=total_mismatches).font = FONT_BOLD
        ws_summary.cell(row=curr_row, column=3, value="100.0%").font = FONT_BOLD
        for col_idx in (1, 2, 3):
            ws_summary.cell(row=curr_row, column=col_idx).border = THIN_BORDER
        curr_row += 2

        # Legend Table
        ws_summary.cell(row=curr_row, column=1, value="คำอธิบายสี (Color Legend)").font = FONT_SUBTITLE
        curr_row += 1
        legend_items = [
            ("รหัสพนักงาน + แผนกไม่ตรง", FILL_BOTH_MISMATCH, "รหัสพนักงานและแผนกใน IT ไม่ตรงกับ ERP AX"),
            ("รหัสพนักงานไม่ตรง", FILL_EMP_MISMATCH, "รหัสพนักงานไม่ตรงกัน แต่แผนกตรงกัน"),
            ("แผนกไม่ตรง", FILL_DEPT_MISMATCH, "แผนกไม่ตรงกัน หรือใน ERP ไม่มีแผนก/รออัปเดต"),
            ("ตัดจำหน่ายใน ERP แต่ IT ยังมีใช้งาน", FILL_SCRAPPED, "ใน ERP มีสถานะ Scrapped / ตัดจำหน่ายแล้ว"),
            ("ไม่พบในระบบ ERP (มีเฉพาะใน IT)", FILL_MISSING_ERP, "ไม่มีรหัสทรัพย์สินนี้ใน Master ERP")
        ]
        for l_title, l_fill, l_desc in legend_items:
            c1 = ws_summary.cell(row=curr_row, column=1, value=l_title)
            c2 = ws_summary.cell(row=curr_row, column=2, value=l_desc)
            c1.fill = l_fill
            c1.font = FONT_BOLD
            c2.font = FONT_REGULAR
            c1.border = THIN_BORDER
            c2.border = THIN_BORDER
            curr_row += 1

        # -------------------------------------------------------------
        # Helper: Create Record Table Sheet
        # -------------------------------------------------------------
        def write_record_sheet(sheet_title: str, records_list: List[Dict[str, Any]]):
            ws = wb.create_sheet(title=sheet_title)
            ws.views.sheetView[0].showGridLines = True

            headers = [
                "IT Record ID",
                "IT รหัสพนักงาน",
                "IT ชื่อพนักงาน",
                "IT สถานที่เก็บ (แผนก)",
                "รหัสทรัพย์สิน",
                "ประเภท Asset",
                "ประเภท Mismatch",
                "AX ชื่อพนักงาน",
                "AX รหัสพนักงาน",
                "AX แผนก",
                "AX ชื่อทรัพย์สิน",
                "AX สถานะ",
                "แนวทางแก้ไข (IT Action)"
            ]

            for col_idx, h in enumerate(headers, start=1):
                c = ws.cell(row=1, column=col_idx, value=h)
                c.font = FONT_HEADER
                c.fill = FILL_HEADER
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.border = THIN_BORDER
            ws.row_dimensions[1].height = 28

            for r_idx, rec in enumerate(records_list, start=2):
                cat_name = rec.get("category", "")
                aid, it_emp_id, it_emp_name, it_dept, sn = comparator._get_key_fields(cat_name, rec.get("cells", {}))
                ax = rec.get("ax_data", {})
                m_type = rec.get("mismatch_type", "")

                if "รหัสพนักงาน + แผนก" in m_type:
                    action = "ตรวจสอบการโอนย้ายทรัพย์สินและผู้ถือครองจริง"
                elif "รหัสพนักงาน" in m_type:
                    action = "ปรับรหัสพนักงานใน IT ให้ตรงกับ ERP หรือแจ้งย้ายสิทธิ์"
                elif "แผนก" in m_type:
                    action = "ปรับแผนกใน IT ให้ตรงกับ ERP หรือตรวจสอบหน่วยงาน"
                elif "ตัดจำหน่าย" in m_type:
                    action = "ตรวจสอบเพื่อตัดจำหน่ายออกจาก IT Inventory"
                elif "ไม่พบ" in m_type:
                    action = "ตรวจสอบเลขที่ทรัพย์สิน หรือลงทะเบียนเข้าระบบ ERP"
                else:
                    action = "ตรวจสอบข้อมูล"

                row_vals = [
                    rec.get("uid", ""),
                    it_emp_id,
                    it_emp_name,
                    it_dept,
                    rec.get("asset_id", aid),
                    cat_name.upper(),
                    m_type,
                    ax.get("ax_emp_name", ""),
                    ax.get("ax_emp_id", ""),
                    ax.get("ax_dept", ""),
                    ax.get("ax_asset_name", ""),
                    ax.get("ax_status", ""),
                    action
                ]

                fill, font_color = get_row_fill_and_font(m_type)

                for c_idx, val in enumerate(row_vals, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=sanitize_excel_value(val))
                    cell.font = FONT_REGULAR
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(vertical="center")

                    if c_idx == 7 and fill:
                        cell.fill = fill
                        cell.font = font_color

                    if c_idx == 2 and "emp_id" in rec.get("diff_fields", []):
                        cell.font = FONT_RED
                    if c_idx == 4 and "dept" in rec.get("diff_fields", []):
                        cell.font = FONT_ORANGE

                ws.row_dimensions[r_idx].height = 20

            for col in ws.columns:
                max_len = max((len(str(cell.value or '')) for cell in col[:250]), default=10)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 48)

        # Sheet 2: ทั้งหมด
        sheet_all_title = f"ทั้งหมด ({len(all_mismatches)})"[:31]
        write_record_sheet(sheet_all_title, all_mismatches)

        # Sub-sheets for specific mismatch categories
        cat_configs = [
            ("รหัสพนักงาน+แผนก", lambda r: "รหัสพนักงาน + แผนก" in r.get("mismatch_type", "")),
            ("รหัสพนักงาน", lambda r: r.get("mismatch_type") == "รหัสพนักงานไม่ตรง"),
            ("แผนก", lambda r: r.get("mismatch_type") == "แผนกไม่ตรง"),
            ("ตัดจำหน่ายใน ERP", lambda r: "ตัดจำหน่าย" in r.get("mismatch_type", "")),
            ("ไม่พบใน ERP", lambda r: "ไม่พบ" in r.get("mismatch_type", ""))
        ]

        for s_title, filter_fn in cat_configs:
            filtered_subset = [r for r in all_mismatches if filter_fn(r)]
            if filtered_subset:
                clean_title = f"{s_title} ({len(filtered_subset)})"[:31]
                write_record_sheet(clean_title, filtered_subset)

        for col in ws_summary.columns:
            max_len = max((len(str(cell.value or '')) for cell in col[:250]), default=10)
            col_letter = get_column_letter(col[0].column)
            ws_summary.column_dimensions[col_letter].width = min(max(max_len + 5, 14), 48)

        return wb

    @classmethod
    def generate_mismatch_report_bytes(cls, comparator, category: Optional[str] = None) -> io.BytesIO:
        """Generate Excel mismatch report directly into in-memory BytesIO buffer."""
        wb = cls.build_mismatch_workbook(comparator, category)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    @classmethod
    def generate_mismatch_report_file(cls, comparator, output_path: str, category: Optional[str] = None) -> str:
        """Save report to file path with fallback to temp directory if write is restricted."""
        wb = cls.build_mismatch_workbook(comparator, category)
        try:
            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            wb.save(output_path)
            return output_path
        except Exception:
            fallback_dir = os.path.join(tempfile.gettempdir(), "fti_reports")
            os.makedirs(fallback_dir, exist_ok=True)
            fallback_path = os.path.join(fallback_dir, os.path.basename(output_path))
            wb.save(fallback_path)
            return fallback_path

    @classmethod
    def build_category_4tab_workbook(cls, comparator, category: str) -> openpyxl.Workbook:
        """
        Build the 4-Tab Excel workbook for a specific category (hw, sw, ups):
        1. 1_ข้อมูลเดิม (Original untouched data & columns from orig_cells)
        2. 2_เปรียบเทียบชี้จุดต่าง (Side-by-side comparison & diff highlighting & user edit notes)
        3. 3_ข้อมูลแก้ไขแล้ว_100% (Latest cells of active inventory with exact original columns)
        4. 4_ข้อมูลที่ถูกตัดออก (Excluded records with exact original columns + exclusion reason & ERP info)
        """
        from engine.comparator import COLUMN_MAP

        if category not in ("hw", "sw", "ups"):
            raise ValueError(f"Unknown category: {category}")

        recs = comparator.records.get(category, [])
        cols = comparator.columns.get(category, [])
        if not cols and recs:
            cols = list(recs[0]["cells"].keys())

        cmap = COLUMN_MAP.get(category, {})

        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        # =====================================================================
        # TAB 1: 1_ข้อมูลเดิม (Original Untouched Data)
        # =====================================================================
        ws_orig = wb.create_sheet(title="1_ข้อมูลเดิม")
        orig_rows = []
        for r in recs:
            orig_rows.append(r.get("orig_cells", r.get("cells", {})))
        df_orig = pd.DataFrame(orig_rows, columns=cols)
        cls._write_dataframe_to_sheet(ws_orig, df_orig, header_fill=FILL_HEADER)

        # =====================================================================
        # TAB 2: 2_เปรียบเทียบชี้จุดต่าง (Detailed Diff & Resolution Sheet)
        # =====================================================================
        ws_comp = wb.create_sheet(title="2_เปรียบเทียบชี้จุดต่าง")
        ws_comp.views.sheetView[0].showGridLines = True

        comp_headers = [
            "ลำดับ",
            "IT Record ID",
            "เลขที่ทรัพย์สิน",
            "ประเภท Asset",
            "ชื่อทรัพย์สิน / Software",
            "IT รหัสพนักงาน (เดิม)",
            "IT ชื่อผู้ใช้งาน (เดิม)",
            "IT สถานที่เก็บ/แผนก (เดิม)",
            "จุดต่างก่อนแก้ไข (Mismatch)",
            "ERP รหัสพนักงาน (AX)",
            "ERP ชื่อพนักงาน (AX)",
            "ERP แผนก (AX)",
            "ERP สถานะ (AX)",
            "การดำเนินการของระบบ",
            "ผลลัพธ์หลังตรวจซ้ำ",
            "คอลัมน์ที่ผู้ใช้แก้ไขเอง"
        ]

        for col_idx, h in enumerate(comp_headers, start=1):
            c = ws_comp.cell(row=1, column=col_idx, value=h)
            c.font = FONT_HEADER
            c.fill = FILL_HEADER_LIGHT
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = THIN_BORDER
        ws_comp.row_dimensions[1].height = 28

        active_rows = []
        excluded_rows = []

        for row_num, r in enumerate(recs, start=2):
            orig_cells = r.get("orig_cells", {})
            curr_cells = r.get("cells", {})
            ax = r.get("ax_data", {})
            m_type = r.get("mismatch_type", "")
            diffs = r.get("diff_fields", [])
            user_edited = r.get("user_edited_cols", [])

            orig_aid = orig_cells.get(cmap.get("asset_id", ""), "")
            orig_emp_id = orig_cells.get(cmap.get("emp_id", ""), "")
            orig_emp_name = orig_cells.get(cmap.get("emp_name", ""), "")
            orig_dept = orig_cells.get(cmap.get("dept", ""), "")

            action_taken = "คงค่าเดิม"
            if r.get("is_excluded"):
                action_taken = f"คัดออกไปยังแท็บ 4 ({r.get('exclude_reason', '')})"
                final_status = "ถูกตัดออกหลังเปรียบเทียบ"
            elif user_edited:
                action_taken = f"แก้ไขโดยผู้ใช้ ({', '.join(user_edited)})"
                final_status = "ตรงกันสมบูรณ์ 100%" if r.get("is_match") else "ผู้ใช้ปรับแต่งเอง (ต่างจาก ERP)"
            elif r.get("is_match"):
                action_taken = "ซิงค์ตรงกับ ERP 100%" if m_type != "ตรงกันสมบูรณ์" else "ข้อมูลตรงกับ ERP อยู่แล้ว"
                final_status = "ตรงกันสมบูรณ์ 100% (ผ่านตรวจซ้ำ 3 รอบ)"
            else:
                action_taken = "รอการแก้ไข"
                final_status = m_type

            row_vals = [
                row_num - 1,
                curr_cells.get("Record-ID") or curr_cells.get("IdAssetRecId") or "",
                orig_aid,
                curr_cells.get("ประเภททรัพย์สิน") or ("Software" if category == "sw" else "PC/Monitor"),
                curr_cells.get("ชื่อทรัพย์สิน") or curr_cells.get("ชื่อ Software") or "",
                orig_emp_id,
                orig_emp_name,
                orig_dept,
                m_type,
                ax.get("ax_emp_id", ""),
                ax.get("ax_emp_name", ""),
                ax.get("ax_dept", ""),
                ax.get("ax_status", ""),
                action_taken,
                final_status,
                ", ".join(user_edited) if user_edited else "-"
            ]

            fill_m, font_m = get_row_fill_and_font(m_type)

            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws_comp.cell(row=row_num, column=c_idx, value=sanitize_excel_value(val))
                cell.font = FONT_REGULAR
                cell.border = THIN_BORDER
                cell.alignment = Alignment(vertical="center")

                # Highlight original mismatch badge column
                if c_idx == 9:
                    if m_type == "ตรงกันสมบูรณ์":
                        cell.fill = FILL_MATCH_GREEN
                        cell.font = FONT_GREEN
                    elif fill_m:
                        cell.fill = fill_m
                        cell.font = font_m

                # Highlight diff cells on IT side
                if c_idx in (6, 7) and "emp_id" in diffs:
                    cell.fill = FILL_EMP_MISMATCH
                    cell.font = FONT_RED
                if c_idx == 8 and "dept" in diffs:
                    cell.fill = FILL_DEPT_MISMATCH
                    cell.font = FONT_ORANGE

                # Highlight final status column
                if c_idx == 15:
                    if r.get("is_excluded"):
                        cell.fill = FILL_SCRAPPED if "ตัดจำหน่าย" in m_type else FILL_MISSING_ERP
                        cell.font = FONT_ORANGE if "ตัดจำหน่าย" in m_type else FONT_BOLD
                    elif r.get("is_match"):
                        cell.fill = FILL_MATCH_GREEN
                        cell.font = FONT_GREEN
                    elif user_edited:
                        cell.fill = FILL_USER_EDITED
                        cell.font = FONT_BOLD

                # User edited column
                if c_idx == 16 and user_edited:
                    cell.fill = FILL_USER_EDITED
                    cell.font = FONT_BOLD

            ws_comp.row_dimensions[row_num].height = 20

            # Partition for Tab 3 & Tab 4
            if r.get("is_excluded"):
                ex_row = copy.deepcopy(curr_cells)
                ex_row["สาเหตุที่ถูกตัดออก"] = r.get("exclude_reason") or m_type
                ex_row["AX สถานะใน ERP"] = ax.get("ax_status", "ไม่พบใน ERP" if not ax else "ตัดจำหน่าย")
                ex_row["AX ชื่อทรัพย์สินใน ERP"] = ax.get("ax_asset_name", "")
                ex_row["AX รหัสพนักงานใน ERP"] = ax.get("ax_emp_id", "")
                ex_row["AX ชื่อพนักงานใน ERP"] = ax.get("ax_emp_name", "")
                ex_row["AX แผนกใน ERP"] = ax.get("ax_dept", "")
                excluded_rows.append(ex_row)
            else:
                active_rows.append(copy.deepcopy(curr_cells))

        for col in ws_comp.columns:
            max_len = max((len(str(cell.value or '')) for cell in col[:250]), default=10)
            col_letter = get_column_letter(col[0].column)
            ws_comp.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)

        # =====================================================================
        # TAB 3: 3_ข้อมูลแก้ไขแล้ว_100% (Active Reconciled Inventory)
        # =====================================================================
        ws_tab3 = wb.create_sheet(title="3_ข้อมูลแก้ไขแล้ว_100%")
        df_tab3 = pd.DataFrame(active_rows, columns=cols)
        cls._write_dataframe_to_sheet(ws_tab3, df_tab3, header_fill=FILL_HEADER_GREEN)

        # =====================================================================
        # TAB 4: 4_ข้อมูลที่ถูกตัดออก (Excluded Inventory)
        # =====================================================================
        extra_cols = [
            "สาเหตุที่ถูกตัดออก",
            "AX สถานะใน ERP",
            "AX ชื่อทรัพย์สินใน ERP",
            "AX รหัสพนักงานใน ERP",
            "AX ชื่อพนักงานใน ERP",
            "AX แผนกใน ERP"
        ]
        ws_tab4 = wb.create_sheet(title="4_ข้อมูลที่ถูกตัดออก")
        df_tab4 = pd.DataFrame(excluded_rows, columns=cols + extra_cols)
        cls._write_dataframe_to_sheet(ws_tab4, df_tab4, header_fill=FILL_HEADER_RED, highlight_col_name="สาเหตุที่ถูกตัดออก")

        return wb

    @classmethod
    def export_corrected_inventory_bytes(cls, comparator, category: str) -> io.BytesIO:
        """Export 4-Tab inventory workbook into BytesIO."""
        wb = cls.build_category_4tab_workbook(comparator, category)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    @classmethod
    def save_all_category_files(cls, comparator, output_dir: str) -> Dict[str, str]:
        """Save all compared category 4-tab Excel files into output_dir."""
        os.makedirs(output_dir, exist_ok=True)
        file_map = {
            "hw": "1-it_asset_HW.xlsx",
            "sw": "2-it_asset_SW.xlsx",
            "ups": "3-it_asset_UPS_Monitor.xlsx"
        }
        saved_paths = {}
        for cat, fname in file_map.items():
            if comparator.category_states.get(cat) == "compared" and len(comparator.records.get(cat, [])) > 0:
                wb = cls.build_category_4tab_workbook(comparator, cat)
                target_path = os.path.join(output_dir, fname)
                wb.save(target_path)
                saved_paths[cat] = target_path
        return saved_paths

    @classmethod
    def export_all_inventories_zip_bytes(cls, comparator) -> io.BytesIO:
        """Export all compared 4-tab Excel files into a single ZIP archive."""
        file_map = {
            "hw": "1-it_asset_HW.xlsx",
            "sw": "2-it_asset_SW.xlsx",
            "ups": "3-it_asset_UPS_Monitor.xlsx"
        }
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for cat, fname in file_map.items():
                if comparator.category_states.get(cat) == "compared" and len(comparator.records.get(cat, [])) > 0:
                    excel_buf = cls.export_corrected_inventory_bytes(comparator, cat)
                    zf.writestr(fname, excel_buf.getvalue())
        zip_buf.seek(0)
        return zip_buf
