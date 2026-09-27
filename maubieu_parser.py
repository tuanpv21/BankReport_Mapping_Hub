# -*- coding: utf-8 -*-
"""
Maubieu & Matrix Report Parser for BankReport Mapping Hub.
Parses Bank / SBV reporting template Excel files (such as A02211, A02224, G01480, G04224, G04235)
extracting report metadata, horizontal columns, vertical rows, and cross-cell calculations.
"""

import os
import re
import openpyxl
import sqlite3

GL_DIR = r"E:\1080 Public Bank\TT35\CIC\GL"
MAUBIEU_CONFIG_DIR = os.path.join(GL_DIR, "Maubieu_Config")
DB_PATH = os.path.join(os.path.dirname(__file__), "mapping_hub.db")

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def read_procedure_file(proc_name):
    """Attempt to read actual .prc or .sql file from GL folder."""
    if not os.path.exists(GL_DIR):
        return ""
    candidates = [
        f"{proc_name}.prc",
        f"pr_{proc_name}.prc",
        f"pr_{proc_name.replace('RPTB_', '')}.prc",
        f"{proc_name.replace('pr_', '')}.prc"
    ]
    for c in candidates:
        fp = os.path.join(GL_DIR, c)
        if os.path.exists(fp):
            try:
                with open(fp, "r", encoding="utf-8", errors="ignore") as f:
                    return f.read()
            except Exception:
                pass
    return ""

def parse_maubieu_excel(file_path):
    """Parses a single Maubieu Excel file and returns structured report and cell mappings."""
    if not os.path.exists(file_path):
        return None

    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    filename = os.path.basename(file_path)

    # Base report code
    base_code = filename.split(".")[0].split("-")[0].upper()

    # Determine Title, Unit, Cycle
    rpt_name = ""
    unit = ""
    cycle = ""

    for r in range(8, 16):
        for c in range(2, 5):
            val = str(ws.cell(row=r, column=c).value or "").strip()
            if any(k in val.upper() for k in ["BÁO CÁO", "BẢNG CÂN ĐỐI", "DANH SÁCH"]):
                rpt_name = val
                # Look for cycle and unit in following rows
                for r2 in range(r + 1, r + 4):
                    v2 = str(ws.cell(row=r2, column=c).value or "").strip()
                    if v2.startswith("(") and v2.endswith(")"):
                        cycle = v2
                    elif "Đơn vị tính" in v2:
                        unit = v2
                break
        if rpt_name:
            break

    if not rpt_name:
        if base_code == "A02211":
            rpt_name = "BẢNG CÂN ĐỐI TÀI KHOẢN KẾ TOÁN (Ngày)"
            cycle = "(Ngày)"
            unit = "Đồng Việt Nam (VND)"
        elif base_code == "A02224":
            rpt_name = "BÁO CÁO MỘT SỐ TÀI KHOẢN KẾ TOÁN VÀ THUYẾT MINH CHI TIẾT TÀI KHOẢN KẾ TOÁN (Tháng)"
            cycle = "(Tháng)"
            unit = "Đồng Việt Nam (VND)"
        elif base_code == "G01480":
            rpt_name = "BÁO CÁO GIÁ TRỊ THỰC CỦA VỐN ĐIỀU LỆ, VỐN ĐƯỢC CẤP (2 kỳ/năm)"
            cycle = "(2 kỳ/năm)"
            unit = "Triệu VND"
        elif base_code == "G04224":
            rpt_name = "BÁO CÁO VỐN TỰ CÓ CỦA NGÂN HÀNG (Tháng)"
            cycle = "(Tháng)"
            unit = "Triệu VND"
        elif base_code == "G04235":
            rpt_name = "BÁO CÁO VỐN TỰ CÓ HỢP NHẤT CỦA NGÂN HÀNG (Quý)"
            cycle = "(Quý)"
            unit = "Triệu VND"
        else:
            rpt_name = f"Mẫu biểu Báo cáo {base_code}"

    target_table = f"RPTB_{base_code}"
    proc_name = f"pr_{base_code}"
    proc_code = read_procedure_file(proc_name)

    # 1. Detect Columns (Row with C-100 or similar)
    col_row = None
    for r in range(12, 19):
        vals = [str(ws.cell(row=r, column=c).value or "") for c in range(1, ws.max_column + 1)]
        if any("C-" in v for v in vals):
            col_row = r
            break

    columns = []
    if col_row:
        for c in range(1, ws.max_column + 1):
            c_code = str(ws.cell(row=col_row, column=c).value or "").strip()
            if "C-" in c_code:
                # Header labels
                l_main = str(ws.cell(row=col_row - 1, column=c).value or "").strip()
                l_sub1 = str(ws.cell(row=col_row + 1, column=c).value or "").strip()
                l_sub2 = str(ws.cell(row=col_row + 2, column=c).value or "").strip()
                full_name = " - ".join([x for x in [l_main, l_sub1, l_sub2] if x and not x.isdigit()])
                if not full_name:
                    full_name = l_sub1 or l_main or c_code

                columns.append({
                    "col_idx": c,
                    "col_code": c_code,
                    "col_name": full_name
                })

    # 2. Detect Rows (Rows starting with R-100 or similar)
    data_start = (col_row + 2) if col_row else 18
    rows = []

    for r in range(data_start, ws.max_row + 1):
        r_id = str(ws.cell(row=r, column=1).value or "").strip()
        if not r_id.startswith("R-"):
            continue

        # Specific row extraction per report template structure
        stt = ""
        name = ""
        acct = ""
        guide = ""

        if base_code == "A02211":
            name = str(ws.cell(row=r, column=2).value or "").strip()
            acct = str(ws.cell(row=r, column=3).value or "").strip()
        elif base_code == "A02224":
            stt = str(ws.cell(row=r, column=2).value or "").strip()
            name = str(ws.cell(row=r, column=3).value or "").strip()
            acct = str(ws.cell(row=r, column=4).value or "").strip()
        elif base_code in ["G04224", "G04235"]:
            stt = str(ws.cell(row=r, column=2).value or "").strip()
            name = str(ws.cell(row=r, column=3).value or "").strip()
            guide = str(ws.cell(row=r, column=5).value or "").strip()
            m = re.search(r"Account\s+([0-9]+)", guide, re.IGNORECASE)
            acct = m.group(1) if m else ""
        elif base_code == "G01480":
            stt = str(ws.cell(row=r, column=2).value or "").strip()
            name = str(ws.cell(row=r, column=3).value or "").strip()
            guide = str(ws.cell(row=r, column=4).value or "").strip()
            m = re.search(r"TK\s*([0-9]+)", guide, re.IGNORECASE)
            acct = m.group(1) if m else ""
        else:
            name = str(ws.cell(row=r, column=2).value or "").strip()
            acct = str(ws.cell(row=r, column=3).value or "").strip()

        rows.append({
            "row_idx": r,
            "row_id": r_id,
            "stt": stt,
            "name": name or f"Chỉ tiêu {r_id}",
            "acct": acct,
            "guide": guide
        })

    # 3. Generate Cell Matrix
    cells = []
    # Column semantic mappings for standard reports:
    # A02211:
    # C-104: DNDK, C-105: DCDK, C-107: PSN, C-108: PSC, C-110: DNCK, C-111: DCCK
    a02211_col_map = {
        "C-104": ("DNDK", "Số dư đầu kỳ - Nợ", "NO_DAU_KY", "DNDK"),
        "C-105": ("DCDK", "Số dư đầu kỳ - Có", "CO_DAU_KY", "DCDK"),
        "C-107": ("PSN", "Số phát sinh - Nợ", "NO_PHAT_SINH", "PSN"),
        "C-108": ("PSC", "Số phát sinh - Có", "CO_PHAT_SINH", "PSC"),
        "C-110": ("DNCK", "Số dư cuối kỳ - Nợ", "NO_CUOI_KY", "DNCK"),
        "C-111": ("DCCK", "Số dư cuối kỳ - Có", "CO_CUOI_KY", "DCCK")
    }

    # A02224:
    # C-102: DCN_LCY, C-103: DCC_LCY, C-105: DCN_FCY, C-106: DCC_FCY, C-108: DCN_SUM, C-109: DCC_SUM
    a02224_col_map = {
        "C-102": ("DCN_LCY", "VND - Dư nợ", "NO_CUOI_KY", "DNCK", "VND"),
        "C-103": ("DCC_LCY", "VND - Dư có", "CO_CUOI_KY", "DCCK", "VND"),
        "C-105": ("DCN_FCY", "Ngoại tệ - Dư nợ", "NO_CUOI_KY", "DNCK", "FCY"),
        "C-106": ("DCC_FCY", "Ngoại tệ - Dư có", "CO_CUOI_KY", "DCCK", "FCY"),
        "C-108": ("DCN_SUM", "Tổng số - Dư nợ", "(3)+(5)", "FORMULA", "ALL"),
        "C-109": ("DCC_SUM", "Tổng số - Dư có", "(4)+(6)", "FORMULA", "ALL")
    }

    for rw in rows:
        r_id = rw["row_id"]
        r_name = rw["name"]
        r_acct = rw["acct"]
        r_guide = rw["guide"]

        if base_code == "A02211":
            for c_code, (tgt_col, col_title, src_col, bal_type) in a02211_col_map.items():
                cells.append({
                    "system_code": "TT35",
                    "report_code": base_code,
                    "field_code": f"{r_id}_{tgt_col}",
                    "field_name_vi": f"{r_name} ({col_title})",
                    "data_type": "NUMBER(18,2)",
                    "target_table": target_table,
                    "target_column": tgt_col,
                    "row_id": r_id,
                    "col_id": c_code,
                    "gl_account": r_acct,
                    "calc_method": "GL_CONFIG",
                    "balance_type": bal_type,
                    "source_system": "ODS",
                    "source_table": "t1080_tb_gl_bal_quy_doi",
                    "source_column": src_col,
                    "transformation_rule": f"GL_CODE LIKE '{r_acct}%' -> SUM({src_col})",
                    "implemented_by": proc_name,
                    "notes": f"Tài khoản {r_acct} ({r_name})",
                    "status": "Active"
                })

        elif base_code == "A02224":
            for c_code, (tgt_col, col_title, src_col, bal_type, *curr) in a02224_col_map.items():
                curr_type = curr[0] if curr else "ALL"
                calc_m = "FORMULA" if bal_type == "FORMULA" else "GL_CONFIG"
                rule = f"Biểu thức: {src_col}" if calc_m == "FORMULA" else f"GL_CODE LIKE '{r_acct}%' [{curr_type}] -> SUM({src_col})"

                cells.append({
                    "system_code": "TT35",
                    "report_code": base_code,
                    "field_code": f"{r_id}_{tgt_col}",
                    "field_name_vi": f"{r_name} ({col_title})",
                    "data_type": "NUMBER(18,2)",
                    "target_table": target_table,
                    "target_column": tgt_col,
                    "row_id": r_id,
                    "col_id": c_code,
                    "gl_account": r_acct,
                    "calc_method": calc_m,
                    "balance_type": bal_type,
                    "source_system": "ODS",
                    "source_table": "t1080_tb_gl_bal_quy_doi" if calc_m == "GL_CONFIG" else "",
                    "source_column": src_col if calc_m == "GL_CONFIG" else "",
                    "transformation_rule": rule,
                    "implemented_by": proc_name,
                    "notes": f"Tài khoản {r_acct} [{curr_type}]",
                    "status": "Active"
                })

        elif base_code in ["G04224", "G04235"]:
            # Single value column GIA_TRI
            calc_m = "FORMULA" if (any(k in r_guide.lower() for k in ["=", "-", "+", "sum", "if", "∑"]) and not r_guide.lower().startswith("account")) else "GL_CONFIG"
            extracted_acct = ""
            m_acc = re.search(r"account\s+([0-9]+)", r_guide, re.IGNORECASE)
            if m_acc:
                extracted_acct = m_acc.group(1)

            cells.append({
                "system_code": "TT35",
                "report_code": base_code,
                "field_code": f"{r_id}_GIA_TRI",
                "field_name_vi": r_name,
                "data_type": "NUMBER(18,2)",
                "target_table": target_table,
                "target_column": "GIA_TRI",
                "row_id": r_id,
                "col_id": "C-100",
                "gl_account": extracted_acct or r_acct,
                "calc_method": calc_m,
                "balance_type": "DNCK",
                "formula_expr": r_guide,
                "source_system": "ODS",
                "source_table": "t1080_tb_gl_bal_quy_doi" if calc_m == "GL_CONFIG" else "",
                "source_column": "NO_CUOI_KY - CO_CUOI_KY" if calc_m == "GL_CONFIG" else "",
                "transformation_rule": f"Công thức: {r_guide}" if r_guide else (f"GL_CODE LIKE '{extracted_acct}%'" if extracted_acct else ""),
                "implemented_by": proc_name,
                "notes": f"Hướng dẫn: {r_guide}" if r_guide else "",
                "status": "Active"
            })

        elif base_code == "G01480":
            # G01480 has SO_DU, SO_TIEN_THAY_DOI, TYLE_THAY_DOI
            g_cols = [
                ("C-101", "SO_DU", "Số dư kỳ này", "GL_CONFIG"),
                ("C-103", "SO_TIEN_THAY_DOI", "Tăng/giảm so với kỳ trước", "FORMULA"),
                ("C-104", "TYLE_THAY_DOI", "Tỷ lệ % thay đổi", "FORMULA")
            ]
            for c_code, tgt_col, col_title, calc_m in g_cols:
                cells.append({
                    "system_code": "TT35",
                    "report_code": base_code,
                    "field_code": f"{r_id}_{tgt_col}",
                    "field_name_vi": f"{r_name} - {col_title}",
                    "data_type": "NUMBER(18,2)",
                    "target_table": target_table,
                    "target_column": tgt_col,
                    "row_id": r_id,
                    "col_id": c_code,
                    "gl_account": r_acct,
                    "calc_method": calc_m,
                    "source_system": "ODS",
                    "source_table": "t1080_tb_gl_bal_quy_doi" if calc_m == "GL_CONFIG" else "",
                    "source_column": "SO_DU" if calc_m == "GL_CONFIG" else "",
                    "transformation_rule": f"Hướng dẫn: {r_guide}" if r_guide else "",
                    "implemented_by": proc_name,
                    "notes": r_guide or "",
                    "status": "Active"
                })

    return {
        "system_code": "TT35",
        "report_code": base_code,
        "report_name": rpt_name,
        "cycle": cycle,
        "unit": unit,
        "target_table": target_table,
        "procedure_name": proc_name,
        "procedure_code": proc_code,
        "template_file": filename,
        "columns_count": len(columns),
        "rows_count": len(rows),
        "columns": columns,
        "rows": rows,
        "cells": cells
    }

def import_all_maubieu_templates():
    """Scans and imports all Maubieu_Config files into mapping_hub.db."""
    if not os.path.exists(MAUBIEU_CONFIG_DIR):
        return {"success": False, "error": f"Folder {MAUBIEU_CONFIG_DIR} not found"}

    target_files = ["A02211.xlsx", "A02224.xlsx", "G01480.xlsx", "G04224.xlsx", "G04235.xlsx"]
    conn = get_db()
    cur = conn.cursor()

    total_imported_reports = 0
    total_imported_cells = 0

    for fn in target_files:
        fp = os.path.join(MAUBIEU_CONFIG_DIR, fn)
        if not os.path.exists(fp):
            continue

        data = parse_maubieu_excel(fp)
        if not data:
            continue

        # 1. Upsert Report
        cur.execute("""
            INSERT INTO reports (
                system_code, report_code, report_name, target_table,
                procedure_name, procedure_code, template_file, cycle, unit, has_matrix, frequency
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
            ON CONFLICT(report_code) DO UPDATE SET
                system_code = excluded.system_code,
                report_name = excluded.report_name,
                target_table = excluded.target_table,
                procedure_name = excluded.procedure_name,
                procedure_code = coalesce(nullif(excluded.procedure_code, ''), reports.procedure_code),
                template_file = excluded.template_file,
                cycle = excluded.cycle,
                unit = excluded.unit,
                has_matrix = 1
        """, (
            data["system_code"],
            data["report_code"],
            data["report_name"],
            data["target_table"],
            data["procedure_name"],
            data["procedure_code"],
            data["template_file"],
            data["cycle"],
            data["unit"],
            data["cycle"] or "Định kỳ"
        ))

        # 2. Insert Cells / Field Mappings
        # Clear old generated cells for this report to avoid duplicates
        cur.execute("DELETE FROM mappings WHERE system_code = ? AND report_code = ?", (data["system_code"], data["report_code"]))

        for c in data["cells"]:
            cur.execute("""
                INSERT INTO mappings (
                    system_code, report_code, field_code, field_name_vi, data_type,
                    target_table, target_column, source_system, source_table, source_column,
                    transformation_rule, implemented_by, notes, status,
                    calc_method, row_id, col_id, gl_account, balance_type, formula_expr
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                c["system_code"],
                c["report_code"],
                c["field_code"],
                c["field_name_vi"],
                c["data_type"],
                c["target_table"],
                c["target_column"],
                c.get("source_system", "ODS"),
                c.get("source_table", ""),
                c.get("source_column", ""),
                c.get("transformation_rule", ""),
                c.get("implemented_by", ""),
                c.get("notes", ""),
                c.get("status", "Active"),
                c.get("calc_method", "GL_CONFIG"),
                c.get("row_id", ""),
                c.get("col_id", ""),
                c.get("gl_account", ""),
                c.get("balance_type", ""),
                c.get("formula_expr", "")
            ))
            total_imported_cells += 1

        total_imported_reports += 1

    conn.commit()
    conn.close()

    return {
        "success": True,
        "reports_imported": total_imported_reports,
        "cells_imported": total_imported_cells
    }

if __name__ == "__main__":
    res = import_all_maubieu_templates()
    print("Import result:", res)
