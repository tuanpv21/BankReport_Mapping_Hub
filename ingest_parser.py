# -*- coding: utf-8 -*-
"""
Smart Ingestion & Parser Engine for BankReport Mapping Hub.
Scans and parses:
  - Excel Specs: Table_Mo_ta (2).xlsx (CIC), TT35 new/*.xlsx
  - PL/SQL Packages & Procedures: pk_cic_calc.pck, pkg_cic.pck, pr_*.prc
  - Table Definitions (*.tab)
"""

import os
import re
import glob
import sqlite3
import openpyxl

PROJECT_ROOT = r"E:\1080 Public Bank"
CIC_DIR = os.path.join(PROJECT_ROOT, r"project_tuanpv\CIC_TUANPV")
ODS_DIR = os.path.join(PROJECT_ROOT, r"project_tuanpv\ODS_TUANPV")
TT35_DIR = os.path.join(PROJECT_ROOT, r"TT35")
DB_PATH = os.path.join(os.path.dirname(__file__), "mapping_hub.db")

def init_db(db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    
    # Tables for Hub
    cur.execute("""
    CREATE TABLE IF NOT EXISTS systems (
        code TEXT PRIMARY KEY,
        name TEXT,
        regulation TEXT,
        description TEXT
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS reports (
        report_code TEXT PRIMARY KEY,
        system_code TEXT,
        report_name TEXT,
        frequency TEXT,
        target_table TEXT,
        procedure_name TEXT,
        regulation_ref TEXT,
        description TEXT,
        FOREIGN KEY (system_code) REFERENCES systems(code)
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS mappings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        system_code TEXT,
        report_code TEXT,
        field_code TEXT,
        field_name_vi TEXT,
        data_type TEXT,
        target_table TEXT,
        target_column TEXT,
        source_system TEXT,
        source_table TEXT,
        source_column TEXT,
        transformation_rule TEXT,
        lookup_ref TEXT,
        implemented_by TEXT,
        regulatory_ref TEXT,
        status TEXT DEFAULT 'Active',
        notes TEXT,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )""")
    
    cur.execute("""
    CREATE TABLE IF NOT EXISTS code_lookups (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        system_code TEXT,
        category_name TEXT,
        bank_code TEXT,
        bank_desc TEXT,
        target_code TEXT,
        target_desc TEXT,
        source_table TEXT,
        source_column TEXT,
        notes TEXT
    )""")
    
    # Seed systems if empty
    cur.execute("SELECT COUNT(*) FROM systems")
    if cur.fetchone()[0] == 0:
        cur.executemany("INSERT INTO systems VALUES (?, ?, ?, ?)", [
            ("CIC", "Hệ thống Thông tin Tín dụng Quốc gia", "Quyết định 573/QĐ-NHNN", "Báo cáo thông tin khách hàng, quan hệ tín dụng, khế ước, thẻ, TSĐB, nợ xấu cho CIC"),
            ("TT35", "Báo cáo Thống kê Ngân hàng Nhà nước", "Thông tư 35/2015/TT-NHNN & TT11/2021", "Báo cáo thống kê chỉ tiêu tiền tệ, tín dụng, rủi ro lãi suất, ngoại hối gửi NHNN")
        ])
    
    conn.commit()
    conn.close()

def parse_cic_excel(db_path=DB_PATH):
    """
    Parses Table_Mo_ta (2).xlsx which contains 23 sheets of CIC data dictionary.
    """
    file_path = os.path.join(TT35_DIR, r"CIC\Table_Mo_ta (2).xlsx")
    if not os.path.exists(file_path):
        file_path = os.path.join(TT35_DIR, r"CIC\Table_Mo_ta.xlsx")
    if not os.path.exists(file_path):
        return 0

    wb = openpyxl.load_workbook(file_path, data_only=True)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    
    count = 0
    # Map sheet name to user-friendly report name
    for sheet_name in wb.sheetnames:
        if sheet_name in ['Tổng hợp', 'CIC_MAP_MACN_BRANCH', 'CIC_BANG_MA_MAP']:
            continue
        
        ws = wb[sheet_name]
        # Row 1 often contains title: e.g. "Thông tin Khế ước Hợp đồng Vay"
        r1_val = ws.cell(row=1, column=1).value
        report_desc = str(r1_val).strip() if r1_val else sheet_name
        
        report_code = sheet_name.replace("CIC_", "")
        target_table = sheet_name.upper()
        
        # Insert or ignore report
        cur.execute("""
        INSERT OR REPLACE INTO reports (report_code, system_code, report_name, target_table, regulation_ref, description)
        VALUES (?, 'CIC', ?, ?, 'Quyết định 573/QĐ-NHNN', ?)
        """, (report_code, report_desc, target_table, f"Chỉ tiêu nhóm {report_code} - {report_desc}"))
        
        # Scan data rows (starting row 3)
        for r in range(3, ws.max_row + 1):
            col_b = ws.cell(row=r, column=2).value # Table
            col_c = ws.cell(row=r, column=3).value # Column / Code
            col_d = ws.cell(row=r, column=4).value # DataType
            col_g = ws.cell(row=r, column=7).value # Note / Viet Name
            col_h = ws.cell(row=r, column=8).value # Condition / Rule
            
            if not col_c:
                continue
            
            field_code = str(col_c).strip()
            tbl_name = str(col_b).strip() if col_b else target_table
            data_type = str(col_d).strip() if col_d else "VARCHAR2(50)"
            field_name_vi = str(col_g).strip() if col_g else field_code
            notes = str(col_h).strip() if col_h else ""
            
            # Lookup ref if mentioned
            lookup_ref = ""
            if "bảng mã" in field_name_vi.lower():
                m = re.search(r"bảng mã\s*(\d+/\w+)", field_name_vi, re.IGNORECASE)
                if m:
                    lookup_ref = f"Bảng mã {m.group(1)}"
            
            cur.execute("""
            INSERT INTO mappings (
                system_code, report_code, field_code, field_name_vi, data_type,
                target_table, target_column, lookup_ref, notes, regulatory_ref, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'QĐ 573/QĐ-NHNN', 'Active')
            """, ('CIC', report_code, field_code, field_name_vi, data_type, tbl_name, field_code, lookup_ref, notes))
            count += 1
            
    conn.commit()
    conn.close()
    return count

def parse_cic_pck(db_path=DB_PATH):
    """
    Parses pk_cic_calc.pck to enrich mappings with:
      - Source table
      - Source column / expression
      - Procedure / Function implementation
    """
    pck_path = os.path.join(CIC_DIR, "pk_cic_calc.pck")
    if not os.path.exists(pck_path):
        return 0

    with open(pck_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    # Regex find functions: function fn_CIC_xxx ... begin ... end;
    fn_pattern = re.compile(r"function\s+(fn_CIC_\w+).*?begin\s*(.*?)\s*end\s*;", re.DOTALL | re.IGNORECASE)
    
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    
    updated = 0
    for match in fn_pattern.finditer(content):
        fn_name = match.group(1)
        body = match.group(2)
        
        # Extract target table from insert into <TABLE>
        ins_match = re.search(r"insert\s+into\s+(\w+)", body, re.IGNORECASE)
        if not ins_match:
            continue
        target_tbl = ins_match.group(1).upper()
        
        # Extract source table from from <TABLE>
        from_match = re.search(r"from\s+([A-Za-z0-9_#]+)(?:\s+([a-zA-Z0-9_]+))?", body, re.IGNORECASE)
        src_tbl = from_match.group(1).upper() if from_match else ""
        src_alias = from_match.group(2) if from_match and from_match.group(2) else ""
        
        # Extract select columns: e.g. a.account_number KU001, ---- S? kh? u?c
        # or to_char(a.value_date, 'YYYYMMDD') KU002, ---- Ngy gi?i ngn
        col_pattern = re.compile(r"^\s*(.*?)\s+([A-Za-z0-9_]+)\s*,\s*[-]{2,}\s*(.*?)$", re.MULTILINE)
        for col_match in col_pattern.finditer(body):
            expr = col_match.group(1).strip()
            field_code = col_match.group(2).strip().upper()
            comment_note = col_match.group(3).strip()
            
            # Detect source column from expression
            clean_col = ""
            if src_alias and expr.startswith(f"{src_alias}."):
                clean_col = expr[len(src_alias)+1:]
            elif "." in expr:
                clean_col = expr.split(".")[-1]
            else:
                clean_col = expr
            
            # Clean up field name
            rule_desc = f"Biểu thức: {expr}" if expr else ""
            
            # Update mapping record
            cur.execute("""
            UPDATE mappings 
            SET source_system = 'ODS/Core',
                source_table = coalesce(nullif(?, ''), source_table),
                source_column = ?,
                transformation_rule = ?,
                implemented_by = ?
            WHERE system_code = 'CIC' AND target_table = ? AND field_code = ?
            """, (src_tbl, clean_col, rule_desc, f"PK_CIC_CALC.{fn_name}", target_tbl, field_code))
            
            if cur.rowcount > 0:
                updated += cur.rowcount
            else:
                # If not found, insert
                report_code = target_tbl.replace("CIC_", "")
                cur.execute("""
                INSERT INTO mappings (
                    system_code, report_code, field_code, field_name_vi, data_type,
                    target_table, target_column, source_system, source_table, source_column,
                    transformation_rule, implemented_by, regulatory_ref, status
                ) VALUES (?, ?, ?, ?, 'VARCHAR2(50)', ?, ?, 'ODS/Core', ?, ?, ?, ?, 'QĐ 573/QĐ-NHNN', 'Active')
                """, ('CIC', report_code, field_code, comment_note or field_code, target_tbl, field_code,
                      src_tbl, clean_col, rule_desc, f"PK_CIC_CALC.{fn_name}"))
                updated += 1

    conn.commit()
    conn.close()
    return updated

def parse_tt35_procedures(db_path=DB_PATH):
    """
    Parses ODS_TUANPV/pr_*.prc files for TT35 reports.
    """
    prc_files = glob.glob(os.path.join(ODS_DIR, "pr_*.prc"))
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    
    count = 0
    for p_path in prc_files:
        p_name = os.path.basename(p_path)
        base_name = os.path.splitext(p_name)[0].upper()
        # Report code usually follows pr_: e.g. pr_b00214 -> B00214
        rpt_code = base_name.replace("PR_", "")
        
        with open(p_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
            
        # Extract Purpose comment
        purpose_match = re.search(r"Purpose:\s*(.*?)(?:\n|\*\/)", content, re.IGNORECASE)
        purpose = purpose_match.group(1).strip() if purpose_match else f"Báo cáo thống kê {rpt_code}"
        
        # Target table: RPTB_B00214 or similar
        target_tbl_match = re.search(r"(?:insert\s+into|update|delete\s+from|delete)\s+(RPTB_[A-Za-z0-9_]+)", content, re.IGNORECASE)
        target_tbl = target_tbl_match.group(1).upper() if target_tbl_match else f"RPTB_{rpt_code}"
        
        # Insert report
        cur.execute("""
        INSERT OR REPLACE INTO reports (report_code, system_code, report_name, target_table, procedure_name, regulation_ref, description)
        VALUES (?, 'TT35', ?, ?, ?, 'Thông tư 35/2015/TT-NHNN', ?)
        """, (rpt_code, purpose, target_tbl, base_name, purpose))
        
        # Find updates or inserts with transformation logic
        # e.g. update RPTB_B00214 set NGAN_HAN = (select sum(...) from ...) where STT in (...)
        updates = re.findall(r"update\s+\w+\s+set\s+([A-Za-z0-9_]+)\s*=\s*(.*?)\s*where\s*(.*?);", content, re.DOTALL | re.IGNORECASE)
        for col_name, expr, where_clause in updates:
            col_name = col_name.strip().upper()
            clean_expr = " ".join(expr.split())
            clean_where = " ".join(where_clause.split())
            rule = f"Cập nhật: {clean_expr} | Điều kiện: {clean_where}"
            
            # Find source table from expr
            src_match = re.search(r"from\s+([A-Za-z0-9_#]+)", clean_expr, re.IGNORECASE)
            src_tbl = src_match.group(1).upper() if src_match else "ODS"
            
            cur.execute("""
            INSERT INTO mappings (
                system_code, report_code, field_code, field_name_vi, data_type,
                target_table, target_column, source_system, source_table, source_column,
                transformation_rule, implemented_by, regulatory_ref, status
            ) VALUES (?, ?, ?, ?, 'NUMBER(18,2)', ?, ?, 'ODS', ?, ?, ?, ?, 'Thông tư 35/2015/TT-NHNN', 'Active')
            """, ('TT35', rpt_code, col_name, f"Chỉ tiêu {col_name} ({clean_where[:30]})", target_tbl, col_name,
                  src_tbl, col_name, rule, base_name))
            count += 1
            
    conn.commit()
    conn.close()
    return count

def parse_tt35_excel_templates(db_path=DB_PATH):
    """
    Parses Excel template files in E:\1080 Public Bank\TT35\TT35 new
    e.g. B00214.xlsx, G02437.xlsx...
    """
    excel_dir = os.path.join(TT35_DIR, "TT35 new")
    if not os.path.exists(excel_dir):
        return 0

    excel_files = glob.glob(os.path.join(excel_dir, "*.xlsx"))
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    
    count = 0
    for ef in excel_files:
        fn = os.path.basename(ef)
        rpt_code = fn.split(".")[0].split(" ")[0].split("-")[0].upper()
        if not rpt_code or len(rpt_code) < 3:
            continue

        try:
            wb = openpyxl.load_workbook(ef, data_only=True)
            sheet_name = wb.sheetnames[0]
            ws = wb[sheet_name]
            
            # Read Help text from row with 'Help'
            help_text = ""
            for r in range(15, min(ws.max_row + 1, 40)):
                c1 = ws.cell(row=r, column=1).value
                if c1 and "help" in str(c1).lower():
                    help_text = str(ws.cell(row=r, column=2).value or "")[:500]
                    break
            
            # Scan rows around 15-30 for indicators
            for r in range(16, min(ws.max_row + 1, 35)):
                c1 = ws.cell(row=r, column=1).value
                c2 = ws.cell(row=r, column=2).value
                c3 = ws.cell(row=r, column=3).value
                c4 = ws.cell(row=r, column=4).value
                
                # Check for R-xxx or numeric STT
                if c1 and str(c1).startswith("R-"):
                    field_code = str(c1).strip()
                    field_name = str(c3).strip() if c3 else str(c2).strip()
                    rule = str(c4).strip() if c4 else ""
                    
                    cur.execute("""
                    INSERT INTO mappings (
                        system_code, report_code, field_code, field_name_vi, data_type,
                        target_table, target_column, transformation_rule, regulatory_ref, notes, status
                    ) VALUES (?, ?, ?, ?, 'NUMBER(18,2)', ?, ?, ?, 'Thông tư 35/2015/TT-NHNN', ?, 'Active')
                    """, ('TT35', rpt_code, field_code, field_name, f"RPTB_{rpt_code}", field_code, rule, help_text[:200]))
                    count += 1
        except Exception:
            continue

    conn.commit()
    conn.close()
    return count

def run_full_sync():
    """
    Executes full synchronization from all local files into SQLite database.
    """
    init_db()
    c1 = parse_cic_excel()
    c2 = parse_cic_pck()
    c3 = parse_tt35_procedures()
    c4 = parse_tt35_excel_templates()
    return {
        "status": "success",
        "cic_excel_fields": c1,
        "cic_pck_enriched": c2,
        "tt35_procedures_mapped": c3,
        "tt35_excel_mapped": c4,
        "total": c1 + c2 + c3 + c4
    }

if __name__ == "__main__":
    print("Initializing DB and running full sync...")
    res = run_full_sync()
    print("Sync completed result:", res)
