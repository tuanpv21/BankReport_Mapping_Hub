# -*- coding: utf-8 -*-
"""
Smart PL/SQL Procedure & Package Parser for BankReport Mapping Hub.
Parses Oracle PL/SQL code:
 - INSERT INTO ... SELECT ...
 - MERGE INTO ... USING (SELECT ...) ... UPDATE SET ...
 - UPDATE ... SET ...
 - OPEN <cursor> FOR SELECT ...
 - COMMENT ON COLUMN ...
and extracts structured column mappings, source tables, and transformation rules.
"""

import re

# Common banking column descriptions
BANK_COLUMN_DICT = {
    "DNDK": ("Dư nợ đầu kỳ", "NUMBER(18,2)"),
    "DCDK": ("Dư có đầu kỳ", "NUMBER(18,2)"),
    "PSN": ("Phát sinh Nợ trong kỳ", "NUMBER(18,2)"),
    "PSC": ("Phát sinh Có trong kỳ", "NUMBER(18,2)"),
    "DNCK": ("Dư nợ cuối kỳ", "NUMBER(18,2)"),
    "DCCK": ("Dư có cuối kỳ", "NUMBER(18,2)"),
    "SO_DU": ("Số dư kỳ báo cáo", "NUMBER(18,2)"),
    "SO_TIEN_THAY_DOI": ("Số tiền tăng/giảm so với kỳ trước", "NUMBER(18,2)"),
    "TYLE_THAY_DOI": ("Tỷ lệ % thay đổi", "NUMBER(10,4)"),
    "RPT_DT": ("Ngày báo cáo", "VARCHAR2(8)"),
    "ID": ("Mã chỉ tiêu (Key)", "VARCHAR2(20)"),
    "STT": ("Số thứ tự", "NUMBER(5)"),
    "CHI_TIEU": ("Tên chỉ tiêu báo cáo", "VARCHAR2(250)"),
    "TEN_CHI_TIEU": ("Tên chỉ tiêu báo cáo", "VARCHAR2(250)"),
    "TEN_TAI_KHOAN": ("Tên tài khoản kế toán", "VARCHAR2(250)"),
    "SO_HIEU_TAI_KHOAN": ("Số hiệu tài khoản GL", "VARCHAR2(30)"),
    "GL_CODE": ("Mã tài khoản GL", "VARCHAR2(30)"),
    "KY_TU_SO": ("Ký tự số lọc GL", "VARCHAR2(30)"),
    "KY_TU_CHU": ("Ký tự chữ lọc GL", "VARCHAR2(30)"),
    "CAL_STEP1": ("Công thức cấu hình GL", "VARCHAR2(250)"),
    "GIA_TRI": ("Giá trị tổng hợp", "NUMBER(18,2)"),
    "MA_KH": ("Mã khách hàng", "VARCHAR2(20)"),
    "TEN_KH": ("Tên khách hàng", "VARCHAR2(250)"),
    "NGAY_MO": ("Ngày mở tài khoản/HĐ", "VARCHAR2(8)"),
    "NGAY_DAO_HAN": ("Ngày đáo hạn", "VARCHAR2(8)"),
    "LAI_SUAT": ("Lãi suất (%)", "NUMBER(8,4)"),
    "TIEN_TE": ("Loại tiền tệ", "VARCHAR2(3)"),
}


def parse_sql_procedure(sql_text, target_table_hint=None, system_hint=None):
    if not sql_text or not str(sql_text).strip():
        return {
            "success": False,
            "error": "Nội dung SQL rỗng",
            "fields": []
        }

    clean_sql = str(sql_text).strip()
    
    # 0. Extract Procedure Name
    proc_name = ""
    proc_match = re.search(r"create\s+(?:or\s+replace\s+)?procedure\s+([A-Za-z0-9_]+)", clean_sql, re.IGNORECASE)
    if proc_match:
        proc_name = proc_match.group(1).upper()

    # 1. Target table detection
    target_table = (target_table_hint or "").strip().upper()
    all_target_candidates = []
    
    # Check MERGE INTO
    for mm in re.finditer(r"merge\s+into\s+([A-Za-z0-9_#]+)", clean_sql, re.IGNORECASE):
        tbl = mm.group(1).upper()
        if not tbl.startswith("TWT_") and tbl not in all_target_candidates:
            all_target_candidates.append(tbl)
            
    # Check UPDATE
    for um in re.finditer(r"update\s+([A-Za-z0-9_#]+)", clean_sql, re.IGNORECASE):
        tbl = um.group(1).upper()
        if not tbl.startswith("TWT_") and tbl not in all_target_candidates:
            all_target_candidates.append(tbl)

    # Check INSERT INTO
    for im in re.finditer(r"insert\s+into\s+([A-Za-z0-9_#]+)", clean_sql, re.IGNORECASE):
        tbl = im.group(1).upper()
        if not tbl.startswith("TWT_") and not tbl.endswith("_HIS") and tbl not in all_target_candidates:
            all_target_candidates.append(tbl)

    # Check OPEN CURSOR FOR SELECT ... FROM <tbl>
    for cm in re.finditer(r"open\s+[A-Za-z0-9_]+\s+for\s+select\s+.*?\s+from\s+([A-Za-z0-9_#]+)", clean_sql, re.IGNORECASE | re.DOTALL):
        tbl = cm.group(1).upper()
        if not tbl.startswith("TWT_") and tbl not in all_target_candidates:
            all_target_candidates.append(tbl)

    if not target_table:
        # Prefer RPTB_ or CIC_ tables
        for cand in all_target_candidates:
            if cand.startswith("RPTB_") or cand.startswith("CIC_"):
                target_table = cand
                break
        if not target_table and all_target_candidates:
            target_table = all_target_candidates[0]

    # 2. Source tables discovery (FROM / JOIN)
    source_tables = []
    source_alias_map = {}
    from_matches = re.finditer(r"(?:from|join)\s+([A-Za-z0-9_#]+)(?:\s+([a-zA-Z0-9_]+))?", clean_sql, re.IGNORECASE)
    for fm in from_matches:
        tbl = fm.group(1).upper()
        alias = fm.group(2) if fm.group(2) else ""
        if tbl not in ["DUAL", "WHERE", "SELECT", "SET", "ON", "INNER", "LEFT", "RIGHT", "OUTER"] and not tbl.startswith("TWT_"):
            if tbl != target_table and not tbl.endswith("_HIS") and tbl not in source_tables:
                source_tables.append(tbl)
            if alias and tbl != "DUAL":
                source_alias_map[alias.lower()] = tbl

    primary_source_table = source_tables[0] if source_tables else ""

    # Result fields storage: keyed by uppercase column name
    column_dict = {}

    def add_or_update_column(col_name, **kwargs):
        col_name = col_name.strip().upper()
        if not col_name or col_name in ["FROM", "WHERE", "GROUP", "ORDER", "JOIN", "SET", "AND", "OR", "ON", "WHEN", "SELECT"]:
            return
        if col_name not in column_dict:
            dict_vi, dict_type = BANK_COLUMN_DICT.get(col_name, (col_name, "VARCHAR2(50)"))
            column_dict[col_name] = {
                "table": target_table,
                "column": col_name,
                "data_type": dict_type,
                "field_name_vi": dict_vi,
                "notes": "",
                "source_table": primary_source_table,
                "source_column": "",
                "transformation_rule": ""
            }
        item = column_dict[col_name]
        for k, v in kwargs.items():
            if v:  # Only update if non-empty
                if k == "transformation_rule" and item["transformation_rule"] and item["transformation_rule"] != v:
                    # Append transformation if different
                    if v not in item["transformation_rule"]:
                        item["transformation_rule"] += f" | {v}"
                elif k == "notes" and item["notes"] and item["notes"] != v:
                    if v not in item["notes"]:
                        item["notes"] += f" | {v}"
                else:
                    item[k] = v

    # ==========================================================
    # Pass 1: OPEN CURSOR FOR SELECT col1, col2, ... FROM ...
    # ==========================================================
    cursor_match = re.search(r"open\s+[a-zA-Z0-9_]+\s+for\s+select\s+(.*?)\s+from\s+([A-Za-z0-9_#]+)", clean_sql, re.IGNORECASE | re.DOTALL)
    if cursor_match:
        cur_cols_str = cursor_match.group(1).strip()
        for raw_c in cur_cols_str.split(","):
            raw_c = raw_c.strip()
            # If alias exists: expr as col_name or expr col_name
            parts = re.split(r"\s+(?:as\s+)?", raw_c, flags=re.IGNORECASE)
            col_name = parts[-1].strip().upper()
            if col_name:
                add_or_update_column(col_name, notes="Cột xuất ra báo cáo (RefCursor)")

    # ==========================================================
    # Pass 2: MERGE INTO ... USING (SELECT ...) ... UPDATE SET ...
    # ==========================================================
    merge_blocks = re.finditer(
        r"merge\s+into\s+([A-Za-z0-9_#]+)(?:\s+([a-zA-Z0-9_]+))?\s+using\s*\((.*?)\)\s*([a-zA-Z0-9_]+)?\s+on\s*\(.*?\)\s*when\s+matched\s+then\s*update\s+set\s*(.*?)(?:where|when|commit|;|end\s+loop)",
        clean_sql,
        re.IGNORECASE | re.DOTALL
    )
    for mb in merge_blocks:
        m_target = mb.group(1).upper()
        using_sql = mb.group(3).strip()
        set_clause = mb.group(5).strip()

        # Parse source table from USING
        src_tbl_using = ""
        m_from = re.search(r"from\s+([A-Za-z0-9_#]+)", using_sql, re.IGNORECASE)
        if m_from:
            cand = m_from.group(1).upper()
            if cand != "DUAL" and not cand.startswith("TWT_") and cand != target_table:
                src_tbl_using = cand

        # Extract projected columns in USING (select col1 as alias1, ...)
        using_select_cols = {}
        using_sel_match = re.search(r"^\s*select\s+(.*?)\s+from\b", using_sql, re.IGNORECASE | re.DOTALL)
        if using_sel_match:
            raw_cols = using_sel_match.group(1).strip()
            # split by comma, ignoring commas inside parentheses
            parts = re.split(r",\s*(?![^()]*\))", raw_cols)
            for p in parts:
                p = p.strip()
                tokens = re.split(r"\s+(?:as\s+)?", p, flags=re.IGNORECASE)
                if len(tokens) >= 2:
                    col_expr = " ".join(tokens[:-1]).strip()
                    alias = tokens[-1].strip().upper()
                    using_select_cols[alias] = col_expr
                elif tokens:
                    alias = tokens[0].strip().upper()
                    using_select_cols[alias] = alias

        # Parse assignments: a.COL = b.ALIAS
        assignments = re.split(r",\s*(?![^()]*\))", set_clause)
        for assign in assignments:
            if "=" in assign:
                left, right = assign.split("=", 1)
                left_col = left.strip().split(".")[-1].upper()
                right_expr = right.strip()
                right_token = right_expr.split(".")[-1].upper()

                matched_expr = using_select_cols.get(right_token, right_expr)
                
                # Check source column from matched_expr
                src_col_found = ""
                m_scol = re.search(r"\b([A-Za-z0-9_]+)\.(NO_CUOI_KY|CO_CUOI_KY|NO_DAU_KY|CO_DAU_KY|NO_PHAT_SINH|CO_PHAT_SINH|[A-Za-z0-9_]+)\b", matched_expr, re.IGNORECASE)
                if m_scol:
                    src_col_found = m_scol.group(2).upper()
                elif right_token in using_select_cols:
                    src_col_found = right_token

                # Guess comment/vi name
                dict_vi, _ = BANK_COLUMN_DICT.get(left_col, (left_col, ""))

                add_or_update_column(
                    left_col,
                    table=m_target if not m_target.startswith("TWT_") else target_table,
                    source_table=src_tbl_using or primary_source_table,
                    source_column=src_col_found,
                    transformation_rule=f"MERGE: {matched_expr}",
                    notes=f"Tính từ {src_tbl_using or 'bảng nguồn'}" if src_tbl_using else ""
                )

    # ==========================================================
    # Pass 3: UPDATE <tbl> SET <col> = <expr> ...
    # ==========================================================
    update_blocks = re.finditer(
        r"update\s+([A-Za-z0-9_#]+)(?:\s+[a-zA-Z0-9_]+)?\s+set\s+(.*?)(?:where|commit|;)",
        clean_sql,
        re.IGNORECASE | re.DOTALL
    )
    for ub in update_blocks:
        tbl_upd = ub.group(1).upper()
        if tbl_upd.startswith("TWT_"):
            continue
        set_body = ub.group(2).strip()
        assignments = re.split(r",\s*(?![^()]*\))", set_body)
        for assign in assignments:
            if "=" in assign:
                left, right = assign.split("=", 1)
                left_col = left.strip().split(".")[-1].upper()
                right_expr = " ".join(right.strip().split())
                
                # If it is empty reset (e.g. SO_DU = ''), just make sure column is recorded
                if right_expr in ["''", "null", "NULL"]:
                    add_or_update_column(left_col, table=tbl_upd)
                    continue

                # Check if right_expr contains a subquery (SELECT ... FROM ...)
                src_tbl_sub = ""
                m_from = re.search(r"from\s+([A-Za-z0-9_#]+)", right_expr, re.IGNORECASE)
                if m_from:
                    src_tbl_sub = m_from.group(1).upper()

                add_or_update_column(
                    left_col,
                    table=tbl_upd or target_table,
                    transformation_rule=f"UPDATE: {right_expr}",
                    source_table=src_tbl_sub or (target_table if "(" in right_expr and "so_du" in right_expr.lower() else primary_source_table)
                )

    # ==========================================================
    # Pass 4: INSERT INTO ... SELECT (with inline comments)
    # ==========================================================
    # Pattern: expr COL_NAME, ---- COMMENT
    # Line must NOT start with comment '--'
    col_pattern = re.compile(r"^\s*([^-\n][^,\n]*?)\s+([A-Za-z0-9_]+)\s*(?:,)?\s*[-]{2,}\s*(.*?)$", re.MULTILINE)
    for m in col_pattern.finditer(clean_sql):
        raw_expr = m.group(1).strip()
        col_name = m.group(2).strip().upper()
        comment = m.group(3).strip()

        # Skip comment lines or lines with comparison operators (=, <, >, <=, >=)
        if raw_expr.startswith("--") or "=" in raw_expr or "<" in raw_expr or ">" in raw_expr:
            continue

        clean_expr = re.sub(r"^\s*select\s+", "", raw_expr, flags=re.IGNORECASE).strip()
        if col_name in ["FROM", "WHERE", "GROUP", "ORDER", "JOIN", "SET", "AND", "OR", "ON", "WHEN", "THEN", "ELSE", "END", "LOOP", "COMMIT", "INTO"]:
            continue

        # Skip procedure parameter names
        if col_name in ["P_DATE", "V_DATE", "V_PRE_DATE", "P_PERIOD"]:
            continue

        src_table = primary_source_table
        src_col = ""
        m_col = re.search(r"\b([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\b", clean_expr)
        if m_col:
            alias = m_col.group(1).lower()
            src_col = m_col.group(2).upper()
            if alias in source_alias_map:
                src_table = source_alias_map[alias]
        elif clean_expr.isidentifier() and clean_expr.lower() not in ["null", "sysdate"]:
            src_col = clean_expr.upper()

        add_or_update_column(
            col_name,
            field_name_vi=comment or col_name,
            source_table=src_table,
            source_column=src_col,
            transformation_rule=f"Cột nguồn: {clean_expr}" if clean_expr and clean_expr != src_col else ""
        )

    # ==========================================================
    # Pass 5: COMMENT ON COLUMN <tbl>.<col> IS '<text>';
    # ==========================================================
    comment_pattern = re.compile(r"comment\s+on\s+column\s+([A-Za-z0-9_#]+)\.([A-Za-z0-9_]+)\s+is\s+'(.*?)'\s*;", re.IGNORECASE)
    for cm in comment_pattern.finditer(clean_sql):
        c_tbl = cm.group(1).upper()
        c_col = cm.group(2).upper()
        c_text = cm.group(3).strip()
        add_or_update_column(c_col, table=c_tbl, field_name_vi=c_text)

    # Refine field results
    fields = list(column_dict.values())

    # Sort fields logically: ID, STT, CHI_TIEU first, then numbers/amounts, then RPT_DT last
    def col_sort_key(f):
        c = f["column"]
        if c == "ID":
            return (0, 0)
        if c in ["STT", "ORD"]:
            return (0, 1)
        if c in ["CHI_TIEU", "TEN_CHI_TIEU", "TEN_TAI_KHOAN"]:
            return (0, 2)
        if c in ["SO_HIEU_TAI_KHOAN", "GL_CODE"]:
            return (0, 3)
        if c in ["RPT_DT", "REPORT_DATE", "CREATED_DATE"]:
            return (2, 0)
        return (1, c)

    fields.sort(key=col_sort_key)

    return {
        "success": True,
        "procedure_name": proc_name,
        "target_table": target_table,
        "source_tables": source_tables,
        "fields_count": len(fields),
        "fields": fields
    }