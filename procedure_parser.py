# -*- coding: utf-8 -*-
"""
Smart PL/SQL Procedure & Package Parser for BankReport Mapping Hub.
Parses Oracle PL/SQL code (INSERT INTO ... SELECT ..., UPDATE ... SET ..., COMMENTS)
and extracts structured column mappings.
"""

import re

def parse_sql_procedure(sql_text, target_table_hint=None, system_hint=None):
    if not sql_text or not sql_text.strip():
        return {
            "success": False,
            "error": "Nội dung SQL rỗng",
            "fields": []
        }

    lines = sql_text.splitlines()
    clean_sql = "\n".join(lines)

    # 1. Detect target table
    target_table = target_table_hint or ""
    if not target_table:
        ins_match = re.search(r"insert\s+into\s+([A-Za-z0-9_#]+)", clean_sql, re.IGNORECASE)
        if ins_match:
            target_table = ins_match.group(1).upper()
        else:
            upd_match = re.search(r"update\s+([A-Za-z0-9_#]+)", clean_sql, re.IGNORECASE)
            if upd_match:
                target_table = upd_match.group(1).upper()

    # 2. Detect source tables (FROM / JOIN)
    source_tables = []
    source_alias_map = {}
    from_matches = re.finditer(r"(?:from|join)\s+([A-Za-z0-9_#]+)(?:\s+([a-zA-Z0-9_]+))?", clean_sql, re.IGNORECASE)
    for fm in from_matches:
        tbl = fm.group(1).upper()
        alias = fm.group(2) if fm.group(2) else ""
        if tbl not in ["DUAL", "WHERE", "SELECT", "SET"] and tbl not in source_tables:
            source_tables.append(tbl)
            if alias:
                source_alias_map[alias.lower()] = tbl

    primary_source_table = source_tables[0] if source_tables else ""

    fields = []
    seen_columns = set()

    # Pattern A: SELECT <expression> <COLUMN_NAME>, ---- <COMMENT>
    # e.g. a.customer_id TTC03, ---- Mã KH
    # e.g. to_char(a.value_date, 'YYYYMMDD') KU002, ---- Ngày giải ngân
    col_pattern = re.compile(r"^\s*(.*?)\s+([A-Za-z0-9_]+)\s*(?:,)?\s*[-]{2,}\s*(.*?)$", re.MULTILINE)
    for m in col_pattern.finditer(clean_sql):
        raw_expr = m.group(1).strip()
        col_name = m.group(2).strip().upper()
        comment = m.group(3).strip()

        # Clean SELECT if at the beginning
        clean_expr = re.sub(r"^\s*select\s+", "", raw_expr, flags=re.IGNORECASE).strip()

        # Skip reserved words
        if col_name in ["FROM", "WHERE", "GROUP", "ORDER", "JOIN", "SET", "AND", "OR"]:
            continue

        # Extract source column and table
        src_table = primary_source_table
        src_col = ""
        m_col = re.search(r"\b([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\b", clean_expr)
        if m_col:
            alias = m_col.group(1).lower()
            src_col = m_col.group(2)
            if alias in source_alias_map:
                src_table = source_alias_map[alias]
        elif clean_expr.isidentifier() and clean_expr.lower() not in ["null", "sysdate"]:
            src_col = clean_expr

        # Determine data type hint
        data_type = "VARCHAR2(50)"
        if "to_char" in clean_expr.lower() and "yyyymmdd" in clean_expr.lower():
            data_type = "VARCHAR2(8)"
        elif clean_expr.isdigit() or any(k in col_name.lower() for k in ["amt", "bal", "rate", "du_no", "so_tien"]):
            data_type = "NUMBER(18,2)"

        fields.append({
            "table": target_table,
            "column": col_name,
            "data_type": data_type,
            "field_name_vi": comment or col_name,
            "notes": "",
            "source_table": src_table,
            "source_column": src_col,
            "transformation_rule": f"Biểu thức: {clean_expr}" if clean_expr and clean_expr != src_col else (clean_expr or "")
        })
        seen_columns.add(col_name)

    # Pattern B: UPDATE <tbl> SET <COLUMN> = (SELECT ... FROM <src_tbl> ...) WHERE <condition>;
    # e.g. update RPTB_B00214 set NGAN_HAN = (select sum(OUTSTANDING) from ods_loan_account_hist ...) where STT in (1,6);
    update_pattern = re.compile(r"update\s+([A-Za-z0-9_#]+)\s+set\s+([A-Za-z0-9_]+)\s*=\s*(.*?)\s*where\s*(.*?);", re.DOTALL | re.IGNORECASE)
    for um in update_pattern.finditer(clean_sql):
        tbl_upd = um.group(1).upper()
        col_upd = um.group(2).upper()
        expr_upd = " ".join(um.group(3).split())
        where_upd = " ".join(um.group(4).split())

        if col_upd in seen_columns:
            continue

        # Extract source table from inside subquery
        src_tbl_sub = ""
        m_src = re.search(r"from\s+([A-Za-z0-9_#]+)", expr_upd, re.IGNORECASE)
        if m_src:
            src_tbl_sub = m_src.group(1).upper()

        # Extract source column
        src_col_sub = ""
        m_scol = re.search(r"\b(?:sum|max|min|avg)?\s*\(\s*([A-Za-z0-9_]+)\s*\)", expr_upd, re.IGNORECASE)
        if m_scol:
            src_col_sub = m_scol.group(1)

        fields.append({
            "table": tbl_upd or target_table,
            "column": col_upd,
            "data_type": "NUMBER(18,2)",
            "field_name_vi": f"Chỉ tiêu {col_upd} ({where_upd[:35]})",
            "notes": f"Điều kiện: {where_upd}",
            "source_table": src_tbl_sub or primary_source_table,
            "source_column": src_col_sub or col_upd,
            "transformation_rule": expr_upd
        })
        seen_columns.add(col_upd)

    # Pattern C: Comment on column <tbl>.<col> is '<comment>';
    comment_pattern = re.compile(r"comment\s+on\s+column\s+([A-Za-z0-9_#]+)\.([A-Za-z0-9_]+)\s+is\s+'(.*?)'\s*;", re.IGNORECASE)
    for cm in comment_pattern.finditer(clean_sql):
        c_tbl = cm.group(1).upper()
        c_col = cm.group(2).upper()
        c_text = cm.group(3).strip()

        # If already in fields, update comment
        found = False
        for f in fields:
            if f["column"] == c_col:
                f["field_name_vi"] = c_text
                found = True
                break

        if not found and c_col not in seen_columns:
            fields.append({
                "table": c_tbl,
                "column": c_col,
                "data_type": "VARCHAR2(50)",
                "field_name_vi": c_text,
                "notes": "",
                "source_table": primary_source_table,
                "source_column": "",
                "transformation_rule": ""
            })
            seen_columns.add(c_col)

    return {
        "success": True,
        "target_table": target_table,
        "source_tables": source_tables,
        "fields_count": len(fields),
        "fields": fields
    }