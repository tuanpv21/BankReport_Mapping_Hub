# -*- coding: utf-8 -*-
"""
BankReport Mapping Hub - Web Application Backend (Flask + SQLite)
Manages Business-to-Technical Mappings for TT35 and CIC systems.
Supports interactive PL/SQL Procedure/Package Parsing & Studio.
"""

import os
import sys
import json
import sqlite3
import webbrowser
import threading
from flask import Flask, render_template, render_template_string, request, jsonify, send_file
from werkzeug.utils import secure_filename
import openpyxl

import ingest_parser as parser
import procedure_parser
import exporter

app = Flask(__name__)
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

DB_PATH = os.path.join(os.path.dirname(__file__), "mapping_hub.db")
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
HTML_PATH = os.path.join(os.path.dirname(__file__), "templates", "index.html")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

# Ensure schema migrations
def init_schema():
    conn = get_db()
    cur = conn.cursor()
    cols = [r[1] for r in cur.execute("PRAGMA table_info(reports)").fetchall()]
    if "procedure_code" not in cols:
        cur.execute("ALTER TABLE reports ADD COLUMN procedure_code TEXT")
        conn.commit()
    conn.close()

init_schema()

@app.route("/")
def index():
    if os.path.exists(HTML_PATH) and os.path.getsize(HTML_PATH) > 0:
        with open(HTML_PATH, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    return render_template("index.html")

@app.route("/api/stats")
def get_stats():
    conn = get_db()
    cur = conn.cursor()
    
    total_reports = cur.execute("SELECT COUNT(*) FROM reports").fetchone()[0]
    total_mappings = cur.execute("SELECT COUNT(*) FROM mappings").fetchone()[0]
    mapped_source = cur.execute("SELECT COUNT(*) FROM mappings WHERE source_table IS NOT NULL AND source_table != ''").fetchone()[0]
    unmapped_source = total_mappings - mapped_source
    
    cic_reports = cur.execute("SELECT COUNT(*) FROM reports WHERE system_code = 'CIC'").fetchone()[0]
    cic_mappings = cur.execute("SELECT COUNT(*) FROM mappings WHERE system_code = 'CIC'").fetchone()[0]
    
    tt35_reports = cur.execute("SELECT COUNT(*) FROM reports WHERE system_code = 'TT35'").fetchone()[0]
    tt35_mappings = cur.execute("SELECT COUNT(*) FROM mappings WHERE system_code = 'TT35'").fetchone()[0]
    
    conn.close()
    return jsonify({
        "total_reports": total_reports,
        "total_mappings": total_mappings,
        "mapped_source": mapped_source,
        "unmapped_source": unmapped_source,
        "mapped_ratio": round((mapped_source / total_mappings * 100), 1) if total_mappings else 0,
        "cic": {"reports": cic_reports, "mappings": cic_mappings},
        "tt35": {"reports": tt35_reports, "mappings": tt35_mappings}
    })

@app.route("/api/reports")
def get_reports():
    system = request.args.get("system")
    conn = get_db()
    cur = conn.cursor()
    
    sql = """
        SELECT r.report_code, r.system_code, r.report_name, r.target_table, 
               r.procedure_name, r.procedure_code, r.regulation_ref, r.description,
               COUNT(m.id) as field_count
        FROM reports r
        LEFT JOIN mappings m ON r.report_code = m.report_code AND r.system_code = m.system_code
        WHERE 1=1
    """
    params = []
    if system and system != "ALL":
        sql += " AND r.system_code = ?"
        params.append(system)
    
    sql += " GROUP BY r.report_code, r.system_code ORDER BY r.system_code, r.report_code"
    rows = cur.execute(sql, params).fetchall()
    conn.close()
    
    return jsonify([dict(r) for r in rows])

@app.route("/api/mappings")
def get_mappings():
    system = request.args.get("system")
    report = request.args.get("report")
    search = request.args.get("search", "").strip()
    status = request.args.get("status")
    
    conn = get_db()
    cur = conn.cursor()
    
    sql = """
        SELECT m.id, m.system_code, m.report_code, r.report_name, m.field_code, 
               m.field_name_vi, m.data_type, m.target_table, m.target_column,
               m.source_system, m.source_table, m.source_column, m.transformation_rule,
               m.lookup_ref, m.implemented_by, m.regulatory_ref, m.status, m.notes, m.updated_at
        FROM mappings m
        LEFT JOIN reports r ON m.report_code = r.report_code AND m.system_code = r.system_code
        WHERE 1=1
    """
    params = []
    if system and system != "ALL":
        sql += " AND m.system_code = ?"
        params.append(system)
    if report and report != "ALL":
        sql += " AND m.report_code = ?"
        params.append(report)
    if status and status != "ALL":
        sql += " AND m.status = ?"
        params.append(status)
    if search:
        s_pat = f"%{search}%"
        sql += """ AND (
            m.field_code LIKE ? OR m.field_name_vi LIKE ? OR 
            m.source_table LIKE ? OR m.source_column LIKE ? OR
            m.transformation_rule LIKE ? OR m.target_table LIKE ? OR
            m.notes LIKE ?
        )"""
        params.extend([s_pat, s_pat, s_pat, s_pat, s_pat, s_pat, s_pat])
        
    sql += " ORDER BY m.system_code, m.report_code, m.id LIMIT 400"
    rows = cur.execute(sql, params).fetchall()
    conn.close()
    
    return jsonify([dict(r) for r in rows])

@app.route("/api/mappings/<int:id>", methods=["GET"])
def get_mapping_detail(id):
    conn = get_db()
    cur = conn.cursor()
    row = cur.execute("SELECT * FROM mappings WHERE id = ?", (id,)).fetchone()
    conn.close()
    if not row:
        return jsonify({"error": "Not found"}), 404
    return jsonify(dict(row))

@app.route("/api/mappings", methods=["POST"])
def create_mapping():
    data = request.json or {}
    conn = get_db()
    cur = conn.cursor()
    
    cur.execute("""
    INSERT INTO mappings (
        system_code, report_code, field_code, field_name_vi, data_type,
        target_table, target_column, source_system, source_table, source_column,
        transformation_rule, lookup_ref, implemented_by, regulatory_ref, status, notes
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data.get("system_code", "CIC"),
        data.get("report_code", ""),
        data.get("field_code", ""),
        data.get("field_name_vi", ""),
        data.get("data_type", "VARCHAR2(50)"),
        data.get("target_table", ""),
        data.get("target_column", data.get("field_code", "")),
        data.get("source_system", "ODS"),
        data.get("source_table", ""),
        data.get("source_column", ""),
        data.get("transformation_rule", ""),
        data.get("lookup_ref", ""),
        data.get("implemented_by", ""),
        data.get("regulatory_ref", ""),
        data.get("status", "Active"),
        data.get("notes", "")
    ))
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    return jsonify({"success": True, "id": new_id})

@app.route("/api/mappings/<int:id>", methods=["PUT"])
def update_mapping(id):
    data = request.json or {}
    conn = get_db()
    cur = conn.cursor()
    
    cur.execute("""
    UPDATE mappings SET
        field_name_vi = ?,
        data_type = ?,
        target_table = ?,
        source_system = ?,
        source_table = ?,
        source_column = ?,
        transformation_rule = ?,
        lookup_ref = ?,
        implemented_by = ?,
        regulatory_ref = ?,
        status = ?,
        notes = ?,
        updated_at = CURRENT_TIMESTAMP
    WHERE id = ?
    """, (
        data.get("field_name_vi"),
        data.get("data_type"),
        data.get("target_table"),
        data.get("source_system"),
        data.get("source_table"),
        data.get("source_column"),
        data.get("transformation_rule"),
        data.get("lookup_ref"),
        data.get("implemented_by"),
        data.get("regulatory_ref"),
        data.get("status"),
        data.get("notes"),
        id
    ))
    conn.commit()
    conn.close()
    return jsonify({"success": True})

@app.route("/api/mappings/<int:id>", methods=["DELETE"])
def delete_mapping(id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM mappings WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})

# ==========================================================
# PROCEDURE / PACKAGE PARSER & STUDIO APIS
# ==========================================================

@app.route("/api/procedure/parse", methods=["POST"])
def parse_procedure():
    data = request.json or {}
    sql_text = data.get("sql_text", "")
    target_table = data.get("target_table", "")
    system_code = data.get("system_code", "")

    res = procedure_parser.parse_sql_procedure(sql_text, target_table_hint=target_table, system_hint=system_code)
    return jsonify(res)

@app.route("/api/procedure/apply", methods=["POST"])
def apply_procedure_mappings():
    data = request.json or {}
    report_code = data.get("report_code")
    system_code = data.get("system_code", "CIC")
    target_table = data.get("target_table", "")
    fields = data.get("fields", [])
    sql_text = data.get("procedure_code", "")
    proc_name = data.get("procedure_name", "")

    if not report_code or not fields:
        return jsonify({"success": False, "error": "Thiếu mã báo cáo hoặc danh sách chỉ tiêu"}), 400

    conn = get_db()
    cur = conn.cursor()

    applied_count = 0
    for f in fields:
        col_name = f.get("column", "").strip().upper()
        if not col_name:
            continue

        field_name_vi = f.get("field_name_vi", col_name)
        data_type = f.get("data_type", "VARCHAR2(50)")
        tbl = f.get("table") or target_table
        src_tbl = f.get("source_table", "")
        src_col = f.get("source_column", "")
        rule = f.get("transformation_rule", "")
        notes = f.get("notes", "")

        # Check if exists
        row = cur.execute("""
            SELECT id FROM mappings 
            WHERE system_code = ? AND report_code = ? AND field_code = ?
        """, (system_code, report_code, col_name)).fetchone()

        if row:
            cur.execute("""
                UPDATE mappings SET
                    field_name_vi = coalesce(nullif(?, ''), field_name_vi),
                    data_type = coalesce(nullif(?, ''), data_type),
                    target_table = coalesce(nullif(?, ''), target_table),
                    source_table = coalesce(nullif(?, ''), source_table),
                    source_column = coalesce(nullif(?, ''), source_column),
                    transformation_rule = coalesce(nullif(?, ''), transformation_rule),
                    notes = coalesce(nullif(?, ''), notes),
                    implemented_by = coalesce(nullif(?, ''), implemented_by),
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (field_name_vi, data_type, tbl, src_tbl, src_col, rule, notes, proc_name, row[0]))
        else:
            cur.execute("""
                INSERT INTO mappings (
                    system_code, report_code, field_code, field_name_vi, data_type,
                    target_table, target_column, source_system, source_table, source_column,
                    transformation_rule, implemented_by, notes, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, 'ODS', ?, ?, ?, ?, ?, 'Active')
            """, (system_code, report_code, col_name, field_name_vi, data_type, tbl, col_name,
                  src_tbl, src_col, rule, proc_name, notes))
        applied_count += 1

    # Save procedure code to report
    if sql_text:
        cur.execute("""
            UPDATE reports SET 
                procedure_code = ?,
                procedure_name = coalesce(nullif(?, ''), procedure_name)
            WHERE system_code = ? AND report_code = ?
        """, (sql_text, proc_name, system_code, report_code))

    conn.commit()
    conn.close()
    return jsonify({"success": True, "applied_count": applied_count})

@app.route("/api/reports/<report_code>/procedure", methods=["GET"])
def get_report_procedure(report_code):
    system = request.args.get("system", "CIC")
    conn = get_db()
    cur = conn.cursor()
    row = cur.execute("SELECT procedure_name, procedure_code FROM reports WHERE system_code = ? AND report_code = ?", (system, report_code)).fetchone()
    conn.close()
    if not row:
        return jsonify({"procedure_name": "", "procedure_code": ""})
    return jsonify({"procedure_name": row[0] or "", "procedure_code": row[1] or ""})

@app.route("/api/reports/<report_code>/procedure", methods=["POST"])
def update_report_procedure(report_code):
    data = request.json or {}
    system = data.get("system_code", "CIC")
    procedure_name = data.get("procedure_name", "")
    procedure_code = data.get("procedure_code", "")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE reports SET 
            procedure_name = ?,
            procedure_code = ?
        WHERE system_code = ? AND report_code = ?
    """, (procedure_name, procedure_code, system, report_code))
    conn.commit()
    conn.close()
    return jsonify({"success": True})

# ==========================================================
# LINEAGE & EXPORT / IMPORT
# ==========================================================

@app.route("/api/lineage")
def get_lineage():
    source_table = request.args.get("source_table", "").strip().upper()
    conn = get_db()
    cur = conn.cursor()
    
    if source_table:
        rows = cur.execute("""
            SELECT m.source_table, m.source_column, m.system_code, m.report_code, 
                   m.field_code, m.field_name_vi, m.target_table, m.implemented_by
            FROM mappings m
            WHERE UPPER(m.source_table) LIKE ? OR UPPER(m.source_column) LIKE ?
            ORDER BY m.system_code, m.report_code
        """, (f"%{source_table}%", f"%{source_table}%")).fetchall()
    else:
        rows = cur.execute("""
            SELECT m.source_table, m.system_code, m.report_code, COUNT(m.id) as field_count
            FROM mappings m
            WHERE m.source_table IS NOT NULL AND m.source_table != ''
            GROUP BY m.source_table, m.system_code, m.report_code
            ORDER BY field_count DESC
            LIMIT 50
        """).fetchall()
        
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route("/api/sync", methods=["POST"])
def trigger_sync():
    res = parser.run_full_sync()
    return jsonify(res)

@app.route("/api/export/excel")
def export_excel():
    system = request.args.get("system")
    report = request.args.get("report")
    if system == "ALL": system = None
    if report == "ALL": report = None
    
    file_path = exporter.export_to_excel(system, report)
    return send_file(file_path, as_attachment=True, download_name=os.path.basename(file_path))

@app.route("/api/export/word")
def export_word():
    system = request.args.get("system")
    report = request.args.get("report")
    if system == "ALL": system = None
    if report == "ALL": report = None
    
    file_path = exporter.export_to_word(system, report)
    return send_file(file_path, as_attachment=True, download_name=os.path.basename(file_path))

@app.route("/api/import/excel", methods=["POST"])
def import_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    if f.filename == "":
        return jsonify({"error": "Filename is empty"}), 400
        
    filename = secure_filename(f.filename)
    save_path = os.path.join(UPLOAD_FOLDER, filename)
    f.save(save_path)
    
    try:
        wb = openpyxl.load_workbook(save_path, data_only=True)
        ws = wb.active
        conn = get_db()
        cur = conn.cursor()
        
        imported = 0
        # Template columns:
        # Col 1: STT
        # Col 2: Table
        # Col 3: Column
        # Col 4: DataType
        # Col 5: Tên nghiệp vụ (Mô tả)
        # Col 6: Ghi chú điều kiện
        # Col 7: Bảng nguồn
        # Col 8: Cột nguồn
        # Col 9: Công thức tính
        start_row = 4
        for r in range(start_row, ws.max_row + 1):
            tbl = ws.cell(row=r, column=2).value
            field_code = ws.cell(row=r, column=3).value
            data_type = ws.cell(row=r, column=4).value
            field_name_vi = ws.cell(row=r, column=5).value
            notes = ws.cell(row=r, column=6).value
            src_table = ws.cell(row=r, column=7).value
            src_col = ws.cell(row=r, column=8).value
            rule = ws.cell(row=r, column=9).value

            if not field_code:
                continue

            tbl_str = str(tbl or "").strip().upper()
            sys_code = "CIC" if "CIC" in tbl_str else "TT35"
            rpt_code = tbl_str.replace("CIC_", "").replace("RPTB_", "")
                
            cur.execute("""
            INSERT INTO mappings (
                system_code, report_code, field_code, field_name_vi, data_type,
                target_table, target_column, source_system, source_table, source_column,
                transformation_rule, notes, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 'ODS', ?, ?, ?, ?, 'Active')
            """, (
                sys_code,
                rpt_code,
                str(field_code).strip(),
                str(field_name_vi or field_code).strip(),
                str(data_type or "VARCHAR2(50)").strip(),
                tbl_str,
                str(field_code).strip(),
                str(src_table or "").strip(),
                str(src_col or "").strip(),
                str(rule or "").strip(),
                str(notes or "").strip()
            ))
            imported += 1
            
        conn.commit()
        conn.close()
        return jsonify({"success": True, "imported_count": imported})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def open_browser():
    webbrowser.open_new("http://127.0.0.1:5050")

if __name__ == "__main__":
    parser.init_db()
    conn = get_db()
    cnt = conn.cursor().execute("SELECT COUNT(*) FROM mappings").fetchone()[0]
    conn.close()
    if cnt == 0:
        print("Empty database detected, running initial sync...")
        parser.run_full_sync()
        
    print("==================================================================")
    print("  BANKREPORT MAPPING HUB (TT35 & CIC) READY!")
    print("  URL: http://127.0.0.1:5050")
    print("==================================================================")
    
    if len(sys.argv) > 1 and sys.argv[1] == "--open":
        threading.Timer(1.5, open_browser).start()
        
    port = int(os.environ.get("PORT", 5050))
    app.run(host="0.0.0.0", port=port, debug=False)