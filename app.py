# -*- coding: utf-8 -*-
"""
BankReport Mapping Hub - Web Application Backend (Flask + SQLite)
Manages Business-to-Technical Mappings for TT35 and CIC systems.
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
               r.procedure_name, r.regulation_ref, r.description,
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
        start_row = 4
        for r in range(start_row, ws.max_row + 1):
            sys_val = ws.cell(row=r, column=2).value
            rpt_val = ws.cell(row=r, column=3).value
            field_code = ws.cell(row=r, column=5).value
            field_name_vi = ws.cell(row=r, column=6).value
            
            if not field_code:
                continue
                
            cur.execute("""
            INSERT INTO mappings (
                system_code, report_code, field_code, field_name_vi, data_type,
                target_table, source_table, source_column, transformation_rule,
                lookup_ref, implemented_by, regulatory_ref, notes, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Active')
            """, (
                str(sys_val or "CIC"),
                str(rpt_val or ""),
                str(field_code),
                str(field_name_vi or field_code),
                str(ws.cell(row=r, column=7).value or "VARCHAR2(50)"),
                str(ws.cell(row=r, column=8).value or ""),
                str(ws.cell(row=r, column=9).value or ""),
                str(ws.cell(row=r, column=10).value or ""),
                str(ws.cell(row=r, column=11).value or ""),
                str(ws.cell(row=r, column=12).value or ""),
                str(ws.cell(row=r, column=13).value or ""),
                str(ws.cell(row=r, column=14).value or ""),
                str(ws.cell(row=r, column=15).value or "")
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
        
    port = int(os.environ.get("PORT", 5050))`r`n    app.run(host="0.0.0.0", port=port, debug=False)