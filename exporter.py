# -*- coding: utf-8 -*-
"""
Export Engine for BankReport Mapping Hub.
Exports mappings to:
  - Excel Spec (.xlsx) matching exact template:
    STT | Table | Column | DataType | Tên nghiệp vụ (Mô tả) | Ghi chú điều kiện | Bảng nguồn | Cột nguồn | Công thức tính
  - Word Document (.docx)
"""

import os
import sqlite3
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

DB_PATH = os.path.join(os.path.dirname(__file__), "mapping_hub.db")

def export_to_excel(system_code=None, report_code=None, output_path=None):
    if not output_path:
        prefix = f"{system_code}_{report_code}" if report_code else (system_code or "ALL")
        output_path = os.path.join(os.path.dirname(__file__), f"Export_Mapping_{prefix}.xlsx")

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    query = """
        SELECT m.target_table, m.field_code, m.data_type, m.field_name_vi,
               m.notes, m.source_table, m.source_column, m.transformation_rule
        FROM mappings m
        WHERE 1=1
    """
    params = []
    if system_code and system_code != "ALL":
        query += " AND m.system_code = ?"
        params.append(system_code)
    if report_code and report_code != "ALL":
        query += " AND m.report_code = ?"
        params.append(report_code)
    query += " ORDER BY m.system_code, m.report_code, m.id"

    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mapping Spec"

    # Header styling
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    mono_font = Font(name="Consolas", size=10)
    thin_border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB")
    )
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 9 standard columns matching user template
    headers = [
        "STT", "Table", "Column", "DataType", "Tên nghiệp vụ (Mô tả)",
        "Ghi chú điều kiện", "Bảng nguồn", "Cột nguồn", "Công thức tính"
    ]

    # Title row
    ws.merge_cells("A1:I1")
    title_cell = ws["A1"]
    title_cell.value = f"BẢNG MA TRẬN MAPPING CHỈ TIÊU BÁO CÁO ({system_code or 'TT35 & CIC'})"
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Header row
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center
        cell.border = thin_border
    ws.row_dimensions[3].height = 26

    # Data rows
    for r_idx, r_data in enumerate(rows, 1):
        row_num = r_idx + 3
        ws.row_dimensions[row_num].height = 20

        # STT
        c_stt = ws.cell(row=row_num, column=1, value=r_idx)
        c_stt.alignment = align_center
        c_stt.font = data_font
        c_stt.border = thin_border

        for c_idx, val in enumerate(r_data, 2):
            cell = ws.cell(row=row_num, column=c_idx, value=val or "")
            cell.border = thin_border
            if c_idx in (2, 3, 4, 7, 8):
                cell.alignment = align_center if c_idx == 4 else align_left
                cell.font = mono_font if c_idx in (2, 3, 4, 7, 8) else data_font
            else:
                cell.alignment = align_left
                cell.font = data_font

    # Auto column width
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row == 1:
                continue
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 45)

    ws.freeze_panes = "A4"
    wb.save(output_path)
    return output_path

def export_to_word(system_code=None, report_code=None, output_path=None):
    if not output_path:
        prefix = f"{system_code}_{report_code}" if report_code else (system_code or "ALL")
        output_path = os.path.join(os.path.dirname(__file__), f"Doc_Spec_{prefix}.docx")

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    query = """
        SELECT m.system_code, m.report_code, r.report_name, m.target_table, m.field_code, 
               m.data_type, m.field_name_vi, m.notes, m.source_table, m.source_column, 
               m.transformation_rule, m.regulatory_ref
        FROM mappings m
        LEFT JOIN reports r ON m.report_code = r.report_code AND m.system_code = r.system_code
        WHERE 1=1
    """
    params = []
    if system_code and system_code != "ALL":
        query += " AND m.system_code = ?"
        params.append(system_code)
    if report_code and report_code != "ALL":
        query += " AND m.report_code = ?"
        params.append(report_code)
    query += " ORDER BY m.system_code, m.report_code, m.id"

    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()

    doc = docx.Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("TÀI LIỆU ĐẶC TẢ ÁNH XẠ NGHIỆP VỤ SANG KỸ THUẬT")
    run_title.font.name = "Times New Roman"
    run_title.font.size = Pt(16)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(30, 58, 138)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run(f"Hệ thống: {system_code or 'TT35 & CIC'} | Báo cáo: {report_code or 'Toàn bộ danh mục'}")
    run_sub.font.name = "Times New Roman"
    run_sub.font.size = Pt(12)
    run_sub.font.italic = True

    doc.add_paragraph()

    reports_dict = {}
    for r in rows:
        rpt = r[1]
        if rpt not in reports_dict:
            reports_dict[rpt] = {
                "name": r[2] or rpt,
                "system": r[0],
                "items": []
            }
        reports_dict[rpt]["items"].append(r)

    for rpt_code, info in reports_dict.items():
        h2 = doc.add_heading(level=2)
        r_h2 = h2.add_run(f"Báo cáo [{info['system']}] {rpt_code} - {info['name']}")
        r_h2.font.name = "Times New Roman"
        r_h2.font.bold = True
        r_h2.font.color.rgb = RGBColor(15, 23, 42)

        tbl = doc.add_table(rows=1, cols=7)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        headers = ["STT", "Table", "Column", "DataType", "Tên Nghiệp vụ (Mô tả)", "Nguồn (Bảng.Cột)", "Công thức tính"]
        widths = [Inches(0.4), Inches(1.1), Inches(1.0), Inches(0.9), Inches(1.7), Inches(1.4), Inches(1.5)]

        hdr_cells = tbl.rows[0].cells
        for idx, text in enumerate(headers):
            hdr_cells[idx].text = text
            hdr_cells[idx].width = widths[idx]
            run = hdr_cells[idx].paragraphs[0].runs[0]
            run.font.name = "Times New Roman"
            run.font.bold = True
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(255, 255, 255)
            shading = docx.oxml.parse_xml(r'<w:shd {} w:fill="1E3A8A"/>'.format(docx.oxml.ns.nsdecls('w')))
            hdr_cells[idx]._tc.get_or_add_tcPr().append(shading)

        for stt, item in enumerate(info['items'], 1):
            row_cells = tbl.add_row().cells
            src_str = f"{item[8]}.{item[9]}" if item[8] and item[9] else (item[8] or item[9] or "-")
            row_values = [
                str(stt),
                item[3] or "-",
                item[4] or "-",
                item[5] or "-",
                item[6] or "-",
                src_str,
                item[10] or item[7] or "-"
            ]
            for idx, val in enumerate(row_values):
                row_cells[idx].text = val
                row_cells[idx].width = widths[idx]
                if row_cells[idx].paragraphs[0].runs:
                    r = row_cells[idx].paragraphs[0].runs[0]
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(8.5)

        doc.add_paragraph()

    doc.save(output_path)
    return output_path