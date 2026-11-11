"""
Lab 2: Figma-to-Code, Layouts & Responsive Design - Report Generator
Student: Nguyễn Văn Linh
Student ID: 26560037
Project: VibeTech E-Commerce Workspace Gear Landing Page (Homework Assignment)
Repository: https://github.com/linh-nguyen123/Web-development-lab
Output: 26560037_NguyễnVănLinh_Lab02.docx
"""

import sys
import os

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
except ImportError:
    print("python-docx is not installed. Please install it using: pip install python-docx")
    sys.exit(1)


def set_cell_background(cell, fill_hex):
    """Set the background color of a table cell."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    """Set internal cell padding (in dxa: 20 dxa = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def create_report():
    doc = Document()

    # 1 inch margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    COLOR_PRIMARY = RGBColor(30, 27, 75)     # #1E1B4B
    COLOR_ACCENT = RGBColor(79, 70, 229)     # #4F46E5
    COLOR_DARK = RGBColor(15, 23, 42)        # #0F172A
    COLOR_MUTED = RGBColor(100, 116, 139)    # #64748B
    HEX_ACCENT_BG = "EEF2FF"
    HEX_PLACEHOLDER_BG = "F8FAFC"

    # =========================================================================
    # TIÊU ĐỀ BÁO CÁO
    # =========================================================================
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    title_run = title_p.add_run("BÁO CÁO THỰC HÀNH LAB 2")
    title_run.font.name = "Arial"
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.color.rgb = COLOR_PRIMARY
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(16)
    sub_run = sub_p.add_run("FIGMA-TO-CODE, LAYOUTS & RESPONSIVE DESIGN")
    sub_run.font.name = "Arial"
    sub_run.font.size = Pt(13)
    sub_run.font.bold = True
    sub_run.font.color.rgb = COLOR_ACCENT
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Bảng thông tin sinh viên & Link GitHub
    cred_table = doc.add_table(rows=7, cols=2)
    cred_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cred_data = [
        ("Họ và tên sinh viên:", "Nguyễn Văn Linh"),
        ("Mã số sinh viên (MSSV):", "26560037"),
        ("Lớp / Học phần:", "Phát triển Ứng dụng Web - Lab 02"),
        ("Đề tài thực hành:", "VibeTech E-Commerce Workspace Gear (Figma to Code)"),
        ("Git Commit Message:", "lab02"),
        ("GitHub Repository URL:", "https://github.com/linh-nguyen123/Web-development-lab"),
        ("GitHub Pages Live Demo:", "https://linh-nguyen123.github.io/Web-development-lab/")
    ]

    for idx, (label, val) in enumerate(cred_data):
        row = cred_table.rows[idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        cell_lbl.width = Inches(2.3)
        cell_val.width = Inches(4.2)

        set_cell_background(cell_lbl, "F1F5F9")
        set_cell_background(cell_val, "FFFFFF")
        set_cell_margins(cell_lbl, top=90, bottom=90, left=140, right=140)
        set_cell_margins(cell_val, top=90, bottom=90, left=140, right=140)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.paragraph_format.space_after = Pt(0)
        r_lbl = p_lbl.add_run(label)
        r_lbl.font.name = "Arial"
        r_lbl.font.size = Pt(9.5)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = COLOR_DARK

        p_val = cell_val.paragraphs[0]
        p_val.paragraph_format.space_after = Pt(0)
        r_val = p_val.add_run(val)
        r_val.font.name = "Arial"
        r_val.font.size = Pt(9.5)

        if "Commit" in label:
            r_val.font.name = "Consolas"
            r_val.font.bold = True
            r_val.font.color.rgb = RGBColor(15, 118, 110)
        elif "URL" in label or "Demo" in label:
            r_val.font.name = "Consolas"
            r_val.font.color.rgb = RGBColor(2, 132, 199)
            r_val.font.underline = True
        elif "tên" in label or "MSSV" in label:
            r_val.font.bold = True
            r_val.font.color.rgb = COLOR_DARK
        else:
            r_val.font.color.rgb = COLOR_DARK

    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # =========================================================================
    # DANH MỤC HÌNH ẢNH MINH CHỨNG
    # =========================================================================
    h_img = doc.add_paragraph()
    h_img.paragraph_format.space_before = Pt(8)
    h_img.paragraph_format.space_after = Pt(10)
    r_img = h_img.add_run("HÌNH ẢNH MINH CHỨNG KẾT QUẢ THỰC HÀNH")
    r_img.font.name = "Arial"
    r_img.font.size = Pt(13)
    r_img.font.bold = True
    r_img.font.color.rgb = COLOR_PRIMARY

    placeholders = [
        (
            "Hình 1. Thanh điều hướng Sticky Navbar & Hiệu ứng Glassmorphism (Flexbox)",
            "Chụp giao diện thanh menu cố định phía trên trang web, căn chỉnh Flexbox với Logo VibeTech, menu liên kết và nút giỏ hàng."
        ),
        (
            "Hình 2. Lưới sản phẩm trên màn hình Desktop (>= 1024px - CSS Grid 4 cột)",
            "Chụp toàn bộ section lưới sản phẩm hoàn chỉnh hiển thị 4 cột theo đúng thiết kế Figma 1920w & 1024w."
        ),
        (
            "Hình 3. Giao diện Responsive trên Điện thoại Di động (< 768px - Lưới 1 cột)",
            "Chụp chế độ giả lập mobile trên trình duyệt DevTools: hiển thị 1 cột duy nhất và menu rút gọn dạng Hamburger."
        ),
        (
            "Hình 4. Giao diện Responsive trên Máy tính bảng Tablet (768px - 1023px - Lưới 2 cột)",
            "Chụp chế độ giả lập tablet (iPad) trên DevTools: hiển thị bố cục lưới 2 cột theo đúng Figma 768w light."
        ),
        (
            "Hình 5. Hiệu ứng tương tác Hover (Phóng to ảnh & Nổi khối thẻ)",
            "Chụp khi rê chuột vào thẻ sản phẩm để thấy ảnh phóng to nhẹ và thẻ nâng lên kèm bóng đổ."
        ),
        (
            "Hình 6. Minh chứng Commit Git (lab02) & Triển khai GitHub Pages thành công",
            "Chụp lịch sử commit git (git log -1 với commit 'lab02') hoặc trang web chạy trực tiếp trên GitHub Pages."
        )
    ]

    for title, desc in placeholders:
        box_table = doc.add_table(rows=2, cols=1)
        box_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        c_title = box_table.rows[0].cells[0]
        c_title.width = Inches(6.5)
        set_cell_background(c_title, HEX_ACCENT_BG)
        set_cell_margins(c_title, top=100, bottom=100, left=140, right=140)
        p_t = c_title.paragraphs[0]
        p_t.paragraph_format.space_after = Pt(2)
        r_t = p_t.add_run(title)
        r_t.font.name = "Arial"
        r_t.font.size = Pt(10.5)
        r_t.font.bold = True
        r_t.font.color.rgb = COLOR_PRIMARY

        p_d = c_title.add_paragraph()
        p_d.paragraph_format.space_after = Pt(0)
        r_d = p_d.add_run(desc)
        r_d.font.name = "Arial"
        r_d.font.size = Pt(8.5)
        r_d.font.color.rgb = COLOR_MUTED

        c_box = box_table.rows[1].cells[0]
        c_box.width = Inches(6.5)
        set_cell_background(c_box, HEX_PLACEHOLDER_BG)
        set_cell_margins(c_box, top=500, bottom=500, left=140, right=140)
        p_b = c_box.paragraphs[0]
        p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_b.paragraph_format.space_after = Pt(0)
        r_b = p_b.add_run("[ DÁN HÌNH ẢNH SCREENSHOT VÀO ĐÂY ]")
        r_b.font.name = "Consolas"
        r_b.font.size = Pt(10)
        r_b.font.bold = True
        r_b.font.color.rgb = COLOR_MUTED

        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    output_path = "26560037_NguyễnVănLinh_Lab02.docx"
    doc.save(output_path)
    print(f"Report successfully generated at: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    create_report()
