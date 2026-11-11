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

    # Set page margins (1 inch)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base Colors
    COLOR_PRIMARY = RGBColor(30, 27, 75)     # #1E1B4B
    COLOR_ACCENT = RGBColor(79, 70, 229)     # #4F46E5
    COLOR_DARK = RGBColor(15, 23, 42)        # #0F172A
    COLOR_MUTED = RGBColor(100, 116, 139)    # #64748B
    HEX_HEADER_BG = "4F46E5"
    HEX_LIGHT_BG = "F8FAFC"
    HEX_ACCENT_BG = "EEF2FF"
    HEX_PLACEHOLDER_BG = "F8FAFC"

    # =========================================================================
    # DOCUMENT COVER HEADER
    # =========================================================================
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    title_run = title_p.add_run("REPORT: LAB 2 - FIGMA-TO-CODE, LAYOUTS & RESPONSIVE DESIGN")
    title_run.font.name = "Arial"
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.color.rgb = COLOR_PRIMARY

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(16)
    sub_run = sub_p.add_run("Homework Assignment: VibeTech Responsive E-Commerce Workspace Gear Landing Page")
    sub_run.font.name = "Arial"
    sub_run.font.size = Pt(12)
    sub_run.font.italic = True
    sub_run.font.color.rgb = COLOR_ACCENT

    # Student Credentials Table
    cred_table = doc.add_table(rows=6, cols=2)
    cred_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cred_data = [
        ("Student Name:", "Nguyễn Văn Linh"),
        ("Student ID:", "26560037"),
        ("Course / Lab:", "Web Development Laboratory - Lab Session 02"),
        ("Figma Project:", "VibeTech E-Commerce Workspace Gear (Unleash the Next Level)"),
        ("Git Commit Message:", "lab02"),
        ("Target Repository:", "https://github.com/linh-nguyen123/Web-development-lab")
    ]

    for idx, (label, val) in enumerate(cred_data):
        row = cred_table.rows[idx]
        cell_lbl, cell_val = row.cells[0], row.cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(4.3)

        set_cell_background(cell_lbl, "F1F5F9")
        set_cell_background(cell_val, "FFFFFF")
        set_cell_margins(cell_lbl, top=100, bottom=100, left=140, right=140)
        set_cell_margins(cell_val, top=100, bottom=100, left=140, right=140)

        p_lbl = cell_lbl.paragraphs[0]
        p_lbl.paragraph_format.space_after = Pt(0)
        r_lbl = p_lbl.add_run(label)
        r_lbl.font.name = "Arial"
        r_lbl.font.size = Pt(10)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = COLOR_DARK

        p_val = cell_val.paragraphs[0]
        p_val.paragraph_format.space_after = Pt(0)
        r_val = p_val.add_run(val)
        r_val.font.name = "Arial"
        r_val.font.size = Pt(10)
        if label.startswith("Git"):
            r_val.font.name = "Consolas"
            r_val.font.bold = True
            r_val.font.color.rgb = RGBColor(15, 118, 110)
        elif label.startswith("Target"):
            r_val.font.color.rgb = RGBColor(2, 132, 199)
            r_val.font.underline = True
        else:
            r_val.font.color.rgb = COLOR_DARK

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # =========================================================================
    # SECTION 1: OVERVIEW & ARCHITECTURE
    # =========================================================================
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(6)
    r1 = h1.add_run("1. Overview & Architectural Analysis")
    r1.font.name = "Arial"
    r1.font.size = Pt(14)
    r1.font.bold = True
    r1.font.color.rgb = COLOR_PRIMARY

    # 1.1 Viewport Meta Tag
    h2_1 = doc.add_paragraph()
    h2_1.paragraph_format.space_before = Pt(8)
    h2_1.paragraph_format.space_after = Pt(4)
    r2_1 = h2_1.add_run("1.1 The Responsive Viewport Meta Tag")
    r2_1.font.name = "Arial"
    r2_1.font.size = Pt(11)
    r2_1.font.bold = True
    r2_1.font.color.rgb = COLOR_DARK

    p_vp = doc.add_paragraph()
    p_vp.paragraph_format.space_after = Pt(6)
    p_vp.add_run(
        "Without configuring the viewport metadata in the HTML <head>, mobile browsers render pages at an artificial "
        "virtual desktop width (typically 980px). This makes font sizes tiny and unusable on mobile devices. "
        "In this homework, the mandatory responsive viewport tag is implemented:"
    )

    p_code_vp = doc.add_paragraph()
    p_code_vp.paragraph_format.space_before = Pt(4)
    p_code_vp.paragraph_format.space_after = Pt(8)
    r_code_vp = p_code_vp.add_run('<meta name="viewport" content="width=device-width, initial-scale=1.0">')
    r_code_vp.font.name = "Consolas"
    r_code_vp.font.size = Pt(9.5)
    r_code_vp.font.bold = True
    r_code_vp.font.color.rgb = COLOR_ACCENT

    # 1.2 Figma Inspection & CSS Variables
    h2_2 = doc.add_paragraph()
    h2_2.paragraph_format.space_before = Pt(8)
    h2_2.paragraph_format.space_after = Pt(4)
    r2_2 = h2_2.add_run("1.2 Figma Dev Mode Inspection and CSS Design Tokens (:root Variables)")
    r2_2.font.name = "Arial"
    r2_2.font.size = Pt(11)
    r2_2.font.bold = True
    r2_2.font.color.rgb = COLOR_DARK

    var_table = doc.add_table(rows=6, cols=3)
    var_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["CSS Token Variable", "Token Value", "Functional UI Purpose"]
    for i, h in enumerate(headers):
        cell = var_table.rows[0].cells[i]
        set_cell_background(cell, HEX_HEADER_BG)
        set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.font.name = "Arial"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    tokens = [
        ("--bg-page", "#f8fafc", "Clean neutral light page background matching Figma frame"),
        ("--brand-primary", "#4f46e5", "VibeTech brand indigo for logo, cart badge, active links"),
        ("--brand-dark", "#0f172a", "Primary button background and Hot/Sale card badge background"),
        ("--bg-glass", "rgba(255, 255, 255, 0.88)", "Sticky navbar glassmorphism background with 12px blur"),
        ("--shadow-card-hover", "0 16px 28px -6px ...", "Smooth elevation box shadow on card hover")
    ]
    for row_idx, row_data in enumerate(tokens, start=1):
        for col_idx, text in enumerate(row_data):
            cell = var_table.rows[row_idx].cells[col_idx]
            set_cell_background(cell, HEX_LIGHT_BG if row_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.name = "Arial" if col_idx != 0 else "Consolas"
            r.font.size = Pt(9)
            r.font.color.rgb = COLOR_DARK

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 1.3 CSS Flexbox vs. CSS Grid Analysis
    h2_3 = doc.add_paragraph()
    h2_3.paragraph_format.space_before = Pt(8)
    h2_3.paragraph_format.space_after = Pt(4)
    r2_3 = h2_3.add_run("1.3 Architectural Comparison: CSS Flexbox vs. CSS Grid")
    r2_3.font.name = "Arial"
    r2_3.font.size = Pt(11)
    r2_3.font.bold = True
    r2_3.font.color.rgb = COLOR_PRIMARY

    comp_table = doc.add_table(rows=5, cols=3)
    comp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_headers = ["Architectural Dimension", "CSS Flexbox (1-Dimensional)", "CSS Grid (2-Dimensional)"]
    for i, h in enumerate(c_headers):
        cell = comp_table.rows[0].cells[i]
        set_cell_background(cell, HEX_HEADER_BG)
        set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.font.name = "Arial"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    comp_data = [
        ("Dimensional Control", "1-Dimensional: Aligns items along one axis at a time (row OR column).", "2-Dimensional: Simultaneously coordinates items across rows AND columns in a unified layout plane."),
        ("Design Paradigm", "Content-First: Items determine their own dimensions based on internal content; container aligns them.", "Layout-First: Container defines explicit grid track structures; items snap within assigned slots."),
        ("Ideal UI Targets", "Sticky navigation bars, toolbar actions, search/cart buttons, mobile drawer menus.", "Product catalogs, multi-column portfolio grids, image galleries, complex web apps."),
        ("Homework Implementation", "Navbar: display: flex; justify-content: space-between; align-items: center; with glassmorphism.", "Product Grid: grid-template-columns across 4 columns (desktop), 2 columns (tablet), 1 column (mobile).")
    ]
    for row_idx, row_data in enumerate(comp_data, start=1):
        for col_idx, text in enumerate(row_data):
            cell = comp_table.rows[row_idx].cells[col_idx]
            set_cell_background(cell, HEX_LIGHT_BG if row_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.name = "Arial"
            r.font.size = Pt(9)
            if col_idx == 0:
                r.font.bold = True
            r.font.color.rgb = COLOR_DARK

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # =========================================================================
    # SECTION 2: UI & FEATURE SCREENSHOT CHECKLIST (PLACEHOLDERS)
    # =========================================================================
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(14)
    h2.paragraph_format.space_after = Pt(6)
    r2 = h2.add_run("2. UI & Feature Implementation Screenshot Checklist")
    r2.font.name = "Arial"
    r2.font.size = Pt(14)
    r2.font.bold = True
    r2.font.color.rgb = COLOR_PRIMARY

    placeholders = [
        (
            "[Screenshot Placeholder 1] Sticky Navbar & Hover Effects (CSS Flexbox)",
            "Sticky header with VibeTech logo, centered navigation links (Home, Products, Deals, About), search & cart icons, and glassmorphism (backdrop-filter: blur(12px))."
        ),
        (
            "[Screenshot Placeholder 2] Product/Portfolio Grid Layout (CSS Grid 8 Items)",
            "Full product grid showing 8 items: Wireless Mouse, Tactix K85 Mechanical, ZenSound ANC Headset, NovaFit Smart Watch, Horizon 27 IPS Display, StreamMic Pro, Lumos Desk Lamp, and ErgoStand Riser."
        ),
        (
            "[Screenshot Placeholder 3] Responsive Mobile Viewport Layout (< 768px - 1 Column)",
            "Mobile simulation (< 768px): 1-column card stack, visible hamburger button, and active mobile navigation drawer."
        ),
        (
            "[Screenshot Placeholder 4] Responsive Tablet Viewport Layout (768px - 1023px - 2 Columns, Figma 768w light)",
            "Tablet simulation (768px - 1023px): 2-column product grid, hidden hamburger button, and horizontal desktop navbar links active."
        ),
        (
            "[Screenshot Placeholder 5] Responsive Desktop Viewport Layout (>= 1024px - 4 Columns, Figma 1024w & 1920w light)",
            "Desktop viewport (>= 1024px): 4-column product grid layout, centered 1280px container, and full desktop navbar."
        ),
        (
            "[Screenshot Placeholder 6] Interactive Hover States (Image Zoom & Elevation)",
            "Active cursor hover showing card image scale (transform: scale(1.08)), card elevation (translateY(-6px)), and glowing box-shadow."
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
        set_cell_margins(c_box, top=450, bottom=450, left=140, right=140)
        p_b = c_box.paragraphs[0]
        p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_b.paragraph_format.space_after = Pt(0)
        r_b = p_b.add_run("--- [PASTE SCREENSHOT HERE] ---")
        r_b.font.name = "Consolas"
        r_b.font.size = Pt(10)
        r_b.font.bold = True
        r_b.font.color.rgb = COLOR_MUTED

        doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # =========================================================================
    # SECTION 3: GIT COMMIT & GITHUB PAGES DEPLOYMENT
    # =========================================================================
    h3 = doc.add_paragraph()
    h3.paragraph_format.space_before = Pt(14)
    h3.paragraph_format.space_after = Pt(6)
    r3 = h3.add_run("3. Git Deployment & GitHub Pages Verification")
    r3.font.name = "Arial"
    r3.font.size = Pt(14)
    r3.font.bold = True
    r3.font.color.rgb = COLOR_PRIMARY

    p_cmd = doc.add_paragraph()
    p_cmd.paragraph_format.space_after = Pt(10)
    cmd_text = (
        "# Stage all homework files\n"
        "git add index.html styles.css assets/ 26560037_NguyễnVănLinh_Lab02.docx generate_report.py README.md\n\n"
        "# Commit with mandatory lab02 message\n"
        'git commit -m "lab02"\n\n'
        "# Set branch to main\n"
        "git branch -M main\n\n"
        "# Push changes to GitHub repository\n"
        "git push -u origin main"
    )
    r_cmd = p_cmd.add_run(cmd_text)
    r_cmd.font.name = "Consolas"
    r_cmd.font.size = Pt(8.5)
    r_cmd.font.color.rgb = RGBColor(15, 23, 42)

    # Placeholder for live GitHub Pages screenshot
    box_gh = doc.add_table(rows=2, cols=1)
    box_gh.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_gh_t = box_gh.rows[0].cells[0]
    c_gh_t.width = Inches(6.5)
    set_cell_background(c_gh_t, HEX_ACCENT_BG)
    set_cell_margins(c_gh_t, top=100, bottom=100, left=140, right=140)
    p_ght = c_gh_t.paragraphs[0]
    p_ght.paragraph_format.space_after = Pt(0)
    r_ght = p_ght.add_run("[Screenshot Placeholder 7] GitHub Pages Deployment Status & Live Page Proof")
    r_ght.font.name = "Arial"
    r_ght.font.size = Pt(10.5)
    r_ght.font.bold = True
    r_ght.font.color.rgb = COLOR_PRIMARY

    c_gh_b = box_gh.rows[1].cells[0]
    c_gh_b.width = Inches(6.5)
    set_cell_background(c_gh_b, HEX_PLACEHOLDER_BG)
    set_cell_margins(c_gh_b, top=450, bottom=450, left=140, right=140)
    p_ghb = c_gh_b.paragraphs[0]
    p_ghb.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ghb.paragraph_format.space_after = Pt(0)
    r_ghb = p_ghb.add_run("--- [PASTE GITHUB PAGES DEPLOYMENT / REPO SCREENSHOT HERE] ---")
    r_ghb.font.name = "Consolas"
    r_ghb.font.size = Pt(10)
    r_ghb.font.bold = True
    r_ghb.font.color.rgb = COLOR_MUTED

    output_path = "26560037_NguyễnVănLinh_Lab02.docx"
    doc.save(output_path)
    print(f"Report successfully generated at: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    create_report()
