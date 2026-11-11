"""
Lab 2: Figma-to-Code, Layouts & Responsive Design - Report Generator
Student: Nguyễn Văn Linh
Student ID: 26560037
Repository: https://github.com/linh-nguyen123/Web-development-lab
Output: 26560037_NguyễnVănLinh_Lab02.docx
"""

import sys
import os

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
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
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base Colors
    COLOR_PRIMARY = RGBColor(15, 37, 55)     # #0F2537
    COLOR_ACCENT = RGBColor(56, 189, 248)    # #38BDF8
    COLOR_DARK = RGBColor(30, 41, 59)        # #1E293B
    COLOR_MUTED = RGBColor(100, 116, 139)    # #64748B
    HEX_HEADER_BG = "0F2537"
    HEX_LIGHT_BG = "F8FAFC"
    HEX_ACCENT_BG = "E0F2FE"
    HEX_BORDER = "CBD5E1"
    HEX_PLACEHOLDER_BG = "F1F5F9"

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
    sub_run = sub_p.add_run("Bridging the Gap between Figma Specifications and Modern Responsive CSS Grids")
    sub_run.font.name = "Arial"
    sub_run.font.size = Pt(12)
    sub_run.font.italic = True
    sub_run.font.color.rgb = COLOR_MUTED

    # Student Credentials Table
    cred_table = doc.add_table(rows=5, cols=2)
    cred_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cred_data = [
        ("Student Name:", "Nguyễn Văn Linh"),
        ("Student ID:", "26560037"),
        ("Course / Lab:", "Web Development Laboratory - Lab 02"),
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
            r_val.font.color.rgb = RGBColor(14, 116, 144)
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
        "Without configuring the viewport metadata in the HTML <head>, mobile browsers default to rendering pages "
        "at a virtual desktop resolution (commonly 980px). This causes text and layout structures to shrink dramatically, "
        "rendering the UI illegible on mobile devices. In this lab, the mandatory viewport meta tag is implemented:"
    )

    p_code_vp = doc.add_paragraph()
    p_code_vp.paragraph_format.space_before = Pt(4)
    p_code_vp.paragraph_format.space_after = Pt(8)
    r_code_vp = p_code_vp.add_run('<meta name="viewport" content="width=device-width, initial-scale=1.0">')
    r_code_vp.font.name = "Consolas"
    r_code_vp.font.size = Pt(9.5)
    r_code_vp.font.bold = True
    r_code_vp.font.color.rgb = RGBColor(3, 105, 161)

    p_vp_desc = doc.add_paragraph()
    p_vp_desc.paragraph_format.space_after = Pt(8)
    p_vp_desc.add_run(
        "• width=device-width: Enforces that the layout viewport precisely matches the device's physical screen width in CSS pixels.\n"
        "• initial-scale=1.0: Sets a 1:1 scale ratio between physical device pixels and CSS pixels on initial page load."
    )

    # 1.2 Figma Inspection & CSS Variables
    h2_2 = doc.add_paragraph()
    h2_2.paragraph_format.space_before = Pt(8)
    h2_2.paragraph_format.space_after = Pt(4)
    r2_2 = h2_2.add_run("1.2 Figma Inspection and CSS Design Tokens (:root Variables)")
    r2_2.font.name = "Arial"
    r2_2.font.size = Pt(11)
    r2_2.font.bold = True
    r2_2.font.color.rgb = COLOR_DARK

    p_figma = doc.add_paragraph()
    p_figma.paragraph_format.space_after = Pt(6)
    p_figma.add_run(
        "Following the lab guidelines, exact specifications were extracted from the Figma Dev Mode canvas instead of "
        "approximating pixel measurements. Spacings were verified holding the Alt key, line-heights were converted into "
        "unitless proportional multipliers, and all core design tokens were structured as CSS Custom Properties in :root:"
    )

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
        ("--bg-dark / --bg-surface", "#0a0e17 / #111827", "Deep high-contrast canvas and card surface backgrounds"),
        ("--accent-cyan / --accent-blue", "#38bdf8 / #3b82f6", "Primary Figma brand accent, active link highlights & badges"),
        ("--bg-glass", "rgba(17, 24, 39, 0.82)", "Sticky navigation bar glassmorphism with 14px backdrop blur"),
        ("--shadow-elevation", "0 20px 30px -10px ...", "Card hover elevation and smooth cyan glow box-shadow"),
        ("--transition-smooth", "0.35s cubic-bezier(...)", "Consistent spring motion for card hover, image zoom & menu")
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

    p_arch = doc.add_paragraph()
    p_arch.paragraph_format.space_after = Pt(6)
    p_arch.add_run(
        "A foundational principle of modern CSS architecture is knowing when to employ Flexbox versus Grid. "
        "The differences and their respective implementations in Lab 2 are summarized below:"
    )

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
        ("Dimensional Control", "1-Dimensional: Aligns items along a single axis (row OR column at a time).", "2-Dimensional: Controls rows AND columns simultaneously in a unified grid coordinate plane."),
        ("Design Paradigm", "Content-First: Child elements dictate sizing based on inner content; container distributes spacing.", "Layout-First: Container dictates strict tracks; child elements snap into predefined slots."),
        ("Ideal Use Cases", "Navigation bars, toolbars, badges, button groups, vertical card content stacking.", "Product catalogs, multi-column portfolio grids, image galleries, complex dashboard dashboards."),
        ("Lab 2 Implementation", "Sticky Navbar: display: flex; justify-content: space-between; align-items: center; with glassmorphism.", "Portfolio Gallery: grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)) with media query breakpoints.")
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

    p_check_intro = doc.add_paragraph()
    p_check_intro.paragraph_format.space_after = Pt(8)
    p_check_intro.add_run(
        "The following placeholders are formatted for pasting execution screenshots directly into the report. "
        "Each section verifies a specific functional requirement established in Lab 2."
    )

    placeholders = [
        (
            "[Screenshot Placeholder 1] Sticky Navbar & Hover Effects (CSS Flexbox)",
            "Sticky navigation bar with glassmorphism (backdrop-filter: blur(14px)), brand logo, horizontal desktop links, "
            "and interactive hover transitions. Implemented via display: flex; justify-content: space-between; align-items: center;."
        ),
        (
            "[Screenshot Placeholder 2] Product/Portfolio Grid Layout (CSS Grid 6+ Items)",
            "CSS Grid section showcasing 6 portfolio items: Creative Interface, Dynamic Dashboard, E-Commerce Storefront, "
            "Social Media App, Finance Tracker, and Fitness Companion, with category badges and action buttons."
        ),
        (
            "[Screenshot Placeholder 3] Responsive Mobile Viewport Layout (< 768px - 1 Column)",
            "Mobile simulation (< 768px): Verified 1-column grid layout (grid-template-columns: 1fr), visible hamburger toggle button, "
            "and collapsible navigation drawer functioning via the menuToggle JavaScript event."
        ),
        (
            "[Screenshot Placeholder 4] Responsive Tablet Viewport Layout (768px - 1023px - 2 Columns)",
            "Tablet simulation (768px - 1023px): Verified 2-column grid layout (grid-template-columns: repeat(2, 1fr)), "
            "automatic expansion of horizontal desktop links, and hamburger menu toggle hidden."
        ),
        (
            "[Screenshot Placeholder 5] Responsive Desktop Viewport Layout (>= 1024px - 3 Columns)",
            "Desktop viewport (>= 1024px): Full 3-column expansive grid layout (grid-template-columns: repeat(3, 1fr)), "
            "centered 1240px container, and full typography scale."
        ),
        (
            "[Screenshot Placeholder 6] Interactive Hover States (Image Zoom & Box-Shadow Elevation)",
            "Demonstration of active cursor hover: Image smooth zoom (transform: scale(1.08)), card elevation (translateY(-8px)), "
            "and dynamic cyan shadow glow (box-shadow: 0 20px 30px -10px rgba(56, 189, 248, 0.25))."
        )
    ]

    for title, desc in placeholders:
        box_table = doc.add_table(rows=2, cols=1)
        box_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Title Cell
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

        # Screenshot Area Cell
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

    p_git_intro = doc.add_paragraph()
    p_git_intro.paragraph_format.space_after = Pt(6)
    p_git_intro.add_run(
        "As specified in the functional requirements, the complete project source code was committed with the exact "
        "message 'lab02' and deployed live to the public internet using GitHub Pages."
    )

    git_table = doc.add_table(rows=4, cols=2)
    git_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    git_specs = [
        ("Target Repository:", "https://github.com/linh-nguyen123/Web-development-lab"),
        ("Required Commit Message:", "lab02"),
        ("Target Branch:", "main (or master)"),
        ("Live GitHub Pages URL:", "https://linh-nguyen123.github.io/Web-development-lab/")
    ]
    for idx, (label, val) in enumerate(git_specs):
        cell_lbl, cell_val = git_table.rows[idx].cells[0], git_table.rows[idx].cells[1]
        cell_lbl.width = Inches(2.2)
        cell_val.width = Inches(4.3)
        set_cell_background(cell_lbl, "F1F5F9")
        set_cell_background(cell_val, "FFFFFF")
        set_cell_margins(cell_lbl, top=80, bottom=80, left=140, right=140)
        set_cell_margins(cell_val, top=80, bottom=80, left=140, right=140)

        p1 = cell_lbl.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(label)
        r1.font.name = "Arial"
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_DARK

        p2 = cell_val.paragraphs[0]
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(val)
        r2.font.name = "Consolas" if "Commit" in label or "Branch" in label else "Arial"
        r2.font.size = Pt(9.5)
        if "URL" in label or "Repository" in label:
            r2.font.color.rgb = RGBColor(2, 132, 199)
            r2.font.underline = True
        else:
            r2.font.bold = True
            r2.font.color.rgb = COLOR_DARK

    p_cmd_title = doc.add_paragraph()
    p_cmd_title.paragraph_format.space_before = Pt(10)
    p_cmd_title.paragraph_format.space_after = Pt(4)
    r_ct = p_cmd_title.add_run("Terminal Execution Commands for Submission:")
    r_ct.font.name = "Arial"
    r_ct.font.size = Pt(10.5)
    r_ct.font.bold = True
    r_ct.font.color.rgb = COLOR_PRIMARY

    p_cmd = doc.add_paragraph()
    p_cmd.paragraph_format.space_after = Pt(10)
    cmd_text = (
        "# 1. Initialize and verify workspace files\n"
        "git add index.html styles.css assets/ 26560037_NguyễnVănLinh_Lab02.docx\n\n"
        "# 2. Commit with mandatory message\n"
        'git commit -m "lab02"\n\n'
        "# 3. Set branch to main and link remote\n"
        "git branch -M main\n"
        "git remote add origin https://github.com/linh-nguyen123/Web-development-lab.git\n\n"
        "# 4. Push changes to GitHub repository\n"
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
    r_ght = p_ght.add_run("[Screenshot Placeholder 7] GitHub Repository & Live GitHub Pages Proof")
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
