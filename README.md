# Lab 2: Figma-to-Code, Layouts & Responsive Design

**Course:** Web Development Laboratory  
**Student Name:** Nguyễn Văn Linh  
**Student ID:** 26560037  
**Submission Commit:** `lab02`  
**GitHub Repository:** [https://github.com/linh-nguyen123/Web-development-lab](https://github.com/linh-nguyen123/Web-development-lab)  
**Live GitHub Pages:** [https://linh-nguyen123.github.io/Web-development-lab/](https://linh-nguyen123.github.io/Web-development-lab/)  
**Word Report:** `26560037_NguyễnVănLinh_Lab02.docx`

---

## 📌 Project Overview
This laboratory project systematically translates user interface designs created in Figma into semantic HTML5 code styled with advanced CSS layouts. It demonstrates:
1. **Figma Dev Mode Inspection:** Translating visual tokens (colors, font hierarchy, unitless line-height, spacings via `Alt` key) into CSS variables (`:root`).
2. **Mandatory Viewport Meta Tag:** `<meta name="viewport" content="width=device-width, initial-scale=1.0">` to ensure accurate 1:1 CSS pixel scaling on mobile devices.
3. **1-Dimensional CSS Flexbox Navigation:** Sticky navbar (`position: sticky; top: 0;`) with modern glassmorphism (`backdrop-filter: blur(14px)`), smooth hover indicator transitions, and an interactive mobile drawer menu.
4. **2-Dimensional CSS Grid Portfolio Section:** 6-item portfolio showcase (`Creative Interface`, `Dynamic Dashboard`, `E-Commerce Storefront`, `Social Media App`, `Finance Tracker`, and `Fitness Companion`).
5. **Mobile-First Responsive Breakpoint Architecture:**
   - **Mobile Viewport (`< 768px`):** 1-column grid stack, compact padding, hamburger navigation drawer.
   - **Tablet Viewport (`768px - 1023px`):** 2-column grid (`repeat(2, 1fr)`), horizontal desktop links visible, hamburger menu toggle hidden.
   - **Desktop Viewport (`>= 1024px`):** 3-column expansive grid (`repeat(3, 1fr)`), max-width 1240px container.
6. **Interactive Hover States:** Smooth image zoom (`transform: scale(1.08)`), card elevation (`transform: translateY(-8px)`), and glowing box-shadow effects.

---

## 📂 Project Structure
```text
├── index.html                           # Semantic HTML5 page with viewport tag and grid structure
├── styles.css                           # Mobile-First CSS stylesheet with :root variables & media queries
├── assets/                              # Exported SVG vector assets and logos
│   ├── logo.svg                         # Main brand logo
│   ├── project1.svg                     # Creative Interface project illustration
│   ├── project2.svg                     # Dynamic Dashboard project illustration
│   ├── project3.svg                     # E-Commerce Storefront project illustration
│   ├── project4.svg                     # Social Media App project illustration
│   ├── project5.svg                     # Finance Tracker project illustration
│   └── project6.svg                     # Fitness Companion project illustration
├── 26560037_NguyễnVănLinh_Lab02.docx    # Official academic Word report document
├── generate_report.py                   # Python report generator script (python-docx)
├── build_docx.ps1                       # Native PowerShell OpenXML document builder
├── .gitignore                           # Git ignore rules
└── README.md                            # Project documentation
```

---

## 🚀 Git Submission Commands
Execute the following commands in the project root:

```bash
# 1. Stage all project files
git add index.html styles.css assets/ 26560037_NguyễnVănLinh_Lab02.docx generate_report.py README.md .gitignore

# 2. Commit with the mandatory lab message
git commit -m "lab02"

# 3. Ensure branch is main
git branch -M main

# 4. Link the target remote repository (if not already linked)
git remote add origin https://github.com/linh-nguyen123/Web-development-lab.git

# 5. Push to GitHub
git push -u origin main
```

---

## 🌐 GitHub Pages Deployment
1. Navigate to your repository on GitHub: `https://github.com/linh-nguyen123/Web-development-lab`.
2. Go to **Settings** &rarr; **Pages** (under Code and automation).
3. Under **Build and deployment** &rarr; **Source**, select **Deploy from a branch**.
4. Choose Branch: `main` (or `master`), Folder: `/ (root)`, and click **Save**.
5. Once built, access the live site at: **`https://linh-nguyen123.github.io/Web-development-lab/`**.
