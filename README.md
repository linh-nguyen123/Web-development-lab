# Lab 2: Figma-to-Code, Layouts & Responsive Design

**Homework Assignment:** VibeTech Responsive E-Commerce Workspace Gear Landing Page  
**Design Heading:** *Unleash the Next Level*  
**Student Name:** Nguyễn Văn Linh  
**Student ID:** 26560037  
**Submission Commit:** `lab02`  
**GitHub Repository:** [https://github.com/linh-nguyen123/Web-development-lab](https://github.com/linh-nguyen123/Web-development-lab)  
**Live GitHub Pages:** [https://linh-nguyen123.github.io/Web-development-lab/](https://linh-nguyen123.github.io/Web-development-lab/)  
**Word Report:** `26560037_NguyễnVănLinh_Lab02.docx`

---

## 📌 Homework Implementation Summary

This project implements the official Homework assignment from Lab 2 ("Responsive Landing Page") based on the Figma design specifications:

1. **Figma Dev Mode Tokens:**
   - Colors: `#f8fafc` page canvas, `#ffffff` card surface, `#4f46e5` brand indigo, `#0f172a` button/badge slate.
   - Typography: `Inter` font-family with unitless responsive line-heights.
   - Assets: Direct vector SVGs & high-fidelity assets inside `assets/`.

2. **Sticky Navigation Bar (CSS Flexbox):**
   - Implemented via `display: flex; justify-content: space-between; align-items: center;`.
   - `position: sticky; top: 0;` with glassmorphism `backdrop-filter: blur(12px)`.
   - Contains VibeTech logo, navigation links (`Home`, `Products`, `Deals`, `About`), search icon, cart counter badge, and mobile hamburger drawer.

3. **Responsive Product Grid (CSS Grid):**
   - 8 Curated Workspace Gear Items:
     1. `Wireless Mouse` (Badge: **Hot**) — `$59.00`
     2. `Tactix K85 Mechanical Keyboard` — `$129.00`
     3. `ZenSound ANC Headset` (Yellow Background) — `$189.00`
     4. `NovaFit Smart Watch` (Badge: **Sale**) — `$99.00`
     5. `Horizon 27" IPS QHD Display` (Badge: **New**) — `$299.00`
     6. `StreamMic Pro Cardioid` — `$149.00`
     7. `Lumos RGB Smart Desk Lamp` (Badge: **Sale**) — `$45.00`
     8. `ErgoStand Aluminum Laptop Riser` — `$69.00`

4. **Mobile-First Responsive Layouts:**
   - **Mobile Viewport (`< 768px`):** 1-column product grid stack (`grid-template-columns: 1fr;`), mobile hamburger drawer menu.
   - **Tablet Viewport (`768px - 1023px`, Figma `768w light`):** 2-column product grid (`repeat(2, 1fr)`), horizontal desktop links.
   - **Desktop Viewport (`>= 1024px`, Figma `1024w` & `1920w light`):** 4-column product grid (`repeat(4, 1fr)`), centered 1280px container.

5. **Interactive Hover States:**
   - Card elevation: `transform: translateY(-6px); box-shadow: 0 16px 28px -6px rgba(0, 0, 0, 0.09);`
   - Image zoom: `.card-image-box img { transform: scale(1.08); transition: transform 0.35s ease; }`
   - Add to Cart interactive counter and visual feedback.

---

## 📂 Project Directory Structure
```text
├── index.html                           # Semantic HTML5 with viewport meta tag & 8-product grid
├── styles.css                           # Mobile-First CSS Grid (4 cols, 2 cols, 1 col) & Flexbox
├── assets/                              # Figma design icons & product illustrations
│   ├── vibetech_logo.svg                # VibeTech vector logo
│   ├── product1.svg / product1.png      # Wireless Mouse
│   ├── product2.svg / product2.png      # Tactix K85 Mechanical
│   ├── product3.svg / product3.png      # ZenSound ANC Headset
│   ├── product4.svg / product4.png      # NovaFit Smart Watch
│   ├── product5.svg / product5.png      # Horizon 27" IPS QHD Display
│   ├── product6.svg / product6.png      # StreamMic Pro Cardioid
│   ├── product7.svg / product7.png      # Lumos RGB Smart Desk Lamp
│   └── product8.svg / product8.png      # ErgoStand Aluminum Laptop Riser
├── 26560037_NguyễnVănLinh_Lab02.docx    # Official Word Report Document
├── generate_report.py                   # Python-docx Report Builder Script
├── build_docx.ps1                       # Native PowerShell OpenXML Document Builder
├── .gitignore                           # Git ignore rules
└── README.md                            # Project documentation
```

---

## 🚀 Git Commands
```bash
# 1. Stage all files
git add .

# 2. Commit with mandatory message
git commit -m "lab02"

# 3. Push to GitHub
git push -u origin main
```
