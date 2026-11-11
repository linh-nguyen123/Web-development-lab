# PowerShell OpenXML Document Builder for Lab 02 Report
# Student: Nguyễn Văn Linh (26560037)
# Output: 26560037_NguyễnVănLinh_Lab02.docx

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

$studentName = "Nguy$([char]0x1EC5)n V$([char]0x103)n Linh"
$outDocx = "26560037_Nguy$([char]0x1EC5)nV$([char]0x103)nLinh_Lab02.docx"
$buildDir = "docx_temp_build"

if (Test-Path $buildDir) { Remove-Item -Path $buildDir -Recurse -Force }
Get-ChildItem *.docx | Remove-Item -Force

New-Item -ItemType Directory -Path "$buildDir\_rels" -Force | Out-Null
New-Item -ItemType Directory -Path "$buildDir\word\_rels" -Force | Out-Null
New-Item -ItemType Directory -Path "$buildDir\docProps" -Force | Out-Null

# 1. [Content_Types].xml
$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
'@
[System.IO.File]::WriteAllText("$buildDir\[Content_Types].xml", $contentTypes, [System.Text.Encoding]::UTF8)

# 2. _rels/.rels
$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText("$buildDir\_rels\.rels", $rootRels, [System.Text.Encoding]::UTF8)

# 3. word/_rels/document.xml.rels
$docRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText("$buildDir\word\_rels\document.xml.rels", $docRels, [System.Text.Encoding]::UTF8)

# 4. docProps/core.xml
$coreProps = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Lab 2 Report: Figma-to-Code, Layouts &amp; Responsive Design</dc:title>
  <dc:subject>Web Development Laboratory - Lab 02</dc:subject>
  <dc:creator>$studentName</dc:creator>
  <cp:lastModifiedBy>$studentName</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
</cp:coreProperties>
"@
[System.IO.File]::WriteAllText("$buildDir\docProps\core.xml", $coreProps, [System.Text.Encoding]::UTF8)

# 5. docProps/app.xml
$appProps = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Template>Normal.dotm</Template>
  <Application>Microsoft Word</Application>
  <DocSecurity>0</DocSecurity>
</Properties>
'@
[System.IO.File]::WriteAllText("$buildDir\docProps\app.xml", $appProps, [System.Text.Encoding]::UTF8)

# 6. word/styles.xml
$stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="140" w:line="260" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
'@
[System.IO.File]::WriteAllText("$buildDir\word\styles.xml", $stylesXml, [System.Text.Encoding]::UTF8)

# 7. word/document.xml
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append(@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
'@)

function Emit-Para($text, $size=22, $bold=$false, $color="1E293B", $align="left", $font="Calibri", $spaceBefore=0, $spaceAfter=120, $italic=$false) {
    $b = if ($bold) { "<w:b/>" } else { "" }
    $i = if ($italic) { "<w:i/>" } else { "" }
    $jc = if ($align -ne "left") { "<w:jc w:val=""$align""/>" } else { "" }
    $escaped = [System.Security.SecurityElement]::Escape($text)
    [void]$sb.Append("<w:p><w:pPr>$jc<w:spacing w:before=""$spaceBefore"" w:after=""$spaceAfter""/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""$font"" w:hAnsi=""$font""/><w:sz w:val=""$size""/>$b$i<w:color w:val=""$color""/></w:rPr><w:t xml:space=""preserve"">$escaped</w:t></w:r></w:p>")
}

# Title & Subtitle
Emit-Para -text "LAB 2: FIGMA-TO-CODE, LAYOUTS & RESPONSIVE DESIGN" -size 36 -bold $true -color "0F2537" -align "center" -font "Arial" -spaceBefore 0 -spaceAfter 60
Emit-Para -text "Bridging the Gap between Figma Specifications and Modern Responsive CSS Grids" -size 24 -italic $true -color "64748B" -align "center" -font "Calibri" -spaceBefore 0 -spaceAfter 240

# Student Credentials Table
[void]$sb.Append(@'
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9360" w:type="dxa"/>
    <w:jc w:val="center"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:left w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:bottom w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:right w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
    </w:tblBorders>
  </w:tblPr>
'@)

$creds = @(
    @("Student Full Name:", $studentName),
    @("Student ID:", "26560037"),
    @("Academic Course:", "Web Development Laboratory - Lab Session 02"),
    @("Git Commit Message:", "lab02"),
    @("Target Repository URL:", "https://github.com/linh-nguyen123/Web-development-lab"),
    @("Live GitHub Pages URL:", "https://linh-nguyen123.github.io/Web-development-lab/")
)

foreach ($c in $creds) {
    $k = [System.Security.SecurityElement]::Escape($c[0])
    $v = [System.Security.SecurityElement]::Escape($c[1])
    $fontVal = if ($k -like "*URL*" -or $k -like "*Commit*") { "Consolas" } else { "Arial" }
    $colVal = if ($k -like "*URL*") { "0284C7" } elseif ($k -like "*Commit*") { "0F766E" } else { "0F172A" }
    [void]$sb.Append(@"
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="3000" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="F1F5F9"/>
        <w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>
      </w:tcPr>
      <w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/><w:color w:val="1E293B"/></w:rPr><w:t>$k</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="6360" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>
        <w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>
      </w:tcPr>
      <w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="$fontVal" w:hAnsi="$fontVal"/><w:b/><w:sz w:val="20"/><w:color w:val="$colVal"/></w:rPr><w:t>$v</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
"@)
}
[void]$sb.Append("</w:tbl>")
Emit-Para -text "" -size 10 -spaceAfter 160

# --- SECTION 1: OVERVIEW & ARCHITECTURE ---
Emit-Para -text "1. Overview & Architectural Analysis" -size 28 -bold $true -color "0F2537" -font "Arial" -spaceBefore 200 -spaceAfter 80

# 1.1 Viewport
Emit-Para -text "1.1 The Responsive Viewport Meta Tag" -size 24 -bold $true -color "1E293B" -font "Arial" -spaceBefore 120 -spaceAfter 60
Emit-Para -text "Without configuring the viewport metadata in the HTML <head>, mobile browsers default to rendering pages inside an artificial virtual desktop viewport (typically 980px wide). This causes the webpage to scale down dramatically, making text unreadable and interactive elements unreachable on smartphone screens. In Lab 2, the mandatory viewport meta tag is implemented:" -size 22 -color "334155" -spaceAfter 80
Emit-Para -text '<meta name="viewport" content="width=device-width, initial-scale=1.0">' -size 22 -bold $true -color "0369A1" -font "Consolas" -spaceBefore 40 -spaceAfter 80
Emit-Para -text "• width=device-width: Sets the layout viewport width to match the exact physical device screen width in CSS pixels (e.g., 390px for iPhone 14)." -size 22 -color "334155" -spaceAfter 40
Emit-Para -text "• initial-scale=1.0: Enforces a 1:1 scale ratio upon initial loading, preventing unexpected synthetic zooming by the browser engine." -size 22 -color "334155" -spaceAfter 120

# 1.2 Figma Inspection & Design Tokens
Emit-Para -text "1.2 Figma Inspection and CSS Design Tokens (:root Variables)" -size 24 -bold $true -color "1E293B" -font "Arial" -spaceBefore 160 -spaceAfter 60
Emit-Para -text "Following standard frontend engineering workflows, all layout constraints, color values, typography rules, and spacing offsets were inspected directly from the Figma canvas via Dev Mode. Margins and paddings were inspected by holding the Alt key over adjacent layers. Line-height values were normalized into unitless multipliers, and design tokens were encapsulated in CSS variables under :root:" -size 22 -color "334155" -spaceAfter 100

# Tokens Table
[void]$sb.Append(@'
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9360" w:type="dxa"/>
    <w:jc w:val="center"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:left w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:bottom w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:right w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="0F2537"/><w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/><w:color w:val="FFFFFF"/></w:rPr><w:t>CSS Variable Token</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="0F2537"/><w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/><w:color w:val="FFFFFF"/></w:rPr><w:t>Token Value</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="4160" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="0F2537"/><w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/><w:color w:val="FFFFFF"/></w:rPr><w:t>UI Purpose</w:t></w:r></w:p></w:tc>
  </w:tr>
'@)

$vars = @(
    @("--bg-dark / --bg-surface", "#0a0e17 / #111827", "Deep high-contrast canvas and card surface backgrounds"),
    @("--accent-cyan / --accent-blue", "#38bdf8 / #3b82f6", "Primary Figma brand accent, active link highlights & badges"),
    @("--bg-glass", "rgba(17, 24, 39, 0.82)", "Sticky navigation bar glassmorphism with 14px backdrop blur"),
    @("--shadow-elevation", "0 20px 30px -10px ...", "Card hover elevation and smooth cyan glow box-shadow"),
    @("--transition-smooth", "0.35s cubic-bezier(...)", "Consistent spring motion for card hover, image zoom & menu")
)
foreach ($v in $vars) {
    $t = [System.Security.SecurityElement]::Escape($v[0])
    $val = [System.Security.SecurityElement]::Escape($v[1])
    $role = [System.Security.SecurityElement]::Escape($v[2])
    [void]$sb.Append(@"
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2800" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F8FAFC"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:b/><w:sz w:val="19"/><w:color w:val="0369A1"/></w:rPr><w:t>$t</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="19"/><w:color w:val="1E293B"/></w:rPr><w:t>$val</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="4160" w:type="dxa"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:sz w:val="19"/><w:color w:val="475569"/></w:rPr><w:t>$role</w:t></w:r></w:p></w:tc>
  </w:tr>
"@)
}
[void]$sb.Append("</w:tbl>")
Emit-Para -text "" -size 10 -spaceAfter 140

# 1.3 CSS Flexbox vs. CSS Grid Analysis
Emit-Para -text "1.3 Architectural Comparison: CSS Flexbox vs. CSS Grid" -size 24 -bold $true -color "1E293B" -font "Arial" -spaceBefore 160 -spaceAfter 60
Emit-Para -text "Choosing between Flexbox and Grid represents the central architectural layout decision in modern frontend design:" -size 22 -color "334155" -spaceAfter 100

[void]$sb.Append(@'
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9360" w:type="dxa"/>
    <w:jc w:val="center"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:left w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:bottom w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:right w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="0F2537"/><w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/><w:color w:val="FFFFFF"/></w:rPr><w:t>Dimension / Property</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="3480" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="0F2537"/><w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/><w:color w:val="FFFFFF"/></w:rPr><w:t>CSS Flexbox (1-Dimensional)</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="3480" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="0F2537"/><w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/><w:color w:val="FFFFFF"/></w:rPr><w:t>CSS Grid (2-Dimensional)</w:t></w:r></w:p></w:tc>
  </w:tr>
'@)

$comp = @(
    @("Dimensions", "1-Dimensional: Aligns items along one axis at a time (row OR column).", "2-Dimensional: Simultaneously coordinates items across rows AND columns in a unified layout plane."),
    @("Design Approach", "Content-First: Items determine their own dimensions based on internal content; container aligns them.", "Layout-First: Container defines explicit grid track structures; items are placed within track slots."),
    @("Ideal UI Targets", "Navigation bars, menu drawers, action buttons, centered badges, card content stacking.", "Product grids, multi-column card galleries, complex application dashboards."),
    @("Lab 2 Implementation", "Sticky Navbar: display: flex; justify-content: space-between; align-items: center; glassmorphism.", "Portfolio Gallery: grid-template-columns with auto-fit/repeat and media query breakpoints.")
)

foreach ($c in $comp) {
    $dim = [System.Security.SecurityElement]::Escape($c[0])
    $fb = [System.Security.SecurityElement]::Escape($c[1])
    $gd = [System.Security.SecurityElement]::Escape($c[2])
    [void]$sb.Append(@"
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F8FAFC"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="19"/><w:color w:val="1E293B"/></w:rPr><w:t>$dim</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="3480" w:type="dxa"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:sz w:val="19"/><w:color w:val="334155"/></w:rPr><w:t>$fb</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="3480" w:type="dxa"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:sz w:val="19"/><w:color w:val="334155"/></w:rPr><w:t>$gd</w:t></w:r></w:p></w:tc>
  </w:tr>
"@)
}
[void]$sb.Append("</w:tbl>")
Emit-Para -text "" -size 10 -spaceAfter 140

# 1.4 Responsive Breakpoints
Emit-Para -text "1.4 Mobile-First Responsive Breakpoint Architecture" -size 24 -bold $true -color "1E293B" -font "Arial" -spaceBefore 160 -spaceAfter 60
Emit-Para -text "Styles are authored with a Mobile-First architecture using progressive enhancement:" -size 22 -color "334155" -spaceAfter 60
Emit-Para -text "• Mobile Viewport (< 768px): Base styles provide a single-column layout (.grid-container { grid-template-columns: 1fr; }). The navigation bar collapses links into an interactive mobile drawer toggled by the hamburger button." -size 22 -color "334155" -spaceAfter 40
Emit-Para -text "• Tablet Viewport (768px - 1023px): @media (min-width: 768px) expands the grid to 2 columns (repeat(2, 1fr)), reveals horizontal nav links, and automatically hides the hamburger button." -size 22 -color "334155" -spaceAfter 40
Emit-Para -text "• Desktop Viewport (>= 1024px): @media (min-width: 1024px) establishes a 3-column grid (repeat(3, 1fr)), limits maximum container width to 1240px, and enables fluid zoom and shadow hover interactions." -size 22 -color "334155" -spaceAfter 180

# --- SECTION 2: UI & FEATURE SCREENSHOT CHECKLIST ---
Emit-Para -text "2. UI & Feature Implementation Screenshot Checklist" -size 28 -bold $true -color "0F2537" -font "Arial" -spaceBefore 240 -spaceAfter 80
Emit-Para -text "This section provides pre-formatted screenshot placeholders for report evaluation. Simply copy and paste the captured browser screenshots directly into each placeholder frame below:" -size 22 -color "334155" -spaceAfter 140

function Emit-Placeholder($title, $spec) {
    $tEsc = [System.Security.SecurityElement]::Escape($title)
    $sEsc = [System.Security.SecurityElement]::Escape($spec)
    [void]$sb.Append(@"
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9360" w:type="dxa"/>
    <w:jc w:val="center"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="6" w:space="0" w:color="0284C7"/>
      <w:left w:val="single" w:sz="6" w:space="0" w:color="0284C7"/>
      <w:bottom w:val="single" w:sz="6" w:space="0" w:color="0284C7"/>
      <w:right w:val="single" w:sz="6" w:space="0" w:color="0284C7"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="BAE6FD"/>
      <w:insideV w:val="none"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="9360" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="E0F2FE"/>
        <w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>
      </w:tcPr>
      <w:p><w:pPr><w:spacing w:before="0" w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="21"/><w:color w:val="0369A1"/></w:rPr><w:t>$tEsc</w:t></w:r></w:p>
      <w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/><w:color w:val="475569"/></w:rPr><w:t>$sEsc</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="9360" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="F8FAFC"/>
        <w:tcMar><w:top w:w="800" w:type="dxa"/><w:bottom w:w="800" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>
      </w:tcPr>
      <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:b/><w:sz w:val="20"/><w:color w:val="94A3B8"/></w:rPr><w:t>--- [PASTE SCREENSHOT HERE] ---</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
"@)
    Emit-Para -text "" -size 10 -spaceAfter 140
}

Emit-Placeholder -title "[Placeholder 1] Sticky Navbar & Hover Effects (Flexbox)" -spec "Sticky header with glassmorphism (backdrop-filter: blur(14px)), brand logo, horizontal links, and hover transitions. Flexbox alignment: display: flex; justify-content: space-between; align-items: center;."
Emit-Placeholder -title "[Placeholder 2] Product/Portfolio Grid Layout (CSS Grid 6+ items)" -spec "CSS Grid container showing 6 items: Creative Interface, Dynamic Dashboard, E-Commerce Storefront, Social Media App, Finance Tracker, and Fitness Companion."
Emit-Placeholder -title "[Placeholder 3] Responsive Mobile Layout (< 768px - 1 column)" -spec "Mobile viewport (< 768px): 1-column card stack, visible hamburger button, and expandable navigation drawer."
Emit-Placeholder -title "[Placeholder 4] Responsive Tablet Layout (768px - 1023px - 2 columns)" -spec "Tablet viewport (768px - 1023px): 2-column card grid, hidden hamburger button, and horizontal desktop navbar links active."
Emit-Placeholder -title "[Placeholder 5] Responsive Desktop Layout (>= 1024px - 3 columns)" -spec "Desktop viewport (>= 1024px): 3-column card grid layout, centered 1240px container, and full desktop navbar."
Emit-Placeholder -title "[Placeholder 6] Interactive Hover States (Image Zoom & Elevation)" -spec "Active cursor hover showing card image scale (transform: scale(1.08)), card elevation (translateY(-8px)), and glowing box-shadow."

# --- SECTION 3: GIT & GITHUB PAGES ---
Emit-Para -text "3. Git Deployment & GitHub Pages Verification" -size 28 -bold $true -color "0F2537" -font "Arial" -spaceBefore 240 -spaceAfter 80
Emit-Para -text "The source code has been structured, committed with message 'lab02', pushed to GitHub, and hosted on GitHub Pages:" -size 22 -color "334155" -spaceAfter 100

[void]$sb.Append(@'
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9360" w:type="dxa"/>
    <w:jc w:val="center"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:left w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:bottom w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:right w:val="single" w:sz="6" w:space="0" w:color="0F2537"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F1F5F9"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/><w:color w:val="1E293B"/></w:rPr><w:t>Target Git Repository:</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="6360" w:type="dxa"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="19"/><w:color w:val="0284C7"/></w:rPr><w:t>https://github.com/linh-nguyen123/Web-development-lab</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F1F5F9"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/><w:color w:val="1E293B"/></w:rPr><w:t>Git Commit Message:</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="6360" w:type="dxa"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:b/><w:sz w:val="19"/><w:color w:val="0F766E"/></w:rPr><w:t>lab02</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F1F5F9"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/><w:color w:val="1E293B"/></w:rPr><w:t>GitHub Pages Live URL:</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="6360" w:type="dxa"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/><w:right w:w="140" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="19"/><w:color w:val="0284C7"/></w:rPr><w:t>https://linh-nguyen123.github.io/Web-development-lab/</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>
'@)
Emit-Para -text "" -size 10 -spaceAfter 120

Emit-Para -text "Exact Terminal Commands for Git Execution:" -size 22 -bold $true -color "0F2537" -font "Arial" -spaceBefore 80 -spaceAfter 40
Emit-Para -text "git add index.html styles.css assets/ 26560037_NguyễnVănLinh_Lab02.docx generate_report.py" -size 20 -bold $true -color "0F172A" -font "Consolas" -spaceAfter 20
Emit-Para -text 'git commit -m "lab02"' -size 20 -bold $true -color "0F766E" -font "Consolas" -spaceAfter 20
Emit-Para -text "git branch -M main" -size 20 -bold $true -color "0F172A" -font "Consolas" -spaceAfter 20
Emit-Para -text "git remote add origin https://github.com/linh-nguyen123/Web-development-lab.git" -size 20 -bold $true -color "0F172A" -font "Consolas" -spaceAfter 20
Emit-Para -text "git push -u origin main" -size 20 -bold $true -color "0369A1" -font "Consolas" -spaceAfter 140

Emit-Placeholder -title "[Placeholder 7] Live GitHub Pages Deployment Proof & Repository Link" -spec "Screenshot from GitHub repository Settings -> Pages displaying the active GitHub Pages green checkmark and public live URL."

# Close document
[void]$sb.Append(@'
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
'@)

[System.IO.File]::WriteAllText("$buildDir\word\document.xml", $sb.ToString(), [System.Text.Encoding]::UTF8)

# Compress to .docx
Add-Type -AssemblyName "System.IO.Compression.FileSystem"
$targetZip = Join-Path (Get-Location).Path $outDocx
[System.IO.Compression.ZipFile]::CreateFromDirectory((Resolve-Path $buildDir).Path, $targetZip)

Remove-Item -Path $buildDir -Recurse -Force
Write-Host "SUCCESS: Generated $outDocx successfully!"
