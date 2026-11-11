# PowerShell OpenXML Document Builder for Lab 02 Report (English Figures & Links)
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
  <dc:subject>Lab 02 Screenshot Verification Report</dc:subject>
  <dc:creator>$studentName</dc:creator>
  <cp:lastModifiedBy>$studentName</cp:lastModifiedBy>
  <cp:revision>4</cp:revision>
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
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120" w:line="240" w:lineRule="auto"/>
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

function Emit-Para($text, $size=22, $bold=$false, $color="0F172A", $align="left", $font="Arial", $spaceBefore=0, $spaceAfter=100, $italic=$false) {
    $b = if ($bold) { "<w:b/>" } else { "" }
    $i = if ($italic) { "<w:i/>" } else { "" }
    $jc = if ($align -ne "left") { "<w:jc w:val=""$align""/>" } else { "" }
    $escaped = [System.Security.SecurityElement]::Escape($text)
    [void]$sb.Append("<w:p><w:pPr>$jc<w:spacing w:before=""$spaceBefore"" w:after=""$spaceAfter""/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""$font"" w:hAnsi=""$font""/><w:sz w:val=""$size""/>$b$i<w:color w:val=""$color""/></w:rPr><w:t xml:space=""preserve"">$escaped</w:t></w:r></w:p>")
}

# Tiêu đề báo cáo
Emit-Para -text "LAB 2 REPORT: FIGMA-TO-CODE & RESPONSIVE DESIGN" -size 32 -bold $true -color "1E1B4B" -align "center" -font "Arial" -spaceBefore 0 -spaceAfter 40
Emit-Para -text "Homework Assignment: VibeTech Responsive E-Commerce Workspace Gear" -size 22 -bold $true -color "4F46E5" -align "center" -font "Arial" -spaceBefore 0 -spaceAfter 180

# Bảng thông tin sinh viên và link GitHub
[void]$sb.Append(@'
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9360" w:type="dxa"/>
    <w:jc w:val="center"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="8" w:space="0" w:color="4F46E5"/>
      <w:left w:val="single" w:sz="8" w:space="0" w:color="4F46E5"/>
      <w:bottom w:val="single" w:sz="8" w:space="0" w:color="4F46E5"/>
      <w:right w:val="single" w:sz="8" w:space="0" w:color="4F46E5"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
    </w:tblBorders>
  </w:tblPr>
'@)

$creds = @(
    @("Student Full Name:", $studentName),
    @("Student ID:", "26560037"),
    @("Course:", "Web Development Laboratory - Lab Session 02"),
    @("Project Name:", "VibeTech E-Commerce Workspace Gear (Unleash the Next Level)"),
    @("Git Commit Message:", "lab02"),
    @("GitHub Repository URL:", "https://github.com/linh-nguyen123/Web-development-lab"),
    @("GitHub Pages Live Demo:", "https://linh-nguyen123.github.io/Web-development-lab/")
)

foreach ($c in $creds) {
    $k = [System.Security.SecurityElement]::Escape($c[0])
    $v = [System.Security.SecurityElement]::Escape($c[1])
    $isLink = $k -like "*URL*" -or $k -like "*Demo*"
    $isCommit = $k -like "*Commit*"
    $fontVal = if ($isLink -or $isCommit) { "Consolas" } else { "Arial" }
    $colVal = if ($isLink) { "0284C7" } elseif ($isCommit) { "0F766E" } else { "0F172A" }
    $boldVal = $isCommit -or ($k -like "*Name*") -or ($k -like "*ID*")
    $bTag = if ($boldVal) { "<w:b/>" } else { "" }
    $uTag = if ($isLink) { "<w:u w:val=""single""/>" } else { "" }

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
      <w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="$fontVal" w:hAnsi="$fontVal"/>$bTag$uTag<w:sz w:val="20"/><w:color w:val="$colVal"/></w:rPr><w:t>$v</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
"@)
}
[void]$sb.Append("</w:tbl>")
Emit-Para -text "" -size 10 -spaceAfter 180

# Section Header
Emit-Para -text "LAB EXECUTION & SCREENSHOT EVIDENCE" -size 24 -bold $true -color "1E1B4B" -font "Arial" -spaceBefore 120 -spaceAfter 120

# Hàm sinh khung dán ảnh tiện lợi bằng tiếng Anh
function Emit-ImagePlaceholder($figureNum, $title, $description) {
    $tEsc = [System.Security.SecurityElement]::Escape("$figureNum: $title")
    $dEsc = [System.Security.SecurityElement]::Escape($description)
    [void]$sb.Append(@"
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9360" w:type="dxa"/>
    <w:jc w:val="center"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="6" w:space="0" w:color="4F46E5"/>
      <w:left w:val="single" w:sz="6" w:space="0" w:color="4F46E5"/>
      <w:bottom w:val="single" w:sz="6" w:space="0" w:color="4F46E5"/>
      <w:right w:val="single" w:sz="6" w:space="0" w:color="4F46E5"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="C7D2FE"/>
      <w:insideV w:val="none"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="9360" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="EEF2FF"/>
        <w:tcMar><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>
      </w:tcPr>
      <w:p><w:pPr><w:spacing w:before="0" w:after="40"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="21"/><w:color w:val="312E81"/></w:rPr><w:t>$tEsc</w:t></w:r></w:p>
      <w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/><w:color w:val="475569"/></w:rPr><w:t>$dEsc</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="9360" w:type="dxa"/>
        <w:shd w:val="clear" w:color="auto" w:fill="F8FAFC"/>
        <w:tcMar><w:top w:w="900" w:type="dxa"/><w:bottom w:w="900" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>
      </w:tcPr>
      <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:b/><w:sz w:val="20"/><w:color w:val="94A3B8"/></w:rPr><w:t>[ PASTE SCREENSHOT HERE ]</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
"@)
    Emit-Para -text "" -size 10 -spaceAfter 140
}

# 6 Figures in English
Emit-ImagePlaceholder -figureNum "Figure 1" -title "Sticky Navigation Bar with CSS Flexbox & Glassmorphism Effect" -description "Capture of the sticky header featuring VibeTech brand identity, centered navigation links (Home, Products, Deals, About), search icon, shopping cart badge, and frosted glass background blur."

Emit-ImagePlaceholder -figureNum "Figure 2" -title "Responsive Desktop Product Grid Layout (CSS Grid - 4 Columns, 8 Items)" -description "Full desktop browser view (>= 1024px) showcasing the complete 4-column product grid with 8 curated workspace gear cards, matching Figma 1920w and 1024w frames."

Emit-ImagePlaceholder -figureNum "Figure 3" -title "Responsive Mobile Viewport & Hamburger Menu Drawer (< 768px - 1 Column)" -description "Browser DevTools mobile device emulation (< 768px, ~390px width) showing single-column product stacking and the interactive mobile navigation drawer toggled open."

Emit-ImagePlaceholder -figureNum "Figure 4" -title "Responsive Tablet Viewport Layout (768px - 1023px - 2 Columns)" -description "Browser DevTools tablet emulation (e.g. iPad, ~820px width) illustrating the 2-column product grid transition with desktop horizontal navigation links visible, matching Figma 768w light frame."

Emit-ImagePlaceholder -figureNum "Figure 5" -title "Interactive Card Hover States (Smooth Image Zoom & Elevation)" -description "Active cursor hover interaction on a product card, demonstrating smooth image scaling (scale(1.08)), card elevation (translateY(-6px)), and glowing box-shadow elevation."

Emit-ImagePlaceholder -figureNum "Figure 6" -title "Git Commit Verification (lab02) & Live GitHub Pages Deployment Proof" -description "Terminal execution proof showing git commit message 'lab02' alongside the live website URL hosted and functioning on GitHub Pages."

# Kết thúc document
[void]$sb.Append(@'
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
'@)

[System.IO.File]::WriteAllText("$buildDir\word\document.xml", $sb.ToString(), [System.Text.Encoding]::UTF8)

# Nén thành file docx
Add-Type -AssemblyName "System.IO.Compression.FileSystem"
$targetZip = Join-Path (Get-Location).Path $outDocx
[System.IO.Compression.ZipFile]::CreateFromDirectory((Resolve-Path $buildDir).Path, $targetZip)

Remove-Item -Path $buildDir -Recurse -Force
Write-Host "SUCCESS: Generated $outDocx successfully"
