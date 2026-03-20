# doc-convert.ps1 — Image & Document converter with progress UI
# v1.0.0
# Supported tools: ImageMagick (images + PDF<->image), LibreOffice (documents), Python/python-docx (image->DOCX)

param(
    [string]$Path,
    [string]$ListFile
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ══════════════════════════════════════
#  TOOL DETECTION
# ══════════════════════════════════════

$magickCmd = $null
if (Get-Command magick -ErrorAction SilentlyContinue) { $magickCmd = "magick" }

$sofficeCmd = $null
$softicePaths = @(
    "soffice",
    "$env:ProgramFiles\LibreOffice\program\soffice.exe",
    "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe",
    "$env:LOCALAPPDATA\Programs\LibreOffice\program\soffice.exe"
)
foreach ($p in $softicePaths) {
    try {
        if (Get-Command $p -ErrorAction SilentlyContinue) { $sofficeCmd = $p; break }
    } catch {}
}
if (-not $sofficeCmd) {
    foreach ($p in $softicePaths[1..3]) {
        if (Test-Path $p) { $sofficeCmd = $p; break }
    }
}

$pythonCmd = $null
foreach ($pc in @("python", "python3")) {
    if (Get-Command $pc -ErrorAction SilentlyContinue) { $pythonCmd = $pc; break }
}

if (-not $magickCmd -and -not $sofficeCmd) {
    [System.Windows.Forms.MessageBox]::Show(
        "No conversion tools found.`n`nPlease install at least one of:`n`n• ImageMagick  (images + PDF↔image)`n  https://imagemagick.org/script/download.php#windows`n`n• LibreOffice  (DOCX/PPTX/XLSX → PDF and PDF → DOCX)`n  https://www.libreoffice.org/download/",
        "Doc Convert", "OK", "Error") | Out-Null
    exit 1
}

# ══════════════════════════════════════
#  FILE TYPE GROUPS
# ══════════════════════════════════════

$imageExts   = @('.jpg','.jpeg','.png','.webp','.bmp','.tiff','.tif','.gif','.heic','.heif','.avif','.jxl','.jp2')
$pdfExts     = @('.pdf')
$docExts     = @('.docx','.doc','.odt','.rtf')
$spreadExts  = @('.xlsx','.xls','.ods')
$presExts    = @('.pptx','.ppt','.odp')
$allDocExts  = $docExts + $spreadExts + $presExts

# ══════════════════════════════════════
#  LOAD FILE LIST
# ══════════════════════════════════════

$initialPaths = @()
if ($ListFile -and (Test-Path $ListFile)) {
    $initialPaths = @(Get-Content $ListFile -Encoding UTF8 |
        Where-Object { $_.Trim() -ne "" } |
        ForEach-Object { $_.Trim() })
    Remove-Item $ListFile -Force -ErrorAction SilentlyContinue
} elseif ($Path) {
    $initialPaths = @($Path)
}

if ($initialPaths.Count -eq 0) { exit 0 }

# ══════════════════════════════════════
#  DETERMINE AVAILABLE CONVERSIONS
# ══════════════════════════════════════

$hasImages = @($initialPaths | Where-Object { $imageExts -contains ([IO.Path]::GetExtension($_).ToLower()) })
$hasPdf    = @($initialPaths | Where-Object { $pdfExts   -contains ([IO.Path]::GetExtension($_).ToLower()) })
$hasDocs   = @($initialPaths | Where-Object { $allDocExts -contains ([IO.Path]::GetExtension($_).ToLower()) })

# Build ordered list: label -> format-key
$sections = [ordered]@{}  # section-name -> [ordered]@{label->key}

if ($magickCmd -and $hasImages.Count -gt 0) {
    $imgTargets = [ordered]@{}
    # Determine current extensions to skip same-format conversions
    $currentImageExts = @($hasImages | ForEach-Object { [IO.Path]::GetExtension($_).ToLower() } | Sort-Object -Unique)
    $skipJpg  = ($currentImageExts -contains '.jpg' -or $currentImageExts -contains '.jpeg') -and $currentImageExts.Count -eq 1
    $skipPng  = ($currentImageExts -contains '.png')  -and $currentImageExts.Count -eq 1
    $skipWebp = ($currentImageExts -contains '.webp') -and $currentImageExts.Count -eq 1
    $skipBmp  = ($currentImageExts -contains '.bmp')  -and $currentImageExts.Count -eq 1
    $skipTiff = ($currentImageExts -contains '.tiff' -or $currentImageExts -contains '.tif') -and $currentImageExts.Count -eq 1
    $skipGif  = ($currentImageExts -contains '.gif')  -and $currentImageExts.Count -eq 1

    if (-not $skipJpg)  { $imgTargets["JPG"]       = "img:jpg"  }
    if (-not $skipPng)  { $imgTargets["PNG"]        = "img:png"  }
    if (-not $skipWebp) { $imgTargets["WebP"]       = "img:webp" }
    if (-not $skipBmp)  { $imgTargets["BMP"]        = "img:bmp"  }
    if (-not $skipTiff) { $imgTargets["TIFF"]       = "img:tiff" }
    if (-not $skipGif)  { $imgTargets["GIF"]        = "img:gif"  }
    $imgTargets["PDF (single file)"] = "img:pdf"
    if ($pythonCmd) {
        $imgTargets["DOCX (embed image)"] = "img:docx"
    }
    $sections["Images"] = $imgTargets
}

if ($magickCmd -and $hasPdf.Count -gt 0) {
    $pdfTargets = [ordered]@{}
    $pdfTargets["JPG  (one file per page)"] = "pdf:jpg"
    $pdfTargets["PNG  (one file per page)"] = "pdf:png"
    $sections["PDF → Image"] = $pdfTargets
}

if ($sofficeCmd) {
    $docTargets = [ordered]@{}
    if ($hasDocs.Count -gt 0) { $docTargets["PDF"] = "doc:pdf" }
    if ($hasPdf.Count -gt 0)  { $docTargets["DOCX"] = "pdf:docx" }
    if ($docTargets.Count -gt 0) { $sections["Documents (LibreOffice)"] = $docTargets }
} elseif ($hasDocs.Count -gt 0 -or $hasPdf.Count -gt 0) {
    # LibreOffice not installed — add an install-prompt entry
    $sections["Documents  (LibreOffice required)"] = [ordered]@{
        "DOCX / XLSX / PPTX → PDF  [install LibreOffice]" = "hint:libreoffice"
        "PDF → DOCX  [install LibreOffice]"               = "hint:libreoffice"
    }
}

if ($sections.Count -eq 0) {
    $allExts = @($initialPaths | ForEach-Object { [IO.Path]::GetExtension($_).ToLower() } | Sort-Object -Unique) -join ", "
    $msgForm = New-Object System.Windows.Forms.Form
    $msgForm.TopMost = $true
    $msgForm.WindowState = "Minimized"
    $msgForm.Show()
    [System.Windows.Forms.MessageBox]::Show(
        $msgForm,
        "No supported conversions for: $allExts`n`nSupported types:`n• Images: jpg, png, webp, bmp, tiff, gif, heic, avif, jp2, jxl`n• PDF (→ jpg/png pages)`n• Documents (docx, xlsx, pptx) → needs LibreOffice",
        "Doc Convert", "OK", "Information") | Out-Null
    $msgForm.Close()
    exit 0
}

# ══════════════════════════════════════
#  FORMAT PICKER UI
# ══════════════════════════════════════

$pickerForm = New-Object System.Windows.Forms.Form
$pickerForm.Text = "Doc Convert"
$pickerForm.StartPosition = "CenterScreen"
$pickerForm.FormBorderStyle = "FixedSingle"
$pickerForm.MaximizeBox = $false
$pickerForm.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 30)
$pickerForm.ForeColor = [System.Drawing.Color]::White
$pickerForm.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$pickerForm.TopMost = $true

$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Text = "Convert to..."
$headerLabel.Location = New-Object System.Drawing.Point(20, 14)
$headerLabel.Size = New-Object System.Drawing.Size(260, 28)
$headerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$headerLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 200, 160)
$pickerForm.Controls.Add($headerLabel)

$script:pickedFormat = $null
$yPos = 50

foreach ($section in $sections.GetEnumerator()) {
    $sectionLabel = New-Object System.Windows.Forms.Label
    $sectionLabel.Text = $section.Key
    $sectionLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $sectionLabel.Size = New-Object System.Drawing.Size(260, 18)
    $sectionLabel.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $sectionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $pickerForm.Controls.Add($sectionLabel)
    $yPos += 20

    foreach ($entry in $section.Value.GetEnumerator()) {
        $fmtKey = $entry.Value
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $entry.Key
        $btn.Location = New-Object System.Drawing.Point(20, $yPos)
        $btn.Size = New-Object System.Drawing.Size(260, 30)
        $btn.FlatStyle = "Flat"
        $btn.FlatAppearance.BorderSize = 0
        $btn.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 52)
        $btn.ForeColor = [System.Drawing.Color]::White
        $btn.TextAlign = "MiddleLeft"
        $btn.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
        $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btn.Tag = $fmtKey

        if ($fmtKey -like "hint:*") {
            $btn.ForeColor = [System.Drawing.Color]::FromArgb(160, 120, 60)
            $btn.BackColor = [System.Drawing.Color]::FromArgb(45, 38, 28)
            $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(100, 70, 30)
            $btn.FlatAppearance.BorderSize = 1
            $btn.Add_Click({
                $pickerForm.TopMost = $false
                $ans = [System.Windows.Forms.MessageBox]::Show(
                    $pickerForm,
                    "This conversion requires LibreOffice (free).`n`nInstall it from libreoffice.org, then re-run Doc Convert.`n`nOpen download page now?",
                    "LibreOffice Required",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Information)
                if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Start-Process "https://www.libreoffice.org/download/"
                }
                $pickerForm.Close()
            })
        } else {
            $btn.Add_Click({
                $script:pickedFormat = $this.Tag
                $pickerForm.Close()
            })
            $btn.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(65, 65, 68) })
            $btn.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 52) })
        }

        $pickerForm.Controls.Add($btn)
        $yPos += 34
    }
    $yPos += 6
}

$pickerForm.ClientSize = New-Object System.Drawing.Size(300, ($yPos + 8))
[System.Windows.Forms.Application]::Run($pickerForm)

if (-not $script:pickedFormat) { exit 0 }
$chosenFormat = $script:pickedFormat

# Parse format key: "category:target"
$fmtParts    = $chosenFormat.Split(":")
$fmtCategory = $fmtParts[0]
$fmtTarget   = $fmtParts[1]

# ══════════════════════════════════════
#  FILTER FILES FOR CHOSEN CONVERSION
# ══════════════════════════════════════

$files = @()
foreach ($p in $initialPaths) {
    if (-not (Test-Path $p -PathType Leaf)) { continue }
    $ext = [IO.Path]::GetExtension($p).ToLower()
    $keep = switch ($fmtCategory) {
        "img" { $imageExts  -contains $ext }
        "pdf" { $pdfExts    -contains $ext }
        "doc" { $allDocExts -contains $ext }
        default { $false }
    }
    if ($keep) { $files += Get-Item $p }
}

if ($files.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "No matching files for this conversion type.",
        "Doc Convert", "OK", "Warning") | Out-Null
    exit 0
}

$formatLabel = switch ($chosenFormat) {
    "img:jpg"   { "Image → JPG" }
    "img:png"   { "Image → PNG" }
    "img:webp"  { "Image → WebP" }
    "img:bmp"   { "Image → BMP" }
    "img:tiff"  { "Image → TIFF" }
    "img:gif"   { "Image → GIF" }
    "img:pdf"   { "Image → PDF" }
    "img:docx"  { "Image → DOCX" }
    "pdf:jpg"   { "PDF → JPG (pages)" }
    "pdf:png"   { "PDF → PNG (pages)" }
    "doc:pdf"   { "Document → PDF" }
    "pdf:docx"  { "PDF → DOCX" }
    default     { $chosenFormat }
}

# ══════════════════════════════════════
#  PROGRESS UI
# ══════════════════════════════════════

$form = New-Object System.Windows.Forms.Form
$form.Text = "Doc Convert"
$form.Size = New-Object System.Drawing.Size(570, 430)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 30)
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Converting $($files.Count) file(s) — $formatLabel"
$titleLabel.Location = New-Object System.Drawing.Point(20, 16)
$titleLabel.Size = New-Object System.Drawing.Size(515, 26)
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(80, 200, 160)
$form.Controls.Add($titleLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 52)
$progressBar.Size = New-Object System.Drawing.Size(515, 24)
$progressBar.Style = "Continuous"
$progressBar.Minimum = 0
$progressBar.Maximum = [Math]::Max($files.Count, 1)
$progressBar.Value = 0
$form.Controls.Add($progressBar)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Starting..."
$statusLabel.Location = New-Object System.Drawing.Point(20, 84)
$statusLabel.Size = New-Object System.Drawing.Size(515, 20)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
$form.Controls.Add($statusLabel)

$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(20, 112)
$listView.Size = New-Object System.Drawing.Size(515, 230)
$listView.View = "Details"
$listView.FullRowSelect = $true
$listView.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 42)
$listView.ForeColor = [System.Drawing.Color]::White
$listView.Font = New-Object System.Drawing.Font("Consolas", 9)
$listView.BorderStyle = "None"
$listView.HeaderStyle = "Nonclickable"
$listView.GridLines = $false
$listView.Columns.Add("File", 330) | Out-Null
$listView.Columns.Add("Status", 90) | Out-Null
$listView.Columns.Add("Size", 80) | Out-Null

foreach ($file in $files) {
    $item = New-Object System.Windows.Forms.ListViewItem($file.Name)
    $item.SubItems.Add("Queued") | Out-Null
    $sizeStr = if ($file.Length -ge 1MB) {
        "$([math]::Round($file.Length / 1MB, 1)) MB"
    } else {
        "$([math]::Round($file.Length / 1KB, 0)) KB"
    }
    $item.SubItems.Add($sizeStr) | Out-Null
    $item.ForeColor = [System.Drawing.Color]::FromArgb(130, 130, 130)
    $listView.Items.Add($item) | Out-Null
}
$form.Controls.Add($listView)

$closeBtn = New-Object System.Windows.Forms.Button
$closeBtn.Text = "Close"
$closeBtn.Location = New-Object System.Drawing.Point(430, 355)
$closeBtn.Size = New-Object System.Drawing.Size(105, 32)
$closeBtn.FlatStyle = "Flat"
$closeBtn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$closeBtn.BackColor = [System.Drawing.Color]::FromArgb(55, 55, 57)
$closeBtn.ForeColor = [System.Drawing.Color]::White
$closeBtn.Visible = $false
$closeBtn.Add_Click({ $form.Close() })
$form.Controls.Add($closeBtn)

# ══════════════════════════════════════
#  CONVERSION ENGINE (async, timer-polled)
# ══════════════════════════════════════

$script:jobQueue    = New-Object System.Collections.Queue
$script:runningJob  = $null   # @{ Process; Idx; OutPath; OutDir; IsPageJob }
$script:doneCount   = 0
$script:successCount = 0
$script:failCount   = 0

for ($i = 0; $i -lt $files.Count; $i++) { $script:jobQueue.Enqueue($i) }

function Start-ConversionJob {
    param([int]$Idx)

    $file     = $files[$Idx]
    $inPath   = $file.FullName
    $dir      = $file.DirectoryName
    $base     = [IO.Path]::GetFileNameWithoutExtension($file.Name)
    $outPath  = $null
    $outDir   = $null
    $isPageJob = $false

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.UseShellExecute        = $false
    $pinfo.RedirectStandardOutput = $true
    $pinfo.RedirectStandardError  = $true
    $pinfo.CreateNoWindow         = $true

    switch ($chosenFormat) {

        # ── Image → Image ──────────────────────────────
        { $_ -match "^img:(jpg|png|webp|bmp|tiff|gif)$" } {
            $ext     = if ($fmtTarget -eq "jpg") { ".jpg" } else { ".$fmtTarget" }
            $outPath = Join-Path $dir "$base$ext"
            if ($outPath -eq $inPath) { $outPath = Join-Path $dir "${base}_converted$ext" }
            $pinfo.FileName  = "magick"
            $args = @()
            # Quality settings per format
            switch ($fmtTarget) {
                "jpg"  { $args = @("$inPath", "-quality", "92", "$outPath") }
                "png"  { $args = @("$inPath", "-quality", "9",  "$outPath") }
                "webp" { $args = @("$inPath", "-quality", "85", "$outPath") }
                "gif"  { $args = @("$inPath", "-dither", "Riemersma", "-colors", "256", "$outPath") }
                default { $args = @("$inPath", "$outPath") }
            }
            $pinfo.Arguments = ($args | ForEach-Object { "`"$_`"" }) -join " "
        }

        # ── Image → PDF ────────────────────────────────
        "img:pdf" {
            $outPath = Join-Path $dir "$base.pdf"
            if ($outPath -eq $inPath) { $outPath = Join-Path $dir "${base}_converted.pdf" }
            $pinfo.FileName  = "magick"
            # -density 72 -units PixelsPerInch ensures correct page size metadata in PDF
            $pinfo.Arguments = "`"$inPath`" -density 72 -units PixelsPerInch -quality 92 `"$outPath`""
        }

        # ── Image → DOCX (embed via python-docx) ───────
        "img:docx" {
            $outPath  = Join-Path $dir "$base.docx"
            if ($outPath -eq $inPath) { $outPath = Join-Path $dir "${base}_converted.docx" }
            # Write temp python script
            $escapedIn  = $inPath.Replace('\','\\')
            $escapedOut = $outPath.Replace('\','\\')
            $pyCode = @"
import sys, subprocess
try:
    import docx
except ImportError:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'python-docx', '-q'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    import docx
from docx.shared import Inches
doc = docx.Document()
doc.add_picture(r"$escapedIn", width=Inches(6))
doc.save(r"$escapedOut")
print("OK")
"@
            $pyFile = Join-Path $env:TEMP "doc_convert_embed_$Idx.py"
            [System.IO.File]::WriteAllText($pyFile, $pyCode, [System.Text.Encoding]::UTF8)
            $pinfo.FileName  = $pythonCmd
            $pinfo.Arguments = "`"$pyFile`""
        }

        # ── PDF → Images (pages via PyMuPDF) ───────────
        { $_ -match "^pdf:(jpg|png)$" } {
            $isPageJob  = $true
            $outDir     = Join-Path $dir "${base}_pages"
            New-Item $outDir -ItemType Directory -Force | Out-Null
            $escapedIn  = $inPath.Replace('\','\\')
            $escapedOut = $outDir.Replace('\','\\')
            $escapedBase = $base.Replace("'","\'")
            $pyFmt      = $fmtTarget   # jpg or png
            $pyQual     = if ($fmtTarget -eq "jpg") { 92 } else { 9 }
            $pyCode = @"
import sys, subprocess, os
try:
    import fitz
except ImportError:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pymupdf', '-q'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    import fitz

doc = fitz.open(r"$escapedIn")
out_dir = r"$escapedOut"
fmt = "$pyFmt"
qual = $pyQual
MAX_DIM = 3000  # safety cap per axis

for i, page in enumerate(doc):
    # Compute a safe DPI (cap output at MAX_DIM px on longest side)
    rect = page.rect
    longest = max(rect.width, rect.height)
    dpi = 150
    if longest * (dpi / 72.0) > MAX_DIM:
        dpi = int(MAX_DIM / longest * 72)
        dpi = max(dpi, 36)

    pix = page.get_pixmap(dpi=dpi)
    name = "${escapedBase}_%04d.%s" % (i, fmt)
    out_path = os.path.join(out_dir, name)
    if fmt == "jpg":
        pix.save(out_path, jpg_quality=qual)
    else:
        pix.save(out_path)

print("OK pages=%d" % doc.page_count)
"@
            $pyFile = Join-Path $env:TEMP "doc_convert_pdf2img_$Idx.py"
            [System.IO.File]::WriteAllText($pyFile, $pyCode, [System.Text.Encoding]::UTF8)
            $pinfo.FileName  = $pythonCmd
            $pinfo.Arguments = "`"$pyFile`""
        }

        # ── Document → PDF (LibreOffice) ────────────────
        "doc:pdf" {
            $outPath = Join-Path $dir "$base.pdf"
            $pinfo.FileName  = $sofficeCmd
            $pinfo.Arguments = "--headless --convert-to pdf --outdir `"$dir`" `"$inPath`""
        }

        # ── PDF → DOCX (LibreOffice) ────────────────────
        "pdf:docx" {
            $outPath = Join-Path $dir "$base.docx"
            $pinfo.FileName  = $sofficeCmd
            $pinfo.Arguments = '--headless --convert-to docx:"MS Word 2007 XML" --outdir "' + $dir + '" "' + $inPath + '"'
        }
    }

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $pinfo
    $proc.Start() | Out-Null
    $proc.StandardOutput.ReadToEndAsync() | Out-Null
    $proc.StandardError.ReadToEndAsync() | Out-Null

    $listView.Items[$Idx].SubItems[1].Text = "Converting"
    $listView.Items[$Idx].ForeColor = [System.Drawing.Color]::FromArgb(255, 215, 80)

    $script:runningJob = @{
        Process    = $proc
        Idx        = $Idx
        OutPath    = $outPath
        OutDir     = $outDir
        IsPageJob  = $isPageJob
        PyFile     = if ($chosenFormat -eq "img:docx") { Join-Path $env:TEMP "doc_convert_embed_$Idx.py" } else { $null }
        PyFile2    = if ($chosenFormat -match "^pdf:(jpg|png)$") { Join-Path $env:TEMP "doc_convert_pdf2img_$Idx.py" } else { $null }
    }
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 250

$timer.Add_Tick({
    # ── Check running job ──
    if ($script:runningJob -ne $null) {
        $job = $script:runningJob
        if (-not $job.Process.HasExited) { return }

        # Job finished
        $idx     = $job.Idx
        $exitCode = $job.Process.ExitCode
        $success = $false

        if ($job.IsPageJob) {
            $pageFiles = @(Get-ChildItem $job.OutDir -ErrorAction SilentlyContinue)
            $success   = ($exitCode -eq 0) -and ($pageFiles.Count -gt 0)
            if ($success) {
                $listView.Items[$idx].SubItems[1].Text = "Done ($($pageFiles.Count) pages)"
            }
        } else {
            $success = ($exitCode -eq 0) -and ($job.OutPath -ne $null) -and (Test-Path $job.OutPath)
        }

        # Clean up temp python scripts
        if ($job.PyFile -and (Test-Path $job.PyFile)) {
            Remove-Item $job.PyFile -Force -ErrorAction SilentlyContinue
        }
        if ($job.PyFile2 -and (Test-Path $job.PyFile2)) {
            Remove-Item $job.PyFile2 -Force -ErrorAction SilentlyContinue
        }

        if ($success) {
            $script:successCount++
            if (-not $job.IsPageJob) {
                $listView.Items[$idx].SubItems[1].Text = "Done"
            }
            $listView.Items[$idx].ForeColor = [System.Drawing.Color]::FromArgb(80, 210, 120)
        } else {
            $script:failCount++
            $listView.Items[$idx].SubItems[1].Text = "Failed"
            $listView.Items[$idx].ForeColor = [System.Drawing.Color]::FromArgb(255, 80, 70)
        }

        $script:doneCount++
        $progressBar.Value = $script:doneCount
        $script:runningJob = $null
    }

    # ── Start next job ──
    if ($script:runningJob -eq $null -and $script:jobQueue.Count -gt 0) {
        $nextIdx = $script:jobQueue.Dequeue()
        $statusLabel.Text = "Converting: $($files[$nextIdx].Name)"
        Start-ConversionJob -Idx $nextIdx
        return
    }

    # ── All done ──
    if ($script:runningJob -eq $null -and $script:jobQueue.Count -eq 0 -and $script:doneCount -ge $files.Count) {
        $timer.Stop()
        $statusLabel.Text = "Done — $($script:successCount) succeeded, $($script:failCount) failed"
        $statusLabel.ForeColor = if ($script:failCount -gt 0) {
            [System.Drawing.Color]::FromArgb(255, 100, 80)
        } else {
            [System.Drawing.Color]::FromArgb(80, 210, 120)
        }
        $progressBar.Value = $files.Count
        $closeBtn.Visible  = $true
    }
})

$form.Add_Shown({
    $statusLabel.Text = "Converting: $($files[0].Name)"
    Start-ConversionJob -Idx ($script:jobQueue.Dequeue())
    $timer.Start()
})

[System.Windows.Forms.Application]::Run($form)
