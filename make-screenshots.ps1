# make-screenshots.ps1 -- Generate README screenshots for DocConvertMenu
# Builds mock forms identical to the real app and captures ONLY the form bounds.
# No full-screen capture, no user interaction, no OOM.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$here   = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$outDir = Join-Path $here "screenshots"
New-Item -Path $outDir -ItemType Directory -Force | Out-Null

function Capture-Form {
    param($Form, $Path)
    $Form.Show()
    $Form.BringToFront()
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 400
    [System.Windows.Forms.Application]::DoEvents()
    $loc = $Form.Location
    $bmp = New-Object System.Drawing.Bitmap($Form.Width, $Form.Height)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($loc.X, $loc.Y, 0, 0, [System.Drawing.Size]::new($Form.Width, $Form.Height))
    $g.Dispose()
    $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $Form.Close()
    $Form.Dispose()
    [System.Windows.Forms.Application]::DoEvents()
}

# -----------------------------------------------------------------------
#  SHARED COLORS  (match doc-convert.ps1 exactly)
# -----------------------------------------------------------------------
$bg        = [System.Drawing.Color]::FromArgb(28, 28, 30)
$btnBg     = [System.Drawing.Color]::FromArgb(50, 50, 52)
$hintBg    = [System.Drawing.Color]::FromArgb(45, 38, 28)
$hintFg    = [System.Drawing.Color]::FromArgb(200, 150, 60)
$hintBdr   = [System.Drawing.Color]::FromArgb(100, 70, 30)
$shrinkBg  = [System.Drawing.Color]::FromArgb(44, 32, 12)
$shrinkFg  = [System.Drawing.Color]::FromArgb(255, 160, 60)
$shrinkBdr = [System.Drawing.Color]::FromArgb(120, 80, 20)
$teal      = [System.Drawing.Color]::FromArgb(80, 200, 160)
$dimGray   = [System.Drawing.Color]::FromArgb(120, 120, 120)
$white     = [System.Drawing.Color]::White
$sepColor  = [System.Drawing.Color]::FromArgb(55, 55, 57)
$darkBg2   = [System.Drawing.Color]::FromArgb(40, 40, 42)
$green     = [System.Drawing.Color]::FromArgb(80, 210, 120)
$yellow    = [System.Drawing.Color]::FromArgb(255, 215, 80)
$grayDim   = [System.Drawing.Color]::FromArgb(130, 130, 130)

# -----------------------------------------------------------------------
#  SCREENSHOT 1: PICKER  (Images + PDF->Image + hint + Shrink button)
# -----------------------------------------------------------------------
function New-PickerForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Doc Convert"
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox = $false
    $form.BackColor = $bg
    $form.ForeColor = $white
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.TopMost = $true

    $hdr = New-Object System.Windows.Forms.Label
    $hdr.Text = "Convert to..."
    $hdr.Location = New-Object System.Drawing.Point(20, 14)
    $hdr.Size = New-Object System.Drawing.Size(260, 28)
    $hdr.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $hdr.ForeColor = $teal
    $form.Controls.Add($hdr)

    $y = 50

    # Section label
    $sl1 = New-Object System.Windows.Forms.Label; $sl1.Text = "Images  (2 files)"
    $sl1.Location = New-Object System.Drawing.Point(20,$y); $sl1.Size = New-Object System.Drawing.Size(260,18)
    $sl1.ForeColor = $dimGray; $sl1.Font = New-Object System.Drawing.Font("Segoe UI",8); $form.Controls.Add($sl1); $y += 20

    foreach ($t in @("PNG","WebP","BMP","TIFF","GIF","PDF  (single file)","DOCX  (embed image)")) {
        $b = New-Object System.Windows.Forms.Button; $b.Text = $t
        $b.Location = New-Object System.Drawing.Point(20,$y); $b.Size = New-Object System.Drawing.Size(260,30)
        $b.FlatStyle = "Flat"; $b.FlatAppearance.BorderSize = 0
        $b.BackColor = $btnBg; $b.ForeColor = $white; $b.TextAlign = "MiddleLeft"
        $b.Padding = New-Object System.Windows.Forms.Padding(10,0,0,0); $form.Controls.Add($b); $y += 34
    }
    $y += 6

    $sl2 = New-Object System.Windows.Forms.Label; $sl2.Text = "PDF -> Image  (1 file)"
    $sl2.Location = New-Object System.Drawing.Point(20,$y); $sl2.Size = New-Object System.Drawing.Size(260,18)
    $sl2.ForeColor = $dimGray; $sl2.Font = New-Object System.Drawing.Font("Segoe UI",8); $form.Controls.Add($sl2); $y += 20

    foreach ($t in @("JPG  (one file per page)","PNG  (one file per page)")) {
        $b = New-Object System.Windows.Forms.Button; $b.Text = $t
        $b.Location = New-Object System.Drawing.Point(20,$y); $b.Size = New-Object System.Drawing.Size(260,30)
        $b.FlatStyle = "Flat"; $b.FlatAppearance.BorderSize = 0
        $b.BackColor = $btnBg; $b.ForeColor = $white; $b.TextAlign = "MiddleLeft"
        $b.Padding = New-Object System.Windows.Forms.Padding(10,0,0,0); $form.Controls.Add($b); $y += 34
    }
    $y += 6

    $sl3 = New-Object System.Windows.Forms.Label; $sl3.Text = "Documents  (LibreOffice required)"
    $sl3.Location = New-Object System.Drawing.Point(20,$y); $sl3.Size = New-Object System.Drawing.Size(260,18)
    $sl3.ForeColor = $dimGray; $sl3.Font = New-Object System.Drawing.Font("Segoe UI",8); $form.Controls.Add($sl3); $y += 20

    foreach ($t in @("DOCX / XLSX / PPTX -> PDF  [install LibreOffice]","PDF -> DOCX  [install LibreOffice]")) {
        $b = New-Object System.Windows.Forms.Button; $b.Text = $t
        $b.Location = New-Object System.Drawing.Point(20,$y); $b.Size = New-Object System.Drawing.Size(260,30)
        $b.FlatStyle = "Flat"; $b.FlatAppearance.BorderSize = 1; $b.FlatAppearance.BorderColor = $hintBdr
        $b.BackColor = $hintBg; $b.ForeColor = $hintFg; $b.TextAlign = "MiddleLeft"
        $b.Padding = New-Object System.Windows.Forms.Padding(10,0,0,0); $form.Controls.Add($b); $y += 34
    }
    $y += 12

    $sep = New-Object System.Windows.Forms.Panel
    $sep.Location = New-Object System.Drawing.Point(20,$y); $sep.Size = New-Object System.Drawing.Size(260,1)
    $sep.BackColor = $sepColor; $form.Controls.Add($sep); $y += 8

    $shrink = New-Object System.Windows.Forms.Button; $shrink.Text = "Shrink..."
    $shrink.Location = New-Object System.Drawing.Point(20,$y); $shrink.Size = New-Object System.Drawing.Size(260,28)
    $shrink.FlatStyle = "Flat"; $shrink.FlatAppearance.BorderSize = 1; $shrink.FlatAppearance.BorderColor = $shrinkBdr
    $shrink.BackColor = $shrinkBg; $shrink.ForeColor = $shrinkFg; $shrink.TextAlign = "MiddleLeft"
    $shrink.Padding = New-Object System.Windows.Forms.Padding(8,0,0,0)
    $shrink.Font = New-Object System.Drawing.Font("Segoe UI",9); $form.Controls.Add($shrink); $y += 34

    $form.ClientSize = New-Object System.Drawing.Size(300, ($y + 8))
    return $form
}

# -----------------------------------------------------------------------
#  SCREENSHOTS 2 & 3: PROGRESS + DONE
# -----------------------------------------------------------------------
function New-ProgressForm {
    param([bool]$Done = $false)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Doc Convert"
    $form.Size = New-Object System.Drawing.Size(570, 430)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox = $false
    $form.BackColor = $bg
    $form.ForeColor = $white
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.TopMost = $true

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Converting 3 file(s)  --  Image -> JPG"
    $title.Location = New-Object System.Drawing.Point(20, 16)
    $title.Size = New-Object System.Drawing.Size(515, 26)
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = $teal
    $form.Controls.Add($title)

    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Location = New-Object System.Drawing.Point(20, 52)
    $pb.Size = New-Object System.Drawing.Size(515, 24)
    $pb.Style = "Continuous"
    $pb.Minimum = 0; $pb.Maximum = 3
    $pb.Value = if ($Done) { 3 } else { 1 }
    $form.Controls.Add($pb)

    $status = New-Object System.Windows.Forms.Label
    $status.Location = New-Object System.Drawing.Point(20, 84)
    $status.Size = New-Object System.Drawing.Size(515, 20)
    if ($Done) {
        $status.Text = "Done -- 3 succeeded, 0 failed"
        $status.ForeColor = $green
    } else {
        $status.Text = "Converting: photo_sunset.jpg"
        $status.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
    }
    $form.Controls.Add($status)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location = New-Object System.Drawing.Point(20, 112)
    $lv.Size = New-Object System.Drawing.Size(515, 230)
    $lv.View = "Details"
    $lv.FullRowSelect = $true
    $lv.BackColor = $darkBg2
    $lv.ForeColor = $white
    $lv.Font = New-Object System.Drawing.Font("Consolas", 9)
    $lv.BorderStyle = "None"
    $lv.HeaderStyle = "Nonclickable"
    $lv.GridLines = $false
    $lv.Columns.Add("File", 330) | Out-Null
    $lv.Columns.Add("Status", 90) | Out-Null
    $lv.Columns.Add("Size", 80) | Out-Null

    $rows = if ($Done) {
        @(
            @{ N="photo_beach.jpg";   St="Done";       Sz="2.4 MB"; C=$green  }
            @{ N="photo_sunset.jpg";  St="Done";       Sz="3.1 MB"; C=$green  }
            @{ N="background.png";    St="Done";       Sz="4.1 MB"; C=$green  }
        )
    } else {
        @(
            @{ N="photo_beach.jpg";   St="Done";       Sz="2.4 MB"; C=$green  }
            @{ N="photo_sunset.jpg";  St="Converting"; Sz="3.1 MB"; C=$yellow }
            @{ N="background.png";    St="Queued";     Sz="4.1 MB"; C=$grayDim}
        )
    }
    foreach ($row in $rows) {
        $item = New-Object System.Windows.Forms.ListViewItem($row.N)
        $item.SubItems.Add($row.St) | Out-Null
        $item.SubItems.Add($row.Sz) | Out-Null
        $item.ForeColor = $row.C
        $lv.Items.Add($item) | Out-Null
    }
    $form.Controls.Add($lv)

    if ($Done) {
        $cb = New-Object System.Windows.Forms.Button
        $cb.Text = "Close"
        $cb.Location = New-Object System.Drawing.Point(430, 355)
        $cb.Size = New-Object System.Drawing.Size(105, 32)
        $cb.FlatStyle = "Flat"
        $cb.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
        $cb.BackColor = [System.Drawing.Color]::FromArgb(55, 55, 57)
        $cb.ForeColor = $white
        $form.Controls.Add($cb)
    }

    return $form
}

# -----------------------------------------------------------------------
#  RUN
# -----------------------------------------------------------------------
Write-Host "Generating screenshots..." -ForegroundColor Yellow

Capture-Form -Form (New-PickerForm) -Path (Join-Path $outDir "picker.png")
Write-Host "  [1/3] picker.png" -ForegroundColor Green

Capture-Form -Form (New-ProgressForm -Done $false) -Path (Join-Path $outDir "progress.png")
Write-Host "  [2/3] progress.png" -ForegroundColor Green

Capture-Form -Form (New-ProgressForm -Done $true) -Path (Join-Path $outDir "done.png")
Write-Host "  [3/3] done.png" -ForegroundColor Green

Write-Host ""
Write-Host "Done. Screenshots saved to: $outDir" -ForegroundColor Green
