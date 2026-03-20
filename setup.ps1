# setup.ps1 — Doc Convert Menu installer with dark GUI
# Copies scripts to %LOCALAPPDATA%\DocConvertMenu, checks tools, registers context menu

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$dest = Join-Path $env:LOCALAPPDATA "DocConvertMenu"

# ── UI ──────────────────────────────────────────────────────────────────────
$form = New-Object System.Windows.Forms.Form
$form.Text = "Doc Convert — Setup"
$form.Size = New-Object System.Drawing.Size(480, 240)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.TopMost = $true

$title = New-Object System.Windows.Forms.Label
$title.Text = "Installing Doc Convert Menu..."
$title.Location = New-Object System.Drawing.Point(20, 15)
$title.Size = New-Object System.Drawing.Size(430, 30)
$title.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$title.ForeColor = [System.Drawing.Color]::FromArgb(0, 200, 150)
$form.Controls.Add($title)

$status = New-Object System.Windows.Forms.Label
$status.Text = "Starting..."
$status.Location = New-Object System.Drawing.Point(20, 58)
$status.Size = New-Object System.Drawing.Size(430, 22)
$status.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
$form.Controls.Add($status)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20, 88)
$progress.Size = New-Object System.Drawing.Size(430, 24)
$progress.Style = "Continuous"
$progress.Minimum = 0
$progress.Maximum = 100
$form.Controls.Add($progress)

$detail = New-Object System.Windows.Forms.Label
$detail.Text = ""
$detail.Location = New-Object System.Drawing.Point(20, 122)
$detail.Size = New-Object System.Drawing.Size(430, 20)
$detail.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 110)
$detail.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$form.Controls.Add($detail)

$closeBtn = New-Object System.Windows.Forms.Button
$closeBtn.Text = "Close"
$closeBtn.Location = New-Object System.Drawing.Point(355, 158)
$closeBtn.Size = New-Object System.Drawing.Size(95, 32)
$closeBtn.FlatStyle = "Flat"
$closeBtn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$closeBtn.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$closeBtn.ForeColor = [System.Drawing.Color]::White
$closeBtn.Visible = $false
$closeBtn.Add_Click({ $form.Close() })
$form.Controls.Add($closeBtn)

function Update-Status($msg, $pct, $det) {
    $status.Text = $msg
    $progress.Value = [Math]::Min($pct, 100)
    if ($det) { $detail.Text = $det }
    $form.Refresh()
}

# ── Install logic (runs on a timer so the form can paint first) ──────────────
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 80
$timer.Add_Tick({
    $timer.Stop()
    try {
        # Step 1: Copy files
        Update-Status "Creating install directory..." 10 $dest
        New-Item -ItemType Directory -Path $dest -Force | Out-Null

        $src = $PSScriptRoot
        foreach ($f in @("doc-convert.ps1","launcher.vbs","install.ps1","uninstall.ps1")) {
            $srcFile = Join-Path $src $f
            if (Test-Path $srcFile) {
                Copy-Item -Path $srcFile -Destination $dest -Force
            }
        }
        Update-Status "Files installed..." 25 "Copied to $dest"

        # Step 2: Detect tools
        Update-Status "Checking for ImageMagick..." 35 ""
        $magickOk = [bool](Get-Command magick -ErrorAction SilentlyContinue)
        $detail.Text = if ($magickOk) { "[OK] ImageMagick found" } else { "[!!] ImageMagick not found — image conversions will be unavailable" }
        $form.Refresh()
        Start-Sleep -Milliseconds 300

        Update-Status "Checking for Python..." 45 ""
        $pythonCmd = $null
        foreach ($pc in @("python","python3")) {
            if (Get-Command $pc -ErrorAction SilentlyContinue) { $pythonCmd = $pc; break }
        }
        $detail.Text = if ($pythonCmd) { "[OK] Python found ($pythonCmd)" } else { "[--] Python not found (optional — needed for DOCX/PDF paths)" }
        $form.Refresh()
        Start-Sleep -Milliseconds 300

        # Step 3: Install Python packages
        if ($pythonCmd) {
            Update-Status "Installing Python packages..." 55 "Installing python-docx..."
            & $pythonCmd -m pip install python-docx -q 2>$null
            Update-Status "Installing Python packages..." 65 "Installing PyMuPDF..."
            & $pythonCmd -m pip install pymupdf -q 2>$null
            $detail.Text = "[OK] python-docx + PyMuPDF ready"
            $form.Refresh()
            Start-Sleep -Milliseconds 300
        }

        # Step 4: Register context menu (HKCU — no admin needed)
        Update-Status "Registering context menu entries..." 80 ""

        $launcherPath = Join-Path $dest "launcher.vbs"
        $scriptPath   = Join-Path $dest "doc-convert.ps1"
        $cmd          = "wscript.exe `"$launcherPath`" `"%1`""
        $menuName     = "DocConvert"
        $menuLabel    = "Convert with Doc Convert"
        $menuIcon     = "shell32.dll,71"

        $extensions = @(
            '.jpg','.jpeg','.png','.webp','.bmp',
            '.tiff','.tif','.gif','.heic','.heif',
            '.avif','.jxl','.jp2',
            '.pdf',
            '.docx','.doc','.odt','.rtf',
            '.xlsx','.xls','.ods',
            '.pptx','.ppt','.odp'
        )

        foreach ($ext in $extensions) {
            $regPath = "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell\$menuName"
            if (Test-Path $regPath) { Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue }
            New-Item -Path $regPath -Force | Out-Null
            Set-ItemProperty -Path $regPath -Name "(Default)"        -Value $menuLabel
            Set-ItemProperty -Path $regPath -Name "Icon"             -Value $menuIcon
            Set-ItemProperty -Path $regPath -Name "MultiSelectModel" -Value "Player"
            $cmdKey = "$regPath\command"
            New-Item -Path $cmdKey -Force | Out-Null
            Set-ItemProperty -Path $cmdKey -Name "(Default)" -Value $cmd
        }

        $detail.Text = "Registered $($extensions.Count) file types in HKCU"
        $form.Refresh()

        # Done!
        Update-Status "Done!" 100 ""
        $title.Text = "Installation complete!"
        $title.ForeColor = [System.Drawing.Color]::FromArgb(0, 220, 120)
        $status.Text = "Right-click any image, PDF, or document to convert it."
        $closeBtn.Visible = $true
        $closeBtn.Focus()

    } catch {
        $title.Text = "Installation failed"
        $title.ForeColor = [System.Drawing.Color]::FromArgb(255, 80, 80)
        $status.Text = $_.Exception.Message
        $detail.Text = "Check that you have write access to $dest"
        $closeBtn.Visible = $true
    }
})

$form.Add_Shown({ $timer.Start() })
[System.Windows.Forms.Application]::Run($form)
