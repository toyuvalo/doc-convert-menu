# install.ps1 -- Register Doc Convert context menu + auto-install Python dependencies
# Run: powershell -ExecutionPolicy Bypass -File install.ps1

$launcherPath = Join-Path $PSScriptRoot "launcher.vbs"
$scriptPath   = Join-Path $PSScriptRoot "doc-convert.ps1"

if (-not (Test-Path $scriptPath)) {
    Write-Host "ERROR: doc-convert.ps1 not found in $PSScriptRoot" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $launcherPath)) {
    Write-Host "ERROR: launcher.vbs not found in $PSScriptRoot" -ForegroundColor Red
    exit 1
}

Write-Host "Doc Convert -- Installer" -ForegroundColor Cyan
Write-Host "========================" -ForegroundColor Cyan
Write-Host ""

# ======================================
#  DETECT TOOLS
# ======================================

Write-Host "Checking tools..." -ForegroundColor White

$magickOk  = $false
$libreOk   = $false
$pythonCmd = $null
$pythonOk  = $false

if (Get-Command magick -ErrorAction SilentlyContinue) {
    $magickOk = $true
    Write-Host "  [OK] ImageMagick found" -ForegroundColor Green
} else {
    Write-Host "  [!!] ImageMagick NOT found" -ForegroundColor Yellow
    Write-Host "       Images and PDF->image conversions will not work." -ForegroundColor DarkGray
    Write-Host "       Install: https://imagemagick.org/script/download.php#windows" -ForegroundColor DarkGray
}

$softicePaths = @(
    "soffice",
    "$env:ProgramFiles\LibreOffice\program\soffice.exe",
    "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe",
    "$env:LOCALAPPDATA\Programs\LibreOffice\program\soffice.exe"
)
foreach ($p in $softicePaths) {
    if ((Get-Command $p -ErrorAction SilentlyContinue) -or (Test-Path $p -ErrorAction SilentlyContinue)) {
        $libreOk = $true; break
    }
}
if ($libreOk) {
    Write-Host "  [OK] LibreOffice found" -ForegroundColor Green
} else {
    Write-Host "  [!!] LibreOffice NOT found" -ForegroundColor Yellow
    Write-Host "       Needed for DOCX/XLSX/PPTX->PDF and PDF->DOCX." -ForegroundColor DarkGray

    # Try to install via winget (available on Win10 1809+ and Win11)
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        Write-Host "  [..] Installing LibreOffice via winget..." -ForegroundColor Yellow
        Write-Host "       (This is a ~350 MB download, may take a few minutes)" -ForegroundColor DarkGray
        & winget install --id TheDocumentFoundation.LibreOffice --silent --accept-package-agreements --accept-source-agreements 2>&1 | Out-Null
        # Re-check after install
        foreach ($p in $softicePaths) {
            if ((Get-Command $p -ErrorAction SilentlyContinue) -or (Test-Path $p -ErrorAction SilentlyContinue)) {
                $libreOk = $true; break
            }
        }
        if ($libreOk) {
            Write-Host "  [OK] LibreOffice installed successfully" -ForegroundColor Green
        } else {
            Write-Host "  [!!] LibreOffice install may need a restart to take effect." -ForegroundColor Yellow
            Write-Host "       If DOCX->PDF still doesn't work after restart, install manually:" -ForegroundColor DarkGray
            Write-Host "       https://www.libreoffice.org/download/" -ForegroundColor DarkGray
        }
    } else {
        Write-Host "       winget not available. Install manually:" -ForegroundColor DarkGray
        Write-Host "       https://www.libreoffice.org/download/" -ForegroundColor DarkGray
    }
}

foreach ($pc in @("python", "python3")) {
    if (Get-Command $pc -ErrorAction SilentlyContinue) { $pythonCmd = $pc; $pythonOk = $true; break }
}
if ($pythonOk) {
    Write-Host "  [OK] Python found ($pythonCmd)" -ForegroundColor Green
} else {
    Write-Host "  [--] Python NOT found (optional)" -ForegroundColor DarkGray
    Write-Host "       Needed for image->DOCX and PDF->image conversions." -ForegroundColor DarkGray
    Write-Host "       Install: https://www.python.org/downloads/" -ForegroundColor DarkGray
}

Write-Host ""

# ======================================
#  AUTO-INSTALL PYTHON PACKAGES
# ======================================

if ($pythonOk) {
    Write-Host "Checking Python packages..." -ForegroundColor White

    # Check / install python-docx
    $hasDocx = & $pythonCmd -c "import docx; print('ok')" 2>$null
    if ($hasDocx -eq "ok") {
        Write-Host "  [OK] python-docx already installed" -ForegroundColor Green
    } else {
        Write-Host "  [..] Installing python-docx..." -ForegroundColor Yellow -NoNewline
        & $pythonCmd -m pip install python-docx -q 2>$null
        $hasDocx2 = & $pythonCmd -c "import docx; print('ok')" 2>$null
        if ($hasDocx2 -eq "ok") {
            Write-Host " done" -ForegroundColor Green
        } else {
            Write-Host " FAILED (will auto-retry on first use)" -ForegroundColor Red
        }
    }

    # Check / install pymupdf
    $hasFitz = & $pythonCmd -c "import fitz; print('ok')" 2>$null
    if ($hasFitz -eq "ok") {
        Write-Host "  [OK] PyMuPDF already installed" -ForegroundColor Green
    } else {
        Write-Host "  [..] Installing PyMuPDF (for PDF->image)..." -ForegroundColor Yellow -NoNewline
        & $pythonCmd -m pip install pymupdf -q 2>$null
        $hasFitz2 = & $pythonCmd -c "import fitz; print('ok')" 2>$null
        if ($hasFitz2 -eq "ok") {
            Write-Host " done" -ForegroundColor Green
        } else {
            Write-Host " FAILED (will auto-retry on first use)" -ForegroundColor Red
        }
    }

    Write-Host ""
}

if (-not $magickOk -and -not $libreOk) {
    Write-Host "WARNING: No conversion tools found. Install ImageMagick and/or LibreOffice" -ForegroundColor Red
    Write-Host "         before using Doc Convert." -ForegroundColor Red
    Write-Host ""
}

# ======================================
#  REGISTER CONTEXT MENU
# ======================================

Write-Host "Registering context menu entries..." -ForegroundColor White

$extensions = @(
    '.jpg', '.jpeg', '.png', '.webp', '.bmp',
    '.tiff', '.tif', '.gif', '.heic', '.heif',
    '.avif', '.jxl', '.jp2',
    '.pdf',
    '.docx', '.doc', '.odt', '.rtf',
    '.xlsx', '.xls', '.ods',
    '.pptx', '.ppt', '.odp'
)

$menuName  = "DocConvert"
$menuLabel = "Convert with Doc Convert"
$icoPath  = Join-Path $PSScriptRoot "doc-convert.ico"
$menuIcon  = if (Test-Path $icoPath) { "$icoPath,0" } else { "shell32.dll,71" }
$cmd       = "wscript.exe `"$launcherPath`" `"%1`""

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

Write-Host "  Registered $($extensions.Count) file types" -ForegroundColor Green
Write-Host ""
Write-Host "Installation complete!" -ForegroundColor Green
Write-Host ""
Write-Host "Right-click any of these to see 'Convert with Doc Convert':" -ForegroundColor White
Write-Host "  Images: jpg jpeg png webp bmp tiff gif heic heif avif jxl jp2" -ForegroundColor Gray
Write-Host "  PDF:    pdf" -ForegroundColor Gray
Write-Host "  Docs:   docx doc odt rtf xlsx xls ods pptx ppt odp" -ForegroundColor Gray
Write-Host ""
Write-Host "If the menu doesn't appear, restart Explorer:" -ForegroundColor Yellow
Write-Host "  taskkill /f /im explorer.exe && start explorer.exe" -ForegroundColor DarkGray
Write-Host ""
Read-Host "Press Enter to close"
