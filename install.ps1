# install.ps1 — Register Doc Convert context menu entries (no admin required)
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

Write-Host "Installing Doc Convert Context Menu..." -ForegroundColor Cyan
Write-Host ""

$cmd = "wscript.exe `"$launcherPath`" `"%1`""

# ── File extensions to register ──
$extensions = @(
    # Images
    '.jpg', '.jpeg', '.png', '.webp', '.bmp',
    '.tiff', '.tif', '.gif', '.heic', '.heif',
    '.avif', '.jxl', '.jp2',
    # PDF
    '.pdf',
    # Documents
    '.docx', '.doc', '.odt', '.rtf',
    '.xlsx', '.xls', '.ods',
    '.pptx', '.ppt', '.odp'
)

$menuName  = "DocConvert"
$menuLabel = "Convert with Doc Convert"
$menuIcon  = "shell32.dll,71"   # document/convert icon

foreach ($ext in $extensions) {
    $regPath = "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell\$menuName"

    # Remove old entry if it exists
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
    }

    New-Item -Path $regPath -Force | Out-Null
    Set-ItemProperty -Path $regPath -Name "(Default)"        -Value $menuLabel
    Set-ItemProperty -Path $regPath -Name "Icon"             -Value $menuIcon
    Set-ItemProperty -Path $regPath -Name "MultiSelectModel" -Value "Player"

    $cmdKey = "$regPath\command"
    New-Item -Path $cmdKey -Force | Out-Null
    Set-ItemProperty -Path $cmdKey -Name "(Default)" -Value $cmd

    Write-Host "  Registered: $ext" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "Installation complete!" -ForegroundColor Green
Write-Host ""
Write-Host 'Right-click any supported file to see "Convert with Doc Convert"' -ForegroundColor White
Write-Host ""
Write-Host "Supported file types:" -ForegroundColor White
Write-Host "  Images : jpg jpeg png webp bmp tiff gif heic heif avif jxl jp2" -ForegroundColor Gray
Write-Host "  PDF    : pdf" -ForegroundColor Gray
Write-Host "  Docs   : docx doc odt rtf xlsx xls ods pptx ppt odp" -ForegroundColor Gray
Write-Host ""
Write-Host "Tools detected:" -ForegroundColor White
if (Get-Command magick -ErrorAction SilentlyContinue) {
    Write-Host "  [OK] ImageMagick — image + PDF<->image conversions" -ForegroundColor Green
} else {
    Write-Host "  [--] ImageMagick not found — install from https://imagemagick.org" -ForegroundColor Yellow
}
$softicePaths = @(
    "soffice",
    "$env:ProgramFiles\LibreOffice\program\soffice.exe",
    "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe",
    "$env:LOCALAPPDATA\Programs\LibreOffice\program\soffice.exe"
)
$libreFound = $false
foreach ($p in $softicePaths) {
    if ((Get-Command $p -ErrorAction SilentlyContinue) -or (Test-Path $p)) {
        $libreFound = $true; break
    }
}
if ($libreFound) {
    Write-Host "  [OK] LibreOffice — document conversions" -ForegroundColor Green
} else {
    Write-Host "  [--] LibreOffice not found — install from https://www.libreoffice.org" -ForegroundColor Yellow
    Write-Host "       (needed for DOCX/PPTX/XLSX -> PDF and PDF -> DOCX)" -ForegroundColor DarkGray
}
if (Get-Command python -ErrorAction SilentlyContinue) {
    Write-Host "  [OK] Python — image-to-DOCX embedding" -ForegroundColor Green
} else {
    Write-Host "  [--] Python not found — image->DOCX will be unavailable" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "If the menu doesn't appear immediately, restart Explorer:" -ForegroundColor Yellow
Write-Host "  taskkill /f /im explorer.exe && start explorer.exe" -ForegroundColor DarkGray
Write-Host ""
Read-Host "Press Enter to close"
