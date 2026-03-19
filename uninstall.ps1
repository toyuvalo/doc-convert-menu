# uninstall.ps1 — Remove Doc Convert context menu entries
# Run: powershell -ExecutionPolicy Bypass -File uninstall.ps1

Write-Host "Uninstalling Doc Convert Context Menu..." -ForegroundColor Cyan
Write-Host ""

$extensions = @(
    '.jpg', '.jpeg', '.png', '.webp', '.bmp',
    '.tiff', '.tif', '.gif', '.heic', '.heif',
    '.avif', '.jxl', '.jp2',
    '.pdf',
    '.docx', '.doc', '.odt', '.rtf',
    '.xlsx', '.xls', '.ods',
    '.pptx', '.ppt', '.odp'
)

$menuName = "DocConvert"

foreach ($ext in $extensions) {
    $regPath = "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell\$menuName"
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed: $ext" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "Uninstall complete." -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to close"
