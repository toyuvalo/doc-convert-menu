# Doc Convert Menu

Windows right-click context menu for file format conversion. Select any image, PDF, or document in Explorer → **Convert with Doc Convert** → pick a format. No admin rights, no dedicated app to open.

**[Project page →](https://webdev.dvlce.ca/doc-convert)**

## Install

### One-click installer (recommended)

Download **[DocConvertSetup.exe](https://github.com/toyuvalo/doc-convert-menu/releases/latest)** and run it.

The installer:
- Copies scripts to `%LOCALAPPDATA%\DocConvertMenu\`
- Auto-installs `python-docx` and `PyMuPDF` via pip (if Python is present)
- Registers all file type context menu entries in HKCU (no admin required)

### Manual install

```powershell
git clone https://github.com/toyuvalo/doc-convert-menu
cd doc-convert-menu
powershell -ExecutionPolicy Bypass -File install.ps1
```

## Supported conversions

| Input | Output | Requires |
|-------|--------|----------|
| JPG / PNG / WebP / BMP / TIFF / GIF / HEIC / AVIF / JP2 | Any image format | ImageMagick |
| Any image | PDF | ImageMagick |
| PDF | JPG / PNG (one file per page) | PyMuPDF (auto-installed) |
| Any image | DOCX (embedded) | python-docx (auto-installed) |
| DOCX / XLSX / PPTX / ODT | PDF | LibreOffice (optional) |
| PDF | DOCX | LibreOffice (optional) |

## Requirements

- Windows 10/11
- [ImageMagick 7+](https://imagemagick.org/script/download.php#windows) — for image and PDF↔image conversions
- Python 3.8+ — optional, needed for DOCX embedding and PDF→image
- LibreOffice — optional, needed for DOCX/XLSX/PPTX→PDF

## Features

- **No admin rights** — all registry entries written to HKCU
- **Multi-file selection** — select any number of files, right-click once, convert all of them
- **Auto dependency install** — `python-docx` and `PyMuPDF` pip-installed automatically on first use
- **Multi-select via VBScript COM** — bypasses the single-file limitation of standard shell extensions

## Build the installer

Requires Windows (IExpress ships with the OS):

```cmd
build.cmd
```

Produces `DocConvertSetup.exe` in the repo root.

## Uninstall

```powershell
powershell -ExecutionPolicy Bypass -File "%LOCALAPPDATA%\DocConvertMenu\uninstall.ps1"
```

## Related

- [ffmpeg-context-menu](https://github.com/toyuvalo/ffmpeg-context-menu) — same idea for audio/video conversion with FFmpeg
- [webdev.dvlce.ca/doc-convert](https://webdev.dvlce.ca/doc-convert) — project page

## License

MIT with [Commons Clause](https://commonsclause.com/) — free to use, modify, and share. Commercial resale not permitted.
