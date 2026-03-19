# Doc Convert — Windows Right-Click Context Menu

Convert images and documents directly from Windows Explorer. Right-click any supported file → **Convert with Doc Convert**.

## Supported conversions

| Input | Output | Tool required |
|-------|--------|---------------|
| JPG, PNG, WebP, BMP, TIFF, GIF, HEIC, AVIF, JP2, JXL | JPG / PNG / WebP / BMP / TIFF / GIF | ImageMagick |
| Any image | PDF | ImageMagick |
| PDF | JPG / PNG (one file per page) | ImageMagick |
| Image | DOCX (image embedded in document) | Python + python-docx (auto-installed) |
| DOCX, DOC, ODT, RTF, XLSX, XLS, ODS, PPTX, PPT, ODP | PDF | LibreOffice |
| PDF | DOCX | LibreOffice |

## Installation

### 1. Install required tools

**ImageMagick** (for all image + PDF↔image conversions):
- Download from https://imagemagick.org/script/download.php#windows
- Use the installer — make sure "Add to PATH" is checked

**LibreOffice** (for DOCX/XLSX/PPTX → PDF and PDF → DOCX):
- Download from https://www.libreoffice.org/download/
- Optional — image conversions work without it

**Python** (for image → DOCX embedding):
- Download from https://www.python.org/downloads/
- `python-docx` is auto-installed on first use

### 2. Register the context menu

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
```

No admin rights required — all registry entries go into `HKCU`.

### 3. Use it

Right-click any supported file in Explorer → **Convert with Doc Convert** → pick your target format.

Multi-select works: select multiple files, right-click any one → all selected files are converted.

## Uninstall

```powershell
powershell -ExecutionPolicy Bypass -File uninstall.ps1
```

## Notes

- Output files are saved in the same folder as the input
- PDF → image creates a `<filename>_pages/` subfolder (one image per page)
- If the source and output extension are the same, `_converted` is appended
- The format picker only shows options valid for the file types you selected

## Similar tools

- [ffmpeg-context-menu](https://github.com/toyuvalo/ffmpeg-context-menu) — Same idea for audio/video conversion
