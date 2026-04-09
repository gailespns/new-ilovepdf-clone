# PDF Tools — Local PDF Processing Server

A self-hosted PDF toolkit that runs entirely on your machine. No file uploads to third-party servers, no account required, no limits.

Built as a Python/Flask backend with a minimal browser UI.

---

## What it does

### PDF operations
| # | Feature | What it does |
|---|---------|-------------|
| 1 | **Merge PDF** | Combine two or more PDFs into one file, in the order you choose |
| 2 | **Split PDF** | Break a PDF into individual pages, custom page ranges, or download all pages at once |
| 3 | **Compress PDF** | Shrink a PDF's file size — choose between four quality levels from tiny (screen) to print-ready (prepress) |
| 4 | **PDF → Word** | Convert a PDF to an editable .docx file. Text, images, and layout are preserved as closely as possible |
| 5 | **PDF → PowerPoint** | Turn each page of a PDF into a slide in a .pptx presentation |

### Table & document conversions
| # | Feature | What it does |
|---|---------|-------------|
| 6 | **PDF → Excel** | Pull tables out of a PDF and drop them into a formatted .xlsx spreadsheet. Falls back to raw text if no tables are detected |
| 7 | **Word → PDF** | Convert a .doc or .docx file to PDF, preserving fonts, images, tables, and layout |
| 8 | **PowerPoint → PDF** | Convert a .ppt or .pptx presentation to PDF — one PDF page per slide |
| 9 | **Excel → PDF** | Convert a .xls or .xlsx spreadsheet to PDF, including all sheets |

### Bonus
- **PDF Info** — instantly see page count, file size, title, author, creator, and whether the file is password-protected

---

## Quick start

### 1. Install system tools

**Linux (Ubuntu/Debian)**
```bash
sudo apt install ghostscript poppler-utils libreoffice
```

**macOS (Homebrew)**
```bash
brew install ghostscript poppler libreoffice
```

**Windows**
Download and install: [Ghostscript](https://www.ghostscript.com/releases/gsdnld.html) · [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases) · [LibreOffice](https://www.libreoffice.org/download/)

### 2. Install Python packages
```bash
pip install -r requirements.txt
```

### 3. Run
```bash
python app.py
```
Then open **http://localhost:5000** in your browser.

---

## How compression works

When you compress a PDF, the tool tries two methods in order:

1. **Ghostscript** (primary) — re-encodes all image data at a lower DPI and repacks all streams. Works great on PDFs with lots of photos or scanned pages. Typical savings: 50–90%.
2. **pikepdf** (fallback) — removes duplicate objects and repacks streams without touching image quality. Safer for already-lean PDFs.

If the "compressed" file ends up *larger* than the original (can happen with already-optimised PDFs), the original is returned unchanged.

**Quality presets explained:**

| Preset | Image DPI | Best for |
|--------|-----------|---------|
| screen | ~72 dpi | Email attachments, WhatsApp, web sharing |
| ebook | ~150 dpi | Reading on screen, e-readers |
| printer | ~300 dpi | Printing at home or the office |
| prepress | ~300 dpi + colour profiles | Commercial/professional printing |

---

## How Office → PDF works

Word, PowerPoint, and Excel files are converted using **LibreOffice** running in headless (no GUI) mode. LibreOffice's rendering engine faithfully reproduces fonts, images, tables, charts, headers, footers, and formatting. The output is a fully searchable PDF — not a scanned image.

---

## How PDF → Excel works

The tool uses **pdfplumber** to scan each page for table structures by analysing the PDF's vector drawing commands (lines and rectangles that form grid borders). When a table is found:
- Each detected table gets its own Excel sheet
- The first row is styled as a bold header
- Data rows alternate between light grey and white (zebra striping)
- Column widths are auto-sized

If no grid tables are found (common in text-heavy PDFs), all text is exported to a "Raw Text" sheet instead.

---

## Tech stack

| Layer | Technology |
|-------|-----------|
| Web server | Python · Flask |
| PDF read/write | pypdf |
| PDF compression | Ghostscript · pikepdf |
| PDF → Word | pdf2docx (PyMuPDF) |
| PDF → Images | pdf2image (poppler) |
| Images → PPTX | python-pptx |
| PDF → Excel | pdfplumber · openpyxl |
| Office → PDF | LibreOffice headless |

---

## File structure
```
pdf_tools/
├── app.py              # Flask server — all 9 endpoints
├── requirements.txt    # Python dependencies
├── README.md           # This file
└── static/
    └── index.html      # Browser UI
```

## Notes
- Maximum upload size: **50 MB** per file
- All uploaded and converted files are deleted from the server immediately after the download starts — nothing is stored permanently
- The server runs locally by default (`http://localhost:5000`). Do not expose it to the internet without adding authentication
