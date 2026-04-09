<<<<<<< HEAD
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
=======
# PDF Tools Backend — iLovePDF Clone (Top 5 Features)

A backend-focused Python/Flask clone of iLovePDF's first five tools.

## Features

| # | Feature | Endpoint | Key Library |
|---|---------|----------|-------------|
| 1 | Merge PDF | `POST /api/merge` | pypdf `PdfWriter` |
| 2 | Split PDF | `POST /api/split` | pypdf `PdfReader` |
| 3 | Compress PDF | `POST /api/compress` | Ghostscript → pikepdf fallback |
| 4 | PDF → Word | `POST /api/pdf-to-word` | pdf2docx `Converter` |
| 5 | PDF → PowerPoint | `POST /api/pdf-to-pptx` | pdf2image + python-pptx |

Plus:
- `POST /api/info` — PDF metadata inspector
- `GET  /api/health` — service health check (ghostscript, pdftoppm)

---

## Setup

### System dependencies
```bash
# Debian / Ubuntu
sudo apt install ghostscript poppler-utils

# macOS (Homebrew)
brew install ghostscript poppler
```

### Python dependencies
>>>>>>> b37428719113e6ba4469ce50fb981c0a35b4f255
```bash
pip install -r requirements.txt
```

<<<<<<< HEAD
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
=======
### Run
```bash
python app.py
# → http://localhost:5000
```

---

## API Reference

### 1 · Merge PDF
```
POST /api/merge
Content-Type: multipart/form-data

files[]  — two or more PDF files (required)
```
Returns: `merged.pdf`

**Algorithm:**  
Opens each uploaded PDF with `PdfReader`, iterates pages, and appends them sequentially to a shared `PdfWriter`. The output is a single PDF containing all pages in upload order.

---

### 2 · Split PDF
```
POST /api/split
Content-Type: multipart/form-data

file    — PDF file (required)
ranges  — page range string (optional, default: "each")
          "each"       → one PDF per page
          "all"        → return whole document
          "1-3,5,7-9"  → custom groups; each group = one PDF
```
Returns: single PDF (one group) or `split_pages.zip` (multiple groups)

**Algorithm:**  
`_parse_page_ranges()` tokenises the range string into lists of 0-based page indices. For each group a new `PdfWriter` is populated and written. Multi-group output is bundled into a ZIP archive via Python's `zipfile` module.

---

### 3 · Compress PDF
```
POST /api/compress
Content-Type: multipart/form-data

file     — PDF file (required)
quality  — "screen" | "ebook" (default) | "printer" | "prepress"
```
Returns: `compressed_<filename>.pdf`

**Algorithm:**  
Primary: invokes **Ghostscript** (`gs -sDEVICE=pdfwrite -dPDFSETTINGS=...`) which down-samples embedded images and re-encodes all streams, often achieving 50-90 % reduction on image-heavy PDFs.  
Fallback: **pikepdf** lossless compression (object deduplication + stream repack) when Ghostscript is not installed.  
Guard: if the output is larger than the input (already-optimised PDFs), the original is returned unchanged.

**Quality presets:**
| Preset | GS flag | Image DPI | Use case |
|--------|---------|-----------|---------|
| screen | `/screen` | ~72 | Email / web display |
| ebook | `/ebook` | ~150 | E-readers, sharing |
| printer | `/printer` | ~300 | Desktop printing |
| prepress | `/prepress` | ~300 + colour | Commercial print |

---

### 4 · PDF → Word
```
POST /api/pdf-to-word
Content-Type: multipart/form-data

file        — PDF file (required)
start_page  — first page to convert, 1-based (optional, default: 1)
end_page    — last page to convert,  1-based (optional, default: last)
```
Returns: `<filename>.docx`

**Algorithm:**  
`pdf2docx.Converter` internally uses **PyMuPDF (fitz)** to render pages at high DPI, then applies a rule-based layout analyser that detects text blocks, fonts, images, tables, and columns. Output is a `.docx` file built with **python-docx** whose XML mirrors the detected layout. Page-range parameters map to pdf2docx's 0-based `start`/`end` arguments.

---

### 5 · PDF → PowerPoint
```
POST /api/pdf-to-pptx
Content-Type: multipart/form-data

file  — PDF file (required)
dpi   — render resolution 72-300 (optional, default: 150)
```
Returns: `<filename>.pptx`

**Algorithm:**  
1. **Rasterise** — `pdf2image.convert_from_path()` calls **poppler's pdftoppm** to render every page as a PIL Image at the requested DPI.  
2. **Save PNGs** — each PIL Image is written to a temporary PNG file.  
3. **Build PPTX** — a `python-pptx` `Presentation` is created with slide dimensions 10 × 7.5 in (standard 4:3). For every PNG a blank slide is added and the image is inserted as a full-slide picture.  
4. **Cleanup** — temp PNGs are deleted after the PPTX is written.

Higher DPI → crisper text in slides but larger file size and slower conversion.

---

## Project Structure
```
pdf_tools/
├── app.py              # Flask backend (all 5 features + utilities)
├── requirements.txt    # Python dependencies
├── README.md           # This file
└── static/
    └── index.html      # Minimal UI (monospace, no frameworks)
```

## Notes
- Max upload size: 50 MB (configurable via `MAX_UPLOAD_MB`)
- All temp files are deleted after each request
- PDF→PPTX requires `poppler-utils` (for `pdftoppm`)
- Compress uses Ghostscript when available; pikepdf otherwise
>>>>>>> b37428719113e6ba4469ce50fb981c0a35b4f255
