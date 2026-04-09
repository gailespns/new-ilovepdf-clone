"""
PDF Tools Backend - iLovePDF Clone (Top 5 Features)
=====================================================
Features implemented:
  1. Merge PDF       - Combine multiple PDFs into one
  2. Split PDF       - Split PDF into individual pages or page ranges
  3. Compress PDF    - Reduce PDF file size via Ghostscript / pikepdf
  4. PDF to Word     - Convert PDF to .docx (via pdf2docx)
  5. PDF to PowerPoint - Convert PDF pages to .pptx (via pdf2image + python-pptx)
"""

import os
import io
import uuid
import zipfile
import shutil
import subprocess
import tempfile
import logging
from pathlib import Path
from flask import Flask, request, jsonify, send_file, after_this_request

# ── PDF libraries ────────────────────────────────────────────────────────────
from pypdf import PdfReader, PdfWriter
import pikepdf
from pdf2docx import Converter
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches, Pt

# ── App setup ────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(message)s")
log = logging.getLogger(__name__)

app = Flask(__name__, static_folder="static", static_url_path="")

UPLOAD_DIR = Path(tempfile.gettempdir()) / "pdf_tools_uploads"
OUTPUT_DIR = Path(tempfile.gettempdir()) / "pdf_tools_outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

MAX_UPLOAD_MB = 50
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

ALLOWED_EXT = {".pdf"}


# ── Helpers ──────────────────────────────────────────────────────────────────

def _uid() -> str:
    return uuid.uuid4().hex


def _save_upload(file_storage, suffix=".pdf") -> Path:
    """Save a Werkzeug FileStorage to a temp file and return its Path."""
    dest = UPLOAD_DIR / f"{_uid()}{suffix}"
    file_storage.save(str(dest))
    log.info("Saved upload → %s  (%.1f KB)", dest.name, dest.stat().st_size / 1024)
    return dest


def _out(suffix: str) -> Path:
    """Return a fresh output path."""
    return OUTPUT_DIR / f"{_uid()}{suffix}"


def _send(path: Path, download_name: str, mimetype: str = "application/octet-stream"):
    """Stream a file to the client and schedule its deletion after the request."""
    @after_this_request
    def _cleanup(response):
        try:
            path.unlink(missing_ok=True)
        except Exception:
            pass
        return response

    return send_file(
        str(path),
        as_attachment=True,
        download_name=download_name,
        mimetype=mimetype,
    )


def _pdf_page_count(path: Path) -> int:
    reader = PdfReader(str(path))
    return len(reader.pages)


def _parse_page_ranges(range_str: str, total_pages: int) -> list[list[int]]:
    """
    Parse a comma-separated range string like "1-3,5,7-9" into lists of
    0-based page indices.  Returns a list of groups; each group becomes one
    output PDF.

    Examples
    --------
    "1-3,5"          → [[0,1,2], [4]]
    "all"            → [[0,1,...,total_pages-1]]
    "" / "each"      → [[0],[1],...] (one PDF per page)
    """
    range_str = range_str.strip().lower()
    if not range_str or range_str == "each":
        return [[i] for i in range(total_pages)]
    if range_str == "all":
        return [list(range(total_pages))]

    groups: list[list[int]] = []
    for part in range_str.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            lo, hi = part.split("-", 1)
            lo, hi = int(lo.strip()) - 1, int(hi.strip()) - 1
            lo = max(0, lo)
            hi = min(total_pages - 1, hi)
            if lo <= hi:
                groups.append(list(range(lo, hi + 1)))
        else:
            idx = int(part) - 1
            if 0 <= idx < total_pages:
                groups.append([idx])
    return groups


# ═══════════════════════════════════════════════════════════════════════════
# FEATURE 1 — MERGE PDF
# ═══════════════════════════════════════════════════════════════════════════

@app.route("/api/merge", methods=["POST"])
def merge_pdf():
    """
    Endpoint: POST /api/merge
    Body    : multipart/form-data, field name 'files[]'  (2+ PDF files)
    Returns : merged.pdf

    Algorithm
    ---------
    1. Validate that at least 2 PDFs are uploaded.
    2. Save each upload to a temp file preserving the client filename.
    3. Iterate uploaded PDFs in order; for each, open with PdfReader and
       append every page to a shared PdfWriter.
    4. Write the writer to a temp output file and stream it back.
    5. Clean up all temp input files.
    """
    files = request.files.getlist("files[]")
    if len(files) < 2:
        return jsonify(error="Please upload at least 2 PDF files."), 400

    saved_inputs: list[Path] = []
    try:
        writer = PdfWriter()

        for f in files:
            if Path(f.filename).suffix.lower() not in ALLOWED_EXT:
                return jsonify(error=f"'{f.filename}' is not a PDF."), 400
            p = _save_upload(f)
            saved_inputs.append(p)

            reader = PdfReader(str(p))
            log.info("Merging '%s'  → %d pages", f.filename, len(reader.pages))
            for page in reader.pages:
                writer.add_page(page)

        out = _out(".pdf")
        with open(out, "wb") as fh:
            writer.write(fh)

        total = sum(_pdf_page_count(p) for p in saved_inputs)
        log.info("Merge complete: %d files, %d pages → %s (%.1f KB)",
                 len(files), total, out.name, out.stat().st_size / 1024)

        return _send(out, "merged.pdf", "application/pdf")

    finally:
        for p in saved_inputs:
            p.unlink(missing_ok=True)


# ═══════════════════════════════════════════════════════════════════════════
# FEATURE 2 — SPLIT PDF
# ═══════════════════════════════════════════════════════════════════════════

@app.route("/api/split", methods=["POST"])
def split_pdf():
    """
    Endpoint: POST /api/split
    Body    : multipart/form-data
      file   – PDF to split
      ranges – (optional) comma-separated page ranges, e.g. "1-3,5,7-9"
               "each" or empty → one PDF per page  (default)
               "all"           → returns the whole PDF as a single output

    Returns : zip archive containing the split PDFs when more than one
              output is produced, otherwise a single PDF.

    Algorithm
    ---------
    1. Parse the 'ranges' field into groups of 0-based page indices using
       _parse_page_ranges().
    2. For every group, open a fresh PdfWriter, add the requested pages,
       and write to a temp file.
    3. If only one group → stream it directly.
       If multiple groups → bundle them into a zip and stream that.
    """
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file provided."), 400
    if Path(f.filename).suffix.lower() not in ALLOWED_EXT:
        return jsonify(error="File must be a PDF."), 400

    ranges_str = request.form.get("ranges", "each")
    input_path = _save_upload(f)
    output_paths: list[Path] = []

    try:
        reader = PdfReader(str(input_path))
        total = len(reader.pages)
        log.info("Split: '%s' has %d pages, ranges='%s'", f.filename, total, ranges_str)

        groups = _parse_page_ranges(ranges_str, total)
        if not groups:
            return jsonify(error="No valid page ranges found."), 400

        for idx, page_indices in enumerate(groups):
            writer = PdfWriter()
            for pi in page_indices:
                writer.add_page(reader.pages[pi])

            out = _out(".pdf")
            with open(out, "wb") as fh:
                writer.write(fh)
            output_paths.append(out)
            log.info("  Part %d/%d: pages %s → %s (%.1f KB)",
                     idx + 1, len(groups),
                     [p + 1 for p in page_indices],
                     out.name, out.stat().st_size / 1024)

        if len(output_paths) == 1:
            out = output_paths[0]
            stem = Path(f.filename).stem
            return _send(out, f"{stem}_split.pdf", "application/pdf")

        # Pack into a zip
        zip_path = _out(".zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, p in enumerate(output_paths, 1):
                zf.write(p, f"page_{i:03d}.pdf")

        return _send(zip_path, "split_pages.zip", "application/zip")

        @after_this_request
        def _cleanup(response):
            for p in output_paths:
                try:
                    p.unlink(missing_ok=True)
                except Exception:
                    pass

            try:
                zip_path.unlink(missing_ok=True)
            except Exception:
                pass

            return response

        return send_file(
            str(zip_path),
            as_attachment=True,
            download_name="split_pages.zip",
            mimetype="application/zip",
        )

    finally:
        input_path.unlink(missing_ok=True)


# ═══════════════════════════════════════════════════════════════════════════
# FEATURE 3 — COMPRESS PDF
# ═══════════════════════════════════════════════════════════════════════════

GS_SETTINGS = {
    "screen":  "/screen",    # ~72 dpi – smallest, screen display
    "ebook":   "/ebook",     # ~150 dpi – good balance
    "printer": "/printer",   # ~300 dpi – high quality print
    "prepress": "/prepress", # ~300 dpi + colour preservation
}


def _compress_ghostscript(input_path: Path, output_path: Path, quality: str) -> bool:
    """Run Ghostscript for PDF compression.  Returns True on success."""
    gs_bin = shutil.which("gs") or shutil.which("gswin64c")
    if not gs_bin:
        return False

    setting = GS_SETTINGS.get(quality, "/ebook")
    cmd = [
        gs_bin,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS={setting}",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        f"-sOutputFile={output_path}",
        str(input_path),
    ]
    log.info("Ghostscript cmd: %s", " ".join(cmd))
    result = subprocess.run(cmd, capture_output=True, timeout=120)
    if result.returncode != 0:
        log.error("Ghostscript stderr: %s", result.stderr.decode())
        return False
    return True


def _compress_pikepdf(input_path: Path, output_path: Path) -> None:
    """Fallback compressor using pikepdf (removes redundant objects)."""
    with pikepdf.open(str(input_path)) as pdf:
        pdf.save(
            str(output_path),
            compress_streams=True,
            preserve_pdfa=False,
            object_stream_mode=pikepdf.ObjectStreamMode.generate,
        )


@app.route("/api/compress", methods=["POST"])
def compress_pdf():
    """
    Endpoint: POST /api/compress
    Body    : multipart/form-data
      file    – PDF to compress
      quality – "screen" | "ebook" (default) | "printer" | "prepress"

    Returns : compressed PDF with original filename prefixed by 'compressed_'

    Algorithm
    ---------
    1. Try Ghostscript first (best compression ratios via lossy image
       downsampling + stream re-encoding with /dPDFSETTINGS).
    2. If Ghostscript is unavailable or fails, fall back to pikepdf which
       does lossless compression (duplicate-object removal, stream repack).
    3. Log original vs. compressed sizes and the compression ratio achieved.
    4. Guard: if compressed file is larger than the original (can happen with
       already-optimised PDFs), return the original unchanged.
    """
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file provided."), 400
    if Path(f.filename).suffix.lower() not in ALLOWED_EXT:
        return jsonify(error="File must be a PDF."), 400

    quality = request.form.get("quality", "ebook")
    if quality not in GS_SETTINGS:
        return jsonify(error=f"Invalid quality. Choose from: {list(GS_SETTINGS)}"), 400

    input_path = _save_upload(f)
    out_path = _out(".pdf")

    try:
        original_size = input_path.stat().st_size
        log.info("Compress: '%s'  original=%.1f KB  quality=%s",
                 f.filename, original_size / 1024, quality)

        success = _compress_ghostscript(input_path, out_path, quality)
        method = "ghostscript"
        if not success or not out_path.exists():
            log.warning("Ghostscript unavailable/failed – falling back to pikepdf")
            _compress_pikepdf(input_path, out_path)
            method = "pikepdf"

        compressed_size = out_path.stat().st_size
        ratio = (1 - compressed_size / original_size) * 100
        log.info("  method=%s  compressed=%.1f KB  reduction=%.1f%%",
                 method, compressed_size / 1024, ratio)

        # If compression made it bigger, return the original
        if compressed_size >= original_size:
            log.info("  Output larger than input – returning original unchanged")
            out_path.unlink(missing_ok=True)
            shutil.copy(input_path, out_path)

        stem = Path(f.filename).stem
        return _send(out_path, f"compressed_{stem}.pdf", "application/pdf")

    finally:
        input_path.unlink(missing_ok=True)


# ═══════════════════════════════════════════════════════════════════════════
# FEATURE 4 — PDF TO WORD
# ═══════════════════════════════════════════════════════════════════════════

@app.route("/api/pdf-to-word", methods=["POST"])
def pdf_to_word():
    """
    Endpoint: POST /api/pdf-to-word
    Body    : multipart/form-data
      file        – PDF to convert
      start_page  – (optional, 1-based) first page to convert (default: 1)
      end_page    – (optional, 1-based) last page to convert  (default: last)

    Returns : .docx file

    Algorithm
    ---------
    pdf2docx works by:
      a) Rendering each PDF page at high DPI using fitz (PyMuPDF).
      b) Analysing text blocks, images, and layout via a rule-based engine.
      c) Writing python-docx XML that mirrors the detected layout.

    Page-range support is passed directly to Converter.convert().
    The conversion can be slow for large documents; ~1 s per page is typical.
    """
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file provided."), 400
    if Path(f.filename).suffix.lower() not in ALLOWED_EXT:
        return jsonify(error="File must be a PDF."), 400

    input_path = _save_upload(f)
    out_path = _out(".docx")

    try:
        total_pages = _pdf_page_count(input_path)
        start = int(request.form.get("start_page", 1))
        end = int(request.form.get("end_page", total_pages))

        # pdf2docx uses 0-based indexing internally but its public API is 0-based
        # start/end in convert() are 0-based page indices
        start_idx = max(0, start - 1)
        end_idx = min(total_pages - 1, end - 1)

        log.info("PDF→Word: '%s'  pages %d-%d of %d",
                 f.filename, start, end, total_pages)

        cv = Converter(str(input_path))
        cv.convert(str(out_path), start=start_idx, end=end_idx)
        cv.close()

        log.info("  Output: %s (%.1f KB)", out_path.name, out_path.stat().st_size / 1024)
        stem = Path(f.filename).stem
        return _send(out_path, f"{stem}.docx",
                     "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    finally:
        input_path.unlink(missing_ok=True)


# ═══════════════════════════════════════════════════════════════════════════
# FEATURE 5 — PDF TO POWERPOINT
# ═══════════════════════════════════════════════════════════════════════════

SLIDE_W_INCHES = 10.0
SLIDE_H_INCHES = 7.5
RENDER_DPI = 150   # Balance between quality and speed/size


@app.route("/api/pdf-to-pptx", methods=["POST"])
def pdf_to_pptx():
    """
    Endpoint: POST /api/pdf-to-pptx
    Body    : multipart/form-data
      file  – PDF to convert
      dpi   – (optional) render resolution (72-300, default 150)

    Returns : .pptx file where each PDF page becomes one slide

    Algorithm
    ---------
    There is no direct PDF→PPTX library in Python, so we use a
    rasterisation approach:

      1. pdf2image.convert_from_path() calls pdftoppm (poppler) to render
         each PDF page as a PIL Image at the requested DPI.
      2. Each image is saved as a temporary PNG.
      3. A python-pptx Presentation is created with a blank layout and
         slide dimensions matching the aspect ratio of a standard 4:3 slide
         (10 × 7.5 in).
      4. For each image, a new slide is added and the image is inserted as a
         full-slide picture (left=0, top=0, width=slide_width,
         height=slide_height).  python-pptx scales the image automatically.
      5. Temp PNGs are deleted after the PPTX is written.

    Higher DPI → sharper slides but larger file and slower conversion.
    """
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file provided."), 400
    if Path(f.filename).suffix.lower() not in ALLOWED_EXT:
        return jsonify(error="File must be a PDF."), 400

    try:
        dpi = int(request.form.get("dpi", RENDER_DPI))
        dpi = max(72, min(300, dpi))
    except ValueError:
        return jsonify(error="Invalid DPI value."), 400

    input_path = _save_upload(f)
    temp_imgs: list[Path] = []
    out_path = _out(".pptx")

    try:
        total_pages = _pdf_page_count(input_path)
        log.info("PDF→PPTX: '%s'  %d pages  dpi=%d", f.filename, total_pages, dpi)

        # Step 1 – Rasterise PDF pages
        log.info("  Rasterising pages via pdf2image/poppler…")
        pil_images = convert_from_path(
            str(input_path),
            dpi=dpi,
            fmt="png",
            thread_count=2,
        )

        # Step 2 – Save each page as a temp PNG
        for i, img in enumerate(pil_images):
            tmp = OUTPUT_DIR / f"{_uid()}_p{i}.png"
            img.save(str(tmp), "PNG")
            temp_imgs.append(tmp)
            log.info("    Page %d/%d → %s", i + 1, total_pages, tmp.name)

        # Step 3 – Build PPTX
        prs = Presentation()
        prs.slide_width = Inches(SLIDE_W_INCHES)
        prs.slide_height = Inches(SLIDE_H_INCHES)

        blank_layout = prs.slide_layouts[6]  # completely blank layout

        for i, img_path in enumerate(temp_imgs):
            slide = prs.slides.add_slide(blank_layout)
            slide.shapes.add_picture(
                str(img_path),
                left=Inches(0),
                top=Inches(0),
                width=prs.slide_width,
                height=prs.slide_height,
            )
            log.info("    Slide %d/%d added", i + 1, len(temp_imgs))

        prs.save(str(out_path))
        log.info("  Output: %s (%.1f KB)", out_path.name, out_path.stat().st_size / 1024)

        stem = Path(f.filename).stem
        return _send(out_path, f"{stem}.pptx",
                     "application/vnd.openxmlformats-officedocument.presentationml.presentation")

    finally:
        input_path.unlink(missing_ok=True)
        for p in temp_imgs:
            p.unlink(missing_ok=True)


# ═══════════════════════════════════════════════════════════════════════════
# UTILITY ENDPOINTS
# ═══════════════════════════════════════════════════════════════════════════

@app.route("/api/info", methods=["POST"])
def pdf_info():
    """Return basic metadata + page count for an uploaded PDF."""
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file provided."), 400

    p = _save_upload(f)
    try:
        reader = PdfReader(str(p))
        meta = reader.metadata or {}
        info = {
            "filename": f.filename,
            "pages": len(reader.pages),
            "file_size_kb": round(p.stat().st_size / 1024, 1),
            "title": meta.get("/Title", ""),
            "author": meta.get("/Author", ""),
            "creator": meta.get("/Creator", ""),
            "encrypted": reader.is_encrypted,
        }
        return jsonify(info)
    finally:
        p.unlink(missing_ok=True)


@app.route("/api/health")
def health():
    gs = bool(shutil.which("gs") or shutil.which("gswin64c"))
    pdftoppm = bool(shutil.which("pdftoppm"))
    return jsonify(status="ok", ghostscript=gs, pdftoppm=pdftoppm)


# ── Serve the single-page UI ─────────────────────────────────────────────────
@app.route("/")
def index():
    return app.send_static_file("index.html")


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
