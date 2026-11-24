#!/usr/bin/env python3
"""render_pages.py -- turn PDF pages into PNG images, so scanned pages (a scanned P&S, deed,
amendment, or probate paper) can be read as pictures.

    python render_pages.py FILE.pdf [--pages 1-3,5] [--dpi 150] [--out FOLDER]

Prints one PNG path per line (then "PAGES: n of m"). Use PyMuPDF when it is installed
(python -m pip install pymupdf), else poppler's `pdftoppm`. Writes only to --out (default: a new
temporary folder) -- never next to the PDF, so nothing lands in a client folder.
Exit 0 on success; 2 when neither renderer is available (says what to install).
"""
import sys as _sys
_sys.dont_write_bytecode = True   # never write __pycache__ into the skill folder
import argparse, os, shutil, subprocess, sys, tempfile

NO_WINDOW = {"creationflags": 0x08000000} if os.name == "nt" else {}   # CREATE_NO_WINDOW


def have_pymupdf():
    try:
        import pymupdf  # noqa: F401
        return True
    except ImportError:
        return False


def renderer():
    """'pymupdf', 'pdftoppm', or None."""
    if have_pymupdf():
        return "pymupdf"
    if shutil.which("pdftoppm"):
        return "pdftoppm"
    return None


def parse_pages(spec, n):
    if not spec:
        return list(range(1, n + 1))
    out = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-", 1)
            out += range(int(a), int(b) + 1)
        elif part:
            out.append(int(part))
    return [p for p in out if 1 <= p <= n]


def render(pdf, pages, dpi, out):
    base = os.path.join(out, os.path.splitext(os.path.basename(pdf))[0][:40] or "page")
    if renderer() == "pymupdf":
        import pymupdf
        paths = []
        with pymupdf.open(pdf) as doc:
            for p in parse_pages(pages, doc.page_count):
                path = f"{base}-p{p:03d}.png"
                doc[p - 1].get_pixmap(dpi=dpi).save(path)
                paths.append(path)
            return paths, doc.page_count
    # pdftoppm: render the requested range page by page (it numbers files itself)
    info = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True, **NO_WINDOW).stdout \
        if shutil.which("pdfinfo") else ""
    n = next((int(l.split()[-1]) for l in info.splitlines() if l.startswith("Pages:")), 9999)
    paths = []
    for p in parse_pages(pages, n) if pages or n != 9999 else range(1, 10000):
        stem = f"{base}-p{p:03d}"
        r = subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-f", str(p), "-l", str(p),
                            "-singlefile", pdf, stem], capture_output=True, **NO_WINDOW)
        if r.returncode != 0 or not os.path.isfile(stem + ".png"):
            break
        paths.append(stem + ".png")
    return paths, (n if n != 9999 else len(paths))


def main():
    ap = argparse.ArgumentParser(description="Render PDF pages to PNG images")
    ap.add_argument("pdf")
    ap.add_argument("--pages", help="e.g. 1-3,5 (default: all)")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--out", help="output folder (default: a new temporary folder)")
    a = ap.parse_args()
    if not renderer():
        print("RENDER: no page renderer -- install PyMuPDF: python -m pip install pymupdf "
              "(or poppler's pdftoppm)")
        return 2
    out = a.out or tempfile.mkdtemp(prefix="poa_pages_")
    os.makedirs(out, exist_ok=True)
    paths, total = render(a.pdf, a.pages, a.dpi, out)
    for p in paths:
        print(p)
    print(f"PAGES: {len(paths)} of {total} rendered with {renderer()} -> {out}")
    return 0 if paths else 1


if __name__ == "__main__":
    sys.exit(main())
