#!/usr/bin/env python3
"""keep_together.py -- place the EXECUTED / signature / notary section of a Seller POA.

Usage: python keep_together.py POA.docx [--mode same|separate] [--address "123 Sample Street, Hingham, MA 02043"]

--mode same      (default) keep the section on the same page as the body text when it fits;
                 otherwise start it on the next page.
--mode separate  always start the section on a new page (the "signature page").
--address        the inline property address for the signature-page caption; read from the
                 opening paragraph ("situated at ... (the "Premises")") when left out.

1. Marks every paragraph of the section (from the caption, or "EXECUTED as a sealed instrument",
   through the notary paragraph -- "My Commission Expires", or the "{{Notary Block}}"
   placeholder) keep-with-next + keep-lines, and every table row in it can't-split, so Word
   moves the whole section as one block.
2. Renders the document with LibreOffice. In "same" mode, if the section fits on the page
   where the body text ends, nothing else is done. Otherwise (or always, in "separate" mode):
   - a page break starts the signature page, which opens with the bold, flush-left caption
     CAPTION ("Signature Page to Seller's Limited Power of Attorney for the sale of [address].")
     followed by one blank line and then "EXECUTED ...";
   - the page-break notice PB_TEXT (bold italic, centered horizontally, and placed about
     one-quarter of the way down the blank space between the last line of text and the bottom
     margin)
     is added below the last body paragraph ONLY when at least NOTICE_MIN_PT (1.5", about six lines) of blank space is left
     on the page where the body text ends.
   With no page break, the caption and notice (and any "{{Page Break Language}}" placeholder)
   are removed. Safe to run again on its own output.
The .docx is updated in place. Prints a one-line result.
"""
import sys as _sys
_sys.dont_write_bytecode = True   # never write __pycache__ into the skill folder
import argparse, os, re, shutil, subprocess, sys, tempfile, zipfile
from xml.sax.saxutils import escape

START = "EXECUTED as a sealed instrument"
END = "My Commission Expires"
END_ALT = "{{Notary Block}}"   # fallback: notary placeholder not filled yet
PB_TEXT = "[Remainder of page intentionally left blank; signature page follows]"
PB_TEXT_MULTI = "[Remainder of page intentionally left blank; signature pages follow]"  # several appointments
PB_MARKERS = (PB_TEXT, PB_TEXT_MULTI, "{{Page Break Language}}", "[Signature to follow on the next page]")  # current, placeholder, earlier wording
CAPTION_START = "Signature Page to Seller’s Limited Power of Attorney"
CAPTION = CAPTION_START + " for the sale of {address}."
COUNTERPART_START = "Counterpart Signature Page"
COUNTERPART = COUNTERPART_START + " {k} of {n} to Seller’s Limited Power of Attorney for the sale of {address}."
NOTICE_MIN_PT = 108.0          # 1.5 inches -- about six lines of blank space
LINE_PT = 14.0                 # approximate height of one line of 11pt text
PARA = r"<w:p(?: [^>]*)?/>|<w:p(?: [^>]*)?>.*?</w:p>"


def para_text(p):
    return "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p))


def set_ppr_flags(p, keep_next, keep_lines, page_break=None):
    """Return paragraph XML with keepNext/keepLines (and optionally pageBreakBefore) set."""
    m = re.match(r"(<w:p(?: [^>]*?)?)(/?)>", p)
    if m.group(2) == "/":  # self-closing <w:p/>
        p = m.group(1) + "></w:p>"
    head_end = p.index(">") + 1
    body = p[head_end:]
    if not body.startswith("<w:pPr>"):
        body = "<w:pPr></w:pPr>" + body
    ppr_end = body.index("</w:pPr>")
    ppr = body[len("<w:pPr>"):ppr_end]
    was_break = "<w:pageBreakBefore/>" in ppr
    ppr = re.sub(r"<w:(keepNext|keepLines|pageBreakBefore)(?: [^>]*)?/>", "", ppr)
    flags = ""
    if keep_next:
        flags += "<w:keepNext/>"
    if keep_lines:
        flags += "<w:keepLines/>"
    if page_break or (page_break is None and was_break):
        flags += "<w:pageBreakBefore/>"
    ps = re.match(r"<w:pStyle [^>]*/>", ppr)
    ppr = (ppr[: ps.end()] + flags + ppr[ps.end():]) if ps else flags + ppr
    return p[:head_end] + "<w:pPr>" + ppr + body[ppr_end:]


def section_start(xml):
    """Start of the paragraph that opens the section: the caption if present, else EXECUTED."""
    hits = [j for j in (xml.find(CAPTION_START), xml.find(COUNTERPART_START), xml.find(START)) if j >= 0]
    i = min(hits)
    return [m.start() for m in re.finditer(r"<w:p[ >]", xml[:i])][-1]


def mark(xml, page_break):
    return mark_at(xml, section_start(xml), page_break)


def mark_at(xml, s, page_break):
    """Keep-together flags on the section whose first paragraph starts at S (through the next
    notary block); PAGE_BREAK puts a page break before it."""
    e = xml.find(END, s)
    e = e if e >= 0 else xml.index(END_ALT, s)
    e = xml.index("</w:p>", e) + len("</w:p>")
    seg = xml[s:e]
    # table rows: can't split
    def row(m):
        tr = m.group(0)
        if "<w:cantSplit/>" in tr:
            return tr
        if "<w:trPr>" in tr:
            return tr.replace("<w:trPr>", "<w:trPr><w:cantSplit/>", 1)
        return re.sub(r"(<w:tr(?: [^>]*)?>(?:<w:tblPrEx>.*?</w:tblPrEx>)?)", lambda h: h.group(1) + "<w:trPr><w:cantSplit/></w:trPr>", tr, count=1, flags=re.S)
    seg = re.sub(r"<w:tr(?: [^>]*)?>.*?</w:tr>", row, seg, flags=re.S)
    paras = list(re.finditer(PARA, seg, re.S))
    out, pos = [], 0
    for i, m in enumerate(paras):
        p = m.group(0)
        first, last = i == 0, i == len(paras) - 1
        p = set_ppr_flags(p, keep_next=not last, keep_lines=True,
                          page_break=(page_break if first else False))
        out.append(seg[pos:m.start()] + p)
        pos = m.end()
    out.append(seg[pos:])
    return xml[:s] + "".join(out) + xml[e:]


def strip_added(xml):
    """Remove the notice, the caption (with the blank line after it) and any placeholder,
    so the script starts from the plain document every time it runs."""
    out, skip_blank = [], False
    pos = 0
    for m in re.finditer(PARA, xml, re.S):
        p = m.group(0)
        out.append(xml[pos:m.start()])
        pos = m.end()
        text = para_text(p)
        if any(mk in text for mk in PB_MARKERS) or CAPTION_START in text or COUNTERPART_START in text:
            skip_blank = True
            continue
        if skip_blank and not text.strip() and "<w:tbl" not in p:
            skip_blank = False
            continue
        skip_blank = False
        out.append(p)
    out.append(xml[pos:])
    return "".join(out)


def _p(text, italic=False, center=False, before_pt=None):
    rpr = "<w:b/><w:bCs/>" + ("<w:i/><w:iCs/>" if italic else "")
    sp = "" if before_pt is None else f'<w:spacing w:before="{max(0, round(before_pt * 20))}" w:after="0"/>'
    jc = "center" if center else "left"
    return (f'<w:p><w:pPr><w:pStyle w:val="BodyText"/>{sp}<w:jc w:val="{jc}"/></w:pPr><w:r><w:rPr>{rpr}'
            f'</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>')


BLANK = '<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr></w:p>'


def add_caption(xml, address):
    """Bold, flush-left caption + one blank line, just before the EXECUTED paragraph."""
    s = [m.start() for m in re.finditer(r"<w:p[ >]", xml[:xml.index(START)])][-1]
    return xml[:s] + _p(CAPTION.format(address=address)) + BLANK + xml[s:]


def executed_starts(xml):
    """Start of each EXECUTED paragraph -- one per notary appointment (counterpart section)."""
    out = []
    for i in [m.start() for m in re.finditer(re.escape(START), xml)]:
        out.append([m.start() for m in re.finditer(r"<w:p[ >]", xml[:i])][-1])
    return out


def counterpart_layout(xml, address, before_pt=None):
    """Several notary appointments: every section starts its own page, opening with the caption
    'Counterpart Signature Page k of n to ...' + a blank line; optional notice before page 1's
    break. Sections are processed last to first so earlier positions stay valid."""
    starts = executed_starts(xml)
    n = len(starts)
    for k in range(n, 0, -1):
        s = starts[k - 1]
        xml = xml[:s] + _p(COUNTERPART.format(k=k, n=n, address=address)) + BLANK + xml[s:]
        xml = mark_at(xml, s, page_break=True)
    if before_pt is not None:
        s = section_start(xml)
        xml = xml[:s] + _p(PB_TEXT_MULTI, italic=True, center=True, before_pt=before_pt) + xml[s:]
    return xml


def add_notice(xml, before_pt):
    """Bold-italic notice, centered horizontally, on its own line just before the caption.
    BEFORE_PT (space above it) places it one-quarter of the way down the blank space."""
    s = section_start(xml)
    return xml[:s] + _p(PB_TEXT, italic=True, center=True, before_pt=before_pt) + xml[s:]


def opening_address(xml):
    text = " ".join(para_texts(xml)).replace("\u200b", "")
    m = re.search(r"situated at (.+?) \(the [“\"]Premises[”\"]\)", text)
    return " ".join(m.group(1).split()) if m else None


def para_texts(xml):
    return [para_text(p) for p in re.findall(r"<w:p(?: [^>]*)?>.*?</w:p>", xml, re.S)]


def previous_text(xml):
    """Last few words of the last non-empty body paragraph before the section."""
    s = section_start(xml)
    for t in reversed(para_texts(xml[xml.index("<w:body>"):s])):
        t = t.strip()
        if t and not any(mk in t for mk in PB_MARKERS):
            return " ".join(t.split()[-4:])
    return None


def usable_bottom_pt(xml):
    """Distance from the top of the page to the bottom margin, in points."""
    h = re.search(r'<w:pgSz [^>]*w:h="(\d+)"', xml)
    b = re.search(r'<w:pgMar [^>]*w:bottom="(\d+)"', xml)
    return (int(h.group(1)) if h else 15840) / 20.0 - (int(b.group(1)) if b else 1440) / 20.0


def write_docx(src_dir, path):
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(src_dir):
            for f in files:
                full = os.path.join(root, f)
                z.write(full, os.path.relpath(full, src_dir))
    os.replace(tmp, path)


def find_soffice():
    for c in ("soffice", "libreoffice"):
        if shutil.which(c):
            return shutil.which(c)
    for c in ("C:/Program Files/LibreOffice/program/soffice.exe", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(c):
            return c
    return None


NO_WINDOW = {"creationflags": 0x08000000} if os.name == "nt" else {}   # CREATE_NO_WINDOW


def render(docx, soffice):
    """Return (pages_text, pages_words): per page, the joined text and a list of (yMin, yMax, word)."""
    out = tempfile.mkdtemp()
    subprocess.run([soffice, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", out, docx],
                   check=True, capture_output=True, timeout=180, **NO_WINDOW)
    pdf = os.path.join(out, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    if pdf_reader() == "pymupdf":
        return render_pymupdf(pdf)
    try:
        html = subprocess.run(["pdftotext", "-bbox", pdf, "-"], capture_output=True, text=True,
                              check=True, timeout=120, **NO_WINDOW).stdout
    except (subprocess.CalledProcessError, OSError):
        if _have_pymupdf():                 # a pdftotext that passed the check but still failed
            return render_pymupdf(pdf)
        raise
    pages_text, pages_words = [], []
    for pg in re.findall(r"<page [^>]*>(.*?)</page>", html, re.S):
        words = [(float(y0), float(y1), w) for y0, y1, w in
                 re.findall(r'<word [^>]*yMin="([\d.]+)"[^>]*yMax="([\d.]+)"[^>]*>([^<]*)</word>', pg)]
        pages_words.append(words)
        pages_text.append(" ".join(w for _, _, w in words).replace("&amp;", "&").replace("&quot;", '"'))
    return pages_text, pages_words


def _have_pymupdf():
    try:
        import pymupdf  # noqa: F401
        return True
    except ImportError:
        return False


_BBOX_OK = {}


def pdftotext_has_bbox():
    """True only for a pdftotext that supports -bbox (poppler's). Other programs named
    pdftotext -- e.g. the xpdf build bundled with some Windows tools -- do not, and the layout
    step would fail with them."""
    exe = shutil.which("pdftotext")
    if not exe:
        return False
    if exe not in _BBOX_OK:
        try:
            r = subprocess.run([exe, "-h"], capture_output=True, text=True, timeout=30, **NO_WINDOW)
            _BBOX_OK[exe] = "-bbox" in (r.stdout + r.stderr)
        except Exception:
            _BBOX_OK[exe] = False
    return _BBOX_OK[exe]


def pdf_reader():
    """The PDF word reader keep_together.py will use: 'pymupdf' (pip install pymupdf) when it is
    installed, else 'pdftotext' when it is poppler's (supports -bbox), else None. PyMuPDF comes
    first: it is a Python package, so there is no question which program a name on the PATH
    points to. Both give identical results on the whole test suite."""
    if _have_pymupdf():
        return "pymupdf"
    if pdftotext_has_bbox():
        return "pdftotext"
    return None


def pdf_reader_problem():
    """Why there is no usable reader, in words, or None."""
    if pdf_reader():
        return None
    if shutil.which("pdftotext"):
        return (f"the pdftotext found at {shutil.which('pdftotext')} does not support -bbox (it is "
                "not poppler's) -- install PyMuPDF: python -m pip install pymupdf")
    return "no PDF reader -- install PyMuPDF: python -m pip install pymupdf"


def render_pymupdf(pdf):
    """Same result as the pdftotext -bbox path: per page, the joined text and (yMin, yMax, word)
    in points from the top of the page."""
    import pymupdf
    pages_text, pages_words = [], []
    with pymupdf.open(pdf) as doc:
        for page in doc:
            words = [(float(w[1]), float(w[3]), w[4]) for w in
                     sorted(page.get_text("words"), key=lambda w: (w[5], w[6], w[7]))]
            pages_words.append(words)
            pages_text.append(" ".join(w for _, _, w in words))
    return pages_text, pages_words


def page_index(pages, needle, last=False):
    """Index of the first (or, with last=True, the last) page containing NEEDLE."""
    needle = " ".join(needle.split())
    hits = [i for i, pg in enumerate(pages) if needle in pg]
    return (hits[-1] if last else hits[0]) if hits else None


def notice_center(pages_words, page):
    """Vertical center (pt from the top) of the notice line on PAGE, or None if it isn't there."""
    if page is None or page >= len(pages_words):
        return None
    hits = [(y0, y1) for y0, y1, w in pages_words[page] if w in ("[Remainder", "follows]", "follow]")]
    if len(hits) < 2:
        return None
    return (min(h[0] for h in hits) + max(h[1] for h in hits)) / 2.0


def blank_space_pt(words, bottom):
    """Blank space between the last line of text and the bottom margin (footer excluded)."""
    ys = [y for _, y, _ in words if y <= bottom + 1]
    return bottom - max(ys) if ys else bottom


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--mode", choices=("same", "separate"), default="same")
    ap.add_argument("--address")
    a = ap.parse_args()
    path = a.docx

    work = tempfile.mkdtemp()
    with zipfile.ZipFile(path) as z:
        z.extractall(work)
    docxml = os.path.join(work, "word", "document.xml")
    xml = strip_added(open(docxml, encoding="utf-8").read())
    address = a.address or opening_address(xml)
    end = END if END in xml else END_ALT

    def save(x):
        open(docxml, "w", encoding="utf-8").write(x)
        write_docx(work, path)

    n_sections = len(executed_starts(xml))
    plain = xml
    for s in reversed(executed_starts(xml)):
        plain = mark_at(plain, s, page_break=False)
    save(plain)
    soffice = find_soffice()
    if not soffice or not pdf_reader():
        missing = "; ".join(x for x, ok in (("LibreOffice not found", soffice),
                             (pdf_reader_problem() or "", pdf_reader())) if not ok)
        if a.mode == "separate":
            print(f"KEEP-TOGETHER: could not render ({missing}) -- separate signature page "
                  "NOT set up; check the layout in Word")
        else:
            print(f"KEEP-TOGETHER: keep-with-next set; could not render to check page fit ({missing}) "
                  "-- no signature-page caption or notice added; check the layout in Word")
        return

    def layout():
        prev = previous_text(xml)
        if n_sections > 1:
            return counterparts_main(xml, address, prev, path, soffice, save, a.mode)
        if a.mode == "same":
            pages, _ = render(path, soffice)
            p_start, p_end = page_index(pages, START), page_index(pages, end)
            p_prev = page_index(pages, prev, last=True) if prev else p_start
            if p_start is None or p_end is None:
                print("KEEP-TOGETHER: keep-with-next set; could not locate the section in the rendered PDF")
                return
            if p_start == p_end and (p_prev is None or p_prev == p_start):
                print(f"KEEP-TOGETHER: same-page mode -- section fits on page {p_start + 1}; no page break needed")
                return

        # The section goes on its own page: caption + page break, then decide on the notice.
        assert address, "could not read the property address for the caption -- pass --address"
        xml2 = mark(add_caption(xml, address), page_break=True)
        save(xml2)
        pages, words = render(path, soffice)
        p_prev = page_index(pages, prev, last=True) if prev else None
        space = blank_space_pt(words[p_prev], usable_bottom_pt(xml)) if p_prev is not None else 0.0
        notice = space >= NOTICE_MIN_PT
        off = None
        if notice:
            # Place the notice, measure where it landed, then correct the space above it so its
            # center sits one-quarter of the way from the last line of text to the bottom margin.
            bottom = usable_bottom_pt(xml)
            target = (bottom - space) + space / 4.0     # one-quarter of the way down the empty space
            before = space / 4.0 - 2 * LINE_PT          # first guess (a blank line already precedes it)
            for _ in range(3):
                xml2 = mark(add_notice(add_caption(xml, address), before), page_break=True)
                save(xml2)
                pages, words = render(path, soffice)
                c = notice_center(words, p_prev)
                if c is None:                            # pushed off the page -- back off
                    before = max(0.0, before - LINE_PT)
                    continue
                off = target - c
                if abs(off) <= 3:
                    break
                before = max(0.0, before + off)
        p_cap = page_index(pages, CAPTION_START)
        p_end = page_index(pages, end)
        p_prev = page_index(pages, prev, last=True) if prev else None
        ok = p_cap is not None and p_cap == p_end and (p_prev is None or p_prev < p_cap)
        why = "separate-page mode" if a.mode == "separate" else "section did not fit"
        note = (f"notice added, 1/4 down ({space / 72:.1f}\" blank"
                + (f"; {abs(off):.0f}pt off target" if off is not None and abs(off) > 3 else "") + ")" if notice
                else f"no notice ({space / 72:.1f}\" blank, under 1.5\")")
        print(f"KEEP-TOGETHER: {why} -- signature page is page {p_cap + 1 if p_cap is not None else '?'} with caption; "
              f"{note} ({'OK' if ok else 'CHECK LAYOUT MANUALLY'})")


    # A layout failure (LibreOffice or the PDF reader erroring or timing out) must not stop the
    # build: fall back to keep-with-next only, say so, and let the POA be saved.
    try:
        return layout()
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError) as e:
        save(plain)
        detail = (getattr(e, "stderr", None) or b"")
        detail = (detail.decode("utf-8", "replace") if isinstance(detail, bytes) else str(detail)).strip()
        print(f"KEEP-TOGETHER: keep-with-next set; the layout check could not run "
              f"({type(e).__name__}{': ' + detail.splitlines()[-1][:120] if detail else ''}) -- no "
              "signature-page caption or notice added; check the layout in Word")

def place_notice(build, xml, prev, path, soffice, save):
    """Render BUILD(None); if the page where the text ends has 1.5" or more of blank space, add the
    notice and move it until its center is one-quarter of the way down that space.
    Returns (pages, words, space, notice, off)."""
    save(build(None))
    pages, words = render(path, soffice)
    p_prev = page_index(pages, prev, last=True) if prev else None
    space = blank_space_pt(words[p_prev], usable_bottom_pt(xml)) if p_prev is not None else 0.0
    notice, off = space >= NOTICE_MIN_PT, None
    if notice:
        bottom = usable_bottom_pt(xml)
        target = (bottom - space) + space / 4.0
        before = space / 4.0 - 2 * LINE_PT
        for _ in range(3):
            save(build(before))
            pages, words = render(path, soffice)
            c = notice_center(words, p_prev)
            if c is None:
                before = max(0.0, before - LINE_PT)
                continue
            off = target - c
            if abs(off) <= 3:
                break
            before = max(0.0, before + off)
    return pages, words, space, notice, off


def counterparts_main(xml, address, prev, path, soffice, save, mode):
    """Several notary appointments (backlog B3): one counterpart signature page per appointment,
    whatever the placement setting (each signer group signs its own page)."""
    assert address, "could not read the property address for the caption -- pass --address"
    n = len(executed_starts(xml))
    pages, words, space, notice, off = place_notice(
        lambda before: counterpart_layout(xml, address, before), xml, prev, path, soffice, save)
    norm = lambda t: " ".join(t.split())
    cap_pages = []
    for k in range(1, n + 1):
        cap_pages.append(page_index(pages, f"{COUNTERPART_START} {k} of {n}"))
    end_pages = [i for i, pg in enumerate(pages) for _ in range(pg.count(END))]
    p_prev = page_index(pages, prev, last=True) if prev else None
    ok = (None not in cap_pages and cap_pages == sorted(set(cap_pages)) and end_pages == cap_pages
          and (p_prev is None or p_prev < cap_pages[0]))
    note = (f"notice added, 1/4 down ({space / 72:.1f}\" blank"
            + (f"; {abs(off):.0f}pt off target" if off is not None and abs(off) > 3 else "") + ")" if notice
            else f"no notice ({space / 72:.1f}\" blank, under 1.5\")")
    pg = ", ".join(str(p + 1) if p is not None else "?" for p in cap_pages)
    extra = " (placement setting not used -- each appointment signs its own page)" if mode == "same" else ""
    print(f"KEEP-TOGETHER: {n} notary appointments -- counterpart signature pages {pg}, each with its "
          f"caption{extra}; {note} ({'OK' if ok else 'CHECK LAYOUT MANUALLY'})")


if __name__ == "__main__":
    main()
