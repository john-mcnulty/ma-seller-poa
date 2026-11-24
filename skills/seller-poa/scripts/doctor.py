#!/usr/bin/env python3
"""doctor.py -- check that this machine can run the Seller POA skill.

    python doctor.py [SETTING FLAGS as for firm_settings.py --show]

REQUIRED  Python 3.9+; pandoc (reads a Word P&S, verifies each POA); LibreOffice (`soffice`) and a
          PDF reader -- the PyMuPDF package (`pip install pymupdf`, preferred) OR poppler's
          `pdftotext` (another program named pdftotext, e.g. xpdf's, does not work) --
          which together render each POA so the signature page can be laid out (caption,
          page-break notice). Without the last two the POAs are still built, but the signature
          section is only marked keep-together and the layout must be checked by hand in Word;
          the skill asks the user before going on that way.
OPTIONAL  Pillow (exact name-width fitting in signature lines), the Times New Roman or Liberation
          Serif fonts (used to measure names).
SETTINGS  attorney_in_fact must be set (see firm_settings.py).

Exit 0 when ready; 2 when only the layout tools (LibreOffice / PDF reader) are missing; 1 otherwise.
Sends nothing over the network.
"""
import sys as _sys
_sys.dont_write_bytecode = True   # never write __pycache__ into the skill folder
import os, platform, shutil, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

INSTALL = {
    "pandoc": "https://pandoc.org/installing.html  (macOS: brew install pandoc; Windows: winget install JohnMacFarlane.Pandoc; Debian/Ubuntu: sudo apt install pandoc)",  # noqa: E501
    "soffice": "https://www.libreoffice.org/download/  (macOS: brew install --cask libreoffice; Windows: winget install TheDocumentFoundation.LibreOffice; Debian/Ubuntu: sudo apt install libreoffice-writer)",  # noqa: E501
    "pdfreader": "python -m pip install pymupdf  (any system; or poppler: macOS brew install poppler, Debian/Ubuntu sudo apt install poppler-utils)",
    "Pillow": "python -m pip install Pillow",
}


def find_soffice():
    for name in ("soffice", "libreoffice"):
        if shutil.which(name):
            return shutil.which(name)
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.isfile(p):
            return p
    return None


NO_WINDOW = {"creationflags": 0x08000000} if os.name == "nt" else {}   # CREATE_NO_WINDOW


def version(cmd):
    try:
        out = subprocess.run([cmd, "--version"], capture_output=True, text=True, timeout=60,
                             **NO_WINDOW).stdout
        return (out.strip().splitlines() or [""])[0][:60]
    except Exception:
        return "found"


def soffice_version(path):
    """LibreOffice's version WITHOUT running it: on Windows `soffice.exe --version` opens a console
    window that waits for Enter. Read it from the install's version file instead."""
    prog = os.path.dirname(os.path.realpath(path))
    for name, key in (("version.ini", "MsiProductVersion="), ("versionrc", "ProductVersion=")):
        f = os.path.join(prog, name)
        if os.path.isfile(f):
            for line in open(f, encoding="utf-8", errors="replace"):
                if line.startswith(key):
                    return "LibreOffice " + line.split("=", 1)[1].strip() + " (" + path + ")"
    return "found at " + path


def main():
    rows, required_ok = [], True

    ok = sys.version_info >= (3, 9)
    required_ok &= ok
    rows.append(("REQUIRED", "Python 3.9+", ok, platform.python_version()))

    pd = shutil.which("pandoc")
    required_ok &= bool(pd)
    rows.append(("REQUIRED", "pandoc", bool(pd), version(pd) if pd else INSTALL["pandoc"]))

    so = find_soffice()
    layout_ok = bool(so)
    rows.append(("REQUIRED", "LibreOffice (soffice)", bool(so),
                 soffice_version(so) if so else "NOT FOUND -- " + INSTALL["soffice"]))

    from keep_together import pdf_reader, pdf_reader_problem
    rd = pdf_reader()
    layout_ok &= bool(rd)
    rows.append(("REQUIRED", "PDF reader (PyMuPDF or poppler pdftotext)", bool(rd),
                 rd if rd else "NOT USABLE -- " + pdf_reader_problem()))

    try:
        import PIL  # noqa: F401
        rows.append(("optional", "Pillow", True, PIL.__version__))
    except ImportError:
        rows.append(("optional", "Pillow", False, "names are measured approximately. " + INSTALL["Pillow"]))

    from build_entity_poa import REGULAR_FONT_CANDIDATES  # the paths the builders search
    fonts = [p for p in REGULAR_FONT_CANDIDATES if os.path.isfile(p)]
    rows.append(("optional", "Times New Roman / Liberation Serif", bool(fonts),
                 fonts[0] if fonts else "not found -- names are measured approximately"))

    import firm_settings  # noqa: E402
    flags = {}
    args = sys.argv[1:]
    for flag, key in (("--attorney-in-fact", "attorney_in_fact"), ("--signing-locations", "signing_locations"),
                      ("--signature-page", "signature_page"), ("--notary-state", "notary_state")):
        if flag in args and args.index(flag) + 1 < len(args):
            flags[key] = args[args.index(flag) + 1]
    st = firm_settings.show(flags)
    settings_ok = not st["errors"]
    rows.append(("REQUIRED", "attorney_in_fact setting", settings_ok,
                 st["attorney_in_fact"] if settings_ok else "NOT SET -- see README, Settings"))
    rows.append(("info", "signing locations", bool(st["signing_locations"]),
                 "; ".join(f"{l['label']} = {l['county'] or '(blank)'}, {l['state']}"
                           for l in st["signing_locations"]) or "none set (the user types the place)"))
    rows.append(("info", "signature page", True, st["signature_page"]))
    rows.append(("info", "settings came from", True,
                 "; ".join(f"{k} <- {v}" for k, v in st["sources"].items())))

    for tier, name, ok, detail in rows:
        mark = "ok  " if ok else ("MISSING" if tier == "REQUIRED" else "--  ")
        print(f"  {mark:8} {tier:12} {name:36} {detail}")
    for w in st["warnings"]:
        print(f"  warning: {w}")
    ready = required_ok and settings_ok and layout_ok
    if ready:
        print("\nREADY")
    elif required_ok and settings_ok:
        print("\nNOT READY (layout) -- LibreOffice or the PDF reader is missing. POAs can still be "
              "built, but without the signature-page layout: ask the user whether to install it "
              "first or go on and check the layout by hand in Word.")
    else:
        print("\nNOT READY -- install or set the REQUIRED items above")
    return 0 if ready else (2 if required_ok and settings_ok else 1)


if __name__ == "__main__":
    sys.exit(main())
