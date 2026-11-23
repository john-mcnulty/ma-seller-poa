#!/usr/bin/env python3
"""build_entity_poa.py -- build an Entity-Seller POA from "POA. Seller. Form 2.docx".

Usage: python build_entity_poa.py TEMPLATE.docx VALUES.json OUTPUT.docx

VALUES.json keys (all strings except signers):
  entity_name, state, entity_type, address, buyers, price,
  ps_date:      P&S date as YYYY-MM-DD        -> "September 1st, 2026"; "YYYY-MM", "YYYY", or ""
                print blanks for the unknown parts ("September _____, 2026")
  include_ps_date: false leaves "dated [P&S date]" out of the opening (default true)
  signing_date: POA signing date -> EXECUTED and notary "September 16th, 2026"; "YYYY-MM",
                "YYYY", or "" print blanks for the unknown parts (see notary_block.py)
  notary_state: notary appointment state (default Massachusetts) -- NOT "state", which is the
                entity's state of organization
  county:       notary county ("" or left out -> blank line)
  appointments: optional -- separate notary appointments (see notary_block.py); signers by name
  signers: [{"name": "...", "title": "..."}, ...]   (in signing order)
  pronoun:      "he" / "she" / "they" / "blank"     (as the user answered; required for one signer)

Follows the firm's Entity Seller POA template (v5); the notary block comes from notary_block.py
(the firm's notary acknowledgment, shared by every seller type). Needs notary_block.py in the
same folder. Prints the entity-name font size used. Run keep_together.py on the output afterwards.
"""
import sys as _sys
_sys.dont_write_bytecode = True   # never write __pycache__ into the skill folder
import datetime, json, os, re, sys, tempfile, zipfile
from xml.sax.saxutils import escape
from notary_block import (attorney_in_fact, SIG_TABS, agreement_runs, appointments, check_ps_date,
                          counterparts_paragraph, executed_runs, notary_notes, notary_paragraph,
                          ps_date_text, sig_line_runs)

COL_TWIPS = 4675                      # each signature column (template v2)
BASE_PT, MIN_PT, STEP_PT = 11.0, 8.0, 0.5
GAP_PT = 14.0                         # minimum gap after the entity name (~0.2")
FONT_CANDIDATES = [
    "C:/Windows/Fonts/timesbd.ttf",
    "/Library/Fonts/Times New Roman Bold.ttf",
    "/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf",
    "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf",   # metric-compatible
    "/usr/share/fonts/liberation/LiberationSerif-Bold.ttf",
]
REGULAR_FONT_CANDIDATES = [           # for a name that is not bold (e.g., the trust name in
    "C:/Windows/Fonts/times.ttf",     # the signature lines)
    "/Library/Fonts/Times New Roman.ttf",
    "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
    "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
    "/usr/share/fonts/liberation/LiberationSerif-Regular.ttf",
]

def text_width_pt(text, size_pt, bold=True):
    try:
        from PIL import ImageFont
        for f in (FONT_CANDIDATES if bold else REGULAR_FONT_CANDIDATES):
            if os.path.exists(f):
                return ImageFont.truetype(f, 1000).getlength(text) * size_pt / 1000.0
    except ImportError:
        pass
    return len(text) * (0.56 if bold else 0.52) * size_pt   # conservative fallback for serif

def fit_size(text, width_twips=COL_TWIPS, bold=True):
    """Largest size from 11pt down to 8pt (0.5pt steps) at which TEXT fits on one line in a
    signature column. BOLD says whether the name is printed bold (measured in the matching font)."""
    avail = width_twips / 20.0 - GAP_PT  # points, leaving a visible gap before the next column
    size = BASE_PT
    while size > MIN_PT and text_width_pt(text, size, bold) > avail:
        size -= STEP_PT
    return size, text_width_pt(text, size, bold) <= avail

def run(text, bold=False, size=None, noproof=False):
    rpr = ""
    if bold: rpr += "<w:b/><w:bCs/>"
    if noproof: rpr += "<w:noProof/>"
    if size and size != BASE_PT:
        hp = int(round(size * 2)); rpr += f'<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/>'
    rpr = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'

def body_p(runs=""):
    return f'<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr>{runs}</w:p>'

def ordinal(n):
    return "th" if 11 <= n % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")

def parse_date(iso):
    return datetime.date.fromisoformat(iso)

def long_date_runs(iso):
    """September 16th, 2026 -- ordinal suffix in superscript."""
    d = parse_date(iso)
    sup = (f'<w:r><w:rPr><w:noProof/><w:vertAlign w:val="superscript"/></w:rPr>'
           f'<w:t>{ordinal(d.day)}</w:t></w:r>')
    return run(f"{d:%B} {d.day}", noproof=True) + sup + run(f", {d.year}", noproof=True)

def cell(content):
    return (f'<w:tc><w:tcPr><w:tcW w:w="{COL_TWIPS}" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:pPr><w:keepNext/><w:keepLines/>{SIG_TABS}<w:jc w:val="left"/></w:pPr>{content}'
            '</w:p></w:tc>')

def by_line():
    """'By:' + the signer's signature line, ending 3.0" from the column's left edge -- the same
    length as an individual seller's line (SIG_LINE_TWIPS). Shared with the
    trustee and estate builders."""
    return run("By:") + sig_line_runs()

def row(left="", right=""):
    return f'<w:tr><w:trPr><w:cantSplit/></w:trPr>{cell(left)}{cell(right)}</w:tr>'

def signature_table(entity, signers, ent_size):
    ent = run(entity, size=ent_size)                # not bold
    by = by_line()
    rows = []
    groups = [signers[i:i + 2] for i in range(0, len(signers), 2)]
    for g, grp in enumerate(groups):
        # one signer overall -> right column; otherwise fill left, then right
        cols = [None, grp[0]] if len(signers) == 1 else (grp + [None])[:2]
        pick = lambda f: [f(s) if s else "" for s in cols]
        if g:
            rows += [row(), row()]                      # space between signer pairs
        rows.append(row(*pick(lambda s: ent)))
        rows.append(row())
        rows.append(row(*pick(lambda s: by)))
        rows.append(row(*pick(lambda s: run(f"{s['name']}, its {s['title']}"))))
    none = '<w:{0} w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    borders = "".join(none.format(b) for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    return ('<w:tbl><w:tblPr><w:tblW w:w="9350" w:type="dxa"/>'
            f'<w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
            '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
            '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
            f'</w:tblPr><w:tblGrid><w:gridCol w:w="{COL_TWIPS}"/><w:gridCol w:w="{COL_TWIPS}"/></w:tblGrid>'
            + "".join(rows) + "</w:tbl>")

WITHOUT = ("WITHOUT limiting the foregoing, the following powers are specifically included: To execute, "
           "deliver and acknowledge and make corrections and additions to all deeds and other documents "
           "necessary to effectuate the transfer of the Premises; to receive and disburse proceeds of the "
           "sale; to execute all documents required by the Buyer’s lender in connection with the granting "
           "of a mortgage and related matters, including, but not limited to, Settlement Statements, "
           "Affidavits regarding mechanics’ liens, tenants and compliance with State and Federal Laws; "
           "and other affidavits required by the lender in connection with the issuance of title insurance "
           "or compliance with the requirements of potential assignees of the mortgage.")

def join_list(items):
    """A | A and B | A, B, and C"""
    if len(items) <= 2:
        return " and ".join(items)
    return ", ".join(items[:-1]) + ", and " + items[-1]

def plural(title):
    if title.endswith("y") and title[-2:-1] not in "aeiou":
        return title[:-1] + "ies"
    return title + "s"

def entity_notary(v, appt=None):
    """The shared notary block (notary_block.py), with the entity capacity clause.
    APPT (one notary appointment) limits it to that appointment's signers, place, and date."""
    appt = appt or appointments(v, [s["name"] for s in v["signers"]], "notary_state")[0]
    signers = [s for s in v["signers"] if s["name"] in appt["signers"]]
    return notary_paragraph(dict(appt, signers=[s["name"] for s in signers],
        capacity={"type": "entity", "entity_name": v["entity_name"],
                  "entity_type": v["entity_type"], "titles": [s["title"] for s in signers]}))

def build_body(v, ent_size):
    tab = "<w:r><w:tab/></w:r>"
    opening = body_p(
        tab + run(v["entity_name"], bold=True)
        + run(f', a {v["state"]} {v["entity_type"]}, hereby appoints ', noproof=True)
        + run(attorney_in_fact(v), bold=True)
        + run(f' as its true and lawful attorney in fact (the “Attorney”) for {v["entity_name"]} and in its name, '
              "to do all things necessary with respect to the sale of property situated at ")
        + run(v["address"], bold=True, noproof=True)
        + agreement_runs(v)                           # "...Agreement dated [P&S date], with "
        + run(v["buyers"], noproof=True)              # Buyer plain
        + run(" (the “P&S”), for not less than ")
        + run("$" + v["price"], bold=True, noproof=True)
        + run(" (“Sales Price”), hereby ratifying and affirming that which its Attorney shall lawfully "
              "do or cause to be done by virtue of the powers herein conferred."))
    appts = appointments(v, [s["name"] for s in v["signers"]], "notary_state")
    out = "<w:p/>" + opening + body_p() + body_p(tab + run(WITHOUT)) + body_p()
    if len(appts) > 1:
        out += counterparts_paragraph() + body_p()
    out += signature_sections(v, appts, lambda subset: signature_table(v["entity_name"], subset, ent_size),
                              lambda a: entity_notary(v, a))
    return out


def signature_sections(v, appts, table, notary):
    """One EXECUTED / signature / notary section per notary appointment (just one when everyone
    signs together). TABLE(subset of v["signers"]) and NOTARY(appointment) build the parts.
    keep_together.py adds the page breaks and captions. Shared with the trustee and estate builders."""
    key = (lambda s: s["name"]) if isinstance(v["signers"][0], dict) else (lambda s: s)
    out = ""
    for i, a in enumerate(appts):
        if i:
            out += body_p()
        subset = [s for s in v["signers"] if key(s) in a["signers"]]
        out += (body_p(executed_runs(a.get("signing_date"))) + body_p() + table(subset)
                + body_p() + notary(a))
    return out

def main(template, values_path, out):
    v = json.load(open(values_path, encoding="utf-8"))
    assert v.get("signers"), "at least one signer is required"
    check_ps_date(v)
    ent_size, fits = fit_size(v["entity_name"], bold=False)
    write_poa(template, build_body(v, ent_size), out)
    note = "" if fits else " -- WARNING: still too long at the minimum size; the name will wrap"
    print(f"ENTITY NAME FONT: {ent_size:g}pt (base {BASE_PT:g}pt){note}")
    print(f"P&S DATE: {ps_date_text(v)}")
    for line in notary_notes(appointments(v, [s["name"] for s in v["signers"]], "notary_state")):
        print(line)

def write_poa(template, body, out):
    """Splice BODY (XML) into TEMPLATE after the title and write OUT. No footer (the
    "Seller's Limited Power of Attorney" footer was removed from every POA).
    Shared with build_trustee_poa.py and build_estate_poa.py."""
    work = tempfile.mkdtemp()
    with zipfile.ZipFile(template) as z:
        z.extractall(work)
    p = os.path.join(work, "word", "document.xml")
    x = open(p, encoding="utf-8").read()
    t = x.index("SELLER’S LIMITED POWER OF ATTORNEY")
    start = x.index("</w:p>", t) + len("</w:p>")          # keep everything through the title
    end = x.index("<w:sectPr")
    x = x[:start] + body + x[end:]
    open(p, "w", encoding="utf-8").write(x)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(work):
            for f in files:
                full = os.path.join(root, f)
                z.write(full, os.path.relpath(full, work))

if __name__ == "__main__":
    main(*sys.argv[1:4])
