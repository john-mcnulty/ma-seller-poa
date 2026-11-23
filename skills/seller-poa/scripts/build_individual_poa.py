#!/usr/bin/env python3
"""build_individual_poa.py -- build an Individual-Seller POA (any number of sellers) from the
bundled templates. Replaces the hand edits of the old Steps 4A-4C.

Usage: python build_individual_poa.py ASSETS_DIR VALUES.json OUTPUT.docx

ASSETS_DIR is the skill's assets folder; the script picks the template by the number of sellers:
  1 seller            -> "POA. Seller. Form 2.docx"  (I / my / me)
  2 or more sellers   -> "POA. Seller. Form 4.docx"  (We / our / us)
(A template file path is also accepted, if it matches the number of sellers.)

VALUES.json keys:
  sellers:      ["Robert M. Sample", "Karen L. Example", "Helen M. Specimen"]
                (names only, in the order of the P&S SELLER entry -- any number)
  address:      inline property address, e.g. "45 Sample Lane, Marshfield, MA 02050"
  ps_date:      P&S date as YYYY-MM-DD        -> "May 18th, 2022"; "YYYY-MM", "YYYY", or ""
                print blanks for the unknown parts ("May _____, 2022")
  include_ps_date: false leaves "dated [P&S date]" out of the opening (default true)
  buyers:       current Buyer name(s), joined with " and " (a list is also accepted)
  price:        current Purchase Price, numeric with cents, e.g. "400,000.00" (a leading $ is ignored)
  signing_date: POA signing date -> EXECUTED line and notary date: "YYYY-MM-DD", "YYYY-MM",
                "YYYY", or "" (unknown parts print as blanks -- see notary_block.py)
  state:        notary state (default Massachusetts)
  county:       notary county, e.g. "Plymouth" ("" or left out -> blank line)
  pronoun:      "he" / "she" / "they" / "blank" -- required for ONE seller (ask the user, never guess)
  appointments: optional -- separate notary appointments (see notary_block.py)
  confirmed_individuals: true -- ONLY when the user confirms that a seller whose name looks like a
                a trust, estate, or entity (e.g., "Trust" in a surname) really is an individual

Backstop: the script stops, building nothing, when a
seller's name looks like a trust, trustee, estate fiduciary, or entity -- that seller belongs in
its own principal (its own POA), listed separately in build_all_poas.py.

What it writes (everything else is the template's own text, unchanged):
  - Opening paragraph: "I, **A**, ..." / "We, **A** and **B**, ..." /
    "We, **A, B, C,** and **D**, ..." (names bold, the final "and" plain), then
    "hereby appoint **[attorney_in_fact]** as my/our true and lawful attorney in fact
    (my/our "Attorney") for me/us and in my/our name, to do all things necessary with respect to
    the sale of property situated at **[Address]** (the "Premises"), as set forth in a Purchase
    and Sale Agreement dated [May 18th, 2022], with [Buyers] (the "P&S"), for not less than
    **$[Price]** ("Sales Price"), hereby ratifying and affirming that which my/our Attorney shall
    lawfully do or cause to be done by virtue of the powers herein conferred."
  - The EXECUTED line, signature lines, and notary block (notary_block.py): one seller -> Form 2's
    right-hand signature line; two or more -> Form 4's borderless two-column table, filled left
    then right, two signers per row (a fifth signer starts a third row, left column).
  - Several notary appointments: the counterparts paragraph and one counterpart section each.
Needs notary_block.py in the same folder. Run keep_together.py and check_agreement.py afterwards.
"""
import sys as _sys
_sys.dont_write_bytecode = True   # never write __pycache__ into the skill folder
import json, os, re, sys, tempfile, zipfile
from xml.sax.saxutils import escape
from notary_block import (BODY_PPR, _date_runs, agreement_runs, appointments, attorney_in_fact,
                          check_ps_date,
                          counterparts_paragraph, individual_counterparts, notary_notes,
                          ps_date_text)

FORM_1 = "POA. Seller. Form 2.docx"      # one seller
FORM_N = "POA. Seller. Form 4.docx"      # two or more sellers
PARA = r"<w:p(?: [^>]*)?/>|<w:p(?: [^>]*[^/>])?>.*?</w:p>"   # self-closing <w:p .../> first


def r(text, bold=False, noproof=False):
    rpr = ("<w:b/><w:bCs/>" if bold else "") + ("<w:noProof/>" if noproof else "")
    rpr = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def words(n):
    return (dict(I="I", my="my", me="me") if n == 1 else dict(I="We", my="our", me="us"))


def seller_runs(names):
    """A | A and B | A, B, C, and D -- names (and the commas between them) bold, 'and' plain."""
    if len(names) == 1:
        return r(names[0], bold=True, noproof=True)
    if len(names) == 2:
        return (r(names[0], bold=True, noproof=True) + r(" and ", noproof=True)
                + r(names[1], bold=True, noproof=True))
    return (r(", ".join(names[:-1]) + ",", bold=True, noproof=True) + r(" and ", noproof=True)
            + r(names[-1], bold=True, noproof=True))


def bookmark(i, name, runs):
    return (f'<w:bookmarkStart w:id="{i}" w:name="{name}"/>' + runs
            + f'<w:bookmarkEnd w:id="{i}"/>')


def opening_runs(v):
    w = words(len(v["sellers"]))
    price = str(v["price"]).strip().lstrip("$").strip()
    buyers = v["buyers"] if isinstance(v["buyers"], str) else " and ".join(v["buyers"])
    return ("<w:r><w:tab/></w:r>" + r(f'{w["I"]}, ') + seller_runs(v["sellers"])
            + r(", hereby appoint ") + r(attorney_in_fact(v), bold=True)
            + r(f' as {w["my"]} true and lawful attorney in fact ({w["my"]} “Attorney”) for {w["me"]} '
                f'and in {w["my"]} name, to do all things necessary with respect to the sale of '
                "property situated at ")
            + bookmark(0, "property_address_inline|1", r(v["address"].strip(), bold=True, noproof=True))
            + agreement_runs(v)                       # "...Agreement dated [P&S date], with "
            + bookmark(1, "borrower_name|2", r(buyers.strip(), noproof=True))
            + r(" (the “P&S”), for not less than ") + r("$", bold=True)
            + bookmark(2, "purchase_price|3", r(price, bold=True, noproof=True))
            + r(f' (“Sales Price”), hereby ratifying and affirming that which {w["my"]} Attorney '
                "shall lawfully do or cause to be done by virtue of the powers herein conferred."))


REGULAR_FONTS = ["C:/Windows/Fonts/times.ttf", "/Library/Fonts/Times New Roman.ttf",
                 "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
                 "/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf",
                 "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
                 "/usr/share/fonts/liberation/LiberationSerif-Regular.ttf"]
COLUMN_PT = 4675 / 20.0 - 11.0        # signature column less Word's default cell margins


def too_wide(name, size_pt=11.0):
    """True when a seller's name will wrap under its signature line in the two-column table."""
    try:
        from PIL import ImageFont
        for f in REGULAR_FONTS:
            if os.path.exists(f):
                return ImageFont.truetype(f, 1000).getlength(name) * size_pt / 1000.0 > COLUMN_PT
    except ImportError:
        pass
    return len(name) * 0.5 * size_pt > COLUMN_PT


NOT_INDIVIDUAL = re.compile(
    r"\b(trusts?|trustees?|u/?d/?t|nominee|estate\s+of|personal\s+representatives?|executor|executrix|"
    r"administrat(?:or|rix|ors|rices)|inc|incorporated|corporation|corp|llc|l\.l\.c|company|"
    r"limited\s+partnership|lp|llp)\b\.?", re.I)


def non_individuals(names):
    """Seller names that look like a trust, trustee, estate fiduciary, or entity -- a sign that a
    seller of another type was put in the individual POA. Backstop for SKILL.md Step 2
    "Principals"."""
    return [(n, m.group(0)) for n in names for m in [NOT_INDIVIDUAL.search(n)] if m]


def pick_template(assets, n):
    if os.path.isdir(assets):
        return os.path.join(assets, FORM_1 if n == 1 else FORM_N)
    base = os.path.basename(assets)
    assert (base == FORM_1) == (n == 1), \
        f"{base} does not fit {n} seller(s) -- use {FORM_1 if n == 1 else FORM_N} (or pass the assets folder)"
    return assets


def build(xml, v):
    paras = list(re.finditer(PARA, xml, re.S))
    text = lambda m: "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", m.group(0)))
    # 1. opening paragraph: keep its <w:p ...> tag and paragraph properties, rebuild the runs
    op = [m for m in paras if "hereby appoint" in text(m)]
    assert len(op) == 1, "template: expected one opening paragraph ('hereby appoint')"
    p = op[0].group(0)
    head = re.match(r"<w:p(?: [^>]*)?>", p).group(0)
    ppr = re.search(r"<w:pPr>.*?</w:pPr>", p, re.S)
    new_open = head + (ppr.group(0) if ppr else BODY_PPR) + opening_runs(v) + "</w:p>"
    # 2. EXECUTED line through the {{Notary Block}} paragraph -> built sections
    ex = [m for m in paras if text(m).startswith("EXECUTED as a sealed instrument")]
    nb = [m for m in paras if "{{Notary Block}}" in text(m)]
    assert len(ex) == 1 and len(nb) == 1, "template: expected one EXECUTED line and one {{Notary Block}}"
    appts = appointments(v, v["sellers"])
    section = ""
    if len(appts) > 1:
        section += counterparts_paragraph() + f"<w:p>{BODY_PPR}</w:p>"
    section += individual_counterparts(dict(v, signers=v["sellers"]), appts)
    out = (xml[:op[0].start()] + new_open + xml[op[0].end():ex[0].start()] + section
           + xml[nb[0].end():])
    return out, appts


def main(assets, values_path, out):
    v = json.load(open(values_path, encoding="utf-8"))
    sellers = [s.strip() for s in v.get("sellers") or []]
    assert sellers, "at least one seller is required"
    assert len(set(sellers)) == len(sellers), "a seller is listed twice"
    v["sellers"] = sellers
    for k in ("address", "buyers", "price"):
        assert v.get(k), f"{k} is required"
    check_ps_date(v)
    odd = non_individuals(sellers)
    if odd and not v.get("confirmed_individuals"):
        sys.exit("MIXED SELLER TYPES: STOPPED -- no POA built. These seller names do not look like "
                 "individuals: " + "; ".join(f"'{n}' ('{w}')" for n, w in odd)
                 + ". A trust, estate, or entity seller is its own principal with its own POA -- "
                 "list it as a separate principal in build_all_poas.py (SKILL.md Step 2, "
                 "\"Principals\"). Only if the user confirms every seller here is an individual, "
                 "add \"confirmed_individuals\": true to the individual principal.")
    template = pick_template(assets, len(sellers))
    work = tempfile.mkdtemp()
    with zipfile.ZipFile(template) as z:
        z.extractall(work)
    p = os.path.join(work, "word", "document.xml")
    xml, appts = build(open(p, encoding="utf-8").read(), v)
    open(p, "w", encoding="utf-8").write(xml)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(work):
            for f in files:
                full = os.path.join(root, f)
                z.write(full, os.path.relpath(full, work))
    n = len(sellers)
    rows = (n + 1) // 2
    layout = ("one signature line (right side)" if n == 1 else
              f"two-column signature table, {rows} row{'s' if rows > 1 else ''}")
    if len(appts) > 1:
        layout = "signature lines on each counterpart page (one seller: right side; two or more: two columns)"
    print(f"TEMPLATE: {os.path.basename(template)} -- {n} seller{'s' if n > 1 else ''} "
          f"({'I/my/me' if n == 1 else 'We/our/us'}); {layout}")
    print(f"P&S DATE: {ps_date_text(v)}")
    for line in notary_notes(appts):
        print(line)
    if n > 1:
        for name in sellers:
            if too_wide(name):
                print(f"NAME WIDTH: '{name}' is too long for its signature column and will wrap onto "
                      "a second line under the signature line -- tell the user (check the page in Word)")


if __name__ == "__main__":
    main(*sys.argv[1:4])
