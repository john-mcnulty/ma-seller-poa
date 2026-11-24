#!/usr/bin/env python3
"""build_estate_poa.py -- build a Probate-Estate-Seller POA from "POA. Seller. Form 2.docx".

Usage: python build_estate_poa.py TEMPLATE.docx VALUES.json OUTPUT.docx
Needs build_entity_poa.py and notary_block.py in the same folder (shared helpers).

VALUES.json keys:
  signers:        ["Ellen K. Example"]          (every fiduciary, in signing order)
  capacity:       "Personal Representative"          (singular, as in the P&S; pluralized
                                                      automatically for 2+ signers)
  estate_name:    "Estate of Edith A. Sample"      (as in the P&S; a leading "The" is dropped
                                                      on the signature lines)
  probate_county: "Worcester"                        (court county, without "County")
  docket:         "WO24P0000EA"
  address, buyers, price
  ps_date:        P&S date as YYYY-MM-DD        -> "February 19th, 2026"; "YYYY-MM", "YYYY", or
                  "" print blanks for the unknown parts ("February _____, 2026")
  include_ps_date: false leaves "dated [P&S date]" out of the opening (default true)
  signing_date:   POA signing date -> EXECUTED and notary dates ("YYYY-MM-DD", "YYYY-MM",
                  "YYYY", or "" -- see notary_block.py)
  state:          notary state (default Massachusetts)
  county:         notary county, e.g. "Plymouth" ("" or left out -> blank line)
  appointments:   optional -- separate notary appointments (see notary_block.py)
  pronoun:        "he" / "she" / "they" / "blank" (required for ONE signer -- ask, never guess)
  include_incapacity: true / false              (default true)
  preposition:    optional override of the word between the capacity and the estate name.
                  Rule: "for" when the capacity is "Public Administrator",
                  otherwise "of" -- e.g. "Public Administrator for the Estate of John Doe".

Opening paragraph (the firm's Probate Estate template):
  I, **[Signer], [Capacity] of the [Estate Name]**, [County] Probate and Family Court Docket No. [Docket],
  hereby appoint **[attorney_in_fact]** as my true and lawful attorney in fact (my "Attorney")
  for me and in my name, in my capacity as [Capacity], to do all things necessary with respect to
  the sale of property situated at ...
Number agreement: I/We, my/our, me/us, capacity/capacities, and the capacity title itself.
"the" rule (firm-wide): "the" before the estate name unless it begins with "The".
Prints ESTATE NAME FONT, DOCKET and (when relevant) SALE AUTHORITY messages.
Run keep_together.py and check_agreement.py on the output afterwards.
"""
import sys as _sys
_sys.dont_write_bytecode = True   # never write __pycache__ into the skill folder
for _stream in (_sys.stdout, _sys.stderr):   # a Windows console can't print "→" etc. otherwise
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass
import json, re, sys
from build_entity_poa import (BASE_PT, COL_TWIPS, WITHOUT, body_p, by_line, cell, fit_size, join_list,
                              long_date_runs, run, signature_sections, write_poa)
from notary_block import (attorney_in_fact, agreement_runs, appointments, check_ps_date, counterparts_paragraph,
                          estate_capacity, notary_notes, notary_paragraph, ps_date_text, the_trust)

# Probate and Family Court division codes at the start of a Massachusetts docket number.
DIVISIONS = {"BA": "Barnstable", "BE": "Berkshire", "BR": "Bristol", "DU": "Dukes",
             "ES": "Essex", "FR": "Franklin", "HD": "Hampden", "HS": "Hampshire",
             "MI": "Middlesex", "NA": "Nantucket", "NO": "Norfolk", "PL": "Plymouth",
             "SU": "Suffolk", "WO": "Worcester"}
STOP_CAPACITIES = ("voluntary", "foreign")     # not handled -- the skill stops before building


def preposition_for(capacity):
    """'for' for a Public Administrator (a county officer acting for the estate); 'of' otherwise."""
    return "for" if capacity.strip().lower() == "public administrator" else "of"


def words(n):
    if n == 1:
        return dict(I="I", my="my", me="me", capacity="capacity")
    return dict(I="We", my="our", me="us", capacity="capacities")


def sig_name(estate):
    """Signature lines: the estate name without a leading 'The'."""
    return re.sub(r"^\s*the\s+", "", estate, flags=re.I)


def county_only(c):
    return re.sub(r"\s+county\s*$", "", c.strip(), flags=re.I)


def docket_check(county, docket):
    code = docket.strip()[:2].upper()
    div = DIVISIONS.get(code)
    if not div:
        return f"DOCKET: '{docket}' -- unknown division code '{code}'; check the docket number"
    if div.lower() != county_only(county).lower():
        return (f"DOCKET: WARNING -- docket prefix '{code}' is {div}, but the court county is "
                f"'{county}'; check which is right")
    return f"DOCKET: prefix '{code}' matches {div} County"


def signature_table(estate, signers, capacity, size):
    """Estate name (not bold) / blank / By:___ / Name, Capacity -- one column per fiduciary.
    Same layout as the entity and trustee tables."""
    none = '<w:{0} w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    borders = "".join(none.format(b) for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    def row(l="", r=""):
        return f'<w:tr><w:trPr><w:cantSplit/></w:trPr>{cell(l)}{cell(r)}</w:tr>'
    rows = []
    for g in range(0, len(signers), 2):
        grp = signers[g:g + 2]
        cols = [None, grp[0]] if len(signers) == 1 else (grp + [None])[:2]
        pick = lambda f: [f(t) if t else "" for t in cols]
        if g:
            rows += [row(), row()]
        rows.append(row(*pick(lambda t: run(estate, size=size))))       # not bold (Oct 29, 2026)
        rows.append(row())
        rows.append(row(*pick(lambda t: by_line())))
        rows.append(row(*pick(lambda t: run(f"{t}, {capacity}"))))
    return ('<w:tbl><w:tblPr><w:tblW w:w="9350" w:type="dxa"/>'
            f'<w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
            '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
            '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
            f'</w:tblPr><w:tblGrid><w:gridCol w:w="{COL_TWIPS}"/><w:gridCol w:w="{COL_TWIPS}"/></w:tblGrid>'
            + "".join(rows) + "</w:tbl>")


def build_body(v, size):
    s = v["signers"]
    w = words(len(s))
    cap = estate_capacity(v["capacity"], len(s))       # "Personal Representative(s)"
    prep = v.get("preposition") or preposition_for(v["capacity"])
    estate = the_trust(v["estate_name"])                # "the Estate of ..." / "The Estate of ..."
    tab = "<w:r><w:tab/></w:r>"
    opening = body_p(
        tab + run(f'{w["I"]}, ')
        + run(f"{join_list(s)}, {cap} {prep} {estate}", bold=True, noproof=True)
        + run(f', {county_only(v["probate_county"])} Probate and Family Court Docket No. {v["docket"].strip()}, '
              "hereby appoint ", noproof=True)
        + run(attorney_in_fact(v), bold=True)
        + run(f' as {w["my"]} true and lawful attorney in fact ({w["my"]} “Attorney”) for {w["me"]} and in '
              f'{w["my"]} name, in {w["my"]} {w["capacity"]} as {cap}, to do all things necessary '
              "with respect to the sale of property situated at ")
        + run(v["address"], bold=True, noproof=True)
        + agreement_runs(v)                           # "...Agreement dated [P&S date], with "
        + run(v["buyers"], noproof=True)
        + run(" (the “P&S”), for not less than ")
        + run("$" + v["price"], bold=True, noproof=True)
        + run(f' (“Sales Price”), hereby ratifying and affirming that which {w["my"]} Attorney shall '
              "lawfully do or cause to be done by virtue of the powers herein conferred."))
    out = "<w:p/>" + opening + body_p() + body_p(tab + run(WITHOUT)) + body_p()
    if v.get("include_incapacity", True):
        out += body_p(tab + run(f'THIS Power of Attorney shall not be affected by {w["my"]} subsequent '
                                "disability or incapacity.")) + body_p()
    appts = appointments(v, v["signers"])
    if len(appts) > 1:
        out += counterparts_paragraph() + body_p()
    out += signature_sections(dict(v, signers=v["signers"]), appts,
                              lambda subset: signature_table(sig_name(v["estate_name"]), subset, v["capacity"], size),
                              lambda a: notary_paragraph(dict(a, capacity={"type": "estate", "capacity": v["capacity"], "estate_name": v["estate_name"], "preposition": prep})))
    return out


def main(template, values_path, out):
    v = json.load(open(values_path, encoding="utf-8"))
    assert v.get("signers"), "at least one signer is required"
    for k in ("capacity", "estate_name", "probate_county", "docket", "address",
              "buyers", "price"):
        assert v.get(k), f"{k} is required"
    check_ps_date(v)
    low = v["capacity"].lower()
    assert not any(x in low for x in STOP_CAPACITIES), \
        f"capacity '{v['capacity']}' is not handled by this skill -- stop and tell the user"
    size, fits = fit_size(sig_name(v["estate_name"]), bold=False)
    write_poa(template, build_body(v, size), out)
    note = "" if fits else " -- WARNING: still too long at the minimum size; the name will wrap"
    print(f"ESTATE NAME FONT: {size:g}pt (base {BASE_PT:g}pt){note}")
    print(f"P&S DATE: {ps_date_text(v)}")
    print(docket_check(v["probate_county"], v["docket"]))
    for line in notary_notes(appointments(v, v["signers"])):
        print(line)
    if "special" in low:
        print("SALE AUTHORITY: WARNING -- a Special Personal Representative does not have the "
              "statutory power to sell real estate (G.L. c. 190B, s. 3-715(b) omits (23 1/2)); "
              "confirm the appointment order or a court decree authorizes this sale")


if __name__ == "__main__":
    main(*sys.argv[1:4])
