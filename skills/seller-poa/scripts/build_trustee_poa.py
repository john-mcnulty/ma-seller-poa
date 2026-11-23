#!/usr/bin/env python3
"""build_trustee_poa.py -- build a Trustee-Seller POA from "POA. Seller. Form 2.docx".

Usage: python build_trustee_poa.py TEMPLATE.docx VALUES.json OUTPUT.docx
Needs build_entity_poa.py in the same folder (shared helpers).

VALUES.json keys:
  trust_name:   trust name exactly as in the P&S, e.g. "Doe Family Trust" or "The Doe Family Trust"
  trustees:     ["Jane Doe", "John Doe"]       (every trustee, in signing order)
  trust_date:   Declaration of Trust date as YYYY-MM-DD   -> "March 3rd, 2015"
  include_trust_date: true / false             (default true). false leaves
                the whole ", u/d/t dated [date]" phrase out of the opening paragraph (one comma
                remains: "...of the Doe Family Trust**, hereby appoint..."); trust_date is then
                not needed and is ignored.
  trust_date_source: optional -- where the date came from (e.g. "P&S SELLER entry", "Trustee's
                Certificate.pdf"); printed on the TRUST DATE line for the report.
  address, buyers, price
  ps_date:      P&S date as YYYY-MM-DD        -> "September 1st, 2026"; "YYYY-MM", "YYYY", or ""
                print blanks for the unknown parts ("September _____, 2026")
  include_ps_date: false leaves "dated [P&S date]" out of the opening (default true)
  signing_date: POA signing date -> EXECUTED "September 16th, 2026" ("YYYY-MM", "YYYY", or ""
                for a partly or wholly blank date -- see notary_block.py)
  include_incapacity: true / false             (default true)
  state:        notary state (default Massachusetts)
  county:       notary county, e.g. "Plymouth" ("" or left out -> blank line)
  appointments: optional -- separate notary appointments (see notary_block.py)
  pronoun:      "he" / "she" / "they" / "blank" (notary block; required for ONE trustee)

Number agreement (one trustee vs. several) is applied throughout the body:
  I / We, my / our, me / us, Trustee / Trustees, capacity / capacities.
"the" rule (firm-wide): outside the signature lines, a lowercase "the" goes before the trust
name unless the name already starts with "The". The signature lines show the name as-is.
The notary block comes from notary_block.py (capacity "as Trustee(s) of [the] Trust").
Needs build_entity_poa.py and notary_block.py in the same folder.
Run keep_together.py on the output afterwards.
"""
import sys as _sys
_sys.dont_write_bytecode = True   # never write __pycache__ into the skill folder
import json, re, sys
from build_entity_poa import (BASE_PT, COL_TWIPS, WITHOUT, body_p, by_line, cell, fit_size, join_list,
                              long_date_runs, run, signature_sections, write_poa)
from notary_block import (attorney_in_fact, agreement_runs, appointments, check_ps_date, counterparts_paragraph,
                          notary_notes, notary_paragraph, ps_date_text, the_trust)


def words(n):
    """Number-agreement vocabulary for one trustee (n == 1) or several."""
    if n == 1:
        return dict(I="I", my="my", me="me", Trustee="Trustee", capacity="capacity")
    return dict(I="We", my="our", me="us", Trustee="Trustees", capacity="capacities")


def signature_table(trust, trustees, size):
    """Trust name (NOT bold) / blank / By:___ / Name, Trustee -- one column
    per trustee.
    Same layout as the entity table: one trustee -> right column; several -> left, right,
    then new groups below."""
    none = '<w:{0} w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    borders = "".join(none.format(b) for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    def row(l="", r=""):
        return f'<w:tr><w:trPr><w:cantSplit/></w:trPr>{cell(l)}{cell(r)}</w:tr>'
    rows = []
    for g in range(0, len(trustees), 2):
        grp = trustees[g:g + 2]
        cols = [None, grp[0]] if len(trustees) == 1 else (grp + [None])[:2]
        pick = lambda f: [f(t) if t else "" for t in cols]
        if g:
            rows += [row(), row()]
        rows.append(row(*pick(lambda t: run(trust, size=size))))
        rows.append(row())
        rows.append(row(*pick(lambda t: by_line())))
        rows.append(row(*pick(lambda t: run(f"{t}, Trustee"))))
    return ('<w:tbl><w:tblPr><w:tblW w:w="9350" w:type="dxa"/>'
            f'<w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
            '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
            '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
            f'</w:tblPr><w:tblGrid><w:gridCol w:w="{COL_TWIPS}"/><w:gridCol w:w="{COL_TWIPS}"/></w:tblGrid>'
            + "".join(rows) + "</w:tbl>")


def include_trust_date(v):
    x = v.get("include_trust_date", True)
    if isinstance(x, str):
        x = x.strip().lower() not in ("false", "no", "0", "")
    return bool(x)


def udt_runs(v):
    """', u/d/t dated March 3rd, 2015' -- or nothing at all when the user chose to leave it out
   , so the opening reads '...of the Doe Family Trust, hereby appoint...'
    with a single comma."""
    if not include_trust_date(v):
        return ""
    return run(", u/d/t dated ") + long_date_runs(v["trust_date"])


def build_body(v, size):
    t = v["trustees"]
    w = words(len(t))
    tab = "<w:r><w:tab/></w:r>"
    opening = body_p(
        tab + run(f'{w["I"]}, ')
        + run(f'{join_list(t)}, as {w["Trustee"]} of {the_trust(v["trust_name"])}', bold=True, noproof=True)
        + udt_runs(v)
        + run(", hereby appoint ") + run(attorney_in_fact(v), bold=True)
        + run(f' as {w["my"]} true and lawful attorney in fact ({w["my"]} “Attorney”) for {w["me"]} and in '
              f'{w["my"]} name, in {w["my"]} {w["capacity"]} as {w["Trustee"]}, to do all things necessary '
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
    appts = appointments(v, v["trustees"])
    if len(appts) > 1:
        out += counterparts_paragraph() + body_p()
    out += signature_sections(dict(v, signers=v["trustees"]), appts,
                              lambda subset: signature_table(v["trust_name"], subset, size),
                              lambda a: notary_paragraph(dict(a, capacity={"type": "trust", "trust_name": v["trust_name"]})))
    return out


def main(template, values_path, out):
    v = json.load(open(values_path, encoding="utf-8"))
    assert v.get("trustees"), "at least one trustee is required"
    for k in ("trust_name", "address", "buyers", "price"):
        assert v.get(k), f"{k} is required"
    check_ps_date(v)
    if include_trust_date(v) and not v.get("trust_date"):
        sys.exit("TRUST DATE: STOPPED -- no POA built. \"trust_date\" is missing. Ask the user for the "
                 "Declaration of Trust date, or (their choice) set \"include_trust_date\": false to "
                 "leave the u/d/t phrase out (SKILL.md Step 3).")
    size, fits = fit_size(v["trust_name"], bold=False)   # 11pt down to 8pt, measured not bold
    body = build_body(v, size)
    if not include_trust_date(v):
        txt = "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", body))
        assert "u/d/t" not in txt and ",," not in txt and ", ," not in txt, "u/d/t phrase not fully removed"
    write_poa(template, body, out)
    note = "" if fits else " -- WARNING: still too long at the minimum size; the name will wrap"
    print(f"TRUST NAME FONT: {size:g}pt (base {BASE_PT:g}pt){note}")
    print(f"P&S DATE: {ps_date_text(v)}")
    src = f" -- source: {v['trust_date_source']}" if v.get("trust_date_source") else ""
    if include_trust_date(v):
        print(f"TRUST DATE: included -- u/d/t dated {v['trust_date']}{src}")
    else:
        print("TRUST DATE: left out (user's choice) -- no u/d/t phrase in the POA" + src)
    for line in notary_notes(appointments(v, v["trustees"])):
        print(line)


if __name__ == "__main__":
    main(*sys.argv[1:4])
