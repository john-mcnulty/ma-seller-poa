#!/usr/bin/env python3
"""check_agreement.py -- check singular/plural agreement in a finished Seller POA (all seller types).

Usage: python check_agreement.py POA.docx

Looks at the body paragraphs above "EXECUTED as a sealed instrument" (the opening, powers and
incapacity paragraphs -- not the signature lines or the notary block, where "me"/"my" refer to
the notary). Works out who is speaking from the opening word and flags any word that disagrees:
  "I,"  (one seller / one trustee)      -> no our, us, Trustees, capacities
  "We," (several sellers / trustees)    -> no my, me, Trustee, capacity
  anything else (entity seller)         -> no my, me, our, us (an entity speaks as "its")
Also checks the trust "the" rule: a trust name not starting with "The" must have "the" before it
(and never "the The"), in the opening paragraph and in the notary capacity clause.
Notary block: one signer -> "person whose name is signed" + he/she/they or the "_____" blank
(+ "as Trustee"); two or more ->
"persons whose names are signed" + they (+ "as Trustees"); "I," openings need one signer and "We,"
openings two or more; no "{{...}}" placeholder may be left anywhere.
Prints one line: AGREEMENT: OK, or AGREEMENT: PROBLEM -- <details>. Exit code 1 on a problem.
"""
import sys as _sys
_sys.dont_write_bytecode = True   # never write __pycache__ into the skill folder
for _stream in (_sys.stdout, _sys.stderr):   # a Windows console can't print "→" etc. otherwise
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass
import re, sys, zipfile
from notary_block import (CAPACITY_HEADS, COUNTERPARTS, PRONOUN_BLANK, plural_capacity,
                          singular_capacity)

BAD = {
    "singular": ["our", "us", "Trustees", "capacities"],
    "plural": ["my", "me", "Trustee", "capacity"],
    "entity": ["my", "me", "our", "us"],
}


def para_texts(xml):
    paras = re.findall(r"<w:p(?: [^>]*)?>.*?</w:p>", xml, re.S)
    return ["".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p)).strip() for p in paras]


def body_text(xml):
    texts = para_texts(xml[: xml.index("EXECUTED as a sealed instrument")])
    return [t for t in texts if t and "LIMITED POWER OF ATTORNEY" not in t]


def the_rule(text, where):
    """Firm-wide: 'the' before a trust name unless it begins with 'The'; never 'the The'."""
    out = []
    m = re.search(r"as Trustees? of (\S+)", text)
    if m and m.group(1) not in ("the", "The"):
        out.append(f"trust name not preceded by 'the' ({where})")
    if re.search(r"\bthe The\b", text):
        out.append(f"double 'the' before the trust name ({where})")
    return out


def estate_problems(opening, notary, mode):
    """Probate-estate POAs: the 'the' rule before the estate name, and the capacity title matching
    the number of fiduciaries in the opening and the notary clause."""
    out = []
    for text, where in [(opening, "opening")] + [(n, "notary") for n in (notary if isinstance(notary, list) else [notary])]:
        if re.search(r"\b(?:of|for) Estate of\b", text):
            out.append(f"estate name not preceded by 'the' ({where})")
        if re.search(r"\bthe The\b", text):
            out.append(f"double 'the' before the estate name ({where})")
    m = re.search(r"in (?:my|our) capacit(?:y|ies) as (.+?), to do all things", opening)
    if not m:
        return out + ["estate opening: can't find 'in my/our capacity as ...'"]
    cap = m.group(1)
    if not re.search(rf", {re.escape(cap)} (?:of|for) ", opening):
        out.append(f"estate opening: capacity '{cap}' differs between the parties and 'capacity as'")
    heads = cap.replace(",", "").split()
    if mode == "singular" and any(w in CAPACITY_HEADS.values() for w in heads):
        out.append(f"'I,' opening but plural capacity '{cap}'")
    if mode == "plural" and not any(w in CAPACITY_HEADS.values() for w in heads):
        out.append(f"'We,' opening but singular capacity '{cap}'")
    for k, n in enumerate(notary if isinstance(notary, list) else [notary] if notary else [], 1):
        # each notary block names its own signers: singular capacity for one, plural for several
        base = singular_capacity(cap)
        want = base if "person whose name is signed" in n else plural_capacity(base)
        if not re.search(rf"\bas {re.escape(want)} (?:of|for) ", n):
            out.append(f"notary {k}: capacity does not read 'as {want} of/for ...'")
    return out


def notary_problems(xml, mode):
    """Every notary block must agree with itself (person/persons, he/she/they, Trustee/Trustees);
    together they must match the opening: "I," -> one block with one signer; "We," -> two or more
    signers in all (one block with several, or several blocks -- separate notary appointments).
    Several blocks also need one EXECUTED line each and the counterparts paragraph."""
    out = []
    notes = [t for t in para_texts(xml) if "before me, the undersigned notary public" in t]
    if not notes:
        return ["no notary acknowledgment found"]
    many_blocks = len(notes) > 1
    singles = 0
    for k, n in enumerate(notes, 1):
        tag = f"notary {k}" if many_blocks else "notary"
        one = "person whose name is signed" in n
        many = "persons whose names are signed" in n
        if one == many:
            out.append(f"{tag} block: can't tell one signer from several (person/persons)")
            continue
        singles += one
        pron = re.search(r"acknowledged to me that (\w+) signed it", n)
        pron = pron.group(1) if pron else None
        if one and pron not in ("he", "she", "they", PRONOUN_BLANK):
            out.append(f"{tag}: one signer but pronoun '{pron}'")
        if many and pron != "they":
            out.append(f"{tag}: several signers but pronoun '{pron}'")
        if one and re.search(r"\bas Trustees of\b", n):
            out.append(f"{tag}: one signer but 'Trustees'")
        if many and re.search(r"\bas Trustee of\b", n):
            out.append(f"{tag}: several signers but 'Trustee'")
        if "{{" in n or ".." in n.replace("...", ""):
            out.append(f"{tag}: placeholder or double period")
        out += the_rule(n, tag)
    if mode == "singular" and (many_blocks or singles != len(notes)):
        out.append("notary: 'I,' opening but several signers")
    if mode == "plural" and not many_blocks and singles:
        out.append("notary: 'We,' opening but one signer")
    if many_blocks:
        executed = [t for t in para_texts(xml) if t.startswith("EXECUTED as a sealed instrument")]
        if len(executed) != len(notes):
            out.append(f"{len(notes)} notary blocks but {len(executed)} EXECUTED lines")
        if COUNTERPARTS not in " ".join(para_texts(xml)):
            out.append("several notary appointments but no counterparts paragraph")
    return out


def main(path):
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    texts = body_text(xml)
    opening = texts[0]
    first = opening.split(",")[0].strip()
    mode = {"I": "singular", "We": "plural"}.get(first, "entity")
    problems = []
    body = " ".join(texts)
    for w in BAD[mode]:
        hits = re.findall(rf"(?<![\w’']){w}(?![\w’'])", body)
        if hits:
            problems.append(f"'{w}' x{len(hits)} ({mode} seller)")
    if "Family Court Docket No." in opening:
        note = [t for t in para_texts(xml) if "before me, the undersigned notary public" in t]
        problems += estate_problems(opening, note, mode)
    else:
        problems += the_rule(opening, "opening")
    problems += notary_problems(xml, mode)
    left = sorted(set(re.findall(r"\{\{[^}]*\}\}", " ".join(para_texts(xml)))))
    if left:
        problems.append("placeholder left in the document: " + ", ".join(left))
    n_notary = sum("before me, the undersigned notary public" in t for t in para_texts(xml))
    if problems:
        print("AGREEMENT: PROBLEM -- " + "; ".join(problems))
        sys.exit(1)
    print((f"AGREEMENT: OK ({mode} -- opening '{first},'" if mode != "entity" else "AGREEMENT: OK (entity -- 'its'")
          + (f"; {n_notary} notary blocks agree)" if n_notary > 1 else "; notary block agrees)"))


if __name__ == "__main__":
    main(sys.argv[1])
