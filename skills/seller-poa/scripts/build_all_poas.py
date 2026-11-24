#!/usr/bin/env python3
"""build_all_poas.py -- build every Seller's Limited POA for one P&S: one POA per PRINCIPAL
. Used for every transaction -- one principal or several.

Usage: python build_all_poas.py ASSETS_DIR VALUES.json OUT_DIR [--mode same|separate]
                                [--docx-skill DOCX_SKILL_DIR]

ASSETS_DIR      the skill's assets folder (Form 2 and Form 4).
OUT_DIR         where the finished POAs are saved -- the folder of the P&S.
--mode          signature page placement for keep_together.py (default same).
--docx-skill    the docx skill's base directory; when given, every POA is also run through its
                scripts/office/validate.py.

A PRINCIPAL is whoever the POA speaks for:
  individual -- ALL the individual sellers together, in one POA (any number)
  trustee    -- the trustee(s) of ONE trust (two different trusts -> two principals)
  estate     -- the fiduciary(ies) of ONE estate
  entity     -- ONE entity (its signers sign for it)
A P&S with sellers of more than one type, or with several trusts / estates / entities, gets one
POA per principal, each built by that type's own builder, with the same Buyer, Price, P&S date,
signing date, and notary location unless a principal overrides them.

VALUES.json:
  {
    "attorney_in_fact": "Alex Q. Attorney, Esq.",          -- the firm setting (required)
    "address": "123 Sample Street, Hingham, MA 02043",       -- shared by every POA
    "ps_date": "2026-09-01", "buyers": "William F. Buyer and Jessica Buyer", "price": "500,000.00",
                                                           -- ps_date may be "2026-09" / "2026" / ""
                                                              (blanks), or set "include_ps_date":
                                                              false to leave the P&S date out
    "signing_date": "2026-09-16",                          -- or "2026-09", "2026", ""
    "notary_state": "Massachusetts", "county": "Plymouth", -- notary location for every POA
    "principals": [
      {"type": "individual", "sellers": ["John Smith"], "pronoun": "he"},
      {"type": "trustee", "trust_name": "Smith Family Trust", "trustees": ["Jane Smith"],
       "trust_date": "2015-03-03", "pronoun": "she"},
      {"type": "entity", "entity_name": "Smith Holdings, LLC", "state": "Massachusetts",
       "entity_type": "limited liability company",
       "signers": [{"name": "Jane Smith", "title": "Manager"}]},
      {"type": "estate", "estate_name": "Estate of Mary Smith", "signers": ["Jane Smith"],
       "capacity": "Personal Representative", "probate_county": "Plymouth",
       "docket": "PL24P1234EA"}
    ]
  }
Each principal takes the keys of its own builder (Steps 4A, 4D, 4E, 4F of SKILL.md) plus "type".
A principal may override any shared key (e.g. its own "signing_date", "notary_state", "county",
or "appointments"). The shared notary state is "notary_state". For an entity, "state" is its
state of organization, so an entity overrides the notary state only with "notary_state"; any
other principal may use "notary_state" or "state". The script passes the notary state to each
builder under the key that builder uses.

Notary appointments across the POAs: when the signers of the POAs will not
all sign at the same notary appointment, list the appointments ONCE, at the top level:
  "appointments": [
    {"signers": ["Grace Sample"], "state": "Massachusetts", "county": "Plymouth",
     "signing_date": "2026-11-30"},
    {"signers": ["Henry J. Sample"], "state": "Massachusetts", "county": "Norfolk",
     "signing_date": "2026-12-02"}],
  "pronouns": {"Grace Sample": "she", "Henry J. Sample": "he"}
Signers are named exactly as in the POAs (an entity signer by "name"); every signer of every POA
must be in exactly one appointment, and a person who signs two POAs is listed once (both POAs
are signed at that appointment). An appointment's missing state / county / signing_date falls
back to the shared value. The script gives each POA the appointments its own signers attend:
all of a POA's signers in one appointment -> that POA's own signing_date / state / county (no
counterparts); its signers split across appointments -> counterpart signature pages, as for a
single POA. "pronouns" gives the pronoun for each person who is alone in a POA's notary block
(a principal's own "pronoun" also works for a one-signer POA). A principal's own
"appointments" (if any) is used instead of the top-level list.

File names:
  one POA          -> "Power of Attorney. Seller. [Address].docx"
  two or more POAs -> "Power of Attorney. Seller. [Address]. [Principal's Name].docx"
  Principal's name: the individuals' names ("John Smith", "John Smith and Jane Smith",
  "A, B, and C"), the trust name, the entity name, or the estate name, as in the P&S. Characters
  Windows does not allow in file names (\\ / : * ? " < > |) become "-"; a trailing period is
  dropped so the name never ends in "..docx".

Safety: every POA is built and checked (builder, keep_together.py, check_agreement.py, and the
validator when --docx-skill is given) in a temporary folder first. Only when EVERY POA passes are
they copied to OUT_DIR, so a failure never leaves a partial set behind. An existing file with the
same name is not overwritten -- the script stops and says so.
Prints one block per POA and a final line: ALL POAS: OK (n built) or ALL POAS: STOPPED -- ...
"""
import sys as _sys
_sys.dont_write_bytecode = True   # never write __pycache__ into the skill folder
import argparse, json, os, re, shutil, subprocess, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
TYPES = ("individual", "trustee", "entity", "estate")
SHARED = ("attorney_in_fact", "address", "ps_date", "include_ps_date", "buyers", "price", "signing_date",
          "notary_state", "county", "include_incapacity")
BAD_CHARS = re.compile(r'[\\/:*?"<>|]')


def join_names(names):
    names = list(names)
    if len(names) <= 2:
        return " and ".join(names)
    return ", ".join(names[:-1]) + ", and " + names[-1]


def principal_name(p):
    t = p["type"]
    if t == "individual":
        return join_names(p["sellers"])
    return {"trustee": p.get("trust_name"), "entity": p.get("entity_name"),
            "estate": p.get("estate_name")}[t]


def file_name(address, name=None):
    parts = ["Power of Attorney", "Seller", address.strip()] + ([name.strip()] if name else [])
    base = ". ".join(BAD_CHARS.sub("-", x).strip() for x in parts)
    return base.rstrip(". ") + ".docx"


def signer_names(p):
    t = p["type"]
    if t == "individual":
        return list(p.get("sellers") or [])
    if t == "trustee":
        return list(p.get("trustees") or [])
    if t == "entity":
        return [s["name"] if isinstance(s, dict) else s for s in p.get("signers") or []]
    return list(p.get("signers") or [])


def check_principals(principals):
    """Problems with the list of principals (before anything is built)."""
    errs = []
    if not principals:
        return ["no principals -- list at least one under \"principals\""]
    for i, p in enumerate(principals, 1):
        t = p.get("type")
        if t not in TYPES:
            errs.append(f"principal {i}: \"type\" must be one of {', '.join(TYPES)} (got {t!r})")
            continue
        if not principal_name(p):
            key = {"individual": "sellers", "trustee": "trust_name", "entity": "entity_name",
                   "estate": "estate_name"}[t]
            errs.append(f"principal {i} ({t}): \"{key}\" is missing")
        if not signer_names(p):
            errs.append(f"principal {i} ({t}): no signers")
    if errs:
        return errs
    if sum(p["type"] == "individual" for p in principals) > 1:
        errs.append("more than one individual principal -- every individual seller goes in ONE "
                    "individual POA (one \"sellers\" list)")
    names = [principal_name(p).strip().lower() for p in principals]
    dups = sorted({n for n in names if names.count(n) > 1})
    if dups:
        errs.append("the same principal is listed twice: " + "; ".join(dups))
    return errs


def all_signers(principals):
    """Every distinct signer across all POAs, in POA order (a person in two POAs once)."""
    out = []
    for p in principals:
        for s in signer_names(p):
            if s not in out:
                out.append(s)
    return out


def check_appointments(shared, principals):
    """Problems with the top-level "appointments" (notary appointments across ALL the POAs).
    Every signer of every POA must be in exactly one appointment; a person
    who signs two POAs is listed once and signs both at that appointment."""
    raw = shared.get("appointments")
    if not raw:
        return []
    from notary_block import check_county, date_parts, state_name
    names, seen, errs = all_signers(principals), [], []
    for i, a in enumerate(raw, 1):
        sub = list(a.get("signers") or [])
        if not sub:
            errs.append(f"appointment {i}: no signers")
        for s in sub:
            if s not in names:
                errs.append(f"appointment {i}: '{s}' is not a signer of any POA (signers: "
                            + "; ".join(names) + ")")
            elif s in seen:
                errs.append(f"'{s}' is in more than one notary appointment")
            else:
                seen.append(s)
        try:
            st = a.get("state", shared.get("notary_state") or shared.get("state"))
            state_name(st)
            check_county(st, a.get("county", shared.get("county")))
            date_parts(a.get("signing_date", shared.get("signing_date")))
        except (ValueError, AssertionError) as e:
            errs.append(f"appointment {i}: {e}")
    missing = [s for s in names if s not in seen]
    if missing:
        errs.append("not in any notary appointment: " + "; ".join(missing))
    if errs:
        return errs
    need = []                                      # alone in a POA's notary block, no pronoun
    for p in principals:
        if p.get("appointments"):
            continue
        ns = signer_names(p)
        for a in raw:
            if any(s in ns for s in a["signers"]):
                cut = appointment_for(shared, p, a, ns)
                if len(cut["signers"]) == 1 and not cut.get("pronoun") \
                        and cut["signers"][0] not in need:
                    need.append(cut["signers"][0])
    if need:
        errs.append("pronoun needed (he / she / they / blank -- ask the user, never guess) for "
                    "each person alone in a POA's notary block: " + "; ".join(need)
                    + " -- add it under \"pronouns\"")
    return errs


def appointment_for(shared, p, a, names):
    """Appointment A (from the top-level list) cut down to principal P's signers NAMES, with the
    shared values filled in and the pronoun for a one-signer notary block."""
    sub = [s for s in names if s in a["signers"]]
    out = {"signers": sub,
           "state": a.get("state") or shared.get("notary_state") or shared.get("state")
                    or "Massachusetts",
           "county": a.get("county", shared.get("county")),
           "signing_date": a.get("signing_date", shared.get("signing_date"))}
    if len(sub) == 1:
        pr = ((shared.get("pronouns") or {}).get(sub[0])
              or (a.get("pronoun") if len(a["signers"]) == 1 else None)
              or (p.get("pronoun") if len(names) == 1 else None))
        if pr:
            out["pronoun"] = pr
    return out


def builder_values(shared, p):
    v = {k: shared[k] for k in SHARED if k in shared}
    v.update({k: val for k, val in p.items() if k != "type"})
    v.pop("notary_state", None)
    shared_notary = shared.get("notary_state") or shared.get("state") or "Massachusetts"
    if p["type"] == "entity":                     # "state" stays the state of organization
        v["notary_state"] = p.get("notary_state") or shared_notary
    else:                                         # a principal's own "state" is its notary state
        v["state"] = p.get("notary_state") or p.get("state") or shared_notary
    raw = shared.get("appointments")
    if raw and not p.get("appointments"):
        # Cross-POA appointments: give this POA only the appointments its own signers attend.
        names = signer_names(p)
        mine = [appointment_for(shared, p, a, names) for a in raw
                if any(s in names for s in a["signers"])]
        mine.sort(key=lambda a: names.index(a["signers"][0]))
        if len(mine) == 1:                        # all its signers together: plain values
            a = mine[0]
            v["signing_date"], v["county"] = a["signing_date"], a["county"]
            v["notary_state" if p["type"] == "entity" else "state"] = a["state"]
            if "pronoun" in a:
                v["pronoun"] = a["pronoun"]
            v.pop("appointments", None)
        else:                                     # its signers split: counterpart pages
            v["appointments"] = mine
    elif len(signer_names(p)) == 1 and not v.get("pronoun"):
        pr = (shared.get("pronouns") or {}).get(signer_names(p)[0])
        if pr:
            v["pronoun"] = pr
    return v


def appointments_summary(shared, principals):
    """Lines describing the cross-POA appointments, for the user."""
    from notary_block import appt_label
    raw = shared.get("appointments") or []
    if not raw:
        return []
    lines = [f"NOTARY APPOINTMENTS (all POAs): {len(raw)}"]
    for i, a in enumerate(raw):
        full = {"state": a.get("state") or shared.get("notary_state") or shared.get("state"),
                "county": a.get("county", shared.get("county")),
                "signing_date": a.get("signing_date", shared.get("signing_date"))}
        who = []
        for s in a["signers"]:
            pois = [principal_name(p) for p in principals if s in signer_names(p)]
            who.append(f"{s} ({'; '.join(pois)} POA{'s' if len(pois) > 1 else ''})")
        lines.append(f"  {chr(65 + i)} -- {appt_label(full)}: " + ", ".join(who))
    return lines


def run(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=HERE,
                       **({"creationflags": 0x08000000} if os.name == "nt" else {}))
    out = (r.stdout + r.stderr).strip()
    return r.returncode, out


def build_one(assets, p, v, out, mode, docx_skill):
    """Build and check one POA at OUT. Returns (ok, lines)."""
    t = p["type"]
    form2 = os.path.join(assets, "POA. Seller. Form 2.docx")
    form4 = os.path.join(assets, "POA. Seller. Form 4.docx")
    vpath = out[:-5] + ".values.json"
    with open(vpath, "w") as f:
        json.dump(v, f, indent=2)
    if t == "individual":
        cmd = [sys.executable, "build_individual_poa.py", assets, vpath, out]
        template = form2 if len(v["sellers"]) == 1 else form4
    else:
        cmd = [sys.executable, f"build_{t}_poa.py", form2, vpath, out]
        template = form2
    lines = []
    steps = [("BUILD", cmd),
             ("KEEP TOGETHER", [sys.executable, "keep_together.py", out, "--mode", mode,
                                "--address", v["address"]]),
             ("AGREEMENT", [sys.executable, "check_agreement.py", out])]
    if docx_skill:
        steps.append(("VALIDATE", [sys.executable,
                                   os.path.join(docx_skill, "scripts", "office", "validate.py"),
                                   out, "--original", template]))
    for label, c in steps:
        code, text = run(c)
        if text:
            lines += ["  " + ln for ln in text.splitlines()]
        if code != 0 or (label == "BUILD" and not os.path.exists(out)):
            lines.append(f"  {label}: FAILED")
            return False, lines
    return True, lines


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("assets")
    ap.add_argument("values")
    ap.add_argument("out_dir")
    ap.add_argument("--mode", choices=("same", "separate"), default="same")
    ap.add_argument("--docx-skill")
    a = ap.parse_args()
    a.assets, a.values, a.out_dir = (os.path.abspath(x) for x in (a.assets, a.values, a.out_dir))

    shared = json.load(open(a.values, encoding="utf-8"))
    principals = shared.get("principals") or []
    for k in ("address", "buyers", "price"):
        if not shared.get(k):
            sys.exit(f"ALL POAS: STOPPED -- \"{k}\" is required (shared by every POA); nothing built")
    try:                                          # the firm setting: who the POA appoints
        from notary_block import attorney_in_fact
        attorney_in_fact(shared)
    except ValueError as e:
        sys.exit(f"ALL POAS: STOPPED -- {e}; nothing built")
    try:                                          # P&S date: full, partly blank, blank, or left out
        from notary_block import check_ps_date
        check_ps_date(shared)
    except (AssertionError, ValueError) as e:
        sys.exit(f"ALL POAS: STOPPED -- P&S date: {e}; nothing built")
    errs = check_principals(principals)
    if errs:
        sys.exit("ALL POAS: STOPPED -- nothing built:\n  " + "\n  ".join(errs))
    errs = check_appointments(shared, principals)
    if errs:
        sys.exit("ALL POAS: STOPPED -- notary appointments; nothing built:\n  " + "\n  ".join(errs))

    multi = len(principals) > 1
    names = [file_name(shared["address"], principal_name(p) if multi else None) for p in principals]
    if len(set(n.lower() for n in names)) < len(names):
        sys.exit("ALL POAS: STOPPED -- two POAs would get the same file name: "
                 + "; ".join(names) + "; nothing built")
    existing = [n for n in names if os.path.exists(os.path.join(a.out_dir, n))]
    if existing:
        sys.exit("ALL POAS: STOPPED -- already in the folder (not overwritten; rename or remove "
                 "it, or ask the user): " + "; ".join(existing) + "; nothing built")

    # Someone signing in more than one capacity (e.g. individually AND as trustee) signs each POA.
    who = {}
    for p in principals:
        for s in signer_names(p):
            who.setdefault(s, []).append(principal_name(p))
    twice = {s: ps for s, ps in who.items() if len(ps) > 1}

    tmp = tempfile.mkdtemp(prefix="poas_")
    ok_all, report = True, []
    print(f"PRINCIPALS: {len(principals)} -> {len(principals)} POA{'s' if multi else ''}")
    for ln in appointments_summary(shared, principals):
        print(ln)
    for i, (p, name) in enumerate(zip(principals, names), 1):
        out = os.path.join(tmp, name)
        ok, lines = build_one(a.assets, p, builder_values(shared, p), out, a.mode, a.docx_skill)
        print(f"POA {i} of {len(principals)} -- {p['type']}: {principal_name(p)}")
        print(f"  FILE: {name}")
        for ln in lines:
            print(ln)
        ok_all &= ok
        report.append((name, out, ok))
    for s, ps in twice.items():
        print(f"SIGNS MORE THAN ONE POA: {s} -- {'; '.join(ps)} (one signature on each POA)")
    if not ok_all:
        bad = [n for n, _, ok in report if not ok]
        sys.exit(f"ALL POAS: STOPPED -- {len(bad)} of {len(report)} failed ({'; '.join(bad)}); "
                 f"NOTHING saved to the folder. Drafts for review: {tmp}")
    os.makedirs(a.out_dir, exist_ok=True)
    for name, out, _ in report:
        shutil.copy2(out, os.path.join(a.out_dir, name))
    shutil.rmtree(tmp, ignore_errors=True)
    print(f"ALL POAS: OK ({len(report)} built) -> {a.out_dir}")


if __name__ == "__main__":
    main()
