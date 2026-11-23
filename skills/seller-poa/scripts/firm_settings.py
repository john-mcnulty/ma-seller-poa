#!/usr/bin/env python3
"""firm_settings.py -- the firm's standing choices for the Seller POA skill, in one place.

    python firm_settings.py --show [SETTING FLAGS]
    python firm_settings.py --location-options [SETTING FLAGS] [--property-county Norfolk]
    python firm_settings.py --resolve-location "Main office" [SETTING FLAGS]

SETTING FLAGS (each optional; SKILL.md passes the plugin settings here):
    --attorney-in-fact    "Alex Q. Attorney, Esq."
    --signing-locations   "Main office = Plymouth County, MA; Boston office = Suffolk County, MA"
    --signature-page      same-page | separate-page | ask
    --notary-state        Massachusetts

WHERE EACH VALUE COMES FROM, in this order (the first usable one wins):
  1. the flag -- in a Claude Code plugin install, SKILL.md fills it from the plugin settings
     (`${user_config.*}`). A value that is empty, or that is still a literal placeholder because
     the setting was never set, counts as missing;
  2. the JSON file named by the SELLER_POA_SETTINGS environment variable, if set;
  3. firm-settings.json in the skill folder (next to SKILL.md) -- the place to put the firm's
     values when the skill is uploaded on its own (e.g. a ZIP added to Claude in the browser);
  4. the built-in default -- EXCEPT attorney_in_fact, which has no default: every POA appoints
     that person, so it must be chosen by the firm. Without it --show reports an error and
     build_all_poas.py refuses to build.

SIGNING LOCATIONS -- the places the firm's signings usually happen (its office or offices), so
the notary location question (SKILL.md Step 3, question 2) is a pick from a list instead of a
state and a county. Each entry is "Label = County, State":
    "Main office (100 Example St, Hingham) = Plymouth County, MA; Boston office = Suffolk County, MA"
The FIRST entry is the default (offered first, as Recommended). Entries are separated by ";" or
new lines; "County" is optional; the state may be a name or a postal code. A Massachusetts county
must be one of the 14 (a close misspelling is corrected). "= State" alone (no county) is allowed:
the notary block then prints a blank county line. In firm-settings.json the list may also be
given as objects: [{"label": "Main office", "county": "Plymouth", "state": "MA"}].

Prints JSON. Exit 0 always for --location-options / --resolve-location (problems are listed
in "warnings"); --show exits 1 when a required setting is missing or a value is invalid.
"""
import sys as _sys
_sys.dont_write_bytecode = True   # never write __pycache__ into the skill folder
import argparse, difflib, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from notary_block import STATES, MA_COUNTIES, state_name, check_county  # noqa: E402

DEFAULTS = {"signature_page": "same-page", "notary_state": "Massachusetts",
            "signing_locations": [], "attorney_in_fact": ""}
PAGE_CHOICES = ("same-page", "separate-page", "ask")
ABBR = {v: k for k, v in STATES.items()}
MAX_OPTIONS = 4                      # the question tool's limit


def unresolved(val):
    """Missing, or still a placeholder: '${user_config.x}', '{{x}}', '[x]', '<x>'."""
    if val is None:
        return True
    if isinstance(val, (list, dict)):
        return not val
    s = str(val).strip()
    return (not s or "${" in s or "user_config" in s or s.startswith("{{")
            or (s.startswith("[") and s.endswith("]")) or (s.startswith("<") and s.endswith(">")))


def load_json(path):
    if path and os.path.isfile(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def file_sources():
    out = []
    env = os.environ.get("SELLER_POA_SETTINGS")
    if env:
        out.append(("SELLER_POA_SETTINGS file", load_json(env)))
    out.append(("firm-settings.json", load_json(os.path.join(SKILL_DIR, "firm-settings.json"))))
    return out


def resolve(flags):
    """Merge flags -> files -> defaults. Returns (settings, sources)."""
    settings, sources = {}, {}
    files = file_sources()
    for key in DEFAULTS:
        val = flags.get(key)
        if not unresolved(val):
            settings[key], sources[key] = val, "plugin setting"
            continue
        for name, data in files:
            fkey = "default_notary_state" if key == "notary_state" and "default_notary_state" in data else key
            if not unresolved(data.get(fkey)):
                settings[key], sources[key] = data[fkey], name
                break
        else:
            settings[key], sources[key] = DEFAULTS[key], "default"
    return settings, sources


def parse_locations(raw, default_state):
    """'Label = County, State; ...' (or a list of strings / objects) -> (locations, warnings)."""
    items = raw
    if isinstance(raw, str):
        items = [x for x in re.split(r"[;\n]+", raw) if x.strip()]
    locs, warns = [], []
    for it in items or []:
        if isinstance(it, dict):
            label, county, state = it.get("label", ""), it.get("county", ""), it.get("state", "")
        else:
            if "=" not in it:
                warns.append(f"signing location '{it.strip()}' skipped -- use 'Label = County, State'")
                continue
            label, place = (x.strip() for x in it.split("=", 1))
            parts = [p.strip() for p in place.split(",") if p.strip()]
            if len(parts) >= 2:
                county, state = ", ".join(parts[:-1]), parts[-1]
            elif len(parts) == 1 and (parts[0].upper() in STATES or parts[0].title() in STATES.values()):
                county, state = "", parts[0]
            else:
                county, state = (parts[0] if parts else ""), default_state
        label = (label or "").strip()
        try:
            st = state_name(state or default_state)
            ct = check_county(st, county)
        except ValueError as e:
            warns.append(f"signing location '{label}' skipped -- {e}")
            continue
        if not label:
            label = f"{ct + ' County, ' if ct else ''}{st}"
        given = re.sub(r"\s+(county|parish)\s*$", "", (county or "").strip(), flags=re.I)
        if ct and given and given.lower() != ct.lower():
            warns.append(f"signing location '{label}': county '{county}' read as '{ct}'")
        locs.append({"label": label, "state": st, "county": ct})
    return locs, warns


def place_text(state, county):
    st = ABBR.get(state, state)
    return f"{county} County, {st}" if county else f"{state} \u2014 county left blank"


def show(flags):
    s, src = resolve(flags)
    errors, warns = [], []
    if unresolved(s["attorney_in_fact"]):
        errors.append("attorney_in_fact is not set. Every POA appoints this person, so there is no "
                      "default: set it in the plugin settings (/plugin configure ma-seller-poa@ma-seller-poa) "
                      "or in firm-settings.json next to SKILL.md. Ask the user for the name "
                      "(e.g. 'Alex Q. Attorney, Esq.') and do not build until it is set.")
    page = str(s["signature_page"]).strip().lower()
    if page not in PAGE_CHOICES:
        warns.append(f"signature_page '{s['signature_page']}' is not one of {', '.join(PAGE_CHOICES)}"
                     " -- using same-page")
        page = "same-page"
    try:
        state = state_name(s["notary_state"])
    except ValueError as e:
        warns.append(f"notary_state: {e} -- using Massachusetts")
        state = "Massachusetts"
    locs, w = parse_locations(s["signing_locations"], state)
    warns += w
    out = {"attorney_in_fact": s["attorney_in_fact"] if not unresolved(s["attorney_in_fact"]) else None,
           "signature_page": page, "mode": {"same-page": "same", "separate-page": "separate"}.get(page),
           "notary_state": state, "signing_locations": locs, "sources": src,
           "errors": errors, "warnings": warns}
    return out


def location_options(flags, property_county=None, property_state=None):
    st = show(flags)
    state, locs = st["notary_state"], st["signing_locations"]
    opts = []
    for i, loc in enumerate(locs[:MAX_OPTIONS - 1]):
        opts.append({"label": f"{loc['label']}{' (Recommended)' if i == 0 else ''}",
                     "description": f"Prints {place_text(loc['state'], loc['county'])} in the notary block",
                     "state": loc["state"], "county": loc["county"]})
    covered = {(o["state"], o["county"]) for o in opts}
    pstate = state_name(property_state) if property_state else "Massachusetts"
    if property_county:
        try:
            pc = check_county(pstate, property_county)
        except ValueError:
            pc = ""
        if pc and (pstate, pc) not in covered and len(opts) < MAX_OPTIONS - 1:
            opts.append({"label": f"{pc} County, {ABBR.get(pstate, pstate)} (where the property is)"
                                  + ("" if opts else " (Recommended)"),
                         "description": f"Prints {pc} County, {ABBR.get(pstate, pstate)}",
                         "state": pstate, "county": pc})
    opts.append({"label": f"{state} \u2014 leave the county blank",
                 "description": "Prints COUNTY OF ______________ for the notary to fill in",
                 "state": state, "county": ""})
    if len(opts) < 2:                      # the question tool needs two options or more
        opts.append({"label": "Somewhere else (type it under Other)",
                     "description": "A county and state, or a city or town", "state": None,
                     "county": None})
    note = None
    if len(locs) > MAX_OPTIONS - 1:
        more = ", ".join(l["label"] for l in locs[MAX_OPTIONS - 1:])
        note = (f"More firm locations than fit as options: {more}. Tell the user they can type one "
                "of these names under Other.")
    if not locs:
        note = ("No firm signing locations are set (setting signing_locations). The user can type "
                "a county and state, or a city or town, under Other.")
    return {"options": opts, "note": note, "warnings": st["warnings"], "errors": st["errors"]}


def resolve_location(text, flags):
    """A typed answer that names a firm location ('boston office') -> its state and county."""
    st = show(flags)
    locs = st["signing_locations"]
    t = re.sub(r"\s*\(recommended\)\s*$", "", (text or "").strip(), flags=re.I).lower()
    for loc in locs:
        if loc["label"].lower() == t:
            return {"match": loc}
    labels = [l["label"].lower() for l in locs]
    close = difflib.get_close_matches(t, labels, n=1, cutoff=0.6)
    hits = [l for l in locs if t and t in l["label"].lower()]
    if close:
        return {"match": locs[labels.index(close[0])], "note": f"read '{text}' as '{locs[labels.index(close[0])]['label']}'"}
    if len(hits) == 1:
        return {"match": hits[0], "note": f"read '{text}' as '{hits[0]['label']}'"}
    return {"match": None, "note": "not a firm signing location -- read it as a county and state, "
                                   "or a city or town (SKILL.md, 'Reading the notary location')"}


def main(argv=None):
    ap = argparse.ArgumentParser(description="Seller POA firm settings")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--show", action="store_true")
    g.add_argument("--location-options", action="store_true")
    g.add_argument("--resolve-location")
    ap.add_argument("--attorney-in-fact")
    ap.add_argument("--signing-locations")
    ap.add_argument("--signature-page")
    ap.add_argument("--notary-state")
    ap.add_argument("--property-county")
    ap.add_argument("--property-state")
    a = ap.parse_args(argv)
    flags = {"attorney_in_fact": a.attorney_in_fact, "signing_locations": a.signing_locations,
             "signature_page": a.signature_page, "notary_state": a.notary_state}
    if a.show:
        out = show(flags)
        print(json.dumps(out, indent=2, ensure_ascii=False))
        return 1 if out["errors"] else 0
    if a.location_options:
        print(json.dumps(location_options(flags, a.property_county, a.property_state), indent=2, ensure_ascii=False))
        return 0
    print(json.dumps(resolve_location(a.resolve_location, flags), indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
