#!/usr/bin/env python3
"""notary_block.py -- the firm's notary acknowledgment, shared by every Seller POA path.

Follows the firm's Notary Acknowledgment form (Massachusetts wording):

  COMMONWEALTH OF MASSACHUSETTS            (or "STATE OF [STATE]" -- see state_heading())
  COUNTY OF [COUNTY]                       (or "COUNTY OF ______________" when left blank)

  On this September 16th, 2026, before me, the undersigned notary public, [**Signers**] personally
  appeared, proved to me through satisfactory evidence of identification, which
  were_________________________, to be the person whose name is signed [persons whose names are
  signed] on the preceding or attached document, and acknowledged to me that [he/she/they] signed
  it voluntarily for its stated purpose[ capacity clause].

  ______________________________
  Notary Public Signature

  My Commission Expires: _____________, 20___

Notary state and county: the state is an input that defaults
to Massachusetts; the county may be left blank, which prints "COUNTY OF ______________".
Signing date -- what is known when the POA is drafted. The EXECUTED line and
the notary date print only the parts of the signing date that are known; the rest are blanks:
  "2026-11-14"  exact date      -> "November 14th, 2026"   (suffix in superscript)
  "2026-11"     month and year  -> "November _____, 2026"
  "2026"        year only       -> "______________ _____, 2026"
  ""            nothing known   -> "______________ _____, _______"
signing_date_options() works out the choices Step 3 offers (the month suggested from today's date
and the closing date); run "python notary_block.py --date-options --today YYYY-MM-DD
[--closing YYYY-MM-DD]" to print them as JSON.

Separate notary appointments: when the signers attend different
notary appointments, each appointment gets its own counterpart signature page -- its own
EXECUTED line and date, only its signers' signature lines, and its own notary block with that
appointment's state, county, date, and he/she/they. keep_together.py adds the page breaks and the
"Counterpart Signature Page k of n" captions.

Usage as a script (fills the "{{Notary Block}}" placeholder of a template-based .docx; kept for
manual use -- since Oct 27, 2026 build_individual_poa.py builds the individual-seller section itself):
  python notary_block.py POA.docx NOTARY_VALUES.json
  python notary_block.py --date-options --today 2026-10-28 [--closing 2026-11-20]
  python notary_block.py --pronoun-options --name "Joseph P. Example, Jr."

NOTARY_VALUES.json keys:
  signing_date: "YYYY-MM-DD" (exact), "YYYY-MM" (month and year), "YYYY" (year only), or
                "" / null (all blank) -- see "Signing date" above
  state:        "Massachusetts" (default when left out; a two-letter code also works)
  county:       "Plymouth" (the word "County" is optional; printed in capitals);
                "" / null / left out -> blank line
  signers:      ["Jane Doe", "John Doe"]          (names only, in signing order)
  pronoun:      "he" / "she" / "they" / "blank"   (required for ONE signer -- ask the user, never
                                                   guess; "blank" prints PRONOUN_BLANK "_____" for
                                                   the notary to fill in;
                                                   ignored for 2+ signers, which get "they")
  capacity:     {"type": "individual"}            (default -- no capacity clause)
              | {"type": "trust", "trust_name": "Doe Family Trust"}
              | {"type": "entity", "entity_name": "...", "entity_type": "Nonprofit Corporation",
                 "titles": ["President", "Treasurer"]}     (one title per signer, in order)
              | {"type": "estate", "capacity": "Personal Representative",
                 "estate_name": "Estate of Mary Smith", "preposition": "of"}
                 (capacity pluralized for 2+ signers; "the" before the estate name unless it
                  begins with "The"; preposition "of" by default. The older "estate_pr" and
                  "estate_other" types are accepted as aliases.)
              | {"type": "aif", "principal_name": "Robert Smith"}
  appointments: OPTIONAL -- only when the signers attend DIFFERENT notary appointments:
                [{"signers": ["Jane Doe"], "state": "Massachusetts", "county": "Plymouth",
                  "signing_date": "2026-10-30", "pronoun": "she"},
                 {"signers": ["John Doe"], "state": "New Hampshire", "county": "Rockingham",
                  "signing_date": "", "pronoun": "he"}]
                Every signer must be in exactly one appointment. With "appointments", the
                top-level signing_date / state / county / pronoun are not used.

Builders import notary_paragraph(), executed_runs(), appointments(), counterparts_paragraph()
and notary_notes() directly; the script form fills the "{{Notary Block}}" paragraph of an
existing .docx (with several appointments it also rebuilds the EXECUTED / signature section).
"""
import sys as _sys
_sys.dont_write_bytecode = True   # never write __pycache__ into the skill folder
import datetime, difflib, json, os, re, sys, tempfile, zipfile
from xml.sax.saxutils import escape

PLACEHOLDER = "{{Notary Block}}"
ID_BLANK = "_" * 25
SIG_LINE = "_" * 30
COUNTY_BLANK = "_" * 14
MONTH_BLANK = "_" * 14          # signing date blanks:
DAY_BLANK = "_" * 5             #   "______________ _____, _______"
YEAR_BLANK = "_" * 7
PRONOUN_BLANK = "_" * 5         # one signer, pronoun not known
# Signer's signature line, every seller type: an UNDERLINED TAB that runs to a
# tab stop SIG_LINE_TWIPS from the start of the line -- 3.0", the length of the one-seller line in
# Form 2 (indent 5040 -> right tab 9360). A tab fills exactly to its stop, so the line is the same
# length in Word and LibreOffice whatever the font, and it can never wrap onto a second line
# (a run of underscores can). "By:" lines end at the same 3.0" point. Every signature column is
# 3.25" (4675 twips), so the line always fits with room to spare.
SIG_LINE_TWIPS = 4320
SIG_TABS = f'<w:tabs><w:tab w:val="left" w:pos="{SIG_LINE_TWIPS}"/></w:tabs>'
COMMISSION = "My Commission Expires: _____________, 20___"
DEFAULT_STATE = "Massachusetts"
PPR = ('<w:pPr><w:pStyle w:val="Footer"/><w:keepNext/><w:keepLines/><w:tabs>'
       '<w:tab w:val="clear" w:pos="4320"/><w:tab w:val="clear" w:pos="8640"/>'
       '<w:tab w:val="left" w:pos="5040"/><w:tab w:val="right" w:pos="10080"/></w:tabs>'
       '<w:jc w:val="left"/></w:pPr>')

COUNTERPARTS = ("THIS Power of Attorney may be executed in counterparts, each of which shall be "
                "deemed an original, and all of which together shall constitute one and the same "
                "instrument.")

# States whose official name is "Commonwealth of ..." -- PROVISIONAL list (backlog B6: confirm the
# full list of which heading each state's notaries use). Every other state gets "STATE OF".
COMMONWEALTHS = {"Massachusetts", "Pennsylvania", "Virginia", "Kentucky"}

STATES = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas", "CA": "California",
    "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware", "DC": "District of Columbia",
    "FL": "Florida", "GA": "Georgia", "HI": "Hawaii", "ID": "Idaho", "IL": "Illinois",
    "IN": "Indiana", "IA": "Iowa", "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana",
    "ME": "Maine", "MD": "Maryland", "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota",
    "MS": "Mississippi", "MO": "Missouri", "MT": "Montana", "NE": "Nebraska", "NV": "Nevada",
    "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico", "NY": "New York",
    "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio", "OK": "Oklahoma", "OR": "Oregon",
    "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina", "SD": "South Dakota",
    "TN": "Tennessee", "TX": "Texas", "UT": "Utah", "VT": "Vermont", "VA": "Virginia",
    "WA": "Washington", "WV": "West Virginia", "WI": "Wisconsin", "WY": "Wyoming",
}

MA_COUNTIES = ["Barnstable", "Berkshire", "Bristol", "Dukes", "Essex", "Franklin", "Hampden",
               "Hampshire", "Middlesex", "Nantucket", "Norfolk", "Plymouth", "Suffolk", "Worcester"]


# ---------- small helpers ----------
def _r(text, bold=False):
    rpr = "<w:b/><w:bCs/>" if bold else ""
    return f'<w:r><w:rPr>{rpr}<w:noProof/></w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'

BR = "<w:r><w:rPr><w:noProof/></w:rPr><w:br/></w:r>"
TAB = "<w:r><w:tab/></w:r>"
BODY_PPR = '<w:pPr><w:pStyle w:val="BodyText"/></w:pPr>'


def ordinal(n):
    return "th" if 11 <= n % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")


def blank_date(iso):
    """True when nothing at all is known about the signing date."""
    return not iso or str(iso).strip().lower() in ("blank", "none", "null", "tbd")


def date_parts(iso):
    """Signing date -> (year, month, day), None for each part not known:
    'YYYY-MM-DD' exact | 'YYYY-MM' month and year | 'YYYY' year only | '' nothing known."""
    if blank_date(iso):
        return (None, None, None)
    s = str(iso).strip()
    m = re.fullmatch(r"(\d{4})(?:-(\d{1,2})(?:-(\d{1,2}))?)?", s)
    if not m:
        raise ValueError(f"signing date '{s}' -- use YYYY-MM-DD, YYYY-MM, YYYY, or \"\" (blank)")
    y, mo, d = (int(x) if x else None for x in m.groups())
    if d is not None:
        datetime.date(y, mo, d)                         # raises on an impossible date
    elif mo is not None and not 1 <= mo <= 12:
        raise ValueError(f"signing date '{s}' -- month must be 01-12")
    return (y, mo, d)


def date_text(iso):
    """Plain text of a signing date: 'November 14th, 2026', 'November _____, 2026',
    '______________ _____, 2026', or '______________ _____, _______'."""
    y, mo, d = date_parts(iso)
    month = datetime.date(2000, mo, 1).strftime("%B") if mo else MONTH_BLANK
    day = f"{d}{ordinal(d)}" if d else DAY_BLANK
    return f"{month} {day}, {y if y else YEAR_BLANK}"


def _date_runs(iso):
    """The date in the POA's date format -- 'September 16th, 2026' with the suffix in superscript;
    a partly known signing date prints blanks for the unknown parts (see date_text)."""
    y, mo, d = date_parts(iso)
    if d is None:
        return _r(date_text(iso))
    sup = (f'<w:r><w:rPr><w:noProof/><w:vertAlign w:val="superscript"/></w:rPr>'
           f'<w:t>{ordinal(d)}</w:t></w:r>')
    return _r(f"{datetime.date(y, mo, d):%B} {d}") + sup + _r(f", {y}")


# ---------- the Step 3 signing-date question ----------
NEXT_MONTH_DAYS = 7          # in the last 7 days of a month, suggest the next month
DECEMBER_FROM = 8            # from December 8th on, offer next year's blanks too
BLANK_YEAR_AFTER = (12, 15)  # all-blank year only when closing is after December 15th


def _add_month(y, m, k=1):
    m += k
    return y + (m - 1) // 12, (m - 1) % 12 + 1


def _sup_free(d):
    return f"{d:%B} {d.day}{ordinal(d.day)}"


def signing_date_options(today, closing=None):
    """The options for the signing-date question when the signing is not booked yet. TODAY and
    CLOSING are datetime.date (CLOSING may be None). Returns a list of dicts
    {label, value, description} in display order; the first is recommended. A booked signing
    date is typed by the user under "Other" (no option for it).

    Rules:
      - Suggested month: the current month, or the next month during the last 7 days of a month
        -- but never later than the closing month (the POA is signed at or before closing).
      - Closing month: also offered when it is later than the suggested month.
      - Year: the current year, without asking -- except from December 8th through December 31st,
        when next year's options are offered too, so the user picks the year (not when the
        closing date is on or before December 15th: the POA is then signed this year).
      - "______________ _____, _______" (year blank): only when the closing date is after
        December 15th of the current year (an extension into the next year is then possible).
      - At most four options (the question tool's limit)."""
    import calendar
    if closing is not None and closing < today:
        closing = None                                  # stale -- caller reports it
    dim = calendar.monthrange(today.year, today.month)[1]
    rec = (today.year, today.month)
    if today.day > dim - NEXT_MONTH_DAYS:
        rec = _add_month(*rec)
    cm = (closing.year, closing.month) if closing else None
    if cm and cm < rec:
        rec = cm
    month = lambda ym: datetime.date(ym[0], ym[1], 1).strftime("%B")
    opts = [{"label": f"{month(rec)} {DAY_BLANK}, {rec[0]} (Recommended)",
             "value": f"{rec[0]:04d}-{rec[1]:02d}",
             "description": "Month and year filled in; the day is written in at signing"
                            + (f" (closing {_sup_free(closing)})" if cm == rec else "")}]
    if cm and cm > rec:
        opts.append({"label": f"{month(cm)} {DAY_BLANK}, {cm[0]}",
                     "value": f"{cm[0]:04d}-{cm[1]:02d}",
                     "description": f"The closing month (closing {_sup_free(closing)})"})
    late = datetime.date(today.year, *BLANK_YEAR_AFTER)
    blank_year = bool(closing and closing > late)
    # December 8th-31st: next year's options too -- unless a closing date on or before
    # December 15th shows the POA will be signed this year.
    december = today.month == 12 and today.day >= DECEMBER_FROM and (closing is None or blank_year)
    last = closing.year if closing else today.year
    years = sorted({today.year, rec[0], last} | ({today.year + 1} if december else set()))
    year_opts = [{"label": f"{MONTH_BLANK} {DAY_BLANK}, {y}", "value": f"{y:04d}",
                  "description": "Only the year filled in"} for y in years]
    room = 4 - len(opts) - (1 if blank_year else 0)
    opts += year_opts[:max(room, 1)]
    if blank_year and len(opts) < 4:
        opts.append({"label": f"{MONTH_BLANK} {DAY_BLANK}, {YEAR_BLANK}", "value": "",
                     "description": f"All blank -- closing {_sup_free(closing)}, {closing.year} "
                                    "is after December 15th"})
    return opts


# ---------- the Step 3 one-signer pronoun question: silent ordering ----------
# The options are always He / She / They / Leave blank (_____), none recommended. Only their ORDER
# changes: the more likely of "He" and "She" for the signer's name goes first, so the user can pick
# it quickly. The ranking is silent -- it never preselects, recommends, or answers the question.
MALE_NAMES = set("""
aaron adam alan albert alexander alfred andrew angelo anthony antonio arnold arthur barry ben
benjamin bernard bill billy bob bobby brad bradley brandon brent brett brian bruce bryan carl
carlos cecil charles chad chester christopher clarence clifford clyde craig curtis dale dan daniel
darren dave david dean dennis derek dick don donald douglas duane dustin earl eddie edgar edward
edwin eric ernest eugene evan felix floyd francis frank franklin fred frederick gary gene geoffrey
george gerald gilbert glen glenn gordon greg gregory harold harry harvey hector henry herbert
howard hugh ian isaac jack jacob jake james jason jeff jeffery jeffrey jeremy jerome jerry jesse
jim jimmy joe joel john johnny jon jonathan jose joseph joshua juan justin keith ken kenneth kevin
kurt kyle lance larry lawrence leo leon leonard lewis lloyd louis luke manuel marc marcus mario mark
marshall martin matthew maurice max michael mike miguel mitchell nathan nathaniel neil
nicholas norman oliver oscar owen patrick paul pedro peter philip phillip ralph randall randy ray
raymond richard rick ricky rob robert roberto rodney roger roland ronald ross roy russell ryan sam
samuel scott sean seth shawn stanley stephen steve steven stuart ted terrence theodore thomas tim
timothy todd tom tommy tony travis troy victor vincent walter warren wayne wesley william willie
zachary
""".split())
FEMALE_NAMES = set("""
abigail ada agnes alice alicia allison amanda amber amy andrea angela anita ann anna anne annette
annie april arlene audrey barbara beatrice becky beth betty beverly bonnie brenda bridget brittany
carla carmen carol caroline carolyn catherine cathy charlotte cheryl christina christine cindy
claire clara colleen connie crystal cynthia daisy darlene dawn debbie deborah debra denise diana
diane dolores donna doris dorothy edith eileen elaine eleanor elizabeth ellen emily emma erica
erin esther ethel eva evelyn faith florence frances gail gina gladys gloria grace gwendolyn hannah
harriet hazel heather heidi helen holly irene jacqueline jane janet janice jean jeanne jennifer
jessica jill joan joann joanne josephine joy joyce judith judy julia julie june karen katherine
kathleen kathryn kathy katie kim kimberly laura lauren linda lisa lois loretta lori lorraine louise
lucy lynn margaret maria marie marilyn marjorie martha mary maureen megan melanie melissa michelle
mildred monica nancy natalie nicole nora norma olivia pamela patricia paula pauline peggy phyllis
rachel rebecca regina rita roberta rose rosemary ruth sally samantha sandra sara sarah sharon
sheila shirley sophia stacy stephanie susan suzanne sylvia tammy teresa theresa tiffany tina
valerie vanessa veronica victoria virginia vivian wanda wendy yvonne
""".split())
MALE_TITLES = {"mr"}
FEMALE_TITLES = {"mrs", "ms", "miss", "madam"}
SKIP_TITLES = {"dr", "rev", "hon", "atty", "prof", "sir"}
MALE_SUFFIXES = {"jr", "sr"}                     # "Jr." / "Sr." are used for men


def pronoun_lean(name):
    """'he', 'she', or None (can't tell) for a signer's name -- used ONLY to order the options.
    Reads a title (Mr. / Mrs. / Ms.), a Jr. / Sr. suffix, then the first given name that is not
    an initial, against the built-in name lists, then a few spelling endings."""
    words = [w.strip(".,").lower() for w in (name or "").replace(",", " ").split()]
    words = [w for w in words if w]
    if not words:
        return None
    if words[0] in MALE_TITLES:
        return "he"
    if words[0] in FEMALE_TITLES:
        return "she"
    if any(w in MALE_SUFFIXES for w in words[1:]):
        return "he"
    given = [w for w in words[:-1] or words if w not in SKIP_TITLES and len(w) > 1]
    if not given:
        return None
    first = given[0]
    if first in MALE_NAMES:
        return "he"
    if first in FEMALE_NAMES:
        return "she"
    if first.endswith(("a", "ette", "elle", "ine", "lyn", "lynn", "leen", "issa")):
        return "she"
    if first.endswith(("o", "us", "ard", "ert", "old")):
        return "he"
    return None


def pronoun_options(name, lean=None):
    """The one-signer pronoun options in display order. LEAN ('he' / 'she') overrides the
    built-in reading -- e.g. Claude's own reading of a name the lists don't cover. No option is
    marked Recommended, and no description hints at the ranking."""
    lean = lean or pronoun_lean(name)
    first, second = (("She", "she"), ("He", "he")) if lean == "she" else (("He", "he"), ("She", "she"))
    return [{"label": first[0], "value": first[1]},
            {"label": second[0], "value": second[1]},
            {"label": "They", "value": "they"},
            {"label": f"Leave blank ({PRONOUN_BLANK})", "value": "blank",
             "description": "The notary fills in the pronoun"}]


def pronoun_options_main(argv):
    """python notary_block.py --pronoun-options --name "Joseph P. Example, Jr." """
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--pronoun-options", action="store_true")
    ap.add_argument("--name", required=True)
    a = ap.parse_args(argv)
    lean = pronoun_lean(a.name)
    print(json.dumps({"options": pronoun_options(a.name), "lean": lean or "unknown"}, indent=2))


def sig_line_runs():
    """The signer's signature line: one underlined tab, to the SIG_TABS stop (put SIG_TABS in the
    paragraph's pPr). Follows "By:" in entity, trust, and estate signature lines."""
    return '<w:r><w:rPr><w:noProof/><w:u w:val="single"/></w:rPr><w:tab/></w:r>'


def executed_runs(iso):
    """'EXECUTED as a sealed instrument this September 16th, 2026.' (flush left -- no tab)."""
    return _r("EXECUTED as a sealed instrument this ") + _date_runs(iso) + _r(".")


def include_ps_date(v):
    """False when the user chose to leave the P&S date out of the POA."""
    x = v.get("include_ps_date", True)
    if isinstance(x, str):
        x = x.strip().lower() not in ("false", "no", "0", "")
    return bool(x)


def ps_date_text(v):
    """The P&S date as it reads in the POA, for messages: 'September 6th, 2023',
    'September _____, 2023', '______________ _____, _______', or 'left out'."""
    return date_text(v.get("ps_date")) if include_ps_date(v) else "left out"


def agreement_runs(v):
    """' (the "Premises"), as set forth in a Purchase and Sale Agreement dated [P&S date], with '
    -- shared by every builder. The P&S date ("ps_date") is printed in
    the POA's date format; like the signing date it may be partly or wholly blank ("YYYY-MM",
    "YYYY", or "" -> "September _____, 2023", "______________ _____, 2023",
    "______________ _____, _______"). With "include_ps_date": false, the words "dated [P&S date]"
    are left out: "...a Purchase and Sale Agreement with [Buyers] (the "P&S")..."."""
    out = _r(" (the “Premises”), as set forth in a Purchase and Sale Agreement")
    if include_ps_date(v):                       # "...Agreement dated September 6th, 2023, with "
        return out + _r(" dated ") + _date_runs(v.get("ps_date")) + _r(", with ")
    return out + _r(" with ")                    # "...Agreement with " -- no comma


def attorney_in_fact(v):
    """The attorney in fact named in every POA -- the firm setting `attorney_in_fact`, passed in
    poa_values.json. Required: there is no default, so a POA can never appoint someone the user
    did not choose. Refuses an unresolved settings placeholder (e.g. a literal
    "${user_config.attorney_in_fact}") or a template blank."""
    name = (v.get("attorney_in_fact") or "").strip()
    if not name:
        raise ValueError("attorney_in_fact is required (the firm setting -- the attorney the POA "
                         "appoints, e.g. \"Alex Q. Attorney, Esq.\"); run firm_settings.py --show")
    if "${" in name or "{{" in name or "user_config" in name or name.startswith("["):
        raise ValueError(f"attorney_in_fact is an unresolved placeholder ({name!r}) -- set the "
                         "firm setting; see the README, \"Settings\"")
    return name


def check_ps_date(v):
    """Raise a clear error for a malformed P&S date (an omitted date needs no value)."""
    if include_ps_date(v):
        assert "ps_date" in v, ("ps_date is required -- \"YYYY-MM-DD\", \"YYYY-MM\", \"YYYY\", or "
                                "\"\" (blank); or \"include_ps_date\": false to leave it out")
        try:
            date_parts(v.get("ps_date"))
        except ValueError as e:
            raise ValueError(f"'{v.get('ps_date')}' is not a valid P&S date ({e})") from None


def join_list(items):
    """A | A and B | A, B, and C"""
    if len(items) <= 2:
        return " and ".join(items)
    return ", ".join(items[:-1]) + ", and " + items[-1]


def plural(title):
    if title.endswith("y") and title[-2:-1] not in "aeiou":
        return title[:-1] + "ies"
    return title + "s"


CAPACITY_HEADS = {"Representative": "Representatives", "Administrator": "Administrators",
                  "Administratrix": "Administratrices", "Executor": "Executors",
                  "Executrix": "Executrices"}


def plural_capacity(cap):
    """Pluralize a probate capacity at its head noun: 'Personal Representative' -> 'Personal
    Representatives'; 'Administrator with the Will Annexed' -> 'Administrators with the Will
    Annexed'; 'Public Administrator' -> 'Public Administrators'."""
    words = cap.split()
    for i in range(len(words) - 1, -1, -1):
        w = words[i].rstrip(",")
        if w in CAPACITY_HEADS:
            words[i] = words[i].replace(w, CAPACITY_HEADS[w])
            return " ".join(words)
    return plural(cap)


def singular_capacity(cap):
    """Undo plural_capacity(): 'Personal Representatives' -> 'Personal Representative'."""
    back = {v: k for k, v in CAPACITY_HEADS.items()}
    words = cap.split()
    for i in range(len(words) - 1, -1, -1):
        w = words[i].rstrip(",")
        if w in back:
            words[i] = words[i].replace(w, back[w])
            return " ".join(words)
    return cap


def estate_capacity(cap, n):
    return cap if n == 1 else plural_capacity(cap)


def the_trust(name):
    """Firm-wide rule: 'the' before a trust name unless the name already begins with 'The'."""
    return name if name.split()[0].lower() == "the" else "the " + name


def article(phrase):
    """'a' / 'an' by the first sound: an LLC, an S corporation, an entity; a Nonprofit Corporation."""
    w = phrase.split()[0]
    if w.isupper() and len(w) > 1:                       # acronym, read letter by letter
        return "an" if w[0] in "AEFHILMNORSX" else "a"
    return "an" if w[0].lower() in "aeiou" else "a"


# ---------- state and county (backlog B4) ----------
def state_name(s):
    """'MA' / 'mass' / 'massachusetts' -> 'Massachusetts'; empty -> the default (Massachusetts)."""
    s = (s or "").strip().rstrip(".")
    if not s:
        return DEFAULT_STATE
    if s.upper() in STATES:
        return STATES[s.upper()]
    s = re.sub(r"^(the\s+)?(commonwealth|state)\s+of\s+", "", s, flags=re.I).strip()
    for full in STATES.values():
        if full.lower() == s.lower():
            return full
    if s.lower() in ("mass", "mass."):
        return "Massachusetts"
    close = difflib.get_close_matches(s.title(), STATES.values(), n=1, cutoff=0.8)
    raise ValueError(f"unknown state '{s}'" + (f" -- did you mean {close[0]}?" if close else ""))


def state_heading(state):
    """First header line: 'COMMONWEALTH OF MASSACHUSETTS', 'STATE OF NEW HAMPSHIRE',
    'DISTRICT OF COLUMBIA'."""
    st = state_name(state)
    if st == "District of Columbia":
        return "DISTRICT OF COLUMBIA"
    return ("COMMONWEALTH OF " if st in COMMONWEALTHS else "STATE OF ") + st.upper()


def county_only(c):
    return re.sub(r"\s+(county|parish)\s*$", "", (c or "").strip(), flags=re.I)


def check_county(state, county):
    """A Massachusetts county must be one of the 14 (a close misspelling is corrected).
    Returns the county as it should print (title case), or '' for a blank county."""
    c = county_only(county)
    if not c or set(c) <= {"_"} or c.lower() in ("blank", "none", "leave blank"):
        return ""
    if state_name(state) == "Massachusetts":
        for m in MA_COUNTIES:
            if m.lower() == c.lower():
                return m
        close = difflib.get_close_matches(c.title(), MA_COUNTIES, n=1, cutoff=0.75)
        if close:
            return close[0]
        raise ValueError(f"'{county}' is not a Massachusetts county (the 14 are: "
                         + ", ".join(MA_COUNTIES) + ") -- ask the user")
    return " ".join(w.capitalize() if w.islower() else w for w in c.split())


def county_line(state, county):
    """'COUNTY OF PLYMOUTH', 'COUNTY OF ______________', 'PARISH OF ...' (Louisiana);
    None for the District of Columbia (no county line)."""
    st = state_name(state)
    if st == "District of Columbia":
        return None
    word = "PARISH OF " if st == "Louisiana" else "COUNTY OF "
    c = check_county(st, county)
    return word + (c.upper() if c else COUNTY_BLANK)


def county_name(c):                       # kept for older callers
    return county_only(c).upper()


# ---------- the capacity clause (no trailing period -- the paragraph supplies it) ----------
def capacity_clause(v):
    cap = v.get("capacity") or {"type": "individual"}
    t, n = cap.get("type", "individual"), len(v["signers"])
    if t == "individual":
        return ""
    if t == "trust":
        return f' as {"Trustee" if n == 1 else "Trustees"} of {the_trust(cap["trust_name"])}'
    if t == "entity":
        titles = cap["titles"]
        assert len(titles) == n, "entity capacity needs one title per signer"
        if n == 1:
            office = titles[0]                              # "President" or "President and Treasurer"
        elif len(set(titles)) == 1:
            office = plural(titles[0])                      # "Managers"
        else:
            office = join_list(titles) + ", respectively,"  # "President and Treasurer, respectively,"
        et = cap["entity_type"]
        return (f" as {office} for {cap['entity_name']}, {article(et)} {et}, "
                f"as the voluntary act of the {et}")
    if t in ("estate", "estate_pr", "estate_other"):
        c = cap.get("capacity") or "Personal Representative"
        prep = cap.get("preposition") or "of"
        return f" as {estate_capacity(c, n)} {prep} {the_trust(cap['estate_name'])}"
    if t == "aif":   # spec wording, verbatim (open follow-up: "of the {{Principal Name}}")
        p = cap["principal_name"]
        return f" as attorney in fact for {p}, the principal, as the voluntary act of the {p}"
    raise ValueError(f"unknown capacity type: {t}")


def pronoun(v):
    if len(v["signers"]) > 1:
        return "they"
    p = (v.get("pronoun") or "").strip().lower()
    p = {"his": "he", "him": "he", "her": "she", "them": "they", "their": "they"}.get(p, p)
    if p in ("blank", "leave blank", "unknown") or (p and set(p) <= {"_"}):
        return PRONOUN_BLANK                  # the notary fills it in
    assert p in ("he", "she", "they"), \
        (f"pronoun (he / she / they / blank) is required for one signer ({v['signers'][0]}) "
         "-- ask the user")
    return p


def notary_paragraph(v):
    """Return the whole notary block as one <w:p> (lines separated by breaks)."""
    s = v["signers"]
    assert s, "at least one signer is required"
    names = ""
    for i, name in enumerate(s):                       # bold names, plain joining words
        if i:
            names += _r(" and " if len(s) == 2 else (", and " if i == len(s) - 1 else ", "))
        names += _r(name, bold=True)
    who = "person whose name is signed" if len(s) == 1 else "persons whose names are signed"
    body = (_r("On this ") + _date_runs(v.get("signing_date"))
            + _r(", before me, the undersigned notary public, ") + names
            + _r(f" personally appeared, proved to me through satisfactory evidence of identification, "
                 f"which were{ID_BLANK}, to be the {who} on the preceding or attached document, and "
                 f"acknowledged to me that {pronoun(v)} signed it voluntarily for its stated purpose"
                 f"{capacity_clause(v)}."))
    head = _r(state_heading(v.get("state")))
    cl = county_line(v.get("state"), v.get("county"))
    if cl:
        head += BR + _r(cl)
    return ("<w:p>" + PPR + head
            + BR + BR + body + BR + BR + _r(SIG_LINE) + BR + _r("Notary Public Signature")
            + BR + BR + _r(COMMISSION) + "</w:p>")


# ---------- notary appointments (backlog B3) ----------
def appointments(v, signer_names, state_key="state"):
    """The notary appointments, in signing order: a list of dicts with signers (names, in the
    POA's order), state, county, signing_date, pronoun. Without "appointments" in V there is one
    appointment with every signer and the top-level values. STATE_KEY names the top-level key
    holding the notary state ("notary_state" for entity sellers, whose "state" is the state of
    organization); each appointment's own key is always "state"."""
    names = list(signer_names)
    top_state = v.get(state_key)
    raw = v.get("appointments")
    if not raw or len(raw) == 1:
        a = dict(raw[0]) if raw else {}
        out = {"signers": names,
               "state": a.get("state", top_state),
               "county": a.get("county", v.get("county")),
               "signing_date": a.get("signing_date", v.get("signing_date")),
               "pronoun": a.get("pronoun", v.get("pronoun") or v.get("signer_pronoun"))}
        date_parts(out["signing_date"])                           # clear error on a bad date
        return [out]
    assert len(names) >= 2, "separate notary appointments need two or more signers"
    seen = []
    appts = []
    for a in raw:
        a = dict(a)
        sub = list(a.get("signers") or [])
        assert sub, "every notary appointment needs at least one signer"
        for s in sub:
            assert s in names, f"appointment signer '{s}' is not one of the signers: {names}"
            assert s not in seen, f"'{s}' is in more than one notary appointment"
            seen.append(s)
        a["signers"] = [s for s in names if s in sub]           # keep the POA's order
        a.setdefault("state", top_state)
        a.setdefault("county", v.get("county"))
        a.setdefault("signing_date", v.get("signing_date"))
        appts.append(a)
    missing = [s for s in names if s not in seen]
    assert not missing, f"not in any notary appointment: {', '.join(missing)}"
    appts.sort(key=lambda a: names.index(a["signers"][0]))       # Signer 1's appointment first
    for a in appts:                                               # validate early, clear errors
        state_name(a.get("state")); check_county(a.get("state"), a.get("county"))
        date_parts(a.get("signing_date"))
        if len(a["signers"]) == 1:
            pronoun(a)
    return appts


def counterparts_paragraph():
    """Body paragraph added only when there are several notary appointments."""
    return f"<w:p>{BODY_PPR}{TAB}{_r(COUNTERPARTS)}</w:p>"


def appt_label(a):
    """'Plymouth County, Massachusetts, October 30th, 2026' -- for messages."""
    st = state_name(a.get("state"))
    c = check_county(st, a.get("county"))
    where = (f"{c} County, " if c else "county blank, ") + st
    when = "date blank" if blank_date(a.get("signing_date")) else date_text(a.get("signing_date"))
    return f"{where}, {when}"


def notary_notes(appts):
    """Messages for the user about the notary blocks (printed by every path)."""
    out = []
    many = len(appts) > 1
    if many:
        out.append(f"APPOINTMENTS: {len(appts)} notary appointments -> {len(appts)} counterpart "
                   "signature pages; counterparts paragraph added to the POA text")
    for i, a in enumerate(appts, 1):
        tag = (f"appointment {i} ({join_list(a['signers'])})" if many else "notary block")
        st = state_name(a.get("state"))
        if many:
            out.append(f"APPOINTMENT {i}: {join_list(a['signers'])} -- {appt_label(a)}")
        if st != "Massachusetts":
            out.append(f"NOTARY: {tag} is in {st} -- the firm's Massachusetts acknowledgment "
                       f"wording is used under '{state_heading(st)}'; confirm a {st} notary will "
                       "accept it")
        if st == "California":
            out.append(f"NOTARY: WARNING -- {tag}: California notaries must use California's own "
                       "all-purpose acknowledgment form (backlog B5, not built yet); arrange a "
                       "California acknowledgment with the notary")
        if st in ("Alaska", "Louisiana", "District of Columbia"):
            out.append(f"NOTARY: {tag} -- {st} does not use ordinary counties; check the "
                       f"'{(county_line(st, a.get('county')) or state_heading(st))}' line")
        typed = county_only(a.get("county"))
        fixed = check_county(st, a.get("county"))
        if typed and fixed and typed.lower() != fixed.lower():
            out.append(f"NOTARY: {tag} -- county '{typed}' read as {fixed} County; confirm")
        if not fixed and st != "District of Columbia":
            out.append(f"NOTARY: {tag} -- county left blank (the notary fills it in)")
        if len(a["signers"]) == 1 and pronoun(a) == PRONOUN_BLANK:
            out.append(f"NOTARY: {tag} -- pronoun left blank ('{PRONOUN_BLANK}'; the notary fills it in)")
        y, mo, d = date_parts(a.get("signing_date"))
        known = ("exact date" if d else "month and year; day blank" if mo else
                 "year only; month and day blank" if y else "all blank")
        out.append(f"SIGNING DATE: {tag} -- {date_text(a.get('signing_date'))} ({known}; "
                   "EXECUTED line and notary date)")
    return out


# ---------- individual sellers: signature lines for one appointment ----------
def individual_signatures(names):
    """Signature lines in the templates' style: one signer -> Form 2's right-hand line and name;
    two or more -> Form 4's borderless two-column table, filled left then right, two signers per
    row (any number of signers: a third or fifth signer starts a new row on the left)."""
    if len(names) == 1:
        ppr = ('<w:pPr><w:keepNext/><w:keepLines/><w:tabs><w:tab w:val="right" w:pos="9360"/>'
               '</w:tabs><w:ind w:left="5040"/></w:pPr>')
        return (f'<w:p>{ppr}<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:tab/></w:r></w:p>'
                f'<w:p>{ppr}{_r(names[0])}</w:p>')
    ppr = f'<w:pPr><w:keepNext/><w:keepLines/>{SIG_TABS}<w:jc w:val="left"/></w:pPr>'
    def tc(name):
        inner = (f'<w:p>{ppr}{sig_line_runs()}</w:p><w:p>{ppr}{_r(name)}</w:p><w:p>{ppr}</w:p>'
                 if name else f'<w:p>{ppr}</w:p>')
        return f'<w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>{inner}</w:tc>'
    rows = "".join(f'<w:tr><w:trPr><w:cantSplit/></w:trPr>{tc(p[0])}{tc(p[1] if len(p) > 1 else None)}</w:tr>'
                   for p in (names[i:i + 2] for i in range(0, len(names), 2)))
    none = '<w:{0} w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    borders = "".join(none.format(b) for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
    return ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>'
            f'<w:tblBorders>{borders}</w:tblBorders>'
            '<w:tblLook w:val="0600" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" '
            'w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="4675"/>'
            '<w:gridCol w:w="4675"/></w:tblGrid>' + rows + '</w:tbl>')


def individual_counterparts(v, appts):
    """EXECUTED line + signature lines + notary block for each appointment (individual sellers)."""
    blank = f"<w:p>{BODY_PPR}</w:p>"
    out = ""
    for i, a in enumerate(appts):
        if i:
            out += blank
        nv = dict(a, capacity=v.get("capacity") or {"type": "individual"})
        out += (f"<w:p>{BODY_PPR}{executed_runs(a.get('signing_date'))}</w:p>" + blank
                + individual_signatures(a["signers"]) + f'<w:p><w:pPr><w:jc w:val="left"/></w:pPr></w:p>'
                + notary_paragraph(nv))
    return out


def para_spans(xml):
    return list(re.finditer(r"<w:p(?: [^>]*)?/>|<w:p(?: [^>]*[^/>])?>.*?</w:p>", xml, re.S))  # self-closing first


def fill(xml, v):
    """Replace the paragraph containing the placeholder with the notary block. With several
    appointments, replace the whole section (EXECUTED line through the placeholder) with one
    counterpart section per appointment, and add the counterparts paragraph to the POA text."""
    paras = para_spans(xml)
    hits = [m for m in paras if PLACEHOLDER in m.group(0)]
    assert len(hits) == 1, f"expected one {PLACEHOLDER} paragraph, found {len(hits)}"
    m = hits[0]
    appts = appointments(v, v["signers"])
    if len(appts) == 1:
        a = appts[0]
        return xml[:m.start()] + notary_paragraph(dict(v, **a)) + xml[m.end():]
    ex = [p for p in paras if "EXECUTED as a sealed instrument" in re.sub(r"<[^>]+>", "", p.group(0))]
    assert len(ex) == 1, "expected one EXECUTED paragraph in the template"
    s = ex[0].start()
    blank = f"<w:p>{BODY_PPR}</w:p>"
    return (xml[:s] + counterparts_paragraph() + blank
            + individual_counterparts(v, appts) + xml[m.end():])


def main(docx, values_path):
    v = json.load(open(values_path, encoding="utf-8"))
    work = tempfile.mkdtemp()
    with zipfile.ZipFile(docx) as z:
        z.extractall(work)
    p = os.path.join(work, "word", "document.xml")
    xml = fill(open(p, encoding="utf-8").read(), v)
    open(p, "w", encoding="utf-8").write(xml)
    tmp = docx + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(work):
            for f in files:
                full = os.path.join(root, f)
                z.write(full, os.path.relpath(full, work))
    os.replace(tmp, docx)
    appts = appointments(v, v["signers"])
    n = len(v["signers"])
    pr = ", ".join(pronoun(a) for a in appts)
    print(f"NOTARY: filled ({n} signer{'s' if n > 1 else ''}, {len(appts)} appointment"
          f"{'s' if len(appts) > 1 else ''}, pronoun{'s' if len(appts) > 1 else ''} '{pr}', "
          f"capacity '{(v.get('capacity') or {}).get('type', 'individual')}')")
    for line in notary_notes(appts):
        print(line)


def date_options_main(argv):
    """python notary_block.py --date-options --today YYYY-MM-DD [--closing YYYY-MM-DD]"""
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--date-options", action="store_true")
    ap.add_argument("--today", required=True)
    ap.add_argument("--closing")
    a = ap.parse_args(argv)
    today = datetime.date.fromisoformat(a.today)
    closing = datetime.date.fromisoformat(a.closing) if a.closing else None
    note = None
    if closing and closing < today:
        note = (f"closing date {closing:%B} {closing.day}, {closing.year} has passed -- not used for "
                "the suggestion; check for a later extension")
    print(json.dumps({"options": signing_date_options(today, closing), "note": note}, indent=2))


if __name__ == "__main__":
    if "--date-options" in sys.argv:
        date_options_main(sys.argv[1:])
    elif "--pronoun-options" in sys.argv:
        pronoun_options_main(sys.argv[1:])
    else:
        main(sys.argv[1], sys.argv[2])
