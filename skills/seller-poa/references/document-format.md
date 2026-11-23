# Document format — what the builders produce

Reference for `SKILL.md`. Everything here is done by the scripts; read it to explain a script
message or to check a finished POA. `[Attorney in fact]` is the firm setting `attorney_in_fact`.

## Formatting rules (all seller types)

**Date format (all seller types).** In the POA, the P&S date (unless the user left it blank or
out — Step 3, "P&S date"), **the EXECUTED line date, and the
notary date**
always read `[Month] [Day][suffix], [Year]` — e.g., "September 1st, 2026" and "EXECUTED as a
sealed instrument this September 16th, 2026." — with the ordinal suffix (st/nd/rd/th) as
superscript. Suffixes: 1st/21st/31st → "st" | 2nd/22nd → "nd" | 3rd/23rd → "rd" | 11th–13th and
all others → "th". In XML the date is three runs:

```xml
<w:r><w:t xml:space="preserve">September 1</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>st</w:t></w:r><w:r><w:t xml:space="preserve">, 2026</w:t></w:r>
```

The notary acknowledgment uses the same format: "On this September 16th, 2026, before me…"

**The signing date may be partly blank.** When the POA is drafted, the
Seller's signing (usually the same meeting at which the Deed and the other wet-ink documents are
signed) is often not scheduled yet. The EXECUTED line and the notary date then print only what is
known, with blanks for the rest:

| Known | The EXECUTED line and the notary date read | `signing_date` value |
|---|---|---|
| Exact date (signing booked) | "…this November 14<sup>th</sup>, 2026." / "On this November 14<sup>th</sup>, 2026, before me…" | `"2026-11-14"` |
| Month and year | "…this November _____, 2026." | `"2026-11"` |
| Year only | "…this ______________ _____, 2026." | `"2026"` |
| Nothing | "…this ______________ _____, _______." | `""` |

(The old blank wording "____ day of ______________, 20___" is no longer used.)
(built by `notary_block.py`).

**Bold and plain text — the same for every seller type.** Every POA uses
the templates' Times New Roman 11pt; only the formatting below is bold:

| Part of the POA | Bold | Plain |
|---|---|---|
| Opening paragraph | The principal (individual: the names; entity: the entity name; trust / estate: the whole "A and B, as Trustees of the … Trust" / "A, Personal Representative of the Estate of …" phrase, as in the firm's templates); **[Attorney in fact]**; the property address; **$[Price]** | Everything else — **including the Buyer name(s)** |
| Signature lines | Nothing | Everything — the trust, entity, or estate name above "By:", the "By:" line, and the signer's name and title |
| Notary block | The signer names | Everything else |
| Captions and notice | The signature-page caption (bold) and the page-break notice (bold italic) | — |

The one size difference: a trust, entity, or estate name too wide for its signature column is
shrunk (11pt down to 8pt); an individual's name wraps instead, with a `NAME WIDTH` warning.

**Signature lines — one length for every seller type.** The line a signer
signs on ends **3.0" from where it starts**, the length of the one-seller line in Form 2:
an individual's line is 3.0" long; an entity, trust, or estate line reads "By:" and then the line,
ending at the same 3.0" point (so the line after "By:" is about 2.8"). Every builder draws it as an
**underlined tab to a fixed tab stop** (`SIG_LINE_TWIPS` and `sig_line_runs()` in
`notary_block.py`), not a run of underscores: a tab fills exactly to its stop, so the length is
the same in Word and LibreOffice whatever the font, and the line can never spill onto a second
line. Every signature column is 3.25" wide, so the line always fits. Do not type underscores for
a signer's line. (The notary's own line in the notary block is unchanged.)

**EXECUTED line — flush left (all seller types).** The line "EXECUTED as a
sealed instrument this [date]." starts at the left margin, with **no tab or indent**, in every
POA. The bundled templates (Form 2 and Form 4) and all builder scripts already do this — do not
add a tab before "EXECUTED".

**Defined term for the agreement (all seller types).** Every POA defines the agreement as
**(the "P&S")**: "…as set forth in a Purchase and Sale Agreement dated [P&S date], with
[Buyers] (the "P&S"), for not less than **$[Price]** ("Sales Price")…" — or, when the user
leaves the P&S date out (Step 3), "…a Purchase and Sale Agreement with [Buyers] (the "P&S")…"
(no "dated", no comma). The bundled templates
(Form 2 and Form 4) and all builder scripts already use this wording — do not change it to
(the "Purchase and Sale Agreement"), and **do not add "as amended"** ("for not less than $[Price]" sets a price floor that "as amended" does not change). Instead,
the skill fills in the most updated Buyer and Purchase Price (see "Amendments and addenda" in SKILL.md, Step 2).

---

### Number agreement — singular vs. plural (all seller types)

Every POA must speak consistently in the number of its principal:

| Principal | Opening | Possessive | Object | Trustee words |
|---|---|---|---|---|
| One individual / one trustee | `I,` | `my` (my true and lawful attorney, my "Attorney", my name, my Attorney, my subsequent disability) | `me` (for me) | `Trustee`, `capacity` |
| Two or more individuals / trustees | `We,` | `our` (everywhere `my` would be) | `us` | `Trustees`, `capacities` |
| One estate fiduciary | `I,` | `my` | `me` | `capacity`; capacity title singular ("Personal Representative") |
| Two or more estate fiduciaries | `We,` | `our` | `us` | `capacities`; capacity title plural ("Personal Representatives") |
| Entity | entity name | `its` | — | — |

- **Individuals:** the single-seller template (Form 2) is written in `I/my/me` and the
  multi-seller template (Form 4) in `We/our/us`; `build_individual_poa.py` picks the template by
  seller count and writes the opening paragraph's words to match.
- **Trustees:** `build_trustee_poa.py` picks the words from the number of trustees.
- **Estates:** `build_estate_poa.py` picks the words and pluralizes the capacity at its head noun
  (Personal Representative**s**, Public Administrator**s**, Administrator**s** with the Will Annexed,
  Executr**ices**).
- **Entities:** always singular `its` (the entity is one principal, however many sign).
- The notary block is excluded — its "me"/"my" refer to the notary.

`check_agreement.py` checks the finished file on every path; see "Number Agreement Check" below.

### The notary block (all seller types)

Every POA ends with the same notary acknowledgment, built by `notary_block.py` from the firm's
Notary Acknowledgment spec. Never type or hand-edit the notary text. The entity and trustee
builders (including `build_individual_poa.py`) insert it themselves.

```
COMMONWEALTH OF MASSACHUSETTS
COUNTY OF [COUNTY]

On this September 16th, 2026, before me, the undersigned notary public, [Signers] personally
appeared, proved to me through satisfactory evidence of identification, which
were_________________________, to be the person whose name is signed on the preceding or
attached document, and acknowledged to me that [he/she/they] signed it voluntarily for its
stated purpose[ capacity clause].

______________________________
Notary Public Signature

My Commission Expires: _____________, 20___
```

- **Heading:** "COMMONWEALTH OF MASSACHUSETTS" by default; the notary state from Step 3 otherwise
  ("STATE OF NEW HAMPSHIRE"; see "Reading the notary location" in SKILL.md, Step 3).
- **Date:** the POA signing date in the Date format ("September 16th", superscript suffix), or
  with blanks for the parts not known ("November _____, 2026", "______________ _____, 2026",
  "______________ _____, _______" — see "The signing date may be partly blank" in Step 2).
- **County:** the county from Step 3, in capitals, without the word "County" — or
  "COUNTY OF ______________" when left blank.
- **Several notary appointments:** one notary block per appointment, each naming only that
  appointment's signers (so a one-signer appointment says "the person whose name is signed" and
  that signer's pronoun — he, she, they, or the _____ blank — even in a "We," POA).
- **Signers:** names only, **bold**, joined "A" / "A and B" / "A, B, and C" (joining words plain).
- **One vs. several:** one signer → "the person whose name is signed" and **he**, **she**,
  **they**, or a five-underscore blank **_____** (the user's Step 3 answer; the blank is for the
  notary to fill in when the user doesn't know the signer's pronoun); two or more → "the persons
  whose names are signed" and **they**.
- **Capacity clause** (no period of its own; the sentence ends with one period):

  | Seller type | Capacity clause |
  |---|---|
  | Individual(s) | *(none)* — "…for its stated purpose." |
  | Trust | "as Trustee of [the] [Trust Name]" (one) / "as Trustees of [the] [Trust Name]" (two or more) — **"the" only if the trust name does not begin with "The"** (firm-wide rule) |
  | Entity | "as [Office(s)] for [Entity Name], a/an [Entity Type], as the voluntary act of the [Entity Type]" — **no state** in either mention; offices: one signer "President" (or "President and Treasurer" if one person holds both); same office "Managers"; different offices "President and Treasurer, respectively," |
  | Attorney-in-fact (edge case) | "as attorney in fact for [Principal], the principal, as the voluntary act of the [Principal]" — supported by `notary_block.py`; the rest of an AIF-signed POA is not built yet, so stop and tell the user if the P&S shows the seller signing by an attorney-in-fact |
  | Probate estate | "as [Capacity] of [the] [Estate Name]" — capacity pluralized for 2+ signers ("as Personal Representatives of…"); "the" before the estate name unless it begins with "The" |


## What each builder produces

### Individual sellers

- **Template by seller count** — the builder picks it from the assets folder: **one seller →
  Form 2** (`I` / `my` / `me`); **two or more → Form 4** (`We` / `our` / `us`). Everything
  except the fields below is the template's own text, unchanged (powers paragraph, incapacity
  paragraph, margins, title).
- **Opening paragraph** — rebuilt from the values (no "as amended"):

  > One: I, **Laura Example**, hereby appoint **[Attorney in fact]** as my true and lawful
  > attorney in fact (my "Attorney") for me and in my name, to do all things necessary with
  > respect to the sale of property situated at **[Address]** (the "Premises"), as set forth in a
  > Purchase and Sale Agreement dated [P&S date], with [Buyers] (the "P&S"), for not less than
  > **$[Price]** ("Sales Price"), hereby ratifying and affirming that which my Attorney shall
  > lawfully do or cause to be done by virtue of the powers herein conferred.
  >
  > Two: We, **A** and **B**, hereby appoint … as our … (our "Attorney") for us and in our name, …
  >
  > Three or more: We, **A, B, C,** and **D**, hereby appoint … — serial
  > comma; the names and the commas between them bold, the final "and" plain.

  The Buyer name(s) stay plain (not bold), as in the templates — and in every seller type's POA.
- **EXECUTED line** — "EXECUTED as a sealed instrument this October 30<sup>th</sup>, 2026." (flush
  left, no tab), or with blanks for the parts of the date not known ("…this November _____, 2026.").
- **Signature lines**:
  - **One seller** — Form 2's single signature line on the right-hand side, with the name below.
  - **Two or more** — Form 4's borderless two-column table, filled **left then right, two
    signers per row**: 2 sellers = 1 row; 3 = 2 rows (third seller on the left); 4 = 2 full rows;
    5 = 3 rows (fifth on the left), and so on. Each row is set not to split across pages.
  - A name too long for its column (about 3.1") wraps onto a second line; the builder prints a
    `NAME WIDTH` warning — tell the user.
- **Notary acknowledgment** — the shared notary block ("The notary block" above), no capacity
  clause; signers bold and joined "A" / "A and B" / "A, B, C, and D".
- **Several notary appointments** — the counterparts paragraph and one counterpart section per
  appointment (its own EXECUTED line, only its sellers' signature lines — one seller: Form 2's
  line; two or more: the two-column table — and its own notary block), as for every seller type.


### Entity sellers

- **Opening paragraph:** "**[Entity Name]**, a [State] [Entity Type], hereby appoints **[Attorney in
  fact]** as its true and lawful attorney in fact (the "Attorney") for [Entity Name]
  and in its name, to do all things necessary with respect to the sale of property situated at
  **[Address]** (the "Premises"), as set forth in a Purchase and Sale Agreement dated
  [September 1<sup>st</sup>, 2026], with [Buyers] (the "P&S"), for not less than **$[Price]**
  ("Sales Price"), hereby ratifying and affirming that which its Attorney shall lawfully do or
  cause to be done by virtue of the powers herein conferred." (No "as amended…" language.)
- The standard "WITHOUT limiting the foregoing…" powers paragraph.
- **No incapacity sentence** — it applies to natural persons, not entities, and is always
  omitted for an entity seller (do not ask the user about it).
- "EXECUTED as a sealed instrument this September 16<sup>th</sup>, 2026." (no leading tab;
  ordinal suffix in superscript).
- **Signature lines** — a borderless two-column table (two 3.25" columns, zero cell margins).
  Each signer gets his or her own column, and each column repeats the entity name:

  ```
  Friends Of The Example Harbor Shelter, Inc.    Friends Of The Example Harbor Shelter, Inc.

  By:______________________                           By:______________________
  Olivia Example, its President                          Beatrice Sample, its Treasurer
  ```

  - Entity name **not bold**; the `By:` line (`By:` and the signature line, ending 3.0" from the column's
    left edge — "Signature lines" above) and the
    `[Name], its [Title]` line are plain too, all flush left in the column.
  - Signers fill the columns left, then right, two per group; a third or fourth signer starts
    a new group two blank rows below, with the entity name repeated above each signer.
  - A **single signer** goes in the right-hand column (the left column is left blank).
  - Each row is set not to split across pages.
- **Entity-name font fitting** — the entity name in the signature table must stay on one line.
  The builder measures the name in Times New Roman (or the metric-identical Liberation
  Serif) and, if it is too wide for its column with a 14-point gap before the next column,
  reduces **only that name's** font size from 11pt in 0.5pt steps to a minimum of 8pt. The
  entity name elsewhere in the document stays 11pt. If the name is still too wide at 8pt it
  will wrap; the builder warns, and you tell the user.
- **Notary acknowledgment** — the shared notary block ("The notary block" above) with the
  entity capacity clause. **No state** appears in it (the opening paragraph keeps "a
  Massachusetts Nonprofit Corporation"):

  | Signers | End of the notary sentence |
  |---|---|
  | **One** (she) | …acknowledged to me that she signed it voluntarily for its stated purpose as President for [Entity Name], a Nonprofit Corporation, as the voluntary act of the Nonprofit Corporation. |
  | **Two, different offices** | …they signed it … as President and Treasurer, respectively, for [Entity Name], a Nonprofit Corporation, as the voluntary act of the Nonprofit Corporation. |
  | **Three or more, different offices** | …they signed it … as President, Treasurer, and Secretary, respectively, for [Entity Name], … |
  | **Two or more, same office** (e.g., LLC managers) | …they signed it … as Managers for Acme Realty LLC, a limited liability company, as the voluntary act of the limited liability company. |
  | **One person, two offices** (he) | …he signed it … as President and Treasurer for [Entity Name], a corporation, as the voluntary act of the corporation. |

  "a" or "an" follows the entity type's first sound ("an LLC", "a corporation").
- **No footer** (the "Seller's Limited Power of Attorney" footer was removed).


### Trustee sellers

- **Opening paragraph** (the firm's Trustee Seller template language; **bold** as shown):

  > One trustee: I, **Jane Doe, as Trustee of the Doe Family Trust**, u/d/t dated March
  > 3<sup>rd</sup>, 2015, hereby appoint **[Attorney in fact]** as my true and lawful attorney in
  > fact (my "Attorney") for me and in my name, in my capacity as Trustee, to do all things
  > necessary with respect to the sale of property situated at **[Address]** (the "Premises"), …
  >
  > Two or more: We, **Jane Doe and John Doe, as Trustees of the Doe Family Trust**, u/d/t dated
  > …, hereby appoint **[Attorney in fact]** as our true and lawful attorney in fact (our
  > "Attorney") for us and in our name, in our capacities as Trustees, to do all things necessary…

  **u/d/t date left out** (`include_trust_date: false`): the whole ", u/d/t dated [date]" phrase
  is omitted and only one comma remains — "I, **Jane Doe, as Trustee of the Doe Family Trust**,
  hereby appoint…". The builder checks that no "u/d/t" and no doubled comma are left.

  Three or more trustees are listed "A, B, and C". The paragraph continues as in the individual
  templates: "…as set forth in a Purchase and Sale Agreement dated [P&S date], with [Buyers]
  (the "P&S"), for not less than **$[Price]** ("Sales Price"), hereby
  ratifying and affirming that which my/our Attorney shall lawfully do or cause to be done by
  virtue of the powers herein conferred."
- **"the" rule** — everywhere except the signature lines, a lowercase **"the"** is placed before
  the trust name unless the name already starts with "The" ("as Trustee of the Doe Family Trust";
  "as Trustees of The Doe Family Trust"). The signature lines show the trust name exactly as
  given, with no "the" added.
- The standard "WITHOUT limiting the foregoing…" powers paragraph.
- The incapacity paragraph with my/our (unless `include_incapacity` is `false`).
- "EXECUTED as a sealed instrument this September 16<sup>th</sup>, 2026."
- **Signature lines** — the same borderless two-column table as the entity path. Each trustee gets
  a column with the trust name repeated above:

  ```
  Doe Family Trust                    Doe Family Trust

  By:______________________           By:______________________
  Jane Doe, Trustee                   John Doe, Trustee
  ```

  Trust name **not bold**; everything in the signature lines is plain. One
  trustee → right-hand column; two → left, right; a third
  or fourth starts a new group below. **Trust-name font fitting** works exactly as for entities
  (measured in the regular, not bold, font):
  if the trust name is too wide for its column, only that name's font size in the signature lines
  is reduced from 11pt in 0.5pt steps to a minimum of 8pt (the trust name elsewhere stays 11pt).
  If the name is still too wide at 8pt it will wrap; the builder warns, and you tell the user.
- **No "as Trustee and not individually" language** anywhere in the POA (the firm's choice).
- **Notary acknowledgment** — the shared notary block ("The notary block" above), ending
  "…acknowledged to me that he/she/they (or _____) signed it voluntarily for its stated purpose as Trustee of the
  Doe Family Trust." (one trustee) or "…they signed it … as Trustees of the Doe Family Trust."
  (two or more). The no-double-"the" rule applies: "as Trustees of The Doe Family Trust".
- **No footer** (the "Seller's Limited Power of Attorney" footer was removed).

---

### Probate estate sellers

- **Opening paragraph** — no "as" before the first capacity:

  > One: I, **Ellen K. Example, Personal Representative of the Estate of Edith A. Sample**,
  > Worcester Probate and Family Court Docket No. WO24P0000EA, hereby appoint **[Attorney in fact]** as my
  > true and lawful attorney in fact (my "Attorney") for me and in my name, in my capacity as
  > Personal Representative, to do all things necessary with respect to the sale of property
  > situated at **[Address]** (the "Premises"), as set forth in a Purchase and Sale Agreement dated
  > [P&S date], with [Buyers] (the "P&S"), for not less than **$[Price]** ("Sales Price"),
  > hereby ratifying and affirming that which my Attorney shall lawfully do…
  >
  > Two or more: We, **A and B, Personal Representatives of the Estate of X**, … as our … (our
  > "Attorney") for us and in our name, in our capacities as Personal Representatives, …

- Court cited "[County] Probate and Family Court Docket No. [Docket]"; no appointment date.
- **Public Administrator → "for"**: when the capacity is "Public
  Administrator", the word between the capacity and the estate name is **"for"** instead of
  "of" — "I, **Paul Publico, Public Administrator for the Estate of John Doe**, …" (plural:
  "Public Administrators for the Estate of…"). The notary clause matches: "…as Public
  Administrator for the Estate of John Doe." Every other capacity uses "of".
- **"the" rule** — "the" before the estate name unless it begins with "The".
- The standard powers paragraph (the "all deeds" power is **always** kept, for every seller type,
  even when the P&S requires the deed signer to sign personally).
- The incapacity paragraph with my/our.
- "EXECUTED as a sealed instrument this May 19<sup>th</sup>, 2026."
- **Signature lines** — the same two-column table as the entity and trustee paths, estate name
  **not bold** and **without a leading "The"**:

  ```
                                      Estate of Edith A. Sample

                                      By:______________________
                                      Ellen K. Example, Personal Representative
  ```

  One signer → right column; long names shrink 11pt → 8pt as for trusts and entities.
- **Notary** — "…acknowledged to me that she signed it voluntarily for its stated purpose as
  Personal Representative of the Estate of Edith A. Sample." (plural: "they … as Personal
  Representatives of …").
- **Docket check** — prints whether the docket's division prefix matches the court county
  (BA Barnstable, BE Berkshire, BR Bristol, DU Dukes, ES Essex, FR Franklin, HD Hampden,
  HS Hampshire, MI Middlesex, NA Nantucket, NO Norfolk, PL Plymouth, SU Suffolk, WO Worcester).

---

### Signature Page and Notary Section (all paths)

The section that begins with "EXECUTED as a sealed instrument…" and runs through the end of
the notary acknowledgment ("My Commission Expires…") — including all signature lines — must
**always stay together on one page**. Where that page is depends on the **signature page
placement** (Firm Settings, or the user's Step 3 answer when the setting is `ask`):

- `same-page` → run with `--mode same`: the section stays on the same page as the POA text when
  it fits; if any part of it would fall on the next page, the whole section starts on the next
  page instead.
- `separate-page` → run with `--mode separate`: the section always starts on its own page.

`[MODE]` in the Step 4 command means `same` or `separate`; `build_all_poas.py` runs
`keep_together.py` on every finished POA in that mode. It:

1. Sets **keep-with-next** and **keep-lines-together** on every paragraph of the section, and
   **can't-split** on every table row in it, so Word moves the section as one block even if
   later edits change the page layout.
2. Renders the document (LibreOffice) and checks where the section lands. In same-page mode, if
   it fits where the POA text ends, nothing else is added.
3. **Signature page** — whenever the section starts on its own page (separate-page mode, or
   same-page mode when it doesn't fit), the script adds a page break and opens the signature page
   with this caption, **bold and flush left**, followed by one blank line and then "EXECUTED…":

   > **Signature Page to Seller's Limited Power of Attorney for the sale of [Inline Property Address].**

   The address comes from `build_all_poas.py` (`--address`, the shared `address` value). No page numbers are added
  .
4. **Page-break notice** — on the page where the POA text ends, if **1.5 inches (about six
   lines) or more** of blank space is left below the last paragraph, the script adds
   ***[Remainder of page intentionally left blank; signature page follows]*** (bold italic,
   brackets included) on its own line, so a reader knows the space is intentional and nothing
   was omitted. The notice is **centered horizontally**, and vertically it sits about
   **one-quarter of the way down** the empty space — the space between the last line of text on
   that page and the bottom margin. The script measures where it lands and
   adjusts the space above it until it is in place. With less blank space, there is no notice. If the POA
   text itself runs past page 1, the notice goes on the last page of text.
5. With no page break, there is no caption and no notice, and any leftover
   `{{Page Break Language}}` placeholder paragraph is removed. The script starts from the plain
   document each time, so it is safe to run again (e.g., in the other mode).
6. Prints the result, e.g. `same-page mode -- section fits on page 1; no page break needed`,
   or `separate-page mode -- signature page is page 2 with caption; notice added, 1/4 down (4.2"
   blank) (OK)`, or `… no notice (1.2" blank, under 1.5")`.

7. **Several notary appointments** (Step 3) — the script finds one EXECUTED section per
   appointment and puts **each on its own page**, whatever the placement mode, with the caption
   "**Counterpart Signature Page [k] of [n] to Seller's Limited Power of Attorney for the sale of
   [address].**" (bold, flush left) and one blank line before "EXECUTED…". The notice rule is the
   same, with the wording ***[Remainder of page intentionally left blank; signature pages
   follow]***. It prints e.g. `2 notary appointments -- counterpart signature pages 2, 3, each
   with its caption; notice added, 1/4 down (3.4" blank) (OK)`.

Do **not** add a page break, caption, or notice by hand. If LibreOffice is not available, the
script sets the keep-together properties only and says so — tell the user to check the page
layout in Word (and, in separate-page mode, that the signature page was not set up).

---

### Number Agreement Check (all paths)

`build_all_poas.py` runs `check_agreement.py` on every finished POA. It reads the body above "EXECUTED"
and prints `AGREEMENT: OK (…)` or `AGREEMENT: PROBLEM -- …` naming each word that disagrees
with the opening ("I," → no our/us/Trustees/capacities; "We," → no my/me/Trustee/capacity;
entity → no my/me/our/us), plus a missing "the" (or a double "the The") before a trust name.
For an estate POA it also checks the "the" rule before the estate name (opening and notary) and
that the capacity title is singular with "I," and plural with "We," in the opening, the
"capacity as" phrase, and the notary clause.
It also checks the notary block: person/persons, he/she/they/_____ vs. they, and Trustee/Trustees match
the number of signers (and the "I,"/"We," opening), the trust "the" rule in the capacity
clause, no double period, and no `{{…}}` placeholder left anywhere. With several notary
appointments it checks **each** notary block on its own (one signer → person/he/she/they/_____, even in a
"We," POA), one EXECUTED line per notary block, and the counterparts paragraph. A problem stops
`build_all_poas.py` before anything is saved; fix it, run again, and tell the user what was fixed.

---
