# ma-seller-poa

A Claude skill (packaged as a Claude Code plugin) that drafts a **Massachusetts Seller's Limited
Power of Attorney** as a Word document from the Purchase and Sale Agreement — for the seller's
attorney in a residential real estate sale.

> **A drafting aid, not legal advice.** Every POA this produces is a draft that a qualified
> attorney must review before anyone signs it. Using it creates no attorney-client relationship
> with anyone. The wording and several checks are **Massachusetts-specific** (see
> [Limitations](#limitations)).

---

## Requirements — install these first

| | What | Why | Install |
|---|---|---|---|
| **Required** | **Python 3.9+** | Runs the builders (standard library only) | [python.org](https://www.python.org/downloads/) |
| **Required** | **pandoc** | Reads a Word P&S; extracts each POA's text to verify it | [pandoc.org/installing](https://pandoc.org/installing.html) — macOS `brew install pandoc` · Windows `winget install JohnMacFarlane.Pandoc` · Debian/Ubuntu `sudo apt install pandoc` |
| **Required** | **LibreOffice** | Renders each POA so the signature page can be laid out: the EXECUTED line, signatures and notary block kept on one page, the signature-page caption, and the "remainder of page intentionally left blank" notice | [libreoffice.org/download](https://www.libreoffice.org/download/) — macOS `brew install --cask libreoffice` · Windows `winget install TheDocumentFoundation.LibreOffice` · Debian/Ubuntu `sudo apt install libreoffice-writer` |
| **Required** | **A PDF reader** — the **PyMuPDF** package (recommended) **or** poppler's `pdftotext` | Reads where things landed on the rendered page | **`python -m pip install pymupdf`** (any system; the simplest choice, especially on Windows) · or poppler: macOS `brew install poppler` · Debian/Ubuntu `sudo apt install poppler-utils`. **Note:** other programs named `pdftotext` (for example the xpdf version bundled with some Windows tools) lack the `-bbox` option and do not work — the skill detects this and uses PyMuPDF instead. |
| Optional | Pillow | Measures long trust / entity / estate names exactly, to fit them in the signature lines | `python -m pip install Pillow` |
| Optional | Times New Roman or Liberation Serif fonts | The fonts names are measured in | Usually already installed |
| Optional | A `docx` skill with `validate.py` | An extra structural check of each finished `.docx` | Claude's built-in docx skill, where available |

And [Claude Code](https://claude.com/claude-code) (or another Claude surface that can run skills
with a shell — see [Install](#install)).

**What happens without LibreOffice or a working PDF reader?** The POAs are still built and checked, but
the layout step cannot run: the signature section is only marked "keep together" in Word, no
signature-page caption or page-break notice is added, and nothing confirms the signatures fit on
the page. The skill tells you this before it builds and asks whether to go on; you would then
check the page layout yourself in Word.

**Check your machine** at any time (it also checks your settings):

```bash
python skills/seller-poa/scripts/doctor.py --attorney-in-fact "Alex Q. Attorney, Esq."
```

```
  ok       REQUIRED     Python 3.9+                          3.12.4
  ok       REQUIRED     pandoc                               pandoc 3.1.3
  ok       REQUIRED     LibreOffice (soffice)                LibreOffice 24.2.7.2
  ok       REQUIRED     PDF reader (pdftotext or PyMuPDF)    pdftotext
  ...
READY
```

---

## What it does

Point it at a P&S (a signed PDF, preferred, or a Word file) and say "do the seller POA". It:

1. Reads the sellers, Buyer, Purchase Price, P&S date, and property address — and the current
   Buyer and price from any signed amendments in the client folder.
2. Works out **who the principal is** and builds the right POA for each:
   - **individual sellers** — any number;
   - **entity sellers** — corporations (President and Treasurer by default), LLCs (all managers),
     and others, cross-checked against the Secretary of the Commonwealth's record when available;
   - **trustee sellers** — one or more trustees, with or without the u/d/t date;
   - **probate estate sellers** — Personal Representatives, Public Administrators, and other
     fiduciaries, with the docket checked against the county;
   - **mixed sellers** (e.g., an individual and a trust) — one POA per principal.
3. Checks the sellers' names against the **vesting deed** when it is in the client folder.
4. Asks a few short questions: signing date (or what to leave blank), **where the POA will be
   notarized — picked from your firm's signing locations**, the notary pronoun for a lone signer,
   whether everyone signs at the same appointment, and whether the Buyer or price changed.
5. Builds each POA, lays out the signature page, checks singular/plural agreement throughout
   (I/We, my/our, Trustee/Trustees, he/she/they), and saves
   `Power of Attorney. Seller. [Property Address].docx` next to the P&S. It never overwrites a file.

### A worked example (fictional)

A P&S for **123 Sample Street, Hingham, MA 02043**, sellers *Peter F. Example and Jessica H.
Example*, Buyers *William F. Buyer and Jessica Buyer*, 500,000.00, dated September 1, 2026. Your
settings: attorney in fact *Alex Q. Attorney, Esq.*, signing locations `Main office = Plymouth
County, MA; Boston office = Suffolk County, MA`.

The skill asks where the POA will be notarized and offers **Main office (Recommended)**,
**Boston office**, and **Massachusetts — leave the county blank**. You pick Main office and a
signing month, and get a one-page POA that opens:

> We, **Peter F. Example and Jessica H. Example**, hereby appoint **Alex Q. Attorney, Esq.** as our
> true and lawful attorney in fact (our "Attorney") for us and in our name, to do all things
> necessary with respect to the sale of property situated at **123 Sample Street, Hingham, MA
> 02043** (the "Premises"), as set forth in a Purchase and Sale Agreement dated September 1st,
> 2026, with William F. Buyer and Jessica Buyer (the "P&S"), for not less than **$500,000.00**
> ("Sales Price") …

with both signature lines and a Massachusetts notary acknowledgment for Plymouth County.

---

## Install

### Claude Code (plugin — recommended)

```
/plugin marketplace add john-mcnulty/ma-seller-poa
/plugin install ma-seller-poa@ma-seller-poa
```

You are asked for the settings below. From a terminal instead:

```bash
claude plugin marketplace add john-mcnulty/ma-seller-poa
claude plugin install ma-seller-poa@ma-seller-poa \
  --config "attorney_in_fact=Alex Q. Attorney, Esq." \
  --config "signing_locations=Main office = Plymouth County, MA; Boston office = Suffolk County, MA" \
  --config "signature_page=same-page" \
  --config "default_notary_state=Massachusetts"
```

Then run the doctor (above), and use it: *"Prepare the seller POA from the P&S in this folder."*
or `/ma-seller-poa:seller-poa`.

To update later, run both commands, then restart Claude Code (check the installed version with
`claude plugin list`; see `CHANGELOG.md` for what changed):

```
claude plugin marketplace update ma-seller-poa
claude plugin update ma-seller-poa@ma-seller-poa
```

### The skill on its own (e.g., uploaded to Claude)

The skill is the folder `skills/seller-poa/`. Zip that folder and add it as a skill where your
Claude plan allows custom skills. **Before zipping, edit `skills/seller-poa/firm-settings.json`**
— there are no plugin settings in this mode, and the skill refuses to build until
`attorney_in_fact` is set. The surface must be able to run Python and the required programs above.

---

## Settings

**Set every one explicitly.** In a plugin install, an option you leave unset can reach the skill as
a literal `${user_config.…}` placeholder; the skill treats that as missing (it never passes a
placeholder into a document), but setting them removes the guesswork. Change them any time with
`/plugin configure ma-seller-poa@ma-seller-poa`.

| Setting | What it does | Default |
|---|---|---|
| `attorney_in_fact` | **Required.** The attorney every POA appoints, exactly as it should print (`Alex Q. Attorney, Esq.`). There is no default — the skill will not build a POA until it is set, so no POA can ever name someone you did not choose. | none |
| `signing_locations` | Your firm's usual signing places, so the notary location is a **pick from a list** instead of a state and a county. Format `Label = County, State`, entries separated by `;`. The **first entry is the default** (offered first, as Recommended). | none — you type the place each time |
| `signature_page` | `same-page` keeps the signatures and notary block on the same page as the POA text when they fit (more transparent to the principal — no one can swap the text later); `separate-page` always uses a captioned signature page (handy when the text may be corrected after signing); `ask` asks every time. | `same-page` |
| `default_notary_state` | The state for a place that gives none, and for the "leave the county blank" option. | `Massachusetts` |

### Signing locations

Examples:

```text
Main office = Plymouth County, MA
Main office (100 Example St, Hingham) = Plymouth County, MA; Boston office = Suffolk County, MA; Nashua office = Hillsborough County, NH
Hingham office = Plymouth, Massachusetts; Mobile signings = MA
```

- The label is what the user sees; put the street address in it if that helps.
- "County" is optional; the state may be a name or a postal code.
- A Massachusetts county must be one of the 14; a close misspelling is corrected (and reported).
- `= MA` with no county means "Massachusetts, county left blank" — for a notary who travels.
- Up to three locations show as options (the question tool allows four, and one is kept for
  "leave the county blank"). The rest can be typed by name — "boston" matches "Boston office".
- The user can always type any other place — a county and state, or just a city or town.

Check how the skill reads them:

```bash
python skills/seller-poa/scripts/firm_settings.py --location-options \
  --signing-locations "Main office = Plymouth County, MA; Boston office = Suffolk County, MA"
```

In `firm-settings.json` (skill-only installs) the same list can be written as objects:

```json
"signing_locations": [
  {"label": "Main office", "county": "Plymouth", "state": "MA"},
  {"label": "Boston office", "county": "Suffolk", "state": "MA"}
]
```

Settings are read in this order: the plugin setting filled into the skill → the plugin settings
Claude Code saved at install (`pluginConfigs` in `~/.claude/settings.json`, read directly) → the JSON
file named by the `SELLER_POA_SETTINGS` environment variable → `skills/seller-poa/firm-settings.json`
→ the default. The doctor's "settings came from" line shows which one each value came from.

---

## Where files go, and what leaves your machine

- **Output:** the POAs are saved in the folder of the P&S you gave it (or, for a P&S attached in a
  chat with no folder, the session's working folder), never overwriting an existing file. Each
  POA is built and checked in a temporary folder first; if any check fails, **nothing** is saved.
- **Reading:** the skill reads the P&S and, when it can, related files in the same client folder
  (amendments, the vesting deed, trust and probate papers, the corporations record). It never
  searches above the client folder.
- **No telemetry.** The scripts make no network connections at all — no analytics, no usage
  reporting. A run's details are client data and stay on your disk.
- The documents you give Claude are processed by Claude like any other file in your session,
  under your Claude plan's terms.

## Limitations

- **Massachusetts.** The notary acknowledgment is the Massachusetts form. For another state the
  heading changes ("STATE OF NEW HAMPSHIRE") and the skill reminds you to confirm that state's
  notary will accept the wording; **California** requires its own acknowledgment form, which is not
  built (the skill says so). Probate docket codes, G.L. c. 190B, and the Secretary of the
  Commonwealth records are Massachusetts-specific.
- **Not handled (the skill stops and says so):** a seller signing through an attorney-in-fact, a
  Voluntary Administrator, a foreign (out-of-state) fiduciary. A corporate trustee ("XYZ Bank, as
  Trustee") is not supported.
- **Not checked:** whether a trustee is still serving, or whether a probate sale needs a license
  (a reminder only).
- The list of states whose notaries use "COMMONWEALTH OF" (MA, PA, VA, KY) is provisional.

## Development

```bash
python tests/run_tests.py          # 45 fictional scenarios, text compared with tests/expected/
python tests/test_scrub_gate.py    # proves the commit gate blocks what it claims to
python scripts/scrub_check.py      # the commit gate
```

`tests/run_tests.py` builds every scenario in `tests/scenarios/` (individuals 1–5, corporations,
LLCs, long names, trustees, estates, notary appointments within and across POAs, dates, out-of-state
notaries, and error cases that must stop with nothing saved) and compares each POA's text with
`tests/expected/`. After an intended change, `--update` rewrites the expected files — review the
diff. **Every name and address in `tests/` is fictional; keep it that way.**

### The commit gate

`scripts/scrub_check.py` runs before every commit (`python scripts/scrub_check.py --install-hook`
once, after cloning) and in CI on every push, over every commit in history. It blocks client data
by **shape**: PDFs, images, Word files, run values files, home-directory paths, email addresses,
credentials. The two Word templates are the only `.docx` files allowed, each pinned by its exact
path **and** SHA-256, and each is opened and checked for comments, tracked changes, hidden text,
author metadata, external links, and media. The maintainer additionally runs a private
name-identity check that never enters this repository.

**Found a real name, address, or document reference in here?** Please report it privately — see
[SECURITY.md](SECURITY.md). A public issue would republish it.

## License

MIT — see [LICENSE](LICENSE).

Author: John McNulty <john.mcnulty80@gmail.com>

Not affiliated with or endorsed by the Massachusetts Probate and Family Court, the Secretary of
the Commonwealth, any registry of deeds, REBA, dotloop, or DocuSign. Names are used only to
describe the documents the skill reads.
