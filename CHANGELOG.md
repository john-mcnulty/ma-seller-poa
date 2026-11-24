# Changelog

The rules live in `skills/seller-poa/SKILL.md`; this file keeps the history of how they came to
be, so the skill itself stays lean.

## 0.2.0 — Cowork-first: firm settings that persist (unreleased)

- **The firm settings are asked once and saved.** Cowork (and a skill installed on its own) has
  no plugin settings, so the skill used to ask for the attorney in fact on every run and never
  offered the firm's signing locations. On the first run it now asks once, in one question, for
  the attorney in fact, the signing locations and the signature page, and saves them as
  `seller-poa-firm-settings.json` in the firm's Clients folder. Later runs find the file in the
  P&S folder's parent or grandparent and ask nothing; the notary location question offers the
  firm's offices. "Change my firm settings" shows each value and where it came from, and saves.
- `firm_settings.py`: new `--search-from` / `--settings-file` (read the file) and `--save` (writes
  only after every value checks out: attorney set and not a placeholder, every signing location
  readable, a valid signature page and state; keeps the file's other values). The file is listed
  in `sources`; the doctor shows it.
- Claude Code is unchanged: plugin settings still win over the file.
- README: "Use it in Cowork (recommended)" comes first; Cowork needs nothing installed.
- CI: the regression suite also runs on Windows; every job has a 30-minute limit. The Windows
  install step retries Chocolatey (its feed can answer 503) and fails at once if pandoc or
  LibreOffice is still missing.
- Tests: 9 settings-file checks (read, save, refusals, precedence, nothing set still asks).
  **The runner now finds LibreOffice in its Windows install folder**: on Windows it used to look
  only on the PATH, so it silently skipped the text comparison. All 45 scenarios now compare
  equal on Windows too.

## 0.1.3 — Windows console text (unreleased)

- **A Windows console can no longer stop a build.** The Word validator printed "→", which a
  Windows console's default code page can't encode, so the validator crashed and the build saved
  nothing (found in a live estate-seller run). Every script now writes UTF-8 to the console, and
  `build_all_poas.py` runs the builders and the validator with UTF-8 output. Reproduced and
  verified with the Windows code page (cp1252).
- The test runner reads its children's output as UTF-8 too.

## 0.1.2 — scanned PDFs (unreleased)

- **New `render_pages.py`** renders PDF pages to images for reading scanned pages (PyMuPDF first,
  poppler's `pdftoppm` as a fallback), writing only to a temporary folder. SKILL.md now names it
  wherever it says to render scanned pages; before, Claude reached for `pdftoppm`, which a
  Windows machine usually doesn't have. The doctor checks it ("Page images for scanned PDFs").
- The doctor finds LibreOffice's version on Windows more reliably.

## 0.1.1 — fixes from the first live runs on Windows (unreleased)

- **No more console window.** On Windows, `soffice.exe --version` opens a console that waits for
  Enter; `doctor.py` now reads LibreOffice's version from its install folder instead of running
  it, and every helper program (LibreOffice, pdftotext, pandoc, the builders) is started without
  a window.
- **A layout failure no longer stops the build.** If LibreOffice or the PDF reader errors or times
  out, the POA is still saved with the signature section kept together, and the message says the
  caption / notice were not added and the layout should be checked in Word.
- **PDF reader:** PyMuPDF is used first when installed; a `pdftotext` is used only if it supports
  `-bbox` (poppler's). The `pdftotext` bundled with Git for Windows does not, and made the layout
  step fail; it is now detected, reported by the doctor, and skipped.
- **Firm settings read directly.** In a live run the attorney in fact was not filled into the
  skill even though the plugin settings were saved, so the skill asked for the name.
  `firm_settings.py` now also reads the saved plugin settings (`pluginConfigs` in
  `~/.claude/settings.json`); the doctor shows where each value came from.
- The version number is bumped on every change from now on: Claude Code installs an update only
  when the version changes.

## 0.1.0 — first public release (unreleased)

- Published as the plugin **ma-seller-poa** (marketplace `ma-seller-poa`); the skill inside is
  `seller-poa`, invoked as `/ma-seller-poa:seller-poa`.

Prepared for publication from the author's private working skill.

- **Firm settings.** The attorney in fact is now a required setting (`attorney_in_fact`) read by
  every builder; there is no default, and the scripts refuse an empty value or an unresolved
  `${user_config…}` placeholder. Signature page placement and the default notary state are
  settings too. Plugin installs use plugin settings; skill-only installs use
  `firm-settings.json`.
- **Signing locations.** New setting `signing_locations` (`Label = County, State; …`): the notary
  location question offers the firm's offices as a pick list, the first one recommended, so the
  user no longer chooses a state and a county. A typed name ("boston") is matched to a location.
  Replaces the single "default signing county" setting.
- **Entity signer question** always offers at least two options (the question tool rejects a
  question with one option): "[signers] (Recommended)" / "Different or additional signers", or
  one option per reading when the sources differ.
- **Estate POAs** cite the "[County] Probate and Family Court Docket No." (was "Probate Court"),
  the court's name as it appears on the P&S and the Letters of Authority.
- **Scripts moved out of SKILL.md** into `skills/seller-poa/scripts/` (they were re-typed on every
  run); SKILL.md went from about 4,200 lines to about 1,200. Formatting detail and the notary
  appointments walk-through moved to `references/`.
- **`doctor.py`** checks the required tools (Python, pandoc, LibreOffice, a PDF reader) and the
  settings. `keep_together.py` can read the rendered page with PyMuPDF when poppler's
  `pdftotext` is not installed (identical results on the whole test suite).
- `build_all_poas.py` accepts relative paths for the assets, values file, and output folder.
- **Templates rebuilt** with placeholder sample text and empty document properties.
- **Regression suite** (`tests/`): 45 fictional scenarios with expected text, plus a test that
  proves the commit gate blocks what it claims to.

## History before open-sourcing (2026)

- **Oct 11–12** — First version: one individual seller from a signed P&S, filled into the firm's
  POA template; the firm's Notary Acknowledgment form.
- **Oct 18–19** — Probate estate sellers (Personal Representatives and other fiduciaries; "for"
  with a Public Administrator; docket division check). The POA always carries the most updated
  Buyer and Purchase Price from signed amendments, and never says "as amended" ("for not less
  than" sets a price floor). The "all deeds" power is kept for every seller type.
- **Oct 26** — 0.75" bottom margin. Signature page placement (`same-page` / `separate-page` /
  `ask`), the signature-page caption, and the "remainder of page intentionally left blank" notice
  placed one-quarter of the way down the blank space. The EXECUTED line is flush left.
- **Oct 27** — Any number of individual sellers (one builder for all). Notary state and county
  are inputs; the county may be left blank. Separate notary appointments with counterpart
  signature pages. Word-file P&S input.
- **Oct 28** — Partly known signing dates ("November _____, 2026"). One POA per principal when the
  sellers are of more than one type, with file names that name each principal. The u/d/t date
  question, with its source. The footer was removed.
- **Oct 29** — One bold / plain text scheme for every seller type; names in signature lines are
  plain.
- **Nov 18–19** — The client folder rule (where related files are looked for). Names checked
  against the vesting deed (the deed's spelling controls for a matched seller or an entity).
- **Nov 20** — Word-file P&S reading restored after it was lost from an earlier version.
- **Nov 23** — The trustee question is skipped when the documents agree (riders do not count).
  One-signer pronoun: He / She / They / Leave blank, ordered silently by the signer's name.
  Notary appointments across several POAs. A blank or partly blank P&S date gets one question
  (suggested date / carry over / leave out / typed). Every signer's line is an underlined tab to a
  fixed 3.0" stop.

## Open items

- **B5** California all-purpose acknowledgment (the skill currently warns).
- **B6** Confirm which states' notaries use "COMMONWEALTH OF".
- **B8** Notary profiles (commission expiration). **B9** Optional ID type in the notary block.
- **B10** Compare the price wording with the REBA form.
- **B14** Several vesting deeds that disagree. **B16** No fixed closing date (date suggestion
  only). **B18** Date options after a past closing date.
- A seller signing through an attorney-in-fact (the skill stops cleanly).
- A corporate trustee ("XYZ Bank, as Trustee") — not handled.
