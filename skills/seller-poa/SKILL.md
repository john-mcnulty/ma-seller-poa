---
name: seller-poa
description: "Generate a Seller's Limited Power of Attorney Word document (.docx) for a Massachusetts residential real estate sale by reading the Purchase and Sale (P&S) Agreement (a signed PDF, or a Word .docx/.doc file) and filling the bundled POA template. Use whenever the user asks to create, generate, or fill out a Seller Power of Attorney or Seller POA, says \"run the POA\" or \"do the POA\" while pointing at a P&S, or provides a P&S and needs the seller's POA. Handles any number of individual sellers, entity sellers (corporation, nonprofit corporation, LLC, or similar), trustee sellers (one or more trustees of a trust), and probate estate sellers (Personal Representatives or other court-appointed fiduciaries). Mixed seller types get one POA per principal."
---

# Seller Power of Attorney Generator

Drafts a Seller's Limited Power of Attorney Word document by reading a Purchase and Sale (P&S)
Agreement — a signed PDF (preferred) or a Word file — and filling a bundled POA template with
what it finds. Supports **individual sellers — any number**, **entity sellers** (corporation,
nonprofit corporation, LLC, etc.), **trustee sellers** (one or more trustees of a trust), and
**probate estate sellers** (one or more Personal Representatives or other court-appointed
fiduciaries). The right wording is chosen from who the SELLER is in the P&S. When the SELLERs
are of **more than one type** (e.g., an individual and a trust), or there are several trusts,
estates, or entities, the skill builds **one POA per principal** (Step 2, "Principals — one POA
each").

**The output is a draft for an attorney to review.** It is not legal advice. The notary wording
and several checks are Massachusetts-specific (the 14 counties, Probate and Family Court docket
codes, G.L. c. 190B, the Secretary of the Commonwealth's corporations records). Say so if the
property or the seller is outside Massachusetts.

Below, `[SKILL_BASE_DIR]` means the base directory shown in the skill header at the top of this
file. Every script is in `[SKILL_BASE_DIR]/scripts/` and the two templates are in
`[SKILL_BASE_DIR]/assets/`. Run the scripts from there — do not copy or re-type them.

## Requirements (check once per session)

```bash
python "[SKILL_BASE_DIR]/scripts/doctor.py" \
  --attorney-in-fact "${user_config.attorney_in_fact}" \
  --signing-locations "${user_config.signing_locations}" \
  --signature-page "${user_config.signature_page}" \
  --notary-state "${user_config.default_notary_state}"
```

It checks the **required** tools — Python 3.9+, **pandoc** (reads a Word P&S and verifies each
POA), **LibreOffice** (`soffice`), and a **PDF reader** (the PyMuPDF package, or poppler's
`pdftotext` — another program named `pdftotext` without the `-bbox` option does not count, and
the doctor says so) — plus the optional Pillow and fonts, and the firm settings below. Its last line says:

- `READY` → go on.
- `NOT READY (layout)` → LibreOffice or the PDF reader is missing. Tell the user, with the install
  command the doctor printed. The POAs can still be built, but without the layout step: no
  signature-page caption, no page-break notice, and no check that the signature section fits —
  the user must check the layout in Word. Ask whether to install first or go on; if they go on,
  repeat that warning in Step 5.
- `NOT READY` → a required item or the attorney-in-fact setting is missing: tell the user what to
  install or set (the README has the commands) and stop until it is fixed (for the setting, see
  below).

## Firm Settings

Every POA appoints the firm's attorney in fact, and the notary location question offers the
firm's usual signing locations. Read the settings first:

```bash
python "[SKILL_BASE_DIR]/scripts/firm_settings.py" --show \
  --attorney-in-fact "${user_config.attorney_in_fact}" \
  --signing-locations "${user_config.signing_locations}" \
  --signature-page "${user_config.signature_page}" \
  --notary-state "${user_config.default_notary_state}"
```

In a plugin install the four values are the user's plugin settings, filled in above. **If a value
still reads as a literal `${user_config…}` placeholder** (the skill was installed on its own, or
the setting was never set), pass an empty string `""` for that flag instead — a shell rejects the
placeholder text. The script then reads the plugin settings that Claude Code saved at install
(`pluginConfigs` in `~/.claude/settings.json`), then `[SKILL_BASE_DIR]/firm-settings.json` (or the
file named by the `SELLER_POA_SETTINGS` environment variable), then the defaults below. So
**never ask the user for a setting the script found**: check `errors` and `sources` in its output
first. Do the same in every
command in this file that passes these flags. Use `python3` wherever `python` is not available.

| Setting | What it is | Default |
|---|---|---|
| `attorney_in_fact` | The attorney the POA appoints, exactly as it should print (e.g. `Alex Q. Attorney, Esq.`) | **None — required** |
| `signing_locations` | The firm's usual signing places, `Label = County, State`, separated by `;` — the **first** is the default (e.g. `Main office = Plymouth County, MA; Boston office = Suffolk County, MA`) | none (the user types the place) |
| `signature_page` | `same-page`, `separate-page`, or `ask` (see below) | `same-page` |
| `default_notary_state` | The state used when a place gives no state, and for the "leave the county blank" option | `Massachusetts` |

- **`errors` not empty** (the attorney in fact is not set): tell the user, ask for the name, and
  do not build until they give it. Use the name they give for this run (put it in
  `attorney_in_fact` in the values file, Step 4) and tell them how to save it: `/plugin configure
  ma-seller-poa@ma-seller-poa` in Claude Code, or `firm-settings.json` in the skill folder.
- **`warnings`** (a signing location skipped or a county corrected): mention them once.
- Use the printed `mode` (`same` / `separate`) in Step 4. When `signature_page` is `ask`, ask in
  Step 3.

**Why the signature page default is `same-page`.** One page is more transparent to the Principal:
with the text and the signatures on the same page, no one can later swap in different text (for
example, turn a limited POA into a "blanket" one). A separate signature page is a practical option
when the POA may be signed first and the text corrected afterwards (a typo, or a change the
Attorney in Fact submits to the Principal for approval) before it goes to the Buyer side. Either
way the script adds the safeguards described in `references/document-format.md`: a caption tying
the signature page to the property, and a notice when page 1 ends with a large blank space.

---

## Output

One new Word file per principal, saved to the same folder as the input P&S Agreement:
```
Power of Attorney. Seller. [Inline Property Address].docx                      (one POA)
Power of Attorney. Seller. [Inline Property Address]. [Principal's Name].docx   (two or more POAs)
```
Most transactions have one principal and so one POA. See Step 5 for the principal's name. When
the P&S has no folder on disk (e.g., it was attached in a chat), save the POAs to the working
folder (the outputs folder, if the session has one) and send them to the user.

---

## Step 1: Word tools

Every POA is built by a script — never hand-edit the XML. If a `docx` skill is available, read its
`SKILL.md`: its `validate.py` checks each finished POA (pass its base directory to Step 4 as
`--docx-skill`; below, `[DOCX_SKILL_DIR]`). Without one, leave `--docx-skill` out — the
builders' own checks still run.

---

## Step 2: Extract Information from the P&S Agreement

### Reading the P&S (PDF or Word)

The signed PDF is the usual input; a Word file works too (for the rare case when the signed PDF
isn't available).

- **Both in the folder?** If the client folder (see "The client folder" below) has a signed PDF
  **and** a Word version of the P&S, use the **signed PDF** and mention the Word file was not used.
- **PDF** — read it with the `Read` tool (in chunks if long). Scanned pages: render them to
  images and read them.
- **Word `.docx`** — extract the text with pandoc, keeping tables (the parties, price, and
  signature blocks are often in tables):

  ```bash
  pandoc "[P&S].docx" --track-changes=accept -t plain --wrap=none -o ps_accepted.txt
  pandoc "[P&S].docx" --track-changes=all -t plain --wrap=none -o ps_redline.txt
  ```

  Read `ps_accepted.txt` (the document as it reads with every change accepted). Then:
  - **Tracked changes / comments** — check for them (`unzip -p "[P&S].docx" word/document.xml`
    and count `<w:ins ` / `<w:del `; a `word/comments.xml` part means comments). If there are any,
    tell the user, and **compare the fields this skill uses** (seller names, buyer, price, P&S date,
    property address) in the accepted text against the original text (pandoc
    `--track-changes=reject`). Name every difference — e.g., "a tracked change adds the middle
    initials: Karen **L.** Example, Helen **M.** Specimen". The POA uses the **accepted** version by
    default; confirm in Step 3.
  - **Unsigned draft** — a Word file normally has no signatures. Tell the user the POA was built
    from a Word version, not the signed P&S, and ask in Step 3 to confirm the sellers, Buyer,
    Purchase Price, and P&S date match the **signed** agreement.
- **Older Word (`.doc`), `.rtf`, `.odt`** — convert to `.docx` first, then read as above:
  `soffice --headless --convert-to docx --outdir . "[P&S].doc"`. (Apple Pages files are not
  supported — ask for a PDF or Word export.)
- Never edit or re-save the user's P&S file; work on a copy.

### The client folder — where to look for related files (all seller types)

Several steps look for other files about the transaction: signed amendments and extensions
(below), the state corporations record (Step 2A), the Trustee's Certificate, deed into the trust,
or Declaration of Trust (Step 2B), the probate papers and deed into the decedent (Step 2C), and
the vesting deed into the Seller (Step 2D).
Look for all of them in the **client folder and its subfolders**, worked out like this:

1. **Start with the folder the P&S is in.** Usually that folder *is* the client folder (e.g.,
   `Clients\Example Development Group LLC\`, with the P&S saved directly in it).
2. **Go up one level only when the P&S is in a document subfolder** — a folder named for a kind of
   document rather than a client or a property (e.g., "P&S", "PS", "Purchase and Sale",
   "Contract", "Contracts", "Offer", "Agreements", "Signed", "Signed Docs"). The client folder is
   then the folder above it (`Clients\Smith\P&S\` → `Clients\Smith\`).
3. **Search the client folder and every subfolder in it** (a few levels deep is plenty) — a
   recorded deed or a signed amendment is often filed in a subfolder such as "Closing Docs
   Signed" or "Title".
4. **Never search above the client folder.** Do not go up when the folder above holds other
   clients' folders or is a general folder ("Clients", "Documents", "Desktop", "Downloads", a
   cloud-sync root, a home folder, or a drive root). If the P&S itself sits in such a general folder,
   do not search it or its subfolders (they are not this client's files): say so, and continue
   with the P&S alone (the questions that ask for a missing file, such as the state record in
   Step 2A, still ask).
5. **No folder to search** (e.g., the P&S was attached in a chat with no folder connected):
   say so once and continue with the P&S alone. A missing related file never stops the workflow
   by itself; each step says what to do without it.
6. **Tell the user which folder you searched** (and that you went up a level, if you did) in the
   first message after reading the P&S, so a wrong guess can be caught. Name each file you use
   from it.

The finished POA is still saved next to the P&S (Step 5), not in the client folder above it.

Extract these fields:

| Field | Where to find it |
|-------|-----------------|
| **Seller name(s)** | PARTIES AND MAILING ADDRESSES — the SELLER entry. Extract **name(s) only**, without any legal descriptor (e.g., extract "Laura Example" not "Laura Example, an unmarried woman"). If there are two or more sellers, note each name separately — Seller 1, Seller 2, Seller 3, … in the order they appear. Each name is then checked against the vesting deed (**Step 2D** — the deed's spelling controls by default for a matched seller). |
| **Seller count** | Count the number of distinct individuals listed as SELLER — **any number** is handled. Watch for sellers joined by repeated "and" ("A and B and C") as well as commas, and check the count against the SELLER signature block. |
| **Seller type** | Individual(s) or **entity**. The SELLER is an entity when the name is a business or organization — e.g., it ends in or contains "Inc.", "Incorporated", "Corporation", "Corp.", "LLC", "L.L.C.", "Limited Liability Company", "Company", "Co.", "LP", or "LLP", or the descriptor calls it a corporation or company. The SELLER is a **trustee seller** when the entry names one or more people "as Trustee(s) of" a trust, or the name is a trust (e.g., "Realty Trust", "Nominee Trust", "Revocable Trust", "Family Trust", "u/d/t"). The SELLER is a **probate estate seller** when the entry names one or more people as Personal Representative(s) (or another probate capacity — Successor / Special Personal Representative, Public Administrator, Executor/Executrix, Administrator/Administratrix, Administrator with the Will Annexed, Administrator d.b.n.) of an estate, usually with a Probate Court docket number. **Stop and tell the user** if the capacity is **Voluntary Administrator / Voluntary Personal Representative** (voluntary administration is not available when the estate owns Massachusetts real estate) or the seller is a **foreign (out-of-state) fiduciary** — neither is handled. If the P&S shows the seller acting **through an attorney-in-fact** (e.g., "by Mary Smith, her attorney in fact"), also stop and tell the user — that edge case is not built yet. **Classify each seller separately** — if the SELLERs are of **more than one type**, or there are several trusts, estates, or entities, each principal gets its own POA (see "Principals — one POA each" below). |
| **Entity name** *(entity only)* | The entity's name exactly as it appears in the SELLER entry (keep "Inc.", "LLC", etc. and its capitalization), without the descriptor. Example: from "Friends Of The Example Harbor Shelter, Inc., a Massachusetts Nonprofit Corporation, with an office location at…" extract "Friends Of The Example Harbor Shelter, Inc." This is the starting point: when the vesting deed is found, **its wording controls by default** (Step 2D). |
| **State of organization** *(entity only)* | The state in the SELLER descriptor (e.g., "Massachusetts" from "a Massachusetts Nonprofit Corporation"). The vesting deed's wording controls by default (Step 2D). |
| **Entity type** *(entity only)* | The entity type in the SELLER descriptor, worded as in the P&S (e.g., "Nonprofit Corporation", "corporation", "limited liability company"). The vesting deed's wording controls by default (Step 2D). |
| **Entity signer(s)** *(entity only)* | The SELLER signature block of the P&S — each `By: [Name], its [Title]` line (e.g., "By: Olivia Example, its President" and "By: Beatrice Sample, its Treasurer"). Record **every** signer, in order, as a name + title pair. Use the name as typed in the signature block, not an e-signature stamp (dotloop/DocuSign stamps may spell it differently — if they differ, point it out to the user). The title is the office only (`President`, `Treasurer`, `Manager`, `Managing Member`), without "its". Who actually signs the POA is decided in **Step 2A**. |
| **Trust name** *(trustee only)* | The trust's name exactly as it appears in the SELLER entry, keeping capitalization and a leading "The" only if the P&S has it (e.g., "Doe Family Trust" or "The Doe Family Trust"). Do not add "the" yourself — the builder applies the "the" rule (Step 4E). |
| **Trustee name(s)** *(trustee only)* | Every person named as trustee in the SELLER entry or the SELLER signature block, names only, in the order they appear. Who signs is confirmed in **Step 2B**. |
| **Declaration of Trust date** *(trustee only)* | The date of the Declaration of Trust in the SELLER entry (e.g., "u/d/t dated March 3, 2015" or "under Declaration of Trust dated…"). If the P&S does not give it, look for a Trustee's Certificate, the deed into the trust, or the Declaration of Trust in the client folder (Step 2B). **Record where it came from** — the source and the words that show it — **or, if not found, every place you looked**; Step 3 shows this to the user. Never guess it. A missing date never stops the workflow: the user can leave the u/d/t phrase out (Step 3). |
| **Probate signer(s)** *(estate only)* | Every fiduciary named in the SELLER entry or signature block, names only, in order. |
| **Probate capacity** *(estate only)* | The capacity **as written in the P&S**, singular, without "as" (e.g., "Personal Representative", "Public Administrator", "Successor Personal Representative"). For a Public Administrator, use just "Public Administrator" (the builder then writes "for the Estate of…"). Free text — never hardcode "Personal Representative". If co-fiduciaries have different capacities, stop and ask the user. |
| **Estate name** *(estate only)* | As in the SELLER entry, e.g. "Estate of Edith A. Sample" — **the P&S controls**; Step 2C flags a different reading of the decedent's name in the deed into the decedent, the title exam, the Letters, or the Decree; keep a leading "The" only if the SELLER entry has it (the builder applies the "the" rule and drops "The" on the signature lines). |
| **Probate county / docket** *(estate only)* | The court county ("Worcester" from "Worcester Probate and Family Court") and the docket number ("WO24P0000EA") as two separate values. If either is missing, ask the user in Step 3. The builder checks the docket's division prefix against the county. |
| **Property address (inline)** | DESCRIPTION section — street address, municipality, and zip all on one line (e.g., "29 Sample Road, Braintree, MA 02184") |
| **P&S signing date** | Near the top: "This [day] day of [Month], [Year]" — record the date (e.g., February 11, 2026). **All seller types** show it in the POA as **"February 11th, 2026"** — month, day with its ordinal suffix in superscript, comma, year (see "Date format" in `references/document-format.md`). **Blank or partly blank date line (all seller types, PDF or Word):** record which parts the line gives (e.g., "This ____ day of February, 2026" → month and year) and any **date to suggest** — on a signed P&S, the **last** e-signature date in the document (dotloop / DocuSign stamps — e.g., buyer 02/18/26, seller 02/19/26 → February 19th, 2026); on a Word draft with no e-signatures, a date in the file name (e.g., "5-18-2022"). Then ask the **P&S date question** in Step 3; never guess and never fill it in without asking. |
| **Buyer name(s)** | PARTIES section — the BUYER entry. Names only — drop "or their nominee" / "or nominee" and any descriptor. |
| **Sale price** | PURCHASE PRICE section — the total amount |
| **Closing date** | TIME FOR PERFORMANCE / DELIVERY OF DEED section — the date the deed is to be delivered; a later date in an **extension** in the client folder (see "Amendments and addenda" below) replaces it, even one signed by only one party (note that in the label). Used **only** to suggest the signing month in Step 3; it is not printed in the POA. If none is found, leave it out. |

**Principals — one POA each (all paths).** Give **each**
person or party named in the SELLER entry and the SELLER signature block its own type:
individual, entity, trustee, or estate fiduciary. Then group them into **principals** — whoever
one POA speaks for:

| Principal | Who is in it | One POA for |
|---|---|---|
| **Individual** | **All** the individual sellers together (any number) | all of them, as today ("I," / "We,") |
| **Trustee** | The trustee(s) of **one** trust | that trust — two different trusts are two principals |
| **Estate** | The fiduciary(ies) of **one** estate | that estate |
| **Entity** | The signer(s) for **one** entity | that entity |

Most P&S Agreements have **one** principal, and the skill builds one POA, exactly as before. When
there are **two or more** — sellers of more than one type (an individual and a trust; two
individuals and an LLC; a trustee and a Personal Representative), or several trusts, estates, or
entities — the skill builds **one Seller's Limited POA per principal**, each with its own type's
wording, signature lines, number agreement, and notary capacity clause. Every POA carries the same
Buyer, Purchase Price, P&S date, and property. There is no combined, mixed-type POA.

Some phrasings that are **one** principal, not several: two or more trustees of the same trust;
co-Personal Representatives of the same estate; several officers or managers signing for one
entity; spouses or other co-owners signing as individuals.

**Before Step 3, tell the user** what you found when there are two or more principals — every
seller, the type you read, the P&S words that show it, and the POAs you will build, e.g.:

> This P&S lists sellers of more than one type, so I'll prepare two POAs:
> 1. John Smith — individual ("John Smith")
> 2. Smith Family Trust — Jane Smith, Trustee ("Jane Smith, as Trustee of the Smith Family Trust
>    u/d/t dated March 3, 2015")

so the user can correct a misreading in the Step 3 answers (e.g., the "trust" is only part of an
individual's mailing address → one individual POA). A person named in **two capacities** (e.g.,
individually **and** as trustee) is in both principals and signs **both** POAs; say so.

All the existing stops still apply to each principal (Voluntary Administrator, foreign fiduciary,
a seller signing through an attorney-in-fact): if any principal hits one, stop and tell the user
before building any POA. `build_individual_poa.py` refuses a seller name that looks like a trust,
estate, or entity (a backstop that catches such a seller put in the individual POA by mistake).

**Amendments and addenda — changes to the Buyer or the Purchase Price (all seller
types).** The POA names the Buyer and the Purchase Price, so it must carry the **current**
values, not just those in the original P&S:

1. **Check the client folder** (see "The client folder" above) for every signed **amendment**
   and **addendum** to the P&S (file names containing "Amendment", "Addendum", "Rider",
   "Extension", or "Modification"; also check a combined file such as "PS Agreement & Both
   Amendments"). Read each one (render scanned pages to images) in date / number order, and
   note any change to the **Buyer** (added, removed, or substituted buyer; an assignment; a
   buyer's nominee) or the **Purchase Price**.
2. **Flag every change** to the user, naming the document (e.g., "Second Amendment: Purchase
   Price amended from 500,000.00 to 530,000.00"). Also mention an unsigned or partly signed
   amendment (e.g., signed by the Buyer only) — it does not count as a change until signed by
   both parties; ask the user.
3. **Default to the most recent values.** Populate the POA with the most updated Buyer and the
   most updated Purchase Price from the latest signed amendment or addendum that changes each.
   With no change found, use the original P&S values.
4. **Ask the user in Step 3** (always, whether or not changes were found) whether the Buyer and
   whether the Purchase Price have changed — see Step 3. The user's answer overrides the
   default.

Changes to other terms (credits, deposits) do not affect the POA — no need to report them. A
change to the **closing date** is not printed in the POA either, but the latest closing date is
used to suggest the signing month (Step 3, question 1).

**How the POA reads** — the date format ("September 1st, 2026", superscript suffix), partly
blank dates, bold and plain text, signature lines, the EXECUTED line, and the defined term (the
"P&S") — is set by the builders and described in `references/document-format.md`. Two rules
matter while gathering the facts: the POA carries the **most updated Buyer and Purchase Price**
(above), and it never says "as amended" ("for not less than $[Price]" sets a price floor that an
amendment does not change).


## Step 2A: Determine the Entity Signers (entity sellers only)

Who signs for an entity seller is a **material issue** in a real estate sale, so work it out
carefully and show the user where each name came from. First decide what kind of entity the
seller is, then apply its default rule below. If neither the P&S descriptor nor the name shows
the kind of entity, ask the user for the entity type before working out the signers.

| Entity kind | How to recognize it | Default signers | Title on the `By:` line |
|---|---|---|---|
| **Corporation** | Type or name says "corporation", "Nonprofit Corporation", "business corporation", "professional corporation", or the name ends in "Inc.", "Incorporated", "Corp.", or "Corporation" | The **President** and the **Treasurer**, in that order | `President` / `Treasurer` |
| **LLC** | Type or name says "limited liability company", or the name ends in "LLC", "L.L.C.", or "Limited Liability Company" | **All managers**, in the order the state record (or else the P&S) lists them | `Manager` |
| **Other** (limited partnership, etc.) | Anything else that is not a trust or estate | No default rule yet — see below | As stated in the P&S |

### The state corporations record

Both default rules may need the entity's record from the state's corporations/business
entity division (in Massachusetts, the Secretary of the Commonwealth Corporations Division
business entity summary, annual report, or certificate of organization). To get it:

- First look in the client folder (Step 2, "The client folder") for a likely file (names containing "Entity
  Summary", "Annual Report", "Certificate of Organization", "Corporations", "Business Entity",
  or "Secretary of State"). If one is there, tell the user which file you found and use it.
- If none is there, **ask the user to upload the record or point you to it** (use
  `AskUserQuestion`; options are listed under each rule below). Wait for the answer before
  going on.
- Records often print names in capitals — convert them to normal capitalization (e.g.,
  `BEATRICE SAMPLE` → `Beatrice Sample`), keeping middle names and initials.

### Corporations — the President and the Treasurer

For a corporation, the default signers are **always the President and the Treasurer, in that
order**, whoever else the P&S lists. Find each officer's name in this order, stopping once both
are known:

1. **The P&S signature block** — a `By: [Name], its President` line and a
   `By: [Name], its Treasurer` line.
2. **The state corporations record**, for whichever officer the P&S does not name. If no record
   is in the client folder, ask: "Upload or point me to the record" / "I don't have it — I'll type the
   officer's name". In the record's officers table, take the name on the `PRESIDENT` row and
   the name on the `TREASURER` row.
3. **Ask the user** — as the backup, if the P&S and the state record still leave the President
   and/or the Treasurer unknown, ask the user for the name of each missing officer.

Even when the P&S names both officers, if a state record is already in the client folder, read it
too for the cross-checks below. (Don't ask the user for the record when the P&S already names
both officers.)

If the **same person holds both offices**, use one signer with both titles:
`[Name], its President and Treasurer`.

### LLCs — all managers

For an LLC, the default signers are **all of the LLC's managers**, each with the title
`Manager`. Because the P&S alone cannot show whether it lists *every* manager, the state
record is the primary source for an LLC:

1. **The state corporations record** — take every person listed as a manager (in a
   Massachusetts record, the "Manager(s)" section). If no record is in the client folder, ask:
   "Upload or point me to the record" / "The P&S lists all of the managers" / "I don't have
   it — I'll type the managers' names".
2. **The P&S signature block** — any `By: [Name], its Manager` lines. Use these when the user
   confirms the P&S lists all managers, and compare them against the record when both exist.
3. **Ask the user** — as the backup, if neither source identifies the managers, ask the user
   for the name of each manager.

**If the LLC has no managers** (e.g., the record lists none, which usually means it is
member-managed), the default rule does not apply: tell the user and ask who will sign and in
what capacity (e.g., "Member" or "Managing Member").

If the Massachusetts record also lists **persons authorized to execute, acknowledge, deliver
and record instruments affecting an interest in real property**, compare that list with the
managers and tell the user about any difference — keep all managers as the default.

If the P&S gives a manager a different title (e.g., "Managing Member"), keep `Manager` as the
default and mention the difference.

### Other entities (limited partnership, etc.)

No default rule yet. Use the signer(s) in the P&S signature block; if there are none, ask the
user for each signer's name and title.

### Cross-checks (all entities)

Never guess a signer and never leave a default signer out. Tell the user about each of these
(do not silently resolve them):
- The P&S and the state record name **different people** for the same role, or **spell a name
  differently** (including e-signature stamps). Prefer the state record's spelling as the
  default, and say so.
- The record shows a **term as expired** before the POA signing date, or the record looks out
  of date — the user should confirm the signer still holds the role.
- The P&S names signers **other than** the default signers (e.g., a Clerk of a corporation, or
  a member of a manager-managed LLC) — keep the default, and mention the difference.

---

## Step 2B: Determine the Trustees (trustee sellers only)

The default signers are **all of the trustees** named in the P&S, each signing as `Trustee`.

- If the client folder has a **Trustee's Certificate**, the **deed into the trust**, or the
  **Declaration of Trust** (file names containing "Trustee", "Certificate", "Declaration", "Trust",
  or "Deed"), read it and tell the user which file you used. Compare its trustees and Declaration
  of Trust date with the P&S.
- Tell the user (do not silently resolve) when the sources name **different trustees**, **spell a
  name differently** (including e-signature stamps), or give a **different trust date**. Prefer the
  recorded document's spelling as the default, and say so.
- Never guess a trustee and never leave one out. Confirm the list in Step 3 — or, when every
  source agrees, use it without asking (Step 3, "Trustees").
- **Trust date — keep a source trail.** For the Declaration of Trust date,
  note for the Step 3 question exactly where it was found — "P&S, SELLER entry (\"u/d/t dated
  March 3, 2015\")", "Trustee's Certificate.pdf, p. 1", "deed into the trust, Book 12345,
  Page 67" — or, when no source gives it, the list of places checked (the P&S, and each file in the
  client folder that you read). When sources give **different** dates, list each date
  with its source. Do not search further afield (registry, web) just for this date: the user can
  leave it out.

---

## Step 2C: Check the Probate Appointment (estate sellers only)

- Look in the client folder (Step 2, "The client folder") for the **Letters of Authority**,
  **Decree and Order** (e.g., "Decree & Order on Petition for Formal Adjudication", "Order of
  Informal Probate"), or a **License / Decree of Sale**. If found, read it (scanned PDFs: render
  the pages to images) and tell the user which file you used.
- Compare with the P&S and **flag every difference** in: **signer name**, **capacity**, **docket
  number**, and **estate name**, including the **decedent's name** (middle name vs. initial,
  spelling, suffix, punctuation). **The P&S controls** for every field — do
  not change the P&S values, just report the differences.
- **Decedent's name vs. the deed into the decedent / title exam.** If the deed by which the
  decedent took title or the title exam is in the client folder (file names containing "Deed" —
  not a draft or signed deed *out* of the estate — or "Title Exam"; the P&S DESCRIPTION's
  Book/Page identifies the deed in), read it (render scanned pages to images) and compare the
  decedent's name in the grantee clause with the decedent's name in the Seller POA. **Flag any
  difference** — the chain of title is strongest when the POA (and the Deed out of the Estate)
  match the deed into the decedent exactly — but keep the P&S name. Tell the user which file you
  checked. If neither is in the client folder, skip this check (no need to ask). Also mention it when the
  decedent took title with someone else (e.g., tenants by the entirety) — the co-owner's death may
  need its own record proof, which is outside this skill.
- **Flag (do not stop):**
  - **Special Personal Representative** — has no statutory power to sell real estate
    (G.L. c. 190B, § 3-715(b) omits (23½)); the builder also prints a `SALE AUTHORITY` warning.
    Build the POA anyway.
  - A Letters of Authority with the **§ 3-108(4) restriction** box checked (late appointment —
    authority limited to confirming title).
  - **No will** (intestate) shown on the decree — a sale needs a Chapter 202 license to sell; with
    a will, the will must give a power of sale or a license must issue. Mention it as a reminder
    only; the skill does not verify the license.
- Never guess a fiduciary and never leave one out. Confirm the list in Step 3.

## Step 2D: Check Names Against the Vesting Deed (entity and individual sellers)

The deed by which the Seller took title — the **vesting deed** (the "deed
in") — is the primary source for how the Seller's name should read, in the POA and later in the
deed out: the chain of title is cleanest when they match. When the vesting deed is available,
check the P&S names against it; when it isn't, use the P&S names. **A missing vesting deed never
stops the workflow.**

This step covers **entity** principals (name, state of organization, and entity type) and **each
individual seller**. Trust and estate principals keep their own rules (Step 2B; Step 2C — for an
estate the P&S controls).

### Find the vesting deed

- Look only in the client folder and its subfolders (Step 2, "The client folder"). Do not search
  the registry or the web for it.
- Identify it from the P&S: the Book and Page (or Land Court document or certificate number) in
  the DESCRIPTION, or another title reference in the P&S.
- Likely files have names containing "Deed", "Vesting", "Current Deed", "Prior Deed", "Deed In",
  or "Title". Read each candidate (render scanned pages to images) and use it only if:
  - it is a **recorded** deed (registry stamp or Book/Page) — not a Word draft; and
  - the **grantee** is the Seller. A deed with the Seller as **grantor** is a deed *out* (e.g.,
    "Deed. 456 Sample St. v1.docx" drafts, or "Deed. Recorded" in a closing folder) — skip it.
- Prefer the deed whose registry stamp matches the P&S Book/Page. If the only deed to the Seller
  shows a different Book/Page, use it and say so.
- A title rundown or title exam can confirm which deed is the vesting deed, but do not take a
  name's spelling from it (rundowns abbreviate and capitalize names).
- If more than one recorded deed into the Seller is found, use the one matching the P&S
  Book/Page and mention the others. (Comparing several vesting deeds that spell the name
  differently is backlog B14, not built.)
- **Not found** → use the P&S values and tell the user where you looked (no question).

### What counts as a difference

- **Counts:** different words, letters, punctuation, or abbreviations — "Group LLC" vs. "Group,
  LLC"; "LLC" vs. "L.L.C."; "Inc." vs. "Incorporated"; "Co." vs. "Company"; a middle initial vs. a
  middle name; "Jr." in one and not the other.
- **Does not count:** capitalization alone. If the deed prints the name in ALL CAPITALS, take the
  words and punctuation from the deed and the capitalization from the P&S (or normal
  capitalization). If the deed uses ordinary mixed case, take its wording as written,
  capitalization included; when capitalization is the **only** difference, use the deed's without
  asking and mention it in Step 5.

### Entity sellers — the deed controls the name, state, and entity type

- From the deed's grantee clause, take the entity **name**, **state of organization**, and
  **entity type** — e.g., "…grant to Example Development Group LLC, a Massachusetts Limited Liability
  Company of 100 Sample Street, Rockland…" → name "Example Development Group LLC", state
  "Massachusetts", type "Limited Liability Company" — and compare each with the P&S.
- **Default: the deed's wording for all three**. They go in the `entity_name`,
  `state`, and `entity_type` values (Step 4D-1), so the opening paragraph, the signature lines,
  the notary clause, and the file name all use them.
- **Possible name change after the deed — flag, do not default.** If the state corporations
  record (Step 2A) names the entity with **different words** from the deed — not merely different
  punctuation, capitalization, or "LLC" / "L.L.C." — the entity may have changed its name, merged,
  or converted since it took title. Then do not default to the deed: in Step 3 offer the deed, P&S,
  and record names with none recommended, and tell the user the POA may need "formerly known as"
  wording (not built). A record that differs only in punctuation (e.g., "Example Development Group,
  LLC") is common — keep the deed's wording and mention the record's.

### Individual sellers — the deed controls each matched name

The grantees on the deed in and the sellers in the P&S are not always the same people — a joint
tenant or a tenant by the entirety may have died, so fewer people sell than took title. So match
**each P&S seller** to a grantee, one at a time, and use judgment:

| Result | When | Default for the POA |
|---|---|---|
| **Exact match** | Same name (capitalization aside) | No change |
| **Match** | Clearly the same person — same first and last name, differing only in a middle name vs. middle initial ("Edward S. Sample" / "Edward Samuel Sample"), a missing middle name or initial, punctuation or spacing ("Van Example" / "VanExample"), or a suffix shown in one name only ("Jr.") | **The deed's spelling** — confirm in Step 3 (deed Recommended) |
| **Possible match** | Could be the same person, but the names alone can't tell you — a nickname or short form ("Bob" / "Robert"), a different last name (marriage or divorce), a different suffix ("Jr." / "Sr.") | **None** — ask in Step 3 with both spellings, none recommended, and say why |
| **No match** | No grantee could be this seller | **The P&S name** — flag it (the seller may have taken title by another deed, by inheritance, or under a former name) |

- Match each grantee to at most one seller, and keep the sellers in the P&S order.
- **A grantee who is not a seller** (e.g., a joint tenant or tenant by the entirety who has died,
  or a co-owner who is not selling): tell the user. It does not change the POA, but the closing may
  need record proof (such as a recorded death certificate) — outside this skill.
- Descriptors ("a married man", "as tenants by the entirety") are not compared — they are not in
  the POA.

### Record the source

For every name used, note where it came from for Step 3 and Step 5 — e.g., "deed in, Book 00000,
Page 000 (Current Deed Probable. 456 Sample St.pdf)" — or "P&S — no vesting deed found (checked: …)".

---

## Step 3: Ask the User Questions

Ask **all questions at once** in a single `AskUserQuestion` call:

1. **POA signing date** — one question that covers both cases: a signing
   that is booked (the user types the date under "Other") and one that is not booked yet (the
   user picks what to fill in). (With several signers, this is the date of **Signer 1's** notary
   appointment.) Build the options with the script — never by hand — from **today's date in the
   user's time zone** and the **closing date** (Step 2):

   ```bash
   python "[SKILL_BASE_DIR]/scripts/notary_block.py" --date-options --today 2026-10-28 --closing 2026-11-20
   ```

   It prints the options as JSON (`label`, `value`, `description`) plus a `note` when the closing
   date has already passed (say so in the question). Use them **as printed and in that order**;
   the first is the recommended one. Question text: "Signing date for the POA — if the signing is
   booked, type the date under Other (e.g., \"Nov 14\"). If not, pick what to fill in:". The rules
   the script applies:
   - **Suggested month** — the current month, or **next month during the last 7 days of a month**;
     but never later than the closing month (the POA is signed at or before closing).
   - **Closing month** — also offered when it is later than the suggested month, labeled with the
     closing date.
   - **Year** — the current year, without asking; but from **December 8th through 31st** next
     year's options are offered too, so the choice of option sets the year (not when the closing
     date is on or before December 15th — the POA is then signed this year).
   - **"______________ _____, _______"** (year blank) — offered **only when the closing date is
     after December 15th** of the current year (an extension into the next year is then the one
     realistic way the POA gets signed next year).
   - At most four options. A booked date always goes under "Other".

   **Reading the answer.** An option → its `value` (`"2026-11"`, `"2026"`, or `""`). A typed date
   → the exact date as `YYYY-MM-DD`; accept any natural form ("Nov 14", "11/14", "14th",
   "Friday the 14th"); with no year, use the first such date **on or after today**. If a typed
   answer is not a date (e.g., "sometime in November"), map it to the nearest option ("2026-11")
   and say so. If an exact date falls **after the closing date**, say so in Step 5 (don't stop).
   No date picker is available in the question tool, so typing is the way to give a booked date.
2. **Notary location** — "Where will the POA be notarized?" (With several signers: "Where will
   Signer 1 ([name]) sign?") Build the options with the script — never by hand — from the firm's
   signing locations (Firm Settings) and the property's county:

   ```bash
   python "[SKILL_BASE_DIR]/scripts/firm_settings.py" --location-options \
     --signing-locations "${user_config.signing_locations}" \
     --notary-state "${user_config.default_notary_state}" --property-county "[Property's county]"
   ```

   It prints up to four options (`label`, `description`, `state`, `county`): the firm's signing
   locations first (the first one Recommended — e.g. "Main office (Recommended)", described as
   "Prints Plymouth County, MA in the notary block"), then the property's county when there is
   room and no firm location already covers it, then "[state] — leave the county blank". Use them
   **as printed and in that order**, and put its `note` in the question text when there is one
   (e.g., more firm locations than fit — the user types one of their names under Other). The user
   can also type anything under "Other": another firm location's name, a county and state
   ("Rockingham County, NH"), or just a **city or town** ("Weymouth", "Concord, NH").
   **Reading the answer:** an option → its `state` and `county` (an option whose `state` is null
   means the user will type the place). Typed text → first try
   `firm_settings.py --resolve-location "[text]" --signing-locations "${user_config.signing_locations}"`
   (a firm location's name, even shortened, e.g. "boston" → "Boston office"); if it is not a
   firm location, see "Reading the notary location" below.
3. **Notary pronoun — only when exactly ONE person signs** (one individual seller, one trustee,
   or one entity signer). Ask which pronoun the notary block should use for that signer —
   options "He" / "She" / "They" / "Leave blank (_____)", no recommended option. "Leave blank" is for when the user doesn't know the signer's gender or pronoun
   preference: the notary block then prints a five-underscore blank — "…acknowledged to me that
   _____ signed it…" — for the notary to fill in (`"pronoun": "blank"`). **Never guess it from a
   name** — but do **order the options silently**: build them with the script,

   ```bash
   python "[SKILL_BASE_DIR]/scripts/notary_block.py" --pronoun-options --name "Joseph P. Example, Jr."
   ```

   which prints the options in order — the more likely of "He" and "She" for the name first, then
   the other, then "They", then "Leave blank (_____)" — and a `lean` (`he`, `she`, or `unknown`).
   Use the options **as printed and in that order** (e.g., "John Example" → He first; "Sarah
   Example" → She first). When the `lean` is `unknown` (a name the built-in lists don't cover, or
   one used by both, such as "Jordan" or "Pat"), put "He" or "She" first by your own reading of
   the name, or keep He first if you can't tell. The ranking is **silent**: no option is marked
   "(Recommended)", no description or question text mentions it, and the user's answer is always
   required — the order never fills in the pronoun. With two
   or more signers, do not ask here: the notary block uses **"they"** automatically (separate
   appointments ask it per appointment — see `references/notary-appointments.md`).
4. **Same notary appointment? — only when TWO OR MORE people sign** (sellers, trustees,
   fiduciaries, or entity signers). "Will all [n] signers sign at the same notary appointment?"
   Options: "Yes — all at the same appointment (Recommended)" / "No — some will sign at a
   different appointment". **No** → follow `references/notary-appointments.md` ("Notary appointments") before building.
   **Several POAs** (two or more principals): ask this **once for all the POAs** — see
   "Several principals" below.

**Entity sellers only** — add these to the same `AskUserQuestion` call:

4. **Entity signer(s)** — Always ask (who signs for an entity is a material issue, so the user
   confirms the signers still hold those roles), and always with **at least two options** — the
   question tool rejects a question with one option:
   - **Sources agree** (or only one source) — "[signers] — [source] (Recommended)" / "Different or
     additional signers (type under Other)". E.g., "Olivia Example, President and Beatrice
     Sample, Treasurer — from the P&S and the state corporations record (Recommended)", or for an
     LLC, "Jane Doe, Manager and John Roe, Manager — all managers per the state record
     (Recommended)".
   - **Sources differ** (Step 2A cross-checks) — one option per reading with its source, the state
     record's reading Recommended, e.g. "Olivia Example, President — state record (Recommended)" /
     "Olive Example, President — P&S signature block". Mention the difference in the question
     text.
   Mention any other cross-check issue from Step 2A (an expired term, extra signers in the P&S)
   in the question text. The user can type different or additional signers under "Other".
5. **State / entity type** — only if the P&S does not state the entity's state of organization
   or entity type (defaults: "Massachusetts" and "corporation"). Do not guess them.

(The signer count for questions 3 and 4 is only known once the signers are confirmed. Ask
whichever fits the recommended signers — question 3 for one person, question 4 for two or more —
and if the user then changes the number of signers, ask the other one in a short follow-up.)

**Trustee sellers only** — add these to the same `AskUserQuestion` call:

4. **Trustees** — **Skip this question when the documents agree**: the P&S
   (SELLER entry and SELLER signature block) and every trust document found in Step 2B
   (Trustee's Certificate, deed into the trust, Declaration of Trust, acceptance or resignation
   of trustee) name the same trustee(s), spelled the same way — capitalization alone does not
   count, as in Step 2D — and at least one such trust document was found. Use those trustees
   without asking, and report them with their sources in Step 5. **Riders, addenda, and
   extensions do not count** for this test: a trustee name that differs
   only there — e.g., Rider B's caption reading "John P. Example, Jr." when the P&S and the
   recorded Trustee Acceptance say "Joseph P. Example, Jr." — does not bring the question back.
   Skip the question anyway, and flag the rider's reading to the user as a likely typo (Step 5).
   Otherwise ask, always with at least two options (the question tool rejects a question with
   one option):
   - **Sources differ** — one option per reading, with its source; the recorded document's
     reading is Recommended (Step 2B), e.g. "Jane A. Doe — Trustee's Certificate and deed into
     the trust (Recommended)" / "Jane Doe — P&S". Mention the difference in the question
     text.
   - **The P&S is the only source** — "[names] — from the P&S (Recommended)" / "Different or
     additional trustees (type under Other)".
   (Same follow-up rule for question 3 as for entities; when the question is skipped, the
   trustee count is already known, so ask question 3 or 4 to match it.)
5. **u/d/t date** — **always ask, for each trust**. The POA includes ", u/d/t
   dated [Trust Date]" by default, but the user may leave it out, so a missing date never stops
   the workflow. The question **names where the date came from**, or that it wasn't found:
   - **Date found:** "Include the u/d/t date in the [Trust Name] POA?" Options: "Include — March
     3, 2015, from the P&S SELLER entry (Recommended)" (or "…from Trustee's Certificate.pdf,
     p. 1"; when two sources agree, name both) / "Leave the u/d/t date out". When sources
     **disagree**, offer one "Include" option per date with its source (none recommended), plus
     "Leave the u/d/t date out".
   - **Date not found:** say so and where you looked, e.g. "No u/d/t date found — checked the P&S
     SELLER entry and the 6 files in the client folder (no Trustee's Certificate, deed into the
     trust, or Declaration of Trust). Leave it out, or type the date under Other." Options:
     "Leave the u/d/t date out for now (Recommended)" / "I'll type the date (use Other)". A typed
     date is used as given — record its source as "typed by the user".
   Leaving it out → `"include_trust_date": false` (Step 4E-1). With several trusts, ask once per
   trust.

**Estate sellers only** — add these to the same `AskUserQuestion` call:

4. **Fiduciaries** — Confirm the signer(s) and capacity from Step 2 / 2C (e.g., "Ellen K.
   Example, Personal Representative — from the P&S (Recommended)"). Mention any Step 2C
   difference in the question text.
5. **Probate county / docket** — only if the P&S doesn't give both.

**All seller types — always ask** (add to the same `AskUserQuestion` call, or a second call
right after if more than four questions apply):

- **Buyer changed?** — "Has the Buyer changed since the P&S?" Options: the Buyer the skill will use
  as the first option, with its source (e.g., "No change — Morgan Buyer (P&S; no amendment
  changes the Buyer) (Recommended)" or "Yes — Jane Roe, per the Third Amendment (Recommended)");
  the user can type a different Buyer via "Other".
- **Purchase Price changed?** — "Has the Purchase Price changed since the P&S?" Options built the
  same way (e.g., "Yes — 530,000.00, per the Second Amendment (Recommended)" / "No — 500,000.00
  (P&S)"); the user can type a different price via "Other".

**P&S date — all seller types, only when the P&S date line is blank or partly blank** (PDF or
Word; backlog B15). A P&S with a full date gets no question. Ask it **once**
(every POA uses it), in the first call or the second call: "The P&S date line is blank
(\"This ____ day of September, 2023\"). What should the POA say?" Options, in this order:

- **The date to suggest** (Step 2), **only when there is one** — Recommended: "September 6th,
  2023 — the last e-signature date (dotloop, 09/06/23) (Recommended)" on a signed P&S, or "May
  18th, 2022 — from the file name (Recommended)" on a Word draft. Never a date from anywhere else.
- **"Carry over incomplete P&S date from the P&S"** (this label, word for word) — keeps the POA's [Month Day, Year] format, fills in the parts the date line gives, and
  blanks the rest. Put what it will print in the option's description: "Prints: dated September
  _____, 2023" (month and year given), "Prints: dated ______________ _____, 2023" (year only), or
  "Prints: dated ______________ _____, _______" (nothing given).
- **"Leave it out"** — "Leave it out — …a Purchase and Sale Agreement with [Buyers] (the
  \"P&S\")…": the words "dated [P&S date]" are dropped from the POA.
- The user can type the date under **"Other"**.

With **no date to suggest**, the question has just the "Carry over…" and "Leave it out" options
(plus Other) and
**neither is Recommended**. There is **no "pause" option**: the question waits, so a
user who needs to look the date up types it under Other when ready, and nothing is built before.

**Reading the answer** → the values file (Step 4): the suggested or a typed date → `ps_date`
`"YYYY-MM-DD"` (a typed date in any natural form, as for the signing date); "Carry over incomplete P&S date" →
`ps_date` with only the parts the line gives (`"2023-09"`, `"2023"`, or `""`); "Leave it out" →
`"include_ps_date": false` (no `ps_date` needed). A date line that gives the **day but not the
month** (rare) can't keep the day in this format — the carried-over date then prints the month and day
blank; say so.

**Word-file P&S only** (Step 2, "Reading the P&S") — add to the same call (or the second call):
- **Confirm against the signed P&S** — "This POA is from a Word version of the P&S, not the
  signed agreement. Do the sellers, Buyer, Purchase Price, and date match the signed P&S?" Options:
  "Yes — they match the signed P&S (Recommended)" / "Not sure — build it and I'll check" (the
  user can type corrections via "Other"). Mention any tracked-change differences in the question.
- **P&S date** — a blank date line is handled by the P&S date question above (the file-name
  date, if any, is its suggested option).

**Names from the vesting deed — entity and individual sellers, when Step 2D found a difference
that counts** (not capitalization alone). Add to the same `AskUserQuestion` call, or a second call
if more than four questions apply:

- **Entity:** "Name for the POA — the deed in (Book [x], Page [y]) reads “[deed wording]”; the
  P&S reads “[P&S wording]”." Options: "Deed: [Name], a [State] [Type] (Recommended)" / "P&S:
  [Name], a [State] [Type]". With a possible name change (Step 2D), offer the deed, P&S, and
  state-record names, none recommended, and say why.
- **Individual sellers — a match:** "The deed in reads “[deed name]” for the P&S's “[P&S
  name]”." Options: "Deed: [deed name] (Recommended)" / "P&S: [P&S name]". When several sellers
  differ, one question may cover them all: "Use the deed spellings for all (Recommended)" / "Use
  the P&S spellings for all" (mixed choices typed under "Other").
- **Individual sellers — a possible match:** both spellings, none recommended, with the reason
  (e.g., "the deed has “Robert J. Smith”; the P&S has “Bob Smith” — possibly a nickname").

**Signature page placement — only when the Firm Settings value is `ask`.** Ask: "Where should the
signature lines and notary block go?" Options: "Same page as the text, when space permits" /
"Separate signature page". (No recommended option.) With `same-page` or `separate-page`, do not
ask — use the setting.

**Several principals (one POA each — Step 2): ask once, apply to all.**

- **Signing date and notary location** (questions 1 and 2) are asked **once** — they set
  **Appointment A**, the first signer's appointment — word them for all ("Signing date for the
  POAs — …", "Where will the POAs be notarized?").
- **Same notary appointment — one question for ALL the POAs**, asked in the
  same call as questions 1 and 2 whenever **two or more different people** sign across the POAs
  (counting a person who signs two POAs once): "Will all the signers — [names, each with the POA
  they sign, e.g. \"Grace Sample (her own POA) and Henry J. Sample (Trust POA)\"] — sign at the
  same notary appointment?" Options: "Yes — all at the same appointment (Recommended)" / "No —
  some will sign at a different place or on a different date". This replaces the per-POA
  question 4. **Yes** → every POA uses questions 1 and 2. **No** → follow "Notary appointments
  across several POAs" in `references/notary-appointments.md`. If the user instead answers under "Other" with a different date or
  place for one POA (e.g., "the trust POA on Oct 31 in Norfolk"), treat it as **No** with those
  appointments already given.
- **Buyer, Purchase Price, and the P&S date question** are asked **once** (every POA uses them).
- **Pronouns** — one answer **per person** who will be **alone in a POA's notary block**, in the
  `--pronoun-options` order for that name, naming the POA ("Smith Family Trust POA — Jane Smith
  signs alone: which pronoun?"). Each POA's notary block names only **its own** signers, so this
  includes every one-signer POA even when its signer attends an appointment with signers of
  another POA. A person who signs two POAs needs only one pronoun answer — use it for both.
  Ask them in the first call when the answer to the same-appointment question can't change who
  is alone (e.g., every POA has one signer); otherwise ask them after the appointments are set up.
- **Type-specific confirmations** (entity signers and state / type, trustees and the u/d/t date,
  fiduciaries and docket) are asked for **each** principal of that type, naming it.

`AskUserQuestion` takes at most four questions per call; if more than four apply, ask the rest in
a second call right after (several principals usually need two or three calls).

The county matters because it goes in the notary acknowledgment section, which certifies *where*
the document was signed — not where the property is located.

### Reading the notary location

- **A firm signing location** — a name from the firm's `signing_locations` (or close to one) is
  that location's state and county; check with `firm_settings.py --resolve-location` (question 2).
- **State** — defaults to the firm's `default_notary_state` (Massachusetts unless set). Accept a
  full name or a postal code ("NH"). The notary block heading is "COMMONWEALTH OF …" for
  Massachusetts, Pennsylvania, Virginia, and Kentucky and "STATE OF …" for every other state (a
  provisional list — backlog B6); "DISTRICT OF COLUMBIA" has no county line and Louisiana prints
  "PARISH OF".
- **County** — for Massachusetts, it must be one of the 14 counties (Barnstable, Berkshire,
  Bristol, Dukes, Essex, Franklin, Hampden, Hampshire, Middlesex, Nantucket, Norfolk, Plymouth,
  Suffolk, Worcester); `notary_block.py` corrects a close misspelling and says so, and stops on
  anything else. "Leave blank" (or an empty value) prints "COUNTY OF ______________".
- **A city or town instead of a county** — work out its county (e.g., Hingham → Plymouth,
  Weymouth → Norfolk, Concord NH → Merrimack). If you are not certain, look it up (web search)
  rather than guess, and name the county you used in the next question or in Step 5 so the user
  can catch a mistake. Never put the town in the notary block.
- **Outside Massachusetts** — the firm's Massachusetts acknowledgment wording is kept, under the
  other state's heading, and the script prints a reminder to confirm that state's notary will
  accept it. **California** needs California's own acknowledgment form (backlog B5, not built
  yet): tell the user plainly that the notary there will need to attach a California
  acknowledgment.

### Notary appointments (signers who will not all sign together)

When the user answers **No** to the same-appointment question (one POA or several), read
`references/notary-appointments.md` and follow it: it walks through the signers one short
question at a time, sets up each appointment (date, place, pronoun), shows the plan, and gives
the `appointments` format for the values file. Each appointment gets its own counterpart
signature page. For a later appointment's place, offer the same options as question 2
(`firm_settings.py --location-options`).

## Step 4: Update the POA Word File

Every POA is built by a script — do not hand-edit the XML. **Every transaction — one principal
or several — is built with one command, `build_all_poas.py`**. For each principal it runs that
type's builder, `keep_together.py`, `check_agreement.py`, and (when given) the docx validator,
and it names the files. It builds and checks every
POA in a temporary folder first and saves them to the P&S folder **only when every POA passes**,
so a failure never leaves a partial set behind.

#### Write the values file

Create `poa_values.json` in the working folder: the **shared** values at the top level, and a
`principals` list with **one entry per principal** from Step 2 (in the order the P&S names them),
each with its `type` and the keys of its own builder:

| `type` | Keys of the entry |
|---|---|
| `individual` | Step 4A-1 (`sellers`, `pronoun`, …) |
| `entity` | Step 4D-1 (`entity_name`, `state` = state of organization, `entity_type`, `signers`, …) |
| `trustee` | Step 4E-1 (`trust_name`, `trustees`, `trust_date`, …) |
| `estate` | Step 4F-1 (`estate_name`, `signers`, `capacity`, `probate_county`, `docket`, …) |

```json
{
  "attorney_in_fact": "Alex Q. Attorney, Esq.",
  "address": "123 Sample Street, Hingham, MA 02043",
  "ps_date": "2026-09-01",
  "buyers": "William F. Buyer and Jessica Buyer",
  "price": "500,000.00",
  "signing_date": "2026-09-16",
  "notary_state": "Massachusetts",
  "county": "Plymouth",
  "principals": [
    {"type": "individual", "sellers": ["John Smith"], "pronoun": "he"},
    {"type": "trustee", "trust_name": "Smith Family Trust", "trustees": ["Jane Smith"],
     "trust_date": "2015-03-03", "pronoun": "she"}
  ]
}
```

| Shared key | Value |
|-----|-------|
| `attorney_in_fact` | **Required.** The `attorney_in_fact` printed by `firm_settings.py --show` (or the name the user gave this run). The scripts refuse an empty value or a `${user_config…}` placeholder |
| `address` | Inline property address |
| `ps_date` | P&S date as `YYYY-MM-DD`; or, when the user chose "Carry over incomplete P&S date from the P&S" (Step 3), only the parts the date line gives: `"YYYY-MM"`, `"YYYY"`, or `""` (blanks for the rest, as for the signing date) |
| `include_ps_date` | Leave it out (default `true`). `false` only when the user chose "Leave it out" (Step 3): the POA then reads "…a Purchase and Sale Agreement with [Buyers] (the "P&S")…", and `ps_date` is not needed |
| `buyers` / `price` | Current Buyer name(s) and Purchase Price (Step 2 amendments / Step 3) |
| `signing_date` | POA signing date from Step 3 question 1: `"YYYY-MM-DD"` (exact), `"YYYY-MM"` (month and year), `"YYYY"` (year only), or `""` (all blank) |
| `notary_state` / `county` | Notary state (default `Massachusetts`) and county from Step 3; `county` `""` for a blank line |
| `appointments` | **Only when the signers will NOT all sign at the same appointment** — the appointments across **all** the POAs, signers by name (Step 3, "Notary appointments across several POAs"); with one POA, the same list works as before |
| `pronouns` | With `appointments`: `{"[name]": "he"/"she"/"they"/"blank"}` for each person alone in a POA's notary block (a one-signer principal's own `pronoun` also works) |

A principal may **override** any shared key in its own entry — e.g., its own `signing_date`,
`county`, or `appointments` (Step 3). With top-level `appointments`, use those instead of
per-principal overrides; a principal's own `appointments`, if given, wins for that POA.
The notary state: an **entity** overrides it only with `notary_state` (its `state` is its state of
organization); any other principal may use `notary_state` or `state`.

#### Run it

```bash
python "[SKILL_BASE_DIR]/scripts/build_all_poas.py" "[SKILL_BASE_DIR]/assets" poa_values.json "[P&S FOLDER]" --mode [MODE] --docx-skill "[DOCX_SKILL_DIR]"
```

`[P&S FOLDER]` is the folder of the input P&S (the output folder, Step 5). `[MODE]` is the `mode`
from `firm_settings.py --show` (`same` or `separate`), or the user's Step 3 answer when the setting
is `ask`. Leave out `--docx-skill` when no docx skill is available. It prints one block per POA — the principal,
its `FILE:` name, and the builder and check messages (listed in each type's section below) — then
`SIGNS MORE THAN ONE POA: …` for anyone in two principals, and a last line:

- `ALL POAS: OK (n built)` → every POA is saved; go to Step 5.
- `ALL POAS: STOPPED -- …` → **nothing was saved** to the folder. Read the reason (a builder's
  check, an `AGREEMENT: PROBLEM`, a validator error, a principal list problem), fix the values and
  run it again, or tell the user. If the reason is that a POA with the same name **is already in
  the folder**, the script does not overwrite it: ask the user whether to replace it, and only on
  a yes rename the old file (e.g., add " (old)") and run again.

The sections below give each type's keys. What each builder produces — the wording, signature
lines, notary block, signature page, and the number-agreement check — is in
`references/document-format.md`; read it when a message needs explaining or the user asks why
the POA reads as it does.

### Step 4A — Individual-Seller POA (any number of sellers)

One script handles every number of individual sellers (`build_individual_poa.py`, run by
`build_all_poas.py`).

#### 4A-1: Write the values file

These are the keys of the **`individual` principal** in `poa_values.json` (Step 4). The example
shows everything the builder receives — `build_all_poas.py` fills in the shared keys (`address`,
`ps_date`, `buyers`, `price`, `signing_date`, the notary state as `state`, `county`) from the top
level. All values from Step 2 / Step 3; plain text — the script handles XML escaping:

```json
{
  "sellers": ["Robert M. Sample", "Karen L. Example", "Helen M. Specimen"],
  "address": "45 Sample Lane, Marshfield, MA 02050",
  "ps_date": "2022-05-18",
  "buyers": "Barbara Buyer",
  "price": "400,000.00",
  "signing_date": "2026-10-30",
  "state": "Massachusetts",
  "county": "Plymouth"
}
```

| Key | Value |
|-----|-------|
| `sellers` | Every seller, names only (no descriptor), in the order of the P&S SELLER entry — **any number**; each name as confirmed in Step 2D / Step 3 (the vesting deed's spelling for a matched seller, by default) |
| `address` | Inline property address |
| `ps_date` | P&S date as `YYYY-MM-DD` (the script formats it "May 18th, 2022"); partly blank or blank, or left out with `include_ps_date`, as in the Step 4 shared-key table |
| `buyers` | Current Buyer name(s) (Step 2 amendments / Step 3), joined with " and " |
| `price` | Current Purchase Price (Step 2 amendments / Step 3), numeric portion with cents (no `$`) |
| `signing_date` | POA signing date from Step 3 question 1: `"YYYY-MM-DD"` (exact), `"YYYY-MM"` (month and year), `"YYYY"` (year only), or `""` (all blank) |
| `state` / `county` | Notary state (default `Massachusetts`) and county from Step 3; `county` `""` for a blank line |
| `appointments` | **Only for separate notary appointments** (Step 3) — signers listed by name; when present, the top-level `signing_date`, `state`, `county`, and `pronoun` are not used |
| `pronoun` | **One seller only:** `he`, `she`, `they`, or `blank` (prints a `_____` blank for the notary to fill in), the user's Step 3 answer (required; never guessed). Leave it out for two or more sellers ("they" is automatic). |
| `confirmed_individuals` | Leave it out. Add `true` **only** after the builder stops with `MIXED SELLER TYPES` on a name that merely looks like a trust, estate, or entity (e.g., the surname "Trust") **and the user confirms** every seller is an individual. |

**Mixed-type backstop:** if any seller name looks like a trust, trustee, estate fiduciary, or entity
("Trustee", "Trust", "Estate of", "Personal Representative", "LLC", "Inc.", …), the builder
**stops without building** and prints `MIXED SELLER TYPES: STOPPED` (and `build_all_poas.py`
saves nothing). That seller belongs in its own principal (Step 2, "Principals — one POA each"):
move it there and run again. Only if it really is an individual (e.g., the surname "Trust") and
the user confirms, add `confirmed_individuals`.

#### 4A-2: Build and check it

`build_all_poas.py` (Step 4) runs the builder, the keep-together check, the agreement check, and the
validator for this POA (against Form 2 for one seller, Form 4 for two or more). Report the
script messages (`TEMPLATE`, any `NAME WIDTH` warning, the `APPOINTMENT` / `NOTARY` lines,
`KEEP-TOGETHER`, `AGREEMENT`).


### Step 4D — Entity-Seller POA (corporation, nonprofit corporation, LLC, etc.)

An entity seller acts through the POA in **its own name**. The document follows the firm's
**Entity Seller POA template (v5)**. `build_all_poas.py` runs `build_entity_poa.py` for it.

#### 4D-1: Write the values file

These are the keys of an **`entity` principal** in `poa_values.json` (Step 4). The example shows
everything the builder receives — `build_all_poas.py` fills in the shared keys from the top
level (the notary state as `notary_state`). All values from Step 2 / Step 2A / Step 3; plain
text — the script handles XML escaping:

```json
{
  "entity_name": "Friends Of The Example Harbor Shelter, Inc.",
  "state": "Massachusetts",
  "entity_type": "Nonprofit Corporation",
  "address": "48 Sample Avenue, Weymouth, MA 02191",
  "ps_date": "2026-09-01",
  "signing_date": "2026-09-16",
  "buyers": "Dana Buyer and Chris Buyer",
  "price": "450,000.00",
  "notary_state": "Massachusetts",
  "county": "Plymouth",
  "signers": [
    {"name": "Olivia Example", "title": "President"},
    {"name": "Beatrice Sample", "title": "Treasurer"}
  ]
}
```

| Key | Value |
|-----|-------|
| `entity_name` | Entity name — from the vesting deed when Step 2D uses it, otherwise the P&S |
| `state` / `entity_type` | State of organization and entity type — the vesting deed's wording when Step 2D uses it, otherwise the P&S's |
| `address` | Inline property address |
| `ps_date` | P&S date as `YYYY-MM-DD` (the script formats it "September 1st, 2026"); partly blank or blank, or left out with `include_ps_date`, as in the Step 4 shared-key table |
| `signing_date` | POA signing date from Step 3 question 1: `"YYYY-MM-DD"`, `"YYYY-MM"`, `"YYYY"`, or `""` (the script formats the EXECUTED line and the notary date, with blanks for the unknown parts) |
| `buyers` | Current Buyer name(s) (Step 2 amendments / Step 3), joined with " and " |
| `price` | Current Purchase Price (Step 2 amendments / Step 3), numeric portion with cents (no `$`) |
| `notary_state` | Notary state from Step 3 (default `Massachusetts`). **Entity path only:** named `notary_state` because `state` is the entity's state of organization |
| `county` | Notary county from Step 3 (e.g., `Plymouth`; the notary block prints it in capitals without "County"); `""` for a blank line |
| `appointments` | **Only for separate notary appointments** (Step 3) — signers listed by `name`; when present, the top-level `signing_date`, `notary_state`, `county`, and `pronoun` are not used |
| `signers` | Every signer confirmed in Step 3, in signing order; `title` is the office only (e.g., `President`, `Manager`, or `President and Treasurer` when one person holds both) |
| `pronoun` | **One signer only:** `he`, `she`, `they`, or `blank` (prints a `_____` blank for the notary to fill in), the user's Step 3 answer (required; never guessed). Leave it out for two or more signers ("they" is automatic). |

#### 4D-2: Build and check it

`build_all_poas.py` (Step 4) runs the builder, the keep-together check, the agreement check, and the
validator for this POA. Report the script messages to the user (`ENTITY NAME FONT: …`,
`KEEP-TOGETHER: …`, and `AGREEMENT: …`). If the builder warns that the name will wrap even at
the minimum size, tell the user.


### Step 4E — Trustee-Seller POA (one or more trustees of a trust)

`build_all_poas.py` runs `build_trustee_poa.py` for it.

#### 4E-1: Write the values file

These are the keys of a **`trustee` principal** in `poa_values.json` (Step 4) — one per trust.
The example shows everything the builder receives; `build_all_poas.py` fills in the shared keys
from the top level (plain text — the script handles escaping):

```json
{
  "trust_name": "Doe Family Trust",
  "trustees": ["Jane Doe", "John Doe"],
  "trust_date": "2015-03-03",
  "address": "123 Sample Street, Hingham, MA 02043",
  "ps_date": "2026-09-01",
  "signing_date": "2026-09-16",
  "buyers": "William F. Buyer and Jessica Buyer",
  "price": "500,000.00",
  "county": "Plymouth",
  "include_incapacity": true
}
```

| Key | Value |
|-----|-------|
| `trust_name` | Trust name as in the P&S (with "The" only if the P&S has it) |
| `trustees` | Every trustee confirmed in Step 3, in signing order |
| `trust_date` | Declaration of Trust date as `YYYY-MM-DD` (formatted "March 3rd", superscript suffix). Not needed when `include_trust_date` is `false` |
| `include_trust_date` | `true` (default) includes ", u/d/t dated [Trust Date]" in the opening; `false` — the user's Step 3 choice — leaves the whole phrase out. With `true` and no `trust_date`, the builder stops and asks for one |
| `trust_date_source` | Where the date came from (e.g., `P&S SELLER entry`, `Trustee's Certificate.pdf, p. 1`, `typed by the user`) — printed on the `TRUST DATE:` line for the Step 5 report |
| `address`, `ps_date`, `signing_date`, `buyers`, `price`, `county` | As for the entity path (Step 4D-1) |
| `state` | Notary state from Step 3 (default `Massachusetts` when left out) |
| `appointments` | Only for separate notary appointments (Step 3) — signers by name |
| `pronoun` | **One trustee only:** `he`, `she`, `they`, or `blank` (prints a `_____` blank for the notary to fill in), the user's Step 3 answer (required; never guessed). Leave it out for two or more trustees. |
| `include_incapacity` | `true` (default) keeps the "THIS Power of Attorney shall not be affected by my/our subsequent disability or incapacity." paragraph |

#### 4E-2: Build and check it

`build_all_poas.py` (Step 4) runs the builder, the keep-together check, the agreement check, and the
validator for this POA. Report the script messages (`TRUST NAME FONT: …`, `TRUST DATE: …`,
`KEEP-TOGETHER: …`, `AGREEMENT: …`).


### Step 4F — Probate-Estate-Seller POA (Personal Representative(s) or other fiduciary)

`build_all_poas.py` runs `build_estate_poa.py` for it.

#### 4F-1: Write the values file

These are the keys of an **`estate` principal** in `poa_values.json` (Step 4) — one per estate.
The example shows everything the builder receives; `build_all_poas.py` fills in the shared keys
from the top level:

```json
{
  "signers": ["Ellen K. Example"],
  "capacity": "Personal Representative",
  "estate_name": "Estate of Edith A. Sample",
  "probate_county": "Worcester",
  "docket": "WO24P0000EA",
  "address": "80 Sample Road, Millbury, MA 01527",
  "ps_date": "2026-02-19",
  "signing_date": "2026-05-19",
  "buyers": "Morgan Buyer",
  "price": "530,000.00",
  "county": "Worcester",
  "pronoun": "she",
  "include_incapacity": true
}
```

| Key | Value |
|-----|-------|
| `signers` | Every fiduciary confirmed in Step 3, in signing order |
| `capacity` | Singular capacity as in the P&S (pluralized automatically for 2+ signers) |
| `estate_name` | As in the P&S ("Estate of …", or "The Estate of …" only if the P&S has it) |
| `probate_county`, `docket` | Court county (no "County") and docket number |
| `address`, `ps_date`, `signing_date`, `buyers`, `price`, `county` | As for the entity path (Step 4D-1) |
| `state` | Notary state from Step 3 (default `Massachusetts` when left out) |
| `appointments` | Only for separate notary appointments (Step 3) — signers by name |
| `pronoun` | **One signer only:** `he`, `she`, `they`, or `blank` (prints a `_____` blank for the notary to fill in), the user's Step 3 answer (never guessed) |
| `include_incapacity` | `true` (default) |
| `preposition` | Leave it out — the builder applies the rule automatically: **"for"** when the capacity is "Public Administrator", **"of"** for every other capacity |

#### 4F-2: Build and check it

`build_all_poas.py` (Step 4) runs the builder, the keep-together check, the agreement check, and the
validator for this POA. Report the script messages (`ESTATE NAME FONT`, `DOCKET`, `SALE
AUTHORITY` if printed, `KEEP-TOGETHER`, `AGREEMENT`).

### Signature page and checks (all seller types)

`build_all_poas.py` runs `keep_together.py` (keeps the EXECUTED line, signature lines, and notary
block together on one page; in separate-page mode, or when the section does not fit, starts it on
a captioned signature page and adds a "[Remainder of page intentionally left blank…]" notice
when 1.5" or more is left blank) and `check_agreement.py` (I/We, my/our, Trustee/Trustees,
person/persons, he/she/they, the "the" rules, no `{{…}}` left) on every POA. Never add a page
break, caption, or notice by hand. Details: `references/document-format.md`.


## Step 5: Save and Verify

**Output filenames** (set by `build_all_poas.py`):
```
Power of Attorney. Seller. [Inline Property Address].docx                      (one POA)
Power of Attorney. Seller. [Inline Property Address]. [Principal's Name].docx   (two or more POAs)
```
Examples: `Power of Attorney. Seller. 120 Sample Street, Hingham, MA 02043.docx`;
`Power of Attorney. Seller. 123 Sample Street, Hingham, MA 02043. John Smith.docx` and
`Power of Attorney. Seller. 123 Sample Street, Hingham, MA 02043. Smith Family Trust.docx`.

The **principal's name** is the individual sellers' names ("John Smith", "John Smith and Jane
Smith", "A, B, and C"), or the trust, entity, or estate name, each as it reads in the POA (so
the vesting deed's spelling when Step 2D uses it). Characters Windows
does not allow in a file name (`\ / : * ? " < > |`) become "-", and a trailing period is dropped
(so "…, Inc." ends "…, Inc.docx").

`build_all_poas.py` saves the files to the folder of the input P&S (never overwriting a file).

**Verify** by running pandoc to extract the text and confirming:
```bash
pandoc "[output path]" -t plain --wrap=none
```
Do this for **each** POA. Check that all fields updated correctly — seller name(s) (or entity name, state, entity type, and
every entity signer in the signature lines and notary block; or trust name, the u/d/t date (or that it was
left out, as the user chose), and every trustee, with Trustee/Trustees and my/our matching the number of trustees; or estate name,
docket, capacity, and every fiduciary), address, dates
(the EXECUTED line and the notary line both read the same date — "[Month] [Day][suffix], [Year]",
or the partly blank form the user chose), county,
buyers, price, the attorney in fact (the firm setting), and the notary block (bold signer names, he/she/they, capacity
clause) — then send the file(s) to the user. With two or more POAs, list them with the
principal each one is for, and say who signs more than one. Also tell the user the notary location(s) used (state, county or "blank", date or "blank") —
with separate appointments, each appointment and its signers (with several POAs, each signer's
POA, from the `NOTARY APPOINTMENTS (all POAs):` lines), as the script prints them — any
`NOTARY:` reminder (out-of-state wording, California, a corrected county), the signature page
placement used,
for each trust, the `TRUST DATE:` line (the u/d/t date and its source, or that it was left out)
and the trustees used with their sources (say so when the Step 3 trustee question was skipped
because the documents agreed, and name any rider, addendum, or extension that spells a trustee
differently, as a likely typo),
the keep-together result (caption and notice, if any), and the agreement
results and, for an entity, trustee, or estate seller, the name font size used in the signature lines.
Report the **name sources** (Step 2D): the vesting deed used (Book/Page and file name), each name
taken from it and what the P&S said, any capitalization-only change, any seller not found on the
deed, and any grantee who is not selling — or that no vesting deed was found and where you looked.
Report the `P&S DATE:` line — what the opening paragraph says ("September 6th, 2023", "September
_____, 2023", or "left out") and, for a blank date line, which option the user chose.
Report the `SIGNING DATE:` line(s) — what the EXECUTED line and the notary date read (e.g.,
"November _____, 2026 — month and year; day blank") — and, for an exact date after the closing
date, say so.

**Disclaimer to keep in mind.** Each POA is a draft for the responsible attorney to review before
it is signed. Do not describe it as final or as legal advice.

---

## Scripts (in `[SKILL_BASE_DIR]/scripts/`)

| Script | Role |
|---|---|
| `doctor.py` | Checks the requirements and the firm settings |
| `firm_settings.py` | Reads the firm settings; builds the notary location options; matches a typed firm location |
| `build_all_poas.py` | Builds and checks every POA for one P&S (Step 4) — the only builder you run |
| `build_individual_poa.py`, `build_entity_poa.py`, `build_trustee_poa.py`, `build_estate_poa.py` | One builder per seller type (run by `build_all_poas.py`) |
| `notary_block.py` | The notary acknowledgment; `--date-options` and `--pronoun-options` for Step 3 |
| `keep_together.py` | Signature page layout (uses LibreOffice) |
| `check_agreement.py` | Number agreement and leftover-placeholder check |

They need Python 3.9+ and nothing outside the standard library (Pillow is optional). They write
only to a temporary folder and to the output folder you give `build_all_poas.py`, and they send
nothing over the network.
