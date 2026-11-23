# Notary appointments — signers who will not all sign together

Reference for `SKILL.md`, Step 3. Read it when the user answers **No** to the same-appointment
question.

### Notary appointments (two or more signers who will NOT all sign together)

Sometimes the sellers (or trustees, fiduciaries, or entity
signers) can't or won't attend the same notary appointment. Each appointment then gets its own
**counterpart signature page**: its own EXECUTED line and date, only its own signers' signature
lines, and its own notary block with that appointment's state, county, date, and he/she/they.
The goal is that the intake feels easy and natural — one short question per signer.

**Appointment A** is Signer 1's appointment: the date and location from questions 1 and 2.
Then go through the remaining signers **in order** (Signer 2, Signer 3, …), one
`AskUserQuestion` call per signer:

- **"Where will [Signer k] sign?"** Options: one option per appointment already set up, labeled
  by who is going and where/when — e.g., "With Peter F. Example — Plymouth County, MA, October
  30th" — then "At a different appointment". (Two signers: just "With [Signer 1] — …" / "At a
  different appointment"; but if the user answered No to question 4, two signers means two
  appointments, so skip this question and go straight to the new appointment.) The question tool
  allows four options, so with four or more appointments already set up, offer the three most
  recent and tell the user they can type "with [name]" under "Other".
- **A different appointment** → in the **same follow-up call**, ask that appointment's **date**
  (the same options as question 1, from `--date-options`; a booked date typed under "Other") and **location** (the same options as question 2, from `firm_settings.py --location-options`), and the next appointment is set up.

When every signer has an appointment:

- **Pronouns** — for each appointment with **exactly one** signer, ask "He" / "She" / "They" /
  "Leave blank (_____)" for that signer, in the `--pronoun-options` order for that signer's name (batch them in one call: "[Name] will sign alone in
  [place] — which pronoun should the notary block use?"). Never guess from a name. Appointments
  with two or more signers use "they".
- If every signer ended up in the **same** appointment after all, treat it as one appointment
  (no counterparts).
- Before building, show the user the appointments in one short list (e.g., "A — Plymouth County,
  MA, Oct 30th: Peter F. Example and Tom Sample; B — Norfolk County, MA, Oct 31st: Jessica
  H. Example"), and build without waiting unless something is unclear.

Write the appointments into the values file as `appointments` (Step 4 — every path uses the same
format):

```json
"appointments": [
  {"signers": ["Peter F. Example", "Tom Sample"], "state": "Massachusetts",
   "county": "Plymouth", "signing_date": "2026-10-30"},
  {"signers": ["Jessica H. Example"], "state": "Massachusetts", "county": "Norfolk",
   "signing_date": "2026-10-31", "pronoun": "she"}
]
```

`signers` are names exactly as in the POA's signer list (for an entity, the `name` of each
signer); every signer must be in exactly one appointment (the scripts check this). `county`
`""` = blank line; `signing_date` as in Step 3 question 1 (`"2026-10-30"`, `"2026-10"`, `"2026"`, or
`""`); `pronoun` only for a one-signer appointment.
With one appointment, leave `appointments` out and use the top-level `signing_date`, state, and
`county` as before.

### Notary appointments across several POAs (several principals)

When the P&S has two or more principals (one POA each — Step 2) and the
user answers **No** to the same-appointment question, set up the appointments **once, for all the
POAs together**, with the same walk-through as above:

- **Signer order** — POA order (the order the P&S names the principals), each POA's signers in
  order; a person who signs two POAs comes up once, at the first of them, and signs **both POAs
  at that one appointment** (if the user says otherwise, stop and ask how to handle it — not
  built).
- **Appointment A** is the first signer's: questions 1 and 2. Then, for each remaining signer,
  one `AskUserQuestion` call: **"Where will [Name] ([principal] POA) sign?"** with one option per
  appointment already set up, labeled by who is going and where/when — e.g., "With Henry J.
  Sample — Norfolk County, MA, December 2nd" — then "At a different appointment" (whose date and
  location are asked in the same call, as above). With just two signers in all, the user's **No**
  already means two appointments: skip "Where will…" and ask the second appointment's date and
  location.
- **Pronouns** — then ask the pronoun for each person alone in a POA's notary block (see
  "Several principals" above), batched in one call.
- **Show the plan** before building, one line per appointment with each signer's POA, e.g.
  "A — Norfolk County, MA, Dec 2nd: Henry J. Sample (Trust POA); B — Plymouth County, MA,
  Nov 30th: Grace Sample (her POA)". Build without waiting unless something is unclear.

Write them **once**, at the top level of `poa_values.json` (Step 4), with the pronouns:

```json
"appointments": [
  {"signers": ["Henry J. Sample"], "state": "Massachusetts", "county": "Norfolk",
   "signing_date": "2026-12-02"},
  {"signers": ["Grace Sample"], "state": "Massachusetts", "county": "Plymouth",
   "signing_date": "2026-11"}
],
"pronouns": {"Henry J. Sample": "he", "Grace Sample": "she"}
```

`build_all_poas.py` gives each POA only the appointments its own signers attend:
- **All of a POA's signers at one appointment** → that POA is built with that appointment's date,
  state, and county — no counterparts (e.g., Grace's POA and the Trust POA each get their own
  place and date).
- **A POA's signers split across appointments** (e.g., co-trustees, one with Grace, one alone) →
  that POA gets counterpart signature pages, exactly as under "Notary appointments" above.
- It stops, saving nothing, if a signer is missing from every appointment or in two, a name
  matches no POA's signer, a county is wrong, or a pronoun is missing for someone alone in a
  POA's notary block. It prints `NOTARY APPOINTMENTS (all POAs):` with one line per appointment
  (A, B, …) naming each signer's POA.

**What changes in the POA with several appointments:**
- A **counterparts paragraph** is added to the POA text, after the other paragraphs: "THIS Power
  of Attorney may be executed in counterparts, each of which shall be deemed an original, and all
  of which together shall constitute one and the same instrument."
- Each appointment's section goes on its **own page**, captioned "**Counterpart Signature Page 1
  of 2 to Seller's Limited Power of Attorney for the sale of [address].**" (bold, flush left) —
  whatever the signature page placement setting, since each signer group signs only its own page.
- The page-break notice (when there is 1.5" or more of blank space) reads ***[Remainder of page
  intentionally left blank; signature pages follow]***.
- Pages are in signing order: Signer 1's appointment first.

---
