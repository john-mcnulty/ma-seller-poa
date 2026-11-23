# Security and privacy reporting

## The report this project cares most about

Most projects' worst-case report is a vulnerability. This one's is different:

> **"Your repository contains a real person's name, address, or document reference."**

This skill was built while doing real conveyancing work: it reads signed Purchase and Sale
Agreements, deeds, trust certificates, probate papers, and corporate records. Every example in the
code, the documentation, the tests, and the two Word templates has been replaced with obviously
fictional values, the templates were rebuilt from scratch with their metadata cleared, and a
commit gate runs on every change (`scripts/scrub_check.py`, plus a private name-identity verifier
the maintainer runs before pushing). That process is not assumed to be perfect — which is why this
page exists.

**If you believe you have found a real identifier — a person, an entity, a street address, a
book/page or document number, a docket number, a price tied to a real sale — please report it
privately, not in a public issue.** A public issue would republish it to a bigger audience.

## How to report privately

1. **Preferred:** GitHub → the **Security** tab → **Report a vulnerability** (private
   vulnerability reporting). It is private to the maintainer.
2. If that is unavailable, open a public issue with **no specifics** — "possible identifier in
   `SKILL.md`, details sent privately" — and ask for a contact address.

Please include the file and line, and why you think it is real. You do not need to be certain. A
false alarm costs a look; the other error is permanent, because git history survives a
private-to-public flip.

## Scope

In scope, most to least urgent:

- A real client identifier anywhere in the tree, **in any commit in history**, or inside the
  Word templates (including their document properties).
- A credential, API key, or absolute path naming a real user's home directory.
- A defect in the gate itself — `scripts/scrub_check.py` passing something it claims to block.
  `tests/test_scrub_gate.py` plants each violation and proves the gate fails; a case it does not
  cover is a real finding.
- Anything that makes the skill **write outside the P&S folder** (or the session's working
  folder), overwrite an existing file, or send run data anywhere. It is designed to do none of
  these.
- A POA that names someone other than the configured attorney in fact.

Out of scope:

- **Massachusetts town and county names.** They are needed for the notary block, the county
  checks, and the docket codes, and identify nobody on their own.
- The statutes, court names, and form wording the skill relies on.

## What a fix looks like

For a leaked identifier, expect the maintainer to widen both halves of the private scrub (the
replacement rules and the verifier), add a regression probe, and rewrite the affected history
rather than only the tip. A later commit does not remove anything from a public repository's
history.

## For contributors

CI runs only the structural half of the gate: it looks for client data by *shape* (PDFs, Word
files, images, run values files, home paths, credentials) and needs no list of real names by
design. **A green CI run does not prove your patch is free of real identifiers.** Use obviously
fictional names and addresses (`Jane Example`, `123 Sample Street`) in any example, test, or
issue — never the matter in front of you.
