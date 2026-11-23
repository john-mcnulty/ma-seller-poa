#!/usr/bin/env python3
"""
scrub_check.py -- pre-commit gate against committing client data.

    python scripts/scrub_check.py                 # check the working tree
    python scripts/scrub_check.py --staged        # check only staged files
    python scripts/scrub_check.py --install-hook  # wire it to git pre-commit
    python scripts/scrub_check.py --hash FILE     # print the SHA-256 to allowlist an inspected file

Exit 0 = clean, 1 = something must not be committed.

WHY THIS FILE CANNOT CONTAIN THE THING IT CHECKS FOR
----------------------------------------------------
The obvious gate is a list of real client names to grep for. That list is exactly the
confidential data the gate exists to protect, and this file is public -- so shipping it here
would publish the names in the act of trying not to. Hashing them is barely better: surnames and
street names are low-entropy and a hash list is brute-forced against a dictionary in seconds.
git history survives a private->public flip, so a mistake here is not fixable by a later commit.

So the gate has two halves:

  1. STRUCTURAL checks (this file, always run). They look for client data by SHAPE, not by
     identity: signed P&S agreements, deeds and other PDFs, generated POAs and other Word files,
     images, values files from a real run, absolute paths naming a user's home directory, stray
     credentials. No name list required, so they are safe to publish and work in a fork.

  2. IDENTITY checks (delegated). If the private verifier is present on this machine -- pointed
     at by SELLER_POA_SCRUB_TOOLS, or found at its conventional location beside the repo -- it is
     executed and its result folded in. It holds the name rules and never enters this repository.

A missing private verifier is a hard failure for the repository OWNER (a machine where the
private tools folder exists, or the variable is set) and only a note for a fork. A gate that
silently downgrades itself when its strongest check is unavailable is the
missing-information-read-as-a-negative-answer failure this project keeps guarding against.

THE TWO WORD TEMPLATES
----------------------
This is a document-generation skill, so it ships two .docx templates. Word files stay BLOCKED by
default; the templates pass only because each one's exact path is listed in TEMPLATE_ALLOWLIST
with the SHA-256 of the exact bytes that were inspected. Any change to a template -- one byte --
fails the gate until it is inspected again and its new hash recorded (--hash). An allowlisted
file must ALSO pass docx_structure(): only the expected package parts, no comments, no tracked
changes, no hidden text, no author / editor / company metadata, no external relationships, no
media, no custom XML, no revision IDs.
"""
import argparse
import fnmatch
import hashlib
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent

# --------------------------------------------------------------- structural

FORBIDDEN_GLOBS = [
    ("*.pdf",  "a PDF (a signed P&S, a deed, a probate paper, a rendered POA)"),
    ("*.jpg",  "an image (a scan or screenshot of a client document)"),
    ("*.jpeg", "an image (a scan or screenshot of a client document)"),
    ("*.png",  "an image (a scan or screenshot of a client document)"),
    ("*.heic", "an image (a phone photo of a client document)"),
    ("*.tif",  "an image (a scan)"),
    ("*.tiff", "an image (a scan)"),
    ("*.docx", "a Word file (a P&S, a generated POA) -- only the inspected templates may ship"),
    ("*.doc",  "a Word file"),
    ("*.rtf",  "a Word file"),
    ("*.odt",  "a Word file"),
    ("*.msg",  "an email"),
    ("*.eml",  "an email"),
    ("poa_values*.json", "a values file from a real run (seller, buyer, address, price)"),
    ("ps_accepted*.txt", "text extracted from a real P&S"),
    ("ps_redline*.txt",  "text extracted from a real P&S"),
    ("*namemap*", "THE NAME MAP -- this must never be in the repository"),
    ("*apply_scrub*", "a scrub tool carrying real names"),
    ("*verify_scrub*", "a scrub tool carrying real names"),
    ("*scrub_inventory*", "a scrub tool carrying real names"),
    ("*rebuild_templates*", "the private template rebuild tool (reads the firm's originals)"),
    ("*PRIVATE*", "a private file"),
    (".env", "credentials"),
    ("*.pem", "credentials"),
    ("*.key", "credentials"),
]

# Exact repo paths -> SHA-256 of the inspected bytes. Update ONLY after inspecting the file at
# the zip level (python scripts/scrub_check.py --hash PATH prints the value).
TEMPLATE_ALLOWLIST = {
    "skills/seller-poa/assets/POA. Seller. Form 2.docx":
        "d060e0f48985b0a857f273e8d136c7412c27c41f86a3ff57e9587c50f99aa9da",
    "skills/seller-poa/assets/POA. Seller. Form 4.docx":
        "894d47a05a9968c9e977cc35e368f6ad09a53fecad12256bf082ae0766a204db",
}
TEMPLATE_PARTS = {
    "[Content_Types].xml", "_rels/.rels", "docProps/core.xml", "docProps/app.xml",
    "word/document.xml", "word/_rels/document.xml.rels", "word/styles.xml",
    "word/settings.xml", "word/fontTable.xml", "word/theme/theme1.xml",
}

CONTENT_PATTERNS = [
    (re.compile(r"[A-Za-z]:\\+Users\\+(?!<)[A-Za-z0-9._-]+", re.I),
     "an absolute Windows user path (names the operator and their layout)"),
    (re.compile(r"/(?:home|Users)/(?!<)[a-z0-9._-]{2,}/", re.I),
     "an absolute POSIX home path"),
    (re.compile(r"sk-ant-[A-Za-z0-9_-]{8,}"),
     "what looks like an Anthropic API key"),
    (re.compile(r"\b\d{3}-\d{2}-\d{4}\b"),
     "what looks like a Social Security number"),
    # These two lines carry the allow marker because the pattern text matches itself.
    (re.compile(r"One" r"Drive", re.I),                       # scrub-check: allow
     "a synced-folder path; the private working copy lives in one"),  # scrub-check: allow
]

# The author's own address is published deliberately in the README and the plugin manifest;
# every OTHER address is a finding.
ALLOWED_EMAILS = {"john.mcnulty80@gmail.com"}
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")

SKIP_DIRS = {".git", "__pycache__", ".venv", "venv", "node_modules"}
ALLOW_MARKER = "scrub-check: allow"


def _rel(p: Path) -> str:
    return p.relative_to(REPO).as_posix()


def _iter_files(staged: bool):
    if staged:
        try:
            out = subprocess.run(
                ["git", "diff", "--cached", "--name-only", "--diff-filter=ACMR"],
                cwd=REPO, capture_output=True, text=True, check=True).stdout
        except Exception as e:
            print(f"  ! could not list staged files ({e}); checking the tree")
            yield from _iter_files(False)
            return
        for rel in out.split("\n"):
            rel = rel.strip()
            if rel and (REPO / rel).is_file():
                yield REPO / rel
        return
    for p in sorted(REPO.rglob("*")):
        if p.is_file() and not any(part in SKIP_DIRS for part in p.relative_to(REPO).parts):
            yield p


def _is_binary(p: Path) -> bool:
    try:
        with open(p, "rb") as fh:
            return b"\x00" in fh.read(4096)
    except OSError:
        return True


def sha256(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def docx_structure(p: Path) -> list:
    """Problems with an allowlisted template, inspected part by part."""
    probs = []
    try:
        z = zipfile.ZipFile(p)
    except zipfile.BadZipFile:
        return ["not a valid .docx (zip)"]
    names = set(z.namelist())
    extra, missing = names - TEMPLATE_PARTS, TEMPLATE_PARTS - names
    if extra:
        probs.append("unexpected package parts: " + ", ".join(sorted(extra)))
    if missing:
        probs.append("missing package parts: " + ", ".join(sorted(missing)))
    for n in sorted(names):
        x = z.read(n).decode("utf-8", errors="replace")
        if n.endswith(".rels") and 'TargetMode="External"' in x:
            probs.append(f"{n}: an external relationship (a linked file or template path)")
        if n.startswith("word/"):
            for tag, why in (("<w:ins ", "a tracked insertion"), ("<w:del ", "a tracked deletion"),
                             ("<w:vanish", "hidden text"), ("commentReference", "a comment"),
                             ("<w:attachedTemplate", "an attached template path"),
                             ("<w:mailMerge", "a mail-merge source"), ("w:rsid", "revision IDs"),
                             ("w14:paraId", "paragraph IDs"), ("docId", "a document ID")):
                if tag in x:
                    probs.append(f"{n}: {why}")
        for rx, why in CONTENT_PATTERNS[:3]:
            if rx.search(x):
                probs.append(f"{n}: {why}")
        for m in EMAIL_RE.finditer(x):
            if m.group(0).lower() not in ALLOWED_EMAILS:
                probs.append(f"{n}: an email address {m.group(0)!r}")
    core = z.read("docProps/core.xml").decode("utf-8") if "docProps/core.xml" in names else ""
    for tag in ("dc:creator", "cp:lastModifiedBy", "cp:keywords", "dc:description",
                "cp:category", "dc:subject"):
        m = re.search(rf"<{tag}>([^<]*)</{tag}>", core)
        if m and m.group(1).strip():
            probs.append(f"docProps/core.xml: {tag} is not empty ({m.group(1)!r})")
    app = z.read("docProps/app.xml").decode("utf-8") if "docProps/app.xml" in names else ""
    for tag in ("Company", "Manager"):
        m = re.search(rf"<{tag}>([^<]*)</{tag}>", app)
        if m and m.group(1).strip():
            probs.append(f"docProps/app.xml: {tag} is not empty ({m.group(1)!r})")
    return probs


def structural_check(staged: bool) -> list:
    findings = []
    files = list(_iter_files(staged))

    for p in files:
        rel = _rel(p)
        if rel in TEMPLATE_ALLOWLIST:
            got = sha256(p)
            if got != TEMPLATE_ALLOWLIST[rel]:
                findings.append(f"{rel}: template changed since it was inspected (sha256 {got[:12]}… "
                                "not allowlisted) -- inspect every part, then update TEMPLATE_ALLOWLIST")
            for prob in docx_structure(p):
                findings.append(f"{rel}: {prob}")
            continue
        name = p.name.lower()
        for pat, why in FORBIDDEN_GLOBS:
            if fnmatch.fnmatchcase(name, pat.lower()):
                findings.append(f"{rel}: {why} -- must not be committed")
                break

    for p in files:
        rel = _rel(p)
        if rel in TEMPLATE_ALLOWLIST or _is_binary(p):
            continue
        try:
            text = p.read_text(encoding="utf-8")
        except Exception:
            continue
        for i, line in enumerate(text.splitlines(), 1):
            if ALLOW_MARKER in line:
                continue
            for rx, why in CONTENT_PATTERNS:
                m = rx.search(line)
                if m:
                    findings.append(f"{rel}:{i}: {why} -- {m.group(0)!r}")
            for m in EMAIL_RE.finditer(line):
                if m.group(0).lower() not in ALLOWED_EMAILS:
                    findings.append(f"{rel}:{i}: an email address -- {m.group(0)!r}")
    return findings


def history_check() -> list:
    """Every commit reachable from HEAD, not just the tip: a client file added in one commit and
    removed in the next is still in history forever. File NAMES are checked against
    FORBIDDEN_GLOBS (the allowlisted template paths excepted), and every version of an
    allowlisted template is opened and checked with docx_structure()."""
    import tempfile
    findings, seen_blobs = [], set()
    try:
        shas = subprocess.run(["git", "rev-list", "HEAD"], cwd=REPO, capture_output=True,
                              text=True, check=True).stdout.split()
    except Exception as e:
        return [f"could not list commits ({e})"]
    for sha in shas:
        tree = subprocess.run(["git", "ls-tree", "-r", sha], cwd=REPO, capture_output=True,
                              text=True, check=True).stdout
        for line in tree.splitlines():
            meta, rel = line.split("\t", 1)
            blob = meta.split()[2]
            if rel in TEMPLATE_ALLOWLIST:
                if blob in seen_blobs:
                    continue
                seen_blobs.add(blob)
                data = subprocess.run(["git", "cat-file", "blob", blob], cwd=REPO,
                                      capture_output=True, check=True).stdout
                with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as t:
                    t.write(data)
                for prob in docx_structure(Path(t.name)):
                    findings.append(f"commit {sha[:10]} {rel}: {prob}")
                os.unlink(t.name)
                continue
            name = rel.rsplit("/", 1)[-1].lower()
            for pat, why in FORBIDDEN_GLOBS:
                if fnmatch.fnmatchcase(name, pat.lower()):
                    findings.append(f"commit {sha[:10]} {rel}: {why}")
                    break
    return sorted(set(findings))


# ----------------------------------------------------------------- identity

def _private_dir() -> Path:
    """Convention: the private tools sit BESIDE the repo, so git inside it cannot see them."""
    return REPO.parent / f"{REPO.name}-PRIVATE-scrub-tools"


def _private_verifier() -> Path:
    env = os.environ.get("SELLER_POA_SCRUB_TOOLS")
    if env:
        p = Path(env)
        return p if p.suffix == ".py" else p / "verify_scrub.py"
    return _private_dir() / "verify_scrub.py"


def _is_owner_machine() -> bool:
    return bool(os.environ.get("SELLER_POA_SCRUB_TOOLS")) or _private_dir().is_dir()


def identity_check(staged: bool) -> tuple:
    v = _private_verifier()
    if not v.is_file():
        if _is_owner_machine():
            return ("fail", f"the private verifier is expected at {v} but is missing. On this "
                            "machine that is a broken gate, not an absent one -- refusing to pass.")
        return ("skip", "no private verifier on this machine (expected for a fork). Structural "
                        "checks only; the name-identity check did not run.")
    cmd = [sys.executable, str(v), "--repo", str(REPO)] + (["--staged"] if staged else [])
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    except Exception as e:
        return ("fail", f"the private verifier could not be executed: {e}")
    if r.returncode == 0:
        return ("ok", "private verifier reports ALL CHECKS PASSED")
    tail = (r.stdout or "").strip().splitlines()
    return ("fail", "private verifier FAILED:\n    " + "\n    ".join(tail[-25:] or ["(no output)"]))


# --------------------------------------------------------------------- hook

HOOK = """#!/bin/sh
# Installed by scripts/scrub_check.py --install-hook
PY=python3; command -v python3 >/dev/null 2>&1 && python3 -c "" 2>/dev/null || PY=python
exec "$PY" "$(git rev-parse --show-toplevel)/scripts/scrub_check.py" --staged
"""


def install_hook() -> int:
    hooks = REPO / ".git" / "hooks"
    if not hooks.is_dir():
        print(f"No git hooks directory at {hooks} -- run `git init` first.")
        return 1
    path = hooks / "pre-commit"
    if path.exists():
        existing = path.read_text(encoding="utf-8", errors="replace")
        if "scrub_check.py" in existing:
            print(f"Already installed: {path}")
            return 0
        backup = path.with_suffix(".pre-scrub-check")
        backup.write_text(existing, encoding="utf-8")
        print(f"Existing hook backed up to {backup}")
    path.write_text(HOOK, encoding="utf-8", newline="\n")
    try:
        path.chmod(0o755)
    except Exception:
        pass
    print(f"Installed pre-commit hook: {path}")
    return 0


# --------------------------------------------------------------------- main

def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--staged", action="store_true", help="check only files staged for commit")
    ap.add_argument("--install-hook", action="store_true", help="install the git pre-commit hook")
    ap.add_argument("--hash", metavar="FILE", help="print a file's SHA-256 (for the allowlist)")
    ap.add_argument("--history", action="store_true",
                    help="check every commit reachable from HEAD (file names and template structure)")
    ap.add_argument("--structural-only", action="store_true",
                    help="skip the private identity check (CI, and the gate's own tests)")
    args = ap.parse_args()

    if args.install_hook:
        return install_hook()
    if args.hash:
        p = Path(args.hash).resolve()
        print(sha256(p))
        for prob in docx_structure(p) if p.suffix.lower() == ".docx" else []:
            print(f"  PROBLEM  {prob}")
        return 0

    if args.history:
        hist = history_check()
        for f in hist:
            print(f"  BLOCK  {f}")
        print("history clean -- no client-data-shaped file in any commit" if not hist else
              "BLOCKED -- history carries client-data-shaped files")
        return 1 if hist else 0

    print(f"scrub-check -- {'staged files' if args.staged else 'working tree'}\n")
    findings = structural_check(args.staged)
    print("Structural checks")
    if findings:
        for f in findings:
            print(f"  BLOCK  {f}")
    else:
        print("  clean -- no client documents, run values, home paths, or stray credentials; "
              "templates match their inspected hashes")

    if args.structural_only:
        status, detail = "skip", "skipped (--structural-only)"
    else:
        status, detail = identity_check(args.staged)
    print("\nName-identity check")
    print(f"  {'ok' if status == 'ok' else status.upper()}  {detail}")

    failed = bool(findings) or status == "fail"
    print("\n" + ("BLOCKED -- do not commit. git history survives a private-to-public flip, so "
                  "this must be clean at the FIRST commit." if failed else "OK -- safe to commit."))
    if status == "skip" and not failed:
        print("Note: the name-identity check did not run. A clean structural result does not "
              "prove the change is free of real names -- see SECURITY.md.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
