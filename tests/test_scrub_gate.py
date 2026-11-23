#!/usr/bin/env python3
"""test_scrub_gate.py -- prove the structural gate (scripts/scrub_check.py) blocks what it claims to.

    python tests/test_scrub_gate.py

Copies the repository to a temporary folder (outside the repo, so no probe file can ever land in
the tree), checks the clean copy PASSES, then plants one violation at a time and checks the gate
FAILS on each. Every planted value is fictional.
"""
import io, os, shutil, subprocess, sys, tempfile, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = "skills/seller-poa/assets/POA. Seller. Form 2.docx"


def copy_repo():
    dst = tempfile.mkdtemp(prefix="gate_test_")
    repo = os.path.join(dst, "ma-seller-poa")
    shutil.copytree(ROOT, repo, ignore=shutil.ignore_patterns(".git", "__pycache__"))
    return repo


def gate(repo):
    r = subprocess.run([sys.executable, os.path.join(repo, "scripts", "scrub_check.py"),
                        "--structural-only"], capture_output=True, text=True)
    return r.returncode, r.stdout


def write(repo, rel, data, mode="w"):
    p = os.path.join(repo, rel)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, mode) as f:
        f.write(data)


def retemplate(repo, part, fn):
    """Rewrite one part of the Form 2 template with fn(xml)."""
    p = os.path.join(repo, TEMPLATE)
    src = zipfile.ZipFile(p)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for n in src.namelist():
            data = src.read(n)
            out.writestr(n, fn(data.decode("utf-8")).encode("utf-8") if n == part else data)
    open(p, "wb").write(buf.getvalue())


PLANTS = [
    ("a PDF", lambda r: write(r, "docs/P&S. Signed. 1 Example Way.pdf", "%PDF-1.4 fictional")),
    ("a generated POA (.docx)", lambda r: write(r, "Power of Attorney. Seller. 1 Example Way.docx", "x")),
    ("a scan image", lambda r: write(r, "skills/seller-poa/assets/scan.png", "x")),
    ("a run values file", lambda r: write(r, "poa_values.json", "{}")),
    ("a name map", lambda r: write(r, "seller-poa-namemap.md", "x")),
    ("a Windows user path", lambda r: write(r, "notes.md", "see C:" + "\\Users\\someone\\Clients\\x\n")),
    ("a POSIX home path", lambda r: write(r, "notes.md", "see /home/" + "someone/clients/x\n")),
    ("an email address", lambda r: write(r, "notes.md", "contact jane" + "@example.com\n")),
    ("an SSN-shaped number", lambda r: write(r, "notes.md", "SSN 000-" + "12-3456\n")),
    ("an API key", lambda r: write(r, "notes.md", "key sk-" + "ant-abcdefghijklmnop\n")),
    ("a changed template (one byte)", lambda r: open(os.path.join(r, TEMPLATE), "ab").write(b"\0")),
    ("template author metadata", lambda r: retemplate(r, "docProps/core.xml", lambda x: x.replace(
        "<dc:creator></dc:creator>", "<dc:creator>Pat Example</dc:creator>"))),
    ("template tracked change", lambda r: retemplate(r, "word/document.xml", lambda x: x.replace(
        "<w:body>", '<w:body><w:p><w:ins w:id="9" w:author="X"><w:r><w:t>x</w:t></w:r></w:ins></w:p>', 1))),
    ("template external link", lambda r: retemplate(r, "word/_rels/document.xml.rels", lambda x: x.replace(
        "</Relationships>", '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/attachedTemplate" Target="file:///x.dotm" '
        'TargetMode="External"/></Relationships>'))),
]


def main():
    fails = 0
    repo = copy_repo()
    code, out = gate(repo)
    print(f"{'ok  ' if code == 0 else 'FAIL'}  clean copy passes")
    fails += code != 0
    if code != 0:
        print(out)
    shutil.rmtree(os.path.dirname(repo))
    for name, plant in PLANTS:
        repo = copy_repo()
        plant(repo)
        code, out = gate(repo)
        print(f"{'ok  ' if code == 1 else 'FAIL'}  blocks {name}")
        fails += code != 1
        shutil.rmtree(os.path.dirname(repo))
    print(f"\n{len(PLANTS) + 1 - fails}/{len(PLANTS) + 1} passed")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
