#!/usr/bin/env python3
"""run_tests.py -- regression suite for the seller-poa scripts, on fictional data only.

    python tests/run_tests.py                 # run every scenario, compare with tests/expected/
    python tests/run_tests.py --update        # rewrite tests/expected/ from the current scripts
    python tests/run_tests.py -k trust        # only scenarios whose name contains "trust"

Each scenario in tests/scenarios/ is a poa_values.json (the file the skill writes in Step 4) plus
a "_test" entry: {"mode": "same"|"separate", "expect": "OK"|"STOPPED"}. The runner builds it with
build_all_poas.py into a temporary folder, converts every POA to plain text with pandoc, and
compares the text with tests/expected/<scenario>/. Error scenarios must stop with nothing saved.
Every POA built is also rendered and must have no blank page ("max_pages" in "_test" caps the
page count). It then runs the firm settings file checks in tests/test_firm_settings.py ("settings_*").

Requirements: Python 3, pandoc. LibreOffice is optional here: without it keep_together.py cannot
render the page layout, so the page-break caption and notice differ. Expected files are
generated WITH LibreOffice, so scenarios are compared only when LibreOffice is present (the
runner says so and runs the error and no-overwrite checks either way).

Every name, address, and number in tests/ is fictional. Never add real client data here.
"""
import argparse, json, os, re, shutil, subprocess, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SKILL = os.path.join(ROOT, "skills", "seller-poa")
TEST_ATTORNEY = "Alex Q. Attorney, Esq."


def have(cmd):
    return shutil.which(cmd) is not None


def soffice():
    """LibreOffice as the builders find it (keep_together.find_soffice): on the PATH, or in its
    usual Windows / macOS install folder, which a Windows install does not put on the PATH."""
    sys.path.insert(0, os.path.join(SKILL, "scripts"))
    sys.dont_write_bytecode = True
    from keep_together import find_soffice
    return bool(find_soffice())


def run_scenario(path, scripts, assets, out_dir):
    v = json.load(open(path, encoding="utf-8"))
    meta = v.pop("_test", {})
    v.setdefault("attorney_in_fact", TEST_ATTORNEY)
    work = tempfile.mkdtemp(prefix="poa_test_")
    vals = os.path.join(work, "poa_values.json")
    json.dump(v, open(vals, "w", encoding="utf-8"), indent=2)
    cmd = [sys.executable, os.path.join(scripts, "build_all_poas.py"), assets, vals, out_dir,
           "--mode", meta.get("mode", "same")]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                       cwd=work, env=dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1"))
    shutil.rmtree(work, ignore_errors=True)
    return meta, r


def final_line(r):
    """The 'ALL POAS: ...' line (build_all_poas.py prints a stop on stderr)."""
    lines = [l for l in (r.stdout + "\n" + r.stderr).splitlines() if l.startswith("ALL POAS:")]
    return lines[-1] if lines else (r.stderr.strip().splitlines() or [""])[-1]


def normalize(text):
    """Compare words, not pandoc's table drawing: pandoc versions draw plain-text tables with
    different rules and column widths (2.9 vs 3.x), so table borders, cell padding and blank lines
    are dropped and runs of spaces collapsed. Every word, in order, is still compared -- including
    the signature-page caption and the page-break notice, which depend on the layout."""
    out = []
    for line in text.replace("|", " ").splitlines():
        if re.fullmatch(r"[\s\-+=:]*", line):
            continue
        out.append(" ".join(line.split()))
    return "\n".join(out)


def page_texts(p):
    """The text of each rendered page (LibreOffice + the skill's own PDF reader)."""
    from keep_together import find_soffice, render
    return render(p, find_soffice())[0]


def page_problem(path, meta):
    """A blank page anywhere, or more pages than the scenario's max_pages."""
    pages = page_texts(path)
    blank = [i + 1 for i, t in enumerate(pages) if not t.strip()]
    if blank:
        return f"blank page(s) {blank} of {len(pages)} in {os.path.basename(path)}"
    if meta.get("max_pages") and len(pages) > meta["max_pages"]:
        return f"{len(pages)} pages, expected at most {meta['max_pages']} in {os.path.basename(path)}"
    return None


def docx_text(p):
    return subprocess.run(["pandoc", p, "-t", "plain", "--wrap=none"], capture_output=True,
                          text=True, encoding="utf-8", check=True).stdout


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--update", action="store_true")
    ap.add_argument("-k", default="")
    ap.add_argument("--scripts", default=os.path.join(SKILL, "scripts"))
    ap.add_argument("--assets", default=os.path.join(SKILL, "assets"))
    ap.add_argument("--expected", default=os.path.join(HERE, "expected"))
    a = ap.parse_args()
    a.scripts, a.assets, a.expected = (os.path.abspath(x) for x in (a.scripts, a.assets, a.expected))

    if not have("pandoc"):
        print("pandoc is required for the tests (https://pandoc.org/installing.html)")
        return 2
    compare = soffice()
    if not compare and not a.update:
        print("NOTE: LibreOffice not found -- text comparison skipped (layout differs without "
              "it); running the stop/no-overwrite checks only.\n")
    if a.update and not compare:
        print("Refusing --update without LibreOffice: the expected files must include the layout.")
        return 2

    names = sorted(f[:-5] for f in os.listdir(os.path.join(HERE, "scenarios"))
                   if f.endswith(".json") and a.k in f)
    fails = 0
    for name in names:
        out = tempfile.mkdtemp(prefix="poa_out_")
        meta, r = run_scenario(os.path.join(HERE, "scenarios", name + ".json"),
                               a.scripts, a.assets, out)
        last = final_line(r)
        built = sorted(f for f in os.listdir(out) if f.endswith(".docx"))
        problem = None
        if meta.get("expect", "OK") == "STOPPED":
            if not last.startswith("ALL POAS: STOPPED") or built:
                problem = f"expected STOPPED with nothing saved; got '{last}' and {len(built)} file(s)"
            elif meta.get("reason") and meta["reason"] not in r.stdout + r.stderr:
                problem = f"stopped, but not for the expected reason ({meta['reason']!r}): {last[:200]}"
        elif not last.startswith("ALL POAS: OK"):
            problem = f"build failed: {last[:300]}\n      " + (r.stdout + r.stderr)[-600:].replace("\n", "\n      ")
        elif compare:
            exp_dir = os.path.join(a.expected, name)
            if a.update:
                shutil.rmtree(exp_dir, ignore_errors=True)
                os.makedirs(exp_dir)
                for f in built:
                    open(os.path.join(exp_dir, f[:-5] + ".txt"), "w", encoding="utf-8",
                         newline="\n").write(docx_text(os.path.join(out, f)))
            else:
                want = sorted(f for f in os.listdir(exp_dir)) if os.path.isdir(exp_dir) else []
                if [f[:-5] + ".txt" for f in built] != want:
                    problem = f"file names differ: built {built}, expected {want}"
                else:
                    for f in built:
                        got = docx_text(os.path.join(out, f))
                        exp = open(os.path.join(exp_dir, f[:-5] + ".txt"), encoding="utf-8").read()
                        if normalize(got) != normalize(exp):
                            problem = f"text differs in {f}"
                            break
            # no blank page (e.g. empty paragraphs spilling past a full page 1), page limit
            for f in built:
                problem = problem or page_problem(os.path.join(out, f), meta)
            # determinism + no-overwrite: a second run into the same folder must stop
            if not problem:
                _, r2 = run_scenario(os.path.join(HERE, "scenarios", name + ".json"),
                                     a.scripts, a.assets, out)
                l2 = final_line(r2)
                if not l2.startswith("ALL POAS: STOPPED"):
                    problem = "second run into the same folder did not refuse to overwrite"
        shutil.rmtree(out, ignore_errors=True)
        print(f"{'FAIL' if problem else 'ok  '}  {name}" + (f"\n      {problem}" if problem else ""))
        fails += bool(problem)
    # the firm settings file (tests/test_firm_settings.py): read, save, refuse, precedence
    sys.path.insert(0, HERE)
    import test_firm_settings
    total = len(names)
    for name, problem in test_firm_settings.run_checks():
        name = "settings_" + name
        if a.k not in name:
            continue
        total += 1
        print(f"{'FAIL' if problem else 'ok  '}  {name}" + (f"\n      {problem}" if problem else ""))
        fails += bool(problem)
    print(f"\n{total - fails}/{total} passed" + (" (expected files updated)" if a.update else ""))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
