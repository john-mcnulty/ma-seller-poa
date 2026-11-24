#!/usr/bin/env python3
"""test_firm_settings.py -- checks of the firm settings file (seller-poa-firm-settings.json).

    python tests/test_firm_settings.py

Runs firm_settings.py in a throwaway folder with an EMPTY Claude config folder (so the machine's
own plugin settings never leak in) and no SELLER_POA_SETTINGS: the Cowork case. Also run by
tests/run_tests.py. Every name here is fictional.
"""
import json, os, shutil, subprocess, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(os.path.dirname(HERE), "skills", "seller-poa", "scripts", "firm_settings.py")
NAME = "seller-poa-firm-settings.json"
ATTY = "Alex Q. Attorney, Esq."
LOCS = "Main office = Plymouth County, MA; Boston office = Suffolk County, MA"


class Sandbox:
    """Clients/<client>/ under a temp folder, plus an empty Claude config folder."""

    def __init__(self):
        self.root = tempfile.mkdtemp(prefix="poa_settings_")
        self.cfg = os.path.join(self.root, "claude-config")
        self.clients = os.path.join(self.root, "Clients")
        self.client = os.path.join(self.clients, "Sample, Pat")
        self.sub = os.path.join(self.client, "P&S")
        for d in (self.cfg, self.sub):
            os.makedirs(d)
        self.file = os.path.join(self.clients, NAME)

    def run(self, *args):
        env = {k: v for k, v in os.environ.items() if k != "SELLER_POA_SETTINGS"}
        env.update(CLAUDE_CONFIG_DIR=self.cfg, PYTHONIOENCODING="utf-8")
        r = subprocess.run([sys.executable, SCRIPT, *args], capture_output=True, text=True,
                           encoding="utf-8", env=env)
        try:
            out = json.loads(r.stdout)
        except ValueError:
            out = {"_raw": r.stdout + r.stderr}
        return r.returncode, out

    def plugin_config(self, **opts):
        with open(os.path.join(self.cfg, "settings.json"), "w", encoding="utf-8") as f:
            json.dump({"pluginConfigs": {"ma-seller-poa@ma-seller-poa": {"options": opts}}}, f)

    def close(self):
        shutil.rmtree(self.root, ignore_errors=True)


def check(cond, msg):
    if not cond:
        raise AssertionError(msg)


def t_nothing_anywhere_asks(sb):
    code, out = sb.run("--show", "--attorney-in-fact", "", "--search-from", sb.client)
    check(code == 1 and out.get("errors"), f"expected an error with no settings: {out}")
    check(out["attorney_in_fact"] is None and out["settings_file"] is None, out)


def t_save_then_read(sb):
    code, out = sb.run("--save", sb.file, "--attorney-in-fact", ATTY, "--signing-locations", LOCS,
                       "--signature-page", "separate-page")
    check(code == 0 and out["saved"] and os.path.isfile(sb.file), f"save failed: {out}")
    code, out = sb.run("--show", "--attorney-in-fact", "", "--search-from", sb.client)
    check(code == 0 and out["attorney_in_fact"] == ATTY, f"not read back: {out}")
    check(out["mode"] == "separate", out)
    check([l["label"] for l in out["signing_locations"]] == ["Main office", "Boston office"], out)
    check(os.path.samefile(out["settings_file"], sb.file), out)
    check(out["sources"]["attorney_in_fact"].startswith("firm settings file"), out["sources"])


def t_second_run_offers_offices(sb):
    sb.run("--save", sb.file, "--attorney-in-fact", ATTY, "--signing-locations", LOCS)
    code, out = sb.run("--location-options", "--search-from", sb.client, "--property-county", "Norfolk")
    labels = [o["label"] for o in out["options"]]
    check(labels[:2] == ["Main office (Recommended)", "Boston office"], labels)


def t_found_from_grandparent(sb):
    sb.run("--save", sb.file, "--attorney-in-fact", ATTY)
    code, out = sb.run("--show", "--search-from", sb.sub)          # P&S one folder deeper
    check(code == 0 and out["attorney_in_fact"] == ATTY, out)


def t_explicit_settings_file(sb):
    other = os.path.join(sb.root, NAME)
    sb.run("--save", other, "--attorney-in-fact", ATTY)
    code, out = sb.run("--show", "--settings-file", other)
    check(code == 0 and out["attorney_in_fact"] == ATTY, out)


def t_invalid_refused(sb):
    bad = [
        (["--attorney-in-fact", ""], "attorney"),
        (["--attorney-in-fact", "${user_config.attorney_in_fact}"], "attorney"),
        (["--attorney-in-fact", "[Attorney Name]"], "attorney"),
        (["--attorney-in-fact", ATTY, "--signing-locations", "Main office = Plymouth County, MA; junk"], "skipped"),
        (["--attorney-in-fact", ATTY, "--signing-locations", "Office = Gotham County, MA"], "skipped"),
        (["--attorney-in-fact", ATTY, "--signature-page", "sometimes"], "signature_page"),
        (["--attorney-in-fact", ATTY, "--notary-state", "Atlantis"], "default_notary_state"),
    ]
    for args, word in bad:
        code, out = sb.run("--save", sb.file, *args)
        check(code == 1 and not out.get("saved"), f"{args}: not refused: {out}")
        check(any(word in e for e in out.get("errors", [])), f"{args}: wrong reason: {out}")
        check(not os.path.exists(sb.file), f"{args}: a file was written")
    code, out = sb.run("--save", os.path.join(sb.clients, "settings.json"), "--attorney-in-fact", ATTY)
    check(code == 1, "a wrongly named file was accepted")
    code, out = sb.run("--save", os.path.join(sb.root, "missing", NAME), "--attorney-in-fact", ATTY)
    check(code == 1, "a missing folder was accepted")


def t_save_keeps_other_values(sb):
    sb.run("--save", sb.file, "--attorney-in-fact", ATTY, "--signing-locations", LOCS)
    code, out = sb.run("--save", sb.file, "--signature-page", "ask")
    check(code == 0, out)
    data = json.load(open(sb.file, encoding="utf-8"))
    check(data["attorney_in_fact"] == ATTY and len(data["signing_locations"]) == 2
          and data["signature_page"] == "ask", data)
    code, out = sb.run("--save", sb.file, "--signing-locations", "")        # clear the list
    check(code == 0 and json.load(open(sb.file, encoding="utf-8"))["signing_locations"] == [], out)


def t_plugin_settings_win(sb):
    sb.run("--save", sb.file, "--attorney-in-fact", ATTY, "--signing-locations", LOCS)
    code, out = sb.run("--show", "--attorney-in-fact", "Pat Plugin, Esq.",
                       "--signing-locations", "${user_config.signing_locations}", "--search-from", sb.client)
    check(out["attorney_in_fact"] == "Pat Plugin, Esq." and out["sources"]["attorney_in_fact"] == "plugin setting", out)
    check(out["sources"]["signing_locations"].startswith("firm settings file"), out["sources"])
    sb.plugin_config(attorney_in_fact="Sam Saved, Esq.")                    # pluginConfigs too
    code, out = sb.run("--show", "--attorney-in-fact", "", "--search-from", sb.client)
    check(out["attorney_in_fact"] == "Sam Saved, Esq.", out)
    check(out["sources"]["attorney_in_fact"].startswith("Claude Code plugin settings"), out["sources"])


def t_unreadable_file_warns(sb):
    with open(sb.file, "w", encoding="utf-8") as f:
        f.write("{ not json")
    code, out = sb.run("--show", "--search-from", sb.client)
    check(code == 1 and any("could not be read" in w for w in out.get("warnings", [])), out)
    code, out = sb.run("--save", sb.file, "--attorney-in-fact", ATTY)
    check(code == 1 and open(sb.file, encoding="utf-8").read() == "{ not json", "a broken file was overwritten")


CHECKS = [(n[2:], f) for n, f in sorted(globals().items()) if n.startswith("t_") and callable(f)]


def run_checks():
    """[(name, problem or None)] -- used by tests/run_tests.py."""
    results = []
    for name, fn in CHECKS:
        sb = Sandbox()
        try:
            fn(sb)
            results.append((name, None))
        except AssertionError as e:
            results.append((name, str(e)[:500]))
        finally:
            sb.close()
    return results


def main():
    res = run_checks()
    for name, problem in res:
        print(f"{'FAIL' if problem else 'ok  '}  settings_{name}" + (f"\n      {problem}" if problem else ""))
    fails = sum(1 for _, p in res if p)
    print(f"\n{len(res) - fails}/{len(res)} passed")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
