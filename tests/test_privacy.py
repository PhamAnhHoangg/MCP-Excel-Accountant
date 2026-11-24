"""Fails if a real name from the PRIVATE profile appears in any file that git would publish."""
import json, os, subprocess, sys
ROOT = os.path.join(os.path.dirname(__file__), "..")
sys.path.insert(0, ROOT)
import company_profile as P


def tracked_files():
    out = subprocess.run(["git", "ls-files", "--cached", "--others", "--exclude-standard"], cwd=ROOT, capture_output=True, text=True)
    if out.returncode != 0:
        raise RuntimeError("git not available")
    return [f for f in out.stdout.splitlines() if f]


def test_no_private_names_in_publishable_files():
    prof = P.load(refresh=True)
    if not prof["path"]:
        print("(no private profile found: privacy scan skipped)")
        return
    words = {v for k, v in prof["tokens"].items() if k not in P.DEFAULT_TOKENS or v != P.DEFAULT_TOKENS[k]}
    words |= {w.strip() for v in prof["tokens"].values() for w in v.split(",")}
    words = {w for w in words if len(w) >= 4 and w.lower() not in {"company"}}
    bad = []
    for f in tracked_files():
        if f == "company_profile.json" or not os.path.isfile(os.path.join(ROOT, f)):
            continue
        try:
            text = open(os.path.join(ROOT, f), encoding="utf-8").read().lower()
        except UnicodeDecodeError:
            continue
        bad += [(f, w) for w in words if w.lower() in text]
    assert not bad, f"private names found in publishable files: {bad}"


if __name__ == "__main__":
    test_no_private_names_in_publishable_files(); print("PRIVACY OK")
