import base64, json, os, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
os.environ["CLAUDE_LAUNCHER_HOME"] = tempfile.mkdtemp()
import launcher as L


def test_all():
    cfg = L.load_config()
    assert L.CONFIG_FILE.exists() and cfg["profiles"][0]["name"] == "Excel"
    p = cfg["profiles"][0]
    s = L.build_script(p)
    assert "$env:CLAUDE_CONFIG_DIR" in s and "Set-Location" in s and "claude mcp add 'excel-accounting'" in s
    assert s.rstrip().endswith("claude")
    s2 = L.build_script(p, "It's a test\nline 2")
    assert s2.rstrip().endswith("claude 'It''s a test\nline 2'")
    assert base64.b64decode(L.encode(s)).decode("utf-16-le") == s
    p2 = dict(p, startup_commands=["git pull"], mcp=None)
    s3 = L.build_script(p2)
    assert "git pull" in s3 and "mcp add" not in s3
    assert L.q("a'b") == "'a''b'"
    assert L.find_profile("excel")["name"] == "Excel" and L.find_profile("nope") is None
    assert L.cli(["--encoded", "nope"]) == 1


if __name__ == "__main__":
    test_all(); print("LAUNCHER OK")
