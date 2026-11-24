"""Claude Launcher: one button per profile/task. Opens a PowerShell window that
sets the profile's Claude account folder, goes to its workspace, registers its MCP,
runs optional start-up commands, then starts Claude Code (optionally with a first prompt).

Terminal use:  cl Excel   (after installing the shortcut, see README)
Window use:  pythonw launcher.py      (or double-click Claude-Launcher.bat)
Profiles live in %USERPROFILE%\\claude-launcher\\profiles.json (created on first run).
"""
from __future__ import annotations

import base64
import json
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
CONFIG_DIR = Path(os.environ.get("CLAUDE_LAUNCHER_HOME", Path.home() / "claude-launcher"))
CONFIG_FILE = CONFIG_DIR / "profiles.json"


def default_config() -> dict:
    home = str(Path.home())
    return {"profiles": [{
        "name": "Excel",
        "config_dir": home + r"\.claude-excel",
        "workspace": home + r"\workspaces\excel",
        "env": {"EXCEL_MCP_ROOT": home + r"\workspaces\excel"},
        "mcp": {"name": "excel-accounting", "command": sys.executable, "args": [str(HERE / "server.py")]},
        "skill_folder": str(HERE / "skills" / "accounting-excel"),
        "startup_commands": [],
        "first_prompt": "",
        "tasks": [
            {"label": "Build company model",
             "prompt": "Use the excel-accounting tools: build the 13-tab company model as SH-Window-model.xlsx, then open it."},
            {"label": "Check a workbook",
             "prompt": "Use the excel-accounting tools: list my workbooks and run check_workbook and check_conventions on each."},
        ],
    }]}


def load_config() -> dict:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if not CONFIG_FILE.exists():
        CONFIG_FILE.write_text(json.dumps(default_config(), indent=2), encoding="utf-8")
    return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))


def q(s: str) -> str:
    """PowerShell single-quoted literal."""
    return "'" + str(s).replace("'", "''") + "'"


def build_script(profile: dict, prompt: str | None = None) -> str:
    """PowerShell text that sets up and starts Claude Code for one profile."""
    cfg, ws = profile["config_dir"], profile["workspace"]
    lines = [
        f"$env:CLAUDE_CONFIG_DIR = {q(cfg)}",
        f"New-Item -ItemType Directory -Force -Path {q(cfg)}, {q(ws)} | Out-Null",
        f"$host.UI.RawUI.WindowTitle = {q('Claude - ' + profile['name'])}",
    ]
    for k, v in (profile.get("env") or {}).items():
        lines.append(f"$env:{k} = {q(v)}")
    skill = profile.get("skill_folder")
    if skill:
        lines.append(f"if (Test-Path {q(skill)}) {{ New-Item -ItemType Directory -Force -Path {q(cfg + chr(92) + 'skills')} | Out-Null; "
                     f"Copy-Item {q(skill)} {q(cfg + chr(92) + 'skills')} -Recurse -Force }}")
    lines.append(f"Set-Location {q(ws)}")
    mcp = profile.get("mcp")
    if mcp:
        env_args = " ".join(f"--env {q(k + '=' + v)}" for k, v in (profile.get("env") or {}).items())
        cmd_args = " ".join(q(a) for a in [mcp["command"], *mcp.get("args", [])])
        lines.append(f"claude mcp add {q(mcp['name'])} {env_args} -- {cmd_args} 2>$null | Out-Null")
    lines += list(profile.get("startup_commands") or [])
    first = prompt if prompt is not None else profile.get("first_prompt", "")
    lines.append(f"claude {q(first)}" if first else "claude")
    return "\n".join(lines)


def encode(script: str) -> str:
    return base64.b64encode(script.encode("utf-16-le")).decode("ascii")


def launch(profile: dict, prompt: str | None = None) -> None:
    enc = encode(build_script(profile, prompt))
    subprocess.Popen(["cmd", "/c", "start", "", "powershell", "-NoExit", "-ExecutionPolicy", "Bypass", "-EncodedCommand", enc])


def find_profile(name: str) -> dict | None:
    for p in load_config()["profiles"]:
        if p["name"].lower() == name.lower():
            return p
    return None


def cli(argv: list[str]) -> int:
    """`launcher.py --encoded NAME [PROMPT]` prints the start-up script as base64 for
    `powershell -EncodedCommand`, so a shell function can start a profile in the CURRENT terminal."""
    if len(argv) >= 2 and argv[0] == "--encoded":
        p = find_profile(argv[1])
        if not p:
            names = ", ".join(x["name"] for x in load_config()["profiles"])
            print(f"No profile named {argv[1]}. Profiles: {names}", file=sys.stderr)
            return 1
        print(encode(build_script(p, argv[2] if len(argv) > 2 else None)))
        return 0
    print("Usage: launcher.py [--encoded NAME [PROMPT]]", file=sys.stderr)
    return 2


def main() -> None:
    import tkinter as tk
    from tkinter import messagebox

    cfg = load_config()
    root = tk.Tk()
    root.title("Claude Launcher")
    root.geometry("460x520")

    def draw():
        for w in root.winfo_children():
            w.destroy()
        for p in load_config()["profiles"]:
            box = tk.LabelFrame(root, text=p["name"], padx=8, pady=6)
            box.pack(fill="x", padx=10, pady=6)
            tk.Button(box, text=f"Open {p['name']} (Claude Code)", command=lambda p=p: launch(p)).pack(fill="x", pady=2)
            for t in p.get("tasks", []):
                tk.Button(box, text="  Task: " + t["label"], anchor="w",
                          command=lambda p=p, t=t: launch(p, t["prompt"])).pack(fill="x", pady=1)
            tk.Button(box, text="Open workspace folder", anchor="w",
                      command=lambda p=p: (Path(p["workspace"]).mkdir(parents=True, exist_ok=True), os.startfile(p["workspace"]))).pack(fill="x", pady=1)
        bar = tk.Frame(root); bar.pack(side="bottom", fill="x", pady=8)
        tk.Button(bar, text="Edit profiles", command=lambda: os.startfile(CONFIG_FILE)).pack(side="left", padx=10)
        tk.Button(bar, text="Reload", command=draw).pack(side="left")

    try:
        draw()
    except Exception as e:  # bad JSON etc.
        messagebox.showerror("Claude Launcher", f"Cannot read {CONFIG_FILE}:\n{e}")
        raise
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        sys.exit(cli(sys.argv[1:]))
    main()
