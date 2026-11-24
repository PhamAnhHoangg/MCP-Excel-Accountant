"""Private company profile. The repository holds ONLY generic placeholders.

Real names and business detail live in `company_profile.json` (git-ignored). Lookup order:
  1. the file named by the EXCEL_MCP_PROFILE environment variable
  2. company_profile.json next to this file
  3. ~/.excel-mcp/company_profile.json
If none exists, generic defaults are used. Text in templates and the catalog may contain tokens
such as @B1@ which are replaced by the profile's real names when a workbook is built.
"""
from __future__ import annotations

import json
import os
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_TOKENS = {
    "B1": "Company A",            # main contracting / fabricating company
    "B2": "Company A Export",     # export arm
    "B3": "Company B",            # system owner / licensor company
    "SUP": "the hardware supplier",
    "STD": "standard brands",
    "JP": "Japanese brands",
}
_TOKEN = re.compile(r"@([A-Z0-9]+)@")
_cache: dict | None = None


def _path() -> Path | None:
    cands = [os.environ.get("EXCEL_MCP_PROFILE"), HERE / "company_profile.json", Path.home() / ".excel-mcp" / "company_profile.json"]
    for c in cands:
        if c and Path(c).is_file():
            return Path(c)
    return None


def load(refresh: bool = False) -> dict:
    global _cache
    if _cache is None or refresh:
        p = _path()
        data = json.loads(p.read_text(encoding="utf-8")) if p else {}
        _cache = {"tokens": {**DEFAULT_TOKENS, **data.get("tokens", {})},
                  "canvas": {int(k): v for k, v in data.get("canvas", {}).items()},
                  "customers": data.get("customers"), "path": str(p) if p else None}
    return _cache


def tokens() -> dict:
    return load()["tokens"]


def sub(text):
    """Replace @TOKEN@ placeholders with the profile's real values (unknown tokens stay as they are)."""
    if not isinstance(text, str) or "@" not in text:
        return text
    t = tokens()
    return _TOKEN.sub(lambda m: t.get(m.group(1), m.group(0)), text)


def default_entity() -> str:
    return tokens()["B1"]


def canvas_override(default: dict) -> dict:
    return {**default, **load()["canvas"]}


def customers_override(default: list) -> list:
    c = load()["customers"]
    return [tuple(x) for x in c] if c else default


def apply_to_workbook(wb) -> int:
    """Replace tokens in every text cell (not formulas that start with '=' unless they contain a token)."""
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "@" in c.value:
                    new = sub(c.value)
                    if new != c.value:
                        c.value = new
                        n += 1
    return n
