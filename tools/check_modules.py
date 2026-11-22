#!/usr/bin/env python3
"""Lightweight static inventory check for Word VBA .bas files (not a VBA compiler)."""
from pathlib import Path
import re, hashlib, sys

BASE = Path(__file__).resolve().parents[1]
WORK = BASE / "src" / "working"
files = sorted(WORK.glob("*.bas"))
if not files:
    print("No .bas modules found.")
    sys.exit(1)

seen = {}
errors = []
for path in files:
    raw = path.read_bytes()
    # Decode only for ASCII structural tokens; preserve source bytes.
    text = raw.decode("latin-1")
    attr = re.search(r'Attribute\s+VB_Name\s*=\s*"([^"]+)"', text, re.I)
    name = attr.group(1) if attr else "(missing VB_Name)"
    procs = re.findall(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function)\s+(\w+)', text, re.M | re.I)
    print(f"{path.name:18} module={name:14} procedures={len(procs):2} sha256={hashlib.sha256(raw).hexdigest()[:12]}")
    if name in seen:
        errors.append(f"Duplicate VB_Name {name}: {seen[name]} and {path.name}")
    seen[name] = path.name
    if not attr:
        errors.append(f"{path.name}: missing Attribute VB_Name")
if errors:
    print("\\n".join("ERROR: " + e for e in errors))
    sys.exit(2)
print(f"OK: {len(files)} .bas files inspected. This is a structural check, not a VBA compile.")
