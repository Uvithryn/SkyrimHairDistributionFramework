#!/usr/bin/env python3
"""Rebuild the EMBEDDED_DATA sample block inside index.html.

The sample is what the tool falls back to when it can't fetch Data/*.json (e.g.
someone opens the raw file). It keeps every hairstyle and requirement, but only
a slice of NPCs and previews so the file stays a reasonable size.
"""
import json, re, sys
from pathlib import Path
from collections import defaultdict

DATA = Path(sys.argv[1] if len(sys.argv) > 1 else "data_v2")
HTML = Path(sys.argv[2] if len(sys.argv) > 2 else "app.html")
OUT = Path(sys.argv[3] if len(sys.argv) > 3 else "index_final.html")

npcs = json.loads((DATA / "npcs.json").read_text())
hairs = json.loads((DATA / "hairstyles.json").read_text())
races = json.loads((DATA / "races.json").read_text())
reqs = json.loads((DATA / "requirements.json").read_text())
previews = json.loads(Path("previews.json").read_text())
previews = previews.get("images", previews)

# NPC sample: a spread across race types and genders, plus a few children.
by_group = defaultdict(list)
for n in npcs:
    by_group[(n.get("raceType"), n.get("gender"))].append(n)
sample_npcs = []
for key in sorted(by_group, key=lambda k: (str(k[0]), str(k[1]))):
    sample_npcs.extend(by_group[key][:5])

# Preview sample: keep a slice per (mod, race, gender) so the fallback chain and
# the browser both have something to show.
by_pv = defaultdict(list)
for p in previews:
    by_pv[(p["mod"], p["race"], p["gender"])].append(p)
sample_pv = []
for key in sorted(by_pv):
    sample_pv.extend(by_pv[key][:6])

blob = {
    "npcs": sample_npcs,
    "hairstyles": hairs,          # all of them: the browser needs the full list
    "races": races,
    "previews": sample_pv,
    "requirements": reqs,
}
payload = json.dumps(blob, ensure_ascii=True, separators=(",", ":"))

html = HTML.read_text()
pat = re.compile(r"(<script>const EMBEDDED_DATA = )(.*?)(;</script>)", re.S)
if not pat.search(html):
    sys.exit("could not find the EMBEDDED_DATA block")
html = pat.sub(lambda m: m.group(1) + payload + m.group(3), html, count=1)
OUT.write_text(html)
print(f"sample: {len(sample_npcs)} NPCs, {len(hairs)} hairstyles, "
      f"{len(sample_pv)} previews, {len(reqs)} requirements")
print(f"wrote {OUT} ({OUT.stat().st_size} bytes)")
