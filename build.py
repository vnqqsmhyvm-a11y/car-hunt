#!/usr/bin/env python3
"""Build the Car Hunt site from the master spreadsheet.

Usage: python3 build.py <path/to/tulsa-used-cars.xlsx>

Reads the 'Cars', 'Sources' and 'Hyundai-Kia Notes' sheets and photos/manifest.json
(written by fetch_photos.py) and writes index.html, tips/index.html, models/index.html (from guide/), robots.txt, .nojekyll.
"""
import datetime as dt
import html
import json
import re
import sys
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl

ROOT = Path(__file__).resolve().parent
CT = ZoneInfo("America/Chicago")
YELLOW = {"FFFFF2CC", "FFFFFF00", "FFFFEB9C", "FFFFF2CC"}

COLS = {  # header text in the sheet -> field name
    "Year": "year", "Make": "make", "Model": "model", "Trim": "trim", "Body": "body",
    "Miles": "miles", "Asking Price": "price", "MPG": "mpg", "City": "city", "State": "state",
    "Seller": "seller", "Title": "title", "Owners": "owners", "Accidents": "accidents",
    "Est. Annual Repair Cost": "repair", "Reliability": "reliability",
    "Known Issues (this year)": "issues", "Est. Ins": "ins", "IIHS": "iihs",
    "NHTSA Stars": "nhtsa", "Notes": "notes", "Listing": "listing",
    "Drive time from Tulsa": "drive", "Service records": "service",  # Budget sheet only
}
OPTIONAL = {"drive", "service"}
LISTS = {  # sheet -> page settings
    "Cars": dict(id="main", path="", label="Main", range="$15–19k",
                 criteria="$15,000–$19,000 · under 100k miles · about 40 miles from Tulsa · 2016+ automatic · clean title · Model Guide picks only",
                 under=(17000, "Under $17k")),
    "Budget": dict(id="budget", path="budget/", label="Budget", range="$10–15k",
                   criteria="$10,000–$15,000 · up to 125k miles (over 100k: Toyota, Honda, Mazda only) · within 2 hours of Tulsa · 2016+ automatic · clean title",
                   under=(12500, "Under $12.5k")),
}


# ---------------- data ----------------

def norm_url(u):
    if not u:
        return ""
    u = u.strip().split("#")[0].split("?")[0]
    return u if u.endswith("/") else u + "/"


def slugify(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def to_int(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return int(v)
    m = re.sub(r"[^\d.]", "", str(v))
    return int(float(m)) if m else None


def read_cars(xlsx, sheet="Cars"):
    wb = openpyxl.load_workbook(xlsx)
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    hdr = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    idx = {COLS[h]: i for i, h in enumerate(hdr) if h in COLS}
    missing = [f for f in COLS.values() if f not in idx and f not in OPTIONAL]
    if missing:
        print(f"warning: columns not found: {missing}", file=sys.stderr)
    cars, slugs = [], set()
    for row in ws.iter_rows(min_row=2):
        get = lambda f: row[idx[f]].value if f in idx else None
        if not get("make") or not get("year"):
            continue
        cell = row[idx["listing"]] if "listing" in idx else None
        url = ""
        if cell is not None:
            if cell.hyperlink and cell.hyperlink.target:
                url = cell.hyperlink.target
            elif isinstance(cell.value, str) and cell.value.startswith("http"):
                url = cell.value
        c = {f: get(f) for f in COLS.values()}
        c = {k: (v.strip() if isinstance(v, str) else v) for k, v in c.items()}
        c["year"] = to_int(c["year"])
        c["miles"] = to_int(c["miles"])
        c["price"] = to_int(c["price"])
        c["url"] = norm_url(url)
        c["make"] = str(c["make"])
        c["model"] = str(c["model"])
        c["yellow"] = any((x.fill.fill_type and x.fill.fgColor.rgb in YELLOW) for x in row)
        c["key"] = c["url"] or f'{c["year"]}|{c["make"]}|{c["model"]}|{c["price"]}'.lower()
        base = slugify(f'{c["year"]} {c["make"]} {c["model"]} {c["price"]}')
        slug, n = base, 2
        while slug in slugs:
            slug, n = f"{base}-{n}", n + 1
        slugs.add(slug)
        c["slug"] = slug
        c["row"] = row[0].row
        c["sheet"] = sheet
        cars.append(c)
    return cars


def avoid_rules():
    data = json.loads((ROOT / "guide" / "models.json").read_text(encoding="utf-8"))
    return [(e, e["match"]) for e in data["entries"] if e["verdict"] == "avoid" and e.get("match")]


def avoid_match(c, rules):
    """The Model Guide Avoid entry this car falls under (make/model/year, plus trim or notes patterns), or None."""
    norm = lambda x: re.sub(r"[^a-z0-9]", "", str(x or "").lower())
    for e, m in rules:
        if norm(c["make"]) not in {norm(x) for x in m.get("makes", [])}:
            continue
        if m.get("models") and norm(c["model"]) not in {norm(x) for x in m["models"]}:
            continue
        if m.get("years") and not (c["year"] and m["years"][0] <= c["year"] <= m["years"][1]):
            continue
        if m.get("trim") and not re.search(m["trim"], str(c["trim"] or ""), re.I):
            continue
        if m.get("notes") and not re.search(m["notes"], str(c["notes"] or ""), re.I):
            continue
        return e
    return None


SOLD_RE = re.compile(r"\b(sold|listing\s+removed|removed)\b", re.I)


def is_sold(c):
    """True when Notes say the car is sold or the listing was removed (not photo recheck)."""
    return bool(SOLD_RE.search(str(c.get("notes") or "")))


def read_all(xlsx):
    """Cars from every list sheet, with slugs unique across sheets."""
    out, seen = {}, set()
    for sheet in LISTS:
        cars = read_cars(xlsx, sheet)
        for c in cars:
            base, n = c["slug"], 2
            while c["slug"] in seen:
                c["slug"], n = f"{base}-{n}", n + 1
            seen.add(c["slug"])
        out[sheet] = cars
    return out


def read_pairs(xlsx, sheet):
    wb = openpyxl.load_workbook(xlsx)
    if sheet not in wb.sheetnames:
        return []
    out = []
    for r in wb[sheet].iter_rows(min_row=2):
        a, b = (r[0].value, r[1].value if len(r) > 1 else None)
        if a is None and b is None:
            continue
        out.append((str(a or "").strip(), str(b or "").strip(), bool(r[0].font and r[0].font.b)))
    return out


def updated_time(xlsx):
    wb = openpyxl.load_workbook(xlsx)
    m = wb.properties.modified
    if m:
        return m.replace(tzinfo=dt.timezone.utc).astimezone(CT)
    return dt.datetime.fromtimestamp(Path(xlsx).stat().st_mtime, CT)


# ---------------- presentation helpers ----------------

def esc(s):
    return html.escape("" if s is None else str(s), quote=True)


def money(n):
    return f"${n:,.0f}" if isinstance(n, (int, float)) else esc(n)


def fmt_ct(t):
    h = t.strftime("%I").lstrip("0")
    return f"{t:%a}, {t:%b} {t.day}, {t.year}, {h}:{t:%M} {t:%p} CT"


def group(make):
    m = make.lower()
    return "hyundai-kia" if m in ("hyundai", "kia") else slugify(make)


def title_badge(c):
    t = str(c["title"] or "")
    if t.lower().startswith("clean"):
        return ("good", "Clean title")
    return ("warn", "Verify title")


def accident_badge(c):
    a = str(c["accidents"] or "").strip().lower()
    if a in ("0", "none", "no") or a.startswith("none"):
        return ("good", "No accidents")
    if "report" in a or (a.isdigit() and int(a) > 0):
        return ("bad", "Accident reported")
    return ("muted", "Accidents not stated")


def reliability_badge(c):
    r = str(c["reliability"] or "")
    m = re.match(r"\s*([\d.]+)\s*/\s*5", r)
    if m:
        return ("info", f"Reliability {m.group(1)}/5")
    m = re.search(r"Consumer Reports ([a-z ]+)", r)
    if m:
        return ("info", f"Reliability: {m.group(1).strip()}")
    return ("muted", "Reliability n/a")


def ins_monthly(c):
    v = c["ins"]
    return int(v) if isinstance(v, (int, float)) else None


def mpg_num(c):
    m = re.match(r"\s*(\d+(?:-\d+)?)", str(c["mpg"] or ""))
    return m.group(1) if m else None


def crash_badges(c):
    out = []
    n = c["nhtsa"]
    if isinstance(n, (int, float)):
        out.append(("info" if n >= 5 else "muted", f"NHTSA {int(n)}★"))
    else:
        out.append(("muted", "NHTSA not rated"))
    i = str(c["iihs"] or "")
    if i.upper().startswith("TOP SAFETY PICK+"):
        out.append(("good", "IIHS Top Safety Pick+"))
    elif i.upper().startswith("TOP SAFETY PICK"):
        out.append(("good", "IIHS Top Safety Pick"))
    return out


def flag(c):
    """AVOID (red) / CAUTION (amber) from the Notes column. Lower-case 'avoid 1.6T/DCT trims'
    is generic advice, not a flag, so AVOID must be upper-case."""
    notes = str(c["notes"] or "")
    parts = [p.strip() for p in notes.split(";")]
    def clean(p):
        m = re.match(r"^(?:AVOID|CAUTION)(?:/(?:AVOID|CAUTION))?\s*(?:per\s+([^:]+))?:\s*(.*)$", p)
        if m:
            src = (m.group(1) or "").replace("Hyundai-Kia", "Hyundai/Kia").strip()
            src = src[0].upper() + src[1:] if src else ""
            p = f"{src}: {m.group(2)}" if src else m.group(2)
        p = re.sub(r"\s*-\s*caution\s*$", "", p, flags=re.I)
        return p[0].upper() + p[1:] if p else p
    for p in parts:
        if re.search(r"\bAVOID\b", p):
            return ("avoid", "Avoid", clean(p))
    for p in parts:
        if re.search(r"\bcaution\b", p, re.I):
            return ("caution", "Caution", clean(p))
    return None


# ---------------- page shell ----------------

CSS = r"""
@font-face{font-family:"Playfair Display";font-weight:700;font-display:swap;src:url(%(base)sassets/fonts/playfair-display-latin-700-normal.woff2) format("woff2")}
:root{color-scheme:light;--bg:#faf6f0;--card:#fff;--ink:#22201c;--soft:#5d574e;--muted:#8a8378;--line:#ebe4d8;
--accent:#c2562b;--accent-ink:#fff;--good:#1f7a4d;--good-bg:#e5f4ec;--warn:#8a5a00;--warn-bg:#fff1cc;--bad:#b42318;--bad-bg:#fde8e6;
--info:#1d4f91;--info-bg:#e8f0fb;--muted-bg:#f1ede6;--radius:18px;
--sans:-apple-system,BlinkMacSystemFont,"SF Pro Text","Segoe UI",system-ui,Roboto,"Helvetica Neue",Arial,sans-serif;
--display:"Playfair Display",Georgia,serif}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%%;text-size-adjust:100%%}
body{margin:0;background:var(--bg);color:var(--ink);font:400 17px/1.5 var(--sans);-webkit-font-smoothing:antialiased}
a{color:var(--accent)}
.wrap{max-width:76rem;margin:0 auto;padding:0 max(1rem,env(safe-area-inset-left)) 2rem max(1rem,env(safe-area-inset-right))}
header.top{padding:1.4rem 0 .6rem}
header.top h1{font:700 2.4rem/1.05 var(--display);margin:0;letter-spacing:-.01em}
header.top h1 a{color:var(--ink);text-decoration:none}
.sub{margin:.35rem 0 .2rem;color:var(--soft);font-size:1.05rem}
.meta{color:var(--muted);font-size:.88rem;margin:0}
.tabs{display:flex;gap:.5rem;margin:.9rem 0 0}
.tabs a{font-weight:600;font-size:.95rem;text-decoration:none;color:var(--ink);padding:.5rem .9rem;border-radius:999px;background:var(--muted-bg)}
.tabs a[aria-current]{background:var(--ink);color:#fff}
.controls{background:rgba(250,246,240,.94);-webkit-backdrop-filter:blur(8px);backdrop-filter:blur(8px);
 margin:0 calc(-1*max(1rem,env(safe-area-inset-left)));padding:.6rem max(1rem,env(safe-area-inset-left));border-bottom:1px solid var(--line)}
.chips{display:flex;flex-wrap:wrap;gap:.45rem}
.chips::-webkit-scrollbar{display:none}
.chip{flex:0 0 auto;font:600 .95rem/1 var(--sans);padding:.62rem .95rem;border-radius:999px;border:1.5px solid var(--line);background:var(--card);color:var(--ink);cursor:pointer}
.chip[aria-pressed=true]{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}
.row2{display:flex;justify-content:space-between;align-items:center;gap:.6rem;margin-top:.5rem;font-size:.92rem;color:var(--soft)}
.row2 label{display:flex;align-items:center;gap:.4rem;white-space:nowrap}
#shown{white-space:nowrap}
.row2 select{font:600 .95rem var(--sans);padding:.45rem .6rem;border-radius:10px;border:1.5px solid var(--line);background:var(--card);color:var(--ink)}
.grid{display:grid;grid-template-columns:1fr;gap:1.1rem;margin-top:1rem;list-style:none;padding:0}
@media (min-width:700px){.grid{grid-template-columns:1fr 1fr}.controls{position:sticky;top:0;z-index:5}}
@media (min-width:1100px){.grid{grid-template-columns:1fr 1fr 1fr}}
.card{background:var(--card);border-radius:var(--radius);overflow:hidden;box-shadow:0 1px 2px rgba(60,40,10,.06),0 6px 18px rgba(60,40,10,.06);display:flex;flex-direction:column}
.card[hidden]{display:none}
.photo{position:relative;background:#eee7dc}
.gallery{display:flex;overflow-x:auto;scroll-snap-type:x mandatory;scrollbar-width:none;aspect-ratio:4/3}
.gallery::-webkit-scrollbar{display:none}
.slide{flex:0 0 100%%;position:relative;overflow:hidden;scroll-snap-align:start}
.slide img{display:block;width:100%%;height:100%%;object-fit:cover}
.slide.tall img.fg{position:relative;object-fit:contain}
.slide img.bg{position:absolute;inset:0;filter:blur(22px) saturate(1.05);opacity:.75;transform:scale(1.2)}
.count{position:absolute;right:.6rem;bottom:.6rem;background:rgba(0,0,0,.62);color:#fff;font-size:.78rem;font-weight:600;padding:.25rem .55rem;border-radius:999px}
.nav{position:absolute;top:50%%;transform:translateY(-50%%);width:2.3rem;height:2.3rem;border-radius:50%%;border:0;background:rgba(255,255,255,.9);color:#222;font-size:1.25rem;line-height:1;cursor:pointer;box-shadow:0 1px 4px rgba(0,0,0,.25)}
.nav.prev{left:.5rem}.nav.next{right:.5rem}
.noimg{aspect-ratio:4/3;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:.4rem;color:var(--muted);background:repeating-linear-gradient(45deg,#f1ebe1,#f1ebe1 12px,#ece5d9 12px,#ece5d9 24px);font-weight:600}
.noimg svg{width:3.2rem;height:3.2rem;opacity:.55}
.ribbon{position:absolute;left:0;right:0;top:0;padding:.5rem .8rem;font-weight:700;font-size:.9rem;color:#fff;background:rgba(40,36,30,.86);text-align:center;letter-spacing:.02em}
.card.removed .gallery,.card.removed .noimg{filter:grayscale(1);opacity:.6}
.sold-wrap{margin-top:1.6rem}
.sold-wrap > summary{font:700 1.05rem/1.3 var(--sans);color:var(--soft)}
.sold-wrap .grid{opacity:.92}
.flag{padding:.6rem .9rem;font-size:.9rem;line-height:1.35}
.flag b{text-transform:uppercase;letter-spacing:.05em;margin-right:.35rem}
.flag.avoid{background:var(--bad);color:#fff}
.flag.caution{background:#f5b301;color:#2a1d00}
.body{padding:.9rem 1rem 1rem;display:flex;flex-direction:column;gap:.55rem;flex:1}
.card h2{font-size:1.18rem;line-height:1.25;margin:0;font-weight:700}
.price{font-size:1.75rem;font-weight:800;letter-spacing:-.01em;line-height:1}
.facts{color:var(--soft);font-size:.95rem;margin:0}
.facts span+span::before{content:"·";margin:0 .4rem;color:var(--muted)}
.badges{display:flex;flex-wrap:wrap;gap:.35rem;margin:.1rem 0 0;padding:0;list-style:none}
.b{font-size:.8rem;font-weight:650;padding:.3rem .55rem;border-radius:8px;white-space:nowrap}
.b.good{background:var(--good-bg);color:var(--good)}.b.warn{background:var(--warn-bg);color:var(--warn)}
.b.bad{background:var(--bad-bg);color:var(--bad)}.b.info{background:var(--info-bg);color:var(--info)}
.b.muted{background:var(--muted-bg);color:var(--soft)}
.actions{margin-top:auto;padding-top:.3rem;display:flex;flex-direction:column;gap:.55rem}
.btn{display:block;text-align:center;background:var(--accent);color:#fff;font-weight:700;font-size:1.05rem;padding:.85rem 1rem;border-radius:12px;text-decoration:none}
.btn.off{background:var(--muted-bg);color:var(--soft)}
details{border:1.5px solid var(--line);border-radius:12px}
summary{cursor:pointer;font-weight:650;padding:.75rem .9rem;list-style:none;display:flex;justify-content:space-between;align-items:center}
summary::-webkit-details-marker{display:none}
summary::after{content:"+";font-size:1.3rem;line-height:1;color:var(--accent)}
details[open] summary::after{content:"–"}
dl{margin:0;padding:0 .9rem .8rem;font-size:.92rem}
dl div{padding:.5rem 0;border-top:1px solid var(--line)}
dt{font-weight:650;color:var(--soft);font-size:.8rem;text-transform:uppercase;letter-spacing:.04em}
dd{margin:.15rem 0 0}
.empty{padding:2rem;text-align:center;color:var(--soft)}
footer.bottom{margin-top:2.2rem;padding-top:1.2rem;border-top:1px solid var(--line);color:var(--soft);font-size:.88rem}
footer.bottom p{margin:.4rem 0}
.tips{max-width:44rem}
.tips h2{font:700 1.5rem/1.2 var(--display);margin:2rem 0 .6rem}
.tips h3{font-size:1.05rem;margin:1.2rem 0 .3rem}
.tips p,.tips li{color:var(--ink)}
.tips ol{padding-left:1.3rem}.tips li{margin:.45rem 0}
.box{background:var(--card);border-radius:var(--radius);padding:1rem 1.1rem;box-shadow:0 1px 2px rgba(60,40,10,.06),0 6px 18px rgba(60,40,10,.06);margin:1rem 0}
.box.verdict{border-left:5px solid var(--accent)}
.small{font-size:.85rem;color:var(--soft)}
.seg{display:grid;grid-template-columns:1fr 1fr;gap:.5rem;margin:1rem 0 .2rem}
.seg a{display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:.7rem .5rem;border-radius:14px;border:2px solid var(--accent);color:var(--accent);background:var(--card);text-decoration:none;line-height:1.2}
.seg a b{font-size:1.15rem}
.seg a span{font-size:.82rem;font-weight:600;opacity:.85;margin-top:.15rem}
.seg a[aria-current]{background:var(--accent);color:#fff}
.criteria{margin:.6rem 0 .5rem;font-size:.88rem;color:var(--soft);line-height:1.4}
.b.drive{background:#efe7fb;color:#5b2a9e}
.added{margin:0 0 -.25rem;font-size:.8rem;color:var(--muted);display:flex;align-items:center;gap:.45rem}
.newb{background:#ff3b30;color:#fff;font-weight:800;font-size:.72rem;letter-spacing:.08em;text-transform:uppercase;padding:.22rem .5rem;border-radius:6px}
.newb[hidden]{display:none}
.chip.new{border-color:#ff3b30;color:#d62a20}
.chip.new[aria-pressed=true]{background:#ff3b30;border-color:#ff3b30;color:#fff}
.intro{color:var(--soft);margin:1rem 0 .2rem;max-width:44rem}
.chips+.chips{margin-top:.45rem}
.chip.sm{font-size:.85rem;padding:.5rem .75rem}
.vgroup h2{font:700 1.45rem/1.2 var(--display);margin:1.8rem 0 .2rem;display:flex;align-items:center;gap:.55rem}
.vgroup h2 .dot{width:.75rem;height:.75rem;border-radius:50%%;flex:0 0 auto}
.vgroup .lede{margin:0 0 .3rem;color:var(--soft);font-size:.92rem}
.mlist{display:grid;grid-template-columns:1fr;gap:.8rem;list-style:none;padding:0;margin:.7rem 0 0}
@media (min-width:760px){.mlist{grid-template-columns:1fr 1fr}}
@media (min-width:1150px){.mlist{grid-template-columns:1fr 1fr 1fr}}
.m{background:var(--card);border-radius:14px;padding:.85rem .95rem .9rem;box-shadow:0 1px 2px rgba(60,40,10,.06),0 4px 12px rgba(60,40,10,.05);border-left:5px solid var(--line)}
.m[hidden],.vgroup[hidden]{display:none}
.m.rec{border-left-color:var(--good)}.m.ok{border-left-color:#e0a100}.m.avoid{border-left-color:var(--bad)}
.m h3{margin:0;font-size:1.08rem;line-height:1.25}
.m .yrs{display:block;color:var(--soft);font-weight:600;font-size:.92rem;margin-top:.1rem}
.m .tags{display:flex;flex-wrap:wrap;gap:.3rem;margin:.45rem 0 .1rem}
.m p.why{margin:.45rem 0 .55rem;font-size:.95rem;line-height:1.45}
.m dl.k{padding:0;margin:0;display:grid;grid-template-columns:auto 1fr;gap:0 .7rem;font-size:.86rem}
.m dl.k dt,.m dl.k dd{border-top:1px solid var(--line);padding:.32rem 0;margin:0}
.m dl.k dt{font-size:.72rem;padding-top:.42rem;white-space:nowrap}
.m .note{font-size:.86rem;color:var(--soft);border-top:1px solid var(--line);padding-top:.4rem;margin:0}
.dot.rec{background:var(--good)}.dot.ok{background:#e0a100}.dot.avoid{background:var(--bad)}
"""

ICON_CAR = ('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true">'
            '<path d="M3 13l2-5a2 2 0 0 1 1.9-1.3h10.2A2 2 0 0 1 19 8l2 5v4a1 1 0 0 1-1 1h-1a2 2 0 0 1-4 0H9a2 2 0 0 1-4 0H4a1 1 0 0 1-1-1z"/>'
            '<path d="M3 13h18"/></svg>')


DEFAULT_SOURCES = ("<p>Sources: listings from Cars.com, Facebook Marketplace and dealer websites. Reliability and repair costs from RepairPal, "
                   "Consumer Reports and J.D. Power. Crash ratings from IIHS and NHTSA; recalls from NHTSA. MPG from fueleconomy.gov. "
                   "Insurance estimates from CarInsurance.com, Insure.com and MoneyGeek.</p>")


def shell(title, base, current, updated, count, main, js="", sources=DEFAULT_SOURCES):
    tabs = (f'<a href="{base}" {"aria-current=page" if current == "cars" else ""}>Cars</a>'
            f'<a href="{base}tips/" {"aria-current=page" if current == "tips" else ""}>Buying tips</a>'
            f'<a href="{base}models/" {"aria-current=page" if current == "models" else ""}>Model guide</a>')
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="robots" content="noindex, nofollow">
<meta name="referrer" content="no-referrer">
<meta name="theme-color" content="#faf6f0">
<title>{esc(title)}</title>
<style>{CSS % {"base": base}}</style>
</head>
<body>
<div class="wrap">
<header class="top">
<h1><a href="{base}">Car Hunt</a></h1>
<p class="sub">A reliable first car for college · Tulsa area</p>
<p class="meta">Updated {esc(fmt_ct(updated))} · {count} cars</p>
<nav class="tabs" aria-label="Pages">{tabs}</nav>
</header>
{main}
<footer class="bottom">
<p><strong>Insurance figures are estimates, not quotes.</strong> They're built from published averages for an 18-year-old added to a parent's full-coverage policy. Get a real quote on the exact VIN before buying, and tell the insurer the car will be kept in Fayetteville.</p>
{sources}
</footer>
</div>
{js}
</body>
</html>
"""


# ---------------- cards ----------------

def photo_block(c, entry, eager, base=""):
    photos = sorted((p for p in (entry or {}).get("photos", []) if (ROOT / p).exists()),
                    key=lambda p: int(re.search(r"-(\d+)\.jpg$", p).group(1)) if re.search(r"-(\d+)\.jpg$", p) else 0)
    sold = is_sold(c)
    removed = sold or (entry or {}).get("status") == "removed"
    ribbon = ('<div class="ribbon">SOLD</div>' if sold else
              '<div class="ribbon">Listing removed</div>' if removed else "")
    name = f'{c["year"]} {c["make"]} {c["model"]}'
    if not photos:
        return (f'<div class="photo">{ribbon}<div class="noimg" role="img" aria-label="Photo unavailable">{ICON_CAR}'
                f'<span>Photo unavailable</span></div></div>')
    imgs = []
    for i, p in enumerate(photos):
        w, h = img_size(p)
        load = "eager" if (eager and i == 0) else "lazy"
        img = (f'<img class="fg" src="{esc(base + p)}" width="{w}" height="{h}" loading="{load}" decoding="async" '
               f'alt="{esc(name)}, photo {i + 1} of {len(photos)}">')
        if h > w * 0.9:  # portrait phone photo: show it whole over a soft blurred copy
            imgs.append(f'<div class="slide tall"><img class="bg" src="{esc(base + p)}" alt="" aria-hidden="true" loading="{load}">{img}</div>')
        else:
            imgs.append(f'<div class="slide">{img}</div>')
    count = (f'<span class="count" data-count="{len(photos)}">1 / {len(photos)}</span>' if len(photos) > 1 else "")
    return f'<div class="photo">{ribbon}<div class="gallery">{"".join(imgs)}</div>{count}</div>'


_SIZES = {}


def img_size(p):
    if p not in _SIZES:
        from PIL import Image
        with Image.open(ROOT / p) as im:
            _SIZES[p] = im.size
    return _SIZES[p]


def badge_html(kind, text):
    return f'<li class="b {kind}">{esc(text)}</li>'


def details_rows(c):
    ins = ins_monthly(c)
    loc = ", ".join(x for x in [c["city"], c["state"]] if x)
    rows = [
        ("Trim", c["trim"]), ("Body style", c["body"]),
        ("Miles", f'{c["miles"]:,}' if c["miles"] else None),
        ("Asking price", money(c["price"]) if c["price"] else None),
        ("Location", loc), ("Drive from Tulsa", c.get("drive")), ("Seller", c["seller"]),
        ("Title", c["title"]), ("Owners", c["owners"]), ("Accidents", c["accidents"]),
        ("Service records", c.get("service")),
        ("Known issues & recalls", c["issues"]),
        ("Notes", c["notes"]),
        ("Est. insurance", f"About ${ins:,}/month (about ${ins * 12:,}/year), estimate" if ins else
         (c["ins"] if c["ins"] not in (None, "") else None)),
        ("Est. repair cost", f'{c["repair"]}/year' if c["repair"] and str(c["repair"]).startswith("$") else c["repair"]),
        ("Reliability", c["reliability"]), ("Gas mileage (MPG)", c["mpg"]),
        ("IIHS crash tests", c["iihs"]), ("NHTSA overall", f'{c["nhtsa"]} stars' if isinstance(c["nhtsa"], (int, float)) else c["nhtsa"]),
    ]
    return "".join(f"<div><dt>{esc(k)}</dt><dd>{esc(v)}</dd></div>" for k, v in rows if v not in (None, ""))


BUILD_DAY = dt.datetime.now(CT).date().isoformat()
FIRST_SEEN = ROOT / "guide" / "first_seen.json"


def stamp_first_seen(cars):
    """Map car key -> first CT date it appeared. New keys get today's CT date; existing dates never change."""
    seen = json.loads(FIRST_SEEN.read_text()) if FIRST_SEEN.exists() else {}
    added = 0
    for c in cars:
        if c["key"] not in seen:
            seen[c["key"]] = BUILD_DAY
            added += 1
        c["first_seen"] = seen[c["key"]]
    FIRST_SEEN.parent.mkdir(exist_ok=True)
    FIRST_SEEN.write_text(json.dumps(dict(sorted(seen.items())), indent=1) + "\n")
    return added


def added_label(day):
    d = dt.date.fromisoformat(day)
    return f"{d:%b} {d.day}"


def card(c, entry, eager, base=""):
    sold = is_sold(c)
    removed = sold or (entry or {}).get("status") == "removed"
    f = flag(c)
    name = " ".join(str(x) for x in [c["year"], c["make"], c["model"], c["trim"] if c["trim"] and c["trim"] != "n/a" else ""] if x)
    loc = ", ".join(x for x in [c["city"], c["state"]] if x)
    seller = re.sub(r"\s*\((?:FB)\)\s*", " ", str(c["seller"] or "")).strip()
    badges = ([("bad", "SOLD")] if sold else []) + ([("drive", f'{c["drive"]} from Tulsa')] if c.get("drive") else []) + [title_badge(c), accident_badge(c), reliability_badge(c)]
    ins = ins_monthly(c)
    badges.append(("info", f"Insurance ~${ins}/mo") if ins else ("muted", "Insurance n/a"))
    if mpg_num(c):
        badges.append(("muted", f"{mpg_num(c)} MPG"))
    badges += crash_badges(c)
    flag_html = (f'<div class="flag {f[0]}"><b>{esc(f[1])}</b>{esc(f[2])}</div>' if f else "")
    if c["url"] and not removed:
        btn = f'<a class="btn" href="{esc(c["url"])}" target="_blank" rel="noopener noreferrer nofollow">See listing ↗</a>'
    elif c["url"]:
        btn = f'<a class="btn off" href="{esc(c["url"])}" target="_blank" rel="noopener noreferrer nofollow">Listing removed (old link)</a>'
    else:
        btn = '<span class="btn off">No listing link</span>'
    attrs = (f'data-group="{group(c["make"])}" data-price="{c["price"] or 0}" data-miles="{c["miles"] or 0}" '
             f'data-year="{c["year"] or 0}" data-removed="{1 if removed else 0}" data-added="{c["first_seen"]}"')
    is_new = c["first_seen"] == BUILD_DAY
    return f"""<li class="card{' removed' if removed else ''}" id="{c['slug']}" {attrs}>
{photo_block(c, entry, eager, base)}
{flag_html}
<div class="body">
<p class="added"><span class="newb"{'' if is_new else ' hidden'}>New</span>Added {added_label(c["first_seen"])}</p>
<h2>{esc(name)}</h2>
<div class="price">{money(c['price'])}</div>
<p class="facts"><span>{f'{c["miles"]:,} miles' if c["miles"] else 'Miles n/a'}</span><span>{esc(loc)}</span><span>{esc(seller)}</span></p>
<ul class="badges">{''.join(badge_html(*b) for b in badges)}</ul>
<div class="actions">
<details><summary>Details</summary><dl>{details_rows(c)}</dl></details>
{btn}
</div>
</div>
</li>"""


JS = r"""<script>
(function(){
var grid=document.getElementById('cars'),cards=[].slice.call(grid.children),shown=document.getElementById('shown');
var chips=[].slice.call(document.querySelectorAll('.chip')),sel=document.getElementById('sort');
document.querySelectorAll('[data-js]').forEach(function(e){e.hidden=false});
var state={f:'all',s:'price'};
var today;try{today=new Intl.DateTimeFormat('en-CA',{timeZone:'America/Chicago',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());}catch(e){}
var nNew=0;cards.forEach(function(c){var b=c.querySelector('.newb');if(today){c.dataset.isnew=c.dataset.added===today?'1':'0';}else{c.dataset.isnew=b&&!b.hidden?'1':'0';}
 if(b)b.hidden=c.dataset.isnew!=='1';if(c.dataset.isnew==='1')nNew++;});
var nc=document.getElementById('newchip');if(nc)nc.textContent='New today ('+nNew+')';
function ok(c){var g=c.dataset.group,p=+c.dataset.price;
 if(state.f==='all')return true; if(state.f==='new')return c.dataset.isnew==='1'; if(state.f.indexOf('under:')===0)return p<+state.f.slice(6); return g===state.f;}
function apply(){var list=cards.slice().sort(function(a,b){
  var r=(+a.dataset.removed)-(+b.dataset.removed); if(r)return r;
  if(state.s==='miles')return a.dataset.miles-b.dataset.miles;
  if(state.s==='added')return (b.dataset.added<a.dataset.added?-1:b.dataset.added>a.dataset.added?1:0)||(a.dataset.price-b.dataset.price);
  if(state.s==='newest')return (b.dataset.year-a.dataset.year)||(a.dataset.price-b.dataset.price);
  return a.dataset.price-b.dataset.price;});
 var n=0; list.forEach(function(c){var v=ok(c);c.hidden=!v;if(v)n++;grid.appendChild(c);});
 shown.textContent=n===cards.length?('Showing all '+n+' cars'):('Showing '+n+' of '+cards.length+' cars');
 document.getElementById('empty').hidden=n>0||state.f==='new';
 document.getElementById('empty-new').hidden=n>0||state.f!=='new';}
chips.forEach(function(ch){ch.addEventListener('click',function(){state.f=ch.dataset.f;
 chips.forEach(function(x){x.setAttribute('aria-pressed',x===ch?'true':'false')});apply();});});
sel.addEventListener('change',function(){state.s=sel.value;apply();});
if(location.hash==='#new'&&nc){nc.click();}else{apply();}
document.querySelectorAll('.photo').forEach(function(ph){var g=ph.querySelector('.gallery'),cnt=ph.querySelector('.count');
 if(!g||!cnt)return; var n=+cnt.dataset.count;
 function idx(){return Math.round(g.scrollLeft/g.clientWidth)}
 g.addEventListener('scroll',function(){cnt.textContent=(idx()+1)+' / '+n},{passive:true});
 if(matchMedia('(hover:hover)').matches){['prev','next'].forEach(function(d){var b=document.createElement('button');
  b.className='nav '+d;b.type='button';b.setAttribute('aria-label',d==='prev'?'Previous photo':'Next photo');b.textContent=d==='prev'?'‹':'›';
  b.addEventListener('click',function(){var i=Math.min(n-1,Math.max(0,idx()+(d==='prev'?-1:1)));g.scrollTo({left:i*g.clientWidth,behavior:'smooth'})});ph.appendChild(b);});}
});
})();
</script>"""


def build_index(cars, sold_cars, manifest, updated, sheet="Cars", counts=None):
    L = LISTS[sheet]
    base = "../" if L["path"] else ""
    def sort_key(c):
        removed = manifest.get(c["key"], {}).get("status") == "removed"
        return (removed, c["price"] or 0)
    ordered = sorted(cars, key=sort_key)
    cards_html = "\n".join(card(c, manifest.get(c["key"]), i < 6, base) for i, c in enumerate(ordered))
    sold_ordered = sorted(sold_cars, key=lambda c: c["price"] or 0)
    sold_html = "\n".join(card(c, manifest.get(c["key"]), False, base) for c in sold_ordered)
    sold_block = ""
    if sold_ordered:
        sold_block = (f'<details class="sold-wrap"><summary>Sold / gone ({len(sold_ordered)})</summary>'
                      f'<ul class="grid">{sold_html}</ul></details>')
    n_new = sum(1 for c in cars if c["first_seen"] == BUILD_DAY)
    chips = [("new", f"New today ({n_new})"), ("all", "All"), ("toyota", "Toyota"), ("honda", "Honda"), ("mazda", "Mazda"),
             ("subaru", "Subaru"), ("hyundai-kia", "Hyundai/Kia"), (f"under:{L['under'][0]}", L["under"][1])]
    chip_html = "".join(f'<button class="chip{' new' if k == 'new' else ''}" type="button" data-f="{k}"{' id="newchip"' if k == "new" else ''} aria-pressed="{"true" if k == "all" else "false"}">{v}</button>'
                        for k, v in chips)
    seg = "".join(
        f'<a href="{base}{o["path"]}"{" aria-current=page" if o is L else ""}><b>{o["label"]}</b><span>{o["range"]} · {(counts or {}).get(sh, 0)} cars</span></a>'
        for sh, o in LISTS.items())
    main = f"""<nav class="seg" aria-label="Car lists">{seg}</nav>
<p class="criteria"><b>{L["label"]} list:</b> {esc(L["criteria"])}</p>
<div class="controls" data-js hidden>
<div class="chips" role="group" aria-label="Filter cars">{chip_html}</div>
<div class="row2"><span id="shown" aria-live="polite"></span>
<label>Sort <select id="sort"><option value="price">Price: low to high</option><option value="miles">Miles: low to high</option><option value="newest">Newest first</option><option value="added">Newest added</option></select></label></div>
</div>
<main>
<ul class="grid" id="cars">
{cards_html}
</ul>
<p class="empty" id="empty" hidden>No cars match this filter.</p>
<p class="empty" id="empty-new" hidden>Nothing new today yet. The daily search adds cars most mornings, so check back later.</p>
{sold_block}
</main>"""
    title = "Car Hunt" if sheet == "Cars" else f"{L['label']} list · Car Hunt"
    return shell(title, base, "cars", updated, len(cars), main, JS)


# ---------------- tips ----------------

HEADINGS = {"VERDICT (plain English)": "The verdict, in plain English", "THEFT ('Kia Boys')": "Theft (the “Kia Boys” problem)",
            "ENGINES": "Engines", "TRANSMISSIONS": "Transmissions", "TIMING BELT vs CHAIN": "Timing belt vs. chain",
            "WARRANTY (used)": "Warranty on a used car", "RELIABILITY": "Reliability", "SAFETY": "Safety"}


def scrub(t):
    # Keep the family's home ZIP off the public site.
    return re.sub(r"\s*\b74114\b\s*", " Tulsa ", t).replace("  ", " ").replace("( Tulsa ", "(Tulsa ").replace(" Tulsa )", " Tulsa)")


def numbered(text):
    items = [s.strip() for s in re.split(r"\s*\d\)\s+", text) if s.strip()]
    return "<ol>" + "".join(f"<li>{esc(i)}</li>" for i in items) + "</ol>"


def build_tips(xlsx, updated, count):
    notes = read_pairs(xlsx, "Hyundai-Kia Notes")
    sources = dict((a, b) for a, b, _ in read_pairs(xlsx, "Sources"))
    out = ['<main class="tips">', '<h2>Before you buy any car</h2>',
           '<div class="box"><ol>'
           '<li><strong>Check the VIN for recalls.</strong> Ask the seller for the VIN (the 17-character vehicle ID). Type it into <a href="https://www.nhtsa.gov/recalls" rel="noopener noreferrer nofollow">nhtsa.gov/recalls</a>. It’s free. Any open recall should be fixed by a dealer, at no cost, before you buy.</li>'
           '<li><strong>Get an independent inspection.</strong> Have a mechanic <em>you</em> choose (not the seller’s) look the car over before any money changes hands. A seller who won’t allow this is a red flag.</li>'
           '<li><strong>Pull a Carfax or AutoCheck report.</strong> It shows the title status, reported accidents, the number of owners and the odometer history. A “Verify title” badge on this site means we couldn’t confirm the title yet.</li>'
           '<li><strong>Get an insurance quote on that exact VIN.</strong> The insurance numbers on this site are estimates. Tell the insurer the car will be kept mostly in Fayetteville, because that can change the price.</li>'
           '</ol></div>']
    if sources.get("IMPORTANT"):
        out.append(f'<p class="small">{esc(scrub(sources["IMPORTANT"]))}</p>')
    out.append('<h2>Hyundai and Kia: what to know</h2>')
    in_verdict = False
    checklist = srcs = None
    for topic, detail, bold in notes:
        if not detail:  # section heading row
            if in_verdict:
                out.append("</div>")
            in_verdict = topic.startswith("VERDICT")
            out.append(f"<h2>{esc(HEADINGS.get(topic, topic.capitalize()))}</h2>")
            if in_verdict:
                out.append('<div class="box verdict">')
            continue
        if topic.upper().startswith("CHECKLIST"):
            checklist = detail
            continue
        if topic.upper() == "SOURCES":
            srcs = detail
            continue
        if topic.upper().startswith("INSURANCE"):
            if in_verdict:
                out.append("</div>")
                in_verdict = False
            out.append("<h2>Insurance</h2>")
            out.append(f"<p>{esc(scrub(detail))}</p>")
            continue
        out.append(f"<h3>{esc(scrub(topic))}</h3><p>{esc(scrub(detail))}</p>")
    if in_verdict:
        out.append("</div>")
    if checklist:
        out.insert(out.index('<h2>Hyundai and Kia: what to know</h2>') + 1,
                   f'<div class="box"><h3 style="margin-top:0">Checklist for any Hyundai or Kia</h3>{numbered(scrub(checklist))}</div>')
    if srcs:
        out.append(f'<h3>Sources</h3><p class="small">{esc(srcs)}</p>')
    out.append("</main>")
    return shell("Buying tips · Car Hunt", "../", "tips", updated, count, "\n".join(out))


# ---------------- model guide ----------------

BASE_74114 = 3419       # Sources sheet: CarInsurance.com OK add-teen cost x Insure.com 74114 factor
MONEYGEEK_AVG = 5654    # Sources sheet: average of MoneyGeek's 96-model 18-yo full-coverage table
VERDICTS = [("rec", "Recommended", "good", "Strong picks. Still check the specific car."),
            ("ok", "OK with conditions", "warn", "Fine if the condition in the note is met."),
            ("avoid", "Avoid", "bad", "Skip these, even when the price looks good.")]
MAKE_GROUPS = [("toyota", "Toyota"), ("honda", "Honda"), ("mazda", "Mazda"), ("subaru", "Subaru"),
               ("hyundai-kia", "Hyundai/Kia"), ("nissan", "Nissan"), ("ford", "Ford")]


def yr_ranges(pairs):
    """[(year, value)] -> 'value (2016–18), value2 (2019)' grouping consecutive years with the same value."""
    out, cur = [], None
    for y, v in sorted(pairs):
        if cur and cur[2] == v and y == cur[1] + 1:
            cur[1] = y
        else:
            cur = [y, y, v]
            out.append(cur)
    return out


def span(a, b):
    return str(a) if a == b else f"{a}–{str(b)[2:]}"


def ins_estimate(rows):
    """Annual 74114 estimate using the Sources-sheet method; None when we have no MoneyGeek figure."""
    t = next((r["insurance_teen_est"] for r in rows if "/yr" in r["insurance_teen_est"]), rows[0]["insurance_teen_est"])
    m = re.search(r"\$([\d,]+)/yr", t)
    if m:
        yr = int(m.group(1).replace(",", ""))
    else:
        m = re.search(r"\$([\d,]+) full", t)
        if not m:
            return None, ""
        yr = round(BASE_74114 * int(m.group(1).replace(",", "")) / MONEYGEEK_AVG)
    return yr, (" (Corolla figure)" if "proxy" in t else "")


def guide_metrics(rows):
    years = [int(r["year"]) for r in rows]
    single = len(set(years)) == 1
    def by_year(vals):
        groups = yr_ranges([(y, v) for y, v in vals if v])
        if len({g[2] for g in groups}) == 1 and len({y for y, v in vals if v}) == len(set(years)):
            return groups[0][2] if groups else ""
        return "; ".join(f"{g[2]} ({span(g[0], g[1])})" for g in groups)
    # reliability
    rel = []
    rp = {re.search(r"(\d\.\d)/5", r["reliability_rating"]).group(1) for r in rows if re.search(r"(\d\.\d)/5", r["reliability_rating"])}
    if rp:
        rank = re.search(r"(\d+(?:st|nd|rd|th) of \d+ [a-zA-Z ]+?)(?:[;)]|$)", rows[0]["reliability_rating"])
        rel.append(f"RepairPal {'/'.join(sorted(rp))}/5" + (f" ({rank.group(1).strip()})" if rank else ""))
    else:
        rel.append("RepairPal: no rating")
    cr = [(int(r["year"]), m.group(1).strip()) for r in rows
          if (m := re.search(r"Consumer Reports ([a-z ]+?)\s*(?:[;(]|$)", r["reliability_rating"])) and m.group(1).strip() != "n/a"]
    if cr:
        rel.append("Consumer Reports " + by_year(cr))
    jd = next((m.group(1) for r in rows if (m := re.search(r"J\.D\. Power 2026 VDS: ([^.]+)", r["known_issues"]))), None)
    if jd:
        rel.append(f"J.D. Power: {jd}")
    # safety
    def award(t):
        a = t.split(".")[0].strip()
        return {"TOP SAFETY PICK+": "Top Safety Pick+", "TOP SAFETY PICK": "Top Safety Pick"}.get(a, a.split("(")[0].strip())
    iihs = "IIHS " + by_year([(int(r["year"]), award(r["iihs"])) for r in rows])
    poor = [int(r["year"]) for r in rows if re.search(r"updated (moderate overlap|side) P\b", r["iihs"])]
    if poor:
        iihs += f"; Poor in an updated test ({', '.join(span(g[0], g[1]) for g in yr_ranges([(y, 1) for y in poor]))})"
    nh = by_year([(int(r["year"]), (r["nhtsa_overall"] + "★") if r["nhtsa_overall"].isdigit() else "not rated") for r in rows])
    safety = f"{iihs}. NHTSA {nh}."
    # repairs
    costs = [int(m.group(1)) for r in rows if (m := re.match(r"\$(\d+)", r["repair_cost_annual"]))]
    repair = (f"${min(costs)}–${max(costs)}/yr" if min(costs) != max(costs) else f"${costs[0]}/yr") if costs else "No RepairPal figure"
    # insurance
    yr, note = ins_estimate(rows)
    ins = f"${round(yr / 12)}/mo · ${yr:,}/yr{note}" if yr else "No estimate"
    # mpg
    mp = [int(m.group(1)) for r in rows if (m := re.match(r"(\d+)", r["mpg_combined"]))]
    mpg = (f"{min(mp)}–{max(mp)}" if min(mp) != max(mp) else str(mp[0])) + " combined" if mp else "n/a"
    return [("Reliability", "; ".join(rel)), ("Safety", safety), ("Repairs", repair),
            ("Insurance", ins), ("MPG", mpg)]


def guide_card(e, ref):
    v = {k: (label, cls) for k, label, cls, _ in VERDICTS}[e["verdict"]]
    mk = e["make"]
    grp = "hyundai-kia" if ("Hyundai" in mk or "Kia" in mk) else slugify(mk)
    tags = [f'<span class="b {v[1]}">{v[0]}</span>']
    if e.get("limited"):
        tags.append('<span class="b muted">Limited data</span>')
    rows = []
    if e.get("ref"):
        rows = [r for r in ref.get(f'{mk} {e["model"]}', []) if r["year"] in e["ref"]]
        if not rows:
            raise SystemExit(f"model guide: no reference rows for {mk} {e['model']} {e['ref']}")
    if rows:
        body = '<dl class="k">' + "".join(f"<dt>{esc(k)}</dt><dd>{esc(val)}</dd>" for k, val in guide_metrics(rows)) + "</dl>"
    else:
        body = ""
    note = e.get("metrics_note", "")
    if e.get("nometrics"):
        note = (note + " " if note else "") + ("Not in our research files, so no ratings, repair or insurance figures are shown."
                                               if e.get("limited") else "Ratings vary by model; see the matching entries above.")
    if note:
        body += f'<p class="note">{esc(note)}</p>'
    return (f'<li class="m {e["verdict"]}" data-v="{e["verdict"]}" data-g="{grp}">'
            f'<h3>{esc(mk)} {esc(e["model"])}<span class="yrs">{esc(e["years"])}</span></h3>'
            f'<div class="tags">{"".join(tags)}</div><p class="why">{esc(e["why"])}</p>{body}</li>')


GUIDE_JS = r"""<script>
(function(){
document.querySelectorAll('[data-js]').forEach(function(e){e.hidden=false});
var st={v:'all',g:'all'},items=[].slice.call(document.querySelectorAll('.m')),shown=document.getElementById('shown');
function apply(){var n=0;items.forEach(function(m){var ok=(st.v==='all'||m.dataset.v===st.v)&&(st.g==='all'||m.dataset.g===st.g);m.hidden=!ok;if(ok)n++;});
 document.querySelectorAll('.vgroup').forEach(function(s){s.hidden=!s.querySelector('.m:not([hidden])')});
 shown.textContent=n===items.length?('Showing all '+n+' entries'):('Showing '+n+' of '+items.length+' entries');
 document.getElementById('empty').hidden=n>0||state.f==='new';
 document.getElementById('empty-new').hidden=n>0||state.f!=='new';}
document.querySelectorAll('.chip').forEach(function(ch){ch.addEventListener('click',function(){var k=ch.dataset.k;st[k]=ch.dataset.f;
 document.querySelectorAll('.chip[data-k="'+k+'"]').forEach(function(x){x.setAttribute('aria-pressed',x===ch?'true':'false')});apply();});});
apply();
})();
</script>"""


def build_models(updated, count):
    import csv
    data = json.loads((ROOT / "guide" / "models.json").read_text(encoding="utf-8"))
    ref = {}
    with open(ROOT / "guide" / "model_reference.csv", newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            ref.setdefault(r["model_key"], []).append(r)
    entries = data["entries"]
    vchips = [("all", "All verdicts")] + [(k, label) for k, label, _, _ in VERDICTS]
    gchips = [("all", "All makes")] + MAKE_GROUPS
    chip = lambda k, f, label: (f'<button class="chip sm" type="button" data-k="{k}" data-f="{f}" '
                                f'aria-pressed="{"true" if f == "all" else "false"}">{esc(label)}</button>')
    out = [f'<p class="intro">{esc(data["intro"])}</p>',
           '<div class="controls" data-js hidden>',
           '<div class="chips" role="group" aria-label="Filter by verdict">' + "".join(chip("v", f, l) for f, l in vchips) + "</div>",
           '<div class="chips" role="group" aria-label="Filter by make">' + "".join(chip("g", f, l) for f, l in gchips) + "</div>",
           '<div class="row2"><span id="shown" aria-live="polite"></span></div></div>', "<main>"]
    for k, label, cls, lede in VERDICTS:
        group = [e for e in entries if e["verdict"] == k]
        out.append(f'<section class="vgroup" id="{k}"><h2><span class="dot {k}"></span>{esc(label)} '
                   f'<span class="small">({len(group)})</span></h2><p class="lede">{esc(lede)}</p><ul class="mlist">')
        out += [guide_card(e, ref) for e in group]
        out.append("</ul></section>")
    out.append('<p class="empty" id="empty" hidden>Nothing matches this filter.</p></main>')
    sources = ("<p>How to read it: year ranges are model years. Reliability, safety, repair and MPG figures cover the years listed. Where we checked only "
               "one year of a design, the note says which. Insurance is one model-level estimate; older cars usually cost a little less.</p>"
               "<p>Sources: reliability from RepairPal (model-level; its ratings cover 2010–2019 model years), Consumer Reports model-year verdicts and the "
               "J.D. Power 2026 Vehicle Dependability Study. Repair costs from RepairPal. Crash ratings from IIHS (G = Good, A = Acceptable, "
               "M = Marginal, P = Poor) and NHTSA 5-Star ratings; recalls from NHTSA and manufacturer notices. MPG is EPA combined, from fueleconomy.gov. "
               "Insurance estimates: CarInsurance.com’s Oklahoma cost of adding an 18-year-old to a parent’s full-coverage policy, adjusted to Tulsa "
               "with Insure.com and to each model with MoneyGeek’s 18-year-old rates. Models MoneyGeek doesn’t list show “No estimate”. "
               "Hyundai/Kia theft, engine and transmission notes come from Hyundai, Kia, NHTSA, IIHS/HLDI and AP reporting. Ford PowerShift years come from "
               "the Vargas v. Ford settlement and the later Barnes v. Ford suit.</p>")
    return shell("Model guide · Car Hunt", "../", "models", updated, count, "\n".join(out), GUIDE_JS, sources)


def main():
    if len(sys.argv) != 2:
        sys.exit("usage: build.py <xlsx>")
    xlsx = sys.argv[1]
    lists = read_all(xlsx)
    rules = avoid_rules()
    hidden = [(sh, c, avoid_match(c, rules)) for sh, cs in lists.items() for c in cs]
    hidden = [(sh, c, e) for sh, c, e in hidden if e]
    lists = {sh: [c for c in cs if not avoid_match(c, rules)] for sh, cs in lists.items()}
    for sh, c, e in hidden:
        print(f"hidden (Model Guide avoid: {e['make']} {e['model']} {e['years']}): {sh} {c['year']} {c['make']} {c['model']} {c['trim']} ${c['price']}")
    # Stamp every remaining row (including Notes-sold) so first_seen stays stable if a listing returns.
    stamped = [c for cs in lists.values() for c in cs]
    new_keys = stamp_first_seen(stamped)
    sold = {sh: [c for c in cs if is_sold(c)] for sh, cs in lists.items()}
    lists = {sh: [c for c in cs if not is_sold(c)] for sh, cs in lists.items()}
    for sh, cs in sold.items():
        for c in cs:
            print(f"sold/gone (Notes): {sh} {c['year']} {c['make']} {c['model']} {c['trim']} ${c['price']}")
    allcars = [c for cs in lists.values() for c in cs]
    mpath = ROOT / "photos" / "manifest.json"
    manifest = json.loads(mpath.read_text()) if mpath.exists() else {}
    updated = updated_time(xlsx)
    counts = {sh: len(cs) for sh, cs in lists.items()}
    n_sold = sum(len(cs) for cs in sold.values())
    for sheet, cars in lists.items():
        d = ROOT / LISTS[sheet]["path"]
        d.mkdir(exist_ok=True)
        (d / "index.html").write_text(build_index(cars, sold.get(sheet, []), manifest, updated, sheet, counts), encoding="utf-8")
    (ROOT / "tips").mkdir(exist_ok=True)
    (ROOT / "tips" / "index.html").write_text(build_tips(xlsx, updated, len(allcars)), encoding="utf-8")
    (ROOT / "models").mkdir(exist_ok=True)
    (ROOT / "models" / "index.html").write_text(build_models(updated, len(allcars)), encoding="utf-8")
    (ROOT / "robots.txt").write_text("User-agent: *\nDisallow: /\n")
    (ROOT / ".nojekyll").write_text("")
    with_photos = sum(1 for c in allcars if manifest.get(c["key"], {}).get("photos"))
    removed = sum(1 for c in allcars if manifest.get(c["key"], {}).get("status") == "removed")
    print(f"first seen today: {new_keys} new key(s)")
    print(f"built {' + '.join(f'{n} {sh}' for sh, n in counts.items())} cars; {with_photos} with photos; {removed} removed; {n_sold} sold/gone")


if __name__ == "__main__":
    main()
