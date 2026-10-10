#!/usr/bin/env python3
"""Fetch real listing photos for cars that don't have any yet.

State lives in photos/manifest.json, keyed by listing URL (or year|make|model|price
when a car has no URL):
  {"<key>": {"slug": ..., "photos": ["photos/<slug>-1.jpg", ...],
             "status": "ok" | "removed" | "unavailable", "source": ..., "checked": ...}}

- Facebook Marketplace: logged-out item page with an iPhone user-agent, og:image only.
  The og:title must name the same year + make, otherwise it's not used.
- Cars.com and other sites: loaded in a real (non-headless) Chrome on a private Xvfb
  display via the DevTools protocol, then gallery / JSON-LD / og:image URLs are read.
  Generic site images (logos, placeholders) are never used.
- Never signs in, never contacts sellers. Images are saved locally (never hotlinked),
  resized to <=1200px wide, JPEG q80.

Usage: python3 fetch_photos.py <xlsx> [--recheck]   (--recheck also re-tests 'ok' cars for removal)
"""
import datetime as dt
import html
import io
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

import requests
from PIL import Image

ROOT = Path(__file__).resolve().parent
PHOTOS = ROOT / "photos"
MANIFEST = PHOTOS / "manifest.json"
MAX_PHOTOS = 3
IPHONE_UA = ("Mozilla/5.0 (iPhone; CPU iPhone OS 17_6 like Mac OS X) AppleWebKit/605.1.15 "
             "(KHTML, like Gecko) Version/17.6 Mobile/15E148 Safari/604.1")
DESKTOP_UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36")
NAV_HEADERS = {"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
               "Accept-Language": "en-US,en;q=0.9", "Sec-Fetch-Dest": "document",
               "Sec-Fetch-Mode": "navigate", "Sec-Fetch-Site": "none", "Upgrade-Insecure-Requests": "1"}
CDP_PORT = 9333
GENERIC_IMG = re.compile(r"logo|placeholder|no[-_]?image|default|Cars-PD_OG|og-img|coming[-_]?soon|sprite|icon", re.I)

sys.path.insert(0, str(ROOT))
import build  # noqa: E402  (reuses the spreadsheet reader and slug/key helpers)


def log(*a):
    print(*a, flush=True)


def load_manifest():
    if MANIFEST.exists():
        return json.loads(MANIFEST.read_text())
    return {}


def save_manifest(m):
    MANIFEST.write_text(json.dumps(m, indent=1, sort_keys=True) + "\n")


def save_image(data, dest):
    im = Image.open(io.BytesIO(data))
    im = im.convert("RGB")
    if im.width < 300:
        raise ValueError(f"image too small ({im.width}px)")
    if im.width > 1200:
        im = im.resize((1200, round(im.height * 1200 / im.width)), Image.LANCZOS)
    dest.parent.mkdir(parents=True, exist_ok=True)
    im.save(dest, "JPEG", quality=80, optimize=True, progressive=True)


def download(url, referer=None):
    h = {"User-Agent": DESKTOP_UA, "Accept": "image/avif,image/webp,image/*,*/*;q=0.8"}
    if referer:
        h["Referer"] = referer
    r = requests.get(url, headers=h, timeout=30)
    r.raise_for_status()
    if not r.headers.get("content-type", "").startswith("image/"):
        raise ValueError(f"not an image: {r.headers.get('content-type')}")
    return r.content


def title_matches(title, car):
    t = (title or "").lower()
    return str(car["year"]) in t and car["make"].lower() in t and car["model"].split()[0].lower() in t


# ---------- Facebook ----------

def fetch_facebook(car):
    r = requests.get(car["url"], headers={"User-Agent": IPHONE_UA, **NAV_HEADERS}, timeout=30, allow_redirects=True)
    if r.status_code == 404:
        return "removed", [], "facebook 404"
    body = r.text
    t = re.search(r'<meta property="og:title" content="([^"]*)"', body)
    title = html.unescape(t.group(1)) if t else ""
    if "/login" in r.url and not t:
        return "unavailable", [], "facebook login wall"
    if not title_matches(title, car):
        # Logged-out item URL resolving to a generic Marketplace page = listing gone.
        if title.startswith("Facebook Marketplace") or "marketplace/item" not in r.url:
            return "removed", [], f"facebook generic page ({title[:60]})"
        return "unavailable", [], f"facebook title mismatch ({title[:60]})"
    imgs = [html.unescape(m) for m in re.findall(r'<meta property="og:image" content="([^"]+)"', body)]
    imgs = [u for u in imgs if "scontent" in u or "fbcdn" in u]
    return "ok", imgs[:1], f"facebook og:image ({title[:60]})"


# ---------- Chrome (DevTools) for everything else ----------

class Chrome:
    def __init__(self):
        self.proc = None

    def up(self):
        try:
            requests.get(f"http://localhost:{CDP_PORT}/json/version", timeout=2)
            return True
        except Exception:
            return False

    def start(self):
        if self.up():
            return
        prof = Path("/tmp/car-hunt-chrome")
        prof.mkdir(exist_ok=True)
        cmd = ["xvfb-run", "-a", "-s", "-screen 0 1366x900x24", "google-chrome", "--no-sandbox", "--disable-gpu",
               f"--remote-debugging-port={CDP_PORT}", f"--user-data-dir={prof}", "--no-first-run",
               "--no-default-browser-check", "--window-size=1366,900", "about:blank"]
        self.proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
        for _ in range(30):
            time.sleep(1)
            if self.up():
                return
        raise RuntimeError("Chrome did not start")

    def stop(self):
        if self.proc:
            os.killpg(self.proc.pid, 15)
            self.proc = None

    def dom(self, url, wait=40):
        import websocket
        self.start()
        t = requests.put(f"http://localhost:{CDP_PORT}/json/new?{url}", timeout=10).json()
        ws = websocket.create_connection(t["webSocketDebuggerUrl"], timeout=30, suppress_origin=True)
        n = [0]

        def call(method, **params):
            n[0] += 1
            ws.send(json.dumps({"id": n[0], "method": method, "params": params}))
            while True:
                msg = json.loads(ws.recv())
                if msg.get("id") == n[0]:
                    return msg.get("result", {})
        end = time.time() + wait
        title = body = href = ""
        try:
            while time.time() < end:
                time.sleep(2.5)
                v = call("Runtime.evaluate", returnByValue=True,
                         expression="document.title+'\\n@@\\n'+location.href+'\\n@@\\n'+document.documentElement.outerHTML")
                v = v.get("result", {}).get("value", "") or ""
                title, href, body = (v.split("\n@@\n", 2) + ["", ""])[:3]
                if title and "moment" not in title.lower() and len(body) > 30000:
                    time.sleep(2)
                    v = call("Runtime.evaluate", returnByValue=True,
                             expression="document.title+'\\n@@\\n'+location.href+'\\n@@\\n'+document.documentElement.outerHTML")
                    title, href, body = (v["result"]["value"].split("\n@@\n", 2) + ["", ""])[:3]
                    break
        finally:
            ws.close()
            requests.get(f"http://localhost:{CDP_PORT}/json/close/{t['id']}", timeout=10)
        return title, href, body


CHROME = Chrome()


def jsonld_images(body):
    out = []
    for blob in re.findall(r'<script[^>]+application/ld\+json[^>]*>(.*?)</script>', body, re.S):
        try:
            data = json.loads(html.unescape(blob.strip()))
        except Exception:
            continue
        stack = [data]
        while stack:
            d = stack.pop()
            if isinstance(d, list):
                stack.extend(d)
            elif isinstance(d, dict):
                img = d.get("image")
                if isinstance(img, str):
                    out.append(img)
                elif isinstance(img, list):
                    out += [i if isinstance(i, str) else i.get("url", "") for i in img]
                elif isinstance(img, dict):
                    out.append(img.get("url", ""))
                stack.extend(v for v in d.values() if isinstance(v, (dict, list)))
    return [u for u in out if u]


def fetch_carscom(car):
    title, href, body = CHROME.dom(car["url"])
    low = body.lower()
    if "moment" in title.lower():
        return "unavailable", [], "cars.com bot check did not clear"
    if ("no longer available" in low or "vehicle is unavailable" in low or "/vehicledetail/" not in href
            or "this vehicle has been sold" in low):
        return "removed", [], f"cars.com: listing gone ({title[:60]})"
    if not title_matches(title, car):
        return "unavailable", [], f"cars.com title mismatch ({title[:60]})"
    urls = []
    for m in re.finditer(r'https://platform\.cstatic-images\.com/[a-z]+/in/v2/([0-9a-f-]+/[0-9a-f-]+)/[A-Za-z0-9_-]+\.(?:jpg|jpeg|png|webp)', body):
        urls.append((m.group(1), m.group(0)))
    if not urls:
        return "unavailable", [], "cars.com: no gallery photos on page"
    main_vehicle = urls[0][0]  # the main gallery comes first; ignore other cars' thumbnails
    seen = []
    for vid, u in urls:
        u = re.sub(r"/(?:[a-z]+)/in/v2/", "/xxlarge/in/v2/", u)
        if vid == main_vehicle and u not in seen:
            seen.append(u)
    return "ok", seen, f"cars.com gallery ({title[:60]})"


def fetch_generic(car):
    title, href, body = CHROME.dom(car["url"], wait=30)
    low = body.lower()
    if not body or "404" in title or "not found" in title.lower() or "sold" in title.lower():
        return "removed", [], f"dealer page gone ({title[:60]})"
    if not title_matches(title, car):
        return "unavailable", [], f"dealer title mismatch ({title[:60]})"
    cands = jsonld_images(body)
    cands += [html.unescape(u) for u in re.findall(r'<meta[^>]+property="og:image"[^>]+content="([^"]+)"', body)]
    # dealer galleries: pick large inventory images that mention the stock/VIN folder
    cands += re.findall(r'https://imagescf\.dealercenter\.net/[^"\'\s)]+\.jpg', body)
    out = []
    for u in cands:
        u = re.sub(r"imagescf\.dealercenter\.net/\d+/\d+/", "imagescf.dealercenter.net/1280/960/", u)
        if u.startswith("http") and "{" not in u and "%7B" not in u and not GENERIC_IMG.search(u) and u not in out:
            out.append(u)
    if not out:
        return "unavailable", [], "dealer page: no vehicle images found"
    return "ok", out, f"dealer page ({title[:60]})"


def fetch_one(car):
    u = car["url"] or ""
    if not u:
        return "unavailable", [], "no listing link"
    if "facebook.com" in u:
        return fetch_facebook(car)
    if "cars.com" in u:
        return fetch_carscom(car)
    return fetch_generic(car)


def next_name(slug, used):
    i = 1
    while f"photos/{slug}-{i}.jpg" in used or (PHOTOS / f"{slug}-{i}.jpg").exists():
        i += 1
    return f"photos/{slug}-{i}.jpg"


def fill(entry, referer):
    """Download candidates (skipping rejected and already-saved ones) until MAX_PHOTOS are saved."""
    photos = [p for p in entry.get("photos", []) if (ROOT / p).exists()]
    sources = {p: u for p, u in entry.get("sources", {}).items() if p in photos}
    tried = set(sources.values()) | set(entry.get("rejected", []))
    attempts = 0
    for u in entry.get("candidates", []):
        if len(photos) >= MAX_PHOTOS or attempts >= 8:
            break
        if u in tried:
            continue
        attempts += 1
        rel = next_name(entry["slug"], set(photos))
        try:
            save_image(download(u, referer), ROOT / rel)
            photos.append(rel)
            sources[rel] = u
        except Exception as ex:
            log(f"    photo failed: {ex}")
    entry["photos"], entry["sources"] = photos, sources
    return entry


def drop(paths, man):
    """Reject specific saved photos (e.g. dealer promo graphics) and refill from the same listing."""
    for p in paths:
        p = str(Path(p)) if not str(p).startswith("photos/") else p
        for key, e in man.items():
            if p in e.get("photos", []):
                src = e.get("sources", {}).pop(p, None)
                e["photos"].remove(p)
                if src:
                    e.setdefault("rejected", []).append(src)
                (ROOT / p).unlink(missing_ok=True)
                log(f"  dropped {p}")
                fill(e, e.get("url"))
                log(f"  now {len(e['photos'])} photo(s): {', '.join(e['photos'])}")
                break
        else:
            log(f"  not in manifest: {p}")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if "--drop" in sys.argv:
        man = load_manifest()
        drop(args, man)
        save_manifest(man)
        return
    if not args:
        sys.exit("usage: fetch_photos.py <xlsx> [--recheck] | fetch_photos.py --drop photos/<file>.jpg ...")
    recheck = "--recheck" in sys.argv
    cars = [c for cs in build.read_all(args[0]).values() for c in cs]
    man = load_manifest()
    now = dt.datetime.now().astimezone().isoformat(timespec="seconds")
    todo = []
    for c in cars:
        e = man.get(c["key"])
        have = e and e.get("photos") and all((ROOT / p).exists() for p in e["photos"])
        if (have or (e and e.get("status") == "removed")) and not recheck:
            continue
        todo.append(c)
    log(f"{len(cars)} cars; {len(todo)} need photos/checks")
    try:
        for c in todo:
            e = man.get(c["key"], {})
            label = f'{c["year"]} {c["make"]} {c["model"]} ${c["price"]}'
            try:
                status, urls, source = fetch_one(c)
            except Exception as ex:  # network hiccup: leave for next run
                log(f"  ERROR {label}: {ex}")
                continue
            e.update({"slug": e.get("slug") or c["slug"], "status": status, "source": source,
                      "checked": now, "url": c["url"]})
            if urls:
                e["candidates"] = urls
            e.setdefault("photos", [])
            if status == "ok" and not [p for p in e["photos"] if (ROOT / p).exists()]:
                fill(e, c["url"])
                if not e["photos"]:
                    e["status"], e["source"] = "unavailable", source + "; image download failed"
            man[c["key"]] = e
            log(f"  {e['status']:11} {len(e['photos'])} photo(s)  {label}  [{source}]")
            save_manifest(man)
            time.sleep(1.5)
    finally:
        CHROME.stop()
    save_manifest(man)


if __name__ == "__main__":
    main()
