#!/usr/bin/env bash
# Usage: ./update.sh /path/to/tulsa-used-cars.xlsx
# Rebuilds the site from the spreadsheet, fetches photos only for cars that have none yet,
# commits "cars: update <date>" and pushes to main. Safe to re-run: no changes -> no commit.
set -euo pipefail
if [ $# -ne 1 ]; then echo "usage: $0 path/to/tulsa-used-cars.xlsx" >&2; exit 2; fi
XLSX="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
[ -s "$XLSX" ] || { echo "missing spreadsheet: $XLSX" >&2; exit 2; }
cd "$(dirname "$0")"
git pull --ff-only --quiet origin main || true
python3 fetch_photos.py "$XLSX"
python3 build.py "$XLSX"
git add -A index.html budget tips models guide broken-arrow south-tulsa midtown robots.txt .nojekyll photos assets build.py fetch_photos.py update.sh README.md .gitignore 2>/dev/null || git add -A
if git diff --cached --quiet; then
  echo "no changes; nothing to commit"
else
  git commit --quiet -m "cars: update $(TZ=America/Chicago date +%Y-%m-%d)"
fi
git push --quiet origin HEAD:main
echo "live: https://vnqqsmhyvm-a11y.github.io/car-hunt/"
