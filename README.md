# car-hunt

Used-car shortlist, built from the master spreadsheet.

- `./update.sh /path/to/tulsa-used-cars.xlsx`: rebuild, fetch photos for new cars, commit, push
- `python3 build.py <xlsx>`: rebuild pages only
- `python3 fetch_photos.py <xlsx> [--recheck]`: fetch missing photos (`--recheck` re-tests every listing for removal)

Needs: python3 with openpyxl, pillow, requests, websocket-client; google-chrome and xvfb-run.
Photo state: `photos/manifest.json` (keyed by listing URL, or year|make|model|price without a URL).
