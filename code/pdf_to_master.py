"""
Fill Full Article Text on Master from manually saved article PDFs.

This is the manual half of the retrieval fallback archive.py describes:
rows it cannot resolve automatically ("needs_factiva" / "failed") get their
text pulled by hand. Saving the page as a PDF and running it through here
writes the same three cells archive.py would have written, so the row
becomes indistinguishable to everything downstream.

Usage
-----
    python3 code/pdf_to_master.py manifest.tsv
    python3 code/pdf_to_master.py --pdf <path> --url <original article URL>

manifest.tsv is two tab-separated columns per line: pdf_path, article_url.
Add --dry-run to see what would be written without touching the workbook.

Matching
--------
Rows are matched on the URL's host+path, ignoring scheme, "www." and any
trailing slash, because the URL you saved from is often a redirect variant
of the one stored in Master (e.g. ".../upper-hand-in-socal-blazes" vs
".../upper-hand-socal-blazes"). If that fails, the script falls back to an
exact title match, and refuses to write if a PDF matches zero rows or more
than one.

Cells written (mirroring archive.py)
------------------------------------
    col 23  Full Article Text   cleaned text, truncated to 32,000 chars
    col 24  Archive URL         only when a Wayback timestamp is recoverable
    col 25  Archive Status      "found"
    col 26  Word Count          len(text.split())

Status is written as "found" rather than a manual-specific value because
llm_coder.py selects exactly the rows whose status is "found"; anything
else would leave the row permanently uncoded. Provenance is still
recoverable from the Engines column and from this script's log.

Text handling
-------------
Ligatures and other compatibility characters are normalised (NFKC), since
PDF extraction yields "eﬀects"/"oﬃce" rather than "effects"/"office" and
those forms would otherwise reach the coder and any keyword matching.
Boilerplate lines are dropped conservatively -- only the Wayback banner,
social-share prompts and trailing "MORE STORIES" navigation. Headlines,
datelines, bylines and photo captions are kept, because that is what
trafilatura leaves in the rows archive.py already populated.
"""

import argparse
import re
import shutil
import sys
import unicodedata
from datetime import datetime
from urllib.parse import urlparse

import openpyxl
import pypdf

REPO = "/Users/hariaksha/Documents/GitHub/climate-disaster"
WORKBOOK = f"{REPO}/attribution.xlsx"

URL_COL, TITLE_COL = 12, 13
TEXT_COL, ARCHIVE_URL_COL, STATUS_COL, WORDCOUNT_COL = 23, 24, 25, 26
EXCEL_TEXT_LIMIT = 32000

BOILERPLATE = [
    re.compile(r"^The Wayback Machine\s*-\s*https?://", re.I),
    re.compile(r"^Don'?t miss the big stories\.?", re.I),
    re.compile(r"^Like\s?\d[\d.,]*[KM]?$", re.I),
    re.compile(r"^MORE STORIES\s*$", re.I),
    re.compile(r"^Share this article", re.I),
]
WAYBACK_TS = re.compile(r"web\.archive\.org/web/(\d{14})/")


def norm_url(u):
    p = urlparse(str(u).strip())
    return f"{p.netloc.lower().removeprefix('www.')}{p.path.rstrip('/')}".lower()


def extract(pdf_path):
    reader = pypdf.PdfReader(pdf_path)
    raw = "\n".join((pg.extract_text() or "") for pg in reader.pages)
    text = unicodedata.normalize("NFKC", raw)

    kept = []
    for line in text.split("\n"):
        s = line.strip()
        if not s:
            continue
        if any(p.search(s) for p in BOILERPLATE):
            continue
        kept.append(s)
    cleaned = "\n".join(kept).strip()

    # Wayback timestamp, from the banner or any embedded link annotation
    ts = None
    m = WAYBACK_TS.search(raw)
    if m:
        ts = m.group(1)
    else:
        for page in reader.pages:
            for annot in (page.get("/Annots") or []):
                try:
                    uri = str((annot.get_object().get("/A") or {}).get("/URI") or "")
                except Exception:
                    continue
                m = WAYBACK_TS.search(uri)
                if m:
                    ts = m.group(1)
                    break
            if ts:
                break
    return cleaned, ts


def find_row(ws, url, pdf_path):
    want = norm_url(url)
    hits = [r for r in range(2, ws.max_row + 1)
            if ws.cell(row=r, column=URL_COL).value
            and norm_url(ws.cell(row=r, column=URL_COL).value) == want]
    how = "url"
    if not hits:  # tolerate redirect slug drift
        stem = want.rstrip("/").rsplit("/", 1)[-1]
        tokens = [t for t in stem.split("-") if len(t) > 3]
        if tokens:
            hits = [r for r in range(2, ws.max_row + 1)
                    if ws.cell(row=r, column=URL_COL).value
                    and sum(t in norm_url(ws.cell(row=r, column=URL_COL).value)
                            for t in tokens) >= max(2, len(tokens) - 1)]
            how = "url-fuzzy"
    if not hits:
        title = pdf_path.rsplit("/", 1)[-1].rsplit(".pdf", 1)[0].strip()
        hits = [r for r in range(2, ws.max_row + 1)
                if str(ws.cell(row=r, column=TITLE_COL).value or "").strip() == title]
        how = "title"
    return hits, how


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("manifest", nargs="?")
    ap.add_argument("--pdf")
    ap.add_argument("--url")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    jobs = []
    if args.manifest:
        for line in open(args.manifest):
            line = line.strip()
            if line and not line.startswith("#"):
                p, u = line.split("\t")[:2]
                jobs.append((p.strip(), u.strip()))
    elif args.pdf and args.url:
        jobs.append((args.pdf, args.url))
    else:
        ap.error("provide a manifest, or both --pdf and --url")

    wb = openpyxl.load_workbook(WORKBOOK)
    ws = wb["Master"]
    planned, problems = [], []

    for pdf_path, url in jobs:
        name = pdf_path.rsplit("/", 1)[-1]
        rows, how = find_row(ws, url, pdf_path)
        if len(rows) != 1:
            problems.append((name, f"matched {len(rows)} rows"))
            print(f"  SKIP {name}: matched {len(rows)} rows")
            continue
        r = rows[0]
        # Guard on text, not on the status label. A row marked 'failed' or
        # 'needs_factiva' has no text and is exactly what a hand-pulled PDF is
        # for; refusing it on the strength of having *a* status blocked the
        # rows this script exists to fill. Only an existing body of text is a
        # real reason not to write.
        existing_text = ws.cell(row=r, column=TEXT_COL).value
        if existing_text is not None and str(existing_text).strip():
            problems.append((name, f"row {r} already has text"))
            print(f"  SKIP {name}: row {r} already has text")
            continue

        text, ts = extract(pdf_path)
        words = len(text.split())
        archive_url = (f"https://web.archive.org/web/{ts}/{url}" if ts else None)
        planned.append((r, name, text, words, archive_url, how))
        print(f"  row {r:5d}  {words:5d} words  match={how:9s}  "
              f"archive_ts={'yes' if ts else 'no ':3s}  {name[:52]}")

    if args.dry_run:
        print(f"\n[DRY RUN] {len(planned)} row(s) would be written; nothing saved.")
        return

    if not planned:
        print("\nNothing to write.")
        return

    backup = f"{REPO}/attribution_backup_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(WORKBOOK, backup)

    for r, name, text, words, archive_url, _ in planned:
        ws.cell(row=r, column=TEXT_COL).value = text[:EXCEL_TEXT_LIMIT]
        ws.cell(row=r, column=STATUS_COL).value = "found"
        ws.cell(row=r, column=WORDCOUNT_COL).value = words
        if archive_url:
            ws.cell(row=r, column=ARCHIVE_URL_COL).value = archive_url

    wb.save(WORKBOOK)
    print(f"\nWrote {len(planned)} row(s). Backup: {backup.rsplit('/',1)[-1]}")
    if problems:
        print(f"{len(problems)} skipped:")
        for n, why in problems:
            print(f"   {n[:60]}: {why}")


if __name__ == "__main__":
    main()
