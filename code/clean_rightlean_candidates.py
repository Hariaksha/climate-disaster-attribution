"""
Clean the right-leaning collection run before it is appended to the Master
sheet of attribution.xlsx.

Input
-----
  data/rightlean_articles_long.csv      raw output of google_search_rightlean.py

Outputs
-------
  data/rightlean_articles_clean.csv           append-ready article rows
  data/rightlean_domains_needing_ratings.csv  domains to look up on
                                              Ad Fontes / AllSides before
                                              the lean analysis is rerun

The raw input file is never modified, so the unfiltered collection output
stays available as a provenance record.

Cleanup steps
-------------
1. Drop sponsored/advertorial subdomains. deloitte.wsj.com is Deloitte-branded
   content hosted on WSJ's platform, not WSJ editorial. google_search.py blocks
   "/sponsored/" and "/advertorial/" URL paths but has no subdomain rule, so
   these slip through.

2. Rewrite Epoch Times CMS URLs. Four articles came back on
   cmsapi.theepochtimes.com, which is the CMS/API backend rather than the
   public article URL, and would likely fail in archive.py. The path component
   is the same, so the host is rewritten to the canonical domain. Flagged in
   the output so the rewrite can be verified during archiving.

3. Drop rows already in Master for the SAME event. These are articles the
   existing corpus already has and has already coded; appending them would
   create true duplicate rows.

   Rows whose URL is in Master under a DIFFERENT event are kept: one article
   can legitimately cover two overlapping events (e.g. a single piece on
   Hurricanes Ivan and Jeanne), and Master already contains 17 such
   cross-event duplicates, so keeping them matches existing practice.

4. Report domains with no Ad Fontes Bias score in the Domain Ratings sheet.
   These would otherwise be coded "Unidentified" in the lean analysis, which
   would defeat the purpose of collecting them -- they are the further-right
   outlets the supplementary run was designed to capture. Nothing is dropped
   for this; it is a checklist.
"""

import re
from urllib.parse import urlparse, urlunparse

import openpyxl
import pandas as pd

REPO = "/Users/hariaksha/Documents/GitHub/climate-disaster"
INPUT_LONG = f"{REPO}/data/rightlean_articles_long.csv"
OUTPUT_CLEAN = f"{REPO}/data/rightlean_articles_clean.csv"
OUTPUT_RATINGS_TODO = f"{REPO}/data/rightlean_domains_needing_ratings.csv"
WORKBOOK = f"{REPO}/attribution.xlsx"

# Subdomains that host branded/sponsored content rather than editorial.
ADVERTORIAL_HOSTS = {"deloitte.wsj.com"}

# Non-canonical hosts -> canonical host (path is preserved).
HOST_REWRITES = {"cmsapi.theepochtimes.com": "www.theepochtimes.com"}


def get_domain(url):
    netloc = urlparse(str(url)).netloc.lower()
    return netloc.removeprefix("www.")


def normalize(url):
    """Host + path, lowercased, no www, no trailing slash -- for matching
    against Master, which stores URLs in varying forms."""
    p = urlparse(str(url))
    return f"{p.netloc.lower().removeprefix('www.')}{p.path.rstrip('/')}"


def rewrite_host(url):
    p = urlparse(str(url))
    new_host = HOST_REWRITES.get(p.netloc.lower())
    if not new_host:
        return url, False
    return urlunparse(p._replace(netloc=new_host)), True


def load_master_keys():
    """Return (set of normalized URLs, set of (normalized URL, event name))."""
    wb = openpyxl.load_workbook(WORKBOOK, read_only=True, data_only=True)
    ws = wb["Master"]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    idx = {h: i for i, h in enumerate(header)}
    urls, pairs = set(), set()
    for r in rows:
        u = r[idx["URL"]]
        if not u:
            continue
        n = normalize(u)
        urls.add(n)
        pairs.add((n, str(r[idx["Name"]]).strip()))
    return urls, pairs


def load_rated_domains():
    """Domains that already have a numeric Ad Fontes Bias score."""
    wb = openpyxl.load_workbook(WORKBOOK, read_only=True, data_only=True)
    ws = wb["Domain Ratings"]
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    idx = {h: i for i, h in enumerate(header)}
    rated = set()
    for r in rows:
        d, af = r[idx["Domain"]], r[idx["Ad Fontes Rating"]]
        if d and af and re.search(r"Bias ([+-]?[0-9.]+)", af):
            rated.add(d)
    return rated


def main():
    df = pd.read_csv(INPUT_LONG)
    start_n = len(df)
    print(f"Loaded {start_n} raw article rows from rightlean_articles_long.csv\n")

    # --- Step 1: drop advertorial subdomains -----------------------------
    df["_host"] = df["URL"].map(lambda u: urlparse(str(u)).netloc.lower())
    advert = df["_host"].isin(ADVERTORIAL_HOSTS)
    if advert.any():
        print(f"[1] Dropping {advert.sum()} sponsored/advertorial rows:")
        for _, r in df[advert].iterrows():
            print(f"      {r['_host']} | {str(r['Title'])[:58]}")
    else:
        print("[1] No advertorial rows found")
    df = df[~advert].copy()

    # --- Step 2: rewrite non-canonical hosts ------------------------------
    rewritten = df["URL"].map(lambda u: rewrite_host(u)[1])
    df["URL Rewritten"] = rewritten
    df["URL"] = df["URL"].map(lambda u: rewrite_host(u)[0])
    if rewritten.any():
        print(f"\n[2] Rewrote {rewritten.sum()} non-canonical host(s) "
              f"({', '.join(HOST_REWRITES)}) -> canonical domain")
        print("    (verify these resolve when archive.py runs)")
    else:
        print("\n[2] No host rewrites needed")

    # --- Step 3: drop rows already in Master for the same event -----------
    master_urls, master_pairs = load_master_keys()
    df["_norm"] = df["URL"].map(normalize)
    df["_event"] = df["Name"].astype(str).str.strip()

    same_event_dup = df.apply(lambda r: (r["_norm"], r["_event"]) in master_pairs, axis=1)
    cross_event_dup = df["_norm"].isin(master_urls) & ~same_event_dup

    print(f"\n[3] Already in Master, same event : {same_event_dup.sum()} -> dropped")
    print(f"    In Master under a different event: {cross_event_dup.sum()} -> kept "
          f"(matches Master's existing cross-event duplicates)")
    df = df[~same_event_dup].copy()

    # Within-run duplicates across events are also legitimate; report only.
    within = df[df.duplicated("_norm", keep=False)]
    if len(within):
        print(f"    Same URL under 2+ events within this run: "
              f"{within['_norm'].nunique()} URL(s), {len(within)} rows -> kept")

    # --- Step 4: domains needing a lean rating ----------------------------
    df["Domain"] = df["URL"].map(get_domain)
    rated = load_rated_domains()
    counts = df["Domain"].value_counts()
    todo = [(d, int(n)) for d, n in counts.items() if d not in rated]

    print(f"\n[4] Domains with no Ad Fontes Bias score on file: {len(todo)}")
    if todo:
        affected = sum(n for _, n in todo)
        print(f"    These cover {affected} articles "
              f"({100 * affected / len(df):.1f}% of the cleaned set).")
        print(f"    Without ratings they will code as 'Unidentified'.\n")
        print(f"    {'Domain':30s} {'Articles':>8s}")
        for d, n in todo:
            print(f"    {d:30s} {n:8d}")
        pd.DataFrame(
            [{"Domain": d, "Articles": n,
              "Ad Fontes Rating": "", "AllSides Rating": "", "Notes": ""}
             for d, n in todo]
        ).to_csv(OUTPUT_RATINGS_TODO, index=False)
        print(f"\n    Checklist written to "
              f"{OUTPUT_RATINGS_TODO.rsplit('/', 1)[-1]} "
              f"(fill in the rating columns, then merge into 'Domain Ratings')")

    # --- Write cleaned output ---------------------------------------------
    df = df.drop(columns=["_host", "_norm", "_event"])
    df.to_csv(OUTPUT_CLEAN, index=False)

    print(f"\n=== SUMMARY ===")
    print(f"Raw rows in          : {start_n}")
    print(f"Clean rows out       : {len(df)}  (removed {start_n - len(df)})")
    print(f"Distinct URLs        : {df['URL'].nunique()}")
    print(f"Events covered       : {df['Name'].nunique()}")
    print(f"Wrote {OUTPUT_CLEAN.rsplit('/', 1)[-1]}")


if __name__ == "__main__":
    main()
