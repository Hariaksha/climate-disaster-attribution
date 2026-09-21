"""
Supplementary article collection: right-leaning outlets only.

Background
----------
The main collection run (google_search.py) retrieved very few articles from
right-leaning outlets: only 20 of the 1,252 Relevant articles came from a
domain with a positive Ad Fontes Bias score. Two causes were identified:

  1. Ranking competition inside a shared Programmable Search Engine. The
     "National" PSE spans ~47 domains, but the Google Custom Search API
     returns only the top 10 organic results per query, so wire-syndicated
     coverage republished across the NYT/WaPo/CNN/NPR/USA Today/AP tends to
     fill every slot. foxnews.com returned 0 candidates and wsj.com returned
     1 across all 313 events, despite both being eligible domains queried on
     every event.

  2. Domain-list gaps. Several right-leaning outlets that ARE in the Access
     World News database with verified URLs (Washington Examiner, Washington
     Times, New York Sun) never made it into pse_domains.json, and
     digital-native outlets (Breitbart, National Review, Daily Wire, Daily
     Caller) were never in the AWN newspaper database at all.

This script addresses both by querying a dedicated PSE containing only
right-leaning outlets, so those domains compete against each other for the
10 available result slots rather than against the largest national papers.

What it does / does not change
------------------------------
Query construction, date-window filtering, geographic relevance, disaster
keyword matching, and de-duplication are all imported unchanged from
google_search.py, so candidates collected here are filtered by exactly the
same rules as the rest of the corpus. The only deliberate difference is
MAX_URLS_PER_ENGINE (see below): with 19 engines the original run capped
each (query, engine) pair at 5 results to keep any one engine from
dominating a disaster's quota, but here there is only one engine, so that
cap would throw away results for no reason.

Outputs (all new files -- nothing existing is overwritten)
----------------------------------------------------------
  - data/rightlean_article_urls.csv       wide format, one row per event,
                                          matching climate_disaster_article_urls.csv
  - data/rightlean_articles_long.csv      long format, one row per article,
                                          convenient for appending to the
                                          Master sheet
  - data/rightlean_checkpoint.csv         resume checkpoint

Next steps after running this (not done here)
---------------------------------------------
  1. Append the new candidates to the Master sheet of attribution.xlsx
  2. Run archive.py to retrieve full article text
  3. Run llm_coder.py to apply the same two-step coding procedure
"""

import argparse
import os
import time

import pandas as pd

import google_search as gs


# The dedicated right-leaning PSE (created separately in Google's
# Programmable Search Engine console).
RIGHTLEAN_CX = "150159cdb4eb04fb2"

EVENTS_FILE = "/Users/hariaksha/Documents/GitHub/climate-disaster/data/events-US-2000-2024-Q4-states.csv"
OUTPUT_WIDE = "/Users/hariaksha/Documents/GitHub/climate-disaster/data/rightlean_article_urls.csv"
OUTPUT_LONG = "/Users/hariaksha/Documents/GitHub/climate-disaster/data/rightlean_articles_long.csv"
CHECKPOINT_FILE = "/Users/hariaksha/Documents/GitHub/climate-disaster/data/rightlean_checkpoint.csv"

# With 19 engines, capping each (query, engine) pair at 5 kept any single
# engine from eating a disaster's whole 30-article quota. Here there is one
# engine, so the cap only discards usable results -- raise it to the API's
# own per-call maximum.
gs.MAX_URLS_PER_ENGINE = 10

MAX_URLS_PER_DISASTER = gs.MAX_URLS_PER_DISASTER

TEST_RANDOM_STATE = 44


def collect_for_event(row, next_event_begin_dt):
    """Run every query for one event against the right-leaning PSE."""
    queries, begin_dt, end_dt, core_name, region_text = gs.build_queries(row)

    states_field = str(row.get("States", "")).strip().upper()
    if not states_field or states_field in ("NATIONAL", "NAN"):
        n_states = None
    else:
        n_states = len([s for s in states_field.split(",") if s.strip()])

    disaster_type = str(row["Disaster"]).strip()
    candidates = []

    for query in queries:
        if len(candidates) >= MAX_URLS_PER_DISASTER:
            break

        results, err = gs.search_one_engine(
            query=query,
            cx=RIGHTLEAN_CX,
            begin_dt=begin_dt,
            end_dt=end_dt,
            disaster_type=disaster_type,
            core_name=core_name,
            num=gs.RESULTS_PER_ENGINE,
            n_states=n_states,
            next_event_begin_dt=next_event_begin_dt,
        )

        if err == "RATE_LIMITED":
            print("  [429] Rate limited — sleeping 60s before retrying…")
            time.sleep(60)
            results, err = gs.search_one_engine(
                query=query,
                cx=RIGHTLEAN_CX,
                begin_dt=begin_dt,
                end_dt=end_dt,
                disaster_type=disaster_type,
                core_name=core_name,
                num=gs.RESULTS_PER_ENGINE,
                n_states=n_states,
                next_event_begin_dt=next_event_begin_dt,
            )
            if err == "RATE_LIMITED":
                raise gs.DailyQuotaExhausted()

        if err:
            print(f"  [API error] query='{query[:50]}': {err}")
        elif results:
            candidates.extend(results)

        time.sleep(gs.SLEEP_BETWEEN_CALLS)

    return gs.dedupe_results(candidates)[:MAX_URLS_PER_DISASTER]


def main(test_n=None):
    df = pd.read_csv(EVENTS_FILE)
    print(f"Loaded {len(df)} events from {os.path.basename(EVENTS_FILE)}")

    # Same "don't let one event's grace period bleed into the next event of
    # the same type" logic as the main run, computed on the full dataset.
    next_event_begin_by_name = {}
    for _dtype, grp in df.groupby("Disaster"):
        begins = sorted(gs.parse_yyyymmdd(b) for b in grp["Begin Date"])
        for _, r in grp.iterrows():
            end_dt = gs.parse_yyyymmdd(r["End Date"])
            later = [b for b in begins if b > end_dt]
            next_event_begin_by_name[str(r["Name"]).strip()] = min(later) if later else None

    if test_n:
        # Use the costliest events, which is a good proxy for "nationally
        # covered". A smoke test needs events these outlets plausibly wrote
        # about at all -- sampling at random (or even just recent) lands on
        # mid-tier regional storms that national outlets legitimately ignored,
        # so a zero result would say nothing about whether the PSE works.
        df = (df.sort_values("CPI-Adjusted Cost", ascending=False)
                .head(test_n).reset_index(drop=True))
        print(f"[SMOKE TEST] Using the {len(df)} costliest events — "
              f"~{len(df) * 4} API calls, no checkpoint written")
    else:
        print(f"[PRODUCTION] Running all {len(df)} events")

    # Resume from checkpoint if a previous run stopped early.
    wide_rows = []
    if os.path.exists(CHECKPOINT_FILE) and not test_n:
        checkpoint_df = pd.read_csv(CHECKPOINT_FILE)
        wide_rows = checkpoint_df.to_dict("records")
        done = set(checkpoint_df["Name"].astype(str).str.strip())
        before = len(df)
        df = df[~df["Name"].astype(str).str.strip().isin(done)].reset_index(drop=True)
        print(f"[Resume] {len(done)} events already done — skipping {before - len(df)}, {len(df)} remaining")

    total = len(df)
    total_articles = 0
    events_with_articles = 0
    interrupted = False

    try:
        for pos, idx in enumerate(df.index, start=1):
            row = df.loc[idx]
            name = str(row["Name"]).strip()
            print(f"\n=== [{pos}/{total}] {name} ===")

            candidates = collect_for_event(row, next_event_begin_by_name.get(name))

            out_row = row.to_dict()
            for i in range(1, MAX_URLS_PER_DISASTER + 1):
                if i <= len(candidates):
                    out_row[f"URL {i}"] = candidates[i - 1]["url"]
                    out_row[f"Title {i}"] = candidates[i - 1]["title"]
                    out_row[f"Snippet {i}"] = candidates[i - 1]["snippet"]
                else:
                    out_row[f"URL {i}"] = ""
                    out_row[f"Title {i}"] = ""
                    out_row[f"Snippet {i}"] = ""
            wide_rows.append(out_row)

            n = len(candidates)
            total_articles += n
            if n:
                events_with_articles += 1
                for c in candidates:
                    print(f"    + {c['displayLink']}: {c['title'][:60]}")
            print(f"  Saved {n} articles")

            if pos % 25 == 0 and not test_n:
                pd.DataFrame(wide_rows).to_csv(CHECKPOINT_FILE, index=False)
                print(f"  [Checkpoint] {pos} events written")

    except KeyboardInterrupt:
        print("\n[!] Interrupted — saving partial results…")
        interrupted = True
    except gs.DailyQuotaExhausted:
        print("\n[!] Daily API quota exhausted — re-run tomorrow to resume from the checkpoint.")
        interrupted = True
    finally:
        if test_n:
            # Smoke test: report only. Writing the checkpoint here would make
            # a later production run think these events were already done.
            print(f"\n=== SMOKE TEST SUMMARY ===")
            print(f"Events processed      : {len(wide_rows)}")
            print(f"Events with ≥1 article: {events_with_articles}")
            print(f"Total articles        : {total_articles}")
            if total_articles == 0:
                print("\n[!] Zero articles on the largest events in the dataset — that points")
                print("    at a config problem, not a coverage gap. Check that the PSE's domain")
                print("    list saved in Google's console and that 'Search the entire web' is OFF.")
            print("\nNo files written (smoke test). Re-run without --test for the full run.")
        elif wide_rows:
            wide_df = pd.DataFrame(wide_rows)
            wide_df.to_csv(CHECKPOINT_FILE, index=False)
            wide_df.to_csv(OUTPUT_WIDE, index=False)

            # Long format: one row per article, easier to append to Master.
            long_rows = []
            for r in wide_rows:
                for i in range(1, MAX_URLS_PER_DISASTER + 1):
                    url = r.get(f"URL {i}", "")
                    if not url:
                        continue
                    long_rows.append({
                        "Name": r.get("Name"),
                        "Disaster": r.get("Disaster"),
                        "Begin Date": r.get("Begin Date"),
                        "End Date": r.get("End Date"),
                        "CPI-Adjusted Cost": r.get("CPI-Adjusted Cost"),
                        "Unadjusted Cost": r.get("Unadjusted Cost"),
                        "Deaths": r.get("Deaths"),
                        "States": r.get("States"),
                        "Article #": i,
                        "URL": url,
                        "Title": r.get(f"Title {i}", ""),
                        "Snippet": r.get(f"Snippet {i}", ""),
                        "Source Engine": "RightLeaning",
                    })
            pd.DataFrame(long_rows).to_csv(OUTPUT_LONG, index=False)

            label = "INTERRUPTED — PARTIAL" if interrupted else "FINAL"
            print(f"\n=== {label} SUMMARY ===")
            print(f"Events processed      : {len(wide_rows)}")
            print(f"Events with ≥1 article: {events_with_articles}")
            print(f"Total articles        : {total_articles}")
            print(f"Wrote {OUTPUT_WIDE}")
            print(f"Wrote {OUTPUT_LONG} ({len(long_rows)} article rows)")
        else:
            print("\nNo data collected — no output written.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Collect candidate articles from right-leaning outlets only."
    )
    parser.add_argument(
        "--test", type=int, metavar="N", default=None,
        help="Smoke test: run N recent events only, write no files "
             "(e.g. --test 3, about 12 API calls).",
    )
    args = parser.parse_args()
    main(test_n=args.test)
