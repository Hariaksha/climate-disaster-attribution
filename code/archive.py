#!/usr/bin/env python3
"""
Phase 1 (v4 — Wayback Machine + stub detection)

Archives all Master sheet URLs via the Wayback Machine and extracts the full
article text directly, verifying it's a real article and not a paywall stub.

Replaces the archive.ph approach, which requires solving a CAPTCHA on every
single request and cannot be automated.

For each URL:
  1. Check the Wayback Machine availability API for an existing snapshot.
  2. If none exists, submit a Save Page Now request (free, no CAPTCHA).
  3. Fetch the snapshot HTML and extract the main article text (trafilatura).
  4. If extracted text has >= MIN_WORDS words, mark 'found' — ready for
     Phase 2 LLM coding, with the full text written directly to col 19.
  5. If too short (paywall stub) or no snapshot exists at all, mark
     'needs_factiva' — you'll need to pull these manually.

Results written to the Master sheet:
  Col 19: Full Article Text   — extracted text (only for 'found' rows)
  Col 20: Archive URL         — Wayback Machine snapshot URL
  Col 21: Archive Status      — 'found', 'needs_factiva', or 'failed'
  Col 22: Word Count          — word count of extracted text (for QA)

Fully resumable — rows already marked 'found' or 'needs_factiva' are skipped.
"""

import time
import re
import requests
import trafilatura
import openpyxl
from datetime import datetime

# ── Config ────────────────────────────────────────────────────────────────────
WORKBOOK_PATH = '/Users/hariaksha/Documents/GitHub/climate-disaster/attribution.xlsx'
SHEET_NAME    = 'Master'
# Column positions are resolved from the Master header row at import time,
# not hardcoded. They were hardcoded (19/20/21/22) until the LLM-coding
# columns were inserted, which pushed these five slots to the right and left
# the constants pointing at Contested / LLM Rationale / National Newspaper /
# Calibration Sample -- writing article text into those would have destroyed
# coded data on every row. Resolving by header name means a future column
# insertion cannot reintroduce that.
def _resolve_columns():
    wb = openpyxl.load_workbook(WORKBOOK_PATH, read_only=True)
    header = [c.value for c in next(wb[SHEET_NAME].iter_rows(max_row=1))]
    wb.close()
    wanted = {
        'URL_COL': 'URL',
        'RELEVANCE_COL': 'Relevance',
        'TEXT_COL': 'Full Article Text',
        'ARCHIVE_URL_COL': 'Archive URL',
        'ARCHIVE_STATUS_COL': 'Archive Status',
        'WORD_COUNT_COL': 'Word Count',
    }
    out = {}
    for const, name in wanted.items():
        if name not in header:
            raise SystemExit(f"Master sheet has no '{name}' column; aborting "
                             f"rather than writing to a guessed position.")
        out[const] = header.index(name) + 1
    return out

_COLS = _resolve_columns()
URL_COL            = _COLS['URL_COL']
RELEVANCE_COL      = _COLS['RELEVANCE_COL']
TEXT_COL           = _COLS['TEXT_COL']
ARCHIVE_URL_COL    = _COLS['ARCHIVE_URL_COL']
ARCHIVE_STATUS_COL = _COLS['ARCHIVE_STATUS_COL']
WORD_COUNT_COL     = _COLS['WORD_COUNT_COL']

MIN_WORDS = 150   # below this, treat as a paywall stub / teaser, not a real article
# NPR's /transcripts/{id} endpoint was given a much lower floor than MIN_WORDS
# on the reasoning that an official transcript is authoritative even when
# short. In practice the endpoint mostly serves a two-sentence *abstract* of an
# audio segment ("NPR's Noah Adams speaks with NPR's Adam Hochberg ... on the
# damage caused by Hurricane Ivan"), not a transcript. At 45-80 words that is a
# summary, not an article: it cannot express attribution framing, so it codes C
# by construction and inflates the no-causal-claim share.
#
# The old value also made the pipeline self-contradictory -- an 80-word Capital
# Gazette stub was rejected as paywalled while a 52-word NPR abstract was
# accepted as full text in the same run. Genuine transcripts run to thousands
# of words and clear the normal floor easily, so the special case now only
# buys a modest allowance rather than a different standard.
MIN_WORDS_NPR_TRANSCRIPT = 150
NPR_ID_PATTERN = re.compile(r'npr\.org/(?:[^/]+/){0,4}?(\d{5,})(?:/|$)')
CNN_FILEID_PATTERN = re.compile(r'transcripts\.cnn\.com/show/([a-z]+)\??.*start_fileid=([a-z]+)_(\d{4}-\d{2}-\d{2})_(\d+)')

DELAY_CHECK  = 1.5   # seconds between availability checks
DELAY_FETCH  = 1.0   # seconds between snapshot fetches
DELAY_SAVE_WAIT = 8  # seconds to wait after submitting Save Page Now
MAX_RETRIES  = 2
SAVE_EVERY   = 25    # rows between workbook saves (see the save block in main)

CONSEC_FAIL_THRESHOLD = 8
BACKOFF_SCHEDULE = [60, 180, 300]   # 1, 3, 5 minutes — WBM is much gentler than archive.ph
BACKOFF_COUNTDOWN_INTERVAL = 30

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    ),
}

WAYBACK_AVAIL_API = 'https://archive.org/wayback/available'
WAYBACK_SAVE_URL  = 'https://web.archive.org/save/'
# ─────────────────────────────────────────────────────────────────────────────

consecutive_failures = 0
backoff_level = 0


def countdown_pause(seconds: int, reason: str):
    print(f'\n⏸  {reason}')
    print(f'   Pausing for {seconds // 60}m {seconds % 60}s — script will resume automatically.')
    end_time = time.time() + seconds
    while True:
        remaining = int(end_time - time.time())
        if remaining <= 0:
            break
        mins, secs = divmod(remaining, 60)
        print(f'   Resuming in {mins}m {secs:02d}s ...', end='\r', flush=True)
        time.sleep(min(BACKOFF_COUNTDOWN_INTERVAL, remaining))
    print(f'\n▶  Resuming now.\n')


def handle_rate_limit():
    global backoff_level, consecutive_failures
    pause_secs = BACKOFF_SCHEDULE[min(backoff_level, len(BACKOFF_SCHEDULE) - 1)]
    backoff_level = min(backoff_level + 1, len(BACKOFF_SCHEDULE) - 1)
    consecutive_failures = 0
    countdown_pause(pause_secs, f'Rate limit detected — auto-pausing for {pause_secs // 60} minutes.')


def is_rate_limited(response) -> bool:
    if response is None:
        return False
    return response.status_code in (429, 503)


def check_existing(url: str, session: requests.Session, _retried=False):
    """Check Wayback Machine availability API. Returns (snapshot_url_or_None, rate_limited)."""
    try:
        r = session.get(WAYBACK_AVAIL_API, params={'url': url}, headers=HEADERS, timeout=20)
        if is_rate_limited(r):
            return None, True
        if r.status_code == 200:
            data = r.json()
            snap = data.get('archived_snapshots', {}).get('closest', {})
            if snap.get('available') and snap.get('url'):
                return snap['url'], False
    except Exception:
        if not _retried:
            time.sleep(3)
            return check_existing(url, session, _retried=True)
    return None, False


def submit_new(url: str, session: requests.Session):
    """Submit Save Page Now request. Returns (snapshot_url_or_None, rate_limited)."""
    try:
        r = session.get(f'{WAYBACK_SAVE_URL}{url}', headers=HEADERS, timeout=90, allow_redirects=True)
        if is_rate_limited(r):
            return None, True
        if 'web.archive.org/web/' in r.url:
            return r.url, False
    except requests.exceptions.Timeout:
        pass
    except Exception:
        pass

    # Save didn't redirect immediately (server error, still processing, etc.)
    # Wait and re-check via the availability API before giving up — the save
    # may have completed in the background, or an existing snapshot may have
    # been missed due to a transient error.
    time.sleep(12)
    return check_existing(url, session)


def fetch_and_extract(snapshot_url: str, session: requests.Session):
    """Fetch the snapshot HTML and extract main article text. Returns (text, word_count)."""
    try:
        r = session.get(snapshot_url, headers=HEADERS, timeout=30)
        if r.status_code != 200:
            return '', 0
        text = trafilatura.extract(
            r.text, url=snapshot_url, include_comments=False, include_tables=False
        )
        if not text:
            return '', 0
        return text, len(text.split())
    except Exception:
        return '', 0


CHROME_MAX_WORDS = 400   # above this an extraction is real content, never player chrome

PLAYER_CHROME_MARKERS = (
    'toggle more options',
    'iframe src=',
    'npr embedded audio player',
    'download embed',
)


def is_player_chrome(text: str) -> bool:
    """
    True when an extraction is audio-player furniture rather than a transcript.

    NPR serves /transcripts/{id} for audio-only segments that have no written
    transcript. trafilatura then returns the player controls -- "Listen Toggle
    more options Download Embed <iframe src=...> Facebook Flipboard Email",
    usually duplicated -- which can exceed MIN_WORDS_NPR_TRANSCRIPT and be
    accepted as an article. Length alone cannot separate that from a genuine
    short transcript, so match on the markup instead.
    """
    if not text:
        return True
    # Player chrome is inherently short -- it is the controls around a missing
    # transcript. Anything substantial is real content, so it is exempt: a
    # genuine transcript can legitimately contain the word "embed" or quote an
    # iframe, and lexical-diversity tests reject real transcripts outright
    # (speaker names and filler words repeat heavily across thousands of
    # words), so length is the reliable discriminator here.
    words = text.split()
    if len(words) > CHROME_MAX_WORDS:
        return False
    low = text.lower()
    return sum(marker in low for marker in PLAYER_CHROME_MARKERS) >= 2


def try_npr_transcript(original_url: str, session: requests.Session):
    """
    For npr.org URLs where the article page is a short teaser, NPR often has
    the full official transcript at npr.org/transcripts/{id}. Try that as a
    fallback. Returns (snapshot_url, text, word_count) or (None, '', 0).
    """
    if 'npr.org' not in original_url:
        return None, '', 0
    m = NPR_ID_PATTERN.search(original_url)
    if not m:
        return None, '', 0
    npr_id = m.group(1)
    transcript_url = f'https://www.npr.org/transcripts/{npr_id}'
    if transcript_url.rstrip('/') == original_url.rstrip('/'):
        return None, '', 0  # already tried this exact URL

    snap_url, rate_limited = check_existing(transcript_url, session)
    if rate_limited:
        handle_rate_limit()
        snap_url, rate_limited = check_existing(transcript_url, session)
    if not snap_url:
        snap_url, rate_limited = submit_new(transcript_url, session)
        if rate_limited:
            handle_rate_limit()
            snap_url, rate_limited = submit_new(transcript_url, session)

    if snap_url:
        text, wc = fetch_and_extract(snap_url, session)
        if wc >= MIN_WORDS_NPR_TRANSCRIPT and not is_player_chrome(text):
            return snap_url, text, wc
    return None, '', 0


def try_cnn_modern_url(original_url: str, session: requests.Session):
    """
    transcripts.cnn.com has two URL formats: a dead legacy query-string format
    (?start_fileid=show_date_NN) that falls back to a generic index page, and
    a working format (/show/{show}/date/{date}/segment/{NN}) with the same
    show/date/segment info. Reconstruct the working URL from the dead one.
    Returns (snapshot_url, text, word_count) or (None, '', 0).
    """
    m = CNN_FILEID_PATTERN.search(original_url)
    if not m:
        return None, '', 0
    show, _show2, date, seg = m.groups()
    new_url = f'https://transcripts.cnn.com/show/{show}/date/{date}/segment/{int(seg):02d}'

    snap_url, rate_limited = check_existing(new_url, session)
    if rate_limited:
        handle_rate_limit()
        snap_url, rate_limited = check_existing(new_url, session)
    if not snap_url:
        snap_url, rate_limited = submit_new(new_url, session)
        if rate_limited:
            handle_rate_limit()
            snap_url, rate_limited = submit_new(new_url, session)

    if snap_url:
        text, wc = fetch_and_extract(snap_url, session)
        if wc >= MIN_WORDS:
            return snap_url, text, wc
    return None, '', 0


def process_url(url: str, session: requests.Session) -> tuple:
    """
    Returns (snapshot_url, status, text, word_count).
    status is 'found', 'needs_factiva', or 'failed'.
    """
    global consecutive_failures, backoff_level
    rl_used = 0

    # Step 0: legacy transcripts.cnn.com URLs must be rewritten BEFORE anything
    # is fetched. The ?start_fileid= format is dead and serves a generic
    # chronological index of every episode of the show -- thousands of words of
    # unrelated segment titles. That sails past MIN_WORDS, so fetching the
    # legacy URL first returns 'found' on an index page and the modern-URL
    # fallback below is never reached. Converting up front means these rows
    # either get the real segment transcript or are correctly rejected; they
    # can no longer be silently filled with an index page.
    if CNN_FILEID_PATTERN.search(url):
        cnn_snap, cnn_text, cnn_wc = try_cnn_modern_url(url, session)
        if cnn_snap:
            consecutive_failures = 0
            return cnn_snap, 'found', cnn_text, cnn_wc
        # Conversion failed: do NOT fall through to fetching the legacy URL,
        # which would capture the index page. This is the 'wrong segment'
        # condition the corpus already records for these URLs.
        return None, 'trash_wrong_segment', '', 0

    # Step 1: check existing snapshot
    snap_url, rate_limited = check_existing(url, session)
    if rate_limited:
        rl_used += 1
        if rl_used <= 1:
            handle_rate_limit()
            snap_url, rate_limited = check_existing(url, session)

    if snap_url:
        text, wc = fetch_and_extract(snap_url, session)
        if wc >= MIN_WORDS:
            consecutive_failures = 0
            backoff_level = max(0, backoff_level - 1)
            return snap_url, 'found', text, wc
        else:
            # Found a snapshot but it's too short — likely a stub. Try saving a fresh copy.
            print(f'      → Snapshot found but only {wc} words (stub?) — trying fresh save...')

    # Step 2: submit Save Page Now (covers both "not found" and "stub" cases)
    print(f'      → Submitting Save Page Now...')
    new_snap_url, rate_limited = submit_new(url, session)
    if rate_limited:
        rl_used += 1
        if rl_used <= 1:
            handle_rate_limit()
            new_snap_url, rate_limited = submit_new(url, session)

    if new_snap_url:
        time.sleep(DELAY_SAVE_WAIT)
        text, wc = fetch_and_extract(new_snap_url, session)
        if wc >= MIN_WORDS:
            consecutive_failures = 0
            backoff_level = max(0, backoff_level - 1)
            return new_snap_url, 'found', text, wc
        else:
            # Still short — try source-specific fallbacks before giving up
            npr_snap, npr_text, npr_wc = try_npr_transcript(url, session)
            if npr_snap:
                consecutive_failures = 0
                return npr_snap, 'found', npr_text, npr_wc
            cnn_snap, cnn_text, cnn_wc = try_cnn_modern_url(url, session)
            if cnn_snap:
                consecutive_failures = 0
                return cnn_snap, 'found', cnn_text, cnn_wc
            return new_snap_url, 'needs_factiva', '', wc

    # Nothing usable found at all — last resort: source-specific fallbacks
    npr_snap, npr_text, npr_wc = try_npr_transcript(url, session)
    if npr_snap:
        consecutive_failures = 0
        return npr_snap, 'found', npr_text, npr_wc
    cnn_snap, cnn_text, cnn_wc = try_cnn_modern_url(url, session)
    if cnn_snap:
        consecutive_failures = 0
        return cnn_snap, 'found', cnn_text, cnn_wc

    # Last resort: fetch the live page directly. Everything above goes through
    # the Wayback Machine, so a freely-readable article whose URL simply has no
    # snapshot -- and which Save Page Now could not capture -- was being given
    # up on without ever being requested from its own server. Many outlets in
    # this corpus (deseret.com, carsonnow.org, newsmax.com, nationalreview.com)
    # serve full article text to an ordinary request, and metered paywalls
    # commonly serve it too. The MIN_WORDS floor still applies, so a hard
    # paywall's teaser is rejected here exactly as it is for a snapshot.
    direct_text, direct_wc = fetch_and_extract(url, session)
    if direct_wc >= MIN_WORDS and not is_player_chrome(direct_text):
        consecutive_failures = 0
        backoff_level = max(0, backoff_level - 1)
        # No snapshot URL to record: the text came from the live page, which is
        # worth distinguishing from an archived capture when tracing provenance.
        return None, 'found', direct_text, direct_wc

    if snap_url:
        # We had an old stub snapshot but couldn't save a better one
        return snap_url, 'needs_factiva', '', 0

    return None, 'failed', '', 0


def add_headers_if_missing(ws):
    headers = {
        TEXT_COL: 'Full Article Text',
        ARCHIVE_URL_COL: 'Archive URL',
        ARCHIVE_STATUS_COL: 'Archive Status',
        WORD_COUNT_COL: 'Word Count',
    }
    for col, name in headers.items():
        if ws.cell(row=1, column=col).value != name:
            ws.cell(row=1, column=col).value = name


def main():
    print(f'Loading workbook: {WORKBOOK_PATH}')
    wb = openpyxl.load_workbook(WORKBOOK_PATH)
    ws = wb[SHEET_NAME]
    add_headers_if_missing(ws)

    total_rows = ws.max_row - 1
    session = requests.Session()

    global consecutive_failures, backoff_level
    found = 0
    needs_factiva = 0
    failed = 0
    skipped = 0
    pending_since_save = 0

    start_time = datetime.now()
    print(f'Starting at {start_time.strftime("%H:%M:%S")} — {total_rows} articles to process\n')

    for row_idx in range(2, ws.max_row + 1):
        url = ws.cell(row=row_idx, column=URL_COL).value
        existing_status = ws.cell(row=row_idx, column=ARCHIVE_STATUS_COL).value

        article_num = row_idx - 1
        title = ws.cell(row=row_idx, column=13).value or ''

        if not url:
            skipped += 1
            continue

        # Work is driven by whether the row actually has article text, not by
        # its status label: a row whose text is empty still needs retrieving
        # whatever its status says, and a row that already has text must not
        # be re-fetched even if its status was never set.
        #
        # 'trash_gallery' is the one status that still short-circuits. Those
        # rows are photo/video gallery pages with no article text to extract
        # -- that is a deliberate classification, not a retrieval failure, and
        # re-running them would overwrite it with 'failed' and lose a
        # distinction the paper reports.
        existing_text = ws.cell(row=row_idx, column=TEXT_COL).value
        has_text = existing_text is not None and str(existing_text).strip() != ''

        # Rows already coded Trash are out of the corpus, so retrieving their
        # text is wasted network time -- llm_coder.py will never read it. This
        # covers the transcripts.cnn.com rows excluded as television, and any
        # other row the coder has already rejected.
        already_trash = str(ws.cell(row=row_idx, column=RELEVANCE_COL).value or '') == 'Trash'

        if has_text or already_trash or existing_status == 'trash_gallery':
            skipped += 1
            if article_num % 100 == 0:
                print(f'[{article_num:4d}/{total_rows}] ... skipping rows that already have text ...')
            continue

        print(f'[{article_num:4d}/{total_rows}] Processing | {title[:60]}')
        print(f'      URL: {str(url)[:80]}')

        snap_url, status, text, wc = None, 'failed', '', 0
        for attempt in range(1, MAX_RETRIES + 1):
            if attempt > 1:
                print(f'      Retry {attempt}/{MAX_RETRIES}...')
                time.sleep(5)
            snap_url, status, text, wc = process_url(str(url), session)
            if status in ('found', 'needs_factiva'):
                break

        ws.cell(row=row_idx, column=ARCHIVE_URL_COL).value = snap_url
        ws.cell(row=row_idx, column=ARCHIVE_STATUS_COL).value = status
        ws.cell(row=row_idx, column=WORD_COUNT_COL).value = wc
        if status == 'found':
            ws.cell(row=row_idx, column=TEXT_COL).value = text[:32000]  # Excel cell limit safety

        # Save periodically rather than after every row. Master is now ~2,400
        # rows carrying ~4,900 formulas, so a full workbook write costs more
        # than the network round trip it was protecting; at one save per row
        # the I/O dominates the run. A crash still loses at most SAVE_EVERY
        # rows, which the text-presence check re-queues on the next run.
        pending_since_save += 1
        if pending_since_save >= SAVE_EVERY:
            wb.save(WORKBOOK_PATH)
            pending_since_save = 0
            print(f'      [saved workbook]')

        if status == 'found':
            found += 1
            print(f'      ✓ Found full text ({wc} words): {snap_url}')
        elif status == 'needs_factiva':
            needs_factiva += 1
            print(f'      ⚠ Stub/paywalled ({wc} words) — needs Factiva: {snap_url}')
        else:
            failed += 1
            consecutive_failures += 1
            print(f'      ✗ No snapshot available — needs Factiva')
            if consecutive_failures >= CONSEC_FAIL_THRESHOLD:
                handle_rate_limit()

        print()
        time.sleep(DELAY_CHECK)

    # Flush whatever the last partial batch wrote.
    if pending_since_save:
        wb.save(WORKBOOK_PATH)
        print(f'[saved workbook — final flush of {pending_since_save} row(s)]\n')

    elapsed = datetime.now() - start_time
    print('=' * 60)
    print(f'DONE in {elapsed}')
    print(f'  Full text found:        {found:4d}')
    print(f'  Needs Factiva (stub):   {needs_factiva:4d}')
    print(f'  Needs Factiva (failed): {failed:4d}')
    print(f'  Skipped:                {skipped:4d}')
    total_needs_factiva = needs_factiva + failed
    print(f'\n  Total needing Factiva: {total_needs_factiva}/{total_rows - skipped}')
    print(f'\nWorkbook saved to: {WORKBOOK_PATH}')
    print('\nNext step: run phase2_llm_coding.py on rows marked "found".')
    print('For rows marked "needs_factiva", pull full text manually and paste into col 19.')


if __name__ == '__main__':
    main()
