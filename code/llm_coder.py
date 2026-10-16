"""
Phase 2 -- LLM coding of the Master sheet.

Codes every Master row that has retrieved article text (Archive Status
"found"), is not one of the 50 rows hand-coded in the Calibration sheet, and
has no Relevance code yet. That last condition makes the script resume-safe:
re-running it only picks up rows still blank.

Each article gets the two-step procedure from paper/attribution_codebook.md --
relevance screening (Relevant / Borderline / Trash), then attribution coding
for Relevant rows (category A-E, plus the D-subtype and Contested tags).

Writes Relevance, Attribution Score, Attribution Category, D-Subtype,
Contested and LLM Rationale back to Master. Attribution Score is derived from
the category in the same write step and stored as a literal rather than a
formula, so openpyxl can read it without Excel having recalculated first.

Spend is capped at MAX_BUDGET_USD and the script stops rather than exceed it.

Usage:
    pip install anthropic openpyxl
    export ANTHROPIC_API_KEY=sk-ant-...
    python code/llm_coder.py        # TEST_MODE first; set False for full run
"""

import os
import sys
import time
import json
import random

import openpyxl

try:
    import anthropic
except ImportError:
    sys.exit(
        "Missing dependency. Run: pip install anthropic openpyxl"
    )

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

XLSX_PATH = "attribution.xlsx"
MODEL = "claude-sonnet-4-6"

# Pricing as of this writing (verify at https://platform.claude.com/docs/en/about-claude/pricing
# before a large run -- prices can change).
PRICE_INPUT_PER_MTOK = 3.00
PRICE_OUTPUT_PER_MTOK = 15.00

MAX_BUDGET_USD = 30.00       # hard stop
CHECKPOINT_EVERY = 25        # save workbook every N coded rows
TEST_MODE = False            # True -> only code TEST_MODE_ROWS rows, then stop
TEST_MODE_ROWS = 5

MAX_RETRIES = 5
BASE_BACKOFF_SECONDS = 2.0

# Master sheet columns, looked up by header text at startup.
#
# These used to be hardcoded indices. The sibling retrieval script
# (archive.py) did the same and its constants went stale when the attribution
# columns were inserted, shifting everything after them five places right --
# it would have written article text into Contested and status into National
# Newspaper on every row it touched. Nothing detects that: the writes succeed
# and land on coded data. Reading the header row costs one pass and removes
# the failure mode.
#
# Each entry is (accepted header names, fallback index). The fallback is used
# only when no listed name is in the header row, which is how the four
# attribution columns get created on a workbook that does not have them yet
# (see NEW_COLUMN_HEADERS below).
COLUMN_SPEC = {
    "COL_DISASTER_ROW": (("Disaster Row #",), 1),
    "COL_ARTICLE_NUM": (("Article #",), 2),
    "COL_NAME": (("Name",), 3),
    "COL_DISASTER_TYPE": (("Disaster Type",), 4),
    "COL_BEGIN_DATE": (("Begin Date",), 5),
    "COL_END_DATE": (("End Date",), 6),
    "COL_STATES": (("States",), 10),
    "COL_URL": (("URL",), 12),
    "COL_TITLE": (("Title",), 13),
    "COL_RELEVANCE": (("Relevance",), 15),
    # "Attribution Index" is the pre-rename header for this column; main()
    # relabels it to "Attribution Score" on the way through.
    "COL_ATTRIBUTION_SCORE": (("Attribution Score", "Attribution Index"), 16),
    "COL_ATTRIBUTION_CATEGORY": (("Attribution Category",), 17),
    "COL_D_SUBTYPE": (("D-Subtype",), 18),
    "COL_CONTESTED": (("Contested",), 19),
    "COL_LLM_RATIONALE": (("LLM Rationale",), 20),
    "COL_NATIONAL_NEWSPAPER": (("National Newspaper",), 21),
    "COL_CALIBRATION_SAMPLE": (("Calibration Sample",), 22),
    "COL_FULL_TEXT": (("Full Article Text",), 23),
    "COL_ARCHIVE_URL": (("Archive URL",), 24),
    "COL_ARCHIVE_STATUS": (("Archive Status",), 25),
    "COL_WORD_COUNT": (("Word Count",), 26),
}

CALIBRATION_URL_HEADERS = ("URL",)
CALIBRATION_URL_FALLBACK = 8


def resolve_columns(header_row, spec, sheet_name):
    """Map each column name to its 1-based index in this header row."""
    lookup, seen_twice = {}, set()
    for idx, value in enumerate(header_row, 1):
        if value is None:
            continue
        name = str(value).strip()
        if name in lookup:
            seen_twice.add(name)
        else:
            lookup[name] = idx

    # A repeated header is not fatal -- the first occurrence wins, which is
    # deterministic -- but it is worth saying out loud, since the duplicate is
    # usually a stray copy and the wrong one could be the first.
    wanted = {n for names, _ in spec.values() for n in names}
    for name in sorted(seen_twice & wanted):
        print(f"{sheet_name}: header {name!r} appears more than once; "
              f"using column {lookup[name]}.")

    resolved, from_fallback = {}, []
    for const, (names, fallback) in spec.items():
        for name in names:
            if name in lookup:
                resolved[const] = lookup[name]
                break
        else:
            resolved[const] = fallback
            from_fallback.append(f"{const}={fallback}")

    clashes = [c for c, i in resolved.items()
               if list(resolved.values()).count(i) > 1]
    if clashes:
        raise SystemExit(
            f"{sheet_name}: two columns resolved to the same index "
            f"({', '.join(sorted(clashes))}). Check the header row for "
            f"duplicate or missing names rather than letting the run "
            f"overwrite the wrong column."
        )

    if from_fallback:
        print(f"{sheet_name}: no header found for "
              f"{', '.join(from_fallback)} -- using default position(s).")
    return resolved

# Columns this script is responsible for creating if the workbook predates
# them. Keyed by the constant rather than an index, since the index is only
# known once the header row has been read.
NEW_COLUMN_HEADERS = {
    "COL_ATTRIBUTION_CATEGORY": "Attribution Category",
    "COL_D_SUBTYPE": "D-Subtype",
    "COL_CONTESTED": "Contested",
    "COL_LLM_RATIONALE": "LLM Rationale",
}

CATEGORY_TO_SCORE = {"A": 1.0, "B": 0.5, "C": 0.0, "D": -1.0, "E": 0.0}

# ---------------------------------------------------------------------------
# Prompt, built from paper/attribution_codebook.md
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """You are coding newspaper articles for the research project \
"Reading the Storm: Media Attribution of U.S. Climate Disasters, 2000-2024." \
You will be given metadata about a specific NOAA-documented billion-dollar \
disaster event, and the full text of one article retrieved as candidate \
coverage of that event. Apply the two-step procedure below exactly.

STEP 1 - RELEVANCE SCREENING
Decide whether the article is actually about THIS SPECIFIC disaster event \
(correct disaster type, correct geography/state, and a date consistent with \
the event window, including reasonable lead-up or retrospective coverage).

- Relevant: the article substantively discusses this specific event -- its \
impacts, recovery, costs, forecasts, or causal framing.
- Borderline: plausibly about this event but ambiguous -- e.g. a wire brief \
too thin to confirm the event match, a "wildfire season" or "hurricane \
season" roundup that may or may not cover this specific event, or a \
retrospective piece (1+ years later) that mentions the event only in passing.
- Trash: off-topic, wrong event/year/disaster type, a paywall/login/cookie \
page, an aggregator or archive listing page with no real article content, a \
dead link, or a photo gallery with only captions (captions cannot sustain \
causal framing).

STEP 2 - ATTRIBUTION CODING (Relevant articles only)
If, and only if, Step 1 = Relevant, assign exactly one attribution category:

- A (score +1.0) Explicit anthropogenic attribution: the article directly \
and affirmatively links the event's occurrence, frequency, or intensity to \
human-caused climate change / global warming. Often cites climate \
scientists, formal attribution studies (e.g. World Weather Attribution), or \
uses confident causal language. Example: "Climate scientists say the \
unprecedented rainfall was supercharged by a warmer atmosphere, which holds \
more moisture."
- B (score +0.5) Hedged / contextual climate mention: climate change is \
raised as a possible or partial factor, with hedged language ("may," \
"could," "some scientists believe"), or mentioned briefly/contextually \
without being central to the causal framing. Example: "The fire comes amid \
a years-long drought that some link to a changing climate."
- C (score 0.0) No causal attribution: pure event reporting -- impacts, \
damages, casualties, response, recovery, forecasts -- with no discussion of \
broader cause (climate or otherwise). Example: "The storm knocked out power \
to 500,000 homes across the Southeast."
- D (score -1.0) Explicit non-climate attribution: the article attributes \
the event to natural variability/cyclical weather patterns (El Nino/La \
Nina, "typical for this time of year," historical cycles) OR to non-climate \
human factors (floodplain development, forest management, infrastructure \
failures) -- and in doing so explicitly downplays or rejects a climate-\
change link. Example: "Meteorologists say this is a fairly typical La Nina \
winter pattern, not evidence of a longer-term trend." For category D only, \
also set d_subtype to "natural_variability" or "other_human_causes" \
depending on which of the two it is.
- E (score 0.0, contested) Mixed / contested: the article presents both \
climate-linked and non-climate framings with roughly comparable weight \
(e.g. a climate scientist and a skeptical official both quoted). Example: \
"While some researchers point to climate change as a contributing factor, \
local officials say the flooding was mainly due to outdated drainage \
infrastructure." Set contested = true only for category E.

If Step 1 is Borderline or Trash, attribution_category, d_subtype, and \
contested must all be null/false -- do not guess a category for an article \
you did not mark Relevant.

Always call the `code_article` tool with your answer. Keep rationale to 1-2 \
sentences -- it is for audit/spot-checking, not for the dataset itself."""

CODE_ARTICLE_TOOL = {
    "name": "code_article",
    "description": "Record the relevance and attribution coding for one article.",
    "input_schema": {
        "type": "object",
        "properties": {
            "relevance": {
                "type": "string",
                "enum": ["Relevant", "Borderline", "Trash"],
            },
            "attribution_category": {
                "type": ["string", "null"],
                "enum": ["A", "B", "C", "D", "E", None],
                "description": "Required if relevance == Relevant, else null.",
            },
            "d_subtype": {
                "type": ["string", "null"],
                "enum": ["natural_variability", "other_human_causes", None],
                "description": "Set only if attribution_category == D, else null.",
            },
            "contested": {
                "type": "boolean",
                "description": "True only if attribution_category == E.",
            },
            "rationale": {
                "type": "string",
                "description": "1-2 sentence justification for audit purposes.",
            },
        },
        "required": ["relevance", "contested", "rationale"],
    },
}


def build_user_message(ws, row):
    name = ws.cell(row=row, column=COL_NAME).value
    disaster_type = ws.cell(row=row, column=COL_DISASTER_TYPE).value
    begin_date = ws.cell(row=row, column=COL_BEGIN_DATE).value
    end_date = ws.cell(row=row, column=COL_END_DATE).value
    states = ws.cell(row=row, column=COL_STATES).value
    title = ws.cell(row=row, column=COL_TITLE).value
    url = ws.cell(row=row, column=COL_URL).value
    full_text = ws.cell(row=row, column=COL_FULL_TEXT).value or ""

    return f"""DISASTER EVENT METADATA
Name: {name}
Disaster type: {disaster_type}
Event window: {begin_date} to {end_date}
States affected: {states}

ARTICLE
URL: {url}
Title: {title}

Full article text:
\"\"\"
{full_text}
\"\"\"

Code this article per the procedure in your instructions."""


# ---------------------------------------------------------------------------
# Cost tracking
# ---------------------------------------------------------------------------

class Budget:
    def __init__(self, max_usd):
        self.max_usd = max_usd
        self.spent_usd = 0.0
        self.calls = 0

    def add(self, input_tokens, output_tokens):
        cost = (input_tokens / 1e6) * PRICE_INPUT_PER_MTOK + (
            output_tokens / 1e6
        ) * PRICE_OUTPUT_PER_MTOK
        self.spent_usd += cost
        self.calls += 1
        return cost

    def exceeded(self):
        return self.spent_usd >= self.max_usd


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def call_with_retry(client, **kwargs):
    for attempt in range(MAX_RETRIES):
        try:
            return client.messages.create(**kwargs)
        except anthropic.RateLimitError:
            wait = BASE_BACKOFF_SECONDS * (2 ** attempt) + random.uniform(0, 1)
            print(f"  rate limited, retrying in {wait:.1f}s...")
            time.sleep(wait)
        except anthropic.APIStatusError as e:
            if e.status_code >= 500 and attempt < MAX_RETRIES - 1:
                wait = BASE_BACKOFF_SECONDS * (2 ** attempt) + random.uniform(0, 1)
                print(f"  server error {e.status_code}, retrying in {wait:.1f}s...")
                time.sleep(wait)
            else:
                raise
    raise RuntimeError("Exceeded max retries")


def main():
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        sys.exit(
            "ANTHROPIC_API_KEY is not set. Set it in your shell "
            "(export ANTHROPIC_API_KEY=sk-ant-...) -- do not hardcode it "
            "here or put it in a file that gets committed to git."
        )

    client = anthropic.Anthropic(api_key=api_key)

    print(f"Loading {XLSX_PATH} ...")
    wb = openpyxl.load_workbook(XLSX_PATH)
    ws = wb["Master"]
    cal = wb["Calibration"]

    # Locate every column by its header before touching anything.
    cols = resolve_columns(
        [c.value for c in next(ws.iter_rows(max_row=1))], COLUMN_SPEC, "Master"
    )
    globals().update(cols)

    cal_url_col = resolve_columns(
        [c.value for c in next(cal.iter_rows(max_row=1))],
        {"COL_CAL_URL": (CALIBRATION_URL_HEADERS, CALIBRATION_URL_FALLBACK)},
        "Calibration",
    )["COL_CAL_URL"]

    # Ensure new column headers exist
    for const, header in NEW_COLUMN_HEADERS.items():
        col = cols[const]
        if ws.cell(row=1, column=col).value != header:
            ws.cell(row=1, column=col).value = header

    # "Attribution Index" is the disaster-level aggregate; this column is the
    # article-level score. Safe to rename -- All Disasters refers to it by
    # letter (Master!P), not by header text.
    ws.cell(row=1, column=COL_ATTRIBUTION_SCORE).value = "Attribution Score"

    # Rows already hand-coded in Calibration, matched by URL -- skip these,
    # the calibration gate already covers them.
    calibration_urls = set()
    for r in range(2, cal.max_row + 1):
        u = cal.cell(row=r, column=cal_url_col).value
        if u:
            calibration_urls.add(u.strip())

    # Select candidate rows
    candidates = []
    for r in range(2, ws.max_row + 1):
        url = ws.cell(row=r, column=COL_URL).value
        status = ws.cell(row=r, column=COL_ARCHIVE_STATUS).value
        full_text = ws.cell(row=r, column=COL_FULL_TEXT).value
        already_coded = ws.cell(row=r, column=COL_RELEVANCE).value
        if not url or status != "found" or not full_text or already_coded:
            continue
        if url.strip() in calibration_urls:
            continue
        candidates.append(r)

    print(f"{len(candidates)} rows queued for coding "
          f"(found, not in Calibration sample, not yet coded).")

    if TEST_MODE:
        candidates = candidates[:TEST_MODE_ROWS]
        print(f"TEST_MODE on -- limiting to first {len(candidates)} rows. "
              f"Set TEST_MODE = False for the full run.")

    budget = Budget(MAX_BUDGET_USD)
    coded_since_save = 0

    for i, row in enumerate(candidates, 1):
        if budget.exceeded():
            print(f"Budget cap (${MAX_BUDGET_USD:.2f}) reached after "
                  f"{budget.calls} calls / ${budget.spent_usd:.2f} spent. Stopping.")
            break

        user_msg = build_user_message(ws, row)

        response = call_with_retry(
            client,
            model=MODEL,
            max_tokens=400,
            system=SYSTEM_PROMPT,
            tools=[CODE_ARTICLE_TOOL],
            tool_choice={"type": "tool", "name": "code_article"},
            messages=[{"role": "user", "content": user_msg}],
        )

        cost = budget.add(response.usage.input_tokens, response.usage.output_tokens)

        tool_use_block = next(
            (b for b in response.content if b.type == "tool_use"), None
        )
        if tool_use_block is None:
            print(f"  row {row}: no tool_use in response, skipping")
            continue

        result = tool_use_block.input
        relevance = result.get("relevance")
        category = result.get("attribution_category")
        d_subtype = result.get("d_subtype")
        contested = result.get("contested", False)
        rationale = result.get("rationale", "")

        # Defensive normalization: enforce the codebook's constraints
        # ourselves rather than trusting the model never deviates from the
        # system prompt's instructions (it's enum-constrained, but cheap
        # insurance against e.g. a Borderline row coming back with a
        # leftover d_subtype/contested value).
        if relevance != "Relevant":
            category = None
            d_subtype = None
            contested = False
        if category != "D":
            d_subtype = None
        if category != "E":
            contested = False

        ws.cell(row=row, column=COL_RELEVANCE).value = relevance
        if relevance == "Relevant" and category in CATEGORY_TO_SCORE:
            ws.cell(row=row, column=COL_ATTRIBUTION_SCORE).value = CATEGORY_TO_SCORE[category]
            ws.cell(row=row, column=COL_ATTRIBUTION_CATEGORY).value = category
        else:
            ws.cell(row=row, column=COL_ATTRIBUTION_SCORE).value = None
            ws.cell(row=row, column=COL_ATTRIBUTION_CATEGORY).value = None
        ws.cell(row=row, column=COL_D_SUBTYPE).value = d_subtype
        ws.cell(row=row, column=COL_CONTESTED).value = bool(contested)
        ws.cell(row=row, column=COL_LLM_RATIONALE).value = rationale

        coded_since_save += 1
        print(f"[{i}/{len(candidates)}] row {row}: {relevance}"
              f"{'/' + category if category else ''}"
              f"  (+${cost:.4f}, running total ${budget.spent_usd:.2f})")

        if coded_since_save >= CHECKPOINT_EVERY:
            wb.save(XLSX_PATH)
            print(f"  checkpoint saved ({budget.calls} calls so far).")
            coded_since_save = 0

    wb.save(XLSX_PATH)
    print(f"Done. {budget.calls} calls made, ${budget.spent_usd:.2f} spent.")


if __name__ == "__main__":
    main()
