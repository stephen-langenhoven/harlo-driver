# harlo-driver

Checks the billing-readiness of loads against the Harlo load-confirmation
service. Input is a "Ready to Confirm TC data" Excel export; the load number
is column `WBSHPGRP` (column T). For each unique load number the tool calls
the Harlo API and records the status. The status that matters is
**"Ready to Confirm"** (no exceptions, not halted) — it means all
prerequisites are met and the load can be billed.

## How it works

- No UI automation. Statuses come straight from the JSON API:
  `POST https://harlo.gambitco.io/api/load-confirmation/initiate` with body
  `{"loadNumber": "70127843"}`. The response carries `success`, `ready`,
  `result.halted`, `result.exceptions`, and a `result.reasoning` list of
  human-readable lines.
- Authentication follows the credential-caching approach from
  `altruos-test-client`: no passwords in the repo. `capture_auth.py` opens a
  visible browser, you log in manually once, and the session (cookies +
  localStorage) is saved to `.auth/harlo.json` (gitignored). `check_loads.py`
  reuses it until it expires, then asks you to re-run the capture.
- Input filenames vary by export date, so the workbook path is an argument.
  Results are written to a copy (`<input stem>_results.xlsx`) with a
  "Load Status" tab — never into the source file, which may be open in Excel.
- Every API response is appended to `<input stem>_results.cache.jsonl` as it
  arrives, so an interrupted run resumes instead of re-checking ~800 loads.

## Setup

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

## Usage

```bash
# One-time (repeat when the session expires): log in and cache the session
python capture_auth.py

# Check every unique load in an export
python check_loads.py Load_Confirmation_2026-07-20.xlsx
```

Useful flags for `check_loads.py`: `--limit N` (smoke-test on a few loads),
`--fresh` (ignore the cache and re-check everything), `--workers` (parallel
API calls, default 4), `--delay` (seconds each worker pauses between calls,
default 0.3), `--sheet` (if the data isn't on the first sheet).

## Output

`<input stem>_results.xlsx` — a copy of the input plus a "Load Status" tab:

| Load # | Status | First Exception | Checked At |
|---|---|---|---|

Statuses: `Ready to Confirm` (no exceptions and not halted — billable),
`Exception` (the load's first exception is in the First Exception column,
e.g. `FLAG: 50114252 BOL not complete (ERRORED)`), `Not Ready`, `ERROR`
(network/API failure). The `Exception` bucket is deliberately a single
category for now: after a full run, the script prints a tally of exception
patterns (with load/shipment numbers masked) — use that to define real
categories in `derive_status()` in `check_loads.py`.

## Notes

- `.auth/` holds live session tokens. It is gitignored; never commit it.
- Spreadsheets (`*.xlsx`, `~$*` lock files) and `*.cache.jsonl` are also
  gitignored — exports contain customer data and don't belong in git.
