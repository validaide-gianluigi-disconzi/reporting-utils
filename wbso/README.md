# WBSO hours — Java (no dependencies)

One self-contained Java file. No Maven, no libraries to download — it has its own
JSON parser and `.xlsx` writer built in, using only the JDK. You compile it once
and run it; it fetches the VALCM `wbso` issues + changelogs from Jira and writes a
finished workbook with the numbers already in it.

## Requirements

A JDK (version 11 or newer; tested on 21). Check with `javac -version`. On a Mac:
`brew install openjdk` if you don't have one.

## Build

```sh
javac WbsoHours.java
```

## Run

```sh
export JIRA_EMAIL="you@company.com"
export JIRA_API_TOKEN="your-token"        # id.atlassian.com/manage-profile/security/api-tokens
java WbsoHours --start 2026-01 --out wbso_hours.xlsx
```

Optional flags:

| Flag        | Default                         | Meaning                              |
|-------------|---------------------------------|--------------------------------------|
| `--start`   | `2026-01`                       | earliest month to count (YYYY-MM)    |
| `--out`     | `wbso_hours.xlsx`               | output filename                      |
| `--project` | `VALCM`                         | Jira project key                     |
| `--label`   | `wbso`                          | label to filter on                   |
| `--base`    | `https://skycell.atlassian.net` | Jira base URL                        |

To see it work without touching Jira (synthetic data):

```sh
java WbsoHours --selftest --out sample.xlsx
```

> **SECURITY:** the token is read from an environment variable and is never
> written to disk or into the workbook. You manage the token; nothing here stores
> it. To refresh next month, just run the command again.

## Output workbook

- **Summary** — person × month matrix of hours (the headline numbers).
- **Detail** — one row per (issue, person, month, phase): the full audit trail,
  including the hours and the time-held that produced them.
- **Issues** — per issue: story points, budgets, what was allocated, the method
  used, and any flag/note.
- **StatusAudit** — every distinct status encountered and how it was classified
  (DEV / QA / neither). **Check this first**: if a status that should count shows
  up as "no phase", add it (lowercase) to `DEV_STATUSES` / `QA_STATUSES` near the
  top of `WbsoHours.java` and recompile.
- **Notes** — the methodology, written out for an auditor.

## Methodology

- 1 story point = 4 hours.
- Base = Story Points DEV (`customfield_12867`) + QA (`customfield_12904`), each
  ×4. The total field (`customfield_10025`) is used only when both are empty.
- DEV hours go to whoever held the ticket during DEV-phase statuses, QA hours
  during QA-phase statuses, proportional to time held; months come from the actual
  status-transition timestamps, split at month boundaries, clipped to `--start`.
- This is a derived estimate (points × 4, allocated by status history), not a log
  of actual hours — a consistent, documented method, with every figure traceable.

Issues whose work-phase activity was entirely before `--start` correctly count as
zero (verified by the self-test). The Dutch-employee filter is **not** applied —
all assignees are included; filter downstream.

## What's verified vs. not

The JSON parsing, the time-weighted allocation (including month-boundary splits,
the total-SP fallback, and pre-window exclusion), and the `.xlsx` output were all
tested on synthetic data. The one thing that can't be tested without your
credentials is the live Jira fetch — if the first real run errors, the likely
causes are an auth issue, a story-point field ID that drifted, or a status name
that needs adding to the phase lists. Paste the error and I can pinpoint it.
