# WBSO hours — Java (no dependencies)

One self-contained Java file. No Maven, no libraries to download — it has its own
JSON parser and `.xlsx` writer built in, using only the JDK. You compile it once
and run it per month; it fetches the VALCM `wbso` issues + changelogs from Jira and
writes a finished workbook with that month's numbers already in it: per person,
per RVO project, per working day, with the Jira evidence behind every hour.

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
java WbsoHours --month 2026-01 --eligible eligible.txt --out reports/wbso_2026-01.xlsx
```

This reports a **single calendar month**: all work performed *within* that month.
Run it once per month, and run it for earlier months to backfill the year — the
months are mutually consistent and add up to each issue's total (no double-count).

Optional flags:

| Flag            | Default                         | Meaning                                                        |
|-----------------|---------------------------------|----------------------------------------------------------------|
| `--month`       | current month                   | the month to report (YYYY-MM)                                  |
| `--out`         | `wbso_hours.xlsx`               | output filename                                                |
| `--eligible`    | none                            | file with one Jira display name per line; only they are counted |
| `--rvo-project` | `2601`                          | S&O project number from the S&O-verklaring (2026: 2601)        |
| `--project`     | `VALCM`                         | Jira project key                                               |
| `--label`       | `wbso`                          | label to filter on                                             |
| `--base`        | `https://skycell.atlassian.net` | Jira base URL                                                  |

A quick way to generate the whole year so far (Mac/Linux):

```sh
for m in 01 02 03 04 05 06 07 08 09; do
  java WbsoHours --month 2026-$m --eligible eligible.txt --out reports/wbso_2026-$m.xlsx
done
```

To see it work without touching Jira (synthetic data):

```sh
java WbsoHours --selftest --month 2026-01 --out sample.xlsx
```

> **SECURITY:** the token is read from an environment variable and is never
> written to disk or into the workbook. You manage the token; nothing here stores
> it. To refresh next month, just run the command again.

Transient Jira errors are retried automatically with backoff: rate limits (429),
server errors (5xx), network hiccups, and the occasional HTTP 400 whose body wraps
a backend 500 (Jira Cloud returns that for database faults). You see a
`transient Jira error ... retry n/8` line and the run continues. If it still
gives up after eight attempts, just run the same command again.

## Who counts: `eligible.txt`

WBSO is a reduction of the wage tax Validaide B.V. withholds, so an hour only
counts when it was worked by someone on the Validaide B.V. payroll (Dutch wage
tax withheld, BSN reported to RVO) and inside the EU. Remote colleagues on a
non-Dutch payroll, contractors and SkyCell AG staff do not count. Cross-border
cases: ask SubVice.

`eligible.txt` holds one Jira display name per line (`#` lines are comments).
Pass it with `--eligible`. The report then counts only those names; everyone
else, and unassigned time, is still shown under **NOT COUNTED** in the Summary
so the totals reconcile with Jira. Without the flag every named assignee is
counted and the Notes tab says so — fine for a quick look, not for filing.

## Which project: `--rvo-project`

RVO wants hours per S&O project. 2026 has one project, 2601, which is the
default. Next year pass the new number. If an issue carries a label
`wbso-<id>` (for example `wbso-2701`), that overrides the default for that
issue, so several projects can coexist in one Jira project.

## Output workbook

- **Summary** — person × month, in two blocks: COUNTED (on the eligibility list)
  and NOT COUNTED (not on the list, or unassigned), each with its own total,
  plus the grand total. The headline numbers.
- **Daily** — one row per (date, person, RVO project, issue, phase): the
  per-person, per-project, per-day record RVO's hours administration asks for.
  Working days only. A person with more than 8 h on one day is flagged.
- **DayTotals** — Daily pivoted: person × day of the month, the classic hours
  sheet.
- **Evidence** — every status and assignee change inside the month, with
  timestamp, author and a link to the issue: the progress trail behind the hours.
- **Detail** — one row per (issue, person, month, phase): hours and the
  working-day time held that produced them.
- **Issues** — per issue: summary, RVO project, story points, budgets, what was
  allocated, the method used, any flag/note, and a link.
- **StatusAudit** — every distinct status encountered and how it was classified
  (DEV / QA / neither). **Check this first**: if a status that should count shows
  up as "no phase", add it (lowercase) to `DEV_STATUSES` / `QA_STATUSES` near the
  top of `WbsoHours.java` and recompile.
- **Notes** — the methodology, written out for an auditor, including whether an
  eligibility list was applied.

## Methodology

- 1 story point = 4 hours.
- Base = Story Points DEV (`customfield_12867`) + QA (`customfield_12904`), each
  ×4. The total field (`customfield_10025`) is used only when both are empty.
- DEV hours go to whoever held the ticket during DEV-phase statuses, QA hours
  during QA-phase statuses, proportional to time held. Time held is counted on
  **working days only** (Monday to Friday), in numerator and denominator alike,
  so no hours land on a weekend.
- The month reported gets the share of the budget equal to **phase-time inside
  that month ÷ phase-time over the issue's whole life**. Running every month
  gives consistent, additive results that sum to each issue's total.
- Within the month, each person's share is spread over the working days the
  ticket was held, proportional to time held per day. Daily rows sum to the
  Detail rows.
- This is a derived estimate (points × 4, allocated by status history), not a
  log of actual hours — a consistent, documented method, with every figure
  traceable to an issue, a person, a project and a day.

Work done outside the reported month simply scores zero for that month (it shows
up when you run that other month). Because an issue's whole-life phase time keeps
growing while it is open, earlier months shift slightly when re-run later:
**re-run the whole year before the realisation report goes to SubVice.**

## What's verified vs. not

The JSON parsing, the within-month time-weighted allocation (including the
boundary case where a ticket's work straddles two months, the total-SP fallback,
tickets whose work falls outside the reported month, the per-day split and the
eligibility blocks), and the `.xlsx` output were tested on synthetic data across
multiple months — and confirmed to sum to the issue totals. The one thing that
can't be tested without your credentials is the live Jira fetch — if the first
real run errors, the likely causes are an auth issue, a story-point field ID that
drifted, or a status name that needs adding to the phase lists.
