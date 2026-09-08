# Poptavej.cz public tender monitor

Scrapes the following filtered search page on poptavej.cz every 3 days,
keeps only public tenders ("verejne zakazky") with an expected value above
10,000,000 CZK, and emails you an Excel file whenever new ones show up:

```
https://www.poptavej.cz/verejne-zakazky?filters%5Bkategorie%5D%5B0%5D=17&filters%5Bkategorie%5D%5B1%5D=1106&filters%5Bkategorie%5D%5B2%5D=10
```

Every run sends exactly one email:
- **New tenders found** -> an email with an `.xlsx` attachment (one row per new tender).
- **Nothing new** -> a short plain-text status email (confirms the job ran).
- **Scrape broke** -> a clearly-labeled failure alert email instead of the above.

## Unique ID used per tender

The listing page does not show "Cislo zakazky" directly, but each tender's
detail page does (confirmed by inspecting a live listing), and it matches
the ID embedded in that tender's own URL, e.g.:

```
/verejna-zakazka/VZ206438/zesileni-stresni-nosne-konstrukce
```

-> tender number `VZ206438`. The scraper extracts this straight from the
listing page's URL (no need to hit the detail page just to confirm it),
and uses it as the unique key for the seen-state file.

## Where the "summary" comes from

The listing page itself only shows name / value / category / region /
deadline - no description. The description ("Predmet plneni...") only
exists on each tender's own detail page. To keep the site footprint small,
the monitor only fetches a detail page for a tender that is **both** above
the 10M threshold **and** not already in the seen-state (i.e. genuinely new)
- so this extra request only happens for the small number of truly new,
qualifying tenders each run, not for every listing.

## Secrets to set (Settings -> Secrets and variables -> Actions)

| Secret | Value |
|---|---|
| `GMAIL_ADDRESS` | The Gmail address to send *from* (currently `maxze.zeiner@gmail.com`) |
| `GMAIL_APP_PASSWORD` | A Gmail **App Password** for that account (not your normal password) - create one at [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) (requires 2-Step Verification enabled) |
| `EMAIL_TO` | The recipient address (currently `maxze.zeiner@gmail.com`, same as sender for testing). Change this secret any time - no code changes needed. |

## Running it manually

Go to the repo's **Actions** tab -> **Poptavej.cz tender monitor** (left
sidebar) -> **Run workflow** button -> **Run workflow**. This uses the same
`workflow_dispatch` trigger as the schedule, so it exercises the full
scrape -> email -> commit flow on demand.

## Schedule

Runs automatically every 3 days via cron (see the comment in
`.github/workflows/poptavej-monitor.yml` for a note on cron's day-of-month
quirk and GitHub's "not exact-minute" scheduling behavior). You can also
just trigger it manually whenever you want, as above.

## Seen-state file

`poptavej-monitor/state/seen_tenders.json` - a JSON object mapping each
seen tender number to when it was first seen and its name. This file is
committed back to the repo by the workflow's last step after every run
(GitHub Actions runners don't persist a filesystem between runs, so this
file - not Actions cache - is the durable record). Do not edit it by hand
unless you want to make the monitor "forget" a tender (deleting an entry
will cause it to be re-reported as new on the next run).

**Heads up on the first run:** since this file starts empty, the very
first run will treat *every* currently-qualifying tender within the
lookback window (~6 days, see `LOOKBACK_DAYS` in `monitor.py`) as "new" -
so expect a bigger first batch than subsequent runs.

## Local files produced during a run

`poptavej-monitor/new_tenders.xlsx` is written locally as a temporary
attachment source when there's something to email; it isn't committed
(only the state JSON is committed by the workflow).

## Future: login step

Once you have poptavej.cz login credentials, they'll unlock more fields
(e.g. contact details on a tender's detail page). There's a clearly marked
`TODO` placeholder near the top of `monitor.py` for adding an authenticated
session - not implemented yet by design.

## A note on robots.txt

This path is disallowed in poptavej.cz's `robots.txt`. This was a deliberate,
informed decision given the very low request frequency (one run every 3
days, one sequential pass, small polite delays, realistic browser headers,
no aggressive retries). See the comments at the top of `monitor.py`.
