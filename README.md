# S&P 500 Dashboard

Daily S&P 500 performance tracker: a Python engine that pulls prices and generates an
interactive HTML dashboard, plus a standalone drag-and-drop version for loading exports
by hand.

## Files

- `sp500_update.py` — fetches S&P 500 constituents (Wikipedia) and prices (Yahoo Finance
  chart API), writes `SP500_Weekly_Performance.xlsx` (Daily / Weekly / By Sector tabs) and
  renders `docs/index.html` from `dashboard_template.html`.
- `dashboard_template.html` — template with `__DATA__` / `__META__` placeholders that
  `sp500_update.py` fills in. Not meant to be opened directly.
- `docs/SP500_DragDrop_Dashboard.html` — standalone dashboard that parses an `.xlsx` export
  entirely in the browser (via SheetJS). No pipeline dependency; drop any matching workbook
  onto it to view it.

## Running locally

```
pip install -r requirements.txt
python sp500_update.py
```

Writes output next to your Desktop by default. Override with env vars:

- `SP500_OUT` — path for the `.xlsx` output
- `SP500_DASH` — path for the generated dashboard HTML
- `SP500_NO_OPEN=1` — skip auto-opening the dashboard in a browser (used in CI)

## GitHub Actions

`.github/workflows/update.yml` runs the pipeline on a schedule (~7:55 PM Pacific daily,
see cron note below), then commits the refreshed `.xlsx` and `docs/index.html` back to the
repo. You can also trigger it manually from the **Actions** tab ("Run workflow") or with:

```
gh workflow run update.yml
```

**Cron/DST note:** GitHub Actions cron runs on UTC and doesn't shift for daylight saving.
The schedule is set for `02:55 UTC`, which lands at 7:55 PM Pacific during PDT (Mar–Nov)
and 6:55 PM Pacific during PST (Nov–Mar). Adjust the cron line in the workflow if you want
it pinned to a specific local time year-round. Also note: GitHub may pause scheduled
workflows on a repo with no other activity for 60+ days — a manual run or any push
re-enables them.

## GitHub Pages

`docs/` is published via GitHub Pages:

- `/` — the live auto-generated dashboard (populated after the first successful run)
- `/SP500_DragDrop_Dashboard.html` — the manual drag-and-drop tool (always available,
  independent of the pipeline)

**Note:** GitHub Pages sites are publicly viewable on the internet even when the
repository itself is private (personal-account private repos don't support a private
Pages audience — that requires GitHub Enterprise). The code/data history in this repo
stays private; the rendered dashboard page does not.

## Relationship to the local Windows Task Scheduler job

This repo runs independently of any local scheduled task on Logan's PC. Both can coexist;
retiring the local job (if redundant) is a separate, manual decision.
