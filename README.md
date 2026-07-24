# Utah County Unfinished-Basement Lead Finder

Finds every residential building in Utah County built in the last 15 years with unfinished basement square footage. No map scraping — it queries the county's public ArcGIS REST API directly (the same backend as maps.utahcounty.gov), so a full county run is ~25 requests instead of thousands of page loads.

## Setup (once)

```
pip install requests openpyxl
```

## Run

```
python utah_basement_leads.py
```

First run pulls everything (~41,000 buildings currently) and looks up addresses/owners — expect a few minutes. Later runs only fetch addresses for *new* parcels, so they're fast.

Test with a small pull first: `python utah_basement_leads.py --limit 1` (one page of 2,000).

## Output

- **basement_leads.xlsx** — five sheets:
  - *All Active Leads* — parcel, address, city, zip, subdivision, owner + owner mailing address, year built, basement/finished/unfinished SF, % unfinished, above-grade SF, garage SF, baths, kitchens, ADU, style, quality, condition, exterior, acreage, land/improvement/total market value (appraisal), taxable value, school district, lat/lon + Google Maps pin link, first seen, link to county record
  - *New This Run* — parcels that appeared since the last run
  - *Dropped This Run* — parcels that stopped matching (usually means the basement got finished)
  - *New Construction Permits* — residential "NEW" permits from the county's recent-permits feed; these appear months before homes reach the assessment roll, so it's your early-warning list for future unfinished basements (builder-owned at permit time)
  - *Basement Finish Permits* — permits pulled to finish a basement; active leads with one of these also get a "Bsmt Finish Permit" date column on the lead sheets (that lead is being finished — act fast or drop it)
- **basement_leads.sqlite** — the persistent "proper table." Keeps `first_seen`, `last_seen`, `status`, `dropped_date` per building. The Excel file is regenerated from it each run, so first-seen dates survive across runs. You can query it directly with any SQLite tool.

## Tuning (top of script)

- `YEARS_BACK = 15` — rolling window, recomputed each run
- `RESIDENTIAL_ONLY = True` — single family / twin / town homes / plexes
- `MIN_UNFINISHED_SF = 1` — raise to e.g. 400 to skip tiny partials
- `INCLUDE_PERMITS = True` — set False to skip the permit sheets

## Map + literature drop tracking

Each run also writes **leads_map.html** — a filterable map of all active leads (city, year built, min unfinished SF, finish-permit status). Pins are colored by how recently that subdivision got literature: red = never, orange = >6 months, yellow = 3–6, green = <3. The sidebar lists subdivisions by lead count with last-drop date; click one to zoom to it.

**Logging drops:** add a row to **drops.csv** each time you distribute:

```
subdivision,date,who,notes
AUTUMN CREST PHASE 1,2026-07-23,Jared,round 1
```

Subdivision names must match the map/spreadsheet (case doesn't matter). Multiple rows per subdivision = multiple rounds; the map uses the most recent date. Use the "hide subdivisions dropped in last N months" filter to plan routes without overlap.

## Hosting the map (access from anywhere)

1. Create a **GitHub repo** (give it an unguessable name — the data is public record, but no need to advertise), enable **Settings → Pages → deploy from main branch**.
2. Put `leads_map.html` (rename to `index.html`) and `drops.csv` in it. After each script run, copy the fresh files over and push (or upload via github.com).
3. The map is now at `https://<user>.github.io/<repo>/` from any phone or computer.
4. Field logging: open `drops.csv` on github.com from your phone → pencil icon → add a row → commit. The hosted map re-reads drops.csv on every reload, so coverage colors update **without regenerating anything**.

## Run on a schedule (Windows)

Task Scheduler → Create Basic Task → weekly → Action "Start a program":
- Program: `python`
- Arguments: `C:\path\to\utah_basement_leads.py`
- Start in: `C:\path\to\` (folder containing the script — DB and Excel land there)

Weekly or monthly is plenty; the assessor data doesn't change daily, and it keeps load on the county server polite (the script also rate-limits itself).

## Notes

- `basement_leads_SAMPLE.xlsx` is a demo built from 6 real parcels (built 2024, >2400 SF unfinished) so you can see the format before a full run.
- The county's server intermittently returns a bogus "Layer not found" error (~half of requests from a flaky node). The script automatically retries, so full runs just take slightly longer.
- One home = one row: buildings with multiple assessment sections are deduplicated.
