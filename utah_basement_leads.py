#!/usr/bin/env python3
"""
Utah County unfinished-basement lead finder.

Pulls every residential building in Utah County built in the last N years
that has unfinished basement square footage, straight from the county's
public ArcGIS REST API (the same backend that powers maps.utahcounty.gov).
No HTML scraping, no map-clicking - two bulk endpoints:

  1. Assessor/Sketches_BuildingSummary - year built, basement SF,
     finished SF, unfinished SF, per building.
  2. Parcels/Parcel_Name_Address - site address, city, zip, owner,
     market value, per parcel (joined on PARCELID).

State is kept in a local SQLite database so every parcel gets a
"first seen" date, and parcels that stop matching (e.g. the basement
got finished) are flagged instead of silently disappearing.

Each run writes an Excel workbook with three sheets:
  All Active Leads / New This Run / Dropped This Run

Usage:
    python utah_basement_leads.py            # normal run
    python utah_basement_leads.py --limit 3  # test run, 3 pages only

Requires:  pip install requests openpyxl
"""

import argparse
import csv
import datetime as dt
import json
import sqlite3
import sys
import time
from pathlib import Path

import requests

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------
YEARS_BACK = 15                 # rolling window: homes built within N years
RESIDENTIAL_ONLY = True         # TXACCTTYPE = 'RESIDENTIAL' (single family,
                                # twin/town homes, plexes). False = everything.
MIN_UNFINISHED_SF = 1           # match if unfinished basement SF >= this
INCLUDE_PERMITS = True          # also pull building permits:
                                #  - NEW residential = future basement leads
                                #    (shows up months before the assessor roll)
                                #  - BASEMENT FINISH = leads being finished
PERMIT_LOOKBACK_YEARS = 3       # how far back to pull permit history
                                # (full county history goes back decades)

SCRIPT_DIR = Path(__file__).resolve().parent
DB_FILE = SCRIPT_DIR / "basement_leads.sqlite"
XLSX_FILE = SCRIPT_DIR / "basement_leads.xlsx"
MAP_FILE = SCRIPT_DIR / "leads_map.html"
DROPS_FILE = SCRIPT_DIR / "drops.csv"   # literature distribution log

BASE = "https://maps.utahcounty.gov/arcgis/rest/services"
BUILDINGS_URL = f"{BASE}/Assessor/Sketches_BuildingSummary/MapServer/0/query"
PARCELS_URL = f"{BASE}/Parcels/Parcel_Name_Address/MapServer/2/query"
# Building_Permits_Status = FULL permit history (~100k records) incl.
# certificate-of-occupancy dates; Building_Permits_Recent only holds the
# assessor's in-progress window (~10k) and drops permits once processed.
PERMITS_URL = f"{BASE}/Assessor/Building_Permits_Status/FeatureServer/0/query"
APPRAISAL_URL = "https://www.utahcounty.gov/LandRecords/AppraisalInfo.asp?avParcelId={pid}"

PAGE_SIZE = 2000                # server maxRecordCount
ADDR_CHUNK = 100                # parcels per address-lookup request
REQUEST_DELAY = 0.3             # seconds between requests (be polite)
TIMEOUT = 60
RETRIES = 10                    # county server has a flaky node that randomly
                                # answers 404 "Layer not found"; retries are cheap

BUILDING_FIELDS = [
    "PARCELID", "BUILDINGID", "IMP_NO", "YEAR_BUILT", "BASEMENT",
    "BSMT_FINISH", "BSMT_UNFINISH", "ABOVE_GRADE_AREA", "ATTACHED_GARAGE",
    "BATHS", "KITCHENS", "ADU", "COST_GROUP_DESCR", "STYLE_DESCR",
    "QUALITY_DESCR", "CONDITION_DESCR", "EXTERIOR_WALL", "PCT_COMPLETE",
    "RCNLD", "TAX_YEAR", "TXACCTTYPE", "OVERALL_RN",
]
PARCEL_FIELDS = [
    "PARCELID", "SITE_FULLADDRESS", "SITE_CITY", "SITE_ZIP5",
    "OWNERNAME", "OWN_FULLADDRESS", "USEDSCRP", "CNVYNAME", "ACREAGE",
    "SCHLDSCRP", "MKT_LANDVALUE", "MKT_IMPVALUE", "MKT_CNTVAL",
    "TXBL_CNTVAL",
]
PERMIT_FIELDS = [
    "PERMITNO", "LOCALPERMITNO", "ACCOUNTNO", "PERMITTYPE",
    "PERMITCLASSIFICATION", "PERMITDATE", "PERMITWORKDATE",
    "PERMITSTATUS", "CERTIFICATEOFOCCUPANCYDATE",
]

session = requests.Session()
session.headers["User-Agent"] = "basement-lead-finder/1.0 (contact: jared@cascade.build)"


# --------------------------------------------------------------------------
# HTTP
# --------------------------------------------------------------------------
def fetch_json(url: str, params: dict) -> dict:
    """POST an ArcGIS query (POST avoids URL-length limits) with retries.

    The county's load balancer has a bad node that instantly returns
    {"error": {"code": 404, "message": "Layer not found"}} for valid
    queries. Keep-alive can pin us to that node, so on any error we drop
    the connection pool to get a fresh backend, then retry quickly.
    """
    last_err = None
    for attempt in range(1, RETRIES + 1):
        try:
            r = session.post(url, data=params, timeout=TIMEOUT)
            r.raise_for_status()
            data = r.json()
            if "error" in data:
                raise RuntimeError(f"ArcGIS error: {data['error']}")
            return data
        except Exception as e:  # noqa: BLE001
            last_err = e
            session.close()  # new connection -> likely a different node
            if attempt < RETRIES:
                # instant bogus 404s need only a short pause; back off a
                # little more each time in case it's real server distress
                time.sleep(min(0.5 * attempt, 5))
    raise RuntimeError(f"Request failed after {RETRIES} tries: {last_err}")


def query_buildings(where: str, max_pages: int | None = None):
    """Yield building attribute dicts, paging through the whole result set."""
    offset, page = 0, 0
    while True:
        params = {
            "where": where,
            "outFields": ",".join(BUILDING_FIELDS),
            "returnGeometry": "false",
            "orderByFields": "OBJECTID",
            "resultOffset": offset,
            "resultRecordCount": PAGE_SIZE,
            "f": "json",
        }
        data = fetch_json(BUILDINGS_URL, params)
        feats = data.get("features", [])
        for f in feats:
            yield f["attributes"]
        page += 1
        print(f"  page {page}: {len(feats)} buildings (total so far ~{offset + len(feats)})")
        if not data.get("exceededTransferLimit") or not feats:
            break
        if max_pages and page >= max_pages:
            print("  --limit reached, stopping early (test mode)")
            break
        offset += PAGE_SIZE
        time.sleep(REQUEST_DELAY)


def polygon_centroid(rings) -> tuple[float | None, float | None]:
    """Centroid (lon, lat) of the largest ring of an esri polygon."""
    best, best_area = None, -1.0
    for ring in rings or []:
        a = 0.0
        for (x1, y1), (x2, y2) in zip(ring, ring[1:]):
            a += x1 * y2 - x2 * y1
        if abs(a) > best_area:
            best_area, best = abs(a), ring
    if not best:
        return None, None
    a = cx = cy = 0.0
    for (x1, y1), (x2, y2) in zip(best, best[1:]):
        cross = x1 * y2 - x2 * y1
        a += cross
        cx += (x1 + x2) * cross
        cy += (y1 + y2) * cross
    if abs(a) < 1e-14:  # degenerate: fall back to vertex average
        xs, ys = zip(*best)
        return sum(xs) / len(xs), sum(ys) / len(ys)
    a *= 0.5
    return cx / (6 * a), cy / (6 * a)


def query_addresses(parcel_ids: list[str]):
    """Yield {parcelid: attrs} per chunk so progress can be saved as we go.

    A chunk that still fails after all retries is skipped with a warning
    instead of aborting the run; those parcels are retried automatically
    on the next run (only parcels without an address get queried).
    """
    failed = 0
    for i in range(0, len(parcel_ids), ADDR_CHUNK):
        chunk = parcel_ids[i:i + ADDR_CHUNK]
        in_list = ",".join(f"'{p}'" for p in chunk)
        params = {
            "where": f"PARCELID IN ({in_list})",
            "outFields": ",".join(PARCEL_FIELDS),
            "returnGeometry": "true",
            "outSR": 4326,
            "f": "json",
        }
        out: dict[str, dict] = {}
        try:
            data = fetch_json(PARCELS_URL, params)
        except RuntimeError as e:
            failed += len(chunk)
            print(f"  WARNING: chunk failed, will retry next run ({e})")
            continue
        for f in data.get("features", []):
            a = f["attributes"]
            lon, lat = polygon_centroid((f.get("geometry") or {}).get("rings"))
            a["_LON"], a["_LAT"] = lon, lat
            out[a["PARCELID"]] = a
        done = min(i + ADDR_CHUNK, len(parcel_ids))
        print(f"  addresses: {done}/{len(parcel_ids)}")
        yield out
        time.sleep(REQUEST_DELAY)
    if failed:
        print(f"  {failed} parcels' addresses deferred to next run")


def ms_to_date(ms) -> str | None:
    """Esri epoch-milliseconds -> 'YYYY-MM-DD' (or None).

    Uses timedelta math instead of fromtimestamp(): Windows raises
    OSError on negative timestamps, and the county data contains bogus
    pre-1970 placeholder dates. Junk dates return None.
    """
    if ms is None:
        return None
    try:
        d = dt.datetime(1970, 1, 1) + dt.timedelta(milliseconds=ms)
    except OverflowError:
        return None
    if not (1900 <= d.year <= 2100):  # placeholder/typo dates
        return None
    return d.date().isoformat()


def query_permits(max_pages: int | None = None):
    """Yield residential NEW + BASEMENT FINISH permits (with lat/lon)."""
    since = dt.date.today().replace(
        year=dt.date.today().year - PERMIT_LOOKBACK_YEARS).isoformat()
    where = ("PERMITTYPE IN ('NEW','BASEMENT FINISH') "
             "AND PERMITCLASSIFICATION = 'RESIDENTIAL' "
             "AND PERMITSTATUS <> 'Void' "
             f"AND PERMITDATE >= DATE '{since}'")
    offset, page = 0, 0
    while True:
        params = {
            "where": where,
            "outFields": ",".join(PERMIT_FIELDS),
            "returnGeometry": "true",
            "outSR": 4326,
            "orderByFields": "OBJECTID_1",
            "resultOffset": offset,
            "resultRecordCount": PAGE_SIZE,
            "f": "json",
        }
        data = fetch_json(PERMITS_URL, params)
        feats = data.get("features", [])
        for f in feats:
            a = f["attributes"]
            g = f.get("geometry") or {}
            a["_LON"], a["_LAT"] = g.get("x"), g.get("y")
            yield a
        page += 1
        print(f"  permits page {page}: {len(feats)}")
        if not data.get("exceededTransferLimit") or not feats:
            break
        if max_pages and page >= max_pages:
            break
        offset += PAGE_SIZE
        time.sleep(REQUEST_DELAY)


# --------------------------------------------------------------------------
# Database
# --------------------------------------------------------------------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS leads (
    building_key   TEXT PRIMARY KEY,   -- PARCELID:IMP_NO
    parcelid       TEXT NOT NULL,
    year_built     INTEGER,
    basement_sf    REAL,
    finished_sf    REAL,
    unfinished_sf  REAL,
    above_grade_sf REAL,
    garage_sf      REAL,
    baths          REAL,
    kitchens       INTEGER,
    adu            TEXT,
    bldg_group     TEXT,
    style          TEXT,
    quality        TEXT,
    condition      TEXT,
    exterior       TEXT,
    pct_complete   REAL,
    bldg_cost_rcnld REAL,
    tax_year       INTEGER,
    acct_type      TEXT,
    address        TEXT,
    city           TEXT,
    zip            TEXT,
    owner          TEXT,
    owner_mailing  TEXT,
    use_descr      TEXT,
    subdivision    TEXT,
    acreage        REAL,
    school_district TEXT,
    land_value     REAL,
    improvement_value REAL,
    market_value   REAL,
    taxable_value  REAL,
    lat            REAL,
    lon            REAL,
    first_seen     TEXT NOT NULL,      -- date this building first matched
    last_seen      TEXT NOT NULL,      -- date it last matched
    status         TEXT NOT NULL DEFAULT 'active',  -- active | dropped
    dropped_date   TEXT
);
CREATE INDEX IF NOT EXISTS idx_leads_parcel ON leads(parcelid);

CREATE TABLE IF NOT EXISTS permits (
    permit_key     TEXT PRIMARY KEY,   -- PERMITNO:ACCOUNTNO
    permitno       TEXT,
    local_permitno TEXT,
    parcelid       TEXT,               -- = ACCOUNTNO (parcel serial)
    permit_type    TEXT,               -- NEW | BASEMENT FINISH
    classification TEXT,
    permit_date    TEXT,
    work_date      TEXT,
    co_date        TEXT,               -- certificate of occupancy
    status         TEXT,
    lat            REAL,
    lon            REAL,
    address        TEXT,
    city           TEXT,
    zip            TEXT,
    owner          TEXT,
    subdivision    TEXT,
    first_seen     TEXT NOT NULL,
    last_seen      TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_permits_parcel ON permits(parcelid);
"""


def open_db() -> sqlite3.Connection:
    con = sqlite3.connect(DB_FILE)
    con.executescript(SCHEMA)
    # lightweight migrations for databases created by older versions
    cols = {r[1] for r in con.execute("PRAGMA table_info(permits)")}
    if "co_date" not in cols:
        con.execute("ALTER TABLE permits ADD COLUMN co_date TEXT")
        con.commit()
    return con


def upsert_building(con, attrs: dict, today: str) -> bool:
    """Insert/update one building. Returns True if it is new."""
    key = f"{attrs['PARCELID']}:{attrs['IMP_NO']}"
    row = con.execute("SELECT status FROM leads WHERE building_key=?", (key,)).fetchone()
    vals = (
        attrs["PARCELID"], attrs.get("YEAR_BUILT"), attrs.get("BASEMENT"),
        attrs.get("BSMT_FINISH"), attrs.get("BSMT_UNFINISH"),
        attrs.get("ABOVE_GRADE_AREA"), attrs.get("ATTACHED_GARAGE"),
        attrs.get("BATHS"), attrs.get("KITCHENS"), attrs.get("ADU"),
        attrs.get("COST_GROUP_DESCR"), attrs.get("STYLE_DESCR"),
        attrs.get("QUALITY_DESCR"), attrs.get("CONDITION_DESCR"),
        attrs.get("EXTERIOR_WALL"), attrs.get("PCT_COMPLETE"),
        attrs.get("RCNLD"), attrs.get("TAX_YEAR"), attrs.get("TXACCTTYPE"),
    )
    if row is None:
        con.execute(
            """INSERT INTO leads (building_key, parcelid, year_built, basement_sf,
               finished_sf, unfinished_sf, above_grade_sf, garage_sf, baths,
               kitchens, adu, bldg_group, style, quality, condition, exterior,
               pct_complete, bldg_cost_rcnld, tax_year, acct_type,
               first_seen, last_seen, status)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, 'active')""",
            (key, *vals, today, today),
        )
        return True
    con.execute(
        """UPDATE leads SET parcelid=?, year_built=?, basement_sf=?, finished_sf=?,
           unfinished_sf=?, above_grade_sf=?, garage_sf=?, baths=?, kitchens=?,
           adu=?, bldg_group=?, style=?, quality=?, condition=?, exterior=?,
           pct_complete=?, bldg_cost_rcnld=?, tax_year=?, acct_type=?,
           last_seen=?, status='active', dropped_date=NULL
           WHERE building_key=?""",
        (*vals, today, key),
    )
    return False


def upsert_permit(con, attrs: dict, today: str) -> bool:
    """Insert/update one permit. Returns True if new."""
    key = f"{attrs.get('PERMITNO')}:{attrs.get('ACCOUNTNO')}"
    row = con.execute("SELECT 1 FROM permits WHERE permit_key=?", (key,)).fetchone()
    vals = (
        attrs.get("PERMITNO"), attrs.get("LOCALPERMITNO"),
        attrs.get("ACCOUNTNO"), attrs.get("PERMITTYPE"),
        attrs.get("PERMITCLASSIFICATION"), ms_to_date(attrs.get("PERMITDATE")),
        ms_to_date(attrs.get("PERMITWORKDATE")),
        ms_to_date(attrs.get("CERTIFICATEOFOCCUPANCYDATE")),
        attrs.get("PERMITSTATUS"),
        attrs.get("_LAT"), attrs.get("_LON"),
    )
    if row is None:
        con.execute(
            """INSERT INTO permits (permit_key, permitno, local_permitno,
               parcelid, permit_type, classification, permit_date, work_date,
               co_date, status, lat, lon, first_seen, last_seen)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (key, *vals, today, today))
        return True
    con.execute(
        """UPDATE permits SET permitno=?, local_permitno=?, parcelid=?,
           permit_type=?, classification=?, permit_date=?, work_date=?,
           co_date=?, status=?, lat=?, lon=?, last_seen=?
           WHERE permit_key=?""",
        (*vals, today, key))
    return False


# --------------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------------
COLUMNS = [
    ("Parcel ID", "parcelid", 12),
    ("Address", "address", 30),
    ("City", "city", 16),
    ("Zip", "zip", 7),
    ("Subdivision", "subdivision", 22),
    ("Owner", "owner", 26),
    ("Owner Mailing Addr", "owner_mailing", 32),
    ("Building Type", "bldg_group", 22),
    ("Year Built", "year_built", 10),
    ("Basement SF", "basement_sf", 12),
    ("Finished SF", "finished_sf", 11),
    ("Unfinished SF", "unfinished_sf", 13),
    ("% Unfinished", "pct_unfinished", 12),
    ("Above Grade SF", "above_grade_sf", 14),
    ("Garage SF", "garage_sf", 10),
    ("Baths", "baths", 7),
    ("Kitchens", "kitchens", 8),
    ("ADU", "adu", 5),
    ("Style", "style", 14),
    ("Quality", "quality", 10),
    ("Condition", "condition", 10),
    ("Exterior", "exterior", 20),
    ("Acreage", "acreage", 9),
    ("Land Value", "land_value", 11),
    ("Improvement Value", "improvement_value", 15),
    ("Total Market Value", "market_value", 15),
    ("Taxable Value", "taxable_value", 12),
    ("School District", "school_district", 15),
    ("Use", "use_descr", 22),
    ("Lat", "lat", 10),
    ("Lon", "lon", 11),
    ("Map Pin", "map_url", 9),
    ("New Constr Permit", "new_permit", 16),
    ("Bsmt Finish Permit", "finish_permit", 16),
    ("First Seen", "first_seen", 11),
    ("Last Seen", "last_seen", 11),
    ("Status", "status", 9),
    ("County Record", "appraisal_url", 14),
]
PERMIT_COLUMNS = [
    ("Parcel ID", "parcelid", 12),
    ("Address", "address", 30),
    ("City", "city", 16),
    ("Zip", "zip", 7),
    ("Subdivision", "subdivision", 22),
    ("Owner", "owner", 26),
    ("Permit No", "permitno", 11),
    ("Local Permit No", "local_permitno", 14),
    ("Permit Date", "permit_date", 12),
    ("Work Date", "work_date", 12),
    ("CO Date", "co_date", 12),
    ("Status", "status", 12),
    ("Lat", "lat", 10),
    ("Lon", "lon", 11),
    ("Map Pin", "map_url", 9),
    ("First Seen", "first_seen", 11),
    ("Became Lead", "became_lead", 12),
    ("County Record", "appraisal_url", 14),
]
LINK_COLS = {"map_url": "pin", "appraisal_url": "view"}
MAP_URL = "https://www.google.com/maps/search/?api=1&query={lat},{lon}"


def _decorate(r: dict):
    """Add computed link/rounding fields shared by all sheets."""
    r["appraisal_url"] = APPRAISAL_URL.format(pid=r["parcelid"])
    if r.get("lat") is not None and r.get("lon") is not None:
        r["lat"], r["lon"] = round(r["lat"], 6), round(r["lon"], 6)
        r["map_url"] = MAP_URL.format(lat=r["lat"], lon=r["lon"])
    else:
        r["map_url"] = None


def export_xlsx(con, today: str):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    rows = [dict(r) for r in con.execute(
        "SELECT * FROM leads ORDER BY status, city, address").fetchall()]
    # most recent permit per parcel, per type -> flag on lead rows
    finish_permits = dict(con.execute(
        """SELECT parcelid, MAX(permit_date) FROM permits
           WHERE permit_type='BASEMENT FINISH' GROUP BY parcelid"""))
    new_permits = dict(con.execute(
        """SELECT parcelid, MAX(permit_date) FROM permits
           WHERE permit_type='NEW' GROUP BY parcelid"""))
    for r in rows:
        b, u = r["basement_sf"] or 0, r["unfinished_sf"] or 0
        r["pct_unfinished"] = round(100 * u / b, 1) if b else None
        r["finish_permit"] = finish_permits.get(r["parcelid"])
        r["new_permit"] = new_permits.get(r["parcelid"])
        _decorate(r)

    # permit parcel now exists as an active lead -> conversion date
    lead_seen = dict(con.execute(
        """SELECT parcelid, MIN(first_seen) FROM leads
           WHERE status='active' GROUP BY parcelid"""))
    permits = [dict(r) for r in con.execute(
        "SELECT * FROM permits ORDER BY permit_date DESC").fetchall()]
    for r in permits:
        r["became_lead"] = lead_seen.get(r["parcelid"])
        _decorate(r)

    sheets = {
        "All Active Leads":
            ([r for r in rows if r["status"] == "active"], COLUMNS),
        "New This Run":
            ([r for r in rows if r["first_seen"] == today
              and r["status"] == "active"], COLUMNS),
        "Dropped This Run":
            ([r for r in rows if r["dropped_date"] == today], COLUMNS),
        "New Construction Permits":
            ([r for r in permits if r["permit_type"] == "NEW"], PERMIT_COLUMNS),
        "Basement Finish Permits":
            ([r for r in permits if r["permit_type"] == "BASEMENT FINISH"],
             PERMIT_COLUMNS),
    }

    wb = Workbook()
    wb.remove(wb.active)
    for name, (data, cols) in sheets.items():
        ws = wb.create_sheet(name)
        ws.append([c[0] for c in cols])
        for r in data:
            ws.append([r.get(c[1]) for c in cols])
        # hyperlink columns (map pin, county record)
        for ci, c in enumerate(cols, 1):
            if c[1] not in LINK_COLS:
                continue
            for i in range(2, len(data) + 2):
                cell = ws.cell(row=i, column=ci)
                if cell.value:
                    cell.hyperlink = cell.value
                    cell.value = LINK_COLS[c[1]]
                    cell.font = Font(color="0563C1", underline="single")
        for i, c in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = c[2]
        if data:
            ref = f"A1:{get_column_letter(len(cols))}{len(data) + 1}"
            t = Table(displayName=name.replace(" ", ""), ref=ref)
            t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2",
                                              showRowStripes=True)
            ws.add_table(t)
        ws.freeze_panes = "A2"
    wb.save(XLSX_FILE)


# --------------------------------------------------------------------------
# Interactive map
# --------------------------------------------------------------------------
DROPS_HEADER = ["subdivision", "date", "who", "notes"]


def read_drops() -> list[dict]:
    """Read the literature-distribution log (creates template if missing)."""
    if not DROPS_FILE.exists():
        with open(DROPS_FILE, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow(DROPS_HEADER)
        return []
    with open(DROPS_FILE, newline="", encoding="utf-8") as f:
        return [{k.strip().lower(): (v or "").strip() for k, v in row.items()}
                for row in csv.DictReader(f)]


def _owner_occupied(site: str | None, mailing: str | None) -> int | None:
    """1 = owner's mailing address matches the site (owner-occupied),
    0 = differs (builder/investor/absentee), None = unknown."""
    if not site or not mailing:
        return None
    def toks(s):
        return "".join(ch if ch.isalnum() or ch.isspace() else " "
                       for ch in s.upper()).split()
    a, b = toks(site), toks(mailing)
    if not a or not b:
        return None
    return 1 if a[:4] == b[:4] else 0


def export_map(con):
    """Write a self-contained filterable Leaflet map of all active leads."""
    rows = con.execute(
        """SELECT parcelid, address, city, zip, subdivision, owner,
                  owner_mailing, year_built, basement_sf, finished_sf,
                  unfinished_sf, market_value, lat, lon, first_seen
           FROM leads WHERE status='active' AND lat IS NOT NULL""").fetchall()
    finish_permits = dict(con.execute(
        """SELECT parcelid, MAX(permit_date) FROM permits
           WHERE permit_type='BASEMENT FINISH' GROUP BY parcelid"""))
    leads = []
    for r in rows:
        (pid, addr, city, zipc, subdiv, owner, mailing, yb, b, fin, unf,
         mval, lat, lon, seen) = r
        leads.append({
            "p": pid, "a": addr, "c": city, "z": zipc,
            "s": (subdiv or "(NO SUBDIVISION)").strip().upper(),
            "o": owner, "y": yb, "b": b, "f": fin, "u": unf,
            "v": mval, "la": round(lat, 6), "lo": round(lon, 6),
            "fs": seen, "fp": finish_permits.get(pid),
            "oo": _owner_occupied(addr, mailing),
        })
    # NEW-construction permits (future leads, ahead of the assessor roll)
    perm_rows = con.execute(
        """SELECT parcelid, address, city, subdivision, owner,
                  permit_date, co_date, status, lat, lon
           FROM permits WHERE permit_type='NEW' AND lat IS NOT NULL""").fetchall()
    lead_parcels = {l["p"] for l in leads}
    min_pdate = dt.date.today().replace(
        year=dt.date.today().year - PERMIT_LOOKBACK_YEARS).isoformat()
    perms = []
    for (pid, addr, city, subdiv, owner, pdate, cdate, stat, lat, lon) in perm_rows:
        if pid in lead_parcels:   # already an active lead; one pin is enough
            continue
        if not pdate or pdate < min_pdate:  # stale/bogus-dated permits
            continue
        perms.append({
            "p": pid, "a": addr, "c": city,
            "s": (subdiv or "(NO SUBDIVISION)").strip().upper(),
            "o": owner, "d": pdate, "cd": cdate, "st": stat,
            "la": round(lat, 6), "lo": round(lon, 6),
        })
    drops = read_drops()
    html = (MAP_TEMPLATE
            .replace("__GENERATED__", dt.date.today().isoformat())
            .replace("__LEADS__", json.dumps(leads, separators=(",", ":")))
            .replace("__PERMS__", json.dumps(perms, separators=(",", ":")))
            .replace("__DROPS__", json.dumps(drops, separators=(",", ":"))))
    MAP_FILE.write_text(html, encoding="utf-8")


MAP_TEMPLATE = r"""<!DOCTYPE html>
<html><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>Basement Leads Map</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet.markercluster/1.5.3/MarkerCluster.min.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet.markercluster/1.5.3/MarkerCluster.Default.min.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet.markercluster/1.5.3/leaflet.markercluster.min.js"></script>
<style>
 html,body{margin:0;height:100%;font-family:system-ui,sans-serif}
 #wrap{display:flex;height:100%}
 #side{width:300px;min-width:300px;overflow-y:auto;padding:12px;
       background:#f7f7f7;border-right:1px solid #ddd;font-size:13px}
 #map{flex:1}
 h1{font-size:15px;margin:0 0 8px}
 label{display:block;margin-top:8px;font-weight:600}
 select,input{width:100%;box-sizing:border-box;padding:4px;margin-top:2px}
 .row{display:flex;gap:6px}.row>div{flex:1}
 #stats{margin:10px 0;padding:8px;background:#fff;border:1px solid #ddd;border-radius:4px}
 .legend span{display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:4px}
 #subs{margin-top:10px}
 .sub{padding:5px 6px;border-bottom:1px solid #e5e5e5;cursor:pointer}
 .sub:hover{background:#eef}
 .sub small{color:#666;display:block}
 .never{color:#c0392b;font-weight:600}
 @media(max-width:700px){#wrap{flex-direction:column}
   #side{width:auto;min-width:0;max-height:45%;border-right:none;border-bottom:1px solid #ddd}}
</style></head><body>
<div id="wrap"><div id="side">
 <h1>Basement Leads <small style="font-weight:400">(data __GENERATED__)</small></h1>
 <label>City <select id="fCity"><option value="">All cities</option></select></label>
 <div class="row">
  <div><label>Built from <input id="fY1" type="number" placeholder="2011"></label></div>
  <div><label>to <input id="fY2" type="number" placeholder="2026"></label></div>
 </div>
 <label>Min unfinished SF <input id="fSF" type="number" value="0"></label>
 <label>Min % of basement unfinished <input id="fPct" type="number" value="0" min="0" max="100"></label>
 <label>Occupancy
  <select id="fOcc"><option value="">All owners</option>
   <option value="own">Owner-occupied</option>
   <option value="abs">Absentee / builder-owned</option></select></label>
 <label>Basement-finish permit
  <select id="fPerm"><option value="">Show all</option>
   <option value="no">Hide (permit pulled)</option>
   <option value="only">Only (permit pulled)</option></select></label>
 <label>Hide subdivisions dropped in last
  <input id="fMonths" type="number" value="0" min="0"> months (0 = show all)</label>
 <label style="margin-top:10px"><input type="checkbox" id="fPerms" style="width:auto">
  Show upcoming construction (<span id="permCount"></span> NEW permits not yet on the assessor roll)</label>
 <div id="stats"></div>
 <div class="legend">
  <span style="background:#c0392b"></span>never dropped&nbsp;
  <span style="background:#e67e22"></span>&gt;6 mo&nbsp;
  <span style="background:#f1c40f"></span>3&ndash;6 mo&nbsp;
  <span style="background:#27ae60"></span>&lt;3 mo&nbsp;
  <span style="background:#2980b9"></span>upcoming constr.
 </div>
 <div id="subs"><label>Subdivisions <input id="subSearch" placeholder="search..."></label>
  <div id="subList"></div></div>
</div><div id="map"></div></div>
<script>
const LEADS=__LEADS__;
const PERMS=__PERMS__;
let DROPS=__DROPS__;
const fmt=n=>n==null?"-":Number(n).toLocaleString();

// last drop date per subdivision (log may hold many rounds; take latest)
let lastDrop={},dropRounds={};
function indexDrops(){
  lastDrop={};dropRounds={};
  for(const d of DROPS){
    const s=(d.subdivision||"").trim().toUpperCase();
    if(!s||!d.date)continue;
    dropRounds[s]=(dropRounds[s]||0)+1;
    if(!lastDrop[s]||d.date>lastDrop[s])lastDrop[s]=d.date;
  }
}
indexDrops();
// live-reload the drop log when hosted (github pages etc.); file:// will just skip
fetch("drops.csv",{cache:"no-store"}).then(r=>r.ok?r.text():null).then(t=>{
  if(!t)return;
  const lines=t.split(/\r?\n/).filter(x=>x.trim());
  const hdr=lines.shift().split(",").map(x=>x.trim().toLowerCase());
  DROPS=lines.map(l=>{const c=l.split(",");const o={};
    hdr.forEach((h,i)=>o[h]=(c[i]||"").trim());return o;});
  indexDrops();refresh();
}).catch(()=>{});

function monthsAgo(iso){
  if(!iso)return null;
  return (Date.now()-new Date(iso).getTime())/2629800000;
}
function colorFor(sub){
  const m=monthsAgo(lastDrop[sub]);
  if(m==null)return"#c0392b";
  if(m>6)return"#e67e22";
  if(m>3)return"#f1c40f";
  return"#27ae60";
}

const map=L.map("map").setView([40.25,-111.7],10);
L.tileLayer("https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png",
  {subdomains:"abcd",maxZoom:20,
   attribution:"&copy; OpenStreetMap contributors &copy; CARTO"}).addTo(map);
let cluster=L.markerClusterGroup({maxClusterRadius:55});
map.addLayer(cluster);

const cities=[...new Set(LEADS.map(l=>l.c).filter(Boolean))].sort();
for(const c of cities){const o=document.createElement("option");o.value=o.textContent=c;fCity.appendChild(o);}

const pctUnf=l=>l.b?100*(l.u||0)/l.b:0;
function visible(){
  const city=fCity.value,y1=+fY1.value||0,y2=+fY2.value||9999,
        sf=+fSF.value||0,pct=+fPct.value||0,occ=fOcc.value,
        perm=fPerm.value,hideM=+fMonths.value||0;
  return LEADS.filter(l=>
    (!city||l.c===city)&&(l.y||0)>=y1&&(l.y||0)<=y2&&(l.u||0)>=sf&&
    pctUnf(l)>=pct&&
    (occ!=="own"||l.oo===1)&&(occ!=="abs"||l.oo===0)&&
    (perm!=="no"||!l.fp)&&(perm!=="only"||l.fp)&&
    (!hideM||(monthsAgo(lastDrop[l.s])==null||monthsAgo(lastDrop[l.s])>hideM)));
}

function refresh(){
  const vis=visible();
  cluster.clearLayers();
  for(const l of vis){
    const m=L.circleMarker([l.la,l.lo],{radius:7,weight:1,color:"#333",
      fillColor:colorFor(l.s),fillOpacity:.9});
    const pc=Math.round(pctUnf(l));
    const occ=l.oo===1?"Owner-occupied":l.oo===0?"Absentee/builder owner":"Occupancy unknown";
    m.bindTooltip(`${l.y} &middot; ${fmt(l.u)} SF unf (${pc}%)`);
    m.bindPopup(`<b>${l.a||"(no address)"}</b><br>${l.o||""} <i>(${occ})</i><br>
      Built ${l.y} &middot; Bsmt ${fmt(l.b)} SF (${fmt(l.u)} unfinished = ${pc}%)<br>
      Value $${fmt(l.v)}<br>Subdivision: ${l.s}<br>
      Last dropped: ${lastDrop[l.s]||"<span class=never>never</span>"}
      ${l.fp?`<br><b style="color:#c0392b">Finish permit ${l.fp}</b>`:""}<br>
      <a target="_blank" href="https://www.utahcounty.gov/LandRecords/AppraisalInfo.asp?avParcelId=${l.p}">county record</a> &middot;
      <a target="_blank" href="https://www.google.com/maps/search/?api=1&query=${l.la},${l.lo}">google maps</a>`);
    cluster.addLayer(m);
  }
  let permShown=0;
  if(fPerms.checked){
    const city=fCity.value,hideM=+fMonths.value||0;
    for(const l of PERMS){
      if(city&&l.c!==city)continue;
      if(hideM&&monthsAgo(lastDrop[l.s])!=null&&monthsAgo(lastDrop[l.s])<=hideM)continue;
      permShown++;
      const m=L.circleMarker([l.la,l.lo],{radius:7,weight:1,color:"#1a5276",
        fillColor:"#2980b9",fillOpacity:.9});
      m.bindTooltip(`NEW permit ${l.d||""}${l.cd?" &middot; CO "+l.cd:""}`);
      m.bindPopup(`<b>${l.a||"(no address)"}</b><br>${l.o||""}<br>
        <b style="color:#2980b9">Upcoming construction</b> &middot; permit ${l.d||"?"}<br>
        ${l.cd?`Occupancy cert: ${l.cd}<br>`:""}Status: ${l.st||"?"}<br>
        Subdivision: ${l.s}<br>
        Last dropped: ${lastDrop[l.s]||"<span class=never>never</span>"}<br>
        <a target="_blank" href="https://www.utahcounty.gov/LandRecords/AppraisalInfo.asp?avParcelId=${l.p}">county record</a> &middot;
        <a target="_blank" href="https://www.google.com/maps/search/?api=1&query=${l.la},${l.lo}">google maps</a>`);
      cluster.addLayer(m);
    }
  }
  stats.innerHTML=`<b>${vis.length.toLocaleString()}</b> of ${LEADS.length.toLocaleString()} leads shown`
    +(fPerms.checked?`<br><b>${permShown.toLocaleString()}</b> upcoming-construction permits`:"");
  renderSubs(vis);
}

function renderSubs(vis){
  const q=(subSearch.value||"").toUpperCase();
  const bySub={};
  for(const l of vis)(bySub[l.s]=bySub[l.s]||[]).push(l);
  const names=Object.keys(bySub).sort((a,b)=>bySub[b].length-bySub[a].length)
    .filter(s=>!q||s.includes(q)).slice(0,150);
  subList.innerHTML=names.map(s=>{
    const d=lastDrop[s];
    return`<div class="sub" data-s="${s}"><b>${s}</b> (${bySub[s].length})
      <small>${d?`last drop ${d} &middot; ${dropRounds[s]} round(s)`:`<span class="never">never dropped</span>`}</small></div>`;
  }).join("");
  for(const el of subList.querySelectorAll(".sub"))
    el.onclick=()=>{const ls=bySub[el.dataset.s];
      map.fitBounds(ls.map(l=>[l.la,l.lo]),{maxZoom:16});};
}

for(const el of[fCity,fY1,fY2,fSF,fPct,fOcc,fPerm,fMonths,fPerms])el.onchange=refresh;
subSearch.oninput=()=>renderSubs(visible());
permCount.textContent=PERMS.length.toLocaleString();
refresh();
</script></body></html>
"""


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--limit", type=int, default=None,
                    help="max pages to fetch (test mode; 1 page = 2000 buildings)")
    args = ap.parse_args()

    today = dt.date.today().isoformat()
    cutoff = dt.date.today().year - YEARS_BACK
    where = f"YEAR_BUILT >= {cutoff} AND BSMT_UNFINISH >= {MIN_UNFINISHED_SF}"
    if RESIDENTIAL_ONLY:
        where += " AND TXACCTTYPE = 'RESIDENTIAL'"

    print(f"Run date: {today}")
    print(f"Query: {where}\n")

    con = open_db()
    con.row_factory = sqlite3.Row

    print("Fetching buildings...")
    new_count = total = 0
    for attrs in query_buildings(where, max_pages=args.limit):
        # buildings with multiple sections repeat per section; keep primary
        if attrs.get("OVERALL_RN") not in (None, 1):
            continue
        if upsert_building(con, attrs, today):
            new_count += 1
        total += 1
    con.commit()

    # Anything not seen today no longer matches (finished, demolished, etc.)
    # Skipped in test mode (--limit) because we didn't fetch everything.
    dropped = 0
    if not args.limit:
        cur = con.execute(
            """UPDATE leads SET status='dropped', dropped_date=?
               WHERE last_seen < ? AND status='active'""", (today, today))
        dropped = cur.rowcount
        con.commit()

    # Address enrichment: only parcels we don't have an address for yet.
    missing = [r[0] for r in con.execute(
        "SELECT DISTINCT parcelid FROM leads WHERE address IS NULL AND status='active'")]
    if missing:
        print(f"\nLooking up addresses for {len(missing)} parcels...")
        for info in query_addresses(missing):
            for pid, a in info.items():
                con.execute(
                    """UPDATE leads SET address=?, city=?, zip=?, owner=?,
                       owner_mailing=?, use_descr=?, subdivision=?, acreage=?,
                       school_district=?, land_value=?, improvement_value=?,
                       market_value=?, taxable_value=?, lat=?, lon=?
                       WHERE parcelid=?""",
                    (a.get("SITE_FULLADDRESS"), a.get("SITE_CITY"),
                     a.get("SITE_ZIP5"), a.get("OWNERNAME"),
                     a.get("OWN_FULLADDRESS"), a.get("USEDSCRP"),
                     a.get("CNVYNAME"), a.get("ACREAGE"), a.get("SCHLDSCRP"),
                     a.get("MKT_LANDVALUE"), a.get("MKT_IMPVALUE"),
                     a.get("MKT_CNTVAL"), a.get("TXBL_CNTVAL"),
                     a.get("_LAT"), a.get("_LON"), pid))
            con.commit()  # progress saved per chunk; a crash loses nothing

    new_permits = 0
    if INCLUDE_PERMITS:
        print("\nFetching permits (residential NEW + BASEMENT FINISH)...")
        for attrs in query_permits(max_pages=args.limit):
            if upsert_permit(con, attrs, today):
                new_permits += 1
        con.commit()
        missing_p = [r[0] for r in con.execute(
            """SELECT DISTINCT parcelid FROM permits
               WHERE address IS NULL AND parcelid IS NOT NULL""")]
        if missing_p:
            print(f"Looking up addresses for {len(missing_p)} permit parcels...")
            for info in query_addresses(missing_p):
                for pid, a in info.items():
                    con.execute(
                        """UPDATE permits SET address=?, city=?, zip=?,
                           owner=?, subdivision=? WHERE parcelid=?""",
                        (a.get("SITE_FULLADDRESS"), a.get("SITE_CITY"),
                         a.get("SITE_ZIP5"), a.get("OWNERNAME"),
                         a.get("CNVYNAME"), pid))
                con.commit()

    print("\nWriting Excel...")
    export_xlsx(con, today)
    print("Writing map...")
    export_map(con)

    active = con.execute("SELECT COUNT(*) FROM leads WHERE status='active'").fetchone()[0]
    n_perm = con.execute("SELECT COUNT(*) FROM permits").fetchone()[0]
    print(f"""
Done.
  Buildings matched this run : {total}
  New since last run         : {new_count}
  Dropped (no longer match)  : {dropped}
  Total active leads         : {active}
  Permits tracked            : {n_perm} ({new_permits} new)
  Output                     : {XLSX_FILE}
  Map                        : {MAP_FILE}
  Drop log                   : {DROPS_FILE}
  Database                   : {DB_FILE}
""")
    con.close()


if __name__ == "__main__":
    sys.exit(main())
