"""Build the data files that power Charter Compass NYC.

Sources (all public, downloaded fresh on every run):
  * NYSED Charter School Directory, 2025-26 (Excel)        -> roster, district, year, authorizer
  * NCES Common Core of Data via Urban Institute API       -> address, coordinates, grades, enrollment, teachers,
                                                             5-year enrollment growth, district charter share
  * NYC Open Data: Borough Boundaries (gthc-hcne)          -> map outlines
  * NYC Open Data: School Districts (8ugf-3d8u)            -> CSD outlines, fallback placement

Writes:
  data/schools.json, data/boroughs.geojson, data/districts.geojson

Usage:
  python3 -m venv .venv && .venv/bin/pip install openpyxl shapely
  .venv/bin/python scripts/build_data.py
"""

import hashlib
import io
import json
import re
import urllib.request
from datetime import date
from pathlib import Path

import openpyxl
from shapely.geometry import Point, mapping, shape

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data"

NYSED_XLSX = "https://www.nysed.gov/sites/default/files/programs/charter-schools/nys-charter-school-directory-as-of-09-02-2025.xlsx"
NYSED_PAGE = "https://www.nysed.gov/charter-schools/charter-schools-directory"
CCD_API = "https://educationdata.urban.org/api/v1/schools/ccd/directory/{year}/?fips=36&charter=1"
CCD_ALL = "https://educationdata.urban.org/api/v1/schools/ccd/directory/{year}/?fips=36"
CCD_HISTORY = range(2019, 2025)  # 2024 = the 2024-25 school year
CCD_YEARS = (2022, 2023, 2024)  # where location fields come from; later years win
SHARE_YEAR = 2024
MIN_TREND_BASE = 100  # don't measure growth from a tiny founding class

# Quality checks on federal (CCD) figures, same rules as CharterCompassNYC.
CCD_CLOSED = 2
RATIO_RANGE = (4, 30)  # students per FTE teacher outside this is a reporting error
MIN_TEACHERS = 2
REVIEWED_INCOMPLETE = {
    "321200860866": "94 students spread across grades 2–11",  # Girls Prep NY; students appear under sibling Public Prep charters
}
BOROUGHS = "https://data.cityofnewyork.us/resource/gthc-hcne.geojson?$limit=10"
DISTRICTS = "https://data.cityofnewyork.us/resource/8ugf-3d8u.geojson?$limit=100"

AUTH = {"SUNY": "S", "Regents": "R", "NYCDOE": "D"}

# Network = a family of schools sharing a name. Inferred from names only; order matters.
NETWORKS = [
    (r"^Success Academy", "Success Academy"),
    (r"^Achievement First", "Achievement First"),
    (r"^Uncommon", "Uncommon Schools"),
    (r"^KIPP", "KIPP NYC"),
    (r"^Icahn", "Icahn Charter Schools"),
    (r"Ascend", "Ascend Public Charter Schools"),
    (r"^Explore", "Explore Schools"),
    (r"^Brilla", "Brilla Public Charter Schools"),
    (r"^Zeta", "Zeta Charter Schools"),
    (r"^Democracy Prep", "Democracy Prep"),
    (r"^Harlem Village Academy", "Harlem Village Academies"),
    (r"^Harlem Children's Zone", "Harlem Children's Zone"),
    (r"^South Bronx Classical", "South Bronx Classical"),
    (r"^Bronx Charter School for Excellence", "Bronx Excellence"),
    (r"^Bronx Charter School for Better Learning", "Bronx Better Learning"),
    (r"^Urban Dove", "Urban Dove Team"),
    (r"^United Charter High School", "United Charter High Schools"),
    (r"^Family Life Academy", "Family Life Academy"),
    (r"^Brooklyn Prospect", "Brooklyn Prospect"),
    (r"^(Girls|Boys) Preparatory", "Public Prep"),
    (r"^Amber", "Amber Charter Schools"),
    (r"^Dream Charter", "DREAM"),
    (r"^Our World Neighborhood", "Our World Neighborhood"),
    (r"Hebrew (Language Academy|Public)", "Hebrew Public"),
    (r"^Hellenic Classical", "Hellenic Classical"),
    (r"^Growing Up Green", "Growing Up Green"),
    (r"Brooklyn Laboratory", "Brooklyn LAB"),
    (r"^Capital Prep", "Capital Prep"),
    (r"^East Harlem Scholars", "East Harlem Scholars"),
    (r"^Storefront", "Storefront Academy"),
    (r"^Manhattan Charter", "Manhattan Charter"),
    (r"^Lamad", "Lamad Academy"),
    (r"^NYC Autism", "NYC Autism Charter"),
    (r"^Central Queens Academy", "Central Queens Academy"),
    (r"^Renaissance Charter", "Renaissance"),
    (r"^New Dawn", "New Dawn"),
    (r"^Neighborhood Charter School", "Neighborhood Charter"),
    (r"^Math, Engineering,? and Science Academy", "MESA"),
    (r"^Bedford Stuyvesant New Beginnings", "New Beginnings"),
    (r"^American Dream", "American Dream"),
    (r"^University Prep", "University Prep"),
]


def fetch(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (charter-compass build)"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read()


def clean_name(name):
    name = re.sub(r"\s+", " ", name.replace("–", "-")).strip()
    m = re.match(r"^(.*) \(The\)$", name)
    return f"The {m.group(1)}" if m else name


def title(s):
    return re.sub(r"(\d)(St|Nd|Rd|Th)\b", lambda m: m.group(1) + m.group(2).lower(), s.title())


def borough(csd):
    if csd <= 6:
        return "Manhattan"
    if csd <= 12:
        return "Bronx"
    if csd <= 23 or csd == 32:
        return "Brooklyn"
    if csd <= 30:
        return "Queens"
    return "Staten Island"


def network(name):
    for pattern, label in NETWORKS:
        if re.search(pattern, name):
            return label
    return "Independent"


def simplify(fc, key, tol):
    feats = []
    for f in fc["features"]:
        geom = shape(f["geometry"]).simplify(tol, preserve_topology=True)
        gj = json.loads(json.dumps(mapping(geom)), parse_float=lambda x: round(float(x), 5))
        feats.append({"type": "Feature", "properties": {"id": f["properties"][key]}, "geometry": gj})
    return {"type": "FeatureCollection", "features": feats}


def point_in(poly, seed):
    """Deterministic pseudo-random point inside a polygon, so approximate pins never land in the water."""
    minx, miny, maxx, maxy = poly.bounds
    h = int(hashlib.sha1(seed.encode()).hexdigest(), 16)
    for i in range(500):
        u = ((h >> (i % 64)) % 10007) / 10007
        v = ((h >> ((i * 7 + 13) % 64)) % 10009) / 10009
        p = Point(minx + u * (maxx - minx), miny + v * (maxy - miny))
        if poly.contains(p):
            return p
        h = int(hashlib.sha1(f"{h}{i}".encode()).hexdigest(), 16)
    return poly.representative_point()


def ccd_all(url):
    """Follow the Urban API's pagination."""
    out = []
    while url:
        page = json.loads(fetch(url))
        out += page["results"]
        url = page.get("next")
    return out


def growth(points):
    """Enrollment change from five years back (or the first year with a real class) to the latest year."""
    if len(points) < 2:
        return None
    last = max(points)
    base = next((y for y in sorted(points) if y >= last - 5 and points[y] >= MIN_TREND_BASE), None)
    if base is None or base >= last:
        return None
    return {"since": base, "from": points[base], "to": points[last],
            "pct": round(100 * (points[last] - points[base]) / points[base])}


def district_share(rows):
    """Share of each CSD's public K-12 students who attend a charter, from every NYC public school in CCD."""
    out = {}
    for r in rows:
        code = r["seasch"].split("-")[-1]
        n = r.get("enrollment")
        if len(code) != 12 or code[:2] not in ("31", "32", "33", "34", "35") or not n or n <= 0:
            continue
        csd = int(code[2:4])
        if not 1 <= csd <= 32:
            continue
        d = out.setdefault(csd, {"enroll": 0, "charter": 0})
        d["enroll"] += n
        if code[6:8] == "86":
            d["charter"] += n
    for d in out.values():
        d["share"] = round(100 * d["charter"] / d["enroll"], 1)
    return {str(k): out[k] for k in sorted(out)}


def looks_like_shell(rec, opened):
    """An established school reporting one grade and a handful of students,
    e.g. a campus whose students moved to a sibling charter."""
    lo, hi, n = rec["lowest_grade_offered"], rec["highest_grade_offered"], rec["enrollment"]
    return rec["year"] - opened >= 2 and lo is not None and lo == hi and n is not None and 0 <= n < 50


def data_note(rec, sed, opened):
    if rec["school_status"] == CCD_CLOSED:
        return (f"Federal records list this school as closed in {rec['year']}–{str(rec['year'] + 1)[2:]}, "
                "so grades and enrollment aren't shown. The state directory still lists it; check with the school.")
    if looks_like_shell(rec, opened) or sed in REVIEWED_INCOMPLETE:
        why = REVIEWED_INCOMPLETE.get(sed, "one grade, a handful of students")
        return (f"Federal figures for this school look incomplete ({why}), likely because students "
                "are reported under a sibling campus. Check with the school.")
    return None


def plausible_ratio(enroll, teachers):
    return teachers >= MIN_TEACHERS and RATIO_RANGE[0] <= enroll / teachers <= RATIO_RANGE[1]


def grade(g):
    return {-1: "PK", 0: "K"}.get(g, str(g)) if g is not None and g >= -1 else None


def main():
    OUT.mkdir(exist_ok=True)

    wb = openpyxl.load_workbook(io.BytesIO(fetch(NYSED_XLSX)), read_only=True)
    rows = [r for r in list(wb.active.iter_rows(values_only=True))[2:] if r[6] == "New York City"]

    ccd, history = {}, {}
    for y in CCD_HISTORY:
        for r in ccd_all(CCD_API.format(year=y)):
            sid = r["seasch"].split("-")[-1]
            if r.get("enrollment") and r["enrollment"] > 0:
                history.setdefault(sid, {})[y] = r["enrollment"]
            if y in CCD_YEARS:
                ccd[sid] = r
    districts_share = district_share(ccd_all(CCD_ALL.format(year=SHARE_YEAR)))

    districts_raw = json.loads(fetch(DISTRICTS))
    csd_shapes = {int(f["properties"]["schooldist"]): shape(f["geometry"]) for f in districts_raw["features"]}

    schools = []
    for _, inst, sed, name, district, county, _, year, authorizer, *_ in rows:
        sed = str(sed)
        csd = int(re.search(r"\d+", district).group())
        name = clean_name(name)
        s = {
            "id": sed,
            "inst": str(inst),
            "name": name,
            "csd": csd,
            "boro": borough(csd),
            "county": county,
            "year": int(year),
            "auth": AUTH[authorizer],
            "network": network(name),
        }
        c = ccd.get(sed)
        if c and c.get("latitude"):
            enroll = c.get("enrollment")
            teachers = c.get("teachers_fte")
            s.update({
                "lat": round(c["latitude"], 6),
                "lon": round(c["longitude"], 6),
                "loc": "exact",
                "address": f"{title(c['street_location'])}, {title(c['city_location'])}, NY {c['zip_location']}",
                "grades": [grade(c.get("lowest_grade_offered")), grade(c.get("highest_grade_offered"))],
                "enroll": enroll if enroll and enroll > 0 else None,
                "teachers": round(teachers, 1) if enroll and enroll > 0 and teachers and plausible_ratio(enroll, teachers) else None,
                "ccdYear": c["year"],
                "nces": c["ncessch"],
            })
            note = data_note(c, sed, int(year))
            if note:
                # Keep the school and its pin, but don't show or count figures we don't trust.
                s.update({"grades": [None, None], "enroll": None, "teachers": None, "dataNote": note})
            else:
                g = growth(history.get(sed, {}))
                if g:
                    s["growth"] = g
        else:
            p = point_in(csd_shapes[csd], sed)
            s.update({"lat": round(p.y, 6), "lon": round(p.x, 6), "loc": "approx"})
        schools.append(s)

    sizes = {}
    for s in schools:
        sizes[s["network"]] = sizes.get(s["network"], 0) + 1
    for s in schools:
        s["netSize"] = sizes[s["network"]] if s["network"] != "Independent" else 1

    schools.sort(key=lambda s: s["name"])
    meta = {
        "built": date.today().isoformat(),
        "count": len(schools),
        "exact": sum(s["loc"] == "exact" for s in schools),
        "flagged": sum("dataNote" in s for s in schools),
        "districts": districts_share,
        "shareYear": SHARE_YEAR,
        "sources": {
            "roster": {"label": "NYSED Charter School Directory 2025-26 (as of Sep 2, 2025)", "url": NYSED_PAGE},
            "locations": {"label": "NCES Common Core of Data, school years 2022-23 to 2024-25", "url": "https://nces.ed.gov/ccd/"},
            "history": {"label": "NCES Common Core of Data, school years 2019-20 to 2024-25", "url": "https://educationdata.urban.org/"},
            "boundaries": {"label": "NYC Open Data: Borough Boundaries & School Districts", "url": "https://data.cityofnewyork.us/"},
        },
    }
    (OUT / "schools.json").write_text(json.dumps({"meta": meta, "schools": schools}, ensure_ascii=False, separators=(",", ":")))
    (OUT / "districts.geojson").write_text(json.dumps(simplify(districts_raw, "schooldist", 0.0004), separators=(",", ":")))
    (OUT / "boroughs.geojson").write_text(json.dumps(simplify(json.loads(fetch(BOROUGHS)), "boroname", 0.0004), separators=(",", ":")))
    print(f"{meta['count']} schools, {meta['exact']} with exact coordinates")


if __name__ == "__main__":
    main()
