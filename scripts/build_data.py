"""Build the data files that power Charter Compass NYC.

Sources (all public, downloaded fresh on every run):
  * NYSED Charter School Directory, 2025-26 (Excel)        -> roster, district, year, authorizer
  * NCES Common Core of Data via Urban Institute API       -> address, coordinates, grades, enrollment, teachers
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
CCD_YEARS = (2022, 2023, 2024)  # later years win
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


def grade(g):
    return {-1: "PK", 0: "K"}.get(g, str(g)) if g is not None and g >= -1 else None


def main():
    OUT.mkdir(exist_ok=True)

    wb = openpyxl.load_workbook(io.BytesIO(fetch(NYSED_XLSX)), read_only=True)
    rows = [r for r in list(wb.active.iter_rows(values_only=True))[2:] if r[6] == "New York City"]

    ccd = {}
    for y in CCD_YEARS:
        for r in json.loads(fetch(CCD_API.format(year=y)))["results"]:
            ccd[r["seasch"].split("-")[-1]] = r

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
                "teachers": round(teachers, 1) if teachers and teachers > 0 else None,
                "ccdYear": c["year"],
                "nces": c["ncessch"],
            })
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
        "sources": {
            "roster": {"label": "NYSED Charter School Directory 2025-26 (as of Sep 2, 2025)", "url": NYSED_PAGE},
            "locations": {"label": "NCES Common Core of Data, school years 2022-23 to 2024-25", "url": "https://nces.ed.gov/ccd/"},
            "boundaries": {"label": "NYC Open Data: Borough Boundaries & School Districts", "url": "https://data.cityofnewyork.us/"},
        },
    }
    (OUT / "schools.json").write_text(json.dumps({"meta": meta, "schools": schools}, ensure_ascii=False, separators=(",", ":")))
    (OUT / "districts.geojson").write_text(json.dumps(simplify(districts_raw, "schooldist", 0.0004), separators=(",", ":")))
    (OUT / "boroughs.geojson").write_text(json.dumps(simplify(json.loads(fetch(BOROUGHS)), "boroname", 0.0004), separators=(",", ":")))
    print(f"{meta['count']} schools, {meta['exact']} with exact coordinates")


if __name__ == "__main__":
    main()
