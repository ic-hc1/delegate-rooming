#!/usr/bin/env python3
"""Build data/summary.json from the per-hotel "Hotel Delegates Master List" workbooks.

Only aggregate counts are written. No names, delegate codes, group codes,
flight numbers, kit boxes or remarks ever leave the workbooks.

Usage:
    python3 scripts/build_data.py <folder-with-xlsx-files> [--out data]

Each workbook must have a "Master List" sheet in the standard layout
(title row, source row, instructions row, header row, then one row per delegate).
"""
import argparse
import datetime as dt
import glob
import json
import os
import re
import sys
from collections import Counter, defaultdict

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required: pip install openpyxl")

# Columns that must never reach the output, in any form.
FORBIDDEN = {
    "Name (Last, First)", "Delegate Code", "Group Code", "Group Contact Name",
    "Flight No.", "Dep. Flight No.", "Kit Box", "HID Remarks",
}

BANDS = ["12AM-3AM", "3AM-6AM", "6AM-9AM", "9AM-12PM", "12PM-3PM", "3PM-6PM", "6PM-9PM", "9PM-12AM"]
TOP_N = 10


def yes(v):
    return isinstance(v, str) and v.strip().lower() in ("yes", "y", "true", "1")


def iso(d):
    return d.strftime("%Y-%m-%d") if isinstance(d, (dt.datetime, dt.date)) else None


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


def as_time(v):
    if isinstance(v, dt.time):
        return v
    if isinstance(v, dt.datetime):
        return v.time()
    return None


def slug(name):
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


def read_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "Master List" not in wb.sheetnames:
        raise ValueError(f"{os.path.basename(path)}: no 'Master List' sheet")
    ws = wb["Master List"]
    rows = list(ws.iter_rows(values_only=True))
    title = str(rows[0][0] or "")
    hotel = title.split("|")[0].strip() or os.path.basename(path)
    meta_line = str(rows[1][0] or "")
    source = ""
    generated = ""
    for part in meta_line.split("|"):
        part = part.strip()
        if part.lower().startswith("source:"):
            source = part.split(":", 1)[1].strip()
        elif part.lower().startswith("generated:"):
            generated = part.split(":", 1)[1].strip()
    header_idx = next(i for i, r in enumerate(rows) if r and r[0] == "No.")
    header = [str(h).strip() if h is not None else "" for h in rows[header_idx]]
    active, cancelled = [], 0
    for r in rows[header_idx + 1:]:
        if r is None or all(c is None for c in r):
            continue
        rec = dict(zip(header, r))
        if rec.get("No.") is None:
            # Grey rows at the bottom carry a Status of "Cancelled ..."
            if rec.get("Name (Last, First)") is not None or rec.get("Status"):
                cancelled += 1
            continue
        active.append(rec)
    return {"hotel": hotel, "source": source, "generated": generated,
            "rows": active, "cancelled": cancelled, "file": os.path.basename(path)}


def band_index(t):
    return t.hour // 3


def is_late_night(t):
    return t is not None and (t.hour >= 22 or t.hour < 5)


def top_counts(counter, n=TOP_N):
    items = counter.most_common()
    head = [[k, v] for k, v in items[:n]]
    other = sum(v for _, v in items[n:])
    if other:
        head.append(["Other", other])
    return head


def aggregate(rows, cancelled, adults_by_group_hotel=None):
    """rows: list of dict records. Returns aggregate-only dict."""
    n = len(rows)
    groups = {r.get("Group Code") for r in rows if r.get("Group Code")}
    arrived = [r for r in rows if yes(r.get("Arrived"))]
    kits = [r for r in rows if yes(r.get("Welcome Kit"))]
    arrived_no_kit = [r for r in arrived if not yes(r.get("Welcome Kit"))]
    minors = [r for r in rows if yes(r.get("Minor"))]

    # Adults per group (within this slice) for the minors check.
    adults = Counter()
    for r in rows:
        age = r.get("Age")
        if isinstance(age, (int, float)) and age >= 18 and r.get("Group Code"):
            adults[r["Group Code"]] += 1
    minors_no_adult = sum(1 for r in minors if adults.get(r.get("Group Code"), 0) == 0)

    checkins = defaultdict(lambda: {"expected": 0, "arrived": 0, "kits": 0, "minors": 0, "lateNight": 0})
    checkouts = defaultdict(lambda: {"delegates": 0, "groups": set(), "flightSameDay": 0, "before9am": 0,
                                     "flightAfter": 0, "flightBefore": 0, "noDeparture": 0, "minors": 0})
    grid = defaultdict(lambda: [0] * (len(BANDS) + 1))  # +1 = no time
    inhouse = Counter()
    gap_buckets = Counter()
    airports = Counter()
    departure = Counter()
    late_night = 0
    gap_1plus = 0
    no_arrival = 0
    first_checkin, last_checkout = None, None

    for r in rows:
        ci, co = as_date(r.get("Check-in")), as_date(r.get("Check-out"))
        ad, at = as_date(r.get("Arrival Date")), as_time(r.get("Arrival Time"))
        dd, dtm = as_date(r.get("Departure Date")), as_time(r.get("Departure Time"))
        minor = yes(r.get("Minor"))
        late = is_late_night(at)
        late_night += late
        if ci:
            c = checkins[ci]
            c["expected"] += 1
            c["arrived"] += yes(r.get("Arrived"))
            c["kits"] += yes(r.get("Welcome Kit"))
            c["minors"] += minor
            c["lateNight"] += late
            first_checkin = ci if first_checkin is None or ci < first_checkin else first_checkin
        if co:
            last_checkout = co if last_checkout is None or co > last_checkout else last_checkout
            c = checkouts[co]
            c["delegates"] += 1
            if r.get("Group Code"):
                c["groups"].add(r["Group Code"])
            c["minors"] += minor
            if dd is None:
                c["noDeparture"] += 1
                departure["noDeparture"] += 1
            elif dd == co:
                c["flightSameDay"] += 1
                departure["sameDay"] += 1
                if dtm is not None and dtm < dt.time(9, 0):
                    c["before9am"] += 1
                    departure["before9am"] += 1
            elif dd > co:
                c["flightAfter"] += 1
                departure["after"] += 1
            else:
                c["flightBefore"] += 1
                departure["before"] += 1
        if ci and co and co > ci:
            night = ci
            while night < co:
                inhouse[night] += 1
                night += dt.timedelta(days=1)
        if ad is None:
            no_arrival += 1
        else:
            grid[ad][band_index(at) if at is not None else len(BANDS)] += 1
        gap = r.get("Gap Nights")
        if isinstance(gap, (int, float)) and gap >= 1:
            gap_1plus += 1
            g = int(gap)
            gap_buckets["1" if g == 1 else "2" if g == 2 else "3-4" if g <= 4 else "5-7" if g <= 7 else "8+"] += 1
        ap = r.get("Airport")
        airports[ap.strip() if isinstance(ap, str) and ap.strip() else "Not set"] += 1

    # Arrival grid: keep every date within the event window, fold earlier ones.
    window_start = (first_checkin - dt.timedelta(days=3)) if first_checkin else None
    folded = [0] * (len(BANDS) + 1)
    grid_rows = []
    for d in sorted(grid):
        if window_start and d < window_start:
            folded = [a + b for a, b in zip(folded, grid[d])]
        else:
            grid_rows.append({"date": iso(d), "counts": grid[d], "total": sum(grid[d])})
    if sum(folded):
        grid_rows.insert(0, {"date": "earlier", "label": f"Before {window_start.strftime('%b %d')}",
                             "counts": folded, "total": sum(folded)})

    out = {
        "delegates": n,
        "cancelled": cancelled,
        "groups": len(groups),
        "groupContacts": sum(1 for r in rows if yes(r.get("Group Contact"))),
        "arrived": len(arrived),
        "welcomeKits": len(kits),
        "arrivedNoKit": len(arrived_no_kit),
        "minors": len(minors),
        "minorsNoAdult": minors_no_adult,
        "sfts": sum(1 for r in rows if yes(r.get("SFTS"))),
        "specialRooming": sum(1 for r in rows if isinstance(r.get("Special Rooming"), str) and r["Special Rooming"].strip()),
        "lateNight": late_night,
        "gap1plus": gap_1plus,
        "gapBuckets": [[k, gap_buckets[k]] for k in ["1", "2", "3-4", "5-7", "8+"] if gap_buckets[k]],
        "noArrivalDate": no_arrival,
        "airports": sorted(airports.items(), key=lambda kv: -kv[1]),
        "departure": {k: departure.get(k, 0) for k in ["sameDay", "before9am", "after", "before", "noDeparture"]},
        "firstCheckin": iso(first_checkin),
        "lastCheckout": iso(last_checkout),
        "checkins": [{"date": iso(d), **checkins[d]} for d in sorted(checkins)],
        "checkouts": [{"date": iso(d), **{k: (len(v) if isinstance(v, set) else v) for k, v in checkouts[d].items()}}
                      for d in sorted(checkouts)],
        "inhouse": [{"night": iso(d), "count": inhouse[d]} for d in sorted(inhouse)],
        "arrivalBands": BANDS + ["No time"],
        "arrivalGrid": grid_rows,
        "branches": top_counts(Counter(str(r.get("Branch")).strip() for r in rows if r.get("Branch"))),
        "languages": top_counts(Counter(str(r.get("Primary Language")).strip() for r in rows if r.get("Primary Language"))),
        "airlines": top_counts(Counter(str(r.get("Airline")).strip() for r in rows if r.get("Airline"))),
        "gender": sorted(Counter(str(r.get("Gender")).strip() for r in rows if r.get("Gender")).items()),
    }
    return out


def scrub_check(obj, path="$"):
    """Fail loudly if anything that looks like a person-level field slipped through."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k in FORBIDDEN:
                raise SystemExit(f"Refusing to write: forbidden key {k!r} at {path}")
            scrub_check(v, f"{path}.{k}")
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            scrub_check(v, f"{path}[{i}]")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("folder", help="folder containing *_Hotel_Delegates_Master_List.xlsx files")
    ap.add_argument("--out", default="data", help="output folder (default: data)")
    args = ap.parse_args()

    files = sorted(glob.glob(os.path.join(args.folder, "*.xlsx")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        sys.exit(f"No .xlsx files found in {args.folder}")

    hotels, all_rows, total_cancelled = [], [], 0
    generated, sources = set(), set()
    for f in files:
        wb = read_workbook(f)
        print(f"  {wb['hotel']:<34} delegates={len(wb['rows']):>4}  cancelled={wb['cancelled']}", file=sys.stderr)
        agg = aggregate(wb["rows"], wb["cancelled"])
        hotels.append({"key": slug(wb["hotel"]), "name": wb["hotel"], **agg})
        all_rows.extend(wb["rows"])
        total_cancelled += wb["cancelled"]
        if wb["generated"]:
            generated.add(wb["generated"])
        if wb["source"]:
            sources.add(wb["source"])

    hotels.sort(key=lambda h: -h["delegates"])
    overall = aggregate(all_rows, total_cancelled)
    summary = {
        "title": "Delegate Rooming Monitor",
        "generated": " / ".join(sorted(generated)) or None,
        "source": " / ".join(sorted(sources)) or None,
        "builtAt": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "hotelCount": len(hotels),
        "overall": overall,
        "hotels": hotels,
    }
    scrub_check(summary)

    os.makedirs(args.out, exist_ok=True)
    json_path = os.path.join(args.out, "summary.json")
    js_path = os.path.join(args.out, "summary.js")
    with open(json_path, "w", encoding="utf-8") as fh:
        json.dump(summary, fh, indent=1, ensure_ascii=False)
    with open(js_path, "w", encoding="utf-8") as fh:
        fh.write("// Generated by scripts/build_data.py - aggregate counts only, no personal data.\n")
        fh.write("window.ROOMING_DATA = ")
        json.dump(summary, fh, separators=(",", ":"), ensure_ascii=False)
        fh.write(";\n")
    print(f"Wrote {json_path} and {js_path}: {len(hotels)} hotels, {overall['delegates']} delegates", file=sys.stderr)


if __name__ == "__main__":
    main()
