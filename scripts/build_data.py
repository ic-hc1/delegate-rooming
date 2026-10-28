#!/usr/bin/env python3
"""Build the dashboard data from the per-hotel "Hotel Delegates Master List" workbooks.

Only aggregate counts are written. No names, delegate codes, group codes,
flight numbers, kit boxes or remarks ever leave the workbooks.

By default the aggregates are encrypted with a password (AES-256-GCM, key derived
with PBKDF2-SHA256) and written to data/summary.enc.js. The published page asks for
that password and decrypts in the browser, so the numbers are never stored in the
repository or served in the clear.

Usage:
    python3 scripts/build_data.py <folder-with-xlsx-files> --password 'secret'
    DASHBOARD_PASSWORD='secret' python3 scripts/build_data.py <folder-with-xlsx-files>
    python3 scripts/build_data.py <folder-with-xlsx-files> --plain   # unencrypted local preview only

Each workbook must have a "Master List" sheet in the standard layout
(title row, source row, instructions row, header row, then one row per delegate).
"""
import argparse
import base64
import zipfile
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

PBKDF2_ITERATIONS = 310_000


def encrypt_payload(plaintext: bytes, password: str) -> dict:
    """AES-256-GCM with a PBKDF2-SHA256 key; matches the WebCrypto calls in index.html."""
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
        from cryptography.hazmat.primitives import hashes
    except ImportError:  # pragma: no cover
        sys.exit("cryptography is required for --password: pip install cryptography")
    salt = os.urandom(16)
    iv = os.urandom(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=PBKDF2_ITERATIONS).derive(password.encode("utf-8"))
    ct = AESGCM(key).encrypt(iv, plaintext, None)  # ciphertext || 16-byte tag, as WebCrypto expects
    b64 = lambda b: base64.b64encode(b).decode("ascii")
    return {"v": 1, "kdf": "PBKDF2-SHA256", "iter": PBKDF2_ITERATIONS, "cipher": "AES-256-GCM",
            "salt": b64(salt), "iv": b64(iv), "ct": b64(ct)}

# Workbook column names that must never appear as keys in the output. Delegate rows are emitted
# only inside the encrypted payload, under short keys, and never carry codes or kit boxes.
FORBIDDEN = {"Delegate Code", "Group Code", "Group Contact Name", "Dep. Flight No.", "Kit Box"}

STANDARD_HEADER = ["No.", "Name (Last, First)", "Group Code", "Delegate Code", "Group Contact", "Gender", "Age", "Minor",
                   "Branch", "SFTS", "Primary Language", "Check-in", "Check-out", "Arrival Date", "Arrival Time", "Airport",
                   "Airline", "Flight No.", "Gap Nights", "Departure Date", "Departure Time", "Special Rooming", "Arrived",
                   "Welcome Kit", "Kit Box", "HID Remarks", "Status"]
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


HOTEL_BY_FILE_TOKEN = {
    "CITADINES": "Citadines Bay City Manila",
    "CONRAD": "Conrad Manila",
    "GRANDWEST": "Grand Westside Manila Bay",
    "HIEX": "Holiday Inn Express Manila",
    "KINGSFORD": "Kingsford Hotel Manila",
    "LANSON": "Lanson Place Mall of Asia",
    "MANILAHOTEL": "The Manila Hotel",
    "OKADA": "Okada Manila",
    "SAVOY": "Savoy Hotel Manila",
    "SEDA": "Seda Manila Bay",
    "TRYP": "TRYP by Wyndham - Mall of Asia",
}


def file_saved_at(path):
    """When the workbook was last saved (from its document properties), as an ISO UTC string, or None."""
    try:
        core = zipfile.ZipFile(path).read("docProps/core.xml").decode("utf-8", "replace")
        m = re.search(r"<dcterms:modified[^>]*>([^<]+)<", core)
        if not m:
            return None
        ts = dt.datetime.fromisoformat(m.group(1).strip().replace("Z", "+00:00"))
        return ts.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    except Exception:
        return None


def _cell_text(v):
    return v.strip() if isinstance(v, str) else ""


def _is_delegate_code(v):
    return isinstance(v, str) and re.fullmatch(r"[A-Z0-9]{11}", v.strip()) is not None


def _is_group_code(v):
    return isinstance(v, str) and re.fullmatch(r"[A-Z0-9]{12}", v.strip()) is not None


def _looks_like_title(text):
    return "|" in text and "Hotel Delegates Master List" in text


def read_workbook(path):
    """Parse one hotel workbook. Hotels edit the top rows by hand (blank, delete or move them), so every
    part of the layout is located by content rather than by row number."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "Master List" not in wb.sheetnames:
        raise ValueError(f"{os.path.basename(path)}: no 'Master List' sheet")
    ws = wb["Master List"]
    rows = list(ws.iter_rows(values_only=True))
    fname = os.path.basename(path)
    notes = []

    # --- hotel name: the file name token is the stable source; the title row is a cross-check ---
    upper = fname.upper()
    token = next((t for t in sorted(HOTEL_BY_FILE_TOKEN, key=len, reverse=True) if t in upper), "")
    title_text = ""
    for r in rows[:10]:
        t = _cell_text(r[0]) if r else ""
        if _looks_like_title(t):
            title_text = t.split("|")[0].strip()
            break
    if not title_text:
        for other in wb.worksheets:
            if other.title == "Master List":
                continue
            first = next(other.iter_rows(min_row=1, max_row=1, values_only=True), None)
            t = _cell_text(first[0]) if first else ""
            if "|" in t and not t.lower().startswith(("source:", "hid:", "updates", "count of", "flight", "sorted", "minors")):
                title_text = t.split("|")[0].strip()
                break
    hotel = HOTEL_BY_FILE_TOKEN.get(token) or title_text
    if not hotel:
        raise ValueError(f"{fname}: cannot determine the hotel name")
    if title_text and title_text != hotel:
        notes.append(f"title row says {title_text!r}; using {hotel!r} from the file name")
    if not title_text:
        notes.append("title row missing; hotel name taken from the file name")

    # --- source / generated line ---
    source = generated = ""
    for r in rows[:10]:
        t = _cell_text(r[0]) if r else ""
        if t.lower().startswith("source:"):
            for part in t.split("|"):
                part = part.strip()
                if part.lower().startswith("source:"):
                    source = part.split(":", 1)[1].strip()
                elif part.lower().startswith("generated:"):
                    generated = part.split(":", 1)[1].strip()
            break
    if not generated:
        notes.append("generated date line missing")

    # --- header row (may be blanked in places, moved, or deleted) ---
    header_idx = next((i for i, r in enumerate(rows[:100]) if r and "Delegate Code" in r and "Name (Last, First)" in r), None)
    first_data_idx = next((i for i, r in enumerate(rows[:40]) if r and any(_is_delegate_code(c) for c in r[:6])), None)
    if first_data_idx is None:
        raise ValueError(f"{fname}: no delegate rows found")
    if header_idx is not None:
        header = [str(h).strip() if h is not None else "" for h in rows[header_idx]]
        # Restore blank header cells from the standard sequence. A blank whose expected label already
        # appears further right is an empty column someone inserted, and is left unmapped.
        filled, inserted, k = [], [], 0
        for i, h in enumerate(header):
            if h in STANDARD_HEADER:
                k = STANDARD_HEADER.index(h) + 1
            elif k < len(STANDARD_HEADER) and STANDARD_HEADER[k] not in header and (
                    not h or STANDARD_HEADER[k + 1] in header[i + 1:i + 2] if k + 1 < len(STANDARD_HEADER) else not h):
                # blank, or typed-over text sitting exactly where the next standard label belongs
                filled.append(STANDARD_HEADER[k] + ("" if not h else f" (was {h!r})"))
                header[i] = STANDARD_HEADER[k]
                k += 1
            elif not h and k < len(STANDARD_HEADER):
                inserted.append(i)
        if filled:
            notes.append("blank header cell(s) restored: " + ", ".join(filled))
        for i in inserted:
            sample = [r[i] for r in rows[header_idx + 1:header_idx + 200] if r and i < len(r) and r[i] is not None]
            notes.append(f"empty inserted column at position {i + 1} ignored" + (f" (it holds {len(sample)} value(s) in the first rows, please check)" if sample else ""))
        missing = [c for c in STANDARD_HEADER if c not in header]
        if missing:
            raise ValueError(f"{fname}: Master List is missing column(s): {', '.join(missing)}")
    else:
        width = max(len(r) for r in rows[first_data_idx:first_data_idx + 5])
        if width < len(STANDARD_HEADER):
            raise ValueError(f"{fname}: header row deleted and the sheet has only {width} columns; cannot map them safely")
        header = STANDARD_HEADER + [""] * (width - len(STANDARD_HEADER))
        notes.append("header row deleted; assuming the standard column order")

    # --- delegate rows: a row is a delegate when it carries a delegate code; the Status column marks
    # cancelled ones (grey rows at the bottom). Anything else (moved title/header rows, typed-over
    # cells, totals) is a stray and is ignored. The "No." column is not trusted: hotels type over it.
    active, cancelled, strays, renumbered, recovered = [], 0, 0, 0, 0
    cancelled_rows = []
    for r in rows[first_data_idx:]:
        if r is None or all(c is None for c in r):
            continue
        rec = dict(zip(header, r))
        if not _is_delegate_code(rec.get("Delegate Code")):
            # A delegate whose code cell was typed over still has a row number, a name and a group code.
            no = rec.get("No.")
            if (isinstance(no, (int, float)) and not isinstance(no, bool)
                    and _cell_text(rec.get("Name (Last, First)")) and _is_group_code(rec.get("Group Code"))):
                recovered += 1
            else:
                strays += 1
                continue
        status = _cell_text(rec.get("Status")).lower()
        if status.startswith("cancel"):
            cancelled += 1
            cancelled_rows.append(rec)
            continue
        no = rec.get("No.")
        if no is None and rec.get("Name (Last, First)") is None:
            strays += 1
            continue
        if not isinstance(no, (int, float)) or isinstance(no, bool):
            renumbered += 1
        active.append(rec)
    if strays:
        notes.append(f"{strays} stray non-delegate row(s) ignored")
    if renumbered:
        notes.append(f"{renumbered} delegate row(s) whose 'No.' cell was typed over were kept")
    if recovered:
        notes.append(f"{recovered} delegate row(s) whose delegate code was blanked were kept")
    for n in notes:
        print(f"  note: {fname}: {n}", file=sys.stderr)
    return {"hotel": hotel, "source": source, "generated": generated, "savedAt": file_saved_at(path),
            "rows": active, "cancelled": cancelled, "cancelledRows": cancelled_rows, "file": fname}


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


def aggregate(rows, cancelled, changes=None):
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

    changes = changes or {}
    changed = [r for r in rows if _cell_text(r.get("Delegate Code")) in changes]
    out = {
        "delegates": n,
        "flightChanges": len(changed),
        "flightChangesNotArrived": sum(1 for r in changed if not yes(r.get("Arrived"))),
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
        "sftsSpecialRooming": sum(1 for r in rows if yes(r.get("SFTS")) and isinstance(r.get("Special Rooming"), str) and r["Special Rooming"].strip()),
        "sftsSpecialRoomingArrived": sum(1 for r in rows if yes(r.get("SFTS")) and isinstance(r.get("Special Rooming"), str) and r["Special Rooming"].strip() and yes(r.get("Arrived"))),
        "sftsArrived": sum(1 for r in rows if yes(r.get("SFTS")) and yes(r.get("Arrived"))),
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


def _generated_date(text):
    """'Oct 24, 2026' -> date; unknown -> date.min so a dated file always wins."""
    for fmt in ("%b %d, %Y", "%B %d, %Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime((text or "").strip(), fmt).date()
        except ValueError:
            pass
    return dt.date.min


def decrypt_existing(path, password):
    """Return the currently published summary, or None if there is none / it cannot be read."""
    if not os.path.exists(path):
        return None
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
        from cryptography.hazmat.primitives import hashes
        m = re.search(r"window\.ROOMING_DATA_ENC = (\{.*\});", open(path, encoding="utf-8").read())
        enc = json.loads(m.group(1))
        key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=base64.b64decode(enc["salt"]),
                         iterations=enc["iter"]).derive(password.encode("utf-8"))
        return json.loads(AESGCM(key).decrypt(base64.b64decode(enc["iv"]), base64.b64decode(enc["ct"]), None))
    except Exception as exc:  # different password, old format, corrupt file
        print(f"  note: could not read the existing {os.path.basename(path)} ({exc.__class__.__name__}); skipping the freshness check", file=sys.stderr)
        return None


def freshness(summary):
    gen = max((_generated_date(g) for g in (summary.get("generated") or "").split(" / ")), default=dt.date.min)
    o = summary["overall"]
    return gen, o.get("arrived", 0) + o.get("welcomeKits", 0)


def guard_against_older(summary, enc_path, password, allow_older):
    """Refuse to overwrite newer published data with an older snapshot (unless --allow-older)."""
    current = decrypt_existing(enc_path, password)
    if not current:
        return
    new_date, new_progress = freshness(summary)
    cur_date, cur_progress = freshness(current)
    older = new_date < cur_date or (new_date == cur_date and new_progress < cur_progress)
    problems = []
    if older:
        problems.append(f"overall: new files (generated {new_date}, {new_progress} arrivals+kits) look older than the "
                        f"published data (generated {cur_date}, {cur_progress} arrivals+kits)")
    # Per hotel: a list saved earlier than the published one, or showing less progress, is a stale copy.
    published = {h["name"]: h for h in current.get("hotels", [])}
    for h in summary["hotels"]:
        ph = published.get(h["name"])
        if not ph:
            continue
        new_saved, cur_saved = h.get("savedAt") or "", ph.get("savedAt") or ""
        new_prog, cur_prog = h["arrived"] + h["welcomeKits"], ph["arrived"] + ph["welcomeKits"]
        if (new_saved and cur_saved and new_saved < cur_saved) or new_prog < cur_prog:
            problems.append(f"{h['name']}: incoming file saved {new_saved or '?'} with {new_prog} arrivals+kits, "
                            f"published one saved {cur_saved or '?'} with {cur_prog}")
    if problems and not allow_older:
        sys.exit("Refusing to publish older data:\n  - " + "\n  - ".join(problems) +
                 "\nSupply the newer file(s), or re-run with --allow-older if this is intended.")
    for pr in problems:
        print(f"  WARNING: publishing older data ({pr})", file=sys.stderr)


def read_flight_changes(path):
    """Planners' 'Change of flights' sheet: delegate ID -> new arrival details. IDs look like
    G-43P2-JQ2B-WXG and match the 11-character delegate code once the prefix and dashes go."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hi = next((i for i, r in enumerate(rows[:10]) if r and "Delegate ID" in r), None)
    if hi is None:
        raise ValueError(f"{os.path.basename(path)}: no 'Delegate ID' column found")
    hdr = [str(h).strip() if h is not None else "" for h in rows[hi]]
    col = lambda name: next((i for i, h in enumerate(hdr) if h.lower().startswith(name.lower())), None)
    ci, cd, ct, cf, cc, cn, ca = (col("Delegate ID"), col("New Arrival Date"), col("New Arrival Time"), col("New Flight"),
                                  col("Hotel Check-in"), col("Notes"), col("Activity affected"))
    out = {}
    for r in rows[hi + 1:]:
        if not r or ci is None or r[ci] is None:
            continue
        code = re.sub(r"[^A-Z0-9]", "", str(r[ci]).upper())
        if code.startswith("G") and len(code) == 12:
            code = code[1:]
        if not re.fullmatch(r"[A-Z0-9]{11}", code):
            continue
        d = as_date(r[cd]) if cd is not None else None
        t = as_time(r[ct]) if ct is not None else None
        raw_date = _cell_text(r[cd]) if cd is not None and isinstance(r[cd], str) else ""
        fl = _cell_text(r[cf]) if cf is not None and isinstance(r[cf], str) else (str(r[cf]) if cf is not None and r[cf] is not None else "")
        rec = {
            "d": iso(d), "t": t.strftime("%H:%M") if t else None,
            "fl": fl if fl.lower() != "unknown" else None,
            "unk": (d is None and (raw_date.lower() == "unknown" or not raw_date)) or fl.lower() == "unknown",
            "ck": (_cell_text(r[cc]) or None) if cc is not None else None,
            "nt": (_cell_text(r[cn]) or None) if cn is not None else None,
            "act": (_cell_text(r[ca]).upper() == "Y") if ca is not None else False,
        }
        if rec["ck"] and rec["ck"].lower() == "unknown":
            rec["ck"] = None
        prev = out.get(code)
        if prev:
            # The same delegate listed twice: a later row is normally the update, but keep any known
            # value the later row lacks, and keep every note.
            merged = {k: (rec[k] if rec[k] not in (None, False) else prev[k]) for k in ("d", "t", "fl", "ck")}
            notes = [n for n in (prev["nt"], rec["nt"]) if n]
            merged["nt"] = " | ".join(dict.fromkeys(notes)) or None
            merged["act"] = prev["act"] or rec["act"]
            merged["unk"] = merged["d"] is None or merged["fl"] is None
            rec = merged
        out[code] = rec
    return out


def read_registry(path):
    """Registration export (Delegates-YYYY-MM-DD.csv): delegate code -> {hotel text, check-in}. No names."""
    import csv as _csv
    out = {}
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for r in _csv.DictReader(fh):
            code = (r.get("Delegate Code") or "").strip()
            if _is_delegate_code(code):
                out[code] = {"acc": (r.get("Accommodations") or "").strip(), "ci": (r.get("Check-in") or "").strip() or None,
                             "status": (r.get("Status") or "").strip()}
    if not out:
        raise ValueError(f"{os.path.basename(path)}: no delegate codes found")
    return out


def hotel_key_for(text, hotels):
    """Map free hotel text ('Kingsford Hotel | IC2026SFTS-016', 'HiEx', 'Okada Manila') to a hotel key."""
    t = re.sub(r"[^a-z]", "", (text or "").split("|")[0].lower())
    if not t:
        return None
    alias = {"hiex": "holiday", "kfh": "kingsford", "gws": "grandwest", "manilahotel": "themanilahotel"}
    t = alias.get(t, t)
    for h in hotels:
        n = re.sub(r"[^a-z]", "", h["name"].lower())
        if n.startswith(t) or t.startswith(n[:8]) or (len(t) >= 5 and t in n):
            return h["key"]
    return None


def read_cancellations(path):
    """Planners' cancellation list: delegate code -> True. Any sheet with a 'Delegate Code' column."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out = {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hi = next((i for i, r in enumerate(rows[:10]) if r and "Delegate Code" in r), None)
        if hi is None:
            continue
        hdr = list(rows[hi]); ci = hdr.index("Delegate Code")
        ni = hdr.index("Name") if "Name" in hdr else None
        hi_ = hdr.index("Hotel") if "Hotel" in hdr else None
        for r in rows[hi + 1:]:
            if r and ci < len(r) and _is_delegate_code(r[ci]):
                out[str(r[ci]).strip()] = {"name": _cell_text(r[ni]) if ni is not None else "", "hotel": _cell_text(r[hi_]) if hi_ is not None else ""}
    if not out:
        raise ValueError(f"{os.path.basename(path)}: no delegate codes found")
    return out


def delegate_rows(rows, hotel_key, changes=None, cancelled_rows=()):
    """Per-delegate rows for the names layer of the dashboard (encrypted payload only)."""
    out = []
    for r in rows:
        ci, co, ad, at = as_date(r.get("Check-in")), as_date(r.get("Check-out")), as_date(r.get("Arrival Date")), as_time(r.get("Arrival Time"))
        fl = r.get("Flight No.")
        rm = r.get("HID Remarks")
        sr = r.get("Special Rooming")
        code = _cell_text(r.get("Delegate Code"))
        fc = (changes or {}).get(code)
        out.append({
            "n": _cell_text(r.get("Name (Last, First)")) or "(no name)",
            "h": hotel_key,
            "ci": iso(ci), "co": iso(co), "ad": iso(ad), "at": at.strftime("%H:%M") if at else None,
            "fl": _cell_text(fl) if isinstance(fl, str) else (str(fl) if fl is not None else None),
            "ap": _cell_text(r.get("Airport")) or None,
            "ar": yes(r.get("Arrived")), "wk": yes(r.get("Welcome Kit")),
            "rm": _cell_text(rm) if isinstance(rm, str) else (str(rm) if rm is not None else None),
            "sr": _cell_text(sr) if isinstance(sr, str) else None,
            "sf": yes(r.get("SFTS")), "mi": yes(r.get("Minor")), "gc": yes(r.get("Group Contact")),
            "br": _cell_text(r.get("Branch")) or None, "lg": _cell_text(r.get("Primary Language")) or None,
            **({"fc": fc} if fc else {}),
        })
    for r in cancelled_rows:
        out.append({"n": _cell_text(r.get("Name (Last, First)")) or "(no name)", "h": hotel_key, "ci": iso(as_date(r.get("Check-in"))),
                    "co": iso(as_date(r.get("Check-out"))), "ad": iso(as_date(r.get("Arrival Date"))), "at": None, "fl": None, "ap": None,
                    "ar": False, "wk": False, "rm": None, "sr": None, "sf": yes(r.get("SFTS")), "mi": yes(r.get("Minor")), "gc": yes(r.get("Group Contact")),
                    "br": _cell_text(r.get("Branch")) or None, "lg": None, "cx": True})
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
    ap.add_argument("--password", default=os.environ.get("DASHBOARD_PASSWORD"),
                    help="password that protects the published dashboard (or set DASHBOARD_PASSWORD)")
    ap.add_argument("--changes", default=None,
                    help="planners' 'Change of flights' workbook; marks rebooked delegates with their new arrival")
    ap.add_argument("--cancelled", default=None,
                    help="planners' cancellation workbook (a 'Delegate Code' column); those delegates are excluded from every count")
    ap.add_argument("--registry", default=None,
                    help="registration export CSV (Delegates-YYYY-MM-DD.csv) used as the registered base for reconciliation")
    ap.add_argument("--allow-older", action="store_true",
                    help="replace the published data even if the new files look older than it")
    ap.add_argument("--plain", action="store_true",
                    help="also write unencrypted summary.json / summary.js (git-ignored; local preview only)")
    args = ap.parse_args()
    if not args.password and not args.plain:
        sys.exit("A password is required: pass --password or set DASHBOARD_PASSWORD (or use --plain for a local, unprotected preview).")
    if args.password is not None and len(args.password) < 8:
        sys.exit("Use a password of at least 8 characters.")

    files = sorted(glob.glob(os.path.join(args.folder, "*.xlsx")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    if not files:
        sys.exit(f"No .xlsx files found in {args.folder}")

    hotels, all_rows, total_cancelled = [], [], 0
    generated, sources = set(), set()
    seen = {}
    kept = {}
    delegates = []
    cancelled_codes = set()
    if args.cancelled:
        cancelled_codes = read_cancellations(args.cancelled)
        print(f"  cancellations: {len(cancelled_codes)} delegate codes in {os.path.basename(args.cancelled)}", file=sys.stderr)
    cancelled_by_list_total = 0
    changes = {}
    if args.changes:
        changes = read_flight_changes(args.changes)
        print(f"  flight changes: {len(changes)} delegates in {os.path.basename(args.changes)}", file=sys.stderr)
    for f in files:
        wb = read_workbook(f)
        key = wb["hotel"].lower()
        if key in seen:
            prev = seen[key]
            progress = lambda w: sum(1 for r in w["rows"] if yes(r.get("Arrived")) or yes(r.get("Welcome Kit")))
            newer = (_generated_date(wb["generated"]), progress(wb), len(wb["rows"])) >= \
                    (_generated_date(prev["generated"]), progress(prev), len(prev["rows"]))
            print(f"  WARNING: {wb['hotel']} appears twice ({prev['file']} and {wb['file']}); keeping the "
                  f"{'newer' if newer else 'earlier'} file {wb['file'] if newer else prev['file']}", file=sys.stderr)
            if not newer:
                continue
            kept.pop(key, None)
            hotels = [h for h in hotels if h["name"].lower() != key]
            all_rows = [r for r in all_rows if r.get("_hotel") != key]
            delegates = [d for d in delegates if d["h"] != slug(prev["hotel"])]
            total_cancelled -= prev["cancelled"]
        seen[key] = wb
        listed = [r for r in wb["rows"] if _cell_text(r.get("Delegate Code")) in cancelled_codes]
        if listed:
            wb["rows"] = [r for r in wb["rows"] if _cell_text(r.get("Delegate Code")) not in cancelled_codes]
            wb["cancelled"] += len(listed)
            wb["cancelledByList"] = len(listed)
            cancelled_by_list_total += len(listed)
            print(f"  note: {wb['file']}: {len(listed)} delegate(s) excluded from the planners' cancellation list", file=sys.stderr)
        wb["listed"] = listed
        for r in wb["rows"]:
            r["_hotel"] = key
        print(f"  {wb['hotel']:<34} delegates={len(wb['rows']):>4}  cancelled={wb['cancelled']}", file=sys.stderr)
        agg = aggregate(wb["rows"], wb["cancelled"], changes)
        agg["cancelledByList"] = wb.get("cancelledByList", 0)
        hotels.append({"key": slug(wb["hotel"]), "name": wb["hotel"], "savedAt": wb["savedAt"], **agg})
        delegates.extend(delegate_rows(wb["rows"], slug(wb["hotel"]), changes))
        kept[key] = wb
        all_rows.extend(wb["rows"])
        total_cancelled += wb["cancelled"]
        if wb["generated"]:
            generated.add(wb["generated"])
        if wb["source"]:
            sources.add(wb["source"])

    hotels.sort(key=lambda h: -h["delegates"])
    overall = aggregate(all_rows, total_cancelled, changes)
    overall["cancelledByList"] = cancelled_by_list_total

    # ---- one combined cancellation record: planners' list + hotel Status column ----
    import collections
    slug_by_lc = {h["name"].lower(): h["key"] for h in hotels}
    by_key = {h["key"]: h for h in hotels}
    active_codes = {_cell_text(r.get("Delegate Code")): r.get("_hotel") for r in all_rows}
    cx = {}
    for key_lc, wbk in kept.items():
        hk = slug(wbk["hotel"])
        for r in wbk.get("cancelledRows", []):
            c = _cell_text(r.get("Delegate Code"))
            if c and c not in active_codes:
                cx[c] = {"n": _cell_text(r.get("Name (Last, First)")) or "(no name)", "h": hk, "ci": iso(as_date(r.get("Check-in"))),
                         "st": _cell_text(r.get("Status")) or None, "hotelMarked": True}
        for r in wbk.get("listed", []):
            c = _cell_text(r.get("Delegate Code"))
            cx[c] = {"n": _cell_text(r.get("Name (Last, First)")) or "(no name)", "h": hk, "ci": iso(as_date(r.get("Check-in"))),
                     "st": _cell_text(r.get("Status")) or None, "hotelMarked": False, "wasActive": True}
    for c, info in cancelled_codes.items():
        if c in cx:
            cx[c]["onList"] = True
        else:
            cx[c] = {"n": info.get("name") or "(no name)", "h": hotel_key_for(info.get("hotel"), hotels), "ci": None,
                     "st": None, "hotelMarked": False, "onList": True}
    registry = read_registry(args.registry) if args.registry else None
    for c, v in cx.items():
        v["src"] = "both" if v.get("onList") and v.get("hotelMarked") else ("list" if v.get("onList") else "hotel")
        v["reg"] = (c in registry) if registry else None
        delegates.append({"n": v["n"], "h": v["h"] or "unknown", "ci": v["ci"], "co": None, "ad": None, "at": None, "fl": None, "ap": None,
                          "ar": False, "wk": False, "rm": None, "sr": None, "sf": False, "mi": False, "gc": False, "br": None, "lg": None,
                          "cx": True, "cs": v["src"], "cst": v["st"], "creg": v["reg"]})
    for h in hotels:
        mine = [v for v in cx.values() if v["h"] == h["key"]]
        h["cancelTotal"] = len(mine)
    overall["cancelTotal"] = len(cx)
    cx_summary = {"total": len(cx), "both": sum(1 for v in cx.values() if v["src"] == "both"),
                  "listOnly": sum(1 for v in cx.values() if v["src"] == "list"), "hotelOnly": sum(1 for v in cx.values() if v["src"] == "hotel"),
                  "excludedFromHotelLists": cancelled_by_list_total}
    reg_meta = None
    if registry:
        m_ = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(args.registry))
        reg_codes = set(registry)
        reg_key = {c: hotel_key_for(v["acc"], hotels) for c, v in registry.items()}
        missing = [c for c in reg_codes if c not in active_codes and c not in cx]
        extra = [c for c in active_codes if c and c not in reg_codes]
        blanks = collections.Counter(slug_by_lc.get(hl) for c, hl in active_codes.items() if not c)
        blank_rows = collections.Counter()
        for r in all_rows:
            if not _cell_text(r.get("Delegate Code")):
                blank_rows[slug_by_lc.get(r.get("_hotel"))] += 1
        paired_total = 0
        for h in hotels:
            k = h["key"]
            h["registered"] = sum(1 for c in reg_codes if reg_key[c] == k)
            h["regCancelled"] = sum(1 for c in reg_codes if reg_key[c] == k and c in cx)
            miss_here = sum(1 for c in missing if reg_key[c] == k)
            paired = min(blank_rows[k], miss_here)
            paired_total += paired
            h["regMissing"] = miss_here - paired
            h["blankCode"] = blank_rows[k]
            h["notRegistered"] = sum(1 for c in extra if active_codes[c] == h["name"].lower()) + (blank_rows[k] - paired)
        no_hotel = [c for c in reg_codes if reg_key[c] is None]
        reg_meta = {"file": os.path.basename(args.registry), "date": m_.group(1) if m_ else None, "count": len(registry),
                    "cancelled": sum(1 for c in cx if c in reg_codes), "inHotelLists": len(reg_codes & set(active_codes)),
                    "missing": len(missing) - paired_total, "notRegistered": len(extra) + sum(blank_rows.values()) - paired_total,
                    "blankCodes": sum(blank_rows.values()), "noHotel": len(no_hotel), "noHotelCancelled": sum(1 for c in no_hotel if c in cx),
                    "cancelledBeforeExport": sum(1 for c in cx if c not in reg_codes)}
        overall["registered"] = len(registry)
        print(f"  registry: {len(registry)} registered; {reg_meta['inHotelLists']} active in hotel lists; {reg_meta['cancelled']} cancelled; "
              f"{len(missing)} missing from hotel lists; {len(extra)} in hotel lists but not registered", file=sys.stderr)
        if missing:
            print("  registered but in no hotel list and not cancelled: " + ", ".join(f"{c} ({registry[c]['acc']}, check-in {registry[c]['ci']})" for c in sorted(missing)), file=sys.stderr)
        if sum(blank_rows.values()):
            print(f"  active hotel rows with a blank delegate code: {dict(blank_rows)} (paired with {paired_total} registered-but-missing at the same hotel)", file=sys.stderr)
        if extra:
            print("  in hotel lists but not in the registration export: " + ", ".join(f"{c} ({active_codes[c]})" for c in sorted(extra)), file=sys.stderr)
    if changes:
        known = {_cell_text(r.get("Delegate Code")) for r in all_rows}
        unmatched = [c for c in changes if c not in known]
        if unmatched:
            print(f"  WARNING: {len(unmatched)} flight-change row(s) match no delegate in the hotel lists", file=sys.stderr)
    summary = {
        "title": "Delegate Rooming Monitor",
        "generated": " / ".join(sorted(generated)) or None,
        "source": " / ".join(sorted(sources)) or None,
        "builtAt": dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "hotelCount": len(hotels),
        "flightChanges": ({"file": os.path.basename(args.changes), "savedAt": file_saved_at(args.changes), "count": len(changes)} if args.changes else None),
        "cancellations": ({"file": os.path.basename(args.cancelled), "savedAt": file_saved_at(args.cancelled), "count": len(cancelled_codes), "inLists": cancelled_by_list_total} if args.cancelled else None),
        "cancelSummary": cx_summary,
        "registry": reg_meta,
        "overall": overall,
        "hotels": hotels,
        "delegates": sorted(delegates, key=lambda d: (d["h"], d["ci"] or "", d["n"].lower())),
    }
    scrub_check(summary)

    if args.password:
        guard_against_older(summary, os.path.join(args.out, "summary.enc.js"), args.password, args.allow_older)

    os.makedirs(args.out, exist_ok=True)
    compact = json.dumps(summary, separators=(",", ":"), ensure_ascii=False)
    written = []
    if args.password:
        enc = encrypt_payload(compact.encode("utf-8"), args.password)
        enc_path = os.path.join(args.out, "summary.enc.js")
        with open(enc_path, "w", encoding="utf-8") as fh:
            fh.write("// Generated by scripts/build_data.py - encrypted aggregate counts, no personal data.\n")
            fh.write("window.ROOMING_DATA_ENC = ")
            json.dump(enc, fh, separators=(",", ":"))
            fh.write(";\n")
        written.append(enc_path)
    if args.plain:
        json_path = os.path.join(args.out, "summary.json")
        js_path = os.path.join(args.out, "summary.js")
        with open(json_path, "w", encoding="utf-8") as fh:
            json.dump(summary, fh, indent=1, ensure_ascii=False)
        with open(js_path, "w", encoding="utf-8") as fh:
            fh.write("// Generated by scripts/build_data.py - aggregate counts only, no personal data. NOT for publishing.\n")
            fh.write("window.ROOMING_DATA = " + compact + ";\n")
        written += [json_path, js_path]
    print(f"Wrote {', '.join(written)}: {len(hotels)} hotels, {overall['delegates']} delegates", file=sys.stderr)


if __name__ == "__main__":
    main()
