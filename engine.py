"""Trimble robotic-laser profile data -> Enterprise EPROD laser workbooks.

Flow for one profile:
  parse_trimble_csv()  -> Shot list (pipe shots) + control shots (BM / BS)
  review_names()       -> suggested names + flags for the technician to confirm
  build_profile()      -> stations with measurement 1/2/3, averages, consistency
  fill_consistency()   -> Consistency Check workbook bytes
  fill_movement()      -> Pipe Movement Calculation workbook bytes

Naming convention (from the field instructions):
  "<ft>' U/S GW <weld> <counter>"  e.g. "3' U/S GW 24500 0001"
  "U/S GW <weld> <counter>" / "D/S GW <weld> <counter>" for the weld station.
  The counter's first digit is the profile (0001 = Profile 0, 1001 = Profile 1).
  Shots are taken in the order the tech walks the pipe; that order is kept.
"""
from __future__ import annotations

import csv
import io
import math
import re
from collections import Counter
from dataclasses import dataclass, field
from statistics import median
from typing import Optional

import openpyxl

from xlsx_patch import XlsxPatcher, column_letters

# Bump whenever parsing or the Workbook/ParsedFile shape changes: the app passes it into
# its cached loaders so Streamlit never serves objects built by older code.
ENGINE_VERSION = 11

CONSISTENCY_SHEET = "Robotic Raw Data Input"
MOVEMENT_SHEET = "Robotic Raw Data Average Input"
FIRST_ROW, LAST_ROW = 8, 107          # 100 stations; rows must never be inserted/deleted
MAX_STATIONS = LAST_ROW - FIRST_ROW + 1
LIMIT_IN = 0.20                       # Enterprise consistency limit between measurements
COVERED_NOTE = "Covered back up before Profile {p} measurement was taken"
ALLOWED_UNITS = ("USSURVEYFEET",)

NAME_RE = re.compile(r"^(?P<station>.*?)\s+(?P<counter>\d{4})$")
STATION_RE = re.compile(
    r"^(?:(?P<dist>\d+(?:\.\d+)?)\s*'+\s*)?(?:(?P<dir>[UD])\s*/?\s*S\s+(?:OF\s+)?)?GW\s*#?\s*(?P<gw>\d+)$",
    re.I)
CONTROL_RE = re.compile(r"\b(BM|BS|BENCH\s*MARK|BACK\s*SIGHT|CHECK)\d*\b|\bDO NOT USE\b", re.I)


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------
@dataclass
class Shot:
    index: int                 # order in the file (= order shot)
    raw_name: str
    n: float
    e: float
    z: float
    station: Optional[str] = None   # canonical station ID parsed from the name
    counter: Optional[int] = None
    gw: Optional[str] = None
    dist: Optional[float] = None
    side: Optional[str] = None      # "U" / "D"
    guess_counter: Optional[int] = None
    cluster: Optional[int] = None
    suggestion: Optional[str] = None
    flags: list = field(default_factory=list)

    @property
    def xyz(self):
        return (self.n, self.e, self.z)

    @property
    def profile(self) -> Optional[int]:
        c = self.counter if self.counter is not None else self.guess_counter
        return None if c is None else c // 1000


@dataclass
class ParsedFile:
    shots: list
    control: list
    units: Optional[str]
    errors: list
    warnings: list


def canonical_station(text: str):
    """'1'' us gw 5000' -> ("1' U/S GW 5000", "5000", 1.0, "U"); None if not standard."""
    cleaned = re.sub(r"\s+", " ", text.strip())
    m = STATION_RE.match(cleaned)
    if not m:
        return None
    side = m.group("dir").upper() if m.group("dir") else None
    gw = str(int(m.group("gw")))
    dist = float(m.group("dist")) if m.group("dist") else None
    if dist and side is None:
        return None                              # "3' GW 500": footage without a side
    label = f"{side}/S GW {gw}" if side else f"GW {gw}"
    if dist:
        d = int(dist) if dist.is_integer() else dist
        label = f"{d}' {label}"
    return label, gw, dist, side


def parse_name(name: str):
    """(station, counter, gw, dist, side) with None where the name doesn't parse."""
    m = NAME_RE.match(name.strip())
    station_text, counter = (m.group("station"), int(m.group("counter"))) if m else (name, None)
    found = canonical_station(station_text)
    if found:
        return found[0], counter, found[1], found[2], found[3]
    return None, counter, None, None, None


def _float(text):
    try:
        return float(str(text).strip())
    except (TypeError, ValueError):
        return None


def parse_trimble_csv(data: bytes) -> ParsedFile:
    """Read a Trimble Access 'CSV with attributes' export (Name, N, E, Elev[, code...]).

    Also accepts the hand-sorted layout (Name, Station ID, Measurement ID, N, E, Elev).
    """
    text = data.decode("utf-8-sig", errors="replace")
    errors, warnings, shots, control, null_points = [], [], [], [], []
    units = None
    for row in csv.reader(io.StringIO(text)):
        row = [c.strip() for c in row]
        if not row or not any(row):
            continue
        joined = ",".join(row)
        if joined.lower().startswith("job:") or "units:" in joined.lower():
            m = re.search(r"Units:\s*([A-Za-z]+)", joined)
            units = m.group(1) if m else units
            continue
        name = row[0]
        coords = [_float(c) for c in row[1:4]]
        if None in coords and len(row) >= 6:
            coords = [_float(c) for c in row[3:6]]
        if any(str(c).strip() == "?" for c in row[1:4]):
            null_points.append(name)
            continue
        if None in coords or not name:
            warnings.append(f"Skipped a line that isn't a point: {joined[:60]}")
            continue
        n, e, z = coords
        if CONTROL_RE.search(name) or (n == 0 and e == 0 and z == 0):
            control.append((name, n, e, z))
            continue
        station, counter, gw, dist, side = parse_name(name)
        shots.append(Shot(len(shots), name, n, e, z, station, counter, gw, dist, side))
    if units and units.upper() not in ALLOWED_UNITS:
        errors.append(f"Job units are '{units}'. Jobs must be US Survey Feet.")
    if units is None:
        warnings.append("No 'Units:' header line found - assuming US Survey Feet.")
    if null_points:
        errors.append(
            f"{len(null_points)} point(s) have no coordinates ('?' in the export), starting with "
            f"'{null_points[0]}'. Trimble couldn't compute them: the station setup has no "
            "backsight azimuth, so nothing measured from it has a position. In Trimble Access, "
            "open Review job, find that station setup's backsight azimuth record, enter the "
            "azimuth (0), then export again - the angles and distances are still stored and "
            "every point recomputes.")
    if not shots and not null_points:
        errors.append("No pipe shots found in the CSV.")
    return ParsedFile(shots, control, units, errors, warnings)


def control_check(control: list) -> list:
    """Compare repeated benchmark / backsight shots from different setups."""
    notes, by_name = [], {}
    for name, n, e, z in control:
        key = re.sub(r"\s+", " ", name.upper())
        by_name.setdefault(key, []).append((n, e, z))
    for key, pts in by_name.items():
        if len(pts) > 1:
            spread = max(math.dist(pts[0], p) for p in pts) * 12
            if spread > 0.10:
                notes.append(f"'{key}' was stored {len(pts)} times and the copies differ by up to "
                             f"{spread:.2f}\" - the setups may not share the same coordinate frame.")
    return notes


# ---------------------------------------------------------------------------
# Name review: find mislabeled shots and suggest the fix
# ---------------------------------------------------------------------------
def _guess_counters(shots: list) -> None:
    """Infer a shot number for shots whose name has none, from the nearest numbered shots."""
    valid = [s for s in shots if s.counter is not None]
    for s in shots:
        if s.counter is not None:
            s.guess_counter = s.counter
            continue
        before = [v for v in valid if v.index < s.index]
        after = [v for v in valid if v.index > s.index]
        g_before = before[-1].counter + (s.index - before[-1].index) if before else None
        g_after = after[0].counter - (after[0].index - s.index) if after else None
        if g_before is not None and g_after is not None and g_before // 1000 != g_after // 1000:
            s.guess_counter = g_after
        else:
            s.guess_counter = g_after if g_after is not None and g_after > 0 else g_before


def _cluster(shots: list) -> None:
    """Group shots of one profile that sit on the same spot of pipe."""
    steps = [math.dist(a.xyz, b.xyz) for a, b in zip(shots, shots[1:])]
    steps = [d for d in steps if 0.05 < d < 5]
    spacing = median(steps) if steps else 1.0
    tol = min(0.35, 0.35 * spacing)
    centers = []   # [sum_n, sum_e, sum_z, count]
    for s in shots:
        best, best_d = None, tol
        for i, c in enumerate(centers):
            d = math.dist(s.xyz, (c[0] / c[3], c[1] / c[3], c[2] / c[3]))
            if d <= best_d:
                best, best_d = i, d
        if best is None:
            centers.append([s.n, s.e, s.z, 1])
            s.cluster = len(centers) - 1
        else:
            c = centers[best]
            c[0] += s.n; c[1] += s.e; c[2] += s.z; c[3] += 1
            s.cluster = best


def review_names(shots: list) -> list:
    """Fill each shot's suggestion and flags. Returns the shots (same objects)."""
    _guess_counters(shots)
    gw_counts = Counter(s.gw for s in shots if s.gw)
    top_gw, top_count = (gw_counts.most_common(1)[0] if gw_counts else (None, 0))

    for profile in sorted({s.profile for s in shots if s.profile is not None}):
        group = [s for s in shots if s.profile == profile]
        _cluster(group)
        clusters = {}
        for s in group:
            clusters.setdefault(s.cluster, []).append(s)
        # Name each spot from the labels of the shots taken there. Clear majorities first;
        # ties (one mislabeled shot at a 2-shot spot) are settled by geometry: the label
        # whose footage matches the measured distance from the weld station wins, and a
        # name another spot already owns loses.
        def centre(cid):
            pts = [m.xyz for m in clusters[cid]]
            return tuple(sum(c) / len(pts) for c in zip(*pts))

        votes = {}
        for cid, members in clusters.items():
            votes[cid] = Counter(m.station for m in members if m.station and
                                 not (top_gw and m.gw != top_gw and gw_counts[m.gw] <= 2
                                      and top_count >= 5))
        names, ties = {}, []
        for cid, v in votes.items():
            ranked = v.most_common()
            if not ranked:
                names[cid] = None
            elif len(ranked) == 1 or ranked[0][1] > ranked[1][1]:
                names[cid] = ranked[0][0]
            else:
                ties.append(cid)

        weld_spot = {}
        for cid, name in names.items():
            if name and "'" not in name:
                weld_spot[name.split()[-1]] = centre(cid)

        def misfit(name, cid):
            found = canonical_station(name)
            if not found or found[1] not in weld_spot:
                return 0.0
            return abs(math.dist(centre(cid), weld_spot[found[1]]) - (found[2] or 0.0))

        owned = Counter(n for n in names.values() if n)
        for cid in ties:
            top = votes[cid].most_common(1)[0][1]
            options = [n for n, c in votes[cid].items() if c == top]
            options.sort(key=lambda n: (owned[n] > 0, misfit(n, cid),
                                        min(m.counter or 9999 for m in clusters[cid]
                                            if m.station == n)))
            names[cid] = options[0]
            owned[options[0]] += 1
            if "'" not in options[0]:
                weld_spot.setdefault(options[0].split()[-1], centre(cid))

        for s in group:
            name = names[s.cluster]
            if s.counter is None:
                s.flags.append("Name doesn't end in a 4-digit shot number")
            if s.station is None:
                s.flags.append("Station name isn't in the standard form (3' U/S GW 24500)")
            elif top_gw and s.gw != top_gw and gw_counts[s.gw] <= 2 and top_count >= 5:
                s.flags.append(f"GW {s.gw} is only used here - typo for GW {top_gw}?")
            if name and s.station and s.station != name:
                s.flags.append(f"Other shot(s) at this exact spot are named '{name}'")
            if name:
                found = canonical_station(name)
                if found and found[2] and found[1] in weld_spot:
                    measured = math.dist(s.xyz, weld_spot[found[1]])
                    if abs(measured - found[2]) > 0.6:
                        s.flags.append(f"Labeled {found[2]:g}' from the weld but measures "
                                       f"{measured:.1f}' from the weld station")
            if name is None:
                s.flags.append("Couldn't work out the station - type the correct name")
            counter = s.counter if s.counter is not None else s.guess_counter
            if name and counter is not None:
                s.suggestion = f"{name} {counter:04d}"
            else:
                s.suggestion = s.raw_name

    counters = Counter(s.counter for s in shots if s.counter is not None)
    for s in shots:
        if s.counter is not None and counters[s.counter] > 1:
            s.flags.append(f"Shot number {s.counter:04d} is used more than once")
        if s.profile is None:
            s.flags.append("Couldn't tell which profile this shot belongs to")
    return shots


# ---------------------------------------------------------------------------
# Build one profile from the confirmed names
# ---------------------------------------------------------------------------
@dataclass
class Station:
    station_id: str
    shots: list                      # Shot objects, in shot order (measurement 1, 2, 3)
    notes: str = ""
    thickness: Optional[float] = None

    def measurement_id(self, i: int) -> str:
        return f"{self.shots[i].counter:04d}"

    west_positive: bool = False      # write Y as -Easting (Enterprise: +Y = west)

    def xyz(self, i: int):
        n, e, z = self.shots[i].xyz
        return (n, -e if self.west_positive else e, z)

    @property
    def average(self):
        pts = [self.xyz(i) for i in range(min(3, len(self.shots)))]
        return tuple(sum(c) / len(pts) for c in zip(*pts))

    def differences(self):
        """[(label, dx, dy, dz, total) in inches] of measurement 2 and 3 vs measurement 1."""
        out, first = [], self.shots[0].xyz
        for i, s in enumerate(self.shots[1:3], start=2):
            d = [(b - a) * 12 for a, b in zip(first, s.xyz)]
            out.append((f"#{i} vs #1", *d, math.sqrt(sum(x * x for x in d))))
        return out

    @property
    def max_difference(self) -> Optional[float]:
        diffs = self.differences()
        return max(d[4] for d in diffs) if diffs else None


@dataclass
class ProfileBuild:
    profile: int
    stations: list
    errors: list
    warnings: list

    @property
    def failures(self):
        return [s for s in self.stations
                if s.max_difference is not None and s.max_difference > LIMIT_IN]

    @property
    def exposed_length(self) -> Optional[float]:
        if len(self.stations) < 2:
            return None
        return round(math.dist(self.stations[0].average, self.stations[-1].average), 2)


def build_profile(shots: list, final_names: dict, profile: int,
                  order: Optional[list] = None, keep_height_outliers: bool = False) -> ProfileBuild:
    """Group the profile's shots by their confirmed names (shot index -> name)."""
    errors, warnings, groups = [], [], {}
    for s in shots:
        name = final_names.get(s.index, s.raw_name)
        if name is None:                     # left out on purpose
            continue
        station, counter, *_ = parse_name(name)
        if station is None or counter is None:
            if s.profile == profile:
                errors.append(f"'{name}' still isn't a valid shot name (e.g. 3' U/S GW 24500 0001).")
            continue
        if counter // 1000 != profile:
            continue
        s.counter = counter
        groups.setdefault(station, []).append(s)
    # Order by shot number, not by position in the file: a station skipped and shot later
    # (with its reserved number) must still land in its place along the pipe.
    for members in groups.values():
        members.sort(key=lambda x: x.counter)
    groups = dict(sorted(groups.items(), key=lambda kv: kv[1][0].counter))
    if len(groups) >= 3 and not order:
        # Pipe order by position: project each station onto the line between the two
        # stations farthest apart, then start from the end nearest the first shot taken.
        cen = {k: tuple(sum(c) / len(v) for c in zip(*[x.xyz for x in v])) for k, v in groups.items()}
        ks = list(cen)
        a, b = max(((i, j) for i in ks for j in ks),
                   key=lambda p: math.dist(cen[p[0]][:2], cen[p[1]][:2]))
        ax = [q - p for p, q in zip(cen[a][:2], cen[b][:2])]
        pos = {k: sum((c - o) * u for c, o, u in zip(cen[k][:2], cen[a][:2], ax)) for k in ks}
        first = next(iter(groups))           # station holding the lowest shot number
        lo, hi = min(pos.values()), max(pos.values())
        reverse = abs(pos[first] - hi) < abs(pos[first] - lo)
        groups = dict(sorted(groups.items(), key=lambda kv: pos[kv[0]], reverse=reverse))
    if order:                                # position mode: stations in pipe order
        rank = {name: i for i, name in enumerate(order)}
        groups = dict(sorted(groups.items(), key=lambda kv: rank.get(kv[0], len(rank))))
    stations = [Station(k, v) for k, v in groups.items()]
    # A shot whose elevation is off from the other shot(s) at its station by more than
    # 0.25 ft is a target-height mistake (5.000 ft backsight rod left on, True height 0
    # instead of Bottom mount = 0.44 ft), not pipe movement. Keep the shot that lines up
    # with the neighbouring stations and leave the other out.
    def mid_z(st):
        zs = sorted(x.z for x in st.shots)
        return zs[len(zs) // 2]
    for i, st in enumerate(stations):
        if len(st.shots) < 2 or max(x.z for x in st.shots) - min(x.z for x in st.shots) <= 0.25:
            continue
        if keep_height_outliers:
            warnings.append(f"{st.station_id}: shots differ in elevation by "
                            f"{max(x.z for x in st.shots) - min(x.z for x in st.shots):.2f} ft "
                            "(kept, as chosen) - the average there is pulled off the pipe line.")
            continue
        near = [mid_z(stations[j]) for j in (i - 1, i + 1) if 0 <= j < len(stations)]
        expected = sum(near) / len(near) if near else mid_z(st)
        bad = max(st.shots, key=lambda x: abs(x.z - expected))
        st.shots.remove(bad)
        warnings.append(f"{st.station_id}: shot {bad.counter:04d} left out - its elevation is "
                        f"{bad.z - expected:+.2f} ft off the pipe line (wrong target height?).")
    for st in stations:
        if len(st.shots) == 1:
            warnings.append(f"{st.station_id}: only one measurement - Enterprise requires two.")
        if len(st.shots) > 3:
            errors.append(f"{st.station_id}: {len(st.shots)} shots under one name (max 3). "
                          "Check for a station label that wasn't changed.")
        spread = max(math.dist(a.xyz, b.xyz) for a in st.shots for b in st.shots)
        if spread > 0.5:
            errors.append(f"{st.station_id}: shots under this name are {spread:.2f} ft apart - "
                          "one of them is mislabeled.")
    if len(stations) >= 3:
        first, last = stations[0].shots[0].xyz, stations[-1].shots[0].xyz
        axis = [b - a for a, b in zip(first, last)]
        along = [sum((p - a) * u for p, a, u in zip(st.shots[0].xyz, first, axis))
                 for st in stations]
        backwards = [stations[i + 1].station_id for i in range(len(along) - 1)
                     if along[i + 1] < along[i] - 0.1 * abs(along[-1] or 1) / len(along)]
        if backwards:
            warnings.append("Stations out of order along the pipe (by shot number): "
                            + ", ".join(backwards[:5]) + ". Check those shot numbers.")
    if len(stations) > MAX_STATIONS:
        errors.append(f"{len(stations)} stations - the template holds {MAX_STATIONS}. "
                      "Use wider spacing (Enterprise allows up to 3 ft on straight pipe over 18 ft).")
    if not stations:
        errors.append(f"No shots for Profile {profile} in this CSV.")
    return ProfileBuild(profile, stations, errors, warnings)


# ---------------------------------------------------------------------------
# Workbooks
# ---------------------------------------------------------------------------
HEADER_CELLS = {"LID": "B2", "AID": "E2", "AID Name": "H2", "Dig No.": "B3",
                "Pipe OD (in)": "B4", "Pipe WT (in)": "E4", "SMYS (psi)": "H4",
                "Exposed Length (ft)": "L4", "Profile No.": "P1",
                "Technician Name": "L3"}


@dataclass
class Workbook:
    kind: str            # "consistency" or "movement"
    data: bytes
    header: dict
    rows: list           # values of A..Y for rows 8..107
    errors: list
    warnings: list


def read_workbook(data: bytes) -> Workbook:
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
    if CONSISTENCY_SHEET in wb.sheetnames:
        kind, sheet = "consistency", CONSISTENCY_SHEET
    elif MOVEMENT_SHEET in wb.sheetnames:
        kind, sheet = "movement", MOVEMENT_SHEET
    else:
        raise ValueError("This isn't an EPROD Robotic Laser Consistency Check or Pipe Movement "
                         "workbook (no 'Robotic Raw Data' input sheet).")
    ws = wb[sheet]
    grid = [list(r) for r in ws.iter_rows(min_row=1, max_row=LAST_ROW, max_col=25, values_only=True)]
    grid += [[None] * 25] * (LAST_ROW - len(grid))
    wb.close()

    def cell(ref):
        col = ord(ref[0]) - 64
        return grid[int(ref[1:]) - 1][col - 1]

    header = {k: cell(v) for k, v in HEADER_CELLS.items()}
    errors, warnings = [], []
    count_formula = str(cell("F5") or "")
    if f"A{LAST_ROW}" not in count_formula:
        errors.append(f"Rows have been deleted from '{sheet}' (station count reads "
                      f"{count_formula or 'nothing'}). Start from a fresh template.")
    broken = XlsxPatcher(data).sheet_xml_containing("#REF!")
    if broken:
        errors.append("Broken #REF! formulas in: " + ", ".join(broken) +
                      ". Rows were deleted at some point - start from a fresh template.")
    for label in ("Pipe WT (in)", "SMYS (psi)", "Pipe OD (in)"):
        if not isinstance(header[label], (int, float)):
            warnings.append(f"{label} is '{header[label]}' - %SMYS will show #VALUE! until it's a number.")
    return Workbook(kind, data, header, grid[FIRST_ROW - 1:LAST_ROW], errors, warnings)


def profile_number(book: Workbook) -> Optional[int]:
    try:
        return int(str(book.header["Profile No."]).strip())
    except (TypeError, ValueError):
        return None


def header_mismatches(a: Workbook, b: Workbook) -> list:
    out = []
    for key in ("LID", "AID", "Dig No."):
        va, vb = str(a.header[key]).strip().lstrip("0"), str(b.header[key]).strip().lstrip("0")
        if va != vb:
            out.append(f"{key} differs: consistency check '{a.header[key]}', "
                       f"pipe movement '{b.header[key]}'.")
    return out


def existing_notes(book: Optional[Workbook], profile: int) -> dict:
    """{station: (notes, thickness)} already in a workbook, to prefill the station table."""
    if book is None:
        return {}
    c0 = 2 if book.kind == "consistency" else 2 + 6 * profile   # B.. block start (1-based)
    out = {}
    for row in book.rows:
        station = row[0]
        if station:
            notes, thick = row[c0], row[c0 + 1]
            if isinstance(notes, str) and notes.startswith("="):
                notes = None
            if isinstance(thick, str) and thick.startswith("="):
                thick = None
            out[str(station).strip()] = (notes if notes not in (None, 0) else "",
                                         thick if isinstance(thick, (int, float)) else None)
    return out


def movement_stations(book: Optional[Workbook]) -> list:
    if book is None:
        return []
    return [str(r[0]).strip() for r in book.rows if r[0] not in (None, "")]


def order_stations(build: ProfileBuild, movement: Optional[Workbook]) -> tuple:
    """Rows for this profile: Pipe Movement's station order when it has one, else shot order.

    Returns (ordered stations for the consistency check, stations missing from this
    profile, stations not in the pipe movement sheet).
    """
    existing = movement_stations(movement)
    by_id = {s.station_id: s for s in build.stations}
    if not existing or build.profile == 0:
        return build.stations, [], []
    ordered = [by_id[k] for k in existing if k in by_id]
    missing = [k for k in existing if k not in by_id]
    extra = [s for s in build.stations if s.station_id not in existing]
    return ordered + extra, missing, extra


def fill_consistency(book: Workbook, stations: list, exposed_length: Optional[float],
                     technician: Optional[str] = None) -> bytes:
    p = XlsxPatcher(book.data)
    values, cached = {}, {}
    for i in range(MAX_STATIONS):
        r = FIRST_ROW + i
        st = stations[i] if i < len(stations) else None
        cols = dict.fromkeys(["A", "B", "C", "D", "E", "F", "G", "H", "K", "L", "M",
                              "N", "Q", "R", "S"])
        cache = dict.fromkeys(["I", "J", "O", "P", "T", "U", "V", "W", "X", "Y"], "")
        if st:
            cols["A"] = st.station_id
            cols["C"] = st.notes or None
            cols["D"] = st.thickness
            for m, (id_col, x_col) in enumerate((("B", "E"), ("H", "K"), ("N", "Q"))):
                if m < len(st.shots):
                    cols[id_col] = st.measurement_id(m)
                    xi = ord(x_col)
                    cols[x_col], cols[chr(xi + 1)], cols[chr(xi + 2)] = st.xyz(m)
            avg = st.average
            cache.update(I=st.notes or 0, J=st.thickness or 0, O=st.notes or 0,
                         P=st.thickness or 0, T=f"{st.measurement_id(0)}_Ave",
                         U=st.notes or 0, V=st.thickness or 0, W=avg[0], X=avg[1], Y=avg[2])
        values.update({f"{c}{r}": v for c, v in cols.items()})
        cached.update({f"{c}{r}": v for c, v in cache.items()})
    if exposed_length is not None and book.header["Exposed Length (ft)"] in (None, "", "xx"):
        values["L4"] = exposed_length
    if technician:
        values["L3"] = technician
    missing = p.set_values(CONSISTENCY_SHEET, values)
    if missing:
        raise ValueError("Could not place cells: " + ", ".join(missing[:10]))
    p.set_cached(CONSISTENCY_SHEET, cached)
    p.force_full_recalc()
    return p.to_bytes()


def fill_movement(book: Workbook, profile: int, stations: list, missing: list,
                  exposed_length: Optional[float], technician: Optional[str] = None) -> bytes:
    """Write this profile's averages into its block, rows aligned by station ID."""
    p = XlsxPatcher(book.data)
    start = 2 + 6 * profile                          # B for Profile 0, H for 1, N for 2, T for 3
    cols = [column_letters(start + k) for k in range(6)]
    existing = movement_stations(book)
    write_ids = profile == 0 or not existing
    order = [s.station_id for s in stations] if write_ids else existing
    by_id = {s.station_id: s for s in stations}
    values = {}
    for i in range(MAX_STATIONS):
        r = FIRST_ROW + i
        key = order[i] if i < len(order) else None
        if write_ids:
            values[f"A{r}"] = key
        st = by_id.get(key) if key else None
        block = [None] * 6
        if st:
            avg = st.average
            block = [f"{st.measurement_id(0)}_Ave", st.notes or None, st.thickness, *avg]
        elif key in missing:
            block[1] = COVERED_NOTE.format(p=profile)
        values.update({f"{c}{r}": v for c, v in zip(cols, block)})
    if exposed_length is not None and book.header["Exposed Length (ft)"] in (None, "", "xx"):
        values["L4"] = exposed_length
    if technician:
        values["L3"] = technician
    missing_cells = p.set_values(MOVEMENT_SHEET, values)
    if missing_cells:
        raise ValueError("Could not place cells: " + ", ".join(missing_cells[:10]))
    p.force_full_recalc()
    return p.to_bytes()


def movement_has_profile(book: Workbook, profile: int) -> bool:
    c = 2 + 6 * profile + 2                          # X column of the block (0-based index)
    return any(isinstance(r[c], (int, float)) for r in book.rows)


# ---------------------------------------------------------------------------
# On-site preview of the workbook's own math (the workbook stays the record)
# ---------------------------------------------------------------------------
def preview(points: list, thickness: list, od_in: Optional[float]) -> dict:
    """Max deflection (in) and fitted curvature strain, the way the hidden Profile sheet
    does it: project onto the first->last chord, fit 4th-order polynomials to the XY and
    Z deflections against projected length, strain = (OD/2) / bending radius.
    """
    import numpy as np
    if len(points) < 6:
        return {}
    pts = np.array(points, float)
    pts[:, 2] -= np.array([t or 0 for t in thickness], float) / 12
    d = pts - pts[0]
    chord = pts[-1] - pts[0]
    u = chord / np.linalg.norm(chord)
    proj = d @ u
    defl = d - np.outer(proj, u)
    sign = np.where(d[:, 0] * u[1] - d[:, 1] * u[0] < 0, -1, 1)
    xy, zz, v = sign * np.hypot(defl[:, 0], defl[:, 1]) * 12, defl[:, 2] * 12, proj * 12
    out = {"max_deflection_in": float(np.max(np.linalg.norm(defl, axis=1)) * 12)}
    if od_in:
        def strain(y):
            c = np.polyfit(v, y, 4)
            f1, f2 = np.polyval(np.polyder(c), v), np.polyval(np.polyder(c, 2), v)
            with np.errstate(divide="ignore"):
                return (od_in / 2) * -f2 / (1 + f1 ** 2) ** 1.5
        out["max_fitted_strain"] = float(np.max(np.sqrt(strain(xy) ** 2 + strain(zz) ** 2)))
    return out


# ---------------------------------------------------------------------------
# Several exports at once, and the Y direction
# ---------------------------------------------------------------------------
def merge_parsed(files: list) -> ParsedFile:
    """Combine several CSV exports. Jobs are cumulative, so the same shot can be in more
    than one export; exact repeats (same name and coordinates) are kept once."""
    shots, control, errors, warnings, seen = [], [], [], [], set()
    units = None
    for f in files:
        errors += f.errors
        warnings += f.warnings
        units = units or f.units
        control += f.control
        for s in f.shots:
            key = (s.raw_name.strip(), round(s.n, 4), round(s.e, 4), round(s.z, 4))
            if key in seen:
                continue
            seen.add(key)
            s.index = len(shots)
            s.flags, s.suggestion, s.cluster, s.guess_counter = [], None, None, None
            shots.append(s)
    if shots:
        errors = [e for e in errors if not e.startswith("No pipe shots")]
    return ParsedFile(shots, control, units, errors, warnings)


def detect_west_positive(book: Optional[Workbook], stations: list) -> Optional[bool]:
    """Which Y convention an uploaded Pipe Movement file already uses, judged from the
    Profile 0 block against these shots' Eastings (the pipe moves inches; the sign
    difference is tens of feet). None when there's nothing to compare."""
    if book is None:
        return None
    by_id = {s.station_id: s for s in stations}
    same = flipped = 0
    for row in book.rows:
        st = by_id.get(str(row[0]).strip()) if row[0] else None
        y = row[5]                                     # F = Profile 0 Y
        if st is None or not isinstance(y, (int, float)):
            continue
        e = sum(x.e for x in st.shots) / len(st.shots)
        if abs(e) < 1:
            continue
        same += abs(y - e) < abs(y + e)
        flipped += abs(y + e) < abs(y - e)
    if same == flipped == 0:
        return None
    return flipped > same



# ---------------------------------------------------------------------------
# Naming by measured position (when the field labels are off in a pattern)
# ---------------------------------------------------------------------------
@dataclass
class Spot:
    along: float          # ft from the GW spot, + downstream
    shots: list
    name: str = ""
    used: list = field(default_factory=list)
    dropped: list = field(default_factory=list)   # (shot, reason)


def position_layout(shots: list, profile: int, shift: float = 0.0):
    """Group a profile's shots by where they were taken and name each spot by its measured
    distance from the weld. The weld is the spot most shots labeled 'GW' (no footage) sit
    on, moved ``shift`` ft downstream if the GW shots weren't on the weld. Downstream is the
    side most 'D/S' labels are on. Returns (spots in pipe order, weld GW number, notes)."""
    group = [s for s in shots if s.profile == profile]
    if not group:
        return [], None, []
    # Group by horizontal position only, so a shot with the wrong target height still
    # lands with the other shots at its spot (and is then dropped for its elevation).
    spots, plan = {}, []
    for s in sorted(group, key=lambda x: x.counter or 9999):
        for k, (n, e, members) in enumerate(plan):
            if math.hypot(s.n - n, s.e - e) <= 0.35:
                members.append(s)
                plan[k] = (sum(x.n for x in members) / len(members),
                           sum(x.e for x in members) / len(members), members)
                break
        else:
            plan.append((s.n, s.e, [s]))
    spots = {k: members for k, (_, _, members) in enumerate(plan)}
    centres = {k: tuple(sum(c) / len(v) for c in zip(*[x.xyz for x in v])) for k, v in spots.items()}
    keys = list(centres)
    a, b = max(((i, j) for i in keys for j in keys),
               key=lambda p: math.dist(centres[p[0]][:2], centres[p[1]][:2]))
    axis = [q - p for p, q in zip(centres[a][:2], centres[b][:2])]
    length = math.hypot(*axis) or 1.0
    axis = [x / length for x in axis]

    def along(k):
        return sum((c - o) * u for c, o, u in zip(centres[k][:2], centres[a][:2], axis))

    gw_votes = Counter(k for k, v in spots.items() for x in v if x.gw and x.dist is None)
    gw_number = Counter(x.gw for x in group if x.gw).most_common(1)
    gw_number = gw_number[0][0] if gw_number else "?"
    notes = []
    if gw_votes:
        origin = gw_votes.most_common(1)[0][0]
    else:
        origin = min(keys, key=lambda k: min(x.counter or 9999 for x in spots[k]))
        notes.append("No shot is labeled as the weld; measuring from the first spot shot.")
    d_side = sum(1 if along(k) > along(origin) else -1
                 for k, v in spots.items() for x in v if x.side == "D")
    sign = 1 if d_side >= 0 else -1

    out = []
    for k, members in spots.items():
        pos = sign * (along(k) - along(origin)) - shift
        feet = round(pos)
        if abs(pos) < 0.5:
            name = f"GW {gw_number}"
        else:
            name = f"{abs(feet)}' {'D' if pos > 0 else 'U'}/S GW {gw_number}"
        spot = Spot(pos, sorted(members, key=lambda x: x.counter or 9999), name)
        zs = sorted(x.z for x in members)
        mid = zs[len(zs) // 2]
        for x in spot.shots:
            if len(members) > 1 and abs(x.z - mid) > 0.25:
                spot.dropped.append((x, f"elevation {x.z - mid:+.2f} ft off the other shots "
                                        "here - wrong target height?"))
            else:
                spot.used.append(x)
        while len(spot.used) > 3:
            extra = spot.used.pop()
            spot.dropped.append((extra, "more than 3 shots at this spot"))
        out.append(spot)
    out.sort(key=lambda sp: sp.along)
    if not any(sp.name.startswith("GW") for sp in out):
        notes.append("No shot sits on the weld itself, so there is no GW station.")
    gaps = [f"{(p.along + q.along) / 2:.1f}" for p, q in zip(out, out[1:])
            if q.along - p.along > 1.5 * median([y.along - x.along for x, y in zip(out, out[1:])])]
    if gaps:
        notes.append("No shots at about " + ", ".join(gaps) + " ft from the weld - marks "
                     "that were skipped in every pass.")
    return out, gw_number, notes


def position_names(spots: list) -> dict:
    """{shot index: new name, or None to leave the shot out}."""
    names = {}
    for sp in spots:
        for x in sp.used:
            names[x.index] = f"{sp.name} {x.counter:04d}"
        for x, _ in sp.dropped:
            names[x.index] = None
    return names


def label_mismatch_share(spots: list) -> float:
    """Share of footage-labeled shots whose label disagrees with their measured spot."""
    total = bad = 0
    for sp in spots:
        for x in sp.used + [d for d, _ in sp.dropped]:
            if x.station:
                total += 1
                bad += x.station != sp.name
    return bad / total if total else 0.0


def pipe_order(spots: list) -> list:
    """Spot names along the pipe, starting from the end the first shot was taken at."""
    names = [sp.name for sp in spots if sp.used]
    first = min((sp for sp in spots if sp.used), key=lambda sp: min(x.counter for x in sp.used))
    if abs(first.along - spots[-1].along) < abs(first.along - spots[0].along):
        names.reverse()
    return names
