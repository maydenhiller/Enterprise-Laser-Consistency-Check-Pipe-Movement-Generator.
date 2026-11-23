"""Enterprise Laser Consistency Check / Pipe Movement Generator (Streamlit).

One dig per run. Upload the Trimble CSV export(s), one or more Consistency Check
workbooks (Profile 0, 1, 2 - whichever you have) and the Pipe Movement workbook.
Every workbook comes back filled, with the same file name it was uploaded with.
"""
import io
import zipfile

import pandas as pd
import streamlit as st

import importlib

import engine
import xlsx_patch

# Streamlit Cloud can keep an older copy of these modules in memory after the repo
# updates while running the new app.py. Reload them every run so they always match.
xlsx_patch = importlib.reload(xlsx_patch)
engine = importlib.reload(engine)

st.set_page_config(page_title="Laser Consistency / Pipe Movement", layout="wide")
st.title("Laser Consistency Check & Pipe Movement Generator")
st.caption("One dig per run. Best run on site after each profile, so a 3rd round can be shot "
           "while the robot is still set up. Never insert or delete rows in the workbooks.")

# Pipe surface choices, matching the As-Built builder's lists.
EXISTING = ["Coal Tar", "Polyken Tape"]
REPAIR = ["PROTAL-7200", "DENSO 7200", "SPC 4888", "SPC 2831", "Aquawrap"]
SLEEVES = ["Petro Sleeve", "B Sleeve"]
SURFACES = ["Bare pipe"] + EXISTING + REPAIR + SLEEVES + ["Other"]
DEFAULT_SURFACE = {0: "Coal Tar", 1: "Bare pipe", 2: "PROTAL-7200", 3: "PROTAL-7200"}
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@st.cache_data(show_spinner=False)
def read_csv(data: bytes, version: int):
    return engine.parse_trimble_csv(data)


@st.cache_data(show_spinner=False)
def load_csvs(blobs: tuple, profiles: tuple, version: int):
    files = [engine.apply_file_profile(engine.parse_trimble_csv(b), p)
             for b, p in zip(blobs, profiles)]
    merged = engine.merge_parsed(files)
    engine.review_names(merged.shots)
    return merged


@st.cache_data(show_spinner=False)
def load_book(data: bytes, version: int):
    return engine.read_workbook(data)


c1, c2, c3 = st.columns(3)
csv_files = c1.file_uploader("1. Trimble export(s) (.csv)", type=["csv"],
                             accept_multiple_files=True,
                             help="One export with every profile, or one export per profile.")
cc_files = c2.file_uploader("2. Consistency Check(s) (.xlsx)", type=["xlsx"],
                            accept_multiple_files=True,
                            help="One per profile. Upload every profile you have shots for.")
mv_file = c3.file_uploader("3. Pipe Movement Calculation (.xlsx)", type=["xlsx"],
                           help="At Profile 0, the blank one from the dig file generator. "
                                "After that, the one this app gave you last time.")
if not (csv_files and cc_files):
    st.info("Upload the CSV export(s) and at least one Consistency Check to start.")
    st.stop()

# Which profile each CSV holds. One file with every profile (shot numbers 0001 / 1001 /
# 2001) works as before; a job kept per profile can number every profile from 0001 and
# is set here (picked up from "Profile 1" in the job or file name when it's there).
FROM_NUMBERS = "From shot numbers"
choices = [FROM_NUMBERS, 0, 1, 2, 3]
file_profiles = []
if len(csv_files) > 1 or any(engine.detect_file_profile(read_csv(f.getvalue(), engine.ENGINE_VERSION),
                                                         f.name) is not None for f in csv_files):
    st.write("**Profile in each CSV**")
    cols = st.columns(min(4, len(csv_files)))
for i, f in enumerate(csv_files):
    found = engine.detect_file_profile(read_csv(f.getvalue(), engine.ENGINE_VERSION), f.name)
    if len(csv_files) > 1 or found is not None:
        pick = cols[i % len(cols)].selectbox(
            f.name, choices, index=choices.index(found) if found is not None else 0,
            key=f"csvprof_{i}_{f.name}",
            format_func=lambda c: c if c == FROM_NUMBERS else f"Profile {c}",
            help="'From shot numbers' reads the profile from each shot's first digit "
                 "(0001 = Profile 0, 1001 = Profile 1). Pick a profile if this file is "
                 "a job for one profile only.")
    else:
        pick = FROM_NUMBERS
    file_profiles.append(None if pick == FROM_NUMBERS else pick)
parsed = load_csvs(tuple(f.getvalue() for f in csv_files), tuple(file_profiles),
                   engine.ENGINE_VERSION)
stop = False
for e in parsed.errors:
    st.error(e)
    stop = True

books = {}
for f in cc_files:
    try:
        book = load_book(f.getvalue(), engine.ENGINE_VERSION)
    except Exception as exc:
        st.error(f"{f.name}: {exc}")
        stop = True
        continue
    if book.kind != "consistency":
        st.error(f"{f.name} is a Pipe Movement workbook - upload it in box 3.")
        stop = True
        continue
    number = engine.profile_number(book)
    if number is None:
        st.error(f"{f.name}: its Profile No. cell (P1) is blank or not a number.")
        stop = True
    elif number in books:
        st.error(f"{f.name} and {books[number][0].name} are both Profile {number}.")
        stop = True
    else:
        books[number] = (f, book)
mv = None
if mv_file:
    try:
        mv = load_book(mv_file.getvalue(), engine.ENGINE_VERSION)
    except Exception as exc:
        st.error(f"{mv_file.name}: {exc}")
        stop = True
    if mv is not None and mv.kind != "movement":
        st.error(f"{mv_file.name} is a Consistency Check - box 3 needs the Pipe Movement workbook.")
        stop = True
for label, book in [(f.name, b) for f, b in books.values()] + \
                   ([(mv_file.name, mv)] if mv is not None else []):
    for e in book.errors:
        st.error(f"{label}: {e}")
        stop = True
    for w in book.warnings:
        st.warning(f"{label}: {w}")
    if mv is not None and book is not mv:
        for m in engine.header_mismatches(book, mv):
            st.warning(f"{label}: {m}")
if stop:
    st.stop()
for w in parsed.warnings:
    st.caption(w)
for note in engine.control_check(parsed.control):
    st.warning(note)

in_csv = sorted({s.profile for s in parsed.shots if s.profile is not None})
profiles = [p for p in sorted(books) if p in in_csv]
for p in sorted(books):
    if p not in in_csv:
        st.warning(f"{books[p][0].name} is Profile {p}, but no Profile {p} shots "
                   f"(numbers {p}001...) are in the CSV. It was skipped.")
for p in in_csv:
    if p not in books:
        st.caption(f"The CSV also has Profile {p} shots; no Consistency Check was uploaded for it.")
if not profiles:
    st.stop()


# ---- per-profile entry ----------------------------------------------------------
bm_dates = engine.dates_in_names(parsed.control)
if bm_dates:
    st.caption("Dates in the benchmark/backsight names: "
               + ", ".join(engine.date_text(d) for d in bm_dates))

results = {}
run_spots = {}          # spots of profiles already named in this run, for matching later ones
tabs = st.tabs([f"Profile {p}" for p in profiles])
for p, tab in zip(profiles, tabs):
    with tab:
        cc_name, cc = books[p]
        on_file = cc.header.get("Technician Name")
        tech = st.text_input(f"Technician who shot Profile {p}",
                             value="" if on_file in (None, "xx") else str(on_file).strip(),
                             key=f"tech_{p}",
                             help="Written to this profile's Consistency Check. Leave blank to "
                                  "keep what the file already has.").strip()
        d1, d2 = st.columns(2)
        start = d1.date_input(f"Profile {p} start date", key=f"start_{p}", format="MM/DD/YYYY",
                              value=engine.parse_date(cc.header.get("Date Start")))
        end = d2.date_input(f"Profile {p} end date", key=f"end_{p}", format="MM/DD/YYYY",
                            value=engine.parse_date(cc.header.get("Date End")))
        if start and end and end < start:
            st.warning("The end date is before the start date.")
        shots = [s for s in parsed.shots if s.profile == p]
        final = {s.index: s.suggestion for s in shots}

        # Two ways to name stations: the field labels (with obvious typos fixed), or the
        # measured position when the labels are off in a pattern. Position is picked
        # automatically when the labels can't produce a clean result.
        keep_odd = st.checkbox(
            "Keep shots with an elevation far off the others at their station",
            key=f"keep_odd_{p}",
            help="Off: a shot more than 0.25 ft off the other shot(s) at its station "
                 "(usually a wrong target height) is left out. On: every shot is averaged in.")
        label_build = engine.build_profile(parsed.shots, final, p)
        order = None
        modes = ["Measured position", "Field labels (typos fixed)"]
        mode = st.radio("Name stations by", modes, horizontal=True, key=f"mode_{p}", index=0,
                        help="Measured position: each station is named by where the robot "
                             "measured it, counting marks from the weld. Field labels: the "
                             "names the crew typed, with obvious typos fixed.")
        if mode == modes[0]:
            # The GW shot is always taken on the weld, so stations are measured straight
            # from it. Labels that sit a whole mark off the measured positions are only
            # reported, never used to move the weld.
            shift = 0.0
            off, _ = engine.suggest_shift(parsed.shots, p)
            if off:
                st.warning(f"Most field labels sit {abs(off):.2f} ft "
                           f"{'downstream' if off > 0 else 'upstream'} of where the robot measured "
                           "those shots. Stations are named from the GW shot as the weld; check "
                           "that the shots labeled 'GW' were really taken on the weld.")
            spots, gw_number, layout_notes = engine.position_layout(parsed.shots, p, shift, keep_odd)
            earlier = [q for q in sorted(run_spots) if q < p]
            refs = (engine.reference_stations(spots=run_spots[earlier[0]]) if earlier else
                    engine.reference_stations(book=mv) if p > 0 and mv is not None and any(
                        engine.movement_has_profile(mv, q) for q in range(p)) else [])
            if p > 0 and refs:
                matched = engine.snap_to_stations(spots, refs)
                st.caption(f"{matched} of {len(spots)} stations matched to Profile "
                           f"{earlier[0] if earlier else 0}'s stations by position.")
            elif not label_build.errors:
                same = sum(1 for sp in spots for x in sp.used
                           if x.station and x.station == sp.name)
                total = sum(1 for sp in spots for x in sp.used if x.station)
                if total and same / total < 0.5:
                    st.warning(f"Only {same} of {total} field labels match the measured "
                               "names. Check the weld offset, or switch to Field labels.")
            run_spots[p] = spots
            for n in layout_notes:
                st.caption(n)
            final = engine.position_names(spots)
            order = engine.pipe_order(spots)
            st.dataframe(pd.DataFrame([{
                "Distance from weld (ft)": round(sp.along, 2),
                "Station": sp.name,
                "Shots used": ", ".join(engine.shot_number(x) for x in sp.used),
                "Field labels": ", ".join(sorted({x.raw_name.rsplit(" ", 1)[0]
                                                  for x in sp.used + [d for d, _ in sp.dropped]})),
                "Left out": "; ".join(f"{engine.shot_number(x)} ({why})" for x, why in sp.dropped),
            } for sp in spots]), hide_index=True, width="stretch")
            fix_rows = [{"#": s.index + 1, "Raw name": s.raw_name,
                         "Used as": final.get(s.index) or "(left out)"}
                        for s in shots
                        if (final.get(s.index) or "").lower() != s.suggestion.lower()]
            renamed = bool(fix_rows)
            if fix_rows:
                with st.expander(f"Shots named differently from their field label - "
                                 f"{len(fix_rows)}", expanded=False):
                    st.dataframe(pd.DataFrame(fix_rows), hide_index=True, width="stretch")
        else:
            renamed = False
            fix_rows = [{"#": s.index + 1, "Raw name": s.raw_name, "Used as": s.suggestion,
                         "Why": "; ".join(s.flags) or "Formatting cleaned up", "_i": s.index}
                        for s in shots if s.flags or s.suggestion != s.raw_name]
            with st.expander(f"Shot names - {len(fix_rows)} corrected automatically",
                             expanded=False):
                if fix_rows:
                    st.write("Each bad name took the name of the other shot(s) at the same "
                             "spot. Change 'Used as' if one is wrong.")
                    fixes = st.data_editor(pd.DataFrame(fix_rows), hide_index=True,
                                           width="stretch", key=f"fixes_{p}",
                                           disabled=["#", "Raw name", "Why"],
                                           column_config={"_i": None, "Why": st.column_config
                                                          .TextColumn(width="large")})
                    final.update(dict(zip(fixes["_i"], fixes["Used as"])))
                else:
                    st.write("All shot names read cleanly.")

        build = engine.build_profile(parsed.shots, final, p, order, keep_odd)
        for note in engine.cross_file_check(parsed.shots, final, p):
            (st.warning if "won't" in note else st.caption)(note)
        if build.errors:
            for e in build.errors:
                st.error(e)
            results[p] = None
            continue

        # Consistency.
        fails = build.failures
        worst = max((s.max_difference for s in build.stations if s.max_difference is not None),
                    default=None)
        m1, m2 = st.columns(2)
        m1.metric("Max difference", f"{worst:.3f} in" if worst is not None else "-",
                  delta=f"limit {engine.LIMIT_IN:.2f} in", delta_color="off")
        m2.metric("Stations over the limit", f"{len(fails)} of {len(build.stations)}")
        st.bar_chart(pd.DataFrame({"Station": [s.station_id for s in build.stations],
                                   "Difference (in)": [s.max_difference or 0
                                                       for s in build.stations]}),
                     x="Station", y="Difference (in)", height=220)

        # Pipe surface: everywhere else -> ranges -> single-station changes.
        st.subheader("What's on the pipe at each station?")
        ids = [s.station_id for s in build.stations]
        a, b = st.columns([2, 1])
        default = a.selectbox("Everywhere not covered by a range below", SURFACES,
                              index=SURFACES.index(DEFAULT_SURFACE.get(p, "Bare pipe")),
                              key=f"surf_{p}")
        default_thick = b.number_input("Thickness (in)", min_value=0.0, step=0.005,
                                       format="%.3f", value=None, key=f"thick_{p}",
                                       disabled=default == "Bare pipe")
        st.write("**Ranges** - add a row for each stretch that's different (a sleeve, a tape "
                 "cut, a repair coating). From and To are included; later rows win where "
                 "ranges overlap.")
        ranges = st.data_editor(
            pd.DataFrame({"From": pd.Series(dtype="str"), "To": pd.Series(dtype="str"),
                          "Surface": pd.Series(dtype="str"),
                          "Thickness (in)": pd.Series(dtype="float")}),
            num_rows="dynamic", hide_index=True, width="stretch", key=f"ranges_{p}",
            column_config={
                "From": st.column_config.SelectboxColumn(options=ids, required=True),
                "To": st.column_config.SelectboxColumn(options=ids, required=True),
                "Surface": st.column_config.SelectboxColumn(options=SURFACES, required=True),
                "Thickness (in)": st.column_config.NumberColumn(min_value=0.0, step=0.001,
                                                                format="%.3f"),
            })
        # Start from what a re-uploaded file already holds, else the "everywhere" choice.
        prefill = {k: v for k, v in engine.existing_notes(cc, p).items() if v != ("", None)}
        assigned, others = {}, {}
        for k in ids:
            if k in prefill:
                notes, thick = prefill[k]
                surface = next((x for x in SURFACES if x.lower() in str(notes).lower()), "Other")
                assigned[k] = (surface, thick)
                others[k] = str(notes) if surface == "Other" else ""
            else:
                assigned[k] = (default, default_thick)
        for _, r in ranges.iterrows():
            if r["From"] in ids and r["To"] in ids and isinstance(r["Surface"], str):
                i, j = sorted((ids.index(r["From"]), ids.index(r["To"])))
                thick = None if pd.isna(r["Thickness (in)"]) else float(r["Thickness (in)"])
                for k in ids[i:j + 1]:
                    assigned[k] = (r["Surface"], thick)

        rows = []
        for s in build.stations:
            surface, thick = assigned[s.station_id]
            other = others.get(s.station_id, "") if surface == "Other" else ""
            md = s.max_difference
            rows.append({"Station": s.station_id,
                         "Shots": ", ".join(f"{x.counter:04d}" for x in s.shots),
                         "Surface": surface,
                         "Thickness (in)": 0.0 if surface == "Bare pipe" else thick,
                         "Other notes": other,
                         "Max diff (in)": round(md, 3) if md is not None else None})
        st.write("**Every station** - the result of the above. Change a single station here "
                 "if needed.")
        table = st.data_editor(
            pd.DataFrame(rows), hide_index=True, width="stretch", key=f"stations_{p}",
            disabled=["Station", "Shots", "Max diff (in)"],
            column_config={
                "Surface": st.column_config.SelectboxColumn(options=SURFACES, required=True),
                "Thickness (in)": st.column_config.NumberColumn(min_value=0.0, step=0.001,
                                                                format="%.3f"),
                "Other notes": st.column_config.TextColumn(
                    help="Skid support, etc. 'no BS' leaves the station out of bending strain."),
                "Max diff (in)": st.column_config.NumberColumn(
                    help="Largest difference between this station's shots. Limit 0.20 in."),
            })
        st.caption("Both shots at a station share one entry - the workbook copies #1's notes "
                   "and thickness to #2 and #3.")
        no_thick = []
        for s, (_, r) in zip(build.stations, table.iterrows()):
            parts = (["GW"] if "'" not in s.station_id else []) + \
                    ([r["Surface"]] if r["Surface"] != "Other" else []) + \
                    [str(r["Other notes"] or "").strip()]
            s.notes = ", ".join(x for x in parts if x)
            if r["Surface"] == "Bare pipe":
                s.thickness = 0.0
            elif pd.isna(r["Thickness (in)"]):
                s.thickness = None
                no_thick.append(s.station_id)
            else:
                s.thickness = float(r["Thickness (in)"])

        od = cc.header["Pipe OD (in)"]
        pv = engine.preview([s.average for s in build.stations],
                            [s.thickness for s in build.stations],
                            od if isinstance(od, (int, float)) else None)
        if pv:
            st.caption(f"Preview with these thicknesses: max deflection "
                       f"{pv['max_deflection_in']:.2f} in, fitted strain (average) "
                       f"{pv.get('max_fitted_strain', 0):.6f}. The workbook is the record.")
        results[p] = {"cc_name": cc_name, "cc": cc, "build": build, "fails": fails,
                      "no_thick": no_thick, "fixed": len(fix_rows), "tech": tech,
                      "dates": (start, end),
                      "by_position": renamed}


download_box = st.container(border=True)   # below the profile tabs

# ---- Y direction ------------------------------------------------------------------
first = next((r for r in results.values() if r), None)
existing_convention = (engine.detect_west_positive(mv, first["build"].stations)
                       if first and mv is not None else None)
west_positive = True if existing_convention is None else existing_convention

# ---- fill everything ----------------------------------------------------------------
outputs, notes = {}, []
crew = []
if mv is not None and mv.header.get("Technician Name") not in (None, "", "xx"):
    crew += [n.strip() for n in str(mv.header.get("Technician Name")).split("/") if n.strip()]
for p in profiles:
    if results.get(p) and results[p]["tech"]:
        crew.append(results[p]["tech"])
# Pipe Movement dates: start = Profile 0's start; end = the end date of the last
# profile on the dig (the highest-numbered one in this run or already in the file).
# Blank keeps what the file has.
mv_start = results[0]["dates"][0] if results.get(0) else None
done = [p for p in results if results[p]]
last_in_file = max([q for q in range(4) if mv is not None and engine.movement_has_profile(mv, q)],
                   default=-1)
last = max(done, default=-1)
mv_end = results[last]["dates"][1] if last >= 0 and last >= last_in_file else None
mv_dates = (mv_start, mv_end)
crew_text = " / ".join({n.lower(): n for n in reversed(crew)}[k] for k in dict.fromkeys(n.lower() for n in crew))  # unique, in order
mv_bytes = mv_file.getvalue() if mv_file else None
for p in profiles:
    r = results.get(p)
    if not r:
        continue
    current_mv = engine.read_workbook(mv_bytes) if mv_bytes else None
    for s in r["build"].stations:
        s.west_positive = west_positive
    stations, missing, extra = engine.order_stations(r["build"], current_mv)
    try:
        outputs[r["cc_name"].name] = engine.fill_consistency(r["cc"], stations,
                                                             r["build"].exposed_length,
                                                             r["tech"] or None, r["dates"])
        if current_mv is not None:
            mv_bytes = engine.fill_movement(current_mv, p, [s for s in stations if s not in extra],
                                            missing, r["build"].exposed_length,
                                            crew_text or None, mv_dates)
    except Exception as exc:
        notes.append(("error", f"Profile {p}: couldn't write the workbooks: {exc}"))
        continue
    if r.get("by_position"):
        notes.append(("info", f"Profile {p}: {r['fixed']} shot(s) named by where they were "
                              "measured rather than their field label (listed on its tab)."))
    elif r["fixed"]:
        notes.append(("info", f"Profile {p}: {r['fixed']} shot name(s) auto-corrected "
                              "(see Shot names on its tab)."))
    if r["no_thick"]:
        notes.append(("warning", f"Profile {p}: no thickness for {len(r['no_thick'])} station(s); "
                                 "left blank, so the calc treats them as 0."))
    if r["fails"]:
        notes.append(("warning", f"Profile {p}: {len(r['fails'])} of {len(stations)} stations over "
                                 "0.20 in - the workbook will say 'Need 3rd Measurement'."))
    for w in r["build"].warnings:
        notes.append(("warning", f"Profile {p}: {w}"))
    for k in missing:
        notes.append(("warning", f"Profile {p}: {k} wasn't shot; noted as covered back up. That "
                                 "profile's calc shows errors until it has coordinates."))
    for s in extra:
        notes.append(("error", f"Profile {p}: {s.station_id} isn't in the Pipe Movement station "
                               "list, so it was left out of that file. Check its name."))
if mv_bytes and mv_file and any(results.values()):
    outputs[mv_file.name] = mv_bytes

with download_box:
    if not outputs:
        st.error("Nothing could be filled yet - see the errors on the profile tabs.")
    else:
        done = [p for p in profiles if results.get(p)]
        st.subheader(f"Filled: Profile {', '.join(map(str, done))}"
                     + (" + Pipe Movement" if mv_file and mv_file.name in outputs else ""))
        cols = st.columns(min(4, len(outputs) + 1))
        for i, (name, data) in enumerate(outputs.items()):
            cols[i % len(cols)].download_button(name, data, file_name=name, mime=XLSX,
                                                key=f"dl_{name}")
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in outputs.items():
                archive.writestr(name, data)
        cols[len(outputs) % len(cols)].download_button(
            "Download all (.zip)", buffer.getvalue(), file_name="Laser workbooks.zip",
            mime="application/zip", type="primary")
        if not mv_file:
            st.info("No Pipe Movement file uploaded, so only the Consistency Checks were filled.")
    for kind, text in notes:
        getattr(st, kind)(text)
    st.caption("Y is written west-positive (+Y = west), as Enterprise specifies"
               if west_positive else
               "Y is written east-positive to match the Pipe Movement file you uploaded, "
               "which already uses Trimble's Easting as-is.")
    st.caption("Values recalculate when the files are opened in Excel.")
