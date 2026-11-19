"""Enterprise Laser Consistency Check / Pipe Movement Generator (Streamlit).

Run one profile at a time. Upload the Trimble CSV export, the Consistency Check
workbook for that profile and (optionally) the Pipe Movement workbook - the one
this app gave you at the last profile. Download both back with the same names.
"""
import pandas as pd
import streamlit as st

import engine

st.set_page_config(page_title="Laser Consistency / Pipe Movement", layout="wide")
st.title("Laser Consistency Check & Pipe Movement Generator")
st.caption("One profile at a time. Run it on site before you pack up so a 3rd round can be "
           "shot while the robot is still set up. Never insert or delete rows in the workbooks.")


@st.cache_data(show_spinner=False)
def load_csv(data: bytes):
    parsed = engine.parse_trimble_csv(data)
    engine.review_names(parsed.shots)
    return parsed


@st.cache_data(show_spinner=False)
def load_book(data: bytes):
    return engine.read_workbook(data)


c1, c2, c3 = st.columns(3)
csv_file = c1.file_uploader("1. Trimble export (.csv)", type=["csv"])
cc_file = c2.file_uploader("2. Consistency Check for this profile (.xlsx)", type=["xlsx"])
mv_file = c3.file_uploader("3. Pipe Movement Calculation (.xlsx) - optional", type=["xlsx"],
                           help="At Profile 1 and 2, upload the Pipe Movement file this app "
                                "gave you at the previous profile.")
if not (csv_file and cc_file):
    st.info("Upload the CSV and the Consistency Check to start.")
    st.stop()

parsed = load_csv(csv_file.getvalue())
try:
    cc = load_book(cc_file.getvalue())
    mv = load_book(mv_file.getvalue()) if mv_file else None
except Exception as exc:  # unreadable workbook
    st.error(str(exc))
    st.stop()

stop = False
if cc.kind != "consistency":
    st.error("Box 2 holds a Pipe Movement workbook - it needs the Consistency Check.")
    stop = True
if mv is not None and mv.kind != "movement":
    st.error("Box 3 holds a Consistency Check - it needs the Pipe Movement workbook.")
    stop = True
for label, book in (("Consistency Check", cc), ("Pipe Movement", mv)):
    if book is not None:
        for e in book.errors:
            st.error(f"{label}: {e}")
            stop = True
        for w in book.warnings:
            st.warning(f"{label}: {w}")
for e in parsed.errors:
    st.error(e)
    stop = True
if stop:
    st.stop()
for w in parsed.warnings:
    st.caption(w)
if mv is not None:
    for m in engine.header_mismatches(cc, mv):
        st.warning(m)
for note in engine.control_check(parsed.control):
    st.warning(note)

# ---- profile -------------------------------------------------------------
in_csv = sorted({s.profile for s in parsed.shots if s.profile is not None})
book_profile = engine.profile_number(cc)
default = book_profile if book_profile in in_csv else (in_csv[-1] if in_csv else 0)
profile = st.selectbox("Profile", in_csv or [0], index=(in_csv or [0]).index(default),
                       help="Taken from the shot numbers (0001 = Profile 0, 1001 = Profile 1).")
if book_profile is not None and book_profile != profile:
    st.warning(f"This Consistency Check is set up as Profile {book_profile}, not Profile {profile}. "
               "Make sure you uploaded the right file.")
if mv is not None and engine.movement_has_profile(mv, profile):
    st.warning(f"The Pipe Movement file already has Profile {profile} data. It will be replaced.")

# Files are filled from the moment everything is uploaded. Everything below this box
# (pipe surface, name fixes) adjusts what goes in it; it refreshes on every change.
download_box = st.container(border=True)

# ---- name fixes (applied automatically) ---------------------------------------
shots = [s for s in parsed.shots if s.profile in (profile, None)]
fix_rows = []
for s in shots:
    if s.flags or s.suggestion != s.raw_name:
        fix_rows.append({"#": s.index + 1, "Raw name": s.raw_name, "Used as": s.suggestion,
                         "Why": "; ".join(s.flags) or "Formatting cleaned up", "_i": s.index})
final = {s.index: s.suggestion for s in shots}

# ---- pipe surface ---------------------------------------------------------------
SURFACES = ["Coating", "Bare pipe", "Sleeve"]
st.header("What's on the pipe at each station?")
st.write("Both shots at a station share one surface entry (the workbook copies measurement "
         "#1's notes and thickness to #2 and #3). Set what most of the pipe has here, then "
         "change any stations that differ in the table below.")
a, b, c = st.columns(3)
default_surface = a.selectbox("Most of the pipe is", SURFACES, key=f"surf_{profile}")
default_type = b.text_input("Type (coal tar, tape, FBE, B-sleeve, Petro sleeve...)",
                            key=f"type_{profile}", disabled=default_surface == "Bare pipe")
default_thick = c.number_input("Total thickness (in)", min_value=0.0, step=0.005, format="%.3f",
                               value=None, key=f"thick_{profile}",
                               disabled=default_surface == "Bare pipe")

# Build the profile with the auto-fixed names first, so stations are known.
build = engine.build_profile(parsed.shots, final, profile)
stations, missing, extra = engine.order_stations(build, mv)
prefill = engine.existing_notes(mv, profile)
prefill.update({k: v for k, v in engine.existing_notes(cc, profile).items() if v != ("", None)})


def surface_row(station):
    if station.station_id in prefill:          # re-upload: keep what's already in the file
        notes, thick = prefill[station.station_id]
        surface = "Bare pipe" if thick == 0 else ("Sleeve" if "sleeve" in str(notes).lower()
                                                  else "Coating")
        return surface, str(notes or ""), thick
    if default_surface == "Bare pipe":
        return "Bare pipe", "", 0.0
    return default_surface, default_type, default_thick


st_rows = []
for s in stations:
    surface, kind, thick = surface_row(s)
    md = s.max_difference
    st_rows.append({"Station": s.station_id,
                    "Shots": ", ".join(f"{x.counter:04d}" for x in s.shots),
                    "Surface": surface, "Type": kind, "Thickness (in)": thick, "Other notes": "",
                    "Max diff (in)": round(md, 3) if md is not None else None})
st_df = st.data_editor(
    pd.DataFrame(st_rows), hide_index=True, width="stretch",
    key=f"stations_{profile}",
    disabled=["Station", "Shots", "Max diff (in)"],
    column_config={
        "Surface": st.column_config.SelectboxColumn(options=SURFACES, required=True),
        "Thickness (in)": st.column_config.NumberColumn(min_value=0.0, step=0.001, format="%.3f"),
        "Other notes": st.column_config.TextColumn(help="Skid support, etc. 'no BS' leaves the "
                                                        "station out of the bending strain."),
    })


def compose_notes(station_id, row):
    parts = []
    if "'" not in station_id:                   # weld station: "U/S GW 181000"
        parts.append("GW")
    if row["Surface"] == "Bare pipe":
        parts.append("Bare")
    else:
        parts.append((row["Type"] or "").strip() or row["Surface"])
    parts.append((row["Other notes"] or "").strip())
    seen, out = set(), []
    for chunk in ", ".join(p for p in parts if p).split(","):
        chunk = chunk.strip()
        if chunk and chunk.lower() not in seen:
            seen.add(chunk.lower())
            out.append(chunk)
    return ", ".join(out)


no_thickness = []
for s, (_, r) in zip(stations, st_df.iterrows()):
    s.notes = compose_notes(s.station_id, r)
    if r["Surface"] == "Bare pipe":
        s.thickness = 0.0
    elif pd.isna(r["Thickness (in)"]):
        s.thickness = None
        no_thickness.append(s.station_id)
    else:
        s.thickness = float(r["Thickness (in)"])

# ---- consistency ---------------------------------------------------------------
st.header("Consistency check")
fails = build.failures
worst = max((s.max_difference for s in stations if s.max_difference is not None), default=None)
m1, m2, m3 = st.columns(3)
m1.metric("Max difference", f"{worst:.3f} in" if worst is not None else "-",
          delta=f"limit {engine.LIMIT_IN:.2f} in", delta_color="off")
m2.metric("Stations over the limit", f"{len(fails)} of {len(stations)}")
od = cc.header["Pipe OD (in)"] if isinstance(cc.header["Pipe OD (in)"], (int, float)) else None
if stations:
    pv = engine.preview([s.average for s in stations], [s.thickness for s in stations], od)
    if pv:
        m3.metric("Fitted strain (average, preview)", f"{pv.get('max_fitted_strain', 0):.6f}",
                  delta=f"max deflection {pv['max_deflection_in']:.2f} in", delta_color="off")
    chart = pd.DataFrame({"Station": [s.station_id for s in stations],
                          "Difference (in)": [s.max_difference or 0 for s in stations]})
    st.bar_chart(chart, x="Station", y="Difference (in)", height=260)
    st.caption("A smooth curve or a sudden step across many stations points at something "
               "systematic (setup, prism handling, pipe moving), not random error.")
if fails:
    st.write("Over 0.20 in: " + ", ".join(f"{s.station_id} ({s.max_difference:.3f})" for s in fails))

# ---- name fixes table ------------------------------------------------------------
st.header("Shot names")
if fix_rows:
    st.write(f"{len(fix_rows)} shot name(s) were corrected automatically, using the other shots "
             "taken at the same spot. If one is wrong, change 'Used as'.")
    fixes = st.data_editor(pd.DataFrame(fix_rows), hide_index=True, width="stretch",
                           key=f"fixes_{profile}", disabled=["#", "Raw name", "Why"],
                           column_config={"_i": None,
                                          "Why": st.column_config.TextColumn(width="large")})
    edits = {i: n for i, n in zip(fixes["_i"], fixes["Used as"]) if n != final[i]}
    if edits:
        final.update(edits)
        build = engine.build_profile(parsed.shots, final, profile)
        old = {s.station_id: s for s in stations}
        stations, missing, extra = engine.order_stations(build, mv)
        for s in stations:                         # keep the surface entries just made
            if s.station_id in old:
                s.notes, s.thickness = old[s.station_id].notes, old[s.station_id].thickness
else:
    st.write("All shot names read cleanly.")

# ---- fill the download box --------------------------------------------------------
with download_box:
    blocking = list(build.errors)
    if blocking:
        st.error("Can't fill the workbooks until these are fixed (see Shot names below):")
        for e in blocking:
            st.write("- " + e)
    else:
        try:
            out_cc = engine.fill_consistency(cc, stations, build.exposed_length)
            out_mv = (engine.fill_movement(mv, profile, [s for s in stations if s not in extra],
                                           missing, build.exposed_length) if mv else None)
        except Exception as exc:
            st.error(f"Couldn't write the workbooks: {exc}")
        else:
            st.subheader(f"Profile {profile}: {len(stations)} stations filled")
            XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            d1, d2 = st.columns(2)
            d1.download_button(f"Download {cc_file.name}", out_cc, file_name=cc_file.name,
                               mime=XLSX, type="primary")
            if out_mv:
                d2.download_button(f"Download {mv_file.name}", out_mv, file_name=mv_file.name,
                                   mime=XLSX, type="primary")
                d2.caption("Keep this file. Upload it again at the next profile.")
            else:
                d2.info("No Pipe Movement file uploaded, so only the Consistency Check was filled.")
    # What to look at before sending
    if fix_rows:
        st.write(f"- {len(fix_rows)} shot name(s) auto-corrected (listed under Shot names).")
    if no_thickness:
        st.warning(f"No thickness entered for {len(no_thickness)} station(s); left blank, so "
                   "the calc treats them as 0. Set it under 'What's on the pipe'.")
    if fails:
        st.warning(f"{len(fails)} of {len(stations)} stations are over the 0.20 in limit. "
                   "The workbook will say 'Need 3rd Measurement'. If you're still set up, shoot "
                   "a 3rd round at those stations, export and re-upload.")
    for w in build.warnings:
        st.warning(w)
    for k in missing:
        st.warning(f"{k} wasn't shot this profile; noted as covered back up. That profile's calc "
                   "will show errors until it has coordinates.")
    for s in extra:
        st.error(f"{s.station_id} isn't in the Pipe Movement sheet's station list, so it was left "
                 "out of the Pipe Movement file. Check the name under Shot names.")
    st.caption("Values recalculate when the file is opened in Excel. Never insert or delete rows.")
