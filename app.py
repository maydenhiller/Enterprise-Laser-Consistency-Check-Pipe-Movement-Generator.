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

# ---- step 1: shot names ----------------------------------------------------
st.header("Step 1 - Check shot names")
shots = [s for s in parsed.shots if s.profile in (profile, None)]
rows = []
for s in shots:
    changed = s.suggestion != s.raw_name
    problem = "; ".join(s.flags) or ("Formatting cleaned up" if changed else "")
    rows.append({"#": s.index + 1, "Raw name": s.raw_name, "Problem": problem,
                 "Use name": s.suggestion, "Confirmed": not problem,
                 "N": s.n, "E": s.e, "Elev": s.z, "_i": s.index})
names_df = pd.DataFrame(rows)
flagged = names_df["Problem"] != ""
st.write(f"{len(shots)} pipe shots for Profile {profile}; **{int(flagged.sum())} need a look**. "
         "The suggestion comes from the other shots taken at the same spot. Fix 'Use name' "
         "if it's wrong, then tick Confirmed.")
only_flagged = st.toggle("Show only shots with problems", value=bool(flagged.any()))
view = names_df[flagged] if only_flagged else names_df
edited = st.data_editor(
    view, hide_index=True, width="stretch", key=f"names_{profile}",
    disabled=["#", "Raw name", "Problem", "N", "E", "Elev"],
    column_config={"_i": None, "Problem": st.column_config.TextColumn(width="large")})
final = dict(zip(names_df["_i"], names_df["Use name"]))
final.update(dict(zip(edited["_i"], edited["Use name"])))
confirmed = dict(zip(names_df["_i"], names_df["Confirmed"]))
confirmed.update(dict(zip(edited["_i"], edited["Confirmed"])))
unconfirmed = [i for i in names_df.loc[flagged, "_i"] if not confirmed[i]]
if unconfirmed:
    st.warning(f"{len(unconfirmed)} flagged shot(s) not confirmed yet.")
    st.stop()

build = engine.build_profile(parsed.shots, final, profile)
for e in build.errors:
    st.error(e)
for w in build.warnings:
    st.warning(w)
if build.errors:
    st.stop()
stations, missing, extra = engine.order_stations(build, mv)
for k in missing:
    st.warning(f"{k} is in the Pipe Movement sheet but wasn't shot this profile. It will be "
               f"noted as covered back up; that profile's calc will show errors until it has "
               "coordinates.")
for s in extra:
    st.error(f"{s.station_id} isn't in the Pipe Movement sheet's station list. Check the name - "
             "it won't be written to the Pipe Movement file.")

# ---- step 2: notes and thickness ---------------------------------------------
st.header("Step 2 - Notes and coating / sleeve thickness")
st.write("Type the total coating or sleeve thickness (in) for **every** station - 0 if bare. "
         "Notes: GW, skid support, B-sleeve, Petro sleeve, composite sleeve, coal tar, tape, FBE. "
         "Typing **no BS** in a station's notes leaves it out of the bending strain.")
prefill = engine.existing_notes(mv, profile)
prefill.update({k: v for k, v in engine.existing_notes(cc, profile).items() if v != ("", None)})
st_rows = []
for s in stations:
    notes, thick = prefill.get(s.station_id, ("", None))
    md = s.max_difference
    st_rows.append({"Station": s.station_id,
                    "Shots": ", ".join(f"{x.counter:04d}" for x in s.shots),
                    "Max diff (in)": round(md, 3) if md is not None else None,
                    "Check": "FAIL" if md is not None and md > engine.LIMIT_IN else
                             ("1 shot" if md is None else "OK"),
                    "Notes": notes or "", "Thickness (in)": thick})
st_df = st.data_editor(
    pd.DataFrame(st_rows), hide_index=True, width="stretch", key=f"stations_{profile}",
    disabled=["Station", "Shots", "Max diff (in)", "Check"],
    column_config={"Thickness (in)": st.column_config.NumberColumn(min_value=0.0, step=0.001,
                                                                   format="%.3f")})
blank = st_df["Thickness (in)"].isna().sum()
if blank:
    st.warning(f"{blank} station(s) have no thickness yet.")
    st.stop()
if st_df["Thickness (in)"].nunique() == 1 and len(st_df) > 1:
    st.warning(f"Every station has the same thickness ({st_df['Thickness (in)'].iloc[0]:g} in). "
               "Make sure that's what was measured.")
for s, (_, r) in zip(stations, st_df.iterrows()):
    s.notes, s.thickness = (r["Notes"] or "").strip(), float(r["Thickness (in)"])

# ---- step 3: consistency -----------------------------------------------------
st.header("Step 3 - Consistency check")
fails = build.failures
worst = max((s.max_difference for s in stations if s.max_difference is not None), default=None)
m1, m2, m3 = st.columns(3)
m1.metric("Max difference", f"{worst:.3f} in" if worst is not None else "-",
          delta=f"limit {engine.LIMIT_IN:.2f} in", delta_color="off")
m2.metric("Stations over the limit", f"{len(fails)} of {len(stations)}")
od = cc.header["Pipe OD (in)"] if isinstance(cc.header["Pipe OD (in)"], (int, float)) else None
pv = engine.preview([s.average for s in stations], [s.thickness for s in stations], od)
if pv:
    m3.metric("Fitted strain (average, preview)", f"{pv.get('max_fitted_strain', 0):.6f}",
              delta=f"max deflection {pv['max_deflection_in']:.2f} in", delta_color="off")
chart = pd.DataFrame({"Station": [s.station_id for s in stations],
                      "Difference (in)": [s.max_difference or 0 for s in stations]})
st.bar_chart(chart, x="Station", y="Difference (in)", height=260)
st.caption("A smooth curve or a sudden step across many stations points at something "
           "systematic (setup, prism handling, pipe moving), not random error.")
proceed = True
if fails:
    st.error("Over 0.20 in: " + ", ".join(f"{s.station_id} ({s.max_difference:.3f})" for s in fails))
    st.write("If you're still set up, shoot a 3rd round at these stations, continuing the shot "
             "numbers, export again and re-upload. Otherwise you can continue; the workbook will "
             "show 'Need 3rd Measurement'.")
    proceed = st.checkbox("Continue without a 3rd measurement")
if not proceed:
    st.stop()

# ---- step 4: files -------------------------------------------------------------
st.header("Step 4 - Download")
try:
    out_cc = engine.fill_consistency(cc, stations, build.exposed_length)
    out_mv = (engine.fill_movement(mv, profile, [s for s in stations if s not in extra],
                                   missing, build.exposed_length) if mv else None)
except Exception as exc:
    st.error(f"Couldn't write the workbooks: {exc}")
    st.stop()
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
d1, d2 = st.columns(2)
d1.download_button(f"Download {cc_file.name}", out_cc, file_name=cc_file.name, mime=XLSX)
if out_mv:
    d2.download_button(f"Download {mv_file.name}", out_mv, file_name=mv_file.name, mime=XLSX)
    d2.caption("Keep this file - upload it again at the next profile.")
else:
    d2.info("No Pipe Movement file uploaded, so only the Consistency Check was filled.")
st.caption("Values recalculate when the file is opened in Excel.")
