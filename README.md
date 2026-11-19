# Enterprise Laser Consistency Check / Pipe Movement Generator

Streamlit app that fills Enterprise's **EPROD Robotic Laser Consistency Check** and
**EPROD Robotic Laser Pipe Movement Calculation** (v2.4) workbooks from a Trimble Access
"CSV with attributes" export (S5 / S7 robotic total station, TSC5).

Headers (LID, AID, line name, dig #, OD, WT, SMYS) come from the
[Enterprise Dig File Generator](https://github.com/maydenhiller/Enterprise-Dig-File-Generator).
This app fills only the measurement area, plus Exposed Length when it's blank.

## One dig per run
1. Upload the Trimble CSV export(s), one or more Consistency Checks (Profile 0, 1, 2 -
   whichever you have shots for) and the Pipe Movement workbook. Several CSVs are fine; a
   shot that's in more than one export (jobs are cumulative) is used once. Each Consistency
   Check is matched to its profile by its Profile No. cell.
2. **Download right away.** Every workbook is filled on upload and comes back with the same
   name, plus a zip of all of them. Profiles are written into the Pipe Movement file in
   order (0, then 1, then 2), with rows matched by station.
3. **Each profile tab** shows the consistency results, then asks what's on the pipe in three
   layers: one choice for everywhere, **ranges** (From station / To station / surface /
   thickness, one row per stretch such as a sleeve or tape cut; later rows win), then
   single-station changes in the full table. The downloads sit below the tabs. Choices match the As-Built builder - Bare pipe; existing coating
   Coal Tar, Polyken Tape; repair coating PROTAL-7200, DENSO 7200, SPC 4888, SPC 2831,
   Aquawrap; Petro Sleeve, B Sleeve; Other. Defaults: Profile 0 Coal Tar, Profile 1 Bare
   pipe, Profile 2 PROTAL-7200. Bare pipe writes 0; the weld station gets "GW" in its notes;
   'no BS' in Other notes leaves a station out of the bending strain. Blank thickness warns
   and is treated as 0. Both shots at a station share one entry (the workbook copies #1's
   notes and thickness to #2 and #3).
4. **Shot names are corrected automatically**: shots at the same spot of pipe are grouped by
   coordinates and a bad name takes the name of the other shot there (missing shot number,
   GW typo, a label that wasn't changed, `''`). Every fix is listed on the profile tab and can
   be overridden.
5. **Consistency**: per-station 3D difference between measurements against Enterprise's
   0.20" limit, a chart, and an on-site preview of max deflection and fitted strain. Failures
   warn; they don't block the download.

## Y direction
Enterprise: +X north (downstream), **+Y west**, +Z up. Trimble's Easting is east-positive,
so the app writes **Y = -Easting**. If an uploaded Pipe Movement file already holds data
written east-positive (made by hand, or before this change), the app detects it from the
coordinates and keeps that file's convention so profiles are never mirrored against each
other. Only the left/right sign of sideways deflection depends on this; deflection size,
strain and movement amounts do not.

## Rules the app follows
- **Shot order is kept.** Rows are written in the order the stations were shot. For Profile
  1 and up, rows follow the station order already in the Pipe Movement file, matched by
  station ID.
- **The first digit of the shot number is the profile** (0001 = Profile 0, 1001 = Profile 1).
  Jobs are cumulative, so one CSV can hold every profile; only the selected profile is used.
- **Stations not shot this profile**: if a station from an earlier profile isn't in this
  one, it gets the note "Covered back up before Profile N measurement was taken" and no
  coordinates.
- **Rows are never inserted or deleted.** The hidden sheets point at rows 8-107. A workbook
  with deleted rows, or with `#REF!` formulas, is rejected; start from a fresh template.
- **Cells are written straight into the sheet XML.** Charts, hidden sheets and formulas are
  copied through untouched, and the workbook recalculates when it's opened in Excel. Saving
  through openpyxl would strip the charts.
- **Units**: US Survey Feet only.

## Run
    pip install -r requirements.txt
    streamlit run app.py

## Tests
    python -m pytest -q tests

The tests use synthetic shots only. Don't commit Enterprise workbooks or field data to this
public repo; `.gitignore` excludes `*.xlsx` and `*.csv`.
