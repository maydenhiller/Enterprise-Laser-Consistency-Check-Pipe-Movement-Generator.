# Enterprise Laser Consistency Check / Pipe Movement Generator

Streamlit app that fills Enterprise's **EPROD Robotic Laser Consistency Check** and
**EPROD Robotic Laser Pipe Movement Calculation** (v2.4) workbooks from a Trimble Access
"CSV with attributes" export (S5 / S7 robotic total station, TSC5).

Headers (LID, AID, line name, dig #, OD, WT, SMYS) come from the
[Enterprise Dig File Generator](https://github.com/maydenhiller/Enterprise-Dig-File-Generator).
This app fills only the measurement area, plus Exposed Length when it's blank.

## One profile per run
1. Upload the Trimble CSV, the Consistency Check for this profile, and the Pipe Movement
   workbook. At Profile 1 and 2, upload the Pipe Movement file the app gave you last time.
2. **Shot names**: shots that sit on the same spot of pipe are grouped by coordinates.
   Mislabeled names are flagged with a suggested fix, for example a missing shot number,
   a GW typo, footage that doesn't match the distance from the weld station, a label that
   wasn't changed, or `''`. The tech confirms or edits each one.
3. **Notes and thickness**: typed for every station. You get a warning if every station has
   the same thickness. `no BS` in a station's notes leaves it out of the bending strain
   (that's the workbook's own rule).
4. **Consistency**: the 3D difference between measurements (#2 vs #1, #3 vs #1) for each
   station, against Enterprise's 0.20" limit, and a chart of the differences. You also get
   an on-site preview of max deflection and fitted curvature strain, using the same chord
   projection and 4th-order fit as the hidden Profile sheet. Failures show a warning; you
   can tick a box to continue without a 3rd measurement.
5. **Download**: both workbooks come back with the same file names they were uploaded with.

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
