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
2. **Download right away.** Both workbooks are filled the moment the files are uploaded, and
   they come back with the same file names. The box at the top lists anything to check.
3. **What's on the pipe**: set what most of the pipe has (Coating, Bare pipe or Sleeve, the
   type, and the total thickness), then change any station that differs in the table.
   - Bare pipe writes 0.
   - The weld station gets "GW" added to its notes automatically.
   - `no BS` in Other notes leaves that station out of the bending strain (that's the
     workbook's own rule).
   - Both shots at a station share one entry, because the workbook copies measurement #1's
     notes and thickness to #2 and #3.
   - If thickness is left blank, a warning shows and the workbook treats it as 0.

   The downloads update with every change.
4. **Shot names are corrected automatically.** Shots that sit on the same spot of pipe are
   grouped by coordinates, and a bad name takes the name of the other shot at that spot.
   That covers a missing shot number, a GW typo, a label that wasn't changed, or `''`. Every
   fix is listed under Shot names and can be overridden there.
5. **Consistency**: for each station, the 3D difference between measurements against
   Enterprise's 0.20" limit, plus a chart. There's also an on-site preview of max deflection
   and fitted strain, using the same math as the hidden Profile sheet. Failures are warned,
   but they don't block the download.

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
