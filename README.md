# LUTO Economics Tab Exporter

Standalone Python tool for exporting LUTO2 Economics tab data into Excel workbooks by state and region.

This project is separate from the older `C:\LUTO_Extractor` dashboard-summary exporter. Do not mix the two projects.

---

## 1. What this exporter does

The exporter reads LUTO2 dashboard JavaScript data files from each run's:

```text
DATA_REPORT\data
```

folder and creates Excel workbooks by state.

Each state workbook contains:

- one sheet per NRM region
- one final direct state sheet for multi-region states
- direct `region_state` values where available, so state-level values match the values shown directly in LUTO2

Each sheet contains:

```text
Sum
Sum - Profitability

Ag
Ag - Profitability

Ag Mgt
Ag Mgt - Profitability

Non-Ag
Non-Ag - Profitability
```

The `Ag Mgt` section also includes the template revenue/cost breakdown for:

```text
Onshore Wind
Utility Solar PV
```

including:

```text
Revenue
Costs (opex)
Costs (capex)
Costs (other)
```

---

## 2. Why this version was created

The previous exporter calculated state totals by summing NRM-region sheets.

That caused small differences from LUTO2 because LUTO2 stores and displays direct state-level values under:

```text
region_state
```

This exporter uses:

```text
regional sheets      = region_NRM
state-total sheets   = direct region_state values where available
single-region states = direct region_state values where available
```

This is why values such as Queensland `Ag profit 2050` now match the value shown directly in LUTO2.

---

## 3. Recommended folder location

Recommended local folder:

```text
C:\LUTO_Economics_Tab_Exporter
```

Keep the older/legacy exporter separate, for example:

```text
C:\LUTO_Extractor
```

or:

```text
C:\LUTO_Extractor_Legacy
```

---

## 4. Required project files

The project folder should contain:

```text
C:\LUTO_Economics_Tab_Exporter
│
├── luto_economics_tab_state_exporter.py
├── install_requirements_helper.py
├── 00_Install_Requirements.bat
├── 01_Run_All_States.bat
├── 02_Run_Queensland_Only.bat
├── 03_Run_All_Runs_All_States.bat
├── 04_Run_All_Runs_Queensland_Only.bat
├── README.md
├── .gitignore
└── outputs
    └── .gitkeep
```

The main script is:

```text
luto_economics_tab_state_exporter.py
```

---

## 5. Python requirements

Required Python packages:

```text
pandas
openpyxl
```

Python 3.9 or newer is recommended.

---

## 6. Install requirements

### Option A — use the BAT installer

Double-click:

```text
00_Install_Requirements.bat
```

### Option B — install manually in PowerShell

Open PowerShell and run:

```powershell
cd C:\LUTO_Economics_Tab_Exporter

python -m pip install pandas openpyxl
```

If `python` is not recognised, try:

```powershell
py -m pip install pandas openpyxl
```

If using Conda:

```powershell
conda install pandas openpyxl
```

### Check installation

```powershell
python -c "import pandas, openpyxl; print('Packages installed successfully')"
```

---

## 7. Correct input folder

For one LUTO2 run, the input must be the folder ending in:

```text
DATA_REPORT\data
```

Example:

```text
C:\LUTO_Economics_Tab_Exporter\Run_G0001\DATA_REPORT\data
```

or:

```text
S:\...\Run_G0001\DATA_REPORT\data
```

Do **not** use:

```text
DATA_REPORT\data\map_layers
```

For multiple runs, use the folder that contains the `Run_*` folders.

Example:

```text
S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 22 May
```

This folder should contain:

```text
Run_G0001
Run_G0002
Run_G0003
...
```

Each run folder should contain:

```text
DATA_REPORT\data
```

---

## 8. Run all RES5 runs for Alexis

For Alexis's current RES5 runs, use this reports base folder:

```text
S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 22 May
```

Run:

```powershell
cd C:\LUTO_Economics_Tab_Exporter

python ".\luto_economics_tab_state_exporter.py" --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 22 May" --output-dir "C:\LUTO_Economics_Tab_Exporter\outputs" --region-level auto --start-year 2020 --end-year 2050
```

If `python` does not work, use:

```powershell
cd C:\LUTO_Economics_Tab_Exporter

py ".\luto_economics_tab_state_exporter.py" --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 22 May" --output-dir "C:\LUTO_Economics_Tab_Exporter\outputs" --region-level auto --start-year 2020 --end-year 2050
```

---

## 9. Run one sample run

If your sample data is here:

```text
C:\LUTO_Economics_Tab_Exporter\Run_G0001\DATA_REPORT\data
```

run:

```powershell
cd C:\LUTO_Economics_Tab_Exporter

python ".\luto_economics_tab_state_exporter.py" --data-dir "C:\LUTO_Economics_Tab_Exporter\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Economics_Tab_Exporter\outputs" --output-prefix Run_G0001 --region-level auto --start-year 2020 --end-year 2050
```

If the sample data is still in the old folder:

```text
C:\LUTO_Extractor\Run_G0001\DATA_REPORT\data
```

run:

```powershell
cd C:\LUTO_Economics_Tab_Exporter

python ".\luto_economics_tab_state_exporter.py" --data-dir "C:\LUTO_Extractor\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Economics_Tab_Exporter\outputs" --output-prefix Run_G0001 --region-level auto --start-year 2020 --end-year 2050
```

---

## 10. BAT file run options

The BAT files are optional. If your system blocks BAT files, use the PowerShell commands above.

### One run, all states

Double-click:

```text
01_Run_All_States.bat
```

### One run, Queensland only

Double-click:

```text
02_Run_Queensland_Only.bat
```

### All `Run_*` folders, all states

Double-click:

```text
03_Run_All_Runs_All_States.bat
```

### All `Run_*` folders, Queensland only

Double-click:

```text
04_Run_All_Runs_Queensland_Only.bat
```

---

## 11. Output folder

All Excel reports are written to:

```text
C:\LUTO_Economics_Tab_Exporter\outputs
```

Example output:

```text
outputs
│
├── Australian_Capital_Territory
│   └── Run_G0001_Australian_Capital_Territory_Economics_Tab_Overview.xlsx
│
├── New_South_Wales
│   └── Run_G0001_New_South_Wales_Economics_Tab_Overview.xlsx
│
├── Queensland
│   └── Run_G0001_Queensland_Economics_Tab_Overview.xlsx
│
├── Victoria
│   └── Run_G0001_Victoria_Economics_Tab_Overview.xlsx
│
└── Economics_Tab_Overview_State_Export_Manifest.csv
```

For batch mode with multiple runs, each state folder will contain one workbook per run.

---

## 12. Expected successful run message

A successful run should show messages similar to:

```text
Found 1 run(s).

================================================================================
Processing Run_G0001
Data folder: C:\...\Run_G0001\DATA_REPORT\data
================================================================================
Region level used: region_NRM
Regions detected: 56
Saved: C:\...\outputs\Queensland\Run_G0001_Queensland_Economics_Tab_Overview.xlsx
Manifest saved: C:\...\outputs\Economics_Tab_Overview_State_Export_Manifest.csv
Finished successfully.
```

If the log shows:

```text
Unknown_State
```

something is wrong with either the region mapping or the source structure.

---

## 13. Important source files used

The exporter reads the LUTO2 dashboard JS files, including:

```text
Economics_overview_sum.js
Economics_overview_Ag.js
Economics_overview_Am.js
Economics_overview_Non_Ag.js
Economics_Sum.js
Economics_Am_revenue.js
Economics_Am_cost.js
```

The exporter does **not** blindly sum every detailed JS series because some detailed files contain overlapping totals/subtotals.

---

## 14. Profitability calculation

There is no explicit `Profitability` field in the JS files.

The exporter calculates:

```text
Profitability (%) = Profit / Total revenue × 100
```

For state sheets, the exporter uses direct LUTO2 state-level values where available.

Blank cells mean no value exists in the source data. Missing values are not forced to zero.

---

## 15. Common errors and fixes

### Error: `python is not recognised`

Try:

```powershell
py ".\luto_economics_tab_state_exporter.py" --help
```

or install Python / check your PATH.

### Error: missing `pandas` or `openpyxl`

Run:

```powershell
python -m pip install pandas openpyxl
```

or:

```powershell
py -m pip install pandas openpyxl
```

### Error: no run folders found

Make sure `--reports-base-dir` points to the folder containing `Run_*` folders, not to a specific `DATA_REPORT\data` folder.

Correct for batch mode:

```text
...\RES5 22 May
```

Incorrect for batch mode:

```text
...\Run_G0001\DATA_REPORT\data
```

### Error: no output created

Check that the selected data folder contains files such as:

```text
Economics_overview_sum.js
Economics_overview_Ag.js
Economics_overview_Am.js
Economics_overview_Non_Ag.js
```

### Error: S drive is not accessible

Make sure the `S:` drive is mapped and accessible in File Explorer.

Test in PowerShell:

```powershell
Test-Path "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 22 May"
```

Expected result:

```text
True
```

### Error: BAT files do not run

Use the direct PowerShell command instead. BAT files are only a convenience.

---

## 16. GitHub / version-control notes

Do not commit generated Excel outputs or large LUTO input data.

The `.gitignore` should exclude:

```text
outputs/*
!outputs/.gitkeep
*.xlsx
*.xls
*.csv
Run_*/
DATA_REPORT/
Report_Data/
map_layers/
*.zip
__pycache__/
*.pyc
```

Only source code, README files, BAT files, `.gitignore`, and `outputs/.gitkeep` should be committed.

---

## 17. Quick commands

### Check project status

```powershell
cd C:\LUTO_Economics_Tab_Exporter
git status
```

### Commit code changes

```powershell
git add .gitignore README.md install_requirements_helper.py 00_Install_Requirements.bat 01_Run_All_States.bat 02_Run_Queensland_Only.bat 03_Run_All_Runs_All_States.bat 04_Run_All_Runs_Queensland_Only.bat luto_economics_tab_state_exporter.py outputs/.gitkeep
git commit -m "Update LUTO economics tab exporter"
git push
```

### Run all RES5 runs

```powershell
cd C:\LUTO_Economics_Tab_Exporter

python ".\luto_economics_tab_state_exporter.py" --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 22 May" --output-dir "C:\LUTO_Economics_Tab_Exporter\outputs" --region-level auto --start-year 2020 --end-year 2050
```
