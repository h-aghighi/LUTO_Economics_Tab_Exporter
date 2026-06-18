# LUTO2 Economics State Excel Exporter

A standalone Python tool for exporting LUTO2 Economics data into organised Excel workbooks by state and NRM region.

The exporter reads the original JavaScript data files used by the LUTO2 dashboard. It does not estimate values from screenshots or read values from the displayed graphs. It follows the same filters, totals and stacked-series logic used to construct the graphs.

This project is separate from the older `C:\LUTO_Extractor` dashboard-summary exporter.

---

## 1. Main script

The final exporter is:

```text
luto2_economics_state_excel_exporter.py
```

Recommended local project folder:

```text
C:\LUTO_Economics_Tab_Exporter
```

Repository:

```text
https://github.com/h-aghighi/LUTO_Economics_Tab_Exporter
```

---

## 2. What the exporter produces

For each state, the script creates an Excel workbook containing:

- one worksheet per NRM region
- one direct state-level worksheet where state-level data are available
- annual values from 2020 to 2050 by default

Each worksheet includes:

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

The `Ag Mgt` section also contains renewable-energy detail for:

```text
Onshore Wind
Utility Solar PV
```

including:

```text
Profit
Revenue
Costs (opex)
Costs (capex)
Costs (other)
```

`Costs (other)` remains blank when the source data do not contain a separate Other Cost category.

---

## 3. Extraction method

The values are extracted from the original LUTO2 data files used to create the Economics graphs.

The exporter:

1. selects the required state or NRM region;
2. follows the same dashboard filters used for the graph;
3. reads a direct total where the source provides an `ALL` series;
4. sums the underlying stacked categories where the graph total is built from several components;
5. avoids adding overlapping totals and subtotals;
6. writes the verified annual values into Excel.

Examples:

```text
Ag Mgt Cost
= sum of all land-use values under the selected ALL cost hierarchy

Onshore Wind Profit
= sum of all land-use values under Onshore Wind > ALL

Utility Solar PV Profit
= sum of all land-use values under Utility Solar PV > ALL
```

The exporter uses direct `region_state` values where available so state sheets match the state-level values shown by LUTO2.

---

## 4. Python requirements

Required packages:

```text
pandas
openpyxl
```

Python 3.9 or newer is recommended.

Install the packages in PowerShell:

```powershell
python -m pip install pandas openpyxl
```

For Conda:

```powershell
conda install pandas openpyxl
```

Check the installation:

```powershell
python -c "import pandas, openpyxl; print('Packages installed successfully')"
```

---

## 5. Input folder structure

### Multiple runs

Use `--reports-base-dir` with the folder that directly contains the `Run_*` folders.

Current example:

```text
S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 14 June\Report_Data
```

Expected structure:

```text
Report_Data
├── Run_G0001
├── Run_G0002
├── Run_G0003
├── Run_G0004
├── Run_G0005
└── Run_G0006
```

Inside each run, the script searches for the folder containing:

```text
Economics_overview_sum.js
```

Normally this is:

```text
Run_G0001\DATA_REPORT\data
```

or:

```text
Run_G0001\data
```

### One run

For a single run, use `--data-dir` and point directly to the folder containing the Economics JavaScript files:

```text
...\Run_G0001\DATA_REPORT\data
```

Do not use:

```text
...\DATA_REPORT\data\map_layers
```

---

## 6. Recommended output folder

Save the extracted Excel files beside the run folders in a clearly named output directory:

```text
Economics_Tab_Excel_Exports
```

For the current RES5 14 June data, the output folder is:

```text
S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 14 June\Report_Data\Economics_Tab_Excel_Exports
```

This keeps the source runs and their extracted Economics reports together without placing files inside individual run folders.

---

## 7. Run all runs and all states

Open Anaconda PowerShell and run:

```powershell
conda activate luto_map

cd "C:\LUTO_Economics_Tab_Exporter"

python ".\luto2_economics_state_excel_exporter.py" `
  --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 14 June\Report_Data" `
  --output-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 14 June\Report_Data\Economics_Tab_Excel_Exports" `
  --region-level auto `
  --start-year 2020 `
  --end-year 2050
```

This command processes all detected `Run_*` folders and all states.

---

## 8. Run a Queensland test first

```powershell
conda activate luto_map

cd "C:\LUTO_Economics_Tab_Exporter"

python ".\luto2_economics_state_excel_exporter.py" `
  --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 14 June\Report_Data" `
  --output-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 14 June\Report_Data\Economics_Tab_Excel_Exports" `
  --states "Queensland" `
  --region-level auto `
  --start-year 2020 `
  --end-year 2050
```

---

## 9. Run one specific run

Example for `Run_G0001`:

```powershell
cd "C:\LUTO_Economics_Tab_Exporter"

python ".\luto2_economics_state_excel_exporter.py" `
  --data-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 14 June\Report_Data\Run_G0001\DATA_REPORT\data" `
  --output-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 14 June\Report_Data\Economics_Tab_Excel_Exports" `
  --output-prefix "Run_G0001" `
  --region-level auto `
  --start-year 2020 `
  --end-year 2050
```

To export only Queensland for this run, add:

```powershell
--states "Queensland"
```

---

## 10. Select particular runs

To process only selected runs from the reports base folder:

```powershell
python ".\luto2_economics_state_excel_exporter.py" `
  --reports-base-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 14 June\Report_Data" `
  --output-dir "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 14 June\Report_Data\Economics_Tab_Excel_Exports" `
  --run-names Run_G0001 Run_G0003 `
  --region-level auto `
  --start-year 2020 `
  --end-year 2050
```

---

## 11. Output structure

The script creates the output directory automatically.

Example:

```text
Report_Data
├── Run_G0001
├── Run_G0002
├── Run_G0003
├── Run_G0004
├── Run_G0005
├── Run_G0006
└── Economics_Tab_Excel_Exports
    ├── Australian_Capital_Territory
    │   ├── Run_G0001_Australian_Capital_Territory_Economics_Tab_Overview.xlsx
    │   └── ...
    ├── New_South_Wales
    │   ├── Run_G0001_New_South_Wales_Economics_Tab_Overview.xlsx
    │   └── ...
    ├── Queensland
    │   ├── Run_G0001_Queensland_Economics_Tab_Overview.xlsx
    │   ├── Run_G0002_Queensland_Economics_Tab_Overview.xlsx
    │   └── ...
    ├── Victoria
    │   └── ...
    └── Economics_Tab_Overview_State_Export_Manifest.csv
```

---

## 12. Important source files

The exporter reads Economics JavaScript files including:

```text
Economics_overview_sum.js
Economics_overview_Ag.js
Economics_overview_Am.js
Economics_overview_Non_Ag.js
Economics_Sum.js
Economics_Am_profit.js
Economics_Am_revenue.js
Economics_Am_cost.js
```

The script does not blindly sum all detailed series because some files contain overlapping totals, subtotals or alternative filter levels.

---

## 13. Profitability calculations

Costs are stored as negative values.

For the accounting-consistent Sum profitability table:

```text
Profit = Total revenue + Total cost
```

Profitability is:

```text
Profitability (%) = Profit / Total revenue × 100
```

Where a valid revenue value is not present, profitability is left blank rather than creating a misleading percentage.

Blank graph cells remain blank when the original source does not contain a value for that year.

---

## 14. Expected successful run

A successful batch run should show messages similar to:

```text
Found 6 run(s).

================================================================================
Processing Run_G0001
Data folder: S:\...\Report_Data\Run_G0001\DATA_REPORT\data
================================================================================

Region level used: region_NRM
Regions detected: ...
Saved: S:\...\Economics_Tab_Excel_Exports\Queensland\Run_G0001_Queensland_Economics_Tab_Overview.xlsx

Manifest saved:
S:\...\Economics_Tab_Excel_Exports\Economics_Tab_Overview_State_Export_Manifest.csv

Finished successfully.
```

---

## 15. Common errors

### `python` is not recognised

Activate the correct Conda environment:

```powershell
conda activate luto_map
```

Then test:

```powershell
python ".\luto2_economics_state_excel_exporter.py" --help
```

### Missing `pandas` or `openpyxl`

```powershell
python -m pip install pandas openpyxl
```

### No run folders found

`--reports-base-dir` must point to the folder containing the `Run_*` directories:

```text
...\RES5 14 June\Report_Data
```

It must not point directly to one run's `DATA_REPORT\data` folder.

### No Economics source files found

Check that each run contains:

```text
Economics_overview_sum.js
```

### The S drive is unavailable

Test the location:

```powershell
Test-Path "S:\VC-DVCResearch\CWA\Projects\Current\Net Zero Industrial Precincts\5 Energy Land Use 25-26\LUTO2 runs\RES5 14 June\Report_Data"
```

Expected result:

```text
True
```

---

## 16. Git and generated files

Generated Excel files and LUTO2 input data should not be committed to GitHub.

A suitable `.gitignore` includes:

```text
outputs*/
Economics_Tab_Excel_Exports/
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

The `Economics_Tab_Excel_Exports` folder in this workflow is located on the shared `S:` drive, outside the local Git repository.

---

## 17. Push the final code and README only

From the project folder:

```powershell
cd "C:\LUTO_Economics_Tab_Exporter"

git add -- "luto2_economics_state_excel_exporter.py" "README.md"

git diff --cached --name-only
```

The staged files should be:

```text
README.md
luto2_economics_state_excel_exporter.py
```

Commit and push:

```powershell
git commit -m "Update final Economics exporter and documentation"

git push origin main
```

Other modified or untracked local files are not included unless they are explicitly staged.
