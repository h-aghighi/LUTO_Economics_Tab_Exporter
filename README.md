# LUTO Economics Tab Exporter

This is the standalone project for the current Alexis economics-tab Excel export.

It is intentionally separated from the older `LUTO_Extractor` / legacy dashboard-summary exporter so the two workflows do not get mixed.

## What this project produces

The exporter creates Excel workbooks by state.

Each workbook contains:

- one sheet per region
- a final direct state sheet for multi-region states
- direct `region_state` values where available, so the state sheet matches LUTO2 directly

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

The Ag Mgt section also includes the template revenue/cost breakdown for:

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

## Recommended project location

Use a separate folder such as:

```text
C:\LUTO_Economics_Tab_Exporter
```

Keep the old/legacy project somewhere else, such as:

```text
C:\LUTO_Extractor_Legacy
```

Do not mix the two projects.

## Required files

```text
LUTO_Economics_Tab_Exporter
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
```

## Install requirements

Double-click:

```text
00_Install_Requirements.bat
```

or run:

```powershell
python -m pip install pandas openpyxl
```

## Run one sample run

If your sample data is in:

```text
C:\LUTO_Economics_Tab_Exporter\Run_G0001\DATA_REPORT\data
```

run:

```powershell
cd C:\LUTO_Economics_Tab_Exporter

python ".\luto_economics_tab_state_exporter.py" --data-dir "C:\LUTO_Economics_Tab_Exporter\Run_G0001\DATA_REPORT\data" --output-dir "C:\LUTO_Economics_Tab_Exporter\outputs" --output-prefix Run_G0001 --region-level auto --start-year 2020 --end-year 2050
```

## Run all Run_* folders

If the project folder contains:

```text
Run_G0001
Run_G0002
Run_G0003
```

then run:

```powershell
cd C:\LUTO_Economics_Tab_Exporter

python ".\luto_economics_tab_state_exporter.py" --reports-base-dir "C:\LUTO_Economics_Tab_Exporter" --output-dir "C:\LUTO_Economics_Tab_Exporter\outputs" --region-level auto --start-year 2020 --end-year 2050
```

## Output folder

All Excel reports are written to:

```text
outputs
```

Generated Excel files should not be committed to GitHub.

## Notes

- State sheets use direct `region_state` values where available.
- Regional sheets use `region_NRM`.
- This avoids the earlier mismatch between NRM-summed state values and direct LUTO2 state values.
- Profitability is calculated because there is no explicit profitability field in the JS files.
