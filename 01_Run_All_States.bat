@echo off
setlocal

echo ==========================================
echo LUTO Economics Tab Exporter
echo Single run - all states
echo ==========================================
echo.

set TOOL_DIR=%~dp0
set SCRIPT=%TOOL_DIR%luto_economics_tab_state_exporter.py
set OUTPUT_DIR=%TOOL_DIR%outputs

if not exist "%SCRIPT%" (
    echo ERROR: Could not find:
    echo %SCRIPT%
    echo.
    echo Make sure luto_economics_tab_state_exporter.py is in the same folder as this BAT file.
    echo.
    pause
    exit /b 1
)

if not exist "%OUTPUT_DIR%" (
    mkdir "%OUTPUT_DIR%"
)

echo Paste the DATA_REPORT\data folder for one run.
echo.
echo Example:
echo C:\LUTO_Extractor\Run_G0001\DATA_REPORT\data
echo.
set /p DATA_DIR=DATA_REPORT\data folder: 

set DATA_DIR=%DATA_DIR:"=%

if not exist "%DATA_DIR%" (
    echo.
    echo ERROR: This folder does not exist:
    echo %DATA_DIR%
    echo.
    pause
    exit /b 1
)

echo.
set /p OUTPUT_PREFIX=Output prefix, for example Run_G0001: 

if "%OUTPUT_PREFIX%"=="" (
    echo.
    echo ERROR: Output prefix cannot be empty.
    echo Example: Run_G0001
    echo.
    pause
    exit /b 1
)

echo.
set /p START_YEAR=Start year [default 2020]: 
set /p END_YEAR=End year [default 2050]: 

if "%START_YEAR%"=="" set START_YEAR=2020
if "%END_YEAR%"=="" set END_YEAR=2050

echo.
echo Running extraction...
echo Input: %DATA_DIR%
echo Output: %OUTPUT_DIR%
echo Prefix: %OUTPUT_PREFIX%
echo Years: %START_YEAR% to %END_YEAR%
echo.

python "%SCRIPT%" ^
  --data-dir "%DATA_DIR%" ^
  --output-dir "%OUTPUT_DIR%" ^
  --output-prefix "%OUTPUT_PREFIX%" ^
  --start-year %START_YEAR% ^
  --end-year %END_YEAR%

if errorlevel 1 (
    echo.
    echo ERROR: Extraction finished with errors. Check the messages above.
    echo.
    pause
    exit /b 1
)

echo.
echo Finished.
echo Output files are here:
echo %OUTPUT_DIR%
echo.
pause
