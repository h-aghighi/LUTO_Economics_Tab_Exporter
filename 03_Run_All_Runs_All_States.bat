@echo off
setlocal

echo ==========================================
echo LUTO Economics Tab Exporter
echo Multiple runs - all states
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

echo Paste the folder containing the Run_* folders.
echo.
echo Example:
echo C:\LUTO_Extractor
echo.
set /p REPORTS_BASE_DIR=Reports base folder: 

set REPORTS_BASE_DIR=%REPORTS_BASE_DIR:"=%

if not exist "%REPORTS_BASE_DIR%" (
    echo.
    echo ERROR: This folder does not exist:
    echo %REPORTS_BASE_DIR%
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
echo Running extraction for all Run_* folders...
echo Input: %REPORTS_BASE_DIR%
echo Output: %OUTPUT_DIR%
echo Years: %START_YEAR% to %END_YEAR%
echo.

python "%SCRIPT%" ^
  --reports-base-dir "%REPORTS_BASE_DIR%" ^
  --output-dir "%OUTPUT_DIR%" ^
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
