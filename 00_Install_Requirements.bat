@echo off
setlocal

echo ==========================================
echo LUTO Economics Exporters
echo Install / Check Required Python Packages
echo ==========================================
echo.

set TOOL_DIR=%~dp0
set HELPER=%TOOL_DIR%install_requirements_helper.py

if not exist "%HELPER%" (
    echo ERROR: Could not find:
    echo %HELPER%
    echo.
    echo Make sure install_requirements_helper.py is in the same folder as this BAT file.
    echo.
    pause
    exit /b 1
)

python "%HELPER%"

if errorlevel 1 (
    echo.
    echo ERROR: Package installation/check failed.
    echo.
    pause
    exit /b 1
)

echo.
echo Required packages are ready.
echo.
pause
