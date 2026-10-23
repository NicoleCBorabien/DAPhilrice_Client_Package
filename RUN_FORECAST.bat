@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title PhilRice Specific Tag Forecasting

cls
echo ============================================================
echo              PHILRICE SPECIFIC TAG FORECASTING
echo ============================================================
echo.
echo 1. Fill in 02_Source_Data\Forecast_Input_Template.xlsx
echo 2. Save and CLOSE Excel
echo 3. This window will create your forecast automatically
echo.

REM Prefer Python 3.12 because the scientific libraries used by the model
REM have stable Windows wheels for this version.
set "PY_CMD="
py -3.12 --version >nul 2>&1
if %errorlevel%==0 set "PY_CMD=py -3.12"

if not defined PY_CMD (
    python --version >nul 2>&1
    if %errorlevel%==0 set "PY_CMD=python"
)

if not defined PY_CMD (
    echo ERROR: Python was not found on this computer.
    echo.
    echo Please ask your administrator to install Python 3.12 from python.org,
    echo then double-click RUN_FORECAST.bat again.
    echo.
    pause
    exit /b 1
)

if not exist ".client_env\Scripts\python.exe" (
    echo First-time setup: preparing the forecasting system...
    echo This only needs to happen once.
    %PY_CMD% -m venv .client_env
    if errorlevel 1 goto :setup_error
)

call ".client_env\Scripts\activate.bat"

if not exist ".client_env\.forecast_ready" (
    echo Installing forecasting components for first use...
    python -m pip install --upgrade pip >nul
    python -m pip install --only-binary=:all: -r requirements.txt
    if errorlevel 1 goto :setup_error
    type nul > ".client_env\.forecast_ready"
)

echo.
echo Running forecast...
echo.
python "03_Model_and_Code\philrice_forecasting.py"
if errorlevel 1 goto :forecast_error

echo.
echo ============================================================
echo                 FORECAST COMPLETED SUCCESSFULLY
echo ============================================================
echo.
echo Opening your result now...
set "RESULT=01_Results\PhilRice_Specific_Tag_Forecast_Results.xlsx"
if exist "%RESULT%" (
    start "" "%RESULT%"
) else (
    start "" "01_Results"
)
echo.
echo You may close this window.
pause
exit /b 0

:setup_error
echo.
echo ============================================================
echo SETUP COULD NOT BE COMPLETED
echo ============================================================
echo.
echo Please contact your project administrator.
echo If this is a company-managed computer, Windows security may be

echo blocking one of the Python forecasting components.
echo.
pause
exit /b 1

:forecast_error
echo.
echo ============================================================
echo FORECAST COULD NOT BE COMPLETED
echo ============================================================
echo.
echo Check that:
echo   - the completed input workbook is saved in 02_Source_Data
echo   - Excel is closed before running the forecast
echo   - the data follows the template format
echo.
echo If the problem continues, contact your project administrator.
echo.
pause
exit /b 1
