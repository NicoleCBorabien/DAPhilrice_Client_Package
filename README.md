# PhilRice Client Package

A simple Windows client package for forecasting monthly **% of Total Workload by Specific Tag**.

## Client workflow

1. Open `02_Source_Data/Forecast_Input_Template.xlsx`.
2. Enter the employee's monthly Specific Tag workload percentages.
3. Save and close Excel.
4. Double-click `RUN_FORECAST.bat`.
5. Open the generated workbook in `01_Results`.

## What the model does

The forecasting script validates the input data, compares ARIMA and Holt using rolling validation and RMSE, selects the better-performing model, forecasts the next 12 months, normalizes monthly forecast shares to 100%, and creates an Excel results workbook with forecast and validation charts.

## Input history requirements

- `5-Year`: 60 consecutive months
- `1-Year`: 12 consecutive months
- `6-Month`: 6 consecutive months

## Main files

- `RUN_FORECAST.bat` — double-click to run the forecast on Windows
- `02_Source_Data/Forecast_Input_Template.xlsx` — client input template
- `03_Model_and_Code/philrice_forecasting.py` — forecasting engine
- `01_Results/` — generated results

## Privacy

Do not commit real employee source data or generated employee forecast results to a public repository.
