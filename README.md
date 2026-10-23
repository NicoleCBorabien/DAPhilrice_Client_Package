# PhilRice Forecast EXE Project

This project builds a single Windows executable:

`PhilRice_Forecast.exe`

## Final client workflow

The client does not need VS Code, Python, Terminal, Jupyter, or GitHub.

1. Double-click `PhilRice_Forecast.exe`.
2. A browser window opens automatically.
3. Choose the PhilRice Excel workbook.
4. Click **Run Forecast**.
5. Wait for the completion message.
6. Click **Download Excel Results**.

The forecasting backend runs locally on the user's computer.

## Build the EXE for free with GitHub Actions

1. Upload this project's CONTENTS to a GitHub repository.
2. Open the repository in GitHub.
3. Click **Actions**.
4. Select **Build Windows EXE**.
5. Click **Run workflow**.
6. Wait for the workflow to finish.
7. Open the completed workflow run.
8. Scroll to **Artifacts**.
9. Download `PhilRice_Forecast_Windows`.
10. Unzip the artifact.
11. The final client file is `PhilRice_Forecast.exe`.

## Development test

If you are developing locally:

```bash
pip install -r requirements-build.txt
python src/app.py
```

## Important Windows note

This EXE is unsigned because the build process is free. Windows SmartScreen or a company Application Control policy may warn about or block unsigned executables. A tightly managed corporate computer may require IT allow-listing.
