# PhilRice Forecast — Free Windows Desktop Build

Uses the tested local web app with live progress, 80/20 validation, adjustable horizon and fully labeled Excel chart images. The forecasting engine remains unchanged.

## For the developer (one-time setup)

1. Upload **the contents of this directory** to the root of the GitHub repository (not inside a nested folder). Include the hidden `.github/workflows` directory. GitHub's browser upload may skip empty/hidden directories; Git via VS Code or GitHub Desktop is most reliable.
2. Go to **Actions → Build Windows EXE → Run workflow → Run workflow**.
3. Open the completed run and download `PhilRice_Forecast_Windows` under **Artifacts**.
4. Extract that artifact. It contains `PhilRice_Forecast.exe`.
5. Test the build on a Windows computer where executing it is permitted before distributing it.

## For the user

1. Double-click `PhilRice_Forecast.exe`.
2. A browser opens to the locally hosted forecasting app.
3. Upload the Excel workload file, choose forecast months, click Run Forecast, view live progress, and download the result workbook.
4. The app runs fully offline and keeps the Excel file locally. A browser window will remain usable while the app process is running; close the app's process to stop the server.

## Technical notes

- Windows EXE is made on a Windows GitHub-hosted runner, not on your PC. A basic GitHub Actions allowance may be free, subject to your account's available minutes, terms, and limits.
- Built with `PyInstaller --onefile`-style EXE spec (`EXE` includes analysis data). Python packages, static assets, and HTML are bundled.
- The server binds only to `127.0.0.1` on an available port.
- The app leaves data files in the OS temporary directory, consistent with the tested app; do not use it with confidential data on an untrusted/shared computer before adding cleanup.
- **Smart App Control:** your current Windows setting can block an unsigned EXE even after selecting Unblock. Building it again does not solve that. Do not disable Windows security. Use approved signing/IT-managed approval or test on a machine where application policy allows execution. Code signing often costs money.
- The EXE is not precompiled inside this ZIP: GitHub Actions creates it.
