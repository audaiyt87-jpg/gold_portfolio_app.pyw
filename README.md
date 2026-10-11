# Gold Portfolio - Windows EXE Build

This package builds `gold_portfolio_app.pyw` (standard-library Tkinter app,
no external dependencies) into a single-file Windows executable `GoldPortfolio.exe`.

Features included: USD/SAR/AED support, Al Rajhi buy/sell spreads, and price alerts.

---

## Option A - Build in the cloud with GitHub Actions (recommended, no local setup)

1. Create a new GitHub repository (or use an existing one).
2. Upload `gold_portfolio_app.pyw` to the repository root.
3. Add the workflow file. Because the `.github` folder is hidden and hard to
   upload from the browser, create it directly on GitHub:
   - Click **Add file -> Create new file**.
   - In the file name box type the full path:
     `.github/workflows/build-exe.yml`
   - Paste the contents of the included `build-exe.yml`.
   - Click **Commit changes**.
4. Go to the **Actions** tab -> select **Build Windows EXE** -> **Run workflow**.
5. Wait for the run to finish (green check), open the run, scroll to
   **Artifacts**, and download **GoldPortfolio-Windows-EXE**.
6. Unzip it to get `GoldPortfolio.exe`.

---

## Option B - Build locally on a Windows PC

1. Install Python 3.11 from https://python.org (check "Add Python to PATH").
2. Double-click `build_exe.bat` (or run the command below in a terminal):
   ```
   python -m pip install --upgrade pyinstaller
   pyinstaller --onefile --windowed --name GoldPortfolio gold_portfolio_app.pyw
   ```
3. The executable will be at `dist\GoldPortfolio.exe`.

---

## Running the EXE

- Double-click `GoldPortfolio.exe`.
- Windows SmartScreen may warn because the app is not code-signed:
  click **More info -> Run anyway**.
