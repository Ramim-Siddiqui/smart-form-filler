# Smart Form Filler

Record yourself entering one record on any website, then let the tool repeat it for every row of an Excel sheet. Built for repetitive data entry into web portals that have no import feature or API.

The tool drives a real Chrome window, so it works with sites behind a login: you sign in by hand once, and the session is reused. No passwords are ever recorded.

## How it works

1. **Open the browser** from the tool and log in to the target website yourself.
2. **Upload an Excel file** and choose which columns to use.
3. **Record** one record by hand. The recorder captures each click, typed value, dropdown choice and date field, with a stable selector for each element.
4. **Map** each recorded value to an Excel column. Columns are suggested automatically by matching field labels to column names, and date fields can be reformatted (for example `dd/mm/yyyy`).
5. **Define a success check** so every row is verified after it is submitted.
6. **Run** a single test row, a few rows, or the whole sheet. Every row gets a `Status`, `Message` and timestamp in a results file. To resume an interrupted run, upload the results file again and finished rows are skipped.

Recorded setups ("recipes") are saved as JSON, so a website only has to be recorded once.

### Several websites at once

Each launch starts the next free copy of the tool (up to nine), each with its own browser window, login profile, and ports (UI on `8500+N`, Chrome DevTools on `9221+N`). Copies never see each other's tabs.

## Tech stack

- **Python 3.9+**
- **Streamlit** for the step-by-step interface (`app.py`)
- **Playwright** connected over the Chrome DevTools Protocol to a user-launched Chrome, for recording and replay (`engine.py`)
- **pandas / openpyxl** for reading the sheet and writing results
- `launcher.py` picks a free slot so several copies can run side by side

## Project structure

```
smart-form-filler/
├── app.py                         # Streamlit UI: upload, record, map, verify, run
├── engine.py                      # Chrome control, recorder (JS injection) and runner
├── launcher.py                    # starts the next free copy on its own ports and profile
├── requirements.txt
├── Setup (first time).bat         # Windows: installs Python if needed, creates .venv, installs deps
├── Start Smart Form Filler.bat    # Windows: starts a copy of the tool
├── recipes/                       # saved website setups (git-ignored)
└── results/                       # run results per Excel file (git-ignored)
```

## Getting started (Windows)

1. Double-click **`Setup (first time).bat`**. It installs Python 3.12 with winget if it is missing, creates a virtual environment and installs the dependencies. If it installs Python, run it once more when prompted.
2. Double-click **`Start Smart Form Filler.bat`** and keep its console window open. The interface opens in your browser.
3. Follow the numbered steps on screen.

Google Chrome or Microsoft Edge must be installed.

### Manual setup (any OS)

```bash
python -m venv .venv
.venv/Scripts/python -m pip install -r requirements.txt   # or .venv/bin/python on macOS/Linux
.venv/Scripts/python launcher.py
```

## Privacy

Browser profiles (`chrome_profile*/`), recipes and results are git-ignored because they can contain logins and personal data. Password fields are detected during recording and never stored.

## Limitations

- Built and tested on Windows; the Chrome path detection and batch files are Windows-specific.
- Sites that change element IDs on every load, or use CAPTCHAs, need manual handling.
- No automated tests yet.

## License

Copyright (c) 2026 Ramim Ali Siddiqui. All rights reserved. The code is published for viewing only; using, copying or modifying it requires written permission. See [LICENSE](LICENSE).
