# Proactive Reporting — Change log

Changes on branch `quarters-licence-search-access`, compared with `main` (`9857180`, "Expand activity logging and add admin log export"). Last updated 2026-10-23.

**Before you deploy:** restart the app on the server and read [Upgrade notes](#upgrade-notes). The server loads the Python code and the page templates only when it starts.

---

## 1. Surveys and quarters

- **Current quarters.** QES data covers Q4 2025, Q1 2026, Q2 2026 and Q3 2026. The "current quarter" used for access and the dashboard is the newest quarter in each survey's data (Q3 2026 today).
- **Newest quarter first.** Quarter pickers, results and every comparison table list the newest quarter at the top. Green/red trends, "Base period", the inventory check and the ±30% flag still compare each quarter with the one before it in time. Charts still run oldest → newest, left to right.
- **Merged multi-quarter Excel files.** When several quarterly extracts are merged into one sheet, the later quarters' columns get `Left_` / `Right_` prefixes (`Right_Q2_2026_Total_Revenue`, `Right_Licence_Number`, …). Each record's blocks are now placed in the quarter that block's own YEAR / QUARTER values name. This happens at upload, and also when reading data stored by older versions, so no re-upload is needed.
- **Long layout.** Data with one row per Frame ID per quarter, with YEAR and QUARTER columns (the usual shape of a database view), is turned into the same one-record-per-Frame-ID layout as the Excel files.

## 2. FISQ renamed to FDI

- The Foreign Investment survey is now **FDI — Foreign Direct Investment Survey** (مسح الاستثمار الأجنبي المباشر) throughout the app.
- Pages are `/survey/fdi` and `/compare/fdi`. Old `/survey/fisq`, `/compare/fisq` and `?tab=fisq` links redirect to the new pages.
- On the first start of the new version, the stored FISQ data (`data/fisq.db`) is copied to `data/fdi.db`. The old file is left in place.
- Renamed files: `calculations_fdi.py`, `survey_fdi.html`, `compare_fdi.html`, `sample_data/FDI_*.xlsx`, `Proactive_Reporting_QES_FDI_Variables.xlsx`.

## 3. Search and the Record card

- **Search by Frame ID or License Number.** Both boxes are searchable dropdowns: click to open the list, or type part of a number or the establishment name. Picking one clears the other.
- **Shared licence numbers.** A search shows every Frame ID the user can open that has the same licence number. For each quarter, the figures come from the one Frame ID that reported the most of them, so two submissions are never mixed within a quarter.
- **Remembered search.** Coming back to a survey page from the menu, Home or the other survey reopens the last Frame ID searched. QES and FDI share Frame IDs.
- **Record card.** Shows the licence number, taken from whichever quarter's block holds it, and the supervisor's name from the account list.

## 4. Access and roles

- **Who sees what.**
  - **ADMIN:** every Frame ID.
  - **Everyone else:** only the Frame IDs they are assigned to in the current quarter, in any assignment column: auditor, enumerator or supervisor, by username or by name. A Frame ID handed to someone else since an earlier quarter is no longer theirs; its earlier quarters still show to the new owner.
  - To let every signed-in user open every Frame ID, set `ALL_USERS_SEE_ALL_RECORDS = True` in `app/auth.py`. It is currently **off**.
- **Exact username matching.** Similar usernames (e.g. `aalshamsi` / `mzaalshamsi`) no longer see each other's records.
- **New CONTROLLER role.** An admin places employees under a controller (Admin → Controller teams). The controller then sees exactly the Frame IDs assigned to those employees.
- **Blank usernames, filled-in names.** When a quarter leaves someone's username blank but has their full name, the app recognises them by learning which username goes with that name from other quarters, in either survey.
  - Safeguards: a name must have appeared next to the same username at least twice and never next to another username, and a name the two surveys disagree on is ignored.
  - Example: `mssalrashdi` (MEERA SALEM SUHAIL TAHI ALRASHDI) went from 0 to 37 QES and 5 FDI Frame IDs in Q3 2026. No other account gained or lost access.
- **Immediate effect.** A role change, suspension or removal applies on the user's next click, not their next sign-in.

## 5. Admin page

- **Users**
  - Suspend / reactivate an account (a suspended user cannot sign in), or remove it for good.
  - **Download users (.xlsx):** every account with Username, Authorization, Name, Job and Location. Passwords are never exported.
- **Controller teams:** choose which employees each CONTROLLER oversees.
- **Action logs**
  - The table shows the **last 7 days** (change `LOG_DAYS_SHOWN` in `app/app.py`).
  - New **Name** column next to Username. It shows "-" when the username has no account, e.g. a failed sign-in.
  - Times are shown in **UAE time**.
  - **Download full log (.xlsx)** still contains the whole history, with the same Name column and UAE time.
- **Database connection (JDBC).** Survey data can be loaded straight from a database instead of an Excel upload.
  - Supports **Oracle** (`jdbc:oracle:thin:…`, via python-oracledb) and **SQL Server** (`jdbc:sqlserver://…`, via pyodbc and the installed ODBC driver). No Java runtime is needed.
  - One SELECT query per dataset (QES, FDI, AES, FIS), with **Test**, **Preview** and **Import**.
  - Optional auto-refresh every N minutes. A refresh whose data has not changed rewrites nothing.
  - Queries are read-only: only SELECT / WITH statements run, and nothing is committed.
  - Credentials are stored in the app's local database (`data/`), never in the code or in git.

## 6. QES checks

- **"Gross output" is now labelled "Total production"** (إجمالي الإنتاج).
- **Compensation more than Total production.** When a quarter's total compensation is above zero and more than its Total production, compensation is shown **in red**:
  - **Calculation page:** the compensation figure is red, with "More than total production" under it, and a red warning box sits under the quarter's formulas.
  - **Comparison page:** red in the Employee comparison table (which now also shows Total production and a check column), and the quarter is listed in a red warning at the top.
- **Compensation more than Expenditure** is flagged the same way, since compensation is part of expenditure.
- **Intermediate consumption from the file.** When the file provides IC, that value is used instead of the calculated one.

## 7. FDI checks

- Column names follow the 20-10 variable list (`app/survey_layouts.py`); older names are still accepted.
- **Equity items check:** 7.1 + 7.2 + 7.3 is compared with Total equity in the form, on the calculation and comparison pages.
- Asset equation: Total assets − Total liabilities = Calculated equity, compared with the equity in the form, at the beginning and end of each quarter.

## 8. Guide and documentation

- The in-app **Guide** is rewritten in English and Arabic: getting started, searching, what the colours mean, step-by-step QES and FDI walkthroughs, every check, and a glossary.
- Updated: the admin variable reference download, `Proactive_Reporting_QES_FDI_Variables.xlsx`, and `docs/Calculations_and_Comparisons_Reference.md` / `_AR.md`.

---

## Upgrade notes

1. **Restart the app** on the server after pulling these changes. Until then it runs the old Python code. Templates read from the shared folder may not match that code, so FDI pages can fail and new admin buttons won't work.
2. **FDI data:** the first start copies `fisq.db` to `fdi.db`. Upload FDI files only after that restart; anything uploaded to an old-version server in the meantime goes to `fisq.db` and won't show under FDI.
3. **Database connection (optional):** `requirements.txt` now lists `oracledb` and `pyodbc`. They are only needed to use Admin → Database connection; the app starts without them. Install with `pip install -r requirements.txt`.
4. `data/` (databases, accounts, settings, logs) and `uploads/` are not in git.

## Open data issues (as of 2026-10-23)

These come from the source files, not the app:

- **`nmalaryani`** (NAJLA MOHAMMED ALI HADDAD ALARYANI) is not in the Q3 2026 QES or FDI data at all, so she has no Q3 Frame IDs.
- **30 QES Q3 2026 rows** have an auditor name but no auditor username, and those names never appear with a username in any quarter. Only admins see them until `Auditor_User_Name` is filled in.
- **FDI Q2 2026:** in 147 rows the auditor name does not match the auditor username.
- **FDI quarters:** the FDI file labels its oldest quarter **Q3 2025**, not Q4 2025, and has no YEAR / QUARTER columns to check against.
- **9 accounts assigned in Q3 2026 are suspended:** `aalali`, `aasalkatheeri`, `assalameri`, `ayalali`, `emalkhulaifi`, `fhbali`, `haalnasseri`, `hralhelali`, `saalromaithi`. They cannot sign in until reactivated.
