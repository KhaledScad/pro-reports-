# Proactive Reporting — Change log

Changes on branch `quarters-licence-search-access`, compared with `main` (`9857180`, "Expand activity logging and add admin log export"). Last updated 2026-10-27.

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
- **Shared licence numbers.** A search shows every Frame ID that shares a licence number with the one searched. For each quarter, the figures come from the one Frame ID that reported the most of them, so two submissions are never mixed within a quarter.
  - **Any quarter's licence counts.** One establishment can carry two licence numbers across quarters (e.g. Frame ID 219980307493: IN-1000657 in Q4 2025 and Q2 2026, CN-1002091 in Q1 2026). Frame IDs are linked when they share any of them, not only the newest one. Values without a digit or shorter than 5 characters ("0", "-") never link Frame IDs. The Record card shows every licence number the Frame ID has had.
  - **This is how the database's Q3 figures reach current Frame IDs.** The QES view files establishments under their other Frame IDs (e.g. 56644 for 219980307493). On today's data, 583 of the 630 current QES Frame IDs now take their Q3 2026 figures from the database.
  - **Users see their establishment's other Frame IDs.** Someone who may open a Frame ID now also sees the quarters of the Frame IDs that share its licence, even when those aren't assigned to them (before, they only saw their own, so the database's figures never reached them). They still cannot open those other Frame IDs on their own.
- **Search lists show the current Frame IDs only.** The Frame ID and License Number lists, and the "(N Frame IDs available to you)" count, offer only the current distribution: 630 QES and 191 FDI today for an admin or supervisor (the survey's auditor list instead, once uploaded). Older Frame IDs a database import files under the current quarter are left out. Licence grouping still uses every Frame ID, and someone allowed to open an older Frame ID can still type it in.
- **Remembered search.** Coming back to a survey page from the menu, Home or the other survey reopens the last Frame ID searched. QES and FDI share Frame IDs.
- **Record card.** Shows the licence number, taken from whichever quarter's block holds it, and the supervisor's name from the account list.
- **Auditor, enumerator and supervisor on the Record card.** These are the people assigned to the Frame ID searched, in the current quarter. Before, when a search combined several Frame IDs with the same licence (most often for an admin, who can open every Frame ID), the card could show another Frame ID's auditor. Each person's name and username now come from the same quarter, so an earlier quarter's auditor or enumerator never replaces the current one. An earlier quarter's name shows only when the Frame ID has nobody in that role in the current quarter.

## 4. Access and roles

- **Who sees what.**
  - **ADMIN:** every Frame ID.
  - **SUPERVISOR authorization:** every Frame ID too, like an admin. Their Home page lists all Frame IDs assigned to anyone in the current quarter, under "All assigned Frame IDs". The admin pages (uploads, users, logs, database) stay ADMIN-only.
  - **Everyone else:** only the Frame IDs they are assigned to in the current quarter, in any assignment column: auditor, enumerator or supervisor, by username or by name. A Frame ID handed to someone else since an earlier quarter is no longer theirs; its earlier quarters still show to the new owner.
  - To let every signed-in user open every Frame ID, set `ALL_USERS_SEE_ALL_RECORDS = True` in `app/auth.py`. It is currently **off**.
- **Exact username matching.** Similar usernames (e.g. `aalshamsi` / `mzaalshamsi`) no longer see each other's records.
- **Auditor list for the current project.** An admin can upload each survey's auditor list (Admin → Auditor assignments): one row per auditor with their Frame IDs separated by commas (`Concat_Frame ID`, `Auditor Name`, `Auditor_User_Name`). While a survey has a list:
  - a normal user (USER) opens exactly the Frame IDs listed under their username or name, with their full history, whoever the data names in this or an earlier quarter;
  - a normal user not in the list opens none in that survey (this includes enumerators, since the list holds auditors);
  - admins and supervisors still see every Frame ID; a controller also gets the listed Frame IDs of their team.
  - Without a list, the data's current-quarter assignment columns decide, as before.
- **Home page lists the current Frame IDs only.** "Frame IDs assigned in Quarter 3 - 2026" shows the current distribution: 630 QES and 191 FDI for an admin or supervisor today, not 1,260 and 384. The extra ones were older Frame IDs that the database import files under Q3 2026. With an auditor list, it shows the list's Frame IDs.
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
- **Auditor assignments — current project:** upload, download (same layout, to edit and re-upload) or remove each survey's auditor list. The card shows the list's size, file and time, and what doesn't line up: Frame IDs in the list but not in the data, data Frame IDs nobody in the list has, auditors without an account, and Frame IDs listed under more than one auditor.
- **Action logs**
  - The table shows the **last 7 days** (change `LOG_DAYS_SHOWN` in `app/app.py`).
  - New **Name** column next to Username. It shows "-" when the username has no account, e.g. a failed sign-in.
  - Times are shown in **UAE time**.
  - **Download full log (.xlsx)** still contains the whole history, with the same Name column and UAE time.
- **Database connection (JDBC).** Survey data can be loaded straight from a database instead of an Excel upload.
  - Supports **Oracle** (`jdbc:oracle:thin:…`, via python-oracledb) and **SQL Server** (`jdbc:sqlserver://…`, via pyodbc and the installed ODBC driver). No Java runtime is needed.
  - One SELECT query per dataset (QES, FDI, AES, FIS), with **Test**, **Preview** and **Import**.
  - **Quarter of the query columns** (set with the connection; starts at the latest ended quarter, Q3 2026 today). A QES / FDI query whose columns have no quarter prefix — a database view returns `LIC_NO`, `MAIN_REVENUE`, … — is read as that quarter: every column except the Frame ID gets its prefix (`Q3_2026_LIC_NO`, …), like the newest block of the Excel file. "Each row's own YEAR / QUARTER" keeps the one-row-per-quarter reading. Queries whose columns already carry prefixes, and the annual AES / FIS, are read as before.
  - **Excel for the earlier quarters, the database for the current one.** Each dataset keeps two parts: the Excel upload's records and the database import's. An import replaces only the previous import's records. While the dataset has a query, an Excel upload replaces only the Excel file's records. For a quarter both hold, the database's values count; the Excel file's still fill the columns the query does not return, such as the usernames and supervisors that decide who opens a Frame ID. The Admin page shows how many records came from the database, and "Loaded now" lists both parts.
  - Access follows the current quarter's assignment columns as before. If the current quarter comes only from a query without username / supervisor columns, auditors and enumerators are still matched by name (names learned from the Excel quarters), but supervisors are not: keep the current quarter's assignment columns in the Excel upload, or add them to the query.
  - The QES pages also read the database view's names: `MAIN_REVENUE` (Total_Revenue), `T_4_TOTA` (Total_Expenses), `T_4_COST_OF_OPERATIONAL` (Operations_Cost), `EMPLYEE_COMPENSATION` (Employee_Salary), `T_4_RAW_MATERIAL`, `T_4_BANKING_CHARGES`, `INTERMEDIATE_CONSUMPTION` (IC). The FDI pages read `CURR_AS_Q3_TOTAL`, `NCURR_AS_Q4_TOTAL`, `CURR_LIAB_Q5_TOTAL`, `LNG_LIAB_Q6_TOTAL`, `EQ_Q7_TOTAL`, `EQ_Q7_7_2`, and show the form items A–C (Q3, Q4) and A–B (Q5, Q6). Search by License Number now also finds `LIC_NO`. An import also stores these columns under the Excel names, and the database's columns take the Excel file's spelling (Oracle returns `COMMENT_REVENUE` for `… AS Comment_Revenue`), so the two parts line up.
  - **Test query** lists which columns are read from the database's own names and which columns the pages read but the query does not return; **Import now** lists the missing ones too. It also says when the rows' own YEAR / QUARTER name another quarter, and when a Frame ID is on more than one row.
  - Optional auto-refresh every N minutes. A refresh whose data has not changed rewrites nothing.
  - Every import, manual or automatic, reports how many records changed since the last import: e.g. "QES: 630 records imported (12 updated, 0 new, 0 removed)", shown on the Admin page. The activity log entry also lists the updated / new / removed Frame IDs (up to 30 of each).
  - **The current quarter comes from the database only.** For the quarter the database holds (Q3 2026), the Excel upload's figures are no longer used; a Frame ID the database doesn't return shows no figures for that quarter. Earlier quarters come from Excel. The Excel quarter's assignment columns and licence number are still used, so access and Frame ID grouping are unchanged.
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
- **47 current QES Frame IDs have no match in the database view** (as of 2026-10-27): none of their licence numbers is on any of the view's rows, so they show only the Excel file's figures for Q3 2026.
- **9 accounts assigned in Q3 2026 are suspended:** `aalali`, `aasalkatheeri`, `assalameri`, `ayalali`, `emalkhulaifi`, `fhbali`, `haalnasseri`, `hralhelali`, `saalromaithi`. They cannot sign in until reactivated.
