# Proactive Reporting — Calculation & Comparison Reference

This document describes every calculated field in the Proactive Reporting app: the formula, which uploaded spreadsheet column(s) feed it, and how the Comparison pages use it. It reflects the code exactly as implemented in `app/calculations_qes.py`, `app/calculations_fisq.py`, `app/calc_common.py`, and `app/app.py`.

---

## 0. How a column is matched to a variable

Uploaded sheets can use either a plain column name (`Total_Revenue`) or a period-prefixed "wide" column (`Q1_2026_Total_Revenue`), one per quarter. For a given period, the app looks for an exact `{period}_{suffix}` match first, then a fuzzy period-prefixed match, then finally a bare (no-period) match — so a single plain value can serve as a fallback across quarters if it isn't provided per-quarter. Another quarter's prefixed column never stands in for a missing one (a quarter without an `IC` column computes IC; it does not borrow the next quarter's), an exact name wins over a longer one that merely ends with it, and `Current_…` never matches `Non-Current_…`.

**Merged files (the 20-10 layout, `app/survey_layouts.py`).** The QES/FISQ extracts arrive merged from several quarterly tables, which leaves `Left_Right_` / `Right_` / `Right_Right_` prefixes behind. On upload (`excel_utils.QuarterLayout`, via `normalize_quarter_columns()`), blocks that share one label — QES's `Q2_2026_…` / `Right_Q2_2026_…` / `Right_Right_Q2_2026_…` — are placed record by record in the quarter their own `YEAR` / `QUARTER` values give. They hold each establishment's latest records in turn: mostly Q2 2026 / Q1 2026 / Q4 2025, but e.g. only Q4 2025, in the first block, for an establishment with nothing newer. A block with no `YEAR` / `QUARTER` for a record takes a quarter that record has no figures for, so every record shows the same quarters. A block whose label is its own (`Q3_2026_`) is taken as labelled; its `YEAR` / `QUARTER` cells read 2026 / Q2 throughout the 20-10 file. Any column without a quarter label that sits inside a block — FISQ's `CURR_AS_Q3_D`, `Right_CURR_AS_Q3_D`, a block's `Licence_Number`, the Q1 2026 `Total_of _Non-Current_Assets` — goes with that block. Columns before the first block (`FRAME_ID`, `ESTABLISHMENT_NAME`) stay plain. Data stored before this placement existed is placed the same way when it is read (`database.all_records()`), so it needs no re-upload. The admin's upload message says how the blocks were placed and how many records differ from the usual quarters.

Column names are compared after normalizing: lower-cased, with spaces/underscores/dashes/dots/slashes/parentheses stripped, and Arabic letter variants folded together (`إ/أ/آ → ا`, `ة → ه`, `ى → ي`). Numeric values are cleaned the same way everywhere: thousands commas are stripped, a value wrapped in parentheses is read as negative (`(500)` → `-500`), and any other non-numeric character is discarded.

**One licence number, several Frame IDs.** An establishment can appear under more than one Frame ID with the same licence number — e.g. an old Frame ID holding Q4 2025–Q1 2026 and a new one holding Q2–Q3 2026. (A Frame ID's licence number is its newest quarter's non-empty `Licence_Number` / `LICENSE_NUMBER`.) The Calculation and Comparison pages therefore search by licence: whichever Frame ID is entered (or the License Number), every Frame ID the user can open with that licence is combined into one record (`app.licence_group()` / `licence_row()`). Each quarter is taken whole from the Frame ID that reported the most figures for it — non-zero figures first, then non-empty ones, counting the columns listed in `FIGURES` in `calculations_qes.py` / `calculations_fisq.py`; on a tie the Frame ID entered wins — so two submissions are never mixed within one quarter. The Frame ID entered supplies the columns that belong to no quarter (`FRAME_ID`, `ESTABLISHMENT_NAME`). The Record card lists every Frame ID with the licence and the quarters each one supplied; each quarter's result, and the first table on the Comparison page, names its Frame ID plus any other Frame ID that also reported figures for that quarter. Frame IDs without a licence number are never combined, and a Frame ID the user is not assigned to is never pulled in.

---

## 1. QES — Quarter Economic Survey

### 1.1 Business-activity classification

Every record is classified per quarter, before any formula is applied. The classifier checks the two-digit ISIC4 code first, then falls back to keyword matching across the activity-type and ISIC4-description columns (English and Arabic):

| Activity | ISIC4 range | Keywords — EN | Keywords — AR |
|---|---|---|---|
| Trade | 45–47 | trade | تجار، تجارة، بيع |
| Industry | 10–33 | industry, manufactur, industrial | صناع، تصنيع |
| Construction | 41–43 | construction | تشييد، انشاء، إنشاء، مقاول |
| Transport | 49–53 | transport, storage | نقل، تخزين |
| Other | *(none of the above)* | — | — |

`Other` is a fallback and, unlike before, is now treated the same as `Construction`/`Transport` (no cost of goods) — not the same as `Industry`.

### 1.2 Gross Output & Intermediate Consumption, by activity

Trade and Industry both have a **cost of goods** (per SCAD's QES + FDI validation rules), so their Gross Output is computed from inventory movement rather than being plain Revenue. The same uploaded `RAW_Material` column is reused for two different real-world concepts depending on the activity: for Trade it is the *net purchases for resale* and comes off gross output; for Industry it is the raw material *used in production* and counts in intermediate consumption instead, together with operations cost and bank charges. Construction, Transport, and every other activity (services, real estate, ...) have no cost-of-goods concept in the survey, so Gross Output is simply Revenue.

| Activity | Gross Output (GO) | Intermediate Consumption (IC) |
|---|---|---|
| **Trade** | Revenue + Inventory(end) − Inventory(begin) − Raw material *(net purchases for resale — this is the "gross profit margin")* | IC as reported in the file, else Operations cost + Expenses |
| **Industry** | Revenue + Inventory(end) − Inventory(begin) | IC as reported in the file, else Operations cost + Bank charges (`Banking_Charges`) + Raw material *(in the 20-10 data the file's IC always equals this sum)* |
| **Construction / Transport / Other** | Revenue | IC as reported in the file, else Operations cost + Expenses |

"IC as reported in the file" means: if the uploaded `IC` column has a non-zero value for that record/period, it is used as-is instead of being computed from its components.

### 1.3 Shared formulas (every activity)

Profit/Loss is also activity-dependent. For Trade, the reported Total Expenses figure already includes the net-purchases line item as one of its components — so when Profit is computed from the gross margin (which has already netted out those purchases via Gross Output above), they are backed out of Total Expenses first, to avoid subtracting them twice. Industry's Gross Output no longer nets out raw material, so its Profit is simply Gross Output − Expenses (the same figure as before the 20-10 change).

| Field | Formula |
|---|---|
| Added Value | Gross Output − Intermediate Consumption |
| Profit — Trade | Gross Output − (Expenses − Raw material) |
| Profit — Industry | Gross Output − Expenses |
| Profit — Construction / Transport / Other | Revenue − Expenses |
| Total Compensation | Employee compensation as reported (`Employee_Salary`) — the quarterly total |
| Average Compensation | Total compensation ÷ Employees ÷ 3 *(converts the quarterly compensation figure to a monthly one)* |
| Profit % | Profit ÷ \|Revenue\| × 100 |
| Added Value % | Added Value ÷ \|Gross Output\| × 100 |
| Intermediate Consumption % | Intermediate Consumption ÷ \|Gross Output\| × 100 |
| Validation status | "Positive added value" / "Negative added value" / "Zero added value" |

### 1.4 QES source columns

| Column(s) matched | Feeds |
|---|---|
| `Total_Revenue` | Revenue |
| `Total_Expenses` | Expenses |
| `Operations_Cost` | Operations cost |
| `Employee_Salary` | Employee compensation |
| `RAW_Material` | Raw material |
| `Total_Finished_Goods_Begining_Of_period`, `Finished_Goods_Begining_Of_period`, `Finished_Goods_Beginning_Of_period`, `Inventory_Start` | Inventory, start of period |
| `Total_Finished_Goods_End_of_period`, `Finished_Goods_End_of_period`, `Inventory_End` | Inventory, end of period |
| `IC` | Intermediate consumption, as reported directly (used in place of the computed value when present) |
| `TOTAL_EMPLOYEES` | Employee count |
| `ACTIVITY_TYPE_EN`, `ACTIVITY_TYPE_AR`, `ISIC4`, `ISIC4_DESC_EN`, `ISIC4_DESC_AR` | Business-activity classification |
| `Comment_Revenue`, `Comment_Expense`, `Comment_Inventory` | Auditor free-text comments, carried through unchanged |
| `Licence_Number` | Groups the Frame IDs a search combines (newest quarter's non-empty value; see §0, *One licence number, several Frame IDs*); License Number search and the Record card |
| `Banking_Charges` | Bank charges — part of Industry's intermediate consumption; also listed in the expense details |

### 1.5 Comparison page — `/compare/qes`

Default period selection is the current quarter plus the previous three (the last four periods found in the record), or whichever quarters the user explicitly checks.

1. **Financial comparison** — Period, Revenue, Expenditure, Profit, Added Value, Intermediate Consumption, Profit %, Added Value %, IC %. Each cell (after the first period) is colored green if it rose versus the prior selected period, red if it fell.
2. **ISIC classification stability** — the record's `ISIC4` value across the compared quarters; flags red if it changed since the base (first-selected) period.
3. **Annual ISIC cross-check (AES)** — see §2.4 below.
4. **Inventory continuity check** — for every period after the first, verifies that this period's start-of-period inventory equals the previous period's end-of-period inventory (within 0.01). Flags "Mismatch" otherwise.
5. **Employee comparison** — Employees, Total compensation, Average compensation, Expenditure, Revenue, plus the period-over-period change in employee count (both count and %).
6. **Trend charts** — see §3 below.

---

## 2. FISQ — Foreign Investment Quarter Survey (asset equation)

### 2.1 Formulas

| Field | Formula |
|---|---|
| Total Assets | Current assets + Non-current assets |
| Total Liabilities | Current liabilities + Non-current liabilities |
| Calculated Equity | Total Assets − Total Liabilities |
| Equation Difference | Calculated Equity − Equity in form *(`Total_Equity`, the figure as reported)*; \|difference\| < 0.01 ⇒ "Balanced" |
| 7.2 + 7.3 check | (7.2 Paid-up capital + 7.3 Other equity) − Equity in form; \|difference\| < 0.01 ⇒ "Correct", else "Flag / review" |
| Overall status | "Balanced" only if **both** the equation difference and the 7.2+7.3 check are within 0.01; otherwise "Needs review" |

### 2.2 FISQ source columns

20-10 names first; the older names after them are read only when a 20-10 name is missing, so files in the earlier layout still work.

| Column(s) matched | Feeds |
|---|---|
| `Total_of_Current_Assets` (older: `Current_Assets`), `Total_of _Non-Current_Assets` (older: `Non_Current_Assets`) | Total assets components |
| `Total_of_Current_Liabilities` (older: `Current_Liabilities`), `Total_of_Long-Term_Liabilities` (older: `Non_Current_Liabilities` / `Long_Term_Liabilities` / `NonCurrent_Liabilities`) | Total liabilities components |
| `Total_Equity` (older: `Equity_In_Form`, `Equity_Form`, `Form_Equity`) | Reported equity — the benchmark both checks compare against |
| `7.2_Paid-up _Capital_Stock`, `7.3_Other_Equity` (older: `Equity_7_2`, `Equity_7_3`) | The 7.2 + 7.3 equity check |
| `CURR_AS_Q3_D/E`, `NCURR_AS_Q4_D/E`, `CURR_LIAB_Q5C/D`, `LNG_LIAB_Q6C/D`, each with its `_PREV` | Form items, shown with their previous-period value — display only, never added into the totals |
| `Total_Assets`, `Total_Liabilities` | As reported in older files — shown alongside the calculated totals, not used in the calculation |
| `…_Begining_Of_period` / `…_End_of_period` forms of the totals | Beginning vs end of period table — only in files that carry opening balances (the 20-10 layout doesn't) |
| `ESTABLISHMENT_NAME`, `LICENSE_NUMBER`, `STRATUM`, `VISIT_STATUS_EN/AR`, `ISIC4`, `AUDITOR_NAME`, `ENUMERATOR_NAME` | Record card; `LICENSE_NUMBER` also groups the Frame IDs a search combines (§0) |

### 2.3 Comparison page — `/compare/fisq`

Same default period selection as QES (current quarter + previous three, or the periods explicitly checked).

1. **Asset equation comparison** — Period, Current assets, Non-current assets, Total assets, Current liabilities, Non-current liabilities, Total liabilities, Calculated equity, Equity in form, Difference. Same green/red trend coloring as QES.
2. **Beginning vs end of period** — see `calc_fisq_period_movement()`: the asset equation computed separately for the start and close of each quarter. Shown only when the file carries beginning-of-period totals.
3. **ISIC classification stability** — same as QES §1.5.2, using the FISQ record's own `ISIC4` value.
4. **Annual ISIC cross-check (FIS)** — see §2.4 below.
5. **Equity details (7.2 + 7.3)** — paid-up capital, other equity, their sum, equity in form, the difference, and status.
6. **Validation** — Assets: file vs. calculated; Liabilities: file vs. calculated; difference; Balanced / Needs review.

### 2.4 Annual ISIC cross-check — QES ↔ AES, FISQ ↔ FIS

Independent of the quarterly QES/FISQ data, the admin can upload two small **annual** reference files — one from the AES filing, one from the FIS filing — each containing just a Frame ID column plus one `{YYYY}_ISIC4` wide column per year (e.g. `2023_ISIC4`, `2024_ISIC4`, `2025_ISIC4`). These are *not* full survey types: no calculation engine, no `/survey` or `/compare` page of their own, no dashboard "access" card — they exist solely to feed this cross-check.

`/compare/qes` shows the establishment's last 3 years from the **AES** file only; `/compare/fisq` shows the last 3 years from the **FIS** file only — each survey checks against its own matching annual filing, not both. The tool flags red if the ISIC4 value changed from one year to the next (green if unchanged). When the licence number has several Frame IDs (§0), all of them are read, each year from the first Frame ID that has a value for it. If nothing has been uploaded for that Frame ID in the relevant file, the table shows "No annual data has been uploaded for this Frame ID yet." instead.

Storage: `database.py`'s `ANNUAL_ISIC_SOURCES = ["AES", "FIS"]`, each with its own SQLite file (`aes.db` / `fis.db`), the same generic `records` table schema as QES/FISQ, uploaded via the same `/admin/upload/{kind}/preview` → `/confirm` flow (`kind` = `aes` or `fis`). Extraction: `calc_common.annual_isic_series()`.
4. **Trend charts** — see §3 below. FISQ has no percentage fields, so its charts only ever split by value scale, never by a percentage group.

---

## 3. Trend charts — how variables get grouped

Both comparison pages let you pick up to **8 variables at once** to plot as a bar chart and a line ("pulse") chart across the selected quarters. To keep the bar chart readable, the app automatically splits the selected variables into separate bar charts:

- Any variable whose name ends in `_pct` (Profit %, Added Value %, IC %) is always plotted in its own **percentages** chart — percentages are a different unit from currency or headcount and should never share an axis with them.
- Among the remaining ("amount") variables, each one's peak value across the shown quarters is compared to the largest peak in the selection. Anything under **12%** of that largest peak is moved into its own **small-scale values** chart, so (for example) Employees in the tens doesn't get flattened to an invisible sliver next to Revenue in the hundreds of thousands.
- If everything selected is within that same order of magnitude, they stay together in a single chart — the split only happens when it's actually needed.

The line/pulse chart is not split this way — it always shows every selected variable together.

---

## 4. Full variable list available for charting

**QES (14 variables):** Revenue, Expenditure, Gross output, Intermediate consumption, Added value, Profit, Employees, Total compensation, Average compensation, Inventory start, Inventory end, Profit %, Added Value %, IC %.

**FISQ (9 variables):** Current assets, Non-current assets, Total assets, Current liabilities, Non-current liabilities, Total liabilities, Calculated equity, Equity in form, Difference.
