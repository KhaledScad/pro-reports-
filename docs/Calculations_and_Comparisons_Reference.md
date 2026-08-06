# Proactive Reporting — Calculation & Comparison Reference

This document describes every calculated field in the Proactive Reporting app: the formula, which uploaded spreadsheet column(s) feed it, and how the Comparison pages use it. It reflects the code exactly as implemented in `app/calculations_qes.py`, `app/calculations_fisq.py`, `app/calc_common.py`, and `app/app.py`.

---

## 0. How a column is matched to a variable

Uploaded sheets can use either a plain column name (`Total_Revenue`) or a period-prefixed "wide" column (`Q1_2026_Total_Revenue`), one per quarter. For a given period, the app looks for an exact `{period}_{suffix}` match first, then a fuzzy period-prefixed match, then finally a bare (no-period) match — so a single value can serve as a fallback across quarters if it isn't provided per-quarter.

Column names are compared after normalizing: lower-cased, with spaces/underscores/dashes/dots/slashes/parentheses stripped, and Arabic letter variants folded together (`إ/أ/آ → ا`, `ة → ه`, `ى → ي`). Numeric values are cleaned the same way everywhere: thousands commas are stripped, a value wrapped in parentheses is read as negative (`(500)` → `-500`), and any other non-numeric character is discarded.

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

Trade and Industry both have a **cost of goods** (per SCAD's QES + FDI validation rules), so their Gross Output is computed from inventory movement rather than being plain Revenue. The same uploaded `RAW_Material` column is reused for two different real-world concepts depending on the activity: for Industry it is the raw material *used in production*; for Trade it is the *net purchases for resale*. Construction, Transport, and every other activity (services, real estate, ...) have no cost-of-goods concept in the survey, so Gross Output is simply Revenue.

| Activity | Gross Output (GO) | Intermediate Consumption (IC) |
|---|---|---|
| **Trade** | Revenue + Inventory(end) − Inventory(begin) − Raw material *(net purchases for resale — this is the "gross profit margin")* | IC as reported in the file, else Operations cost + Expenses |
| **Industry** | Revenue + Inventory(end) − Inventory(begin) − Raw material *(raw material used in production — this is the "gross profit margin")* | IC as reported in the file, else Operations cost + Expenses + Raw material |
| **Construction / Transport / Other** | Revenue | IC as reported in the file, else Operations cost + Expenses |

"IC as reported in the file" means: if the uploaded `IC` column has a non-zero value for that record/period, it is used as-is instead of being computed from its components.

### 1.3 Shared formulas (every activity)

Profit/Loss is also activity-dependent. For Trade and Industry, the reported Total Expenses figure already includes the raw material / net-purchases line item as one of its components (it's one of the itemized expense lines that sum to the reported total) — so when Profit is computed from the gross margin (which has already netted out that same raw-material figure via Gross Output above), the raw-material amount is backed out of Total Expenses first, to avoid subtracting it twice.

| Field | Formula |
|---|---|
| Added Value | Gross Output − Intermediate Consumption |
| Profit — Trade / Industry | Gross Output − (Expenses − Raw material) |
| Profit — Construction / Transport / Other | Revenue − Expenses |
| Average Salary | Employee compensation ÷ Employees ÷ 3 *(converts the quarterly compensation figure to a monthly one)* |
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

### 1.5 Comparison page — `/compare/qes`

Default period selection is the current quarter plus the previous three (the last four periods found in the record), or whichever quarters the user explicitly checks.

1. **Financial comparison** — Period, Revenue, Expenditure, Profit, Added Value, Intermediate Consumption, Profit %, Added Value %, IC %. Each cell (after the first period) is colored green if it rose versus the prior selected period, red if it fell.
2. **Inventory continuity check** — for every period after the first, verifies that this period's start-of-period inventory equals the previous period's end-of-period inventory (within 0.01). Flags "Mismatch" otherwise.
3. **Employee comparison** — Employees, Compensation, Average salary, Expenditure, Revenue, plus the period-over-period change in employee count (both count and %).
4. **Trend charts** — see §3 below.

---

## 2. FISQ — Foreign Investment Quarter Survey (asset equation)

### 2.1 Formulas

| Field | Formula |
|---|---|
| Total Assets | Current assets + Non-current assets |
| Total Liabilities | Current liabilities + Non-current liabilities |
| Calculated Equity | Total Assets − Total Liabilities |
| Equation Difference | Calculated Equity − Equity in form *(the figure as reported)*; \|difference\| < 0.01 ⇒ "Balanced" |
| Equity details total | Equity 7.1 + 7.2 + 7.3 + 7.4 + 7.5 + 7.6 |
| 7.2 + 7.3 check | (Equity 7.2 + Equity 7.3) − Equity in form; \|difference\| < 0.01 ⇒ "Correct", else "Flag / review" |
| Overall status | "Balanced" only if **both** the equation difference and the 7.2+7.3 check are within 0.01; otherwise "Needs review" |

### 2.2 FISQ source columns

| Column(s) matched | Feeds |
|---|---|
| `Current_Assets`, `Non_Current_Assets` | Total assets components |
| `Total_Assets` | As reported in the file — shown alongside the calculated total for comparison, not used in the calculation itself |
| `Current_Liabilities`; `Non_Current_Liabilities` / `Total_of_Long_Term_Liabilities` / `Long_Term_Liabilities` / `NonCurrent_Liabilities` | Total liabilities components |
| `Total_Liabilities` | As reported in the file — shown alongside the calculated total |
| `Equity_In_Form`, `Equity_Form`, `Form_Equity` | Reported equity — the benchmark both checks compare against |
| `Equity_Calculated`, `Calculated_Equity` | The file's own calculated-equity figure, if present |
| `Equity_7_1` … `Equity_7_6` | The six equity detail line items |
| `ESTABLISHMENT_NAME` | Record label |

### 2.3 Comparison page — `/compare/fisq`

Same default period selection as QES (current quarter + previous three, or the periods explicitly checked).

1. **Asset equation comparison** — Period, Current assets, Non-current assets, Total assets, Current liabilities, Non-current liabilities, Total liabilities, Calculated equity, Equity in form, Difference. Same green/red trend coloring as QES.
2. **Equity details (7.1–7.6)** — each line item, their total, the 7.2+7.3 subtotal, equity in form, the difference, and status.
3. **Validation** — Assets: file vs. calculated; Liabilities: file vs. calculated; difference; Balanced / Needs review.
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

**QES (14 variables):** Revenue, Expenditure, Gross output, Intermediate consumption, Added value, Profit, Employees, Average salary, Compensation, Inventory start, Inventory end, Profit %, Added Value %, IC %.

**FISQ (9 variables):** Current assets, Non-current assets, Total assets, Current liabilities, Non-current liabilities, Total liabilities, Calculated equity, Equity in form, Difference.
