"""
Admin-only downloadable XLSX variable reference for QES and FISQ.

Unlike `guide.html` (deliberately worded for end users, with no raw
spreadsheet column names), this file is FOR the people preparing the
upload spreadsheets — it lists every technical column name the app looks
for, the label auditors see for it in the tool, a description, and
whether it feeds the Calculation page, the Comparison page, or both.

Every technical column listed here may also be supplied as a
period-prefixed "wide" column (`Q1_2026_Total_Revenue` instead of plain
`Total_Revenue`) — the app tries the period-prefixed form first, then
falls back to the plain column name. That convention applies uniformly
and is noted once in the sheet header rather than on every row.
"""
from __future__ import annotations

import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .calculations_qes import DETAIL_GROUPS as QES_DETAIL_GROUPS
from .calculations_fisq import COMPONENT_GROUPS as FISQ_COMPONENT_GROUPS, EQUITY_FORM_DETAIL as FISQ_EQUITY_FORM_DETAIL

HEADER_FILL = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SECTION_FILL = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
SECTION_FONT = Font(bold=True, color="991B1B")
WRAP = Alignment(wrap_text=True, vertical="top")

COLUMNS = [
    "Section", "Technical column name(s)", "Label shown to auditors (EN)",
    "Label shown to auditors (AR)", "Description (EN)", "الوصف (AR)", "Used in",
]
COLUMN_WIDTHS = [22, 46, 26, 26, 58, 58, 16]

# Identification / who's-who fields shared by QES and FISQ alike (from
# calc_common.record_meta() and excel_utils.get_frame_id()).
IDENTIFICATION_ROWS = [
    ("Frame_ID (or FRAME, ID, ESTABLISHMENTID, ESTABLISHMENTNUMBER — any Frame-ID-looking header)",
     "Frame ID", "رقم الإطار",
     "Uniquely identifies the establishment. Used to look up every quarterly submission for it, and to search on the Calculation/Comparison pages.",
     "يحدد المنشأة بشكل فريد. يُستخدم لاسترجاع كل استمارة ربعية لها، وللبحث في صفحتي الاحتساب والمقارنة.",
     "Both"),
    ("ESTABLISHMENT_NAME", "Establishment", "المنشأة",
     "Establishment's display name.", "الاسم المعروض للمنشأة.", "Both"),
    ("ISIC4", "ISIC4", "ISIC4",
     "Four-digit ISIC classification code. The first two digits drive automatic business-activity classification (Trade 45–47, Industry 10–33, Construction 41–43, Transport 49–53, else Other). Also feeds the ISIC classification-stability check on the Comparison page.",
     "رمز التصنيف ISIC4 المكوّن من أربعة أرقام. أول رقمين يحددان تصنيف النشاط تلقائياً (تجارة 45–47، صناعة 10–33، إنشاءات 41–43، نقل 49–53، وإلا أخرى). يغذي أيضاً فحص استقرار تصنيف ISIC4 في صفحة المقارنة.",
     "Both"),
    ("ISIC4_DESC_EN, ISIC4_DESC_AR", "ISIC4 description", "وصف ISIC4",
     "Text description of the ISIC4 code. Also scanned for activity keywords as a fallback classification signal.",
     "وصف نصي لرمز ISIC4. يُفحص أيضاً بحثاً عن كلمات مفتاحية للنشاط كإشارة تصنيف احتياطية.",
     "Both"),
    ("ACTIVITY_TYPE_EN, ACTIVITY_TYPE_AR", "Activity type", "نوع النشاط",
     "Free-text business activity as reported. Scanned for activity keywords (trade/industry/construction/transport, EN+AR) as a fallback classification signal alongside the ISIC4 code range.",
     "النشاط التجاري كنص حر كما ورد. يُفحص بحثاً عن كلمات مفتاحية للنشاط (تجارة/صناعة/إنشاءات/نقل، عربي وإنجليزي) كإشارة تصنيف احتياطية إلى جانب نطاق رمز ISIC4.",
     "Both"),
    ("AUDITOR_NAME_EN, AUDITOR_NAME_AR, Auditor_User_Name (or Auditor_Username)",
     "Auditor", "المدقق",
     "Name / username of the auditor of record. Shown on Calculation/Comparison pages; also drives auditor-scoped row visibility (an Auditor role only sees rows matching their own username).",
     "اسم/اسم مستخدم المدقق المسؤول. يظهر في صفحتي الاحتساب والمقارنة؛ ويحدد أيضاً رؤية الصفوف حسب دور المدقق (يرى المدقق الصفوف الخاصة باسم مستخدمه فقط).",
     "Both"),
    ("ENUMERATOR_NAME_EN, ENUMERATOR_NAME_AR, _ENUMERATOR_User_name (or ENUMERATOR_User_name / Enumerator_User_Name)",
     "Enumerator", "الباحث الميداني",
     "Name / username of the enumerator of record. Same role as Auditor above, for the Enumerator role's row visibility.",
     "اسم/اسم مستخدم الباحث الميداني المسؤول. نفس دور المدقق أعلاه، لتحديد رؤية الصفوف لدور الباحث الميداني.",
     "Both"),
    ("Supervisor_USER_Name (or Supervisor_Username)", "Supervisor username", "اسم مستخدم المشرف",
     "Drives row visibility for the Supervisor role.", "يحدد رؤية الصفوف لدور المشرف.", "Both"),
    ("YEAR", "Year", "السنة", "Reported year, shown for reference alongside the record.",
     "السنة المُبلغ عنها، تُعرض للمرجعية إلى جانب السجل.", "Both"),
]

QES_ROWS = [
    ("Revenue, expenses & cost of goods", "Total_Revenue", "Revenue", "الإيرادات",
     "Total revenue for the quarter. Feeds Gross output, Profit, Profit %.",
     "إجمالي الإيرادات خلال الربع. يغذي الإنتاج الإجمالي، الربح، ونسبة الربح.", "Both"),
    ("Revenue, expenses & cost of goods", "Total_Expenses", "Expenditure", "المصروفات",
     "Total expenses for the quarter. For Trade/Industry establishments this already includes the Raw material figure as one of its line items — the Profit formula backs Raw material back out to avoid double-counting it. Feeds Profit, Intermediate consumption.",
     "إجمالي المصروفات خلال الربع. بالنسبة للتجارة والصناعة، يتضمن هذا الرقم أصلاً قيمة المواد الخام كأحد بنوده — وتطرح معادلة الربح قيمة المواد الخام مرة أخرى لتفادي احتسابها مرتين. يغذي الربح، والاستهلاك الوسيط.", "Both"),
    ("Revenue, expenses & cost of goods", "Operations_Cost", "Operations cost", "تكلفة التشغيل",
     "Feeds Intermediate consumption (Operations cost + Expenditure, plus Raw material for Industry).",
     "يغذي الاستهلاك الوسيط (تكلفة التشغيل + المصروفات، بالإضافة إلى المواد الخام للصناعة).", "Both"),
    ("Revenue, expenses & cost of goods", "RAW_Material", "Raw material", "المواد الخام",
     "For Industry establishments: raw material used in production. For Trade establishments: net purchases for resale — the same column serves both meanings depending on the establishment's classified activity. Feeds Gross output, Intermediate consumption (Industry only), and Profit for both Trade & Industry. Not used for Construction/Transport/Other.",
     "للصناعة: المواد الخام المستخدمة في الإنتاج. للتجارة: صافي المشتريات لغرض إعادة البيع — نفس العمود يخدم كلا المعنيين حسب تصنيف نشاط المنشأة. يغذي الإنتاج الإجمالي، والاستهلاك الوسيط (للصناعة فقط)، والربح لكل من التجارة والصناعة. لا يُستخدم للإنشاءات/النقل/أخرى.", "Both"),
    ("Revenue, expenses & cost of goods",
     "Total_Finished_Goods_Begining_Of_period (or Finished_Goods_Begining_Of_period / Finished_Goods_Beginning_Of_period / Inventory_Start)",
     "Inventory start", "المخزون، أول المدة",
     "Combined finished-goods + work-in-progress inventory value at the start of the quarter (Trade & Industry only). Feeds Gross output; checked against the previous quarter's Inventory end in the Inventory continuity check.",
     "قيمة مخزون البضاعة التامة وغير التامة الصنع مجتمعة في بداية الربع (للتجارة والصناعة فقط). يغذي الإنتاج الإجمالي؛ ويُقارَن بمخزون آخر المدة للربع السابق في فحص استمرارية المخزون.", "Both"),
    ("Revenue, expenses & cost of goods",
     "Total_Finished_Goods_End_of_period (or Finished_Goods_End_of_period / Inventory_End)",
     "Inventory end", "المخزون، آخر المدة",
     "Same as Inventory start, at the close of the quarter.",
     "نفس مخزون أول المدة، في نهاية الربع.", "Both"),
    ("Revenue, expenses & cost of goods", "IC", "Intermediate consumption (as reported)", "الاستهلاك الوسيط (كما ورد)",
     "If present and non-zero, used directly instead of being computed from Operations cost + Expenditure (+ Raw material for Industry).",
     "إذا وُجدت وكانت غير صفرية، تُستخدم مباشرة بدلاً من احتسابها من تكلفة التشغيل + المصروفات (+ المواد الخام للصناعة).", "Both"),
    ("Employment", "Employee_Salary", "Compensation", "التعويضات",
     "Total employee compensation for the quarter. Feeds Average salary (÷ Employees ÷ 3).",
     "إجمالي تعويضات الموظفين خلال الربع. يغذي متوسط الراتب (÷ عدد الموظفين ÷ 3).", "Both"),
    ("Employment", "TOTAL_EMPLOYEES (or Employee_Number)", "Employees", "عدد الموظفين",
     "Headcount for the quarter. Feeds Average salary; also shown in the Employee comparison table with the period-over-period change (count and %).",
     "عدد الموظفين خلال الربع. يغذي متوسط الراتب؛ ويظهر أيضاً في جدول مقارنة الموظفين مع التغيّر بين الفترات (بالعدد والنسبة).", "Both"),
    ("Comments", "Comment_Revenue", "Revenue comment", "ملاحظة الإيرادات",
     "Free-text auditor note shown alongside the revenue figure.", "ملاحظة نصية من المدقق تظهر بجانب رقم الإيرادات.", "Calculation"),
    ("Comments", "Comment_Expense", "Expense comment", "ملاحظة المصروفات",
     "Free-text auditor note shown alongside the expenditure figure.", "ملاحظة نصية من المدقق تظهر بجانب رقم المصروفات.", "Calculation"),
    ("Comments", "Comment_Inventory", "Inventory comment", "ملاحظة المخزون",
     "Free-text auditor note shown alongside the inventory figures.", "ملاحظة نصية من المدقق تظهر بجانب أرقام المخزون.", "Calculation"),
]

_FISQ_ASSET_LIAB_NOTE_EN = "Same End_of_period-preferred, Begining_Of_period-fallback priority as the rest of FISQ's begin/end columns."
_FISQ_ASSET_LIAB_NOTE_AR = "نفس أولوية التفضيل (نهاية المدة أولاً، ثم بداية المدة كبديل) المتبعة في بقية أعمدة FISQ للبداية/النهاية."

FISQ_ROWS = [
    ("Assets", "Current_Assets_Begining_Of_period (or _Beginning_Of_period / Current_Assets_Start)",
     "Current assets — beginning of period", "الأصول المتداولة — بداية المدة",
     "Assets expected to be used or converted to cash within the year, as of the start of the quarter. Feeds the beginning-of-period asset equation.",
     "الأصول المتوقع استخدامها أو تحويلها إلى نقد خلال السنة، كما في بداية الربع. يغذي معادلة الأصول لبداية المدة.", "Both"),
    ("Assets", "Current_Assets_End_of_period (or Current_Assets_End; also read as plain Current_Assets)",
     "Current assets — end of period", "الأصول المتداولة — نهاية المدة",
     "Same as above, at the close of the quarter — this is also the tool's primary \"current assets\" figure for the quarter as a whole.",
     "نفس ما سبق، في نهاية الربع — وهو أيضاً الرقم الأساسي الذي تعتمده الأداة \"للأصول المتداولة\" لكامل الربع.", "Both"),
    ("Assets", "Non_Current_Assets_Begining_Of_period (or _Beginning_Of_period / Non_Current_Assets_Start)",
     "Non-current assets — beginning of period", "الأصول غير المتداولة — بداية المدة",
     "Longer-term assets (property, equipment, ...), as of the start of the quarter.",
     "الأصول طويلة الأجل (ممتلكات، معدات، ...)، كما في بداية الربع.", "Both"),
    ("Assets", "Non_Current_Assets_End_of_period (or Non_Current_Assets_End; also read as plain Non_Current_Assets)",
     "Non-current assets — end of period", "الأصول غير المتداولة — نهاية المدة",
     "Same as above, at the close of the quarter — also the primary quarter-level figure.",
     "نفس ما سبق، في نهاية الربع — وهو أيضاً الرقم الأساسي على مستوى الربع.", "Both"),
    ("Assets", "Total_Assets_Begining_Of_period (or _Beginning_Of_period)",
     "Total assets (as reported) — beginning of period", "إجمالي الأصول (كما ورد) — بداية المدة",
     "The file's own reported total, shown alongside the calculated total (Current + Non-current assets) for comparison — not used in the calculation itself.",
     "الإجمالي المُبلغ عنه في الملف، يُعرض بجانب الإجمالي المحتسب (المتداولة + غير المتداولة) للمقارنة — لا يُستخدم في الاحتساب نفسه.", "Comparison"),
    ("Assets", "Total_Assets_End_of_period (also read as plain Total_Assets)",
     "Total assets (as reported) — end of period", "إجمالي الأصول (كما ورد) — نهاية المدة",
     "Same as above, end of quarter.", "نفس ما سبق، نهاية الربع.", "Comparison"),
    ("Liabilities", "Current_Liabilities_Begining_Of_period (or _Beginning_Of_period / Current_Liabilities_Start)",
     "Current liabilities — beginning of period", "المطلوبات المتداولة — بداية المدة",
     "Obligations due within the year, as of the start of the quarter.",
     "الالتزامات المستحقة خلال السنة، كما في بداية الربع.", "Both"),
    ("Liabilities", "Current_Liabilities_End_of_period (or Current_Liabilities_End; also read as plain Current_Liabilities)",
     "Current liabilities — end of period", "المطلوبات المتداولة — نهاية المدة",
     "Same as above, end of quarter — also the primary quarter-level figure.",
     "نفس ما سبق، نهاية الربع — وهو أيضاً الرقم الأساسي على مستوى الربع.", "Both"),
    ("Liabilities",
     "Non_Current_Liabilities_Begining_Of_period (or _Beginning_Of_period / Total_of_Long_Term_Liabilities_Begining_Of_period / Long_Term_Liabilities_Begining_Of_period / Non_Current_Liabilities_Start)",
     "Non-current liabilities — beginning of period", "المطلوبات غير المتداولة — بداية المدة",
     "Longer-term obligations, as of the start of the quarter.", "الالتزامات طويلة الأجل، كما في بداية الربع.", "Both"),
    ("Liabilities",
     "Non_Current_Liabilities_End_of_period (or Non_Current_Liabilities_End / Total_of_Long_Term_Liabilities_End_of_period / Long_Term_Liabilities_End_of_period; also read as plain Non_Current_Liabilities / Total_of_Long_Term_Liabilities / Long_Term_Liabilities / NonCurrent_Liabilities)",
     "Non-current liabilities — end of period", "المطلوبات غير المتداولة — نهاية المدة",
     "Same as above, end of quarter — also the primary quarter-level figure.",
     "نفس ما سبق، نهاية الربع — وهو أيضاً الرقم الأساسي على مستوى الربع.", "Both"),
    ("Liabilities", "Total_Liabilities_Begining_Of_period (or _Beginning_Of_period)",
     "Total liabilities (as reported) — beginning of period", "إجمالي المطلوبات (كما ورد) — بداية المدة",
     "The file's own reported total, shown alongside the calculated total for comparison — not used in the calculation itself.",
     "الإجمالي المُبلغ عنه في الملف، يُعرض بجانب الإجمالي المحتسب للمقارنة — لا يُستخدم في الاحتساب نفسه.", "Comparison"),
    ("Liabilities", "Total_Liabilities_End_of_period (also read as plain Total_Liabilities)",
     "Total liabilities (as reported) — end of period", "إجمالي المطلوبات (كما ورد) — نهاية المدة",
     "Same as above, end of quarter.", "نفس ما سبق، نهاية الربع.", "Comparison"),
    ("Equity", "Equity_In_Form_Begining_Of_period (or _Beginning_Of_period / Equity_Form_Begining_Of_period / Form_Equity_Begining_Of_period)",
     "Equity in form — beginning of period", "حقوق الملكية في الاستمارة — بداية المدة",
     "The equity figure the establishment itself reported, as of the start of the quarter — the benchmark the calculated equity is checked against.",
     "رقم حقوق الملكية الذي أبلغت عنه المنشأة نفسها، كما في بداية الربع — المرجع الذي تُقارَن به حقوق الملكية المحتسبة.", "Both"),
    ("Equity", "Equity_In_Form_End_of_period (or Equity_Form_End_of_period / Form_Equity_End_of_period; also read as plain Equity_In_Form / Equity_Form / Form_Equity)",
     "Equity in form — end of period", "حقوق الملكية في الاستمارة — نهاية المدة",
     "Same as above, end of quarter — also the primary quarter-level figure used in the main asset-equation check.",
     "نفس ما سبق، نهاية الربع — وهو أيضاً الرقم الأساسي المستخدم في فحص معادلة الأصول الرئيسي.", "Both"),
    ("Equity", "Equity_Calculated (or Calculated_Equity)", "Calculated equity (as reported)", "حقوق الملكية المحتسبة (كما وردت)",
     "The file's own calculated-equity figure, if present — shown for reference only, not used in the tool's own calculation.",
     "رقم حقوق الملكية المحتسب الخاص بالملف نفسه، إن وُجد — يُعرض للمرجعية فقط، ولا يُستخدم في احتساب الأداة نفسها.", "Comparison"),
    ("Equity details (7.1–7.6)", "Equity_7_1", "Equity detail 7.1", "تفصيل حقوق الملكية 7.1",
     "First of six equity line items that together should sum to Equity in form.",
     "أول ستة بنود تفصيلية لحقوق الملكية، يجب أن يكون مجموعها مساوياً لحقوق الملكية في الاستمارة.", "Comparison"),
    ("Equity details (7.1–7.6)", "Equity_7_2", "Equity detail 7.2", "تفصيل حقوق الملكية 7.2",
     "Also checked, together with 7.3 alone, against Equity in form (a common way this figure is expected to reconcile).",
     "يُتحقق منه أيضاً، مع 7.3 وحدهما، مقابل حقوق الملكية في الاستمارة (طريقة شائعة يُتوقع أن يتوافق بها هذا الرقم).", "Comparison"),
    ("Equity details (7.1–7.6)", "Equity_7_3", "Equity detail 7.3", "تفصيل حقوق الملكية 7.3",
     "See Equity detail 7.2.", "انظر تفصيل حقوق الملكية 7.2.", "Comparison"),
    ("Equity details (7.1–7.6)", "Equity_7_4", "Equity detail 7.4", "تفصيل حقوق الملكية 7.4",
     "Part of the six-item equity total.", "جزء من إجمالي بنود حقوق الملكية الستة.", "Comparison"),
    ("Equity details (7.1–7.6)", "Equity_7_5", "Equity detail 7.5", "تفصيل حقوق الملكية 7.5",
     "Part of the six-item equity total.", "جزء من إجمالي بنود حقوق الملكية الستة.", "Comparison"),
    ("Equity details (7.1–7.6)", "Equity_7_6", "Equity detail 7.6", "تفصيل حقوق الملكية 7.6",
     "Part of the six-item equity total.", "جزء من إجمالي بنود حقوق الملكية الستة.", "Comparison"),
]

ANNUAL_ISIC_ROWS = [
    ("Frame_ID (or FRAME, ID, ESTABLISHMENTID, ESTABLISHMENTNUMBER)", "Frame ID", "رقم الإطار",
     "Same Frame ID as the QES/FISQ quarterly files — used to match this annual record to the establishment.",
     "نفس رقم الإطار المستخدم في ملفات QES/FISQ الربعية — يُستخدم لمطابقة هذا السجل السنوي بالمنشأة.", "Comparison"),
    ("{YYYY}_ISIC4 — one column per year, e.g. 2023_ISIC4, 2024_ISIC4, 2025_ISIC4",
     "(year column, e.g. \"2025\")", "(عمود السنة، مثل \"2025\")",
     "The ISIC4 code reported for that year. The tool reads every {YYYY}_ISIC4 column present, sorts them by year, and shows the last 3 on the Comparison page — flagging red the moment the value changes from one year to the next.",
     "رمز ISIC4 المُبلغ عنه لتلك السنة. تقرأ الأداة كل عمود {YYYY}_ISIC4 موجود، وترتبها حسب السنة، وتعرض آخر 3 سنوات في صفحة المقارنة — وتُشير بالأحمر لحظة تغيّر القيمة من سنة لأخرى.", "Comparison"),
]


def _write_note(ws: Worksheet, text_en: str, text_ar: str, row: int) -> int:
    ws.cell(row=row, column=1, value=text_en).font = Font(italic=True, color="6B7280")
    ws.cell(row=row, column=1).alignment = WRAP
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(COLUMNS))
    row += 1
    ws.cell(row=row, column=1, value=text_ar).font = Font(italic=True, color="6B7280")
    ws.cell(row=row, column=1).alignment = WRAP
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(COLUMNS))
    return row + 2


def _write_sheet(ws: Worksheet, rows: list[tuple], note_en: str, note_ar: str) -> None:
    ws.sheet_view.rightToLeft = False
    for i, width in enumerate(COLUMN_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    r = _write_note(ws, note_en, note_ar, 1)

    header_row = r
    for c, title in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=header_row, column=c, value=title)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = WRAP
    ws.freeze_panes = f"A{header_row + 1}"
    r += 1

    current_section = None
    for section, tech, label_en, label_ar, desc_en, desc_ar, used_in in rows:
        if section != current_section:
            cell = ws.cell(row=r, column=1, value=section)
            for c in range(1, len(COLUMNS) + 1):
                ws.cell(row=r, column=c).fill = SECTION_FILL
            cell.font = SECTION_FONT
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(COLUMNS))
            current_section = section
            r += 1
        values = ["", tech, label_en, label_ar, desc_en, desc_ar, used_in]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = WRAP
        r += 1


_PERIOD_NOTE_EN = ("Every technical column below may also be supplied as a period-prefixed \"wide\" column "
                    "(e.g. Q1_2026_Total_Revenue instead of plain Total_Revenue) — one column per quarter. "
                    "The tool tries the period-prefixed form first, then falls back to the plain column name.")
_PERIOD_NOTE_AR = ("يمكن توفير أي عمود تقني أدناه أيضاً كعمود \"عريض\" مسبوق بالفترة "
                    "(مثل Q1_2026_Total_Revenue بدلاً من Total_Revenue) — عمود واحد لكل ربع. "
                    "تحاول الأداة أولاً الشكل المسبوق بالفترة، ثم تعود إلى اسم العمود العادي كبديل.")

_ANNUAL_NOTE_EN = ("This sheet is not part of the QES/FISQ upload file — it's the separate annual AES / FIS "
                    "reference file (Admin panel → AES / FIS upload cards) used only for the ISIC "
                    "classification-stability cross-check on the Comparison page.")
_ANNUAL_NOTE_AR = ("هذه الورقة ليست جزءاً من ملف رفع QES/FISQ — بل هي ملف AES / FIS السنوي المرجعي المنفصل "
                    "(لوحة الإدارة ← بطاقتا رفع AES / FIS) المستخدم فقط لمطابقة استقرار تصنيف ISIC4 "
                    "في صفحة المقارنة.")


_DETAIL_DESC_EN = "Optional survey-form line item, shown for reference on the Calculation/Comparison pages — read-only, does not feed any formula above."
_DETAIL_DESC_AR = "بند اختياري من استمارة المسح، يُعرض للمرجعية في صفحتي الاحتساب والمقارنة — للعرض فقط، ولا يدخل في أي معادلة أعلاه."

_DETAIL_NOTE_EN = ("Every column below is OPTIONAL and read-only: if present, it appears in a collapsible "
                    "\"Survey form details\" section on the Calculation/Comparison pages, purely for reference. "
                    "None of these feed Total_Revenue / Total_Expenses / Current_Assets / etc. above — those keep "
                    "being read exactly as before, regardless of whether these columns are supplied.")
_DETAIL_NOTE_AR = ("كل عمود أدناه اختياري وللعرض فقط: إذا وُجد، يظهر في قسم قابل للطي بعنوان \"تفاصيل استمارة المسح\" "
                    "في صفحتي الاحتساب والمقارنة، للمرجعية فقط. لا يدخل أي منها في احتساب Total_Revenue / "
                    "Total_Expenses / Current_Assets / إلخ أعلاه — فتلك تبقى كما هي بغض النظر عن توفر هذه الأعمدة.")


def _qes_detail_rows() -> list[tuple]:
    rows = []
    for title_en, title_ar, items in QES_DETAIL_GROUPS:
        for label_en, label_ar, codes in items:
            tech = " + ".join(codes) if len(codes) > 1 else codes[0]
            rows.append((title_en, tech, label_en, label_ar, _DETAIL_DESC_EN, _DETAIL_DESC_AR, "Both"))
    return rows


def _fisq_detail_rows() -> list[tuple]:
    rows = []
    for g in FISQ_COMPONENT_GROUPS:
        for label_en, label_ar, codes in g["components"]:
            tech = " (or) ".join(codes) if len(codes) > 1 else codes[0]
            rows.append((g["title_en"], tech, label_en, label_ar, _DETAIL_DESC_EN, _DETAIL_DESC_AR, "Both"))
    for label_en, label_ar, codes in FISQ_EQUITY_FORM_DETAIL:
        tech = " (or) ".join(codes) if len(codes) > 1 else codes[0]
        rows.append(("Equity details (form Q7)", tech, label_en, label_ar, _DETAIL_DESC_EN, _DETAIL_DESC_AR, "Both"))
    return rows


def build_workbook(survey_type: str) -> bytes:
    """survey_type: 'QES' or 'FISQ'. Returns the .xlsx file bytes."""
    survey_type = survey_type.upper()
    wb = Workbook()
    ws = wb.active
    ws.title = survey_type

    rows = IDENTIFICATION_ROWS_FOR(survey_type) + (QES_ROWS if survey_type == "QES" else FISQ_ROWS)
    _write_sheet(ws, rows, _PERIOD_NOTE_EN, _PERIOD_NOTE_AR)

    ws2 = wb.create_sheet("AES-FIS annual ISIC")
    annual_rows = [("Frame ID & annual ISIC4", *row) for row in ANNUAL_ISIC_ROWS]
    _write_sheet(ws2, annual_rows, _ANNUAL_NOTE_EN, _ANNUAL_NOTE_AR)

    ws3 = wb.create_sheet("Survey form details (optional)")
    detail_rows = _qes_detail_rows() if survey_type == "QES" else _fisq_detail_rows()
    _write_sheet(ws3, detail_rows, _DETAIL_NOTE_EN, _DETAIL_NOTE_AR)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def IDENTIFICATION_ROWS_FOR(survey_type: str) -> list[tuple]:
    return [("Identification", *row) for row in IDENTIFICATION_ROWS]
