from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI, Request, Form, UploadFile, File, Query
from fastapi.responses import RedirectResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from . import database as db
from . import auth
from . import excel_utils
from . import tempstore
from . import i18n
from .calc_common import fmt, pct, period_label, get_periods, record_meta
from .calculations_qes import calc_qes
from .calculations_fisq import calc_fisq

BASE_DIR = Path(__file__).resolve().parent
db.init_db()

app = FastAPI(title="Proactive Reporting")
app.add_middleware(SessionMiddleware, secret_key="proactive-reporting-local-secret-key-change-me")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")

templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
templates.env.filters["fmt"] = fmt
templates.env.filters["pct"] = pct
templates.env.filters["period_label"] = period_label


def trend_cls(prev_val, curr_val) -> str:
    """green if this quarter is up vs the previous one, red if down."""
    if prev_val is None:
        return ""
    if curr_val > prev_val:
        return "trend-up"
    if curr_val < prev_val:
        return "trend-down"
    return "trend-same"


templates.env.globals["trend_cls"] = trend_cls
LAST_N_PERIODS = 4  # current quarter + previous 3

SURVEY_LABELS = {
    "QES": "Quarter Economic Survey",
    "FISQ": "Foreign Investment Quarter Survey",
}
CALCULATED_TYPES = {"QES", "FISQ"}

# Variables selectable in the compare-page trend charts: (data key, i18n label key).
CHART_VARIABLES = {
    "QES": [
        ("total_revenue", "revenue"), ("total_expenses", "expenditure"),
        ("gross_output", "gross_output"), ("intermediate_consumption", "intermediate_consumption"),
        ("added_value", "added_value"), ("profit", "profit"),
        ("employees", "employees"), ("average_salary", "average_salary"),
        ("employee_salary", "compensation"),
        ("begin_inventory", "inventory_start"), ("end_inventory", "inventory_end"),
        ("profit_pct", "profit_pct"), ("av_pct", "av_pct"), ("ic_pct", "ic_pct"),
    ],
    "FISQ": [
        ("current_assets", "current_assets"), ("non_current_assets", "non_current_assets"),
        ("total_assets", "total_assets"), ("current_liabilities", "current_liabilities"),
        ("non_current_liabilities", "non_current_liabilities"), ("total_liabilities", "total_liabilities"),
        ("calculated_equity", "calculated_equity"), ("equity_in_form", "equity_in_form"),
        ("equation_difference", "difference"),
    ],
}
CHART_DEFAULT_VARIABLES = {
    "QES": ["added_value", "total_revenue"],
    "FISQ": ["calculated_equity", "total_assets"],
}


# --------------------------------------------------------------- helpers ---

def get_lang(request: Request) -> str:
    return request.session.get("lang", "en")


def render(request: Request, template: str, **ctx):
    lang = get_lang(request)
    ctx["user"] = auth.current_user(request)
    ctx["is_admin"] = auth.is_admin(ctx["user"])
    ctx["survey_labels"] = SURVEY_LABELS
    ctx["lang"] = lang
    ctx["t"] = lambda key: i18n.t(key, lang)
    return templates.TemplateResponse(request, template, ctx)


@app.get("/lang/{lang}")
def set_lang(request: Request, lang: str, next: str = "/dashboard"):
    request.session["lang"] = "ar" if lang == "ar" else "en"
    return RedirectResponse(next, status_code=303)


def require_login(request: Request) -> dict:
    user = auth.current_user(request)
    if not user:
        return None
    return user


def merged_row(records: list[dict]) -> dict:
    merged: dict = {}
    for r in records:
        merged.update(r["data"])
    return merged


# ------------------------------------------------------------------ auth ---

@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    if auth.current_user(request):
        return RedirectResponse("/dashboard")
    return RedirectResponse("/login")


@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    if auth.current_user(request):
        return RedirectResponse("/dashboard")
    return render(request, "login.html", error=None)


@app.post("/login", response_class=HTMLResponse)
def login_submit(request: Request, username: str = Form(...), password: str = Form(...)):
    user = auth.login(username.strip(), password.strip())
    if not user:
        db.add_log(username, "UNKNOWN", "LOGIN_FAILED", "Invalid username or password")
        return render(request, "login.html", error=i18n.t("invalid_login", get_lang(request)))
    auth.set_session_user(request, user)
    db.add_log(user["username"], user["authorization"], "LOGIN_SUCCESS", "")
    return RedirectResponse("/dashboard", status_code=303)


@app.get("/logout")
def logout(request: Request):
    user = auth.current_user(request)
    if user:
        db.add_log(user["username"], user["authorization"], "LOGOUT", "")
    request.session.clear()
    return RedirectResponse("/login", status_code=303)


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    user = require_login(request)
    if not user:
        return RedirectResponse("/login")
    access = [
        {
            "type": t,
            "label": SURVEY_LABELS[t],
            "frame_ids": sorted({r["frame_id"] for r in auth.visible_records(user, t)}),
        }
        for t in db.SURVEY_TYPES
    ]
    return render(request, "dashboard.html", access=access)


# ----------------------------------------------------------------- admin ---

@app.get("/admin", response_class=HTMLResponse)
def admin_page(request: Request):
    user = require_login(request)
    if not user or not auth.is_admin(user):
        return RedirectResponse("/dashboard")
    counts = {t: db.record_count(t) for t in db.SURVEY_TYPES}
    return render(request, "admin.html", counts=counts, users=db.list_users(),
                  logs=db.list_logs(200), message=request.query_params.get("message"),
                  error=request.query_params.get("error"))


UPLOAD_KINDS = {*db.SURVEY_TYPES, "ACCOUNTS"}


@app.post("/admin/upload/{kind}/preview")
async def upload_preview(request: Request, kind: str, file: UploadFile = File(...)):
    """Step 1: stash the file and let the admin pick which sheet to import."""
    user = require_login(request)
    if not user or not auth.is_admin(user):
        return RedirectResponse("/dashboard")
    kind = kind.upper()
    if kind not in UPLOAD_KINDS:
        return RedirectResponse("/admin?error=Unknown+upload+type", status_code=303)
    content = await file.read()
    try:
        names = excel_utils.sheet_names(content)
        token = tempstore.save(content, file.filename)
        return render(request, "choose_sheet.html", kind=kind, token=token,
                      filename=file.filename, sheet_names=names)
    except Exception as exc:  # noqa: BLE001
        return RedirectResponse(f"/admin?error=Could+not+read+file:+{exc}", status_code=303)


@app.post("/admin/upload/{kind}/confirm")
def upload_confirm(request: Request, kind: str, token: str = Form(...), sheet: str = Form(default="__ALL__")):
    """Step 2: import the chosen sheet (or all sheets) from the stashed file."""
    user = require_login(request)
    if not user or not auth.is_admin(user):
        return RedirectResponse("/dashboard")
    kind = kind.upper()
    loaded = tempstore.load(token)
    if not loaded:
        return RedirectResponse("/admin?error=Upload+expired,+please+try+again", status_code=303)
    content, filename = loaded
    try:
        if kind == "ACCOUNTS":
            users = excel_utils.parse_account_sheet(content, sheet)
            if not users or not any(u["authorization"] == "ADMIN" for u in users):
                raise ValueError("Include at least one ADMIN account in the Account sheet.")
            db.replace_users(users)
            db.add_log(user["username"], user["authorization"], "AUTH_FILE_LOADED", f"{filename} / sheet: {sheet}")
            msg = f"Account+database+loaded+({len(users)}+users)"
        else:
            rows = excel_utils.rows_for_survey(content, sheet)
            if not rows:
                raise ValueError("No records were imported. Make sure the sheet has a Frame ID / FRAME_ID column.")
            count = db.replace_records(kind, rows, filename)
            db.add_log(user["username"], user["authorization"], f"{kind}_DATA_UPLOADED", f"{filename} / sheet: {sheet} / records: {count}")
            msg = f"{kind}:+{count}+records+loaded"
        return RedirectResponse(f"/admin?message={msg}", status_code=303)
    except Exception as exc:  # noqa: BLE001
        return RedirectResponse(f"/admin?error={exc}", status_code=303)
    finally:
        tempstore.delete(token)


# ---------------------------------------------------------------- survey ---

@app.get("/survey/{survey_type}", response_class=HTMLResponse)
def survey_search(request: Request, survey_type: str, frame_id: str = "",
                   mode: str = "compare", quarter: str = "", quarters: list[str] = Query(default=[])):
    user = require_login(request)
    if not user:
        return RedirectResponse("/login")
    survey_type = survey_type.upper()
    if survey_type not in db.SURVEY_TYPES:
        return RedirectResponse("/dashboard")
    lang = get_lang(request)

    frame_options = sorted({r["frame_id"] for r in auth.visible_records(user, survey_type)})

    result = None
    error = None
    row = None
    periods = []
    meta = None
    raw_fields = None
    frame_id = frame_id.strip()

    if frame_id:
        records = [r for r in auth.visible_records(user, survey_type) if r["frame_id"].strip().lower() == frame_id.lower()]
        if not records:
            error = i18n.t("not_found_or_access", lang)
        else:
            row = merged_row(records)
            periods = get_periods(row)
            if survey_type in CALCULATED_TYPES:
                if mode == "compare":
                    # Default: current quarter + previous 3 (last 4 available periods).
                    selected = [p for p in quarters if p in periods] or periods[-LAST_N_PERIODS:]
                else:
                    selected = [quarter] if quarter in periods else (periods[-1:] if periods else [])
                calc_fn = calc_qes if survey_type == "QES" else calc_fisq
                extra = {"lang": lang} if survey_type == "QES" else {}
                result = calc_fn(row, selected, mode, **extra) if periods else None
                if not periods:
                    error = "No quarter/year period was detected in the uploaded columns for this record."
                else:
                    meta = record_meta(row, result["first"]["period"] if result and result.get("first") else periods[-1], lang)
            else:
                raw_fields = {k: v for k, v in row.items() if not str(k).startswith("__")}
            db.add_log(user["username"], user["authorization"], "SURVEY_FRAME_VIEW", f"{survey_type} / {frame_id}")

    template = f"survey_{survey_type.lower()}.html" if survey_type in CALCULATED_TYPES else "survey_generic.html"
    return render(request, template, survey_type=survey_type, frame_id=frame_id, mode=mode,
                  quarter=quarter, quarters=quarters, frame_options=frame_options, periods=periods,
                  result=result, error=error, row=row, meta=meta, raw_fields=raw_fields)


# ------------------------------------------------------------- comparison ---

@app.get("/compare/{survey_type}", response_class=HTMLResponse)
def compare_view(request: Request, survey_type: str, frame_id: str = "", periods: list[str] = Query(default=[])):
    user = require_login(request)
    if not user:
        return RedirectResponse("/login")
    survey_type = survey_type.upper()
    if survey_type not in CALCULATED_TYPES:
        return RedirectResponse("/dashboard")
    lang = get_lang(request)

    frame_options = sorted({r["frame_id"] for r in auth.visible_records(user, survey_type)})
    frame_id = frame_id.strip()
    error = None
    row = None
    all_periods = []
    calcs = []
    meta = None

    if frame_id:
        records = [r for r in auth.visible_records(user, survey_type) if r["frame_id"].strip().lower() == frame_id.lower()]
        if not records:
            error = i18n.t("not_found_or_access", lang)
        else:
            row = merged_row(records)
            all_periods = get_periods(row)
            # Default: current quarter + previous 3 (last 4 available periods).
            chosen = [p for p in periods if p in all_periods] or all_periods[-LAST_N_PERIODS:]
            calc_period_fn = (lambda p: calc_qes(row, [p], "single", lang)["first"]) if survey_type == "QES" else \
                              (lambda p: calc_fisq(row, [p], "single")["first"])
            calcs = [calc_period_fn(p) for p in chosen]
            if calcs:
                meta = record_meta(row, calcs[0]["period"], lang)
            db.add_log(user["username"], user["authorization"], "PROJECT_COMPARISON_VIEW", f"{survey_type} / {frame_id}")

    # Inventory continuity check (QES only): start should equal previous end.
    inventory_rows = []
    if survey_type == "QES" and calcs:
        for i, c in enumerate(calcs):
            prev_end = calcs[i - 1]["end_inventory"] if i > 0 else None
            ok = i == 0 or abs((c["begin_inventory"] or 0) - (prev_end or 0)) < 0.01
            inventory_rows.append({"calc": c, "ok": ok, "is_base": i == 0})

    employee_rows = []
    if survey_type == "QES" and calcs:
        for i, c in enumerate(calcs):
            prev = calcs[i - 1]["employees"] if i > 0 else None
            change = (c["employees"] - prev) if i > 0 else 0
            change_pct = (change / abs(prev) * 100) if (i > 0 and prev) else 0
            employee_rows.append({"calc": c, "is_base": i == 0, "change": change, "change_pct": change_pct})

    chart = None
    if calcs:
        var_defs = CHART_VARIABLES[survey_type]
        chart = {
            "labels": [period_label(c["period"], lang) for c in calcs],
            "variables": [{"key": k, "label": i18n.t(label_key, lang)} for k, label_key in var_defs],
            "series": {k: [c.get(k) or 0 for c in calcs] for k, _ in var_defs},
            "default": CHART_DEFAULT_VARIABLES[survey_type],
        }

    template = f"compare_{survey_type.lower()}.html"
    return render(request, template, survey_type=survey_type, frame_id=frame_id, frame_options=frame_options,
                  all_periods=all_periods, selected_periods=[c["period"] for c in calcs], calcs=calcs,
                  inventory_rows=inventory_rows, employee_rows=employee_rows, error=error, row=row, meta=meta,
                  chart=chart)


# ------------------------------------------------------------------ main ---

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.app:app", host="127.0.0.1", port=8000, reload=True)
