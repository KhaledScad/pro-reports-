from __future__ import annotations

import json
import threading
from collections import Counter
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, Request, Form, UploadFile, File, Query
from fastapi.responses import RedirectResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from . import database as db
from . import auth
from . import excel_utils
from . import tempstore
from . import i18n
from . import variable_reference
from . import db_import
from .logs_export import build_logs_workbook, build_users_workbook, build_assignments_workbook
from .calc_common import (fmt, pct, period_label, get_periods, key_period, record_meta, assignment_meta, text_val,
                          to_num, clean, latest_text, annual_isic_series)
from .calculations_qes import calc_qes, qes_detail_groups, FIGURES as QES_FIGURES, LICENSE_NUMBER
from .calculations_fdi import calc_fdi, calc_fdi_period_movement, fdi_component_groups, FIGURES as FDI_FIGURES

BASE_DIR = Path(__file__).resolve().parent
db.init_db()

app = FastAPI(title="Proactive Reporting")
app.add_middleware(SessionMiddleware, secret_key=db.get_or_create_session_secret())
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


@app.middleware("http")
async def no_store_pages(request: Request, call_next):
    # Pages are per-user; without this the browser can show a previous
    # user's cached page (Back button / restored tabs) after switching accounts.
    response = await call_next(request)
    if not request.url.path.startswith("/static/"):
        response.headers["Cache-Control"] = "no-store"
    elif "v" in request.query_params:
        # ?v=static_version changes on every deploy, so the browser can keep
        # the file instead of re-checking it (a slow network-share read) on every page.
        response.headers["Cache-Control"] = "public, max-age=31536000, immutable"
    return response


@app.on_event("shutdown")
def _flush_local_db_to_durable_storage():
    db.sync_now()

templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
# Templates change only on a deploy, which restarts the server. Without this
# Jinja re-checks every template file on every page — tens of ms per file
# when the app runs from a network share.
templates.env.auto_reload = False
templates.env.filters["fmt"] = fmt
templates.env.filters["pct"] = pct
templates.env.filters["period_label"] = period_label
templates.env.filters["uae_time"] = db.uae_time
# Cache-bust /static assets whenever any of them changes (i.e. every deploy) so
# browsers can't keep serving a stale style.css/charts.js after an update.
templates.env.globals["static_version"] = str(int(max(p.stat().st_mtime for p in (BASE_DIR / "static").iterdir() if p.is_file())))


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
    "FDI": "Foreign Direct Investment Survey",
}
CALCULATED_TYPES = {"QES", "FDI"}
# Old survey name -> new one (FISQ is now FDI), so old links still work.
RENAMED_TYPES = {old: new for new, old in db.RENAMED_DBS.items()}
USER_ROLES = ("ADMIN", "SUPERVISOR", "CONTROLLER", "USER")  # CONTROLLER: sees only their team's Frame IDs
REPORTED_FIGURES = {"QES": QES_FIGURES, "FDI": FDI_FIGURES}

# Variables selectable in the compare-page trend charts: (data key, i18n label key).
CHART_VARIABLES = {
    "QES": [
        ("total_revenue", "revenue"), ("total_expenses", "expenditure"),
        ("gross_output", "gross_output"), ("intermediate_consumption", "intermediate_consumption"),
        ("added_value", "added_value"), ("profit", "profit"),
        ("employees", "employees"),
        ("employee_salary", "compensation"), ("average_salary", "average_salary"),
        ("begin_inventory", "inventory_start"), ("end_inventory", "inventory_end"),
        ("profit_pct", "profit_pct"), ("av_pct", "av_pct"), ("ic_pct", "ic_pct"),
    ],
    "FDI": [
        ("current_assets", "current_assets"), ("non_current_assets", "non_current_assets"),
        ("total_assets", "total_assets"), ("current_liabilities", "current_liabilities"),
        ("non_current_liabilities", "non_current_liabilities"), ("total_liabilities", "total_liabilities"),
        ("calculated_equity", "calculated_equity"), ("equity_in_form", "equity_in_form"),
        ("equation_difference", "difference"),
    ],
}
CHART_DEFAULT_VARIABLES = {
    "QES": ["added_value", "total_revenue"],
    "FDI": ["calculated_equity", "total_assets"],
}

# ±30% quarter-over-quarter movement flag (FDI only), per the source
# form's own validation rule: (calc dict key, label EN, label AR).
FDI_FLAG_METRICS = [
    ("current_assets", "Current assets", "الأصول المتداولة"),
    ("non_current_assets", "Non-current assets", "الأصول غير المتداولة"),
    ("current_liabilities", "Current liabilities", "المطلوبات المتداولة"),
    ("non_current_liabilities", "Non-current liabilities", "المطلوبات غير المتداولة"),
    ("equity_in_form", "Equity", "حقوق الملكية"),
]


# --------------------------------------------------------------- helpers ---

def get_lang(request: Request) -> str:
    return request.session.get("lang", "en")


def client_ip(request: Request) -> str:
    return request.client.host if request.client else ""


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


def frame_choices(visible: list[dict]) -> list[dict]:
    """Every Frame ID in `visible` (the records this user is assigned to),
    with its establishment name and licence number, for the Frame ID /
    License Number search dropdowns."""
    by_frame: dict[str, list[dict]] = {}
    for r in visible:
        by_frame.setdefault(r["frame_id"], []).append(r)
    out = []
    for fid in sorted(by_frame):
        row = merged_row(by_frame[fid])
        out.append({
            "frame_id": fid,
            "name": latest_text(row, ["ESTABLISHMENT_NAME"]),
            "license": latest_text(row, LICENSE_NUMBER),
            "licenses": frame_licences(row),
        })
    return out


def frame_licences(row: dict) -> set[str]:
    """Every licence number a Frame ID has had, in any quarter (lower-case) —
    an establishment's quarters can carry different ones, e.g. IN-1000657 in
    Q2 and CN-1002091 in Q1, and another Frame ID (a database view's) may use
    either. Values without a digit or shorter than 5 characters ("0", "-")
    are not licence numbers and never link Frame IDs."""
    found = {text_val(row, p, LICENSE_NUMBER).strip().lower() for p in [*get_periods(row), None]}
    return {v for v in found if len(v) >= 5 and any(ch.isdigit() for ch in v)}


def license_choices(frame_options: list[dict]) -> list[dict]:
    """Distinct licence numbers among `frame_options`, each with the Frame
    IDs it belongs to, for the License Number search dropdown."""
    by_license: dict[str, list[dict]] = {}
    for f in frame_options:
        if f["license"]:
            by_license.setdefault(f["license"], []).append(f)
    return [{"license": lic, "frames": frames} for lic, frames in sorted(by_license.items())]


templates.env.globals["license_choices"] = license_choices


@app.on_event("startup")
def _warm_up_in_background():
    """Do the one-off work of the first page views right after a restart —
    compile every template (read over the network share) and load each
    survey's records and column lookups — so no user's first click waits on it."""
    def warm():
        for name in templates.env.list_templates():
            templates.get_template(name)
        for survey_type in db.SURVEY_TYPES:
            frame_choices(db.all_records(survey_type))

    threading.Thread(target=warm, daemon=True, name="warm-up").start()


@app.on_event("startup")
def _start_database_auto_refresh():
    db_import.start_auto_refresh()


def resolve_frame(choices: list[dict], frame_id: str, license: str) -> str:
    """The Frame ID to open: the Frame ID dropdown wins; otherwise the first
    one with the licence number picked — its page covers every Frame ID that
    shares the licence anyway (licence_group)."""
    frame_id, license = frame_id.strip(), license.strip()
    if frame_id or not license:
        return frame_id
    hit = (next((c for c in choices if c["license"].strip().lower() == license.lower()), None)
           or next((c for c in choices if license.lower() in c.get("licenses", ())), None))
    return hit["frame_id"] if hit else license  # unknown: fall through to the normal not-found / access-denied path


def search_choices(survey_type: str, frame_options: list[dict]) -> list[dict]:
    """What the Frame ID / License Number search lists offer: the current
    distribution only (auth.current_frames). frame_options keeps every Frame
    ID the user may open, for licence grouping and remembered searches."""
    current = auth.current_frames(survey_type)
    return [f for f in frame_options if f["frame_id"].strip() in current]


def search_frame(request: Request, survey_type: str, frame_options: list[dict], frame_id: str, license: str) -> str:
    """The Frame ID a survey page opens. A search (the Frame ID / License
    Number boxes) wins, and a search with both boxes empty forgets the
    remembered one. Opened without a search — from the menu, Home, or back
    from another tab — it reopens the last Frame ID the user searched, in
    this survey or the other one (QES and FDI share Frame IDs), when this
    survey has it and the user may open it here; else this survey's own
    last one (see remember_frame)."""
    if "frame_id" in request.query_params or "license" in request.query_params:
        if not frame_id.strip() and not license.strip():
            request.session.pop("last_frame_id", None)
            request.session["last_frames"] = {k: v for k, v in (request.session.get("last_frames") or {}).items() if k != survey_type}
        return resolve_frame(frame_options, frame_id, license)
    openable = {f["frame_id"].strip().lower(): f["frame_id"] for f in frame_options}
    for remembered in (request.session.get("last_frame_id"), (request.session.get("last_frames") or {}).get(survey_type)):
        if remembered and remembered.strip().lower() in openable:
            return openable[remembered.strip().lower()]
    return ""


def remember_frame(request: Request, survey_type: str, frame_id: str) -> None:
    """Keep the Frame ID just opened for search_frame: as this survey's own
    last one and — when the user asked for it (a search or a link carrying
    it), not when it was only reopened from memory — as the last one
    searched in any survey."""
    if "frame_id" in request.query_params or "license" in request.query_params:
        request.session["last_frame_id"] = frame_id
    request.session["last_frames"] = {**(request.session.get("last_frames") or {}), survey_type: frame_id}


def licence_group(frame_options: list[dict], frame_id: str) -> list[str]:
    """frame_id, then every other Frame ID this user can open that shares a
    licence number with it, in any quarter (frame_licences) — a search for
    one Frame ID shows them all, e.g. a current Frame ID and the older one a
    database view files the same establishment under."""
    target = frame_id.strip().lower()
    current = next((f for f in frame_options if f["frame_id"].strip().lower() == target), None)
    if not current:
        return [frame_id]
    licences = current.get("licenses") or set()
    return [current["frame_id"]] + [f["frame_id"] for f in frame_options
                                    if f is not current and licences & f.get("licenses", set())]


def figure_score(row: dict, period: str, figures: list[list[str]]) -> tuple[int, int]:
    """How fully `row` reported `period`: (non-zero figures, non-empty figures)."""
    values = [text_val(row, period, cols) for cols in figures]
    return sum(1 for v in values if to_num(v)), sum(1 for v in values if v)


def licence_row(visible: list[dict], frame_ids: list[str], figures: list[list[str]]) -> tuple[dict, dict]:
    """One row for every Frame ID in `frame_ids` (from licence_group). The
    first supplies the plain columns (FRAME_ID, ESTABLISHMENT_NAME, ...); each
    quarter is taken whole from the Frame ID that reported the most figures
    for it — ties go to the earlier one — so two submissions are never mixed
    within one quarter. Returns (row, sources): sources maps each quarter, in
    time order, to {"frame_id": where it came from, "also": other Frame IDs
    that reported figures for it too}. Both are empty if none is visible."""
    rows = {}
    for fid in frame_ids:
        records = [r for r in visible if r["frame_id"].strip().lower() == fid.strip().lower()]
        if records:
            rows[fid] = merged_row(records)
    if not rows:
        return {}, {}
    base = next(iter(rows))
    row = dict(rows[base])
    row_periods = {fid: set(get_periods(r)) for fid, r in rows.items()}
    sources = {}
    for p in set().union(*row_periods.values()):
        scores = {fid: figure_score(r, p, figures) for fid, r in rows.items() if p in row_periods[fid]}
        owner = max(scores, key=scores.get)  # first of equals wins
        sources[p] = {"frame_id": owner, "also": [f for f, s in scores.items() if f != owner and s[1]]}
        if owner != base:
            for k in [k for k in row if key_period(k) == p]:
                del row[k]
            row.update((k, v) for k, v in rows[owner].items() if key_period(k) == p)
    return row, {p: sources[p] for p in get_periods(row)}


_all_choices_cache: dict[str, tuple[list, list[dict]]] = {}


def establishment_group(survey_type: str, frame_options: list[dict], frame_id: str) -> tuple[list[str], list[dict]]:
    """The Frame IDs to combine for `frame_id` and the records to read them
    from. A Frame ID the user may open (one of their frame_options) brings in
    every Frame ID that shares a licence number with it — the same
    establishment's other Frame IDs, e.g. the one a database view files its
    current quarter under — even when those are not the user's own; a Frame
    ID they may not open brings in nothing."""
    target = frame_id.strip().lower()
    if not any(f["frame_id"].strip().lower() == target for f in frame_options):
        return [frame_id], []
    records = db.all_records(survey_type)
    cached = _all_choices_cache.get(survey_type)
    if not cached or cached[0] is not records:
        cached = _all_choices_cache[survey_type] = (records, frame_choices(records))
    return licence_group(cached[1], frame_id), records


def add_record_details(meta: dict, visible: list[dict], frame_options: list[dict], frame_ids: list[str],
                       sources: dict, lang: str) -> None:
    """Fill in the Record card's licence number, every Frame ID this user can
    open that shares it (with the quarters each one supplied), who the Frame
    ID searched is assigned to, and the supervisor's name from the account list."""
    current = next((f for f in frame_options if f["frame_id"] == frame_ids[0]), None)
    newest = current["license"] if current else ""
    others = sorted(l.upper() for l in (current or {}).get("licenses", set()) if l != newest.lower())
    meta["license"] = " · ".join([newest, *others] if newest else others)  # every licence the Frame ID has had
    meta["license_frames"] = [{"frame_id": fid, "periods": [p for p, s in sources.items() if s["frame_id"] == fid]}
                              for fid in frame_ids]
    if len(frame_ids) > 1:
        # The auditor, enumerator and supervisor are the searched Frame ID's own,
        # from its current quarter — not those of another Frame ID with the same
        # licence that supplied the combined row's current quarter.
        target = frame_ids[0].strip().lower()
        own = merged_row([r for r in visible if r["frame_id"].strip().lower() == target])
        meta.update(assignment_meta(own, lang))
    meta["supervisor_name"] = auth.account_name(meta.get("supervisor_username", ""))


def chronological(selected: list[str], available: list[str]) -> list[str]:
    """Selected periods in time order, whatever order the browser sent them."""
    return [p for p in available if p in selected]


def annual_isic_rows(source: str, frame_ids: list[str]) -> list[dict] | None:
    """Last 3 years of {YYYY}_ISIC4 data for these Frame IDs (one licence's,
    from licence_group) from an annual ISIC source (AES/FIS): each year from
    the first Frame ID that has a value for it. None if nothing was ever
    uploaded for any of them."""
    data: dict = {}
    for fid in frame_ids:
        for k, v in db.annual_isic_by_frame(source, fid).items():
            if not clean(data.get(k)):
                data[k] = v
    if not data:
        return None
    series = annual_isic_series(data)[-3:]
    rows = []
    prev = None
    for i, (year, isic4) in enumerate(series):
        changed = i > 0 and isic4 != prev
        rows.append({"year": year, "isic4": isic4, "is_base": i == 0, "changed": changed})
        prev = isic4
    return rows


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
        db.add_log(username, "UNKNOWN", "LOGIN_FAILED", "Invalid username or password", client_ip(request))
        return render(request, "login.html", error=i18n.t("invalid_login", get_lang(request)))
    if user.get("suspended"):
        db.add_log(user["username"], user["authorization"], "LOGIN_SUSPENDED", "Account is suspended", client_ip(request))
        return render(request, "login.html", error=i18n.t("account_suspended", get_lang(request)))
    auth.set_session_user(request, user)
    db.add_log(user["username"], user["authorization"], "LOGIN_SUCCESS", "", client_ip(request))
    return RedirectResponse("/dashboard", status_code=303)


@app.get("/logout")
def logout(request: Request):
    user = auth.current_user(request)
    if user:
        db.add_log(user["username"], user["authorization"], "LOGOUT", "", client_ip(request))
    request.session.clear()
    response = RedirectResponse("/login", status_code=303)
    response.headers["Clear-Site-Data"] = '"cache"'
    return response


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    user = require_login(request)
    if not user:
        return RedirectResponse("/login")
    # Only the current quarter's assignments: a Frame ID someone else has
    # taken over since an earlier quarter is no longer listed as yours.
    access = []
    team = auth.team_members(user) if auth.is_controller(user) else None
    for t in db.SURVEY_TYPES:
        period = db.current_period(t)
        current = auth.current_frames(t)  # the current distribution only, not older Frame IDs a database import files under this quarter
        records = [r for r in auth.visible_records(user, t, period) if r["frame_id"].strip() in current]
        access.append({
            "type": t,
            "label": SURVEY_LABELS[t],
            "period": period,
            "frame_ids": sorted({r["frame_id"] for r in records}),
        })
        if team:  # a CONTROLLER's Home page: each employee's Frame IDs this quarter (and from the auditor list)
            assigned = [(r["frame_id"], auth.assigned_values(r["data"], period)) for r in records]
            for member in team:
                listed = auth.listed_frames(t, member["tokens"]) or set()
                member.setdefault("counts", {})[t] = len({fid for fid, people in assigned
                                                          if member["tokens"] & people or fid.strip() in listed})
    db.add_log(user["username"], user["authorization"], "DASHBOARD_VIEW", "", client_ip(request))
    return render(request, "dashboard.html", access=access, team=team)


@app.get("/guide", response_class=HTMLResponse)
def guide(request: Request, tab: str = "qes"):
    user = require_login(request)
    if not user:
        return RedirectResponse("/login")
    db.add_log(user["username"], user["authorization"], "GUIDE_VIEW", tab, client_ip(request))
    tab = RENAMED_TYPES.get(tab.upper(), tab).lower()
    return render(request, "guide.html", guide_tab="fdi" if tab == "fdi" else "qes")


# ----------------------------------------------------------------- admin ---

@app.get("/admin", response_class=HTMLResponse)
def admin_page(request: Request):
    user = require_login(request)
    if not user or not auth.is_admin(user):
        if user:
            db.add_log(user["username"], user["authorization"], "ADMIN_ACCESS_DENIED", "/admin", client_ip(request))
        return RedirectResponse("/dashboard")
    counts = {t: db.record_count(t) for t in db.SURVEY_TYPES}
    annual_isic_counts = {s: db.record_count(s) for s in db.ANNUAL_ISIC_SOURCES}
    database_counts = {k: db.record_count(k, database=True) for k in db_import.DATASETS}
    db.add_log(user["username"], user["authorization"], "ADMIN_PANEL_VIEW", "", client_ip(request))
    account_users = sorted(
        (_public_user(u) for u in db.list_users()),
        key=lambda account: (account["username"] or "").lower(),
    )
    settings = db.get_settings()
    database = {  # the password itself never goes back to the page
        "url": settings.get("db_url", ""), "user": settings.get("db_user", ""),
        "has_password": bool(settings.get("db_password")), "refresh_minutes": settings.get("refresh_minutes", "0"),
        "quarter": db_import.quarter_setting(settings), "quarters": db_import.quarter_choices(settings.get("db_quarter", "")),
        "row_quarters": db_import.ROW_QUARTERS,
        "datasets": [{"kind": k, "query": settings.get(f"query_{k}", ""), "status": db_import.status(k, settings),
                      "loaded": [p for p in (db.last_load(k, database=False), db.last_load(k, database=True)) if any(p)]}
                     for k in db_import.DATASETS],
    }
    teams = db.all_teams()
    controllers = []
    names = {u["username"].lower(): u["name"] or u["username"] for u in account_users}
    for c in (u for u in account_users if u["authorization"] == "CONTROLLER"):
        members = teams.get(c["username"].lower(), [])
        team = {m.lower() for m in members}
        candidates = [u for u in account_users if u["authorization"] != "ADMIN" and u["username"].lower() != c["username"].lower()]
        # Members set elsewhere (the Account sheet) that the list would not offer —
        # an admin, or a username without an account — stay listed, so saving keeps them.
        listed = {u["username"].lower() for u in candidates}
        candidates += [{"username": m, "name": names.get(m.lower(), ""), "job": ""} for m in members if m.lower() not in listed]
        controllers.append({**c, "team": team, "team_names": sorted(names.get(m.lower(), m) for m in members),
                            "candidates": sorted(candidates, key=lambda u: (u["username"].lower() not in team,
                                                                          (u["name"] or u["username"]).lower()))})
    return render(request, "admin.html", counts=counts, annual_isic_counts=annual_isic_counts,
                  database_counts=database_counts, database=database,
                  controllers=controllers, assignment_lists={t: assignment_summary(t) for t in db.SURVEY_TYPES},
                  users=account_users, user_roles=USER_ROLES, logs=with_account_names(db.recent_logs(LOG_DAYS_SHOWN)),
                  message=request.query_params.get("message"), error=request.query_params.get("error"),
                  section=request.query_params.get("section", ""))


LOG_DAYS_SHOWN = 7  # the Admin page's Action logs table; the .xlsx download has everything


def with_account_names(logs: list[dict]) -> list[dict]:
    """Each log entry plus its account's display name — empty when the
    username has no account (e.g. a failed login with an unknown name)."""
    names = {auth.norm_user(u["username"]): u.get("name") or "" for u in db.list_users()}
    return [{**entry, "name": names.get(auth.norm_user(entry.get("username")), "")} for entry in logs]


def _public_user(user: dict | None) -> dict | None:
    if not user:
        return None
    return {
        "username": user.get("username") or "",
        "authorization": str(user.get("authorization") or "USER").upper(),
        "name": user.get("name") or "",
        "job": user.get("job") or "",
        "location": user.get("location") or "",
        "suspended": bool(user.get("suspended")),
    }


@app.post("/admin/users")
def admin_save_user(
    request: Request,
    username: str = Form(""),
    password: str = Form(""),
    authorization: str = Form(""),
    name: str = Form(""),
    job: str = Form(""),
    location: str = Form(""),
    original_username: str = Form(""),
):
    """Create one account, or update the account named by original_username."""
    actor = require_login(request)
    if not actor or not auth.is_admin(actor):
        if actor:
            db.add_log(actor["username"], actor["authorization"], "ADMIN_ACCESS_DENIED", "/admin/users", client_ip(request))
        return RedirectResponse("/dashboard")
    username = username.strip()
    original = original_username.strip()
    role = authorization.strip().upper()
    password = password.strip()
    if not username:
        return RedirectResponse(f"/admin?section=users&error={quote('Username is required.')}#users", status_code=303)
    if role not in USER_ROLES:
        return RedirectResponse(
            f"/admin?section=users&error={quote('Authorization must be ' + ', '.join(USER_ROLES[:-1]) + ' or ' + USER_ROLES[-1] + '.')}#users",
            status_code=303,
        )
    try:
        action = db.save_user(
            {
                "username": username,
                "password": password,
                "authorization": role,
                "name": name.strip() or username,
                "job": job.strip(),
                "location": location.strip(),
            },
            original_username=original,
        )
    except ValueError as exc:
        return RedirectResponse(f"/admin?section=users&error={quote(str(exc))}#users", status_code=303)
    saved = db.get_user(username)
    if saved and actor["username"].lower() == (original or username).lower():
        auth.set_session_user(request, saved)
    db.add_log(actor["username"], actor["authorization"], action, username, client_ip(request))
    if saved and actor["username"].lower() == (original or username).lower() and not auth.is_admin(saved):
        return RedirectResponse("/dashboard", status_code=303)
    message = f"User {username} created." if action == "USER_CREATED" else f"User {username} updated."
    return RedirectResponse(f"/admin?section=users&message={quote(message)}#users", status_code=303)


@app.post("/admin/users/suspend")
def admin_suspend_user(request: Request, original_username: str = Form(""), suspended: str = Form("1")):
    """Suspend an account (suspended=1) — it can no longer sign in, and is
    signed out on its next click — or reactivate it (suspended=0)."""
    actor = require_login(request)
    if not actor or not auth.is_admin(actor):
        if actor:
            db.add_log(actor["username"], actor["authorization"], "ADMIN_ACCESS_DENIED", "/admin/users/suspend", client_ip(request))
        return RedirectResponse("/dashboard")
    username = original_username.strip()
    suspend = suspended.strip() != "0"
    if suspend and username.lower() == actor["username"].lower():
        return RedirectResponse(f"/admin?section=users&error={quote('You cannot suspend your own account.')}#users", status_code=303)
    try:
        account = db.set_user_suspended(username, suspend)
    except ValueError as exc:
        return RedirectResponse(f"/admin?section=users&error={quote(str(exc))}#users", status_code=303)
    db.add_log(actor["username"], actor["authorization"], "USER_SUSPENDED" if suspend else "USER_REACTIVATED",
               account["username"], client_ip(request))
    message = f"User {account['username']} {'suspended' if suspend else 'reactivated'}."
    return RedirectResponse(f"/admin?section=users&message={quote(message)}#users", status_code=303)


@app.post("/admin/users/remove")
def admin_remove_user(request: Request, original_username: str = Form("")):
    """Delete an account for good, and take it out of every controller team."""
    actor = require_login(request)
    if not actor or not auth.is_admin(actor):
        if actor:
            db.add_log(actor["username"], actor["authorization"], "ADMIN_ACCESS_DENIED", "/admin/users/remove", client_ip(request))
        return RedirectResponse("/dashboard")
    username = original_username.strip()
    if username.lower() == actor["username"].lower():
        return RedirectResponse(f"/admin?section=users&error={quote('You cannot remove your own account.')}#users", status_code=303)
    try:
        account = db.delete_user(username)
    except ValueError as exc:
        return RedirectResponse(f"/admin?section=users&error={quote(str(exc))}#users", status_code=303)
    db.add_log(actor["username"], actor["authorization"], "USER_REMOVED",
               f"{account['username']} ({account['authorization']}, {account['name'] or '-'})", client_ip(request))
    return RedirectResponse(f"/admin?section=users&message={quote('User ' + account['username'] + ' removed.')}#users", status_code=303)


@app.post("/admin/controllers/{controller}")
def admin_save_team(request: Request, controller: str, members: list[str] = Form(default=[])):
    """Set which employees a CONTROLLER oversees — they then see exactly the
    Frame IDs assigned to those employees (auth.team_tokens)."""
    actor = require_login(request)
    if not actor or not auth.is_admin(actor):
        if actor:
            db.add_log(actor["username"], actor["authorization"], "ADMIN_ACCESS_DENIED", f"/admin/controllers/{controller}", client_ip(request))
        return RedirectResponse("/dashboard")
    account = db.get_user(controller)
    if not account or str(account["authorization"]).upper() != "CONTROLLER":
        return RedirectResponse(f"/admin?section=controllers&error={quote(controller + ' is not a CONTROLLER account.')}#controllers", status_code=303)
    team = db.set_team(account["username"], members)
    db.add_log(actor["username"], actor["authorization"], "CONTROLLER_TEAM_SAVED",
               f"{account['username']}: {', '.join(team) or '(none)'}", client_ip(request))
    message = f"{account['name'] or account['username']}: {len(team)} employee{'s' if len(team) != 1 else ''} saved."
    return RedirectResponse(f"/admin?section=controllers&message={quote(message)}#controllers", status_code=303)


@app.get("/admin/logs/export.xlsx")
def download_logs_export(request: Request):
    """Admin-only full activity-log export — every entry ever recorded,
    not just the most recent 200 shown on the Admin page."""
    user = require_login(request)
    if not user or not auth.is_admin(user):
        if user:
            db.add_log(user["username"], user["authorization"], "ADMIN_ACCESS_DENIED", "/admin/logs/export.xlsx", client_ip(request))
        return RedirectResponse("/dashboard")
    logs = with_account_names(db.all_logs())
    content = build_logs_workbook(logs)
    db.add_log(user["username"], user["authorization"], "LOGS_EXPORTED", str(len(logs)), client_ip(request))
    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="activity_logs.xlsx"'},
    )


@app.get("/admin/users/export.xlsx")
def download_users_export(request: Request):
    """Admin-only export of every user account, in the Admin page's order
    and columns (no passwords)."""
    user = require_login(request)
    if not user or not auth.is_admin(user):
        if user:
            db.add_log(user["username"], user["authorization"], "ADMIN_ACCESS_DENIED", "/admin/users/export.xlsx", client_ip(request))
        return RedirectResponse("/dashboard")
    users = sorted((_public_user(u) for u in db.list_users()), key=lambda account: (account["username"] or "").lower())
    teams = db.all_teams()
    for u in users:  # a CONTROLLER's employees, in the form the Account sheet's Employees column takes
        u["employees"] = ", ".join(teams.get(u["username"].lower(), [])) if u["authorization"] == "CONTROLLER" else ""
        u["status"] = "Suspended" if u["suspended"] else "Active"
    content = build_users_workbook(users)
    db.add_log(user["username"], user["authorization"], "USERS_EXPORTED", str(len(users)), client_ip(request))
    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="users.xlsx"'},
    )


@app.get("/admin/variables/{survey_type}.xlsx")
def download_variable_reference(request: Request, survey_type: str):
    """Admin-only technical variable reference: every column the app looks
    for, the label auditors see, a description, and calc/comparison usage."""
    user = require_login(request)
    if not user or not auth.is_admin(user):
        if user:
            db.add_log(user["username"], user["authorization"], "ADMIN_ACCESS_DENIED", f"/admin/variables/{survey_type}.xlsx", client_ip(request))
        return RedirectResponse("/dashboard")
    survey_type = survey_type.upper()
    if survey_type not in db.SURVEY_TYPES:
        return RedirectResponse("/admin?error=Unknown+survey+type", status_code=303)
    content = variable_reference.build_workbook(survey_type)
    db.add_log(user["username"], user["authorization"], "VARIABLE_REFERENCE_DOWNLOADED", survey_type, client_ip(request))
    filename = f"{survey_type}_variables.xlsx"
    return StreamingResponse(
        iter([content]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


UPLOAD_KINDS = {*db.SURVEY_TYPES, *db.ANNUAL_ISIC_SOURCES, "ACCOUNTS", *(f"ASSIGN_{t}" for t in db.SURVEY_TYPES)}


def assignment_summary(survey_type: str) -> dict | None:
    """The Admin page's view of a survey's auditor list (None without one):
    each auditor's Frame IDs, and what does not line up with the data or the
    accounts."""
    rows = db.assignments(survey_type)
    if rows is None:
        return None
    users = db.list_users()
    accounts = {auth.norm_user(u["username"]) for u in users} | {auth.norm_user(u["name"]) for u in users if u.get("name")}
    by_auditor: dict[str, dict] = {}
    for r in rows:
        a = by_auditor.setdefault(r["username"].lower(), {"username": r["username"], "name": r["name"], "frames": set()})
        a["frames"].add(r["frame_id"])
    auditors = sorted(by_auditor.values(), key=lambda a: (a["name"] or a["username"]).lower())
    listed = {r["frame_id"].strip() for r in rows}
    in_data = {r["frame_id"].strip() for r in db.all_records(survey_type)}
    owners = Counter(r["frame_id"] for r in rows)
    return {
        "meta": json.loads(db.get_settings().get(f"assignments_{survey_type}") or "{}"),
        "auditors": auditors,
        "frames": len(listed),
        "not_in_data": sorted(listed - in_data),
        "unassigned": len(in_data - listed),
        "no_account": [a for a in auditors if auth.norm_user(a["username"]) not in accounts
                       and auth.norm_user(a["name"]) not in accounts],
        "shared": sorted(f for f, n in owners.items() if n > 1),
    }


def assignments_redirect(message: str = "", error: str = "") -> RedirectResponse:
    text = f"message={quote(message)}" if message else f"error={quote(error)}"
    return RedirectResponse(f"/admin?section=assignments&{text}#assignments", status_code=303)


@app.post("/admin/upload/{kind}/preview")
async def upload_preview(request: Request, kind: str, file: UploadFile = File(...)):
    """Step 1: stash the file and let the admin pick which sheet to import."""
    user = require_login(request)
    if not user or not auth.is_admin(user):
        if user:
            db.add_log(user["username"], user["authorization"], "ADMIN_ACCESS_DENIED", f"/admin/upload/{kind}/preview", client_ip(request))
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
        if user:
            db.add_log(user["username"], user["authorization"], "ADMIN_ACCESS_DENIED", f"/admin/upload/{kind}/confirm", client_ip(request))
        return RedirectResponse("/dashboard")
    kind = kind.upper()
    if kind not in UPLOAD_KINDS:
        return RedirectResponse("/admin?error=Unknown+upload+type", status_code=303)
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
            for u in users:  # an Employees column, when the sheet has one, sets each CONTROLLER's team
                if u["authorization"] == "CONTROLLER" and u.get("employees") is not None:
                    db.set_team(u["username"], u["employees"])
            db.add_log(user["username"], user["authorization"], "AUTH_FILE_LOADED", f"{filename} / sheet: {sheet}", client_ip(request))
            msg = f"Account+database+loaded+({len(users)}+users)"
        elif kind.startswith("ASSIGN_"):
            survey = kind.removeprefix("ASSIGN_")
            entries = excel_utils.parse_assignment_sheet(content, sheet)
            db.replace_assignments(survey, entries, {"file": filename, "sheet": sheet, "by": user["username"]})
            s = assignment_summary(survey)
            notes = [f"{len(s['not_in_data'])} of its Frame IDs are not in the {survey} data" if s["not_in_data"] else "",
                     f"{s['unassigned']} {survey} Frame IDs are in nobody's list" if s["unassigned"] else "",
                     f"{len(s['no_account'])} auditors have no account" if s["no_account"] else ""]
            message = f"{survey} auditor list loaded: {len(s['auditors'])} auditors, {s['frames']} Frame IDs" + "".join(f"; {n}" for n in notes if n)
            db.add_log(user["username"], user["authorization"], f"{survey}_ASSIGNMENTS_UPLOADED",
                       f"{filename} / sheet: {sheet} / {message}", client_ip(request))
            return assignments_redirect(message=message)
        else:
            rows, notes = excel_utils.rows_for_survey(content, sheet)
            if not rows:
                raise ValueError("No records were imported. Make sure the sheet has a Frame ID / FRAME_ID column.")
            # While the dataset has a database query, the upload holds the other
            # quarters: it replaces only the Excel part (database.replace_records).
            keep = bool(db.get_settings().get(f"query_{kind}", "").strip())
            count = db.replace_records(kind, rows, filename, keep_database=keep)
            if keep and (kept := db.record_count(kind, database=True)):
                notes.append(f"the database import's {kept} records are kept, and replace this file's values for their quarter")
            db.add_log(user["username"], user["authorization"], f"{kind}_DATA_UPLOADED",
                       f"{filename} / sheet: {sheet} / records: {count}" + "".join(f" / {n}" for n in notes), client_ip(request))
            msg = quote(f"{kind}: {count} records loaded" + "".join(f"; {n}" for n in notes))
        return RedirectResponse(f"/admin?message={msg}", status_code=303)
    except Exception as exc:  # noqa: BLE001
        return RedirectResponse(f"/admin?error={exc}", status_code=303)
    finally:
        tempstore.delete(token)


@app.post("/admin/assignments/{survey_type}/clear")
def admin_clear_assignments(request: Request, survey_type: str):
    """Remove a survey's auditor list: the data's assignment columns decide again."""
    user = require_login(request)
    if not user or not auth.is_admin(user):
        if user:
            db.add_log(user["username"], user["authorization"], "ADMIN_ACCESS_DENIED", f"/admin/assignments/{survey_type}/clear", client_ip(request))
        return RedirectResponse("/dashboard")
    survey_type = survey_type.upper()
    if survey_type not in db.SURVEY_TYPES:
        return assignments_redirect(error="Unknown survey.")
    db.clear_assignments(survey_type)
    db.add_log(user["username"], user["authorization"], f"{survey_type}_ASSIGNMENTS_REMOVED", "", client_ip(request))
    return assignments_redirect(message=f"{survey_type} auditor list removed: the data's assignment columns decide again.")


@app.get("/admin/assignments/{survey_type}.xlsx")
def download_assignments(request: Request, survey_type: str):
    """A survey's current auditor list, in the layout it is uploaded in."""
    user = require_login(request)
    if not user or not auth.is_admin(user):
        if user:
            db.add_log(user["username"], user["authorization"], "ADMIN_ACCESS_DENIED", f"/admin/assignments/{survey_type}.xlsx", client_ip(request))
        return RedirectResponse("/dashboard")
    survey_type = survey_type.upper()
    summary = assignment_summary(survey_type) if survey_type in db.SURVEY_TYPES else None
    if not summary:
        return assignments_redirect(error=f"{survey_type} has no auditor list.")
    db.add_log(user["username"], user["authorization"], f"{survey_type}_ASSIGNMENTS_DOWNLOADED", "", client_ip(request))
    return StreamingResponse(
        iter([build_assignments_workbook(summary["auditors"])]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{survey_type}_auditor_assignments.xlsx"'},
    )


def database_redirect(message: str = "", error: str = "") -> RedirectResponse:
    text = f"message={quote(message)}" if message else f"error={quote(error)}"
    return RedirectResponse(f"/admin?section=database&{text}#database", status_code=303)


@app.post("/admin/database")
def admin_database_connection(request: Request, db_url: str = Form(""), db_user: str = Form(""),
                              db_password: str = Form(""), refresh_minutes: str = Form("0"), db_quarter: str = Form(""),
                              action: str = Form("save")):
    """Save the database connection (JDBC URL, credentials, auto-refresh, the
    quarter of the queries' unlabelled columns); with action=test, then try
    it. A blank password keeps the saved one."""
    user = require_login(request)
    if not user or not auth.is_admin(user):
        if user:
            db.add_log(user["username"], user["authorization"], "ADMIN_ACCESS_DENIED", "/admin/database", client_ip(request))
        return RedirectResponse("/dashboard")
    db_url = db_url.strip()
    try:
        if db_url:
            db_import.parse_jdbc_url(db_url)
    except ValueError as exc:
        return database_redirect(error=str(exc))
    try:
        minutes = max(0, int(refresh_minutes.strip() or 0))
    except ValueError:
        return database_redirect(error="Auto-refresh must be a whole number of minutes (0 = off).")
    values = {"db_url": db_url, "db_user": db_user.strip(), "refresh_minutes": str(minutes)}
    try:
        values["db_quarter"] = db_import.clean_quarter(db_quarter or db_import.quarter_setting(db.get_settings()))
    except ValueError as exc:
        return database_redirect(error=str(exc))
    if db_password:
        values["db_password"] = db_password
    db.save_settings(values)
    db.add_log(user["username"], user["authorization"], "DB_CONNECTION_SAVED",
               f"{db_import.describe(db_url)} / auto-refresh: {minutes or 'off'} min / quarter: {values['db_quarter']}",
               client_ip(request))
    if action != "test":
        return database_redirect(message="Database connection saved.")
    try:
        message = db_import.test_connection(db.get_settings())
    except Exception as exc:  # noqa: BLE001
        db.add_log(user["username"], user["authorization"], "DB_CONNECTION_TEST_FAILED", str(exc), client_ip(request))
        return database_redirect(error=f"Connection failed: {exc}")
    db.add_log(user["username"], user["authorization"], "DB_CONNECTION_TESTED", message, client_ip(request))
    return database_redirect(message=message)


@app.post("/admin/database/{kind}")
def admin_database_query(request: Request, kind: str, query: str = Form(""), action: str = Form("save")):
    """Save one dataset's SELECT query; with action=preview, run it and say
    what it would load; with action=import, load it in place of the dataset."""
    user = require_login(request)
    if not user or not auth.is_admin(user):
        if user:
            db.add_log(user["username"], user["authorization"], "ADMIN_ACCESS_DENIED", f"/admin/database/{kind}", client_ip(request))
        return RedirectResponse("/dashboard")
    kind = kind.upper()
    if kind not in db_import.DATASETS:
        return database_redirect(error="Unknown dataset.")
    query = query.strip()
    try:
        if query:
            db_import.clean_query(query)
    except ValueError as exc:
        return database_redirect(error=f"{kind}: {exc}")
    if query != db.get_settings().get(f"query_{kind}", ""):
        db.save_settings({f"query_{kind}": query})
        db.add_log(user["username"], user["authorization"], f"{kind}_DB_QUERY_SAVED", query[:500], client_ip(request))
    if action == "preview":
        try:
            return database_redirect(message=f"{kind}: {db_import.preview(kind)}")
        except Exception as exc:  # noqa: BLE001
            return database_redirect(error=f"{kind}: {exc}")
    if action == "import":
        ok, message = db_import.import_dataset(kind, user["username"], user["authorization"], client_ip(request))
        return database_redirect(message=message) if ok else database_redirect(error=message)
    return database_redirect(message=f"{kind}: query saved." if query else f"{kind}: query cleared — this dataset is updated by Excel upload only.")


# ---------------------------------------------------------------- survey ---

def renamed_type_redirect(request: Request, survey_type: str) -> RedirectResponse:
    """The same page under the survey's new name, keeping the query string:
    /survey/fisq?frame_id=1 -> /survey/fdi?frame_id=1."""
    path = request.url.path.rsplit("/", 1)[0] + "/" + RENAMED_TYPES[survey_type.upper()].lower()
    query = request.url.query
    return RedirectResponse(path + (f"?{query}" if query else ""), status_code=301)


@app.get("/survey/{survey_type}", response_class=HTMLResponse)
def survey_search(request: Request, survey_type: str, frame_id: str = "", license: str = "",
                   mode: str = "compare", quarter: str = "", quarters: list[str] = Query(default=[])):
    if survey_type.upper() in RENAMED_TYPES:
        return renamed_type_redirect(request, survey_type)
    user = require_login(request)
    if not user:
        return RedirectResponse("/login")
    survey_type = survey_type.upper()
    if survey_type not in db.SURVEY_TYPES:
        return RedirectResponse("/dashboard")
    lang = get_lang(request)

    visible = auth.visible_records(user, survey_type)
    frame_options = frame_choices(visible)
    frame_id = search_frame(request, survey_type, frame_options, frame_id, license)

    result = None
    error = None
    row = None
    periods = []
    meta = None
    raw_fields = None
    period_sources = {}

    if frame_id:
        frame_ids, pool = establishment_group(survey_type, frame_options, frame_id)
        row, sources = licence_row(pool, frame_ids, REPORTED_FIGURES.get(survey_type, []))
        if not row:
            error = i18n.t("not_found_or_access", lang)
            db.add_log(user["username"], user["authorization"], "RECORD_ACCESS_DENIED", f"{survey_type} / {frame_id}", client_ip(request))
        else:
            remember_frame(request, survey_type, frame_ids[0])
            if len(frame_ids) > 1:
                period_sources = sources
            periods = get_periods(row)
            if survey_type in CALCULATED_TYPES:
                if mode == "compare":
                    # Default: current quarter + previous 3 (last 4 available periods).
                    selected = chronological(quarters, periods) or periods[-LAST_N_PERIODS:]
                else:
                    selected = [quarter] if quarter in periods else (periods[-1:] if periods else [])
                calc_fn = calc_qes if survey_type == "QES" else calc_fdi
                extra = {"lang": lang} if survey_type == "QES" else {}
                result = calc_fn(row, selected, mode, **extra) if periods else None
                if not periods:
                    error = "No quarter/year period was detected in the uploaded columns for this record."
                else:
                    meta = record_meta(row, lang)
                    add_record_details(meta, visible, frame_options, frame_ids, sources, lang)
            else:
                raw_fields = {k: v for k, v in row.items() if not str(k).startswith("__")}
            db.add_log(user["username"], user["authorization"], "SURVEY_FRAME_VIEW", f"{survey_type} / {' + '.join(frame_ids)}", client_ip(request))

    # Beginning-of-quarter vs end-of-quarter asset equation, per calculated
    # quarter (FDI only) — بداية المدة و نهاية المدة.
    movement_map = {}
    if survey_type == "FDI" and result and result.get("calcs"):
        movement_map = {c["period"]: calc_fdi_period_movement(row, c["period"]) for c in result["calcs"]}

    # Granular survey-form line items, shown alongside the main figures —
    # read-only detail, not part of any calculation above.
    detail_map = {}
    if survey_type == "QES" and result and result.get("calcs"):
        detail_map = {c["period"]: qes_detail_groups(row, c["period"], lang) for c in result["calcs"]}
    elif survey_type == "FDI" and result and result.get("calcs"):
        detail_map = {c["period"]: fdi_component_groups(row, c["period"], lang) for c in result["calcs"]}

    template = f"survey_{survey_type.lower()}.html" if survey_type in CALCULATED_TYPES else "survey_generic.html"
    return render(request, template, survey_type=survey_type, frame_id=frame_id, mode=mode,
                  license=license.strip(), period_sources=period_sources,
                  quarter=quarter, quarters=quarters, frame_options=frame_options, periods=periods,
                  search_options=search_choices(survey_type, frame_options),
                  result=result, error=error, row=row, meta=meta, raw_fields=raw_fields, movement_map=movement_map,
                  detail_map=detail_map)


# ------------------------------------------------------------- comparison ---

@app.get("/compare/{survey_type}", response_class=HTMLResponse)
def compare_view(request: Request, survey_type: str, frame_id: str = "", license: str = "",
                 periods: list[str] = Query(default=[])):
    if survey_type.upper() in RENAMED_TYPES:
        return renamed_type_redirect(request, survey_type)
    user = require_login(request)
    if not user:
        return RedirectResponse("/login")
    survey_type = survey_type.upper()
    if survey_type not in CALCULATED_TYPES:
        return RedirectResponse("/dashboard")
    lang = get_lang(request)

    visible = auth.visible_records(user, survey_type)
    frame_options = frame_choices(visible)
    frame_id = search_frame(request, survey_type, frame_options, frame_id, license)
    error = None
    row = None
    all_periods = []
    calcs = []
    meta = None
    frame_ids = []
    period_sources = {}

    if frame_id:
        frame_ids, pool = establishment_group(survey_type, frame_options, frame_id)
        row, sources = licence_row(pool, frame_ids, REPORTED_FIGURES[survey_type])
        if not row:
            error = i18n.t("not_found_or_access", lang)
            db.add_log(user["username"], user["authorization"], "RECORD_ACCESS_DENIED", f"{survey_type} / {frame_id}", client_ip(request))
        else:
            remember_frame(request, survey_type, frame_ids[0])
            if len(frame_ids) > 1:
                period_sources = sources
            all_periods = get_periods(row)
            # Default: current quarter + previous 3 (last 4 available periods).
            chosen = chronological(periods, all_periods) or all_periods[-LAST_N_PERIODS:]
            calc_period_fn = (lambda p: calc_qes(row, [p], "single", lang)["first"]) if survey_type == "QES" else \
                              (lambda p: calc_fdi(row, [p], "single")["first"])
            calcs = [calc_period_fn(p) for p in chosen]
            if calcs:
                meta = record_meta(row, lang)
                add_record_details(meta, visible, frame_options, frame_ids, sources, lang)
            db.add_log(user["username"], user["authorization"], "PROJECT_COMPARISON_VIEW", f"{survey_type} / {' + '.join(frame_ids)}", client_ip(request))

    # First-to-last change across the compared quarters: Added value (QES),
    # Calculated equity (FDI).
    headline = None
    if len(calcs) >= 2:
        key = "added_value" if survey_type == "QES" else "calculated_equity"
        first, last = calcs[0], calcs[-1]
        change = (last[key] or 0) - (first[key] or 0)
        percent = (change / abs(first[key]) * 100) if first[key] else 0.0
        headline = {"first": first, "last": last, "change": change, "percent": percent}

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

    # ISIC4 classification stability across the compared quarters (QES + FDI),
    # with the licence number each quarter was reported under. Many quarters'
    # files carry no licence number (FDI's older quarter blocks have no such
    # column), so those fall back to the licence of the Frame ID that reported
    # the quarter — flagged, so the page can tell the two apart.
    isic_rows = []
    if calcs:
        frame_licence = {f["frame_id"]: f["license"] for f in frame_options}
        prev_isic = None
        for i, c in enumerate(calcs):
            isic = text_val(row, c["period"], ["ISIC4"])
            changed = i > 0 and isic != prev_isic
            own_licence = text_val(row, c["period"], LICENSE_NUMBER)
            owner = (period_sources.get(c["period"]) or {}).get("frame_id") or frame_ids[0]
            isic_rows.append({"period": c["period"], "license": own_licence or frame_licence.get(owner, ""),
                              "license_in_quarter": bool(own_licence),
                              "isic4": isic, "is_base": i == 0, "changed": changed})
            prev_isic = isic

    # Beginning-of-quarter vs end-of-quarter asset equation, within each
    # quarter (FDI only) — بداية المدة و نهاية المدة.
    movement_rows = []
    if survey_type == "FDI" and calcs:
        movement_rows = [mv for c in calcs if (mv := calc_fdi_period_movement(row, c["period"]))]

    # ±30% quarter-over-quarter movement flag (FDI only): flags any of the
    # 5 tracked totals that moved more than 30% since the previous selected
    # quarter, per the source form's validation rule.
    movement_flags = []
    if survey_type == "FDI" and calcs:
        for i, c in enumerate(calcs):
            if i == 0:
                continue
            prev = calcs[i - 1]
            for key, label_en, label_ar in FDI_FLAG_METRICS:
                prev_val, cur_val = prev.get(key) or 0, c.get(key) or 0
                if not prev_val:
                    continue
                change_pct = (cur_val - prev_val) / abs(prev_val) * 100
                if abs(change_pct) > 30:
                    movement_flags.append({
                        "period": c["period"], "metric": label_ar if lang == "ar" else label_en,
                        "prev_value": prev_val, "value": cur_val, "change_pct": change_pct,
                    })

    # Granular survey-form line items, per compared quarter — read-only
    # detail, not part of any calculation above.
    detail_rows = []
    if survey_type == "QES" and calcs:
        detail_rows = [{"period": c["period"], "groups": qes_detail_groups(row, c["period"], lang)} for c in calcs]
    elif survey_type == "FDI" and calcs:
        detail_rows = [{"period": c["period"], "groups": fdi_component_groups(row, c["period"], lang)} for c in calcs]

    # Annual ISIC cross-check against the same Frame IDs' annual filing,
    # last 3 years — QES compares against AES, FDI compares against FIS.
    aes_isic_rows = annual_isic_rows("AES", frame_ids) if (survey_type == "QES" and frame_id) else None
    fis_isic_rows = annual_isic_rows("FIS", frame_ids) if (survey_type == "FDI" and frame_id) else None

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
    return render(request, template, survey_type=survey_type, frame_id=frame_id,
                  license=license.strip(), period_sources=period_sources,
                  frame_options=frame_options, search_options=search_choices(survey_type, frame_options),
                  all_periods=all_periods, selected_periods=[c["period"] for c in calcs], calcs=calcs,
                  inventory_rows=inventory_rows, employee_rows=employee_rows, isic_rows=isic_rows,
                  movement_rows=movement_rows, movement_flags=movement_flags, detail_rows=detail_rows,
                  aes_isic_rows=aes_isic_rows, fis_isic_rows=fis_isic_rows,
                  headline=headline, error=error, row=row, meta=meta, chart=chart)


# ------------------------------------------------------------------ main ---

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.app:app", host="127.0.0.1", port=8000, reload=True)
