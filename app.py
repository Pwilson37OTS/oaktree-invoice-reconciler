"""Bullhorn ↔ QBO Revenue Reconciler — Streamlit dashboard.

Two top-level tabs: OTS and CTS. Each tab is a self-contained reconciler
view (the auditor for one entity may not be the auditor for the other).

Within each tab:
  - Upload source files (or auto-load if running on Phil's machine where
    the canonical files live at known paths)
  - Continuous Audit  — billing-team daily/weekly self-serve
  - Month Close       — month-end sign-off (Excel snapshot via download)

Cloud-friendly: files uploaded by the user are processed in-memory in a
TemporaryDirectory and never persist on the server. Each user-session has
its own state.

Run locally with:
    streamlit run app.py
"""
from __future__ import annotations
import io
import json
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st

from reconcile import (
    ENTITY_CONFIG, entity_paths, reconcile_entity, write_excel_snapshot,
    STATUS_CLEAN, STATUS_PENNY, STATUS_ZERO_NO_ACTION, STATUS_EXPENSE_RECLASS,
)
from period_calendar import to_period_anchor


HERE = Path(__file__).resolve().parent
OUT = HERE / "output"

ENTITIES: dict[str, dict] = {
    "ots": {"label": "OTS — OakTree Software",            "short": "OTS"},
    "cts": {"label": "CTS — Complete Technical Services", "short": "CTS"},
}

STATUS_COLORS = {
    "CLEAN": "#e8f5e9",
    "PENNY_ROUND": "#fff8e1",
    "MINOR_VARIANCE": "#ffe0b2",
    "MATERIAL_VARIANCE": "#ffcdd2",
    "ATS_ONLY_REVENUE_NOT_IN_QBO": "#ffcdd2",
    "QBO_ONLY_NO_ATS_SOURCE": "#bbdefb",
    "ZERO_AMOUNT_NO_ACTION": "#eceff1",
    "EXPENSE_RECLASS_NO_ACTION": "#e1f5fe",
}

STATUS_LABEL = {
    "CLEAN": "Clean",
    "PENNY_ROUND": "Penny rounding",
    "MINOR_VARIANCE": "Minor variance (≤ $1)",
    "MATERIAL_VARIANCE": "Material variance",
    "ATS_ONLY_REVENUE_NOT_IN_QBO": "Bullhorn only — QBO missing this",
    "QBO_ONLY_NO_ATS_SOURCE": "QBO only — no Bullhorn source",
    "ZERO_AMOUNT_NO_ACTION": "$0 — no invoice expected",
    "EXPENSE_RECLASS_NO_ACTION": "Expense reclass — reporting only",
}

NON_ACTIONABLE = {"CLEAN", "PENNY_ROUND", "ZERO_AMOUNT_NO_ACTION", "EXPENSE_RECLASS_NO_ACTION"}

RENDER_ROW_CAP = 500

# Per-entity QBO-account headline breakdown. Each entry is a label shown on
# the tile and a list of account-name prefixes (startswith match against the
# QBO line's "Full name" / "Distribution account" column).
# Add categories or rename labels here — the dashboard picks them up
# automatically on next refresh. An empty list = no breakdown row for that entity.
ACCOUNT_BREAKDOWN: dict[str, list[tuple[str, list[str]]]] = {
    "ots": [
        ("Employee Advance", ["Employee Advance"]),
        ("Primary Sales",    ["Primary Sales"]),     # matches "Primary Sales:Revenue-Placement" etc.
        ("Employee Wages",   ["Employee Wages"]),
    ],
    # CTS account-name conventions differ — fill in when Phil shares the
    # canonical account categories.
    "cts": [],
}

# QBO account prefixes that count as P&L revenue. The QBO P&L tile sums
# only lines whose account starts with one of these prefixes. Clearing
# accounts (Employee Advance) and cost accounts (Employee Wages) are
# intentionally excluded — they're not P&L revenue.
# Empty list = include all accounts (gross sum).
PL_REVENUE_PREFIXES: dict[str, list[str]] = {
    "ots": ["Primary Sales"],
    "cts": [],  # TBD when Phil shares CTS account conventions
}

# Entities that have month-end period cutoffs (per period_calendar). These get
# a deferral section in Month Close: revenue invoices booked in the current
# month whose accounting period is actually the next month. OakTree's OTS
# billing cycle straddles month-ends; CTS bills clean Sundays with no cutoff.
ENTITIES_WITH_DEFERRAL = {"ots"}

# Description substrings that disqualify a line from the period-end deferral.
# Direct Hire fees are one-time placement fees, not recurring weekly billing,
# so they don't follow the service-week cutoff even when booked in the window.
# Case-insensitive. Add more patterns here if other one-off charge types surface.
DEFERRAL_EXCLUDE_DESC = ("direct hire",)

# Product/Service codes that post to flush accounts (balance-sheet clearing or
# expense flush), NOT P&L revenue — so they inflate the QBO revenue figure.
# Month Close shows each flush total and a "Book Revenue" (revenue net of all
# flush lines). Per-entity list of (display label, exact Product/Service code).
# Add a tuple here to flush another Product/Service code.
FLUSH_PRODUCTS: dict[str, list[tuple[str, str]]] = {
    "cts": [
        ("CPL2 (flush account)", "CPL2"),         # -> Interco Rec-OTS
        ("Per Diem (flush account)", "Per Diem"),  # -> Employee Advance
        ("Expense/EXP (flush account)", "EXP"),    # -> Employee Advance
    ],
}

# Entities that get the service-week accrual section in Month Close. CTS bills
# bi-weekly, so a single invoice booked in one month can carry a line whose
# service week falls in the adjacent month — that line's revenue needs to be
# accrued to the month it was earned.
ENTITIES_WITH_ACCRUAL = {"cts"}

# Revenue (P&L) account prefixes per entity — used to scope the accrual to
# actual revenue lines (excludes flush/clearing accounts).
REVENUE_ACCOUNT_PREFIXES = {
    "ots": ["Primary Sales"],
    "cts": ["Revenue-Sales", "Revenue-Placement"],
}


# -------------------------- date helpers --------------------------

def _week_start(d):
    if d is None or pd.isna(d):
        return None
    return d - timedelta(days=d.weekday())


def _week_label(week_start):
    if week_start is None or pd.isna(week_start):
        return ""
    end = week_start + timedelta(days=6)
    if week_start.year == end.year:
        return f"{week_start.strftime('%b')} {week_start.day} - {end.strftime('%b')} {end.day}, {end.year}"
    return f"{week_start.strftime('%b')} {week_start.day}, {week_start.year} - {end.strftime('%b')} {end.day}, {end.year}"


# -------------------------- data layer --------------------------

def frame_payload(payload: dict) -> pd.DataFrame:
    """Convert the JSON `rows` list into a DataFrame with helper columns."""
    rows = payload.get("rows", [])
    df = pd.DataFrame(rows)
    if not df.empty:
        df["billing_date"] = pd.to_datetime(df["billing_date"]).dt.date
        df["month"] = df["billing_date"].apply(lambda d: f"{d.year:04d}-{d.month:02d}")
        df["week_start"] = df["billing_date"].apply(_week_start)
        df["week_label"] = df["week_start"].apply(_week_label)
        df["qbo_amount"] = df["qbo_amount"].astype("float64")
        df["ats_amount"] = df["ats_amount"].astype("float64")
        df["diff"] = df["diff"].astype("float64")
    return df


def run_from_uploads(entity: str, qbo_file, ats_file) -> dict:
    """Save uploads to a tempdir, run reconcile_entity, return the payload."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        qbo_path = tmp_path / (qbo_file.name or "qbo.xlsx")
        ats_path = tmp_path / (ats_file.name or "ats.xlsx")
        qbo_path.write_bytes(qbo_file.getbuffer())
        ats_path.write_bytes(ats_file.getbuffer())
        return reconcile_entity(
            entity,
            qbo_path=qbo_path,
            ats_path=ats_path,
            write_outputs=False,
            verbose=False,
        )


def maybe_autoload(entity: str) -> dict | None:
    """Auto-run the engine if source files are present at any configured
    candidate path. Order: local OneDrive (live data on Phil's machine) →
    repo-bundled data/ (Streamlit Cloud deploy). If neither is present,
    return None and the user uploads manually."""
    qbo, ats = entity_paths(entity)
    if not (qbo and ats):
        return None
    try:
        return reconcile_entity(entity, write_outputs=False, verbose=False)
    except Exception as e:
        st.warning(f"Auto-load failed for {entity.upper()}: {e}")
        return None


def get_entity_payload(entity: str) -> dict | None:
    """Return the cached payload for `entity`, attempting auto-load on miss."""
    key = f"{entity}_payload"
    if key in st.session_state:
        return st.session_state[key]
    p = maybe_autoload(entity)
    if p is not None:
        st.session_state[key] = p
    return p


# -------------------------- in-memory close log --------------------------

CLOSE_LOG_FILE = OUT / "close_log.json"


def load_close_log() -> dict:
    # Cloud filesystem is ephemeral, so this is best-effort. The user gets a
    # permanent record via the downloaded snapshot Excel.
    if "close_log" in st.session_state:
        return st.session_state["close_log"]
    if CLOSE_LOG_FILE.exists():
        try:
            log = json.loads(CLOSE_LOG_FILE.read_text())
        except Exception:
            log = {}
    else:
        log = {}
    st.session_state["close_log"] = log
    return log


def save_close_log(log: dict) -> None:
    st.session_state["close_log"] = log
    try:
        CLOSE_LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        CLOSE_LOG_FILE.write_text(json.dumps(log, indent=2, default=str))
    except Exception:
        pass  # ephemeral fs on Cloud — session_state is the canonical store


# -------------------------- rendering --------------------------

def _qbo_all(payload: dict) -> list[dict]:
    """The unclipped QBO line set for P&L/flush/deferral reporting. Falls back
    to reconciled-row lines for older payloads that predate qbo_lines_all."""
    lines = payload.get("qbo_lines_all")
    if lines is not None:
        return lines
    # Backward-compat: flatten from reconciled rows (clipped) if needed.
    return [q for row in payload.get("rows", []) for q in row.get("qbo_lines", [])]


def _pl_total(payload: dict, month: str | None, revenue_prefixes: list[str]) -> float:
    """Sum every QBO line whose Transaction Date is in `month` (YYYY-MM) AND
    whose account starts with one of `revenue_prefixes`. Matches what shows
    up on the QBO P&L statement — excludes clearing accounts like Employee
    Advance and cost accounts like Employee Wages.

    Empty `revenue_prefixes` list = no account filter (gross sum).
    `month=None` = no date filter (all months).

    Reads the UNCLIPPED QBO line set: revenue reporting must include every
    booked line (e.g. a June credit memo for an old service week that the
    reconciliation coverage clip drops).
    """
    total = 0.0
    for q in _qbo_all(payload):
        d = q.get("txn_date")
        if not d:
            continue
        if month is not None and not d.startswith(month):
            continue
        acct = (q.get("account") or "").strip()
        if revenue_prefixes and not any(acct.startswith(p) for p in revenue_prefixes):
            continue
        total += q.get("amount") or 0.0
    return round(total, 2)


def _next_month_label(month: str) -> str:
    """'2026-05' -> '2026-06'."""
    y, m = int(month[:4]), int(month[5:7])
    return f"{y + 1:04d}-01" if m == 12 else f"{y:04d}-{m + 1:02d}"


def _prev_month_label(month: str) -> str:
    """'2026-06' -> '2026-05'."""
    y, m = int(month[:4]), int(month[5:7])
    return f"{y - 1:04d}-12" if m == 1 else f"{y:04d}-{m - 1:02d}"


def accrual_items(payload: dict, month: str, revenue_prefixes: list[str],
                  direction: str) -> tuple[list[dict], float]:
    """Revenue lines whose service-week month and booking (txn) month differ
    across an adjacent month boundary — the bi-weekly straddle.

    direction="back": booked (txn) in `month`, service week in the PRIOR month
                      -> revenue to accrue OUT of `month` back to prior month.
    direction="in":   service week in `month`, booked (txn) in the NEXT month
                      -> revenue earned in `month` to accrue IN from next month.

    Restricted to revenue accounts; Direct Hire (one-off, no weekly service
    week) is excluded.
    """
    if direction == "back":
        txn_m, svc_m = month, _prev_month_label(month)
    elif direction == "in":
        txn_m, svc_m = _next_month_label(month), month
    else:
        raise ValueError(direction)

    items: list[dict] = []
    for q in _qbo_all(payload):
        txn, parsed = q.get("txn_date"), q.get("parsed_date")
        if not txn or not parsed:
            continue
        if not (txn.startswith(txn_m) and parsed.startswith(svc_m)):
            continue
        acct = (q.get("account") or "").strip()
        if revenue_prefixes and not any(acct.startswith(p) for p in revenue_prefixes):
            continue
        if any(p in (q.get("description") or "").lower() for p in DEFERRAL_EXCLUDE_DESC):
            continue
        items.append({
            "contractor": q.get("contractor", ""),
            "invoice_num": q.get("num", ""),
            "booked": txn,
            "service_week": parsed,
            "customer": q.get("customer", ""),
            "amount": q.get("amount", 0.0),
            "description": q.get("description", ""),
        })
    total = round(sum(i["amount"] for i in items), 2)
    return items, total


def deferral_items(payload: dict, month: str, revenue_prefixes: list[str]) -> tuple[list[dict], float]:
    """Revenue QBO lines whose Transaction Date is in `month` but whose
    accounting period (per period_calendar) is a LATER month — i.e. booked
    this month, belong to next month, need a deferral JE to move them forward.

    Returns (items, total). Each item is the raw qbo_line dict plus the
    contractor it belongs to and its anchor month.
    """
    items: list[dict] = []
    for q in _qbo_all(payload):
        d = q.get("txn_date")
        if not d or not d.startswith(month):
            continue
        acct = (q.get("account") or "").strip()
        if revenue_prefixes and not any(acct.startswith(p) for p in revenue_prefixes):
            continue
        # Exclude one-off charge types (e.g. Direct Hire placement fees) —
        # they don't follow the service-week deferral even when in-window.
        desc_l = (q.get("description") or "").lower()
        if any(p in desc_l for p in DEFERRAL_EXCLUDE_DESC):
            continue
        try:
            dt = date.fromisoformat(d)
        except (TypeError, ValueError):
            continue
        anchor = to_period_anchor(dt)
        if (anchor.year, anchor.month) > (dt.year, dt.month):
            items.append({
                "contractor": q.get("contractor", ""),
                "txn_date": d,
                "num": q.get("num", ""),
                "customer": q.get("customer", ""),
                "account": acct,
                "amount": q.get("amount", 0.0),
                "anchor_month": f"{anchor.year:04d}-{anchor.month:02d}",
                "description": q.get("description", ""),
            })
    total = round(sum(i["amount"] for i in items), 2)
    return items, total


def flush_lines(payload: dict, month: str, product_code: str) -> tuple[list[dict], float]:
    """All QBO lines with Product/Service == product_code and a Transaction
    Date in `month`. These post to a flush account, not P&L revenue.
    Returns (items, total)."""
    items: list[dict] = []
    for q in _qbo_all(payload):
        if (q.get("product") or "").strip().upper() != product_code.upper():
            continue
        d = q.get("txn_date")
        if not d or not d.startswith(month):
            continue
        items.append({
            "contractor": q.get("contractor", ""),
            "txn_date": d,
            "num": q.get("num", ""),
            "customer": q.get("customer", ""),
            "amount": q.get("amount", 0.0),
            "description": q.get("description", ""),
        })
    total = round(sum(i["amount"] for i in items), 2)
    return items, total


def _account_breakdown(payload: dict, filter_keys: set[str], categories: list[tuple[str, list[str]]]) -> dict[str, float]:
    """Sum sign-corrected qbo_line amounts by account category. Restricted
    to reconciled rows whose `key` is in `filter_keys` (respects the active
    Month/Week/Status/Client filters)."""
    totals = {label: 0.0 for label, _ in categories}
    if not filter_keys:
        return totals
    for row in payload.get("rows", []):
        if row["key"] not in filter_keys:
            continue
        for q in row.get("qbo_lines", []):
            acct = (q.get("account") or "").strip()
            if not acct:
                continue
            for label, prefixes in categories:
                if any(acct.startswith(p) for p in prefixes):
                    totals[label] = round(totals[label] + (q.get("amount") or 0.0), 2)
                    break  # first matching category wins; no double-counting
    return totals


def render_tiles(df: pd.DataFrame, payload: dict | None = None, entity: str | None = None,
                 pl_month: str | None = None) -> None:
    qbo_total = round(df["qbo_amount"].fillna(0).sum(), 2)
    ats_total = round(df["ats_amount"].fillna(0).sum(), 2)
    net = round(ats_total - qbo_total, 2)
    exceptions = (~df["status"].isin(NON_ACTIONABLE)).sum() if not df.empty else 0
    material = (df["status"] == "MATERIAL_VARIANCE").sum() if not df.empty else 0

    # Two rows of 3 tiles each — 6 tiles in a single row clip dollar
    # amounts on narrow screens. Top row: financial totals.
    # Bottom row: variance / health metrics.
    r1c1, r1c2, r1c3 = st.columns(3)
    r1c1.metric("Bullhorn billed", f"${ats_total:,.2f}")
    r1c2.metric("QBO revenue (recon)", f"${qbo_total:,.2f}",
                help="Reconciliation view — buckets each QBO line by its service-week date "
                     "(parsed from the Description). Used for matching against Bullhorn.")
    if payload is not None:
        revenue_prefixes = PL_REVENUE_PREFIXES.get(entity or "", [])
        pl = _pl_total(payload, pl_month, revenue_prefixes)
        r1c3.metric("QBO P&L (txn date)", f"${pl:,.2f}",
                    help="P&L view — sums revenue-account QBO lines (Primary Sales for OTS) "
                        "by their Transaction Date column (when QBO actually booked the entry). "
                        "Excludes clearing accounts (Employee Advance) and cost accounts "
                        "(Employee Wages). May differ from QBO recon when Credit Memos for "
                        "prior-month service weeks are booked in the current month.")
    else:
        r1c3.metric("QBO P&L (txn date)", "—")

    r2c1, r2c2, r2c3 = st.columns(3)
    r2c1.metric("Net variance", f"${net:,.2f}",
                delta=None if abs(net) < 0.01 else f"{net:+,.2f}",
                delta_color="inverse" if abs(net) >= 0.01 else "off")
    r2c2.metric("Exceptions", int(exceptions))
    r2c3.metric("Material variances", int(material), delta_color="inverse")

    # Optional second-row breakdown by QBO account category
    categories = ACCOUNT_BREAKDOWN.get(entity or "", [])
    if not categories or payload is None or df.empty:
        return
    filter_keys = set(df["key"])
    totals = _account_breakdown(payload, filter_keys, categories)
    cols = st.columns(len(categories))
    for col, (label, _) in zip(cols, categories):
        col.metric(label, f"${totals[label]:,.2f}")


def style_status(val):
    return f"background-color: {STATUS_COLORS.get(val, '#ffffff')}"


def render_table(df: pd.DataFrame) -> None:
    if df.empty:
        st.info("No rows match the current filters.")
        return

    total = len(df)
    truncated = total > RENDER_ROW_CAP
    if truncated:
        priority = df["status"].map({
            "MATERIAL_VARIANCE": 0,
            "ATS_ONLY_REVENUE_NOT_IN_QBO": 1,
            "QBO_ONLY_NO_ATS_SOURCE": 2,
            "MINOR_VARIANCE": 3,
            "PENNY_ROUND": 4,
            "EXPENSE_RECLASS_NO_ACTION": 5,
            "ZERO_AMOUNT_NO_ACTION": 6,
            "CLEAN": 7,
        }).fillna(99)
        df = df.assign(_pri=priority).sort_values(
            ["_pri", "diff"], ascending=[True, False],
            key=lambda s: s.abs() if s.name == "diff" else s,
        ).drop(columns=["_pri"]).head(RENDER_ROW_CAP)
        st.warning(
            f"Showing the top {RENDER_ROW_CAP:,} of {total:,} matching rows "
            f"(prioritized by status severity). Narrow with the filters above."
        )

    show = df[["contractor", "client", "billing_date", "ats_amount", "qbo_amount", "diff", "status", "candidate_id", "key"]].rename(
        columns={
            "contractor": "Contractor", "client": "Client", "billing_date": "Billing date",
            "ats_amount": "Bullhorn", "qbo_amount": "QBO", "diff": "Diff (BH − QBO)",
            "status": "Status", "candidate_id": "Cand ID", "key": "_key",
        }
    )
    styled = show.style.format({
        "Bullhorn": lambda v: "" if pd.isna(v) else f"${v:,.2f}",
        "QBO": lambda v: "" if pd.isna(v) else f"${v:,.2f}",
        "Diff (BH − QBO)": lambda v: f"${v:,.2f}",
    }).map(style_status, subset=["Status"])

    st.dataframe(styled, hide_index=True, width="stretch", height=520,
                 column_config={"_key": None})


def render_drill_down(df: pd.DataFrame, payload: dict, prefix: str) -> None:
    if df.empty:
        return
    options = df["key"].tolist()
    labels = {
        r["key"]: f"{r['contractor']}  |  {r['billing_date']}  |  diff ${r['diff']:.2f}  ({r['status']})"
        for _, r in df.iterrows()
    }
    key = st.selectbox("Drill into a reconciled key", options=options,
                       format_func=lambda k: labels.get(k, k), key=f"{prefix}_drill")
    full = next((r for r in payload["rows"] if r["key"] == key), None)
    if not full:
        return

    a, b = st.columns(2)
    with a:
        st.subheader(f"Bullhorn lines · {len(full['ats_lines'])}")
        if full["ats_lines"]:
            st.dataframe(pd.DataFrame(full["ats_lines"]), hide_index=True, width="stretch")
        else:
            st.info("No Bullhorn lines for this key.")
    with b:
        st.subheader(f"QBO lines · {len(full['qbo_lines'])}")
        if full["qbo_lines"]:
            st.dataframe(pd.DataFrame(full["qbo_lines"]), hide_index=True, width="stretch")
        else:
            st.info("No QBO lines for this key.")


def render_filters(df: pd.DataFrame, prefix: str) -> pd.DataFrame:
    with st.expander("Filters", expanded=True):
        c1, c2, c3 = st.columns([1, 1.4, 1.4])
        with c1:
            months = sorted(df["month"].dropna().unique(), reverse=True)
            month_choice = st.selectbox("Month", options=["(all)"] + list(months), key=f"{prefix}_month")
        with c2:
            week_pairs = (
                df.dropna(subset=["week_start"])
                  .drop_duplicates(subset=["week_start"])
                  .sort_values("week_start", ascending=False)[["week_start", "week_label"]]
                  .to_records(index=False).tolist()
            )
            WEEK_ALL = "(all)"
            week_options = [WEEK_ALL] + [ws for ws, _ in week_pairs]
            week_label_by_ws = {ws: lbl for ws, lbl in week_pairs}
            week_choice = st.selectbox(
                "Week (Mon–Sun)", options=week_options,
                format_func=lambda v: WEEK_ALL if v == WEEK_ALL else week_label_by_ws.get(v, str(v)),
                key=f"{prefix}_week",
            )
        with c3:
            statuses = sorted(df["status"].dropna().unique())
            status_choice = st.multiselect(
                "Status", options=statuses,
                default=[s for s in statuses if s not in NON_ACTIONABLE],
                format_func=lambda s: STATUS_LABEL.get(s, s),
                key=f"{prefix}_status",
            )

        c4, c5 = st.columns([1.5, 1.5])
        with c4:
            clients = sorted(df["client"].dropna().unique())
            client_choice = st.multiselect("Client", options=clients, key=f"{prefix}_client")
        with c5:
            search = st.text_input("Contractor search", key=f"{prefix}_search")

    out = df.copy()
    if month_choice != "(all)":
        out = out[out["month"] == month_choice]
    if week_choice != WEEK_ALL:
        out = out[out["week_start"] == week_choice]
    if status_choice:
        out = out[out["status"].isin(status_choice)]
    if client_choice:
        out = out[out["client"].isin(client_choice)]
    if search:
        out = out[out["contractor"].str.contains(search, case=False, na=False)]
    return out


# -------------------------- upload form --------------------------

def render_upload_form(entity: str) -> None:
    """Two-uploader form per entity. Replaces session_state on submit."""
    cfg = ENTITIES[entity]
    short = cfg["short"]

    st.markdown(
        f"Upload the two source files for **{short}** to run the audit. "
        f"Files are processed in your browser session — nothing is stored on the server."
    )

    c1, c2 = st.columns(2)
    qbo_file = c1.file_uploader(
        f"QBO Revenue Audit Report (audit_rev_{entity}.xlsx)",
        type=["xlsx"], key=f"{entity}_qbo_uploader",
    )
    ats_file = c2.file_uploader(
        f"Bullhorn billable charges ({'billable_charges' if entity == 'ots' else 'billable_charges_CTS'}.xlsx)",
        type=["xlsx"], key=f"{entity}_ats_uploader",
    )

    cols = st.columns([1, 1, 4])
    run_disabled = qbo_file is None or ats_file is None
    if cols[0].button(f"Run {short} reconciliation", key=f"{entity}_run", disabled=run_disabled, type="primary"):
        with st.spinner(f"Reconciling {short}…"):
            payload = run_from_uploads(entity, qbo_file, ats_file)
        st.session_state[f"{entity}_payload"] = payload
        st.success(f"{short} reconciled. {payload['summary']['actionable_exceptions']} actionable exceptions.")
        st.rerun()

    has_payload = f"{entity}_payload" in st.session_state
    if cols[1].button(f"Clear {short}", key=f"{entity}_clear", disabled=not has_payload):
        del st.session_state[f"{entity}_payload"]
        st.rerun()


# -------------------------- per-tab pages --------------------------

def page_continuous(payload: dict, df: pd.DataFrame, entity: str) -> None:
    filtered = render_filters(df, prefix=f"{entity}_cont")
    selected_month = st.session_state.get(f"{entity}_cont_month")
    pl_month = selected_month if selected_month and selected_month != "(all)" else None
    render_tiles(filtered, payload=payload, entity=entity, pl_month=pl_month)
    st.markdown("### Reconciled rows")
    render_table(filtered)
    st.markdown("### Drill-down")
    render_drill_down(filtered, payload, prefix=f"{entity}_cont")


def render_flush_section(payload: dict, month: str, entity: str) -> None:
    """Flush-account adjustment: certain Product/Service codes (CPL2, Per Diem)
    post to flush accounts (balance sheet / expense), not P&L revenue, so they
    inflate the QBO revenue figure. Show each flush total and a Book Revenue
    (revenue net of all flush lines)."""
    flush_cfg = FLUSH_PRODUCTS.get(entity)
    if not flush_cfg:
        return

    revenue_prefixes = PL_REVENUE_PREFIXES.get(entity, [])
    gross = _pl_total(payload, month, revenue_prefixes)  # revenue as booked (incl. flush)

    results = []  # (label, code, items, total)
    total_flush = 0.0
    for label, code in flush_cfg:
        items, total = flush_lines(payload, month, code)
        results.append((label, code, items, total))
        total_flush = round(total_flush + total, 2)
    book = round(gross - total_flush, 2)

    st.markdown("### Book revenue (flush-account adjustment)")

    # Headline tiles: Revenue as booked + one per flush code, chunked 3/row so
    # large numbers don't get truncated on narrow screens.
    tiles = [("Revenue (as booked)", f"${gross:,.2f}",
              "QBO revenue booked by Transaction Date, including flush lines.")]
    for label, code, items, total in results:
        tiles.append((label, f"${total:,.2f}",
                      f"Lines with Product/Service = {code}. These post to a flush "
                      f"account, not P&L revenue, so they're removed from Book Revenue."))
    for start in range(0, len(tiles), 3):
        row = tiles[start:start + 3]
        cols = st.columns(3)
        for col, (lbl, val, help_txt) in zip(cols, row):
            col.metric(lbl, val, help=help_txt)

    st.metric("Book Revenue", f"${book:,.2f}",
              help="Revenue as booked minus all flush lines = what hits the P&L.")

    # Per-code line-item detail for verification.
    for label, code, items, total in results:
        if not items:
            continue
        with st.expander(f"Show the {len(items)} {code} line(s) in {month} — ${total:,.2f}", expanded=False):
            cdf = pd.DataFrame(items)[
                ["contractor", "txn_date", "num", "customer", "amount", "description"]
            ].rename(columns={
                "contractor": "Contractor", "txn_date": "Txn date", "num": "Invoice #",
                "customer": "Customer", "amount": "Amount", "description": "Description",
            })
            styled = cdf.style.format({"Amount": lambda v: f"${v:,.2f}"})
            st.dataframe(styled, hide_index=True, width="stretch", height=280)


def _accrual_detail_df(items: list[dict]) -> "pd.DataFrame":
    return pd.DataFrame(items)[
        ["contractor", "invoice_num", "booked", "service_week", "customer", "amount", "description"]
    ].rename(columns={
        "contractor": "Contractor", "invoice_num": "Invoice #", "booked": "Booked (txn)",
        "service_week": "Service week", "customer": "Customer", "amount": "Amount",
        "description": "Description",
    })


def render_accrual_section(payload: dict, month: str, entity: str) -> None:
    """Service-week accrual for bi-weekly invoices straddling a month boundary.
    Accrue-back: booked this month, earned last month -> move OUT to prior month.
    Accrue-in:   earned this month, booked next month -> move IN from next month."""
    if entity not in ENTITIES_WITH_ACCRUAL:
        return

    rev = REVENUE_ACCOUNT_PREFIXES.get(entity, [])
    prev_m, next_m = _prev_month_label(month), _next_month_label(month)
    back_items, back_total = accrual_items(payload, month, rev, "back")
    in_items, in_total = accrual_items(payload, month, rev, "in")

    if not back_items and not in_items:
        return

    st.markdown("### Accruals (bi-weekly service-week timing)")

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Accrue back → {prev_m}", f"${back_total:,.2f}",
              help=f"Revenue lines booked in {month} whose service week is in {prev_m}. "
                   f"Move this OUT of {month} into {prev_m}.")
    c2.metric(f"Accrue in ← {next_m}", f"${in_total:,.2f}",
              help=f"Revenue lines earned in {month} (service week) but booked in {next_m}. "
                   f"Move this INTO {month}. May be incomplete until {next_m} is fully booked.")
    net = round(in_total - back_total, 2)
    c3.metric(f"Net accrual to {month}", f"${net:,.2f}",
              help=f"Accrue-in minus accrue-out. Positive = {month} revenue rises on a "
                   f"service-week basis; negative = it falls.")

    st.caption(
        f"{len(back_items)} line(s) booked in {month} belong to {prev_m}; "
        f"{len(in_items)} line(s) earned in {month} were booked in {next_m}. "
        f"These are bi-weekly invoices whose two weeks straddle the month boundary."
    )

    if back_items:
        with st.expander(f"Accrue back to {prev_m} — {len(back_items)} line(s), ${back_total:,.2f}", expanded=False):
            df = _accrual_detail_df(back_items)
            st.dataframe(df.style.format({"Amount": lambda v: f"${v:,.2f}"}),
                         hide_index=True, width="stretch", height=280)
            st.download_button(f"Download — {entity}_{month}_accrue_back.csv",
                               data=df.to_csv(index=False).encode("utf-8"),
                               file_name=f"{entity}_{month}_accrue_back_to_{prev_m}.csv",
                               mime="text/csv", key=f"{entity}_{month}_accr_back_dl")
    if in_items:
        with st.expander(f"Accrue in from {next_m} — {len(in_items)} line(s), ${in_total:,.2f}", expanded=False):
            df = _accrual_detail_df(in_items)
            st.dataframe(df.style.format({"Amount": lambda v: f"${v:,.2f}"}),
                         hide_index=True, width="stretch", height=280)
            st.download_button(f"Download — {entity}_{month}_accrue_in.csv",
                               data=df.to_csv(index=False).encode("utf-8"),
                               file_name=f"{entity}_{month}_accrue_in_from_{next_m}.csv",
                               mime="text/csv", key=f"{entity}_{month}_accr_in_dl")


def render_deferral_section(payload: dict, month: str, entity: str) -> None:
    """Show the period-end deferral: revenue invoices booked this month that
    belong (per accounting cutoff) to next month, so finance can book the JE."""
    if entity not in ENTITIES_WITH_DEFERRAL:
        return

    revenue_prefixes = PL_REVENUE_PREFIXES.get(entity, [])
    items, total = deferral_items(payload, month, revenue_prefixes)
    next_month = _next_month_label(month)

    st.markdown("### Period-end deferral")
    if not items:
        st.caption(
            f"No invoices booked in {month} fall in a cutoff window. "
            f"Nothing to defer into {next_month}."
        )
        return

    pl_booked = _pl_total(payload, month, revenue_prefixes)
    c1, c2, c3 = st.columns(3)
    c1.metric(
        f"Defer {month} → {next_month}", f"${total:,.2f}",
        help=f"Total revenue invoices booked (Transaction Date) in {month} whose "
             f"accounting period is {next_month}. Book a deferral JE for this amount.",
    )
    c2.metric(f"{month} P&L as booked", f"${pl_booked:,.2f}",
              help="Revenue booked by Transaction Date, before the deferral JE.")
    c3.metric(f"{month} P&L after deferral", f"${round(pl_booked - total, 2):,.2f}",
              help="What stays in this month once the deferral moves out.")

    st.caption(
        f"{len(items)} revenue invoice line(s) dated in {month}'s cutoff window "
        f"belong to {next_month}. Move this total with a deferral journal entry."
    )

    with st.expander(f"Show the {len(items)} invoice line(s) to defer", expanded=False):
        ddf = pd.DataFrame(items)[
            ["contractor", "txn_date", "num", "customer", "account", "amount", "description"]
        ].rename(columns={
            "contractor": "Contractor", "txn_date": "Txn date", "num": "Invoice #",
            "customer": "Customer", "account": "Account", "amount": "Amount",
            "description": "Description",
        })
        styled = ddf.style.format({"Amount": lambda v: f"${v:,.2f}"})
        st.dataframe(styled, hide_index=True, width="stretch", height=320)

        csv = ddf.to_csv(index=False).encode("utf-8")
        st.download_button(
            f"Download deferral detail · {entity}_{month}_deferral.csv",
            data=csv,
            file_name=f"{entity}_{month}_deferral.csv",
            mime="text/csv",
            key=f"{entity}_{month}_deferral_dl",
        )


def page_month_close(payload: dict, df: pd.DataFrame, entity: str) -> None:
    months = sorted(df["month"].dropna().unique(), reverse=True)
    if not months:
        st.warning("No data loaded.")
        return
    month = st.selectbox("Close period", options=months, key=f"{entity}_close_month")
    sub = df[df["month"] == month]

    log = load_close_log()
    closed = log.get(entity, {}).get(month)
    if closed:
        st.success(
            f"**Closed** by {closed['signer']} on {closed['signed_at']}. "
            f"(Session-scoped — re-download the snapshot below if you need a permanent record.)"
        )

    render_tiles(sub, payload=payload, entity=entity, pl_month=month)

    render_flush_section(payload, month, entity)

    render_accrual_section(payload, month, entity)

    render_deferral_section(payload, month, entity)

    st.markdown(f"### Reconciled rows — {ENTITIES[entity]['short']} · {month}")
    render_table(sub)
    st.markdown("### Drill-down")
    render_drill_down(sub, payload, prefix=f"{entity}_close")

    st.markdown("### Sign off and snapshot")
    with st.form(f"signoff_{entity}"):
        signer = st.text_input("Signer name", key=f"{entity}_close_signer")
        notes = st.text_area("Notes (optional)", height=80, key=f"{entity}_close_notes")
        submitted = st.form_submit_button("Sign off")
        if submitted:
            if not signer.strip():
                st.error("Signer name is required.")
            else:
                log.setdefault(entity, {})[month] = {
                    "signer": signer.strip(),
                    "signed_at": datetime.now().isoformat(timespec="seconds"),
                    "notes": notes.strip(),
                }
                save_close_log(log)
                st.success(f"{ENTITIES[entity]['short']} {month} signed off by {signer}. Download the snapshot Excel below for your permanent record.")

    # Generate snapshot in-memory and offer for download
    reconciled = payload.get("_reconciled")
    if reconciled:
        buf = io.BytesIO()
        write_excel_snapshot(reconciled, buf, period_label=month, entity_label=ENTITIES[entity]["short"])
        buf.seek(0)
        st.download_button(
            label=f"Download snapshot · audit_{entity}_{month}.xlsx",
            data=buf.getvalue(),
            file_name=f"audit_{entity}_{month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{entity}_close_download",
        )


def render_entity_tab(entity: str) -> None:
    cfg = ENTITIES[entity]
    short = cfg["short"]

    payload = get_entity_payload(entity)

    with st.expander(
        "📂 Upload source files" if payload is None else f"📂 Re-upload {short} files",
        expanded=(payload is None),
    ):
        render_upload_form(entity)

    if payload is None:
        st.info(f"Upload the two {short} source files above to see the dashboard.")
        return

    df = frame_payload(payload)

    st.caption(
        f"Engine ran: {payload.get('generated_at', '?')} · "
        f"{len(df):,} reconciled rows · "
        f"{payload.get('summary', {}).get('actionable_exceptions', 0)} actionable exceptions"
    )

    # Coverage window: the engine reconciles only where both exports overlap.
    # A subtle informational caption (no warning) shows the loaded date ranges.
    cov = payload.get("coverage") or {}
    if cov.get("qbo_min") and cov.get("ats_min"):
        st.caption(
            f"Coverage — QBO: {cov['qbo_min']} → {cov['qbo_max']} · "
            f"Bullhorn: {cov['ats_min']} → {cov['ats_max']}"
        )

    view = st.radio(
        "View", ["Continuous Audit", "Month Close"],
        horizontal=True, label_visibility="collapsed", key=f"{entity}_view",
    )
    if view == "Continuous Audit":
        page_continuous(payload, df, entity)
    else:
        page_month_close(payload, df, entity)


# -------------------------- main --------------------------

def main():
    st.set_page_config(page_title="OakTree Revenue Reconciler", layout="wide")

    with st.sidebar:
        st.markdown("## OakTree Reconciler")
        st.markdown("Audit Bullhorn ↔ QBO revenue for OTS and CTS.")
        st.markdown("---")

        for ent_key, cfg in ENTITIES.items():
            p = st.session_state.get(f"{ent_key}_payload")
            if p:
                when = p.get("generated_at", "?")
                actionable = p.get("summary", {}).get("actionable_exceptions", "?")
                st.markdown(f"**{cfg['short']}** · ran `{when[:19]}` · actionable **{actionable}**")
            else:
                st.markdown(f"**{cfg['short']}** · *no data — upload in tab*")

        st.markdown("---")
        st.caption(
            "Files you upload stay in your browser session only. "
            "Reload the page or click **Clear** in a tab to drop them."
        )

    tabs = st.tabs([ENTITIES[k]["label"] for k in ("ots", "cts")])
    for tab, entity in zip(tabs, ("ots", "cts")):
        with tab:
            render_entity_tab(entity)


if __name__ == "__main__":
    main()
