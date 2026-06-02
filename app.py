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


def render_tiles(df: pd.DataFrame, payload: dict | None = None, entity: str | None = None) -> None:
    qbo_total = round(df["qbo_amount"].fillna(0).sum(), 2)
    ats_total = round(df["ats_amount"].fillna(0).sum(), 2)
    net = round(ats_total - qbo_total, 2)
    exceptions = (~df["status"].isin(NON_ACTIONABLE)).sum() if not df.empty else 0
    material = (df["status"] == "MATERIAL_VARIANCE").sum() if not df.empty else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Bullhorn billed", f"${ats_total:,.2f}")
    c2.metric("QBO revenue", f"${qbo_total:,.2f}")
    c3.metric("Net variance", f"${net:,.2f}",
              delta=None if abs(net) < 0.01 else f"{net:+,.2f}",
              delta_color="inverse" if abs(net) >= 0.01 else "off")
    c4.metric("Exceptions", int(exceptions))
    c5.metric("Material variances", int(material), delta_color="inverse")

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
    render_tiles(filtered, payload=payload, entity=entity)
    st.markdown("### Reconciled rows")
    render_table(filtered)
    st.markdown("### Drill-down")
    render_drill_down(filtered, payload, prefix=f"{entity}_cont")


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

    render_tiles(sub, payload=payload, entity=entity)
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
