"""Load Bullhorn billable charges from BULLHORN_MIRROR, returning the same
``BHRow`` shape the xlsx parser produces so the reconcile engine is unchanged.

Source: ``RAW.BillableCharge`` (one row per billable charge), joined to
``ClientCorporation`` for the client name. Validated against the Canvas export
for OTS: monthly totals tie within ~1% (the mirror is slightly more complete,
which is the whole reason for the cutover), and per-contractor weekly amounts
match to the penny on ``PeriodEndDate`` + ``Pre-taxAmount``.

Field mapping (xlsx column -> mirror source):
  Invoice Date         -> PeriodEndDate            (the match-key week date)
  Candidate ID         -> Candidateid
  Candidate Name       -> parsed from Description ("Name - YYYY-MM-DD")
  Client Company Name  -> ClientCorporation.CompanyName
  Gross Sales Amount   -> Pre-taxAmount
  Placement ID         -> Placementid
  Billing Period End   -> PeriodEndDate
  Branch (OTS/CTS)     -> Branch

Not yet sourced (display-only in the drill-down, not used for matching): hours,
bill/pay rate, job title, gross margin, work state, employee type. These need
Timesheet / Placement / JobOrder joins and can be enriched later; leaving them
blank does not affect the reconciliation, which keys on candidate + date +
amount.

CTS attribution is intentionally NOT handled here yet: some CTS contractors are
billed inter-company through OTS (CPL2 / "Interco Rec-OTS") and appear under
Branch='OTS' in the mirror while the Canvas report lists them under CTS. That
inclusion rule is not derivable from the charge/placement/candidate fields and
is pending Brenda's Canvas report definition. OTS is clean (Branch='OTS').
"""
from __future__ import annotations

from datetime import date, datetime, timedelta

import mirror
from parse_bullhorn import BHRow

_BRANCH = {"ots": "OTS", "cts": "CTS"}

# Pull a bounded history by default so the reconciler's date range roughly
# matches the Canvas export window; the engine's coverage clip aligns it to QBO.
_DEFAULT_LOOKBACK_DAYS = 220

_SQL = """
SELECT bc."ID"            AS charge_id,
       bc."Candidateid"   AS candidate_id,
       bc."Description"    AS descr,
       cc."CompanyName"    AS client,
       bc."PeriodEndDate"  AS period_end,
       bc."Pre-taxAmount"  AS gross_sales,
       bc."Placementid"    AS placement_id,
       bc."Branch"         AS branch
FROM BULLHORN_MIRROR."RAW"."BillableCharge" bc
LEFT JOIN BULLHORN_MIRROR."RAW"."ClientCorporation" cc
       ON cc."ID" = bc."ClientCorporationid"
WHERE bc."Branch" = ?
  AND bc."Candidateid" IS NOT NULL
  AND bc."PeriodEndDate" IS NOT NULL
  AND bc."PeriodEndDate" >= ?
ORDER BY bc."PeriodEndDate", bc."Candidateid"
"""


def _candidate_name(descr: str) -> str:
    """Description is "{Candidate Name} - {YYYY-MM-DD}" (and occasionally has
    further " - " segments). The name is everything before the first date-like
    trailing segment; fall back to the whole string."""
    s = (descr or "").strip()
    if " - " in s:
        return s.rsplit(" - ", 1)[0].strip()
    return s


def _to_date(v) -> date | None:
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def _to_float(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _to_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def load_bullhorn_mirror(
    entity: str,
    *,
    apply_period_remap: bool = True,
    since: date | None = None,
) -> list[BHRow]:
    """Billable charges for `entity` ('ots'/'cts') from the live mirror, as
    BHRow objects. `since` bounds the pull (default ~220 days back)."""
    branch = _BRANCH.get(entity)
    if branch is None:
        raise ValueError(f"Unknown entity {entity!r}")
    if since is None:
        since = date.today() - timedelta(days=_DEFAULT_LOOKBACK_DAYS)

    df = mirror.query(_SQL, [branch, since.isoformat()])

    rows: list[BHRow] = []
    for rec in df.itertuples(index=False):
        ped = _to_date(rec.PERIOD_END)
        cand_id = _to_str(rec.CANDIDATE_ID)
        if not ped or not cand_id:
            continue
        rows.append(
            BHRow(
                row_num=int(rec.CHARGE_ID),
                invoice_date=ped,
                candidate_id=cand_id,
                candidate_name=_candidate_name(rec.DESCR),
                client=_to_str(rec.CLIENT),
                state="",
                hours=0.0,
                bill_rate=0.0,
                gross_sales_amount=round(_to_float(rec.GROSS_SALES), 2),
                pay_rate=0.0,
                job_title="",
                gross_margin=0.0,
                placement_id=_to_str(rec.PLACEMENT_ID),
                period_end=ped,
                branch=_to_str(rec.BRANCH),
                employee_type="",
                apply_period_remap=apply_period_remap,
            )
        )
    return rows
