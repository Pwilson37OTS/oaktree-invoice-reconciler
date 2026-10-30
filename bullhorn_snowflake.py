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

Entity selection (reproduces Brenda's Canvas report sets, see _entity_predicate):
  OTS = Branch='OTS' (this already includes the CITGO inter-company contractors;
        they appear in billable_charges.xlsx too and are flushed to CTS via CPL2
        downstream).
  CTS = Branch='CTS'  OR  (Branch='OTS' AND ClientCorporationid = 1462 / CITGO).
        The second term is the inter-company set: CTS contractors billed on OTS
        paper (CPL2 / "Interco Rec-OTS"). In the Canvas export these carry the
        Branch2='1' tag. Validated 2026-10-29: the mirror's (Branch='OTS',
        client=CITGO) set is byte-for-byte the 14 Branch2='1' contractors in
        billable_charges_CTS.xlsx, with matching per-contractor gross.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta

import mirror
from parse_bullhorn import BHRow

# Clients whose OTS-billed charges belong to the other entity (inter-company).
# CITGO (ClientCorporation 1462) is staffed by CTS but billed on OTS paper via
# CPL2 / "Interco Rec-OTS", so its OTS charges roll into the CTS report.
_INTERCO_CLIENT_IDS = {"cts": [1462]}

# Canvas "Branch2" tag the drill-down shows, so the mirror matches the export:
# 1 = billed on OTS paper, 2 = billed by CTS.
_BRANCH2 = {"OTS": "1", "CTS": "2"}

# Pull a bounded history by default so the reconciler's date range roughly
# matches the Canvas export window; the engine's coverage clip aligns it to QBO.
_DEFAULT_LOOKBACK_DAYS = 220

# This engine reconciles INVOICES against QBO, not raw billable charges. A
# billable charge exists as soon as time is approved, but QBO only holds what was
# actually invoiced, so a never-invoiced charge has no QBO counterpart and
# manufactures a phantom ATS-only exception. This matters most for flat-rate
# clients (e.g. ConocoPhillips): the weekly billable charges are never invoiced
# at all -- the client is billed one flat monthly charge -- so summing the weekly
# charges both double-counts and mis-states the invoice. Example (verified
# 2026-10-30): Jose Guerrero's weekly charges ($4,191 / $8,543 / $2,901) carry
# zero invoiced transactions, while a single flat $20,220 charge is the actual
# invoice.
#
# So the invoiced universe (and the billed amount) comes from
# BillableChargeInvoicedTransaction, the charge->invoice link -- NOT from
# BillableCharge. We do NOT use the "Invoiced?" boolean: it is stale in the mirror
# (Aug 2026: 199 of 213 OTS charges flagged false actually have invoiced
# transactions). The `invoiced` CTE sums each charge's linked transaction amounts
# (netting any credit / rebill lines -- ~1 charge in 500 nets to something other
# than its Pre-taxAmount), and the inner join to it both restricts to invoiced
# charges and supplies the billed amount. The Canvas export did the equivalent --
# its "Invoice Date" column (== Billing Period End Date on every row) is only
# populated for invoiced charges, and the loader skips blank-date rows.
#
# invoiced_only=False keeps every billable charge at its raw Pre-taxAmount (a
# left join, amount off BillableCharge) for inspecting the un-invoiced backlog.
_SQL_TEMPLATE = """
WITH invoiced AS (
    SELECT it."ChargeID" AS charge_id, SUM(it."Amount") AS invoiced_amount
    FROM BULLHORN_MIRROR."RAW"."BillableChargeInvoicedTransaction" it
    GROUP BY it."ChargeID"
)
SELECT bc."ID"            AS charge_id,
       bc."Candidateid"   AS candidate_id,
       bc."Description"    AS descr,
       cc."CompanyName"    AS client,
       bc."PeriodEndDate"  AS period_end,
       {amount_col}        AS gross_sales,
       bc."Placementid"    AS placement_id,
       bc."Branch"         AS branch
FROM BULLHORN_MIRROR."RAW"."BillableCharge" bc
{invoiced_join} invoiced inv ON inv.charge_id = bc."ID"
LEFT JOIN BULLHORN_MIRROR."RAW"."ClientCorporation" cc
       ON cc."ID" = bc."ClientCorporationid"
WHERE {entity_pred}
  AND bc."Candidateid" IS NOT NULL
  AND bc."PeriodEndDate" IS NOT NULL
  AND bc."PeriodEndDate" >= ?
ORDER BY bc."PeriodEndDate", bc."Candidateid"
"""

# (invoiced_join, amount_col) per mode.
_INVOICED_MODE = ("JOIN", "inv.invoiced_amount")       # billed amount, invoiced only
_RAW_MODE = ("LEFT JOIN", 'bc."Pre-taxAmount"')        # every charge, raw amount


def _entity_predicate(entity: str) -> tuple[str, list]:
    """SQL predicate + qmark params selecting the charges that belong to `entity`,
    reproducing Brenda's Canvas report selection (see module docstring).

    OTS is simply Branch='OTS'. CTS is Branch='CTS' plus the inter-company set:
    OTS-billed charges for an inter-company client (CITGO). Literals here are
    trusted internal constants, not user input."""
    if entity == "ots":
        return 'bc."Branch" = ?', ["OTS"]
    if entity == "cts":
        ids = _INTERCO_CLIENT_IDS["cts"]
        placeholders = ",".join("?" for _ in ids)
        pred = (
            '(bc."Branch" = ? '
            f'OR (bc."Branch" = ? AND bc."ClientCorporationid" IN ({placeholders})))'
        )
        return pred, ["CTS", "OTS", *ids]
    raise ValueError(f"Unknown entity {entity!r}")


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
    invoiced_only: bool = True,
) -> list[BHRow]:
    """Invoiced charges for `entity` ('ots'/'cts') from the live mirror, as
    BHRow objects. `since` bounds the pull (default ~220 days back).

    invoiced_only (default True) restricts to charges Bullhorn has actually
    invoiced, so the set is comparable to QBO; pass False only to inspect the
    raw billable-charge universe (e.g. un-invoiced backlog)."""
    pred, pred_params = _entity_predicate(entity)
    if since is None:
        since = date.today() - timedelta(days=_DEFAULT_LOOKBACK_DAYS)

    invoiced_join, amount_col = _INVOICED_MODE if invoiced_only else _RAW_MODE
    sql = _SQL_TEMPLATE.format(
        entity_pred=pred, invoiced_join=invoiced_join, amount_col=amount_col,
    )
    df = mirror.query(sql, [*pred_params, since.isoformat()])

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
                branch=_BRANCH2.get(_to_str(rec.BRANCH), _to_str(rec.BRANCH)),
                employee_type="",
                apply_period_remap=apply_period_remap,
            )
        )
    return rows
