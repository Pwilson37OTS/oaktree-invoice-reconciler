"""Load Bullhorn billing data from BULLHORN_MIRROR as ``BHRow`` objects.

Two source modes (see ``load_bullhorn_mirror``):

  "invoices" (DEFAULT) -- one BHRow per FINALIZED invoice line item
      (``RAW.InvoiceStatementLineItem`` on a finalized ``RAW.InvoiceStatement``).
      This is exactly what Bullhorn invoiced, which is what the engine reconciles
      against QBO, and it is immune to the mirror's junction-table sync lag.

  "charges" -- one BHRow per invoiced BILLABLE CHARGE
      (``RAW.BillableChargeInvoicedTransaction`` inner-joined to
      ``RAW.BillableCharge``). The original cutover source. Subject to the junction
      sync lag. Kept for validation / rollback via ``mode="charges"`` or env
      ``BULLHORN_MIRROR_MODE=charges``.

WHY FINALIZED INVOICE LINE ITEMS
--------------------------------
The engine reconciles what Bullhorn *invoiced* against what QBO booked. QBO holds
finalized invoices, so the faithful Bullhorn side is the finalized invoice line
items. Summing them per candidate+week naturally:
  - nets credit memos (their line amounts are negative) against the originals, and
  - counts a re-issued (rebill) week once -- it includes the FINALIZED rebill and
    excludes the superseded DRAFT, because drafts are not finalized (and never
    reach QBO either).

WHY NOT THE CHARGE->INVOICE JUNCTION FOR THE AMOUNT
---------------------------------------------------
The charge path sums ``BillableChargeInvoicedTransaction`` (the per-day
charge->invoice link). CData replicates the parent ``InvoiceStatement`` + line
items promptly, but those junction rows (and ``BillableCharge."Invoiced?"``) land
~1-1.5 weeks later, so a freshly finalized invoice looks un-billed for a week-plus
-- understating Bullhorn and manufacturing phantom MATERIAL_VARIANCE / QBO_ONLY
exceptions. Measured 2026-10-30: 222 finalized invoices / ~$1.10M were in
``InvoiceStatement`` but missing from the junction, ALL created in the prior ~10
days. The junction also counts linked DRAFT rebills (not yet finalized, not in
QBO), which the finalized-line-item sum correctly ignores. Example: Peter Landry,
service week 2026-10-18 -- invoice 25243 ($1,575 Per Diem) was linked, invoice
25388 ($11,308.37 Regular+OT, created 10/21) had zero junction rows; the charge
path showed $1,575 vs QBO's $12,883.37, the line-item sum shows $12,883.37.

RESOLVING candidate_id  (precise where possible, name only where necessary)
-------------------------------------------------------------------------
The match key is ``candidate_id|service_date`` on both sides. QBO carries the id
as the leading comma-token of each description ("185245, Peter Landry - ..."), but
line items carry only the contractor NAME. We resolve each line item's candidate_id
by, in order:
  1. a JUNCTION ORACLE -- if the line's invoice is linked in the junction and all
     its linked charges belong to ONE candidate (6,372 of 6,475 linked invoices as
     of 2026-10-30), that candidate_id is used directly. No name guessing, so
     settled invoices (incl. middle-initial / suffix / nickname / name-change
     contractors) resolve exactly.
  2. NAME resolution from ``RAW.Candidate`` -- only for invoices the junction has
     not linked yet (the lag tail) and the handful of multi-candidate invoices
     (whose line items do carry names). The id is matched on two normalized keys
     per candidate -- full name, and first+last token (middle names, initials and
     Jr/Sr/III suffixes stripped) -- built from BOTH ``Name`` and ``FirstName`` +
     ``LastName`` (they can differ, e.g. "Jimmy" vs "James"). Keys mapping to more
     than one candidate are dropped rather than mis-attributed. (A legal name
     CHANGE can't be resolved by name, but those invoices are essentially always
     linked, so the oracle handles them before name resolution is reached.)
Branch (for entity selection) is taken per candidate from ``BillableCharge`` and
is unambiguous (no candidate's charges span two branches).

ENTITY SELECTION (reproduces Brenda's Canvas report sets)
  OTS = candidate Branch 'OTS'.
  CTS = candidate Branch 'CTS'  OR  (Branch 'OTS' AND the invoice's client is an
        inter-company client: CITGO / ClientCorporation 1462). CITGO is staffed by
        CTS but billed on OTS paper via CPL2 / "Interco Rec-OTS".

TAX
  Services are not taxed here (``Pre-taxAmount`` == ``TotalAmount`` on every
  finalized invoice), so ``Pre-taxAmount`` is the billed figure.

Display-only fields (hours, bill/pay rate, job title, GM, state, employee type)
are not sourced; they do not affect matching (candidate + date + amount).
"""
from __future__ import annotations

import os
import re
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

# Default Bullhorn source mode. Override with env BULLHORN_MIRROR_MODE.
_DEFAULT_MODE = "invoices"

# Name suffixes stripped before building the first+last match key.
_NAME_SUFFIXES = {"jr", "sr", "ii", "iii", "iv", "v"}


# --------------------------------------------------------------------------- #
# Small coercers (shared by all paths)
# --------------------------------------------------------------------------- #
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


def _to_int(v) -> int | None:
    try:
        if v is None or v == "":
            return None
        return int(float(v))
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------- #
# Lookups for the invoice-line-item path
# --------------------------------------------------------------------------- #
def _name_keys(raw: str) -> list[str]:
    """Ordered normalized match keys for a person name: full name first, then
    first+last token. Punctuation dropped, whitespace collapsed, case folded,
    Jr/Sr/III-style suffixes removed. Bullhorn renders names with a double space
    ("Peter  Landry"), which collapses here.

    Examples:
      "Peter  Landry"       -> ["peter landry"]
      "James E. Wingerter"  -> ["james e wingerter", "james wingerter"]
      "Loren Estrada Jr."   -> ["loren estrada"]
    """
    toks = [t for t in re.split(r"\s+", re.sub(r"[.,]", " ", raw or "")) if t]
    toks = [t for t in toks if t.lower() not in _NAME_SUFFIXES]
    if not toks:
        return []
    full = " ".join(toks).casefold()
    keys = [full]
    if len(toks) >= 2:
        fl = f"{toks[0]} {toks[-1]}".casefold()
        if fl != full:
            keys.append(fl)
    return keys


# Candidate lookups, restricted to candidates that appear in BillableCharge (the
# only ones that can be invoiced): name key -> id, id -> branch, id -> placement.
_MAPS_SQL = """
SELECT c."ID"        AS id,
       c."Name"      AS name,
       c."FirstName" AS fn,
       c."LastName"  AS ln,
       bx.branch     AS branch,
       bx.placement  AS placement
FROM BULLHORN_MIRROR."RAW"."Candidate" c
JOIN (
    SELECT "Candidateid"     AS cid,
           MAX("Branch")      AS branch,
           MAX("Placementid") AS placement
    FROM BULLHORN_MIRROR."RAW"."BillableCharge"
    WHERE "Candidateid" IS NOT NULL
    GROUP BY "Candidateid"
) bx ON bx.cid = c."ID"
"""

# Junction oracle: invoice -> candidate_id, ONLY for invoices whose linked charges
# all belong to one candidate (the overwhelming majority). Bounded to the pull
# window. Gives a precise id for settled invoices without touching names.
_ORACLE_SQL = """
SELECT it."Invoiceid" AS inv_id, MIN(bc."Candidateid") AS cand_id
FROM BULLHORN_MIRROR."RAW"."BillableChargeInvoicedTransaction" it
JOIN BULLHORN_MIRROR."RAW"."BillableCharge" bc ON bc."ID" = it."ChargeID"
JOIN BULLHORN_MIRROR."RAW"."InvoiceStatement" inv ON inv."ID" = it."Invoiceid"
WHERE bc."Candidateid" IS NOT NULL
  AND inv."InvoiceStatementDate" >= ?
GROUP BY it."Invoiceid"
HAVING COUNT(DISTINCT bc."Candidateid") = 1
"""

# Finalized invoice line items since a cutoff. The match/service date is parsed
# from the line Description (not InvoiceStatementDate), so back-dated invoices
# land in the right service week.
_LINEITEM_SQL = """
SELECT li."ID"                   AS line_id,
       li."InvoiceStatementid"   AS invoice_id,
       li."Description"          AS descr,
       li."Pre-taxAmount"        AS amount,
       inv."ClientCorporationid" AS client_id,
       cc."CompanyName"          AS client
FROM BULLHORN_MIRROR."RAW"."InvoiceStatementLineItem" li
JOIN BULLHORN_MIRROR."RAW"."InvoiceStatement" inv
       ON inv."ID" = li."InvoiceStatementid"
LEFT JOIN BULLHORN_MIRROR."RAW"."ClientCorporation" cc
       ON cc."ID" = inv."ClientCorporationid"
WHERE inv."Finalized?" = TRUE
  AND inv."InvoiceStatementDate" >= ?
ORDER BY inv."InvoiceStatementDate", li."ID"
"""


def _build_maps() -> tuple[dict[str, str], dict[str, str], dict[str, str]]:
    """Return (name_key -> candidate_id, candidate_id -> branch,
    candidate_id -> placement_id). A name key that maps to more than one
    candidate_id is ambiguous and dropped (left unresolved, never mis-attributed)."""
    df = mirror.query(_MAPS_SQL)

    key_to_id: dict[str, str] = {}
    collided: set[str] = set()
    branch_by_id: dict[str, str] = {}
    placement_by_id: dict[str, str] = {}

    for rec in df.itertuples(index=False):
        cid = _to_str(rec.ID)
        if not cid:
            continue
        branch_by_id[cid] = _to_str(rec.BRANCH)
        placement_by_id[cid] = _to_str(rec.PLACEMENT)

        keys = set(_name_keys(_to_str(rec.NAME)))
        keys |= set(_name_keys(f"{_to_str(rec.FN)} {_to_str(rec.LN)}"))
        for k in keys:
            if k in key_to_id and key_to_id[k] != cid:
                collided.add(k)
            else:
                key_to_id[k] = cid

    for k in collided:
        key_to_id.pop(k, None)

    return key_to_id, branch_by_id, placement_by_id


def _invoice_candidate_oracle(since: date) -> dict[str, str]:
    """invoice_id -> candidate_id for single-candidate linked invoices."""
    df = mirror.query(_ORACLE_SQL, [since.isoformat()])
    return {_to_str(rec.INV_ID): _to_str(rec.CAND_ID) for rec in df.itertuples(index=False)}


def _resolve_name(line_name: str, key_to_id: dict[str, str]) -> str | None:
    for k in _name_keys(line_name):
        cid = key_to_id.get(k)
        if cid:
            return cid
    return None


def _lineitem_name(descr: str) -> str:
    """Candidate name from a line Description like
    "Peter  Landry - Regular - 2026-10-18" -> "Peter  Landry": everything before
    the first " - " segment."""
    s = (descr or "").strip()
    if " - " in s:
        return s.split(" - ", 1)[0].strip()
    return s


def _lineitem_date(descr: str) -> date | None:
    """Service-week date from a line Description: the rightmost " - " segment that
    parses as a date. Handles the plain form ("... - 2026-10-18") and the extended
    form ("... - 2026-10-18 - El Paso - TX - 79930")."""
    for part in reversed((descr or "").split(" - ")):
        d = _to_date(part.strip())
        if d:
            return d
    return None


def _keep_for_entity(entity: str, branch: str, client_id: int | None) -> bool:
    """Apply the Canvas entity selection to one line (candidate `branch` + the
    invoice's `client_id`)."""
    if entity == "ots":
        return branch == "OTS"
    if entity == "cts":
        interco = _INTERCO_CLIENT_IDS.get("cts", [])
        return branch == "CTS" or (branch == "OTS" and client_id in interco)
    raise ValueError(f"Unknown entity {entity!r}")


def _load_from_invoices(entity: str, *, apply_period_remap: bool, since: date) -> list[BHRow]:
    key_to_id, branch_by_id, placement_by_id = _build_maps()
    oracle = _invoice_candidate_oracle(since)

    df = mirror.query(_LINEITEM_SQL, [since.isoformat()])

    rows: list[BHRow] = []
    for rec in df.itertuples(index=False):
        descr = _to_str(rec.DESCR)
        sdate = _lineitem_date(descr)
        if not sdate:
            continue  # no service date -> cannot form a match key
        # Precise id from the junction oracle; fall back to name resolution.
        cand_id = oracle.get(_to_str(rec.INVOICE_ID)) or _resolve_name(
            _lineitem_name(descr), key_to_id
        )
        if not cand_id:
            continue  # unresolved (or ambiguous) name on an unlinked invoice
        branch = branch_by_id.get(cand_id, "")
        if not _keep_for_entity(entity, branch, _to_int(rec.CLIENT_ID)):
            continue
        rows.append(
            BHRow(
                row_num=int(rec.LINE_ID),
                invoice_date=sdate,
                candidate_id=cand_id,
                candidate_name=_lineitem_name(descr),
                client=_to_str(rec.CLIENT),
                state="",
                hours=0.0,
                bill_rate=0.0,
                gross_sales_amount=round(_to_float(rec.AMOUNT), 2),
                pay_rate=0.0,
                job_title="",
                gross_margin=0.0,
                placement_id=placement_by_id.get(cand_id, ""),
                period_end=sdate,
                branch=_BRANCH2.get(branch, branch),
                employee_type="",
                apply_period_remap=apply_period_remap,
            )
        )
    return rows


# --------------------------------------------------------------------------- #
# BillableCharge path (legacy / validation / rollback)
# --------------------------------------------------------------------------- #
# The `invoiced` CTE sums each charge's linked invoiced-transaction amounts
# (netting credit/rebill lines); the inner join to it both restricts to invoiced
# charges and supplies the billed amount. invoiced_only=False keeps every charge
# at its raw Pre-taxAmount (left join) for inspecting the un-invoiced backlog.
# NOTE: subject to the junction sync lag -- prefer the default "invoices" mode for
# anything that must tie to QBO.
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
    """SQL predicate + qmark params selecting the charges that belong to `entity`.

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
    """Charge Description is "{Candidate Name} - {YYYY-MM-DD}" (occasionally with
    further " - " segments). The name is everything before the last segment; fall
    back to the whole string."""
    s = (descr or "").strip()
    if " - " in s:
        return s.rsplit(" - ", 1)[0].strip()
    return s


def _load_from_charges(
    entity: str, *, apply_period_remap: bool, since: date, invoiced_only: bool,
) -> list[BHRow]:
    pred, pred_params = _entity_predicate(entity)
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


# --------------------------------------------------------------------------- #
# Public entry point
# --------------------------------------------------------------------------- #
def load_bullhorn_mirror(
    entity: str,
    *,
    apply_period_remap: bool = True,
    since: date | None = None,
    invoiced_only: bool = True,
    mode: str | None = None,
) -> list[BHRow]:
    """Bullhorn billing for `entity` ('ots'/'cts') from the live mirror, as BHRow
    objects. `since` bounds the pull (default ~220 days back).

    mode (explicit arg > env BULLHORN_MIRROR_MODE > default "invoices"):
      "invoices" -- finalized invoice line items; candidate_id via the junction
                    oracle, name only for unlinked/multi-candidate invoices.
                    Immune to the junction sync lag. This is what reconciles to QBO.
      "charges"  -- invoiced charges only. With invoiced_only=True (default)
                    restrict to junction-linked charges; False keeps every charge at
                    raw Pre-taxAmount (un-invoiced backlog). Subject to the lag.
    """
    if since is None:
        since = date.today() - timedelta(days=_DEFAULT_LOOKBACK_DAYS)

    resolved = (mode or os.environ.get("BULLHORN_MIRROR_MODE") or _DEFAULT_MODE).strip().lower()

    if resolved in ("charge", "charges", "junction"):
        return _load_from_charges(
            entity, apply_period_remap=apply_period_remap, since=since,
            invoiced_only=invoiced_only,
        )
    if resolved in ("invoice", "invoices", "lineitem", "lineitems"):
        return _load_from_invoices(
            entity, apply_period_remap=apply_period_remap, since=since,
        )
    raise ValueError(
        f"Unknown BULLHORN_MIRROR_MODE {resolved!r} (want 'invoices' or 'charges')"
    )
