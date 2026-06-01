"""Tag each ReconciledRow with a status. Mutates rows in place.

Statuses:
  CLEAN                          diff == $0.00 and both sides present
  PENNY_ROUND                    0 < |diff| <= $0.01  (QBO/ATS rounding)
  MINOR_VARIANCE                 $0.01 < |diff| <= MATERIAL_THRESHOLD
  MATERIAL_VARIANCE              |diff| > MATERIAL_THRESHOLD
  ATS_ONLY_REVENUE_NOT_IN_QBO    ATS has rows, QBO has none, ATS sum != 0
  QBO_ONLY_NO_ATS_SOURCE         QBO has rows, ATS has none, QBO sum != 0
  ZERO_AMOUNT_NO_ACTION          One-sided AND that side sums to $0.00.
                                 Bullhorn-side example: contractor with hours
                                 but $0 bill rate (training/PTO/internal).
                                 QBO-side example: recon JE where debits +
                                 credits net to zero. Either way there's no
                                 invoice expected and no action needed.
  EXPENSE_RECLASS_NO_ACTION      QBO entry whose description matches a known
                                 expense-reclass pattern (e.g. "1220 RECON",
                                 "DPDIEM"). These are JEs moving expenses
                                 from Employee Advance to COGS — they appear
                                 in QBO without a Bullhorn billable, but
                                 they are reporting artifacts, not real
                                 discrepancies. Add patterns to
                                 EXPENSE_RECLASS_PATTERNS to extend.
"""
from __future__ import annotations
from match import ReconciledRow

PENNY_THRESHOLD = 0.01
MATERIAL_THRESHOLD = 1.00  # configurable in BUILD_PLAN section 10 #1
ZERO_TOLERANCE = 0.005     # anything that rounds to $0.00 at 2dp

# Substrings in QBO line descriptions that identify reporting-only entries
# (expense reclasses, recon JEs that double-count). Case-insensitive.
# Extend this tuple when new patterns are identified.
EXPENSE_RECLASS_PATTERNS = (
    "1220 RECON",   # JEs moving Employee Advance -> COGS
    "DPDIEM",       # per diem reclass code
)

STATUS_CLEAN = "CLEAN"
STATUS_PENNY = "PENNY_ROUND"
STATUS_MINOR = "MINOR_VARIANCE"
STATUS_MATERIAL = "MATERIAL_VARIANCE"
STATUS_ATS_ONLY = "ATS_ONLY_REVENUE_NOT_IN_QBO"
STATUS_QBO_ONLY = "QBO_ONLY_NO_ATS_SOURCE"
STATUS_ZERO_NO_ACTION = "ZERO_AMOUNT_NO_ACTION"
STATUS_EXPENSE_RECLASS = "EXPENSE_RECLASS_NO_ACTION"


def _matches_expense_reclass(r: ReconciledRow) -> bool:
    """True if any QBO line on this row carries a known reclass pattern."""
    for q in r.qbo_rows:
        desc = (q.description or "").upper()
        if any(p in desc for p in EXPENSE_RECLASS_PATTERNS):
            return True
    return False


def classify(rows: list[ReconciledRow]) -> None:
    for r in rows:
        # Reporting-only QBO entries take precedence — they aren't real
        # reconciliation exceptions regardless of which side has data.
        if _matches_expense_reclass(r):
            r.status = STATUS_EXPENSE_RECLASS
            continue

        if r.qbo_amount is None and r.ats_amount is not None:
            # Bullhorn-only. If the Bullhorn side nets to $0 there's nothing
            # to invoice and no QBO match will ever exist — not an exception.
            if abs(r.ats_amount) < ZERO_TOLERANCE:
                r.status = STATUS_ZERO_NO_ACTION
            else:
                r.status = STATUS_ATS_ONLY
        elif r.ats_amount is None and r.qbo_amount is not None:
            # QBO-only. If the QBO side nets to $0 it's a recon JE that
            # canceled itself out — also no action.
            if abs(r.qbo_amount) < ZERO_TOLERANCE:
                r.status = STATUS_ZERO_NO_ACTION
            else:
                r.status = STATUS_QBO_ONLY
        else:
            d = abs(r.diff)
            if d == 0:
                r.status = STATUS_CLEAN
            elif d <= PENNY_THRESHOLD:
                r.status = STATUS_PENNY
            elif d <= MATERIAL_THRESHOLD:
                r.status = STATUS_MINOR
            else:
                r.status = STATUS_MATERIAL


def status_counts(rows: list[ReconciledRow]) -> dict[str, int]:
    out: dict[str, int] = {}
    for r in rows:
        out[r.status] = out.get(r.status, 0) + 1
    return out
