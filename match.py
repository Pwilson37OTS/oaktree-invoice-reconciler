"""Aggregate + full-outer-join of QBO and Bullhorn rows on `candidate_id|YYYY-MM-DD`.

Mirrors the VBA logic: per-key sum on both sides, then `diff = ats - qbo`
rounded to 2 decimal places.

Each ReconciledRow also carries the contributing source rows so the dashboard
can show drill-down.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from datetime import date
from collections import defaultdict

from parse_qbo import QBORow
from parse_bullhorn import BHRow


@dataclass
class ReconciledRow:
    key: str
    candidate_id: str
    billing_date: date
    contractor_name: str

    qbo_amount: float | None
    ats_amount: float | None

    qbo_rows: list[QBORow] = field(default_factory=list)
    ats_rows: list[BHRow] = field(default_factory=list)

    # Convenience context fields (populated from the most informative source row).
    client: str = ""
    placement_id: str = ""
    job_title: str = ""

    status: str = ""  # filled in by classify()

    @property
    def diff(self) -> float:
        """ATS - QBO, rounded 2dp. Missing sides count as 0."""
        return round((self.ats_amount or 0.0) - (self.qbo_amount or 0.0), 2)


def reconcile(qbo_rows: list[QBORow], ats_rows: list[BHRow]) -> list[ReconciledRow]:
    """Build the union of keys and produce one ReconciledRow per key."""
    qbo_by_key: dict[str, list[QBORow]] = defaultdict(list)
    ats_by_key: dict[str, list[BHRow]] = defaultdict(list)

    for r in qbo_rows:
        if r.key:
            qbo_by_key[r.key].append(r)
    for r in ats_rows:
        if r.key:
            ats_by_key[r.key].append(r)

    all_keys = set(qbo_by_key) | set(ats_by_key)
    out: list[ReconciledRow] = []

    for key in all_keys:
        qbo = qbo_by_key.get(key, [])
        ats = ats_by_key.get(key, [])

        candidate_id, date_str = key.split("|", 1)
        billing_date = date.fromisoformat(date_str)

        # Contractor name: prefer Bullhorn (canonical), fall back to QBO description.
        contractor_name = ats[0].candidate_name if ats else _name_from_qbo(qbo)
        client = ats[0].client if ats else (qbo[0].customer if qbo else "")
        placement_id = ats[0].placement_id if ats else ""
        job_title = ats[0].job_title if ats else ""

        qbo_total = round(sum(r.amount for r in qbo), 2) if qbo else None
        ats_total = round(sum(r.gross_sales_amount for r in ats), 2) if ats else None

        out.append(
            ReconciledRow(
                key=key,
                candidate_id=candidate_id,
                billing_date=billing_date,
                contractor_name=contractor_name,
                qbo_amount=qbo_total,
                ats_amount=ats_total,
                qbo_rows=qbo,
                ats_rows=ats,
                client=client,
                placement_id=placement_id,
                job_title=job_title,
            )
        )

    # Sort like the VBA: contractor name asc, then billing date asc.
    out.sort(key=lambda r: (r.contractor_name.casefold(), r.billing_date))
    return out


def _name_from_qbo(qbo_rows: list[QBORow]) -> str:
    """Extract contractor name from a QBO description like
    '58363, Angela Hardin - Expense - 2026-04-03'."""
    if not qbo_rows:
        return ""
    desc = qbo_rows[0].description
    if "," in desc:
        after_comma = desc.split(",", 1)[1]
        if " - " in after_comma:
            return after_comma.split(" - ", 1)[0].strip()
        return after_comma.strip()
    return desc.strip()
