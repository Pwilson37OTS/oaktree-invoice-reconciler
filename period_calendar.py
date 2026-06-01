"""Period-end remapping for OakTree's accounting cutoff dates.

When a billing date falls in a transitional window around month-end,
OakTree's books anchor it to a specific date in the destination accounting
period. Same logic Brenda has in her SQL CASE statement on the Invoice
Statement Date — kept here so the reconciler matches across the cutoff.

To extend (e.g. next fiscal year, or a rule change), append tuples to
PERIOD_REMAPS. Each tuple is (window_start, window_end_inclusive, anchor).
"""
from __future__ import annotations
from datetime import date


PERIOD_REMAPS: list[tuple[date, date, date]] = [
    (date(2026,  1, 26), date(2026,  2,  2), date(2026,  1, 31)),
    (date(2026,  2, 25), date(2026,  3,  2), date(2026,  2, 28)),
    (date(2026,  5, 28), date(2026,  5, 31), date(2026,  6,  1)),
    (date(2026,  6, 26), date(2026,  6, 28), date(2026,  7,  1)),
    (date(2026,  7, 31), date(2026,  8,  2), date(2026,  7, 31)),
    (date(2026, 10, 28), date(2026, 11,  1), date(2026, 11,  1)),
]


def to_period_anchor(d: date | None) -> date | None:
    """Remap a billing date to its accounting-period anchor.
    Returns the date unchanged when no transition rule applies."""
    if d is None:
        return None
    for lo, hi, anchor in PERIOD_REMAPS:
        if lo <= d <= hi:
            return anchor
    return d


def is_remapped(d: date | None) -> bool:
    """True if `d` falls inside any transition window (i.e. would be remapped)."""
    if d is None:
        return False
    return any(lo <= d <= hi for lo, hi, _ in PERIOD_REMAPS)
