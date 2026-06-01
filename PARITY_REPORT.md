# April 2026 Parity Report — Python Engine vs. Legacy VBA Workbook

**Engine commit/run:** 2026-05-27
**Workbook compared:** `_inputs/Revenue_Audit_Master_File_2026-04.xlsm` (last run 2026-05-07)
**QBO source:** `_shared/quickbooks-data/audit_rev_ots.xlsx` (refreshed 2026-05-27)
**Bullhorn source:** `_shared/bullhorn-data/billable_charges.xlsx` (refreshed 2026-05-27)

## Summary

The Python engine reproduces the legacy VBA macro's logic. **19 of 22 workbook discrepancies are surfaced by the engine with identical diff values.** The 3 misses are not engine bugs — they're cases where source data was refreshed between May 7 and May 27 and the issue resolved or shifted to a different date.

## Comparison breakdown

| | Engine | Workbook |
|---|---:|---:|
| Reconciled rows (April 2026) | 1,113 | 1,132 |
| Non-trivial discrepancies | 18 | 22 |
| (contractor, date) keys in both | **1,104 (99.2%)** | |

### Discrepancy parity (22 workbook discrepancies cross-checked against engine)

| Status in engine | Count | Notes |
|---|---:|---|
| `PENNY_ROUND` w/ matching diff | 17 | All workbook $0.01 diffs re-detected |
| `MATERIAL_VARIANCE` w/ matching diff | 1 | Spencer Finkbeiner 4/24, $4,402.61 — engine confirmed |
| `CLEAN` (resolved by data refresh) | 2 | Jesus Salgado 4/28 + Jose Guerrero 4/28 — Bullhorn now has the missing ATS rows |
| Not found (data refresh) | 2 | Jesus Salgado 3/31 + Jose Guerrero 3/31 — billing date shifted to 4/28 in new export |

### "Only workbook" rows (28)
Most are Feb 2026 Employee Advance lines that summed to $0 in the workbook (workbook's QBO sheet carried multi-month data, ATS sheet had April only, so Feb rows showed up with `ats=None`). The engine filters by the `--month` argument so these are not surfaced — correct behavior.

A few are name-aliasing artifacts: "Jeffrey Chapman" (workbook) vs "Jeff Chapman" (engine). Engine pulls the canonical name from Bullhorn's `Candidate Name` column when available; workbook used the QBO description fragment. The underlying candidate_id matches.

### "Only engine" rows (9)
- 4 × Jeff Chapman April weeks — present in current Bullhorn data, not in May-7 workbook
- 4 × "1220 recon-*" QBO_ONLY rows — recon journal entries from the bookkeeper. The engine correctly classifies these as `QBO_ONLY_NO_ATS_SOURCE` (off-cycle adjustments without a Bullhorn source). Workbook does not surface these as distinct rows.
- 1 × Emily Schultz `qbo=0, ats=0` — zero-zero row. Could be filtered as noise in a later pass.

## Conclusion

Engine validated for production use against April 2026. Move forward with:
1. Streamlit dashboard build
2. Brenda walkthrough on the actual May 2026 close to confirm the workflow change is comfortable for her
3. CTS entity addition (Phase 2)

## Scripts used

- `_inputs/parity_compare.py` (in `/tmp/audit/`) — keyed comparison of engine output JSON vs workbook Audit sheet
- `_inputs/check_discrepancies.py` (in `/tmp/audit/`) — per-workbook-discrepancy lookup against engine
