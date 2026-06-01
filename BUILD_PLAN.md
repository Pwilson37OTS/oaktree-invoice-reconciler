# Bullhorn ↔ QBO Revenue Reconciler — Build Plan

**Status:** Awaiting (1) Phil's confirmation of Bullhorn ATS report drop path, (2) sign-off on matching logic translation, (3) sign-off on output substrate (Streamlit).
**Owner:** Phil
**Date:** 2026-05-27
**Source workbook reverse-engineered:** `_inputs/Revenue_Audit_Master_File_2026-04.xlsm` (310 KB, Apr 2026 period)

---

## 1. Goal

Replace the manual VBA-driven Revenue Audit workbook with an always-on Python engine + Streamlit dashboard that reconciles Bullhorn (ATS) billable activity against QBO revenue every time the source files refresh. Billing team gets continuous self-serve visibility; Phil + Brenda get a one-click month-close snapshot.

OTS-only in MVP. CTS in phase 2. Invoices/revenue only in MVP. Payroll (Bullhorn → iSolved) is a separate later module.

## 2. What the existing VBA actually does

Read this section before writing any engine code. Everything in section 3 is a translation of this.

### Inputs

**Sheet `QBO Revenue`** — pasted from QBO's "Revenue Audit Report (Includes Expenses)" memorized report. 10 columns. Header on row 5.

| Col | Field | Notes |
|---|---|---|
| B | Transaction date | Often blank on most rows; VBA backfills from row above (QBO groups by date) |
| C | Transaction type | Invoice / Credit Memo / Journal Entry / etc. |
| D | Num | Invoice number (not used in matching) |
| E | Name | Customer/vendor name |
| F | Description | **Carries the contractor key + billing-week date**, e.g. `182449, Anselmo Morales - Regular - 2025-12-14` (extended form may append `- City - State - ZIP`) |
| G | Item split account | Used to detect "Employee Advance" rows for sign-flipping |
| I | Amount | The dollar value being matched |

**Sheet `ATS Hours`** — pasted from a Bullhorn report. 18 columns. Header on row 1.

| Col | Field | Notes |
|---|---|---|
| B | Data Item2 | Billing-week start or period-end date (the date used as the match key) |
| C | Candidate ID | The numeric contractor key — same as the leading token in QBO Description |
| D | Candidate Name | Display name |
| K | Gross Sales Amount | The dollar value being matched |
| (others) | Bill rate, hours, GM, branch, etc. | Carried for context only, not used in matching |

### Pre-processing the VBA performs on QBO Revenue

1. **Parse the billing-week date out of the Description.** Split column F by `" - "`. Walk the parts from right to left; the rightmost part that `IsDate()`-parses is the billing-week date. Write it into column B.
   - Standard: `182449, Anselmo Morales - Regular - 2025-12-14` → date `2025-12-14`
   - Extended: `182827, Erika Rivera - Regular - 2025-12-13 - El Paso - TX - 79930` → date `2025-12-13`

2. **Sign-correct Employee Advance lines** (column G == `"Employee Advance"`):
   - If transaction type is `Credit Memo` → force amount to **negative** absolute value.
   - Otherwise → force amount to **positive** absolute value.
   - This normalizes QBO's natural sign convention (which flips between Invoice and Credit Memo on the Employee Advance account) so a per-contractor-per-week sum nets to a comparable number.

3. **Backfill blank dates** in column B by walking upward to the nearest non-empty row (QBO groups multi-line invoices under one date).

### Matching key & aggregation

- **Key:** `candidate_id | yyyy-mm-dd` where:
  - `candidate_id` on the QBO side = the first comma-delimited token of column F Description.
  - `candidate_id` on the ATS side = column C.
  - Date on the QBO side = the parsed/backfilled column B.
  - Date on the ATS side = column B.
- **Aggregate by sum:** both QBO Amount (col I) and ATS Gross Sales Amount (col K) are summed per key. This is why multi-line invoices and split bills still match — they roll up to one contractor-week total.
- **All numbers rounded to 2 decimals** before comparison.

### Output

- **`Audit` sheet** — one row per unique key (full outer join). Columns: Contractor Name, QBO Date, QBO Amount, ATS Date, ATS Amount, Difference (`ROUND(ATS - QBO, 2)`).
- **`Discrepancies` sheet** — same columns, filtered to rows where Difference ≠ 0.

### What the April 2026 run produced

- 1,132 audit rows (one per unique contractor-week-amount across both sources).
- **23 discrepancies** — overwhelmingly $0.01 rounding diffs (e.g., Air Liquide $1,045.97 vs $1,045.98). One large item (Wesley Qualls, $14,271.60/.61). One contractor-only row (Aaron Schehl, QBO row with no matching ATS row → suggests credit memo or off-cycle adjustment).

**Implication:** the current system is already accurate at ~98%. The automation's job isn't to improve accuracy — it's to **eliminate the manual effort** of pasting data + running the macro + reviewing it, and to give the billing team continuous visibility so anomalies are caught the day they happen instead of three weeks later.

## 3. Engine spec (Python translation)

### Inputs

- **QBO source:** the same "Revenue Audit Report" QBO memorized report, exported to `_shared/quickbooks-data/Revenue_Audit.xlsx` (path TBD with Brenda — confirm the file name she saves it as). Refresh weekly with the existing QBO exports.
- **ATS/Bullhorn source:** **PATH TBD** — Phil is wiring up the Bullhorn report drop into `_shared/bullhorn-data/`. Until that path is fixed, the engine is stubbed.

### Modules

| Module | Purpose |
|---|---|
| `parse_qbo.py` | Read the Revenue Audit Report sheet. Parse date from Description, sign-flip Employee Advance lines, backfill dates, return list of `QBORow` records. |
| `parse_bullhorn.py` | Read the ATS Hours report. Return list of `ATSRow` records. (Write once Bullhorn report format is fixed.) |
| `match.py` | Build the per-key dictionaries (`atsDict`, `qboSums`), union the keys, produce reconciled rows. |
| `classify.py` | Tag each reconciled row with one of: `CLEAN` (diff == 0), `PENNY_ROUND` (\|diff\| ≤ $0.01), `MINOR_VARIANCE` ($0.01 < \|diff\| ≤ $1.00), `MATERIAL_VARIANCE` (\|diff\| > $1.00), `ATS_ONLY_REVENUE_NOT_IN_QBO` (QBO understated), `QBO_ONLY_NO_ATS_SOURCE` (QBO overstated or off-cycle), `ZERO_AMOUNT_NO_ACTION` (one-sided AND that side sums to $0.00 — e.g. Bullhorn row with $0 bill rate, or QBO recon JE that nets to zero; no invoice ever expected). |
| `reconcile.py` | Orchestrator. Calls the parsers, runs match, classifies, writes JSON + Excel snapshot. |
| `app.py` | Streamlit dashboard. Continuous view + month-close view. |

### Output (per run)

- `output/invoice_reconciliation_data.json` — full reconciled dataset for the Streamlit app.
- `output/close_YYYY-MM.xlsx` — month-close snapshot with tabs: `Audit`, `Discrepancies`, `Summary`. Identical column shape to the existing workbook so Brenda can open it and recognize it immediately during the transition.

### Match parameters (defaults — same as VBA today)

- Match key: `candidate_id | yyyy-mm-dd` (exact).
- Aggregation: sum per key on both sides.
- Tolerance: $0.00 absolute (VBA uses exact ROUND comparison). We'll add a `PENNY_ROUND` bucket on top for the $0.01 noise.
- Sign rules: Employee Advance + Credit Memo → negate; Employee Advance + anything else → positive abs.
- Window: full period present in the source files (no time-window cutoff at engine level; the dashboard filters by month/week).

## 4. Streamlit dashboard

### Continuous audit view (default landing page)

- Tiles: total billed ATS, total revenue QBO, net variance, # discrepancies, # material-variance discrepancies.
- Filterable, sortable table — every reconciled row (full join). Filter chips: Week, Status, Contractor, Client, AM/Recruiter (joined from `_shared/bullhorn-data/299.xlsx`).
- Click any row → drawer showing the underlying ATS lines + QBO lines that aggregated to that key, plus parsed description, transaction type, customer, branch.
- Color: clean = neutral; penny = pale yellow; ATS-only / QBO-only / material variance = red.

### Month-close view

- Lock to a chosen month. Show same tiles + table filtered to that month.
- "Sign off and snapshot" button → writes `output/close_YYYY-MM/snapshot.xlsx` + records the signer's name and timestamp.
- After sign-off, the month displays a "Closed" badge; subsequent changes are flagged as `POST_CLOSE_ADJUSTMENT`.

### Hosting (MVP)

- Run locally first (`streamlit run app.py`). Validate against the April 2026 workbook output before any deployment.
- Once validated: Streamlit Community Cloud private, OAuth via Microsoft (matches OakTree's M365 tenant).

## 5. Validation plan

Before any production run, prove parity with the existing workbook:

1. Run the new engine against the April 2026 input data already present in the workbook (`_inputs/Revenue_Audit_Master_File_2026-04.xlsm`).
2. Compare the engine's output row-by-row to the workbook's `Audit` sheet (1,132 rows) and `Discrepancies` sheet (23 rows).
3. **Acceptance criterion:** 100% row-for-row match. Anything different is either a VBA bug we found or a translation error in the engine — investigate every one.
4. After parity, run it against the most recent closed month independently and have Brenda compare to her notes.

## 6. Phase 0 — what's still blocking

- [ ] **Bullhorn report path & format.** Phil is wiring up the canned report to drop into `_shared/bullhorn-data/`. Once we have a sample file + the agreed path, `parse_bullhorn.py` gets written and the engine can run end-to-end.
- [ ] **QBO Revenue Audit Report path.** Confirm with Brenda the exact filename she'll save the QBO "Revenue Audit Report (Includes Expenses)" as in `_shared/quickbooks-data/`. Recommend: `Revenue_Audit.xlsx`.
- [ ] **Sign-off on the classification tiers in section 3.** Phil to confirm `MATERIAL_VARIANCE` threshold ($1.00 default) and that `PENNY_ROUND` is OK to auto-bucket (vs. surfacing every penny diff).

## 7. Phase 1 — what gets built once Phase 0 clears

1. `parse_qbo.py` (the description-parsing + sign-flip logic is mostly a port of the VBA).
2. `parse_bullhorn.py` (depends on actual Bullhorn report format).
3. `match.py`, `classify.py`, `reconcile.py`.
4. Local Streamlit app with both views.
5. April 2026 parity test (section 5).
6. Brenda walkthrough → tune classifications, finalize tiles.

## 8. Phase 2 — CTS + hosting + notifications

- Add CTS entity (parallel exports already in `_shared/quickbooks-data/*_CTS.xlsx`).
- Deploy Streamlit Cloud private with M365 SSO.
- Daily Slack/email digest of any new material variance ≥ threshold.

## 9. Phase 3 — Live integrations + payroll

- Replace file-based QBO ingest with direct QBO MCP queries (no waiting for the export refresh).
- Build sibling `payroll-reconciler/` module (Bullhorn → iSolved) once iSolved structured data is available.

## 10. Open questions for Phil

1. Material-variance threshold: $1.00 default — too tight, too loose?
2. Penny-round bucket: auto-hide (default) or always surface?
3. Should the dashboard pull AM/recruiter context from `299.xlsx` for every contractor, or only on drill-down (faster)?
4. Who can sign off and close the month — Phil only, or Phil + Brenda?

## 11. Files & references

- `_inputs/Revenue_Audit_Master_File_2026-04.xlsm` — the workbook this plan was derived from.
- `_inputs/Revenue_Audit_VBA_source.bas` — extracted VBA source (`AuditSheets` macro).
- Pattern reference: `jarvis-phil/employee-advance-reconciler/` (especially `reconcile.py` v9 — same idioms apply here).
