# Tally Financial Statements Generator

A standalone Python desktop app that reads Tally data exported by the **[TSF Exporter](https://github.com/dhruvdua88/Tally-TSF-Exporter)** and generates publication-ready **Schedule III financial statements** in Excel — with linked notes, 3-year projections (simple or banker-grade), cash flow, ratios, charts, and a full data-validation report. Multiple branches are auto-consolidated with per-branch columns and a Consolidated total.

**Accepts three input formats from the TSF Exporter:**
- **ZIP file** — the `.zip` produced directly by the TSF Exporter (recommended, one click)
- **SQLite file** — the `.sqlite` / `.db` extracted from the ZIP
- **CSV folder** — the folder of `.csv` files inside the ZIP

> The TSF Exporter is a separate Windows app that connects to TallyPrime and exports your data. Run it once per branch/company to get the input files for this app.

---

## ⬇ Download

**[Download latest release (ZIP)](https://github.com/dhruvdua88/tally-fin-statements/releases/latest)**

The ZIP contains everything: the app, a one-click Mac launcher, and a one-click Windows launcher. No coding required.

---

## What's New

### v2.0 — Multi-branch consolidation

- **Add multiple Tally SQLite files** (one per company branch) in the GUI and generate **one Excel workbook** with per-branch columns plus a Consolidated total in the Schedule III Balance Sheet and Statement of P&L.
- **Source files panel** replaces the single-file picker — a list with **Add Branch…**, **Rename**, and **Remove** buttons. Each branch loads independently, and the consolidated card refreshes on every change.
- **Side-by-side columns** in BS and P&L: `Particulars | Note | <Branch 1> | <Branch 2> | … | Consolidated`. Notes prefix each ledger with `[BranchName]` to disambiguate identical names across branches.
- **Period validation** — refuses to consolidate if `period_from` / `period_to` don't match across branches (clear error popup naming the mismatched branches).
- **Stock overrides** apply to the consolidated total; projections still run on the consolidated data.
- **Backwards compatible** — load a single file and you get the original single-column layout unchanged.

### v1.2 — Simple/Detailed projection modes + optional analytics sheets

- **Simple projection mode** — only 3 inputs needed (revenue growth, gross margin, tax rate). Working-capital days, OpEx growth, interest and depreciation rates are derived automatically from the base year. Choose Simple or Detailed via a radio on the Projections tab.
- **Optional Excel sheets** (all toggleable via checkboxes on the Projections tab):
  - **Ratios** — profitability (Gross/EBITDA/PAT margin, RoE, RoCE), liquidity (Current, Quick), leverage (D/E, Interest Coverage, DSCR), efficiency days (Inventory, Debtor, Creditor)
  - **Cash Flow Statement** — 3-year, indirect method (CFO + CFI + CFF, with opening/closing cash)
  - **Common-Size statements** — P&L as % of revenue, Balance Sheet as % of total assets
  - **Charts** — Revenue / EBITDA / PAT bar chart + PAT-margin trend line
  - **Banker view** — adds DSCR + Interest Coverage rows directly on the Projected P&L
- **Projection engine — modelling bugs fixed:**
  - Interest now computed on average LT-borrowings (opening + closing) / 2 — not period-end.
  - CapEx no longer double-counted (mid-year-convention depreciation).
  - Inventory days re-applied every year (was previously frozen after year 0).
  - Cash plug going negative now surfaces as **"Additional Short-Term Borrowing Needed"** in red instead of silently showing negative cash.
- **Print-ready page setup** on every sheet (A4, fit-to-width, company header, page numbers, date footer).
- **Settings persistence** — last DB path, output folder, projection inputs, mode, and output options are saved to `~/.tallyfin_settings.json` and restored on next launch.

### v1.1 — Balance Sheet linking & formatting

- BS face cells **linked via Excel formulas** to each note's total cell — change a number on a note sheet, the BS face updates automatically (`='N1 Share Capital'!B4`).
- Grand totals (`TOTAL EQUITY & LIABILITIES`, `TOTAL ASSETS`) and the BS Difference are formula-driven.
- Indian accounting number format (negatives in parens, zero as `–`).
- Total rows have full-width fills (A:D), thin borders on every data row, bold subtotal amounts.
- Sheets reordered: Balance Sheet → P&L → Projected BS → Projected P&L → Notes; file opens on Balance Sheet.

---

## Install & Run (no coding needed)

### Mac

1. Download the ZIP from the link above and unzip it anywhere (e.g. your Desktop)
2. Right-click **`run_mac.command`** → **Open** → click **Open** in the security dialog
   *(first time only — macOS flags new downloads; after that you can double-click directly)*
3. A Terminal window opens, installs the one dependency automatically, then launches the app
4. The Terminal window can be closed once the app window appears

> **No Python?** The launcher will open `python.org/downloads` for you automatically.

### Windows

1. Download the ZIP and unzip it anywhere
2. Double-click **`run_windows.bat`**
3. A Command Prompt window installs the dependency then opens the app
4. The Command Prompt can be closed once the app appears

> **No Python?** The launcher opens `python.org/downloads` and explains what to tick ("Add Python to PATH").

### Manual (for developers)

```bash
pip install openpyxl
python financial_statements.py
```

---

## Features

| Area | Detail |
|---|---|
| **Multi-Branch Consolidation** | Load multiple Tally SQLite files (one per branch). Balance Sheet and P&L emit per-branch columns plus a Consolidated total. Periods must match across branches. |
| **Balance Sheet** | Schedule III format (Indian Companies Act 2013); face cells are formula-linked to note totals; note numbers are clickable hyperlinks |
| **P&L Statement** | Revenue, purchases, employee costs, finance costs, depreciation, other expenses; key ratios |
| **Note Sheets** | 9 separate Excel sheets (Share Capital → Cash & Bank), each with a "← Back" link |
| **Notes Index** | One-page index of all notes with amounts and hyperlinks |
| **Projections — Simple Mode** | Just 3 inputs (revenue growth, gross margin, tax rate). Everything else is derived from base year. |
| **Projections — Detailed Mode** | All 15 banker-grade inputs (working capital days, OpEx growth, interest, CapEx, etc.) |
| **Optional Analytics** | Ratios sheet · Cash Flow Statement · Common-Size statements · Charts · Banker view (DSCR + Interest Coverage) — all opt-in |
| **Print-Ready Output** | A4 fit-to-width page setup on every sheet, with company header and page footer |
| **Validation Report** | 15+ automated checks — BS balance equation, P&L reconciliation, sign flips, natural-sign breaches, stale balances, unclassified groups |
| **Group Mapping** | Auto-infers Schedule III head for any non-standard Tally primary group; user can override via GUI dropdown |
| **Stock Overrides** | Opening and closing stock can be corrected without editing Tally |
| **Settings Persistence** | Remembers last file paths, projection inputs and output options across launches |
| **GUI** | `tkinter` desktop app — no web server, single `pip install openpyxl` dependency |

---

## Requirements

- Python 3.11+
- `openpyxl` — `pip install openpyxl`
- A Tally export produced by the **[TSF Exporter](https://github.com/dhruvdua88/Tally-TSF-Exporter)** (ZIP, SQLite, or CSV folder)

---

## Quick Start

### Step 1 — Export from TallyPrime (using TSF Exporter)

1. Download and run the **[TSF Exporter](https://github.com/dhruvdua88/Tally-TSF-Exporter)** on the computer running TallyPrime
2. Select your date range and click **Run export** — it produces a `.zip` file
3. For **two branches** (e.g. two companies or offices): run the TSF Exporter once for each and save two separate `.zip` files

### Step 2 — Generate statements

```bash
pip install openpyxl
python financial_statements.py
```

1. Click **+ Add Branch…**, select your `.zip` file (or `.sqlite` / `.db` if you extracted it). Give the branch a name when prompted.
2. For two branches: click **+ Add Branch…** again and load the second `.zip`. Both branches must cover the same period.
   - *Have only CSV files?* Use **+ From CSV Folder…** instead and select the folder that contains `mst_ledger.csv`.
3. Optionally enter corrected opening/closing stock values
4. Click **Generate Statements (Excel)** — the file lands on your Desktop (or chosen folder)
5. For projections: fill in the 3-Year Projections section and click **Generate + 3-Year Projections**

Settings persist — next launch remembers your last files and inputs.

---

## Excel Output Structure

| Sheet | Always included? | Contents |
|---|---|---|
| `Balance Sheet` | ✓ | Schedule III face — formula-linked to note totals. With ≥2 branches loaded, columns become `Particulars · Note · <Branch1> · <Branch2> · … · Consolidated`. |
| `P&L Statement` | ✓ | Schedule III face + key financial ratios. Same per-branch + Consolidated column layout when multiple branches are loaded. |
| `N1 Share Capital` | ✓ | Ledger-level detail |
| `N2 Reserves Surplus` | ✓ | Reserves & Surplus breakdown |
| `N3 LT Borrowings` | ✓ | Secured / Unsecured loans by ledger |
| `N4 ST Borrowings` | ✓ | Bank OD / CC accounts |
| `N5 Trade Payables` | ✓ | Creditors by parent sub-group |
| `N8 Fixed Assets` | ✓ | 7-column schedule: Gross Open → Net Block |
| `N11 Inventories` | ✓ | Opening and closing stock |
| `N12 Trade Receivables` | ✓ | Debtors by parent sub-group |
| `N13 Cash & Bank` | ✓ | Cash-in-hand + bank accounts (debit-balance only) |
| `Notes Index` | ✓ | All notes with amounts and hyperlinks |
| `Validation` | ✓ | Color-coded ERROR / WARNING / INFO check results |
| `Projected P&L` | If projections enabled | Base year + Years 1–3; can include DSCR + Interest Coverage rows |
| `Projected Balance Sheet` | If projections enabled | Base year + Years 1–3; shows funding-shortfall flag if cash plug went negative |
| `Assumptions` | If projections enabled | All projection inputs recorded for audit trail |
| `Ratios` | **Optional (default ON)** | Profitability, liquidity, leverage, efficiency days |
| `Cash Flow` | **Optional** | 3-year, indirect method |
| `Common-Size` | **Optional** | P&L as % of revenue, BS as % of total assets |
| `Charts` | **Optional (default ON)** | Revenue/EBITDA/PAT bar chart + PAT-margin line |

Every note sheet has a **← Back to Balance Sheet** hyperlink in cell A2. The file opens on the Balance Sheet.

---

## TSF Schema

The app reads Tally data exported by the **[TSF Exporter](https://github.com/dhruvdua88/Tally-TSF-Exporter)**.  
It accepts the ZIP directly, or a SQLite / CSV folder extracted from it.  
The data must contain:

### Required tables

#### `mst_ledger`
| Column | Type | Description |
|---|---|---|
| `name` | TEXT | Ledger name |
| `parent` | TEXT | Direct parent group name |
| `closing_balance` | TEXT | Year-end balance in Tally's TEXT format (negative = debit) |
| `opening_balance` | TEXT | *(optional)* Prior-year closing balance |

**Sign convention:** Tally stores amounts as TEXT with Indian comma formatting (`1,23,456.78`). Negative values are debit balances (assets/expenses); positive values are credit balances (liabilities/income). Some older exports append `Cr` or `Dr` suffixes — the `safe_float()` parser handles all variants.

#### `mst_group`
| Column | Type | Description |
|---|---|---|
| `name` | TEXT | Group name (matches `mst_ledger.parent`) |
| `primary_group` | TEXT | Tally primary group — the Schedule III classification key |
| `is_deemedpositive` | TEXT | `'1'` if credit is the natural balance for this group |

#### `_export_info`
Row-per-field table (name / value columns):

| `name` value | `value` description |
|---|---|
| `company_name` | Registered company name |
| `period_from` | ISO date (`YYYY-MM-DD`) of period start |
| `period_to` | ISO date (`YYYY-MM-DD`) of period end |

### Optional tables

| Table | Used for |
|---|---|
| `trn_accounting` | Cross-check totals (not used for primary P&L computation) |
| `trn_voucher` | Not used by this tool |

### Primary group → Schedule III mapping

The app uses two static maps to classify every Tally primary group:

**`BS_MAP`** — Balance Sheet groups:

| Tally Primary Group | Schedule III Head | Side |
|---|---|---|
| Fixed Assets | Fixed Assets | Non-Current Asset |
| Investments | Non-Current Investments | Non-Current Asset |
| Deposits (Asset) | Long-Term Loans & Advances | Non-Current Asset |
| Loans & Advances (Asset) | Long-Term Loans & Advances | Non-Current Asset |
| Misc. Expenses (ASSET) | Other Non-Current Assets | Non-Current Asset |
| Stock-in-hand | Inventories | Current Asset |
| Sundry Debtors | Trade Receivables | Current Asset |
| Cash-in-hand | Cash & Cash Equivalents | Current Asset |
| Bank Accounts | Cash & Cash Equivalents | Current Asset |
| Current Assets | Other Current Assets | Current Asset |
| Capital Account | Share Capital | Equity |
| Reserves & Surplus | Reserves & Surplus | Equity |
| Secured Loans | Long-Term Borrowings | Non-Current Liability |
| Unsecured Loans | Long-Term Borrowings | Non-Current Liability |
| Loans (Liability) | Long-Term Borrowings | Non-Current Liability |
| Bank OD A/c | Short-Term Borrowings | Current Liability |
| Sundry Creditors | Trade Payables | Current Liability |
| Current Liabilities | Other Current Liabilities | Current Liability |
| Branch / Divisions | Other Current Liabilities | Current Liability |
| Duties & Taxes | Duties & Taxes (Net) | Current Liability |
| Provisions | Short-Term Provisions | Current Liability |
| Suspense A/c | Other Current Liabilities | Current Liability |

**`PNL_MAP`** — P&L groups:

| Tally Primary Group | Schedule III Line |
|---|---|
| Sales Accounts | Revenue from Operations |
| Direct Incomes | Revenue from Operations |
| Indirect Incomes | Other Income |
| Purchase Accounts | Cost of Materials / Purchases |
| Direct Expenses | Direct Expenses |
| Indirect Expenses | Indirect Expenses |

**Auto-inference:** If a primary group is not in either map, the app uses 30 keyword rules to infer the closest standard group (e.g. `"Internet Sales"` → `Sales Accounts`). Users can override this via the **Group Mapping** tab.

---

## Tally Sign Convention

Tally stores all ledger balances as TEXT with a critical sign convention:

```
negative value  →  debit balance   (assets, expenses — natural debit)
positive value  →  credit balance  (liabilities, equity, income — natural credit)
```

The app displays all amounts as positive on the face of statements by negating debit-natured figures:

```python
# Asset display (e.g. Trade Receivables):
display_value = -ledger.closing   # debit → positive for display

# Liability display (e.g. Trade Payables):
display_value = ledger.closing    # credit → positive for display

# Revenue (Sales Accounts): credit → positive as-is
# Expense (Purchase Accounts): debit → negate to show positive
```

**Bank Accounts special case:** Bank Accounts under "Bank Accounts" primary group can have either sign:
- Negative closing = debit = cash held → shown under **Cash & Cash Equivalents**
- Positive closing = credit = overdraft → shown under **Short-Term Borrowings**

---

## P&L Data Source

The app uses `mst_ledger.closing_balance` (not `trn_accounting`) as the primary P&L data source. This matches Tally's own P&L report because `trn_accounting` includes journal entries and inter-branch adjustments that Tally excludes from its P&L computation.

**P&L A/c balance treatment:**
```
pnl_balance (closing)  = cumulative balance = prior retained earnings + current year profit
pnl_opening            = prior year retained earnings not yet transferred to Reserves
current year profit    = pnl_balance - pnl_opening
```

---

## Projection Engine

The 3-year projection uses a **cash-plug balance sheet** method, with two modes:

### Simple Mode (recommended for quick estimates)
Only three inputs required — everything else is derived from the base year:
- Revenue Growth Y1/Y2/Y3 (%)
- Gross Margin (%)
- Effective Tax Rate (%)

Inventory days, debtor days, creditor days are auto-computed from base-year ratios; OpEx growth = minimum of revenue growth; interest rate defaults to 12 %; depreciation rate defaults to 15 %.

### Detailed Mode (for banker submissions)
Full control over all 15 levers:

| Parameter | Default | Description |
|---|---|---|
| Revenue Growth Y1/Y2/Y3 | 15 / 15 / 10 % | Year-on-year revenue growth |
| Gross Margin % | 30% | (Revenue − Purchases ± Stock) / Revenue |
| OpEx Growth % | 10% | Annual growth in employee + other indirect expenses |
| Inventory Days | 45 | Days of COGS held as stock |
| Debtor Days | 60 | Days outstanding for trade receivables |
| Creditor Days | 45 | Days outstanding for trade payables |
| Loan Repayment p.a. | 0 | Annual long-term loan repayment (₹) |
| New Borrowings p.a. | 0 | Annual new long-term borrowings (₹) |
| Interest Rate % | 14% | Annualised rate on **average** borrowings |
| CapEx p.a. | 0 | Annual capital expenditure (₹) — mid-year convention |
| Depreciation Rate % | 15% | Written-down-value depreciation rate |
| Tax Rate % | 25% | Effective corporate tax rate |
| Other Income p.a. | 0 | Fixed other income per year (₹) |

### Modelling notes (v1.2)
- **Interest:** computed on the average of opening and closing long-term borrowings each year, not the period-end balance.
- **CapEx:** depreciated on a mid-year convention (half-year in the year of addition) and lands once in the closing net block — no double-counting.
- **Inventory:** recomputed every year as `COGS / 365 × inventory_days` (was previously frozen after year 0).
- **Cash plug:** if total liabilities can't fund total assets, the shortfall is surfaced as "Additional Short-Term Borrowing Needed" in red on the Projected BS, instead of showing negative cash.

---

## Validation Checks

The app runs 15+ automated checks on every load:

| Category | Check | Severity |
|---|---|---|
| Schema | Required tables present (`mst_ledger`, `mst_group`, `_export_info`) | ERROR |
| Schema | Required columns present | ERROR |
| Schema | Optional tables absent (`trn_accounting`) | WARNING |
| Balance Sheet | Assets = Equity + Liabilities (within ₹1) | INFO / WARNING / ERROR |
| P&L | Computed profit = Tally's P&L A/c current-year balance | INFO / WARNING / ERROR |
| Classification | Primary groups not in BS_MAP or PNL_MAP with material balances | WARNING |
| Data Quality | Opening stock = 0 but closing stock non-zero | WARNING |
| Data Quality | Closing stock > 60% of revenue | WARNING |
| Data Quality | Net cash & bank negative | WARNING |
| Data Quality | Trade receivables > annual revenue | WARNING |
| Data Quality | Share capital negative | ERROR |
| Data Quality | Reserves & surplus highly negative | WARNING |
| Data Quality | Long-term borrowings negative | WARNING |
| Data Quality | Period not approximately 12 months | WARNING |
| Data Quality | Sign flips (opening and closing opposite signs) | WARNING |
| Data Quality | Natural sign breaches (asset with credit, liability with debit) | WARNING |
| Data Quality | Large stale balances (≥₹1L, no movement) | INFO |

Results appear in the **Validation** tab in the GUI and in the **Validation** sheet in the Excel output (color-coded red / amber / green).

---

## File Structure

```
tally-fin-statements/
├── financial_statements.py   # Single-file app — everything is here
├── run_mac.command           # Mac one-click launcher
├── run_windows.bat           # Windows one-click launcher
├── requirements.txt          # openpyxl>=3.1.0
└── README.md                 # This file
```

A small settings file (`~/.tallyfin_settings.json`) is created on first run to remember last paths and projection inputs.

---

## Limitations

- **Single period only** — all loaded branches must cover the same `period_from` / `period_to` (mismatched periods are rejected on consolidate). One SQLite file = one year's statements per branch.
- **No previous-year column** — the "Previous Year" column in the BS face is reserved (requires two exports — planned for a future release)
- **Simplified tax** — tax expense uses the Deferred Tax Liability ledger as a proxy on the Actual sheets; projections use a flat effective rate
- **Depreciation from FA schedule** — if Tally has not booked depreciation entries (common when using a separate depreciation workbook), the P&L depreciation line will be zero; add it via a manual adjustment or use the Indirect Expenses sub-group instead
- **Projection accuracy** — projections are management estimates based on user inputs; the cash-plug method means cash absorbs all modelling errors, and a negative plug is flagged as a funding shortfall

---

## Licence

MIT — see [LICENSE](../LICENSE)
