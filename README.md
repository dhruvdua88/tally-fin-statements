# Tally Financial Statements Generator

## ⬇ Download

> ### **➡ [Download ZIP (direct, always latest)](https://github.com/dhruvdua88/tally-fin-statements/archive/refs/heads/main.zip) ⬅**
>
> Clicking the link downloads `tally-fin-statements-main.zip` immediately — no intermediate page. The ZIP contains the app plus one-click launchers for **Mac** (`run_mac.command`) and **Windows** (`run_windows.bat`). No coding required — see [Install & Run](#install--run-no-coding-needed) below.
>
> Prefer a tagged release? Browse all versions at [github.com/dhruvdua88/tally-fin-statements/releases](https://github.com/dhruvdua88/tally-fin-statements/releases).

You will also need the companion **[TSF Exporter](https://github.com/dhruvdua88/Tally-TSF-Exporter)** to pull data out of TallyPrime — see [How it connects to Tally](#how-it-connects-to-tally).

---

## What it does

A standalone Python desktop app that reads Tally data exported by the **[TSF Exporter](https://github.com/dhruvdua88/Tally-TSF-Exporter)** and generates **Schedule III financial statements** in Excel:

- **Balance Sheet** and **Statement of P&L** in Indian Companies Act 2013 format, with formula-linked note hyperlinks
- **9 note schedules** (Share Capital, Reserves, Borrowings, Trade Payables/Receivables, Fixed Assets, Inventories, Cash & Bank)
- **3-Year Projected P&L and Balance Sheet** with cash-plug method
- **Validation sheet** — 26 automated checks (BS balance, P&L reconciliation, sign breaches, etc.)
- **Multi-branch consolidation** — load multiple TSF files, get per-branch columns + a Consolidated total

**Accepts three input formats from the TSF Exporter:**
- **ZIP file** — the `.zip` produced directly by the TSF Exporter (recommended, one click)
- **SQLite file** — the `.sqlite` / `.db` extracted from the ZIP
- **CSV folder** — the folder of `.csv` files inside the ZIP

---

## How it connects to Tally

This app does **not** speak to TallyPrime directly. It reads files produced by the **[TSF Exporter](https://github.com/dhruvdua88/Tally-TSF-Exporter)**, which is the trusted bridge:

```
  TallyPrime  ──XML over TCP 9000──▶  TSF Exporter (Node.js + Python)  ──ZIP/CSV/SQLite──▶  This app  ──▶  Excel
```

**Why two tools?** The TSF Exporter handles all TallyPrime XML/TDL protocol complexity on the Windows machine where Tally runs. This app then works on **any** OS (Mac, Windows, Linux) against the exported files — your auditor or analyst doesn't need Tally running locally.

**For two branches:** run the TSF Exporter once per company/branch in TallyPrime, save the two `.zip` files, then load both into this app. Periods must match across branches.

---

## What's New

### v2.1 — ZIP + CSV folder input

- **Accept TSF Exporter ZIP files directly** — no need to unzip first; the app extracts the SQLite into a temp file, loads it, and cleans up.
- **+ From CSV Folder…** button — load a folder of CSV files (`mst_ledger.csv`, `mst_group.csv`, `_export_info.csv`) when you only have the extracted CSVs.
- File dialog filters updated to show `.zip` alongside `.sqlite` / `.db`.

### v2.0 — Multi-branch consolidation

- **Add multiple Tally exports** (one per company branch) in the GUI and generate **one Excel workbook** with per-branch columns plus a Consolidated total in the Schedule III Balance Sheet and Statement of P&L.
- **Source files panel** with a list and **Add Branch…**, **Rename**, **Remove** buttons. Each branch loads independently.
- **Side-by-side columns** in BS and P&L: `Particulars | Note | <Branch 1> | <Branch 2> | … | Consolidated`. Notes prefix each ledger with `[BranchName]` to disambiguate identical names across branches.
- **Period validation** — refuses to consolidate if `period_from` / `period_to` don't match across branches.
- **Stock overrides** apply to the consolidated total; projections still run on the consolidated data.
- **Backwards compatible** — load a single file and you get the original single-column layout unchanged.

### v1.2 — Projection engine fixes + print-ready output

- **Projection engine — modelling bugs fixed:**
  - Interest now computed on average LT-borrowings (opening + closing) / 2 — not period-end.
  - CapEx no longer double-counted (mid-year-convention depreciation).
  - Inventory days re-applied every year (was previously frozen after year 0).
  - Cash plug going negative now surfaces as **"Additional Short-Term Borrowing Needed"** in red instead of silently showing negative cash.
- **Print-ready page setup** on every sheet (A4, fit-to-width, company header, page numbers, date footer).
- **Settings persistence** — last paths, projection inputs, and output options saved to `~/.tallyfin_settings.json` and restored on next launch.

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
| **Multi-Branch Consolidation** | Load multiple Tally exports (one per branch) as ZIP, SQLite, or CSV folder. Balance Sheet and P&L emit per-branch columns plus a Consolidated total. Periods must match across branches. |
| **Three Input Formats** | TSF Exporter `.zip` (auto-extracts), `.sqlite` / `.db`, or a folder of CSVs (`mst_ledger.csv`, `mst_group.csv`, `_export_info.csv`). |
| **Balance Sheet** | Schedule III format (Indian Companies Act 2013); face cells are formula-linked to note totals; note numbers are clickable hyperlinks. |
| **P&L Statement** | Revenue, purchases, employee costs, finance costs, depreciation, other expenses; inline key ratios (Gross Margin, EBITDA, PAT margin). |
| **Note Sheets** | 9 separate Excel sheets (Share Capital → Cash & Bank), each with a "← Back" link. |
| **Notes Index** | One-page index of all notes with amounts and hyperlinks. |
| **3-Year Projections** | Banker/investor format with 15 inputs (revenue growth Y1/Y2/Y3, gross margin, OpEx growth, working-capital days, loan schedule, interest, CapEx, depreciation, tax rate, other income). Cash-plug method with funding-shortfall flag. |
| **Print-Ready Output** | A4 fit-to-width page setup on every sheet, with company header and page footer. |
| **Validation Report** | 26 automated checks — BS balance equation, P&L reconciliation, sign flips, natural-sign breaches, stale balances, unclassified groups, sub-12-month period detection, etc. |
| **Group Mapping** | Auto-infers Schedule III head for any non-standard Tally primary group (28 keyword rules); user can override via GUI dropdown. |
| **Stock Overrides** | Opening and closing stock can be corrected without editing Tally. |
| **Settings Persistence** | Remembers last file paths, projection inputs and output options across launches (`~/.tallyfin_settings.json`). |
| **GUI** | `tkinter` desktop app — no web server, single `pip install openpyxl` dependency. |

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
| `P&L Statement` | ✓ | Schedule III face + inline key ratios (Gross Margin, EBITDA, PAT margin). Same per-branch + Consolidated column layout when multiple branches are loaded. |
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
| `Validation` | ✓ | Color-coded ERROR / WARNING / INFO check results (26 checks) |
| `Projected P&L` | If projections enabled | Base year + Years 1–3 |
| `Projected Balance Sheet` | If projections enabled | Base year + Years 1–3; shows funding-shortfall flag if cash plug went negative |
| `Assumptions` | If projections enabled | All projection inputs recorded for audit trail |

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

The 3-year projection uses a **cash-plug balance sheet** method. All 15 levers are exposed in the GUI with sensible defaults — fill in what matters for your case and leave the rest as-is:

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

The app runs **26 automated checks** on every load:

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

## How robust are the statements?

The TSF Exporter + this app combination is designed to produce statements that match Tally's own reports:

**What is enforced:**
- Schema validation rejects malformed inputs before any computation runs
- P&L numbers come from `mst_ledger.closing_balance` (the same source Tally's P&L report uses) — not `trn_accounting`, which includes inter-branch and journal entries Tally excludes
- BS face cells are formula-linked to note totals, so the printed Balance Sheet always reconciles with its notes
- Sign convention is handled in one place (`safe_float`) for both Indian-comma format and `Cr` / `Dr` suffixes
- Periods must match across branches before consolidation runs
- 26 post-load checks flag the common Tally pitfalls (BS not balancing, P&L not reconciling with the P&L A/c ledger, missing depreciation, stock anomalies, mis-classified groups)

**What still depends on you:**
- The TSF Exporter must have run cleanly against a closed period in Tally (mid-year exports get a sub-12-month warning, not an error)
- Tally's group hierarchy should follow the standard (`Sales Accounts`, `Bank Accounts`, etc.) — non-standard groups get auto-inferred via 28 keyword rules, but you can override in the GUI
- Tax expense uses the Deferred Tax Liability ledger as a proxy on Actuals; complex tax workings need a manual adjustment
- Depreciation comes from the Fixed Assets schedule in Tally — if your firm uses an external depreciation workbook, that line will be zero unless booked back in

In short: the data path is well-validated, but the source data must reflect a year-end book close. The Validation sheet tells you when it doesn't.

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
