# Tally Financial Statements Generator

A standalone Python desktop app that reads a **Tally SQLite export** (produced by the [FinAnalyzer TSF Exporter](https://github.com/dhruvdua88/finanalyzer/tree/main/python-tsf-exporter)) and generates publication-ready **Schedule III financial statements** in Excel — with linked notes, 3-year projections, and a full data-validation report.

> **Requires the new TSF SQLite schema** — exported by the bundled TSF Exporter (v2+). The schema uses `mst_ledger`, `mst_group`, and `_export_info` tables with `closing_balance` stored as TEXT in Tally's native sign convention.

---

## Download

**[⬇ Download latest release (ZIP)](https://github.com/dhruvdua88/tally-fin-statements/releases/latest)**

The ZIP contains everything you need: the app, a one-click Mac launcher, and a one-click Windows launcher. No coding required.

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
| **Balance Sheet** | Schedule III format (Indian Companies Act 2013), note numbers are clickable hyperlinks to dedicated note sheets |
| **P&L Statement** | Revenue, purchases, employee costs, finance costs, depreciation, other expenses; key ratios |
| **Note Sheets** | 9 separate Excel sheets (Share Capital → Cash & Bank), each with a "← Back" link |
| **Notes Index** | One-page index of all notes with amounts and hyperlinks |
| **3-Year Projections** | Projected P&L + Balance Sheet with cash-plug method; user inputs growth %, margins, working-capital days, CapEx |
| **Assumptions Sheet** | All projection inputs recorded for audit trail |
| **Validation Report** | 15+ automated checks — BS balance equation, P&L reconciliation, sign flips, natural-sign breaches, stale balances, unclassified groups |
| **Group Mapping** | Auto-infers Schedule III head for any non-standard Tally primary group; user can override via GUI dropdown |
| **Stock Overrides** | Opening and closing stock can be corrected without editing Tally |
| **GUI** | `tkinter` desktop app — no web server, no installation beyond `pip install openpyxl` |

---

## Requirements

- Python 3.11+
- `openpyxl` — `pip install openpyxl`
- A Tally SQLite export produced by the TSF Exporter (see [TSF Schema](#tsf-schema) below)

---

## Quick Start

```bash
pip install openpyxl
python financial_statements.py
```

1. Click **Browse…** and select your `.sqlite` Tally export file
2. Optionally enter corrected opening/closing stock values
3. Click **Apply & Preview Numbers** to see a live balance-sheet and P&L summary
4. Check the **Validation** tab for any data-quality warnings
5. Use the **Group Mapping** tab to assign non-standard Tally groups to Schedule III heads
6. Click **Generate Actual Statements (Excel)** — or fill in the **3-Year Projections** tab and click **Generate Projected Statements + Actual**

The output Excel file is saved to your chosen output folder (default: Desktop).

---

## Excel Output Structure

| Sheet | Contents |
|---|---|
| `Balance Sheet` | Schedule III face — note numbers link to note sheets |
| `P&L Statement` | Schedule III face + key financial ratios |
| `N1 Share Capital` | Ledger-level detail for Share Capital |
| `N2 Reserves Surplus` | Reserves & Surplus breakdown |
| `N3 LT Borrowings` | Secured / Unsecured loans by ledger |
| `N4 ST Borrowings` | Bank OD / CC accounts |
| `N5 Trade Payables` | Creditors by parent sub-group |
| `N8 Fixed Assets` | 7-column schedule: Gross Open → Net Block |
| `N11 Inventories` | Opening and closing stock |
| `N12 Trade Receivables` | Debtors by parent sub-group |
| `N13 Cash & Bank` | Cash-in-hand + bank accounts (debit-balance only) |
| `Notes Index` | All notes with amounts and hyperlinks |
| `Projected P&L` | *(if projections enabled)* Base year + Years 1–3 |
| `Projected Balance Sheet` | *(if projections enabled)* Base year + Years 1–3 |
| `Assumptions` | *(if projections enabled)* All 15 projection inputs |
| `Validation` | Color-coded ERROR / WARNING / INFO check results |

Every note sheet has a **← Back to Balance Sheet** hyperlink in cell A2.

---

## TSF Schema

The app reads Tally data exported via the **TSF Exporter** tool (in `../python-tsf-exporter/`). The SQLite file must contain:

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

The 3-year projection uses a **cash-plug balance sheet** method:

1. **P&L:** Revenue grows at user-specified rates; COGS = Revenue × (1 − Gross Margin %); operating expenses compound at `opex_growth_pct`; interest on average borrowings; WDV depreciation on fixed assets
2. **Working capital:** Inventory = COGS / 365 × `inventory_days`; Debtors = Revenue / 365 × `debtor_days`; Creditors = COGS / 365 × `creditor_days`
3. **Cash (plug):** Total assets − all other assets = residual cash, which auto-balances the sheet

Inputs:

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
| Interest Rate % | 14% | Annualised rate on total borrowings |
| CapEx p.a. | 0 | Annual capital expenditure (₹) |
| Depreciation Rate % | 15% | Written-down-value depreciation rate |
| Tax Rate % | 25% | Effective corporate tax rate |
| Other Income p.a. | 0 | Fixed other income per year (₹) |

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
fin-statements/
├── financial_statements.py   # Single-file app — everything is here
└── README.md                 # This file
```

---

## Limitations

- **Single company, single period** — one SQLite file = one year's statements
- **No previous-year column** — the "Previous Year" column in the BS face is blank (requires two exports)
- **Simplified tax** — tax expense uses the Deferred Tax Liability ledger as a proxy; actual current tax workings are not computed
- **Depreciation from FA schedule** — if Tally has not booked depreciation entries (common when using a separate depreciation workbook), the P&L depreciation line will be zero; add it via a manual adjustment or use the Indirect Expenses sub-group instead
- **Projection accuracy** — projections are management estimates based on user inputs; the cash-plug method means cash absorbs all modelling errors

---

## Licence

MIT — see [LICENSE](../LICENSE)
