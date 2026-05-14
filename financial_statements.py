#!/usr/bin/env python3
"""
Tally Financial Statements Generator
Reads a Tally SQLite export and produces:
  • Schedule III Balance Sheet
  • Statement of Profit & Loss
  • 3-Year Projected P&L and Balance Sheet (banker/investor format)

Run:  python financial_statements.py
Requires: openpyxl  (pip install openpyxl)
"""
import os
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
try:
    import customtkinter as ctk
    _HAS_CTK = True
except ImportError:
    _HAS_CTK = False
from pathlib import Path
from datetime import date
from dataclasses import dataclass, field
from typing import Any

import openpyxl
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.worksheet.page import PageMargins

# ─── Schedule III primary-group classification ────────────────────────────────
# Values: (side, schedule_head, is_asset)
#   side:  'equity' | 'noncurrent_liab' | 'current_liab' | 'noncurrent_asset' | 'current_asset'
#   schedule_head: display label used in grouping
#   is_asset: True if the natural balance is debit (asset)
#
# Tally sign convention in closing_balance:
#   negative → debit balance   (normal for assets / expenses)
#   positive → credit balance  (normal for liabilities / income)
#
# For display we always show amounts as POSITIVE on the face of statements:
#   asset display  =  – closing_balance   (negate the debit)
#   liab display   =  + closing_balance   (credit as-is)
#   revenue display=  + closing_balance   (credit as-is)
#   expense display=  – closing_balance   (negate the debit)

BS_MAP: dict[str, tuple[str, str, bool]] = {
    # ── Assets ───────────────────────────────────────────────────────────────
    "Fixed Assets":             ("noncurrent_asset", "Fixed Assets",                     True),
    "Investments":              ("noncurrent_asset", "Non-Current Investments",           True),
    "Deposits (Asset)":         ("noncurrent_asset", "Long-Term Loans & Advances",        True),
    "Loans & Advances (Asset)": ("noncurrent_asset", "Long-Term Loans & Advances",        True),
    "Misc. Expenses (ASSET)":   ("noncurrent_asset", "Other Non-Current Assets",          True),
    "Stock-in-hand":            ("current_asset",    "Inventories",                       True),
    "Sundry Debtors":           ("current_asset",    "Trade Receivables",                 True),
    "Cash-in-hand":             ("current_asset",    "Cash & Cash Equivalents",           True),
    "Bank Accounts":            ("current_asset",    "Cash & Cash Equivalents",           True),
    "Current Assets":           ("current_asset",    "Other Current Assets",              True),
    # ── Equity ───────────────────────────────────────────────────────────────
    "Capital Account":          ("equity",           "Share Capital",                     False),
    "Reserves & Surplus":       ("equity",           "Reserves & Surplus",                False),
    # ── Non-current liabilities ──────────────────────────────────────────────
    "Secured Loans":            ("noncurrent_liab",  "Long-Term Borrowings",              False),
    "Unsecured Loans":          ("noncurrent_liab",  "Long-Term Borrowings",              False),
    "Loans (Liability)":        ("noncurrent_liab",  "Long-Term Borrowings",              False),
    # ── Current liabilities ──────────────────────────────────────────────────
    "Bank OD A/c":              ("current_liab",     "Short-Term Borrowings",             False),
    "Sundry Creditors":         ("current_liab",     "Trade Payables",                    False),
    "Current Liabilities":      ("current_liab",     "Other Current Liabilities",         False),
    "Branch / Divisions":       ("current_liab",     "Other Current Liabilities",         False),
    "Duties & Taxes":           ("current_liab",     "Duties & Taxes (Net)",              False),
    "Provisions":               ("current_liab",     "Short-Term Provisions",             False),
    "Suspense A/c":             ("current_liab",     "Other Current Liabilities",         False),
}

PNL_MAP: dict[str, tuple[str, str]] = {
    "Sales Accounts":   ("revenue",  "Revenue from Operations"),
    "Direct Incomes":   ("revenue",  "Revenue from Operations"),
    "Indirect Incomes": ("revenue",  "Other Income"),
    "Purchase Accounts":("expense",  "Cost of Materials / Purchases"),
    "Direct Expenses":  ("expense",  "Direct Expenses"),
    "Indirect Expenses":("expense",  "Indirect Expenses"),
}

# Sub-groups of Indirect Expenses that are carved out on the face of P&L
FINANCE_COST_PARENTS = {"Finance Costs", "Interest & Late Filing Fees"}
EMPLOYEE_COST_PARENTS = {
    "Employee benefit expenses", "Contribution to Provident Funds & Others",
    "Salary", "Salaries", "Staff Salary",
}

# ─── Data classes ─────────────────────────────────────────────────────────────

@dataclass
class LedgerRow:
    name: str
    parent: str
    primary_group: str
    opening: float
    closing: float
    is_deemedpositive: bool

@dataclass
class PnlRow:
    """Aggregated P&L transaction total from trn_accounting."""
    primary_group: str
    parent: str           # Tally sub-group (e.g. "Finance Costs")
    total: float          # raw sum; negative = debit (expense/purchase), positive = credit (income)

@dataclass
class FinancialData:
    company: str
    period_from: str
    period_to: str
    period_label: str
    ledgers: list[LedgerRow]
    pnl_rows: list[PnlRow]    # from trn_accounting (kept for future use)
    pnl_balance: float         # P&L A/c closing (cumulative, used in Reserves & Surplus on BS)
    pnl_opening: float         # P&L A/c opening (prior year retained profit not yet in Reserves)
    opening_stock_override: float | None = None
    closing_stock_override: float | None = None

    # ── Balance-sheet ledger helpers ─────────────────────────────────────────

    def _ledgers_for(self, primary_group: str) -> list[LedgerRow]:
        return [l for l in self.ledgers if l.primary_group == primary_group]

    def _sum_closing(self, primary_group: str) -> float:
        return sum(l.closing for l in self._ledgers_for(primary_group))

    def _sum_opening(self, primary_group: str) -> float:
        return sum(l.opening for l in self._ledgers_for(primary_group))

    # ── P&L helpers (use mst_ledger closing_balance — matches Tally's own P&L) ──
    # trn_accounting totals include journal/inter-branch adjustments that Tally
    # excludes from its P&L computation. mst_ledger.closing_balance is the
    # authoritative figure Tally uses.
    # Sign convention for P&L ledgers:
    #   Revenue (Sales, Incomes): positive closing = credit = income as-is
    #   Expenses (Purchases, Costs): negative closing = debit = negate to display positive

    def _pnl_group_total(self, primary_group: str) -> float:
        """Sum of closing_balances for a primary group (matches Tally's P&L)."""
        return sum(l.closing for l in self._ledgers_for(primary_group))

    def _pnl_by_parent(self, primary_group: str) -> dict[str, float]:
        result: dict[str, float] = {}
        for l in self._ledgers_for(primary_group):
            result[l.parent] = result.get(l.parent, 0) + l.closing
        return result

    # ── Fixed assets note ─────────────────────────────────────────────────────

    def fixed_assets_schedule(self) -> list[dict]:
        by_subgroup: dict[str, dict] = {}
        for l in self._ledgers_for("Fixed Assets"):
            sg = l.parent
            if sg not in by_subgroup:
                by_subgroup[sg] = {"gross_open": 0, "gross_close": 0,
                                   "depr_open": 0, "depr_close": 0}
            d = by_subgroup[sg]
            is_depr = "depreciation" in l.name.lower() or "accumulated" in l.name.lower()
            if is_depr:
                d["depr_open"]  += l.opening   # credit = positive
                d["depr_close"] += l.closing
            else:
                d["gross_open"]  += -l.opening  # debit = negative → negate for display
                d["gross_close"] += -l.closing
        rows = []
        for sg, d in by_subgroup.items():
            rows.append({
                "name":        sg,
                "gross_open":  d["gross_open"],
                "additions":   max(0, d["gross_close"] - d["gross_open"]),
                "disposals":   max(0, d["gross_open"] - d["gross_close"]),
                "gross_close": d["gross_close"],
                "depr_open":   d["depr_open"],
                "depr_charge": max(0, d["depr_close"] - d["depr_open"]),
                "depr_close":  d["depr_close"],
                "net_open":    d["gross_open"]  - d["depr_open"],
                "net_close":   d["gross_close"] - d["depr_close"],
            })
        return sorted(rows, key=lambda r: r["name"])

    def net_fixed_assets(self) -> float:
        return sum(r["net_close"] for r in self.fixed_assets_schedule())

    # ── Stock ─────────────────────────────────────────────────────────────────

    def opening_stock(self) -> float:
        if self.opening_stock_override is not None:
            return self.opening_stock_override
        return -self._sum_opening("Stock-in-hand")   # debit balance → negate

    def closing_stock(self) -> float:
        if self.closing_stock_override is not None:
            return self.closing_stock_override
        return -self._sum_closing("Stock-in-hand")

    # ── P&L totals (from mst_ledger closing_balance) ─────────────────────────

    def revenue_from_ops(self) -> float:
        # Revenue accounts: credit nature → positive closing_balance → use as-is
        return (self._pnl_group_total("Sales Accounts")
                + self._pnl_group_total("Direct Incomes"))

    def other_income(self) -> float:
        return self._pnl_group_total("Indirect Incomes")

    def purchases(self) -> float:
        # Expense accounts: debit nature → negative closing_balance → negate to display positive
        return -self._pnl_group_total("Purchase Accounts")

    def direct_expenses(self) -> float:
        return -self._pnl_group_total("Direct Expenses")

    def finance_costs(self) -> float:
        ibd = self._pnl_by_parent("Indirect Expenses")
        return sum(-v for k, v in ibd.items() if k in FINANCE_COST_PARENTS)

    def employee_costs(self) -> float:
        ibd = self._pnl_by_parent("Indirect Expenses")
        return sum(-v for k, v in ibd.items() if k in EMPLOYEE_COST_PARENTS)

    def depreciation_from_fa(self) -> float:
        return sum(r["depr_charge"] for r in self.fixed_assets_schedule())

    def other_indirect_expenses(self) -> float:
        ibd = self._pnl_by_parent("Indirect Expenses")
        skip = FINANCE_COST_PARENTS | EMPLOYEE_COST_PARENTS
        return sum(-v for k, v in ibd.items() if k not in skip)

    def total_expenses(self) -> float:
        return (self.purchases()
                + self.direct_expenses()
                + self.employee_costs()
                + self.finance_costs()
                + self.depreciation_from_fa()
                + self.other_indirect_expenses()
                + self.opening_stock()
                - self.closing_stock())

    def profit_before_tax(self) -> float:
        return self.revenue_from_ops() + self.other_income() - self.total_expenses()

    def tax_expense(self) -> float:
        # Deferred Tax Liability ledger treatment: positive = credit = liability = tax expense
        dtl = next((l.closing for l in self.ledgers
                    if "deferred tax" in l.name.lower()), 0.0)
        return dtl  # simplified; actual tax computation needs I-T workings

    def profit_after_tax(self) -> float:
        return self.profit_before_tax() - self.tax_expense()

    # ── Balance sheet totals ──────────────────────────────────────────────────

    def share_capital(self) -> float:
        # Capital Account with EQUITY SHARE CAPITAL parent (exclude Reserve & Surplus ledger)
        total = 0.0
        for l in self._ledgers_for("Capital Account"):
            if "reserve" not in l.name.lower() and "profit" not in l.name.lower():
                total += l.closing   # credit = positive = equity
        return total

    def reserves_surplus(self) -> float:
        res = self._sum_closing("Reserves & Surplus")
        # Reserve & Surplus ledger under Capital Account
        for l in self._ledgers_for("Capital Account"):
            if "reserve" in l.name.lower():
                res += l.closing
        # Add current year P&L
        res += self.pnl_balance
        return res

    def long_term_borrowings(self) -> float:
        return (self._sum_closing("Secured Loans")
                + self._sum_closing("Unsecured Loans")
                + self._sum_closing("Loans (Liability)"))

    def short_term_borrowings(self) -> float:
        return self._sum_closing("Bank OD A/c")

    def trade_payables(self) -> float:
        return max(0, self._sum_closing("Sundry Creditors"))

    def duties_and_taxes_net(self) -> float:
        net = self._sum_closing("Duties & Taxes")
        # Positive net = payable (liability); negative net = refund (asset)
        return net

    def other_current_liabilities(self) -> float:
        # Duties & Taxes gets its own line; don't include it here
        cl     = self._sum_closing("Current Liabilities")
        branch = self._sum_closing("Branch / Divisions")
        susp   = self._sum_closing("Suspense A/c")
        ocl    = cl + max(0, branch) + susp
        # Remove Deferred Tax Liability from OCL (it's non-current)
        dtl = next((l.closing for l in self.ledgers
                    if "deferred tax" in l.name.lower()), 0.0)
        return max(0, ocl - dtl)

    def short_term_provisions(self) -> float:
        return max(0, self._sum_closing("Provisions"))

    def non_current_investments(self) -> float:
        return -self._sum_closing("Investments")

    def long_term_loans_advances(self) -> float:
        dep = -self._sum_closing("Deposits (Asset)")
        la  = -self._sum_closing("Loans & Advances (Asset)")
        return dep + la

    def other_noncurrent_assets(self) -> float:
        return -self._sum_closing("Misc. Expenses (ASSET)")

    def trade_receivables(self) -> float:
        return max(0, -self._sum_closing("Sundry Debtors"))

    def cash_and_bank(self) -> float:
        # Tally: for is_deemedpositive groups, negative closing = debit = asset (cash/bank balance)
        # Only count banks with a debit (negative) closing as assets here.
        # Banks with credit (positive) closing are picked up in bank_od_in_bank_accounts().
        cash = -self._sum_closing("Cash-in-hand")
        bank_asset = sum(-l.closing for l in self._ledgers_for("Bank Accounts")
                         if l.closing < 0)
        return cash + bank_asset

    def bank_od_in_bank_accounts(self) -> float:
        """Bank Accounts ledgers with credit (positive) closing = additional OD / liability."""
        return sum(l.closing for l in self._ledgers_for("Bank Accounts") if l.closing > 0)

    def duties_and_taxes_asset(self) -> float:
        """Net GST/TDS refund receivable (when D&T net is debit/negative)."""
        net = self.duties_and_taxes_net()
        return -net if net < 0 else 0.0

    def other_current_assets(self) -> float:
        oca       = -self._sum_closing("Current Assets")
        branch_dr = max(0, -self._sum_closing("Branch / Divisions"))
        dt_asset  = self.duties_and_taxes_asset()
        return oca + branch_dr + dt_asset

    def deferred_tax_asset(self) -> float:
        dtl = next((l.closing for l in self.ledgers
                    if "deferred tax" in l.name.lower()), 0.0)
        return max(0, -dtl)  # if DTL < 0, it's actually a DTA

    def total_noncurrent_assets(self) -> float:
        return (self.net_fixed_assets()
                + self.non_current_investments()
                + self.long_term_loans_advances()
                + self.other_noncurrent_assets()
                + self.deferred_tax_asset())

    def total_current_assets(self) -> float:
        return (self.closing_stock()
                + self.trade_receivables()
                + self.cash_and_bank()
                + self.other_current_assets())

    def total_assets(self) -> float:
        return self.total_noncurrent_assets() + self.total_current_assets()

    def total_equity(self) -> float:
        return self.share_capital() + self.reserves_surplus()

    def deferred_tax_liability(self) -> float:
        dtl = next((l.closing for l in self.ledgers
                    if "deferred tax" in l.name.lower()), 0.0)
        return max(0, dtl)

    def total_noncurrent_liab(self) -> float:
        return self.long_term_borrowings() + self.deferred_tax_liability()

    def total_current_liab(self) -> float:
        return (self.short_term_borrowings()
                + self.bank_od_in_bank_accounts()   # OD banks under Bank Accounts group
                + self.trade_payables()
                + max(0, self.duties_and_taxes_net())
                + self.other_current_liabilities()
                + self.short_term_provisions())

    def total_equity_liabilities(self) -> float:
        return self.total_equity() + self.total_noncurrent_liab() + self.total_current_liab()


# ─── Validation ───────────────────────────────────────────────────────────────

ERROR   = "ERROR"
WARNING = "WARNING"
INFO    = "INFO"

@dataclass
class Check:
    severity: str     # ERROR | WARNING | INFO
    category: str     # Schema | Balance | P&L | Data Quality | Classification
    message: str
    detail: str = ""

class ValidationResult:
    def __init__(self) -> None:
        self.checks: list[Check] = []

    def error(self, cat: str, msg: str, detail: str = "") -> None:
        self.checks.append(Check(ERROR, cat, msg, detail))

    def warning(self, cat: str, msg: str, detail: str = "") -> None:
        self.checks.append(Check(WARNING, cat, msg, detail))

    def info(self, cat: str, msg: str, detail: str = "") -> None:
        self.checks.append(Check(INFO, cat, msg, detail))

    @property
    def has_errors(self) -> bool:
        return any(c.severity == ERROR for c in self.checks)

    @property
    def errors(self) -> list[Check]:
        return [c for c in self.checks if c.severity == ERROR]

    @property
    def warnings(self) -> list[Check]:
        return [c for c in self.checks if c.severity == WARNING]

    def summary(self) -> str:
        e = len(self.errors)
        w = len(self.warnings)
        parts = []
        if e: parts.append(f"{e} error{'s' if e > 1 else ''}")
        if w: parts.append(f"{w} warning{'s' if w > 1 else ''}")
        return ", ".join(parts) if parts else "✓ All checks passed"


# ── Amount parser ─────────────────────────────────────────────────────────────

_INFER_KEYWORDS: list[tuple[list[str], str]] = [
    # P&L — Revenue
    (["sales account", "sales accounts", "revenue", "turnover"], "Sales Accounts"),
    (["direct income", "direct incomes"],                         "Direct Incomes"),
    (["indirect income", "indirect incomes", "other income"],     "Indirect Incomes"),
    # P&L — Expenses
    (["purchase account", "purchase accounts"],                   "Purchase Accounts"),
    (["direct expense", "direct expenses"],                       "Direct Expenses"),
    (["indirect expense", "indirect expenses"],                   "Indirect Expenses"),
    (["expense", "expenditure", "cost", "consumption"],           "Indirect Expenses"),
    # BS — Assets
    (["fixed asset", "tangible asset", "intangible asset",
      "capital work", "cwip"],                                     "Fixed Assets"),
    (["investment"],                                               "Investments"),
    (["stock", "inventory", "inventories"],                       "Stock-in-hand"),
    (["sundry debtor", "trade receivable", "debtor"],             "Sundry Debtors"),
    (["cash-in-hand", "cash in hand", "petty cash"],              "Cash-in-hand"),
    (["bank account", "bank"],                                    "Bank Accounts"),
    (["loan and advance", "loans and advance", "advance"],        "Loans & Advances (Asset)"),
    (["deposit"],                                                  "Deposits (Asset)"),
    (["current asset"],                                           "Current Assets"),
    (["misc. expense", "deferred expense", "preliminary"],        "Misc. Expenses (ASSET)"),
    # BS — Equity
    (["share capital", "equity capital", "paid-up capital",
      "capital account"],                                          "Capital Account"),
    (["reserve", "surplus"],                                      "Reserves & Surplus"),
    # BS — Liabilities
    (["secured loan", "term loan", "vehicle loan"],               "Secured Loans"),
    (["unsecured loan"],                                           "Unsecured Loans"),
    (["bank od", "overdraft", "cc limit", "cash credit",
      "working capital loan"],                                     "Bank OD A/c"),
    (["sundry creditor", "trade payable", "creditor"],            "Sundry Creditors"),
    (["provision"],                                                "Provisions"),
    (["duties", "taxes payable", "gst", "tds payable"],          "Duties & Taxes"),
    (["current liabilit", "other liabilit"],                      "Current Liabilities"),
    (["branch", "division"],                                      "Branch / Divisions"),
    (["suspense"],                                                 "Suspense A/c"),
]

def infer_standard_group(primary: str, parents: set[str]) -> str | None:
    """Heuristically map an unknown primary group to the nearest BS_MAP / PNL_MAP key.
    Returns None if no confident match is found.
    """
    text = f"{primary} {' '.join(parents)}".lower()
    for keywords, target in _INFER_KEYWORDS:
        if any(kw in text for kw in keywords):
            return target
    return None


def safe_float(value: Any, default: float = 0.0) -> float:
    """Parse a Tally TEXT amount field robustly.

    Handles: None, empty string, plain floats, Indian comma format
    (1,23,456.78), Cr/Dr suffixes, and unexpected non-numeric content.

    Tally sign convention (stored in the DB):
      negative → debit balance   positive → credit balance
    The Cr/Dr suffix in some older exports reverses this per usual
    double-entry convention, so we handle both.
    """
    if value is None:
        return default
    s = str(value).strip()
    if not s:
        return default
    # Strip Indian-format commas: "1,23,456.78" → "123456.78"
    s_clean = s.replace(",", "")
    low = s_clean.lower()
    # Handle "123456.78 Cr" / "123456.78 Dr" suffix (some Tally versions)
    if low.endswith("cr"):
        try:
            return float(low[:-2].strip())   # Cr in Tally = credit = positive
        except ValueError:
            return default
    if low.endswith("dr"):
        try:
            return -float(low[:-2].strip())  # Dr in Tally = debit = negative
        except ValueError:
            return default
    try:
        return float(s_clean)
    except (ValueError, TypeError):
        return default


# ── Schema validator ──────────────────────────────────────────────────────────

_REQUIRED_TABLES = {"mst_ledger", "mst_group", "_export_info"}
_OPTIONAL_TABLES = {"trn_accounting", "trn_voucher"}

_REQUIRED_LEDGER_COLS = {"name", "parent", "closing_balance"}
_REQUIRED_GROUP_COLS  = {"name", "primary_group"}


def validate_schema(con: sqlite3.Connection) -> ValidationResult:
    """Check that the database has the structure we expect."""
    vr = ValidationResult()

    existing = {r[0] for r in con.execute(
        "SELECT name FROM sqlite_master WHERE type='table'"
    )}

    for t in _REQUIRED_TABLES:
        if t not in existing:
            vr.error("Schema", f"Required table missing: '{t}'",
                     "This may not be a valid Tally SQLite export from this app.")

    for t in _OPTIONAL_TABLES:
        if t not in existing:
            vr.warning("Schema", f"Optional table absent: '{t}'",
                       "Some cross-checks will be skipped.")

    if "mst_ledger" in existing:
        cols = {r[1] for r in con.execute("PRAGMA table_info(mst_ledger)")}
        for c in _REQUIRED_LEDGER_COLS:
            if c not in cols:
                vr.error("Schema", f"mst_ledger missing column '{c}'")
        if "opening_balance" not in cols:
            vr.warning("Schema", "mst_ledger has no 'opening_balance' column",
                       "Opening stock and prior-year P&L carry-forward will be zero.")

    if "mst_group" in existing:
        cols = {r[1] for r in con.execute("PRAGMA table_info(mst_group)")}
        for c in _REQUIRED_GROUP_COLS:
            if c not in cols:
                vr.error("Schema", f"mst_group missing column '{c}'")

    if "_export_info" in existing:
        cols = {r[1] for r in con.execute("PRAGMA table_info(_export_info)")}
        if "name" not in cols or "value" not in cols:
            vr.warning("Schema", "_export_info table has unexpected structure",
                       "Company name and period may not load correctly.")

    return vr


# ── Financial-data validator ──────────────────────────────────────────────────

def validate_financial_data(fd: "FinancialData") -> ValidationResult:
    """Business-logic checks on the loaded financial data."""
    vr = ValidationResult()

    # ── Balance sheet equation ────────────────────────────────────────────────
    diff = fd.total_assets() - fd.total_equity_liabilities()
    if abs(diff) < 1:
        vr.info("Balance Sheet", "Balance sheet balances to zero. ✓")
    elif abs(diff) < 50_000:
        vr.warning("Balance Sheet",
                   f"Small imbalance: ₹{diff:,.2f}",
                   "Likely a rounding difference in stock or depreciation entries.")
    else:
        vr.error("Balance Sheet",
                 f"Balance sheet does not balance — difference: ₹{diff:,.0f}",
                 "Check for ledgers with unusual group assignments or missing classifications.")

    # ── P&L reconciliation ────────────────────────────────────────────────────
    cy_profit    = fd.pnl_balance - fd.pnl_opening   # current year as Tally computed it
    our_pbt      = fd.profit_before_tax()
    pnl_diff     = abs(our_pbt - cy_profit)
    if pnl_diff < 1:
        vr.info("P&L", f"P&L reconciles exactly with Tally (₹{our_pbt:,.0f}). ✓")
    elif pnl_diff < 50_000:
        vr.warning("P&L",
                   f"Minor P&L gap: computed ₹{our_pbt:,.0f} vs Tally ₹{cy_profit:,.0f} "
                   f"(diff ₹{pnl_diff:,.0f})",
                   "May be caused by rounding in opening/closing stock or minor ledger mismatches.")
    else:
        vr.error("P&L",
                 f"P&L does not reconcile — computed ₹{our_pbt:,.0f}, "
                 f"Tally ₹{cy_profit:,.0f}, diff ₹{pnl_diff:,.0f}",
                 "Some ledgers may be under a group not mapped in BS_MAP or PNL_MAP. "
                 "See 'Unclassified' section below.")

    # ── Unclassified primary groups ───────────────────────────────────────────
    all_pgs = {l.primary_group for l in fd.ledgers if l.primary_group}
    for pg in sorted(all_pgs):
        if pg not in BS_MAP and pg not in PNL_MAP:
            net = sum(l.closing for l in fd.ledgers if l.primary_group == pg)
            if abs(net) > 1_000:
                vr.warning("Classification",
                           f"Unclassified primary group '{pg}': net ₹{net:,.0f}",
                           "Add this group to BS_MAP or PNL_MAP at the top of the script "
                           "to include it in statements.")

    # ── Stock checks ──────────────────────────────────────────────────────────
    if fd.opening_stock() == 0 and fd.closing_stock() > 0 and fd.pnl_balance != 0:
        vr.warning("Data Quality",
                   "Opening stock is zero but closing stock is non-zero",
                   "If this is not a new business, verify the opening stock ledger.")

    rev = fd.revenue_from_ops()
    cs  = fd.closing_stock()
    if rev > 0 and cs > rev * 0.6:
        vr.warning("Data Quality",
                   f"Closing stock ₹{cs:,.0f} is >60% of revenue ₹{rev:,.0f}",
                   "Verify the closing stock value entered is correct.")

    # ── Cash & bank ───────────────────────────────────────────────────────────
    if fd.cash_and_bank() < -10_000:
        vr.warning("Data Quality",
                   f"Net cash & bank is negative (₹{fd.cash_and_bank():,.0f})",
                   "Some bank accounts may have credit balances (OD) that are classified "
                   "under 'Bank Accounts' rather than 'Bank OD A/c'. Review Note 13.")

    # ── Receivables vs revenue ────────────────────────────────────────────────
    tr = fd.trade_receivables()
    if rev > 0 and tr > rev:
        vr.warning("Data Quality",
                   f"Trade receivables ₹{tr:,.0f} exceed annual revenue ₹{rev:,.0f}",
                   "Debtor days > 365. Check for stale/uncollected debtors or "
                   "misclassified items under Sundry Debtors.")

    # ── Capital & equity ──────────────────────────────────────────────────────
    if fd.share_capital() < 0:
        vr.error("Data Quality",
                 f"Share Capital is negative (₹{fd.share_capital():,.0f})",
                 "Capital Account ledgers have a net debit balance. "
                 "Check Capital Account entries in Tally.")

    if fd.reserves_surplus() < -100_000:
        vr.warning("Data Quality",
                   f"Reserves & Surplus is significantly negative (₹{fd.reserves_surplus():,.0f})",
                   "Accumulated losses exceed paid-up capital. May be technically insolvent.")

    # ── Borrowings sign check ─────────────────────────────────────────────────
    if fd.long_term_borrowings() < 0:
        vr.warning("Data Quality",
                   f"Long-term borrowings net is negative (₹{fd.long_term_borrowings():,.0f})",
                   "Secured/Unsecured Loan ledgers have a net debit balance. "
                   "Check if loan repayments exceeded drawdowns or if ledgers are miscoded.")

    # ── Period check ──────────────────────────────────────────────────────────
    try:
        pf = date.fromisoformat(fd.period_from)
        pt = date.fromisoformat(fd.period_to)
        days = (pt - pf).days
        if days < 300 or days > 400:
            vr.warning("Data Quality",
                       f"Period is {days} days ({fd.period_from} → {fd.period_to})",
                       "Expected a 12-month financial year (365 days). "
                       "Projections assume a full year as the base period.")
    except (ValueError, TypeError):
        vr.warning("Schema", "Could not parse period dates from _export_info.")

    # ── Significant ledgers with no BS/PL classification ─────────────────────
    unclass = [l for l in fd.ledgers
               if l.primary_group not in BS_MAP
               and l.primary_group not in PNL_MAP
               and abs(l.closing) > 50_000]
    if unclass:
        examples = "; ".join(f"{l.name} (₹{l.closing:,.0f})" for l in unclass[:4])
        vr.warning("Classification",
                   f"{len(unclass)} ledger(s) with material balances are unclassified",
                   f"Examples: {examples}")

    # ── Per-ledger sign/stale checks (from BalanceSheetCleanlinessAnalytics) ──
    asset_groups = {pg for pg, (side, _, _) in BS_MAP.items() if side in ("noncurrent_asset", "current_asset")}
    liab_groups  = {pg for pg, (side, _, _) in BS_MAP.items() if side in ("equity", "noncurrent_liab", "current_liab")}

    sign_flips, natural_breaches, stale = [], [], []
    for l in fd.ledgers:
        # Sign flip: opening and closing have opposite signs
        if abs(l.opening) > 1 and abs(l.closing) > 1 and l.opening * l.closing < 0:
            sign_flips.append(l.name)

        # Natural sign breach: asset with credit balance or liability with debit balance
        if l.primary_group in asset_groups and l.closing > 1_000:
            natural_breaches.append(f"{l.name} (asset with credit balance ₹{l.closing:,.0f})")
        if l.primary_group in liab_groups and l.closing < -1_000:
            natural_breaches.append(f"{l.name} (liability with debit balance ₹{l.closing:,.0f})")

        # Stale: large balance but opening == closing (no movement)
        if abs(l.closing) >= 100_000 and abs(l.opening - l.closing) < 1 and l.opening != 0:
            stale.append(f"{l.name} (₹{l.closing:,.0f})")

    if sign_flips:
        examples = "; ".join(sign_flips[:3])
        vr.warning("Data Quality",
                   f"{len(sign_flips)} ledger(s) have sign flips (opening/closing opposite signs)",
                   f"Examples: {examples}")
    if natural_breaches:
        examples = "; ".join(natural_breaches[:3])
        vr.warning("Data Quality",
                   f"{len(natural_breaches)} ledger(s) breach their natural sign",
                   f"Examples: {examples}")
    if stale:
        examples = "; ".join(stale[:3])
        vr.info("Data Quality",
                f"{len(stale)} ledger(s) have large stale balances (no movement in period)",
                f"Examples: {examples}")

    return vr


# ─── Database loader ──────────────────────────────────────────────────────────

def load_from_sqlite(db_path: str,
                     opening_stock_override: float | None = None,
                     closing_stock_override: float | None = None,
                     schema_vr: ValidationResult | None = None,
                     reclassify_map: dict[str, str] | None = None) -> FinancialData:
    """Load financial data from a Tally SQLite export.

    Raises RuntimeError if schema validation finds blocking errors.
    Non-fatal warnings are accumulated into schema_vr (passed in or a new one).
    """
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row

    # Schema validation — fail fast on missing tables / columns
    sv = validate_schema(con)
    if schema_vr is not None:
        schema_vr.checks.extend(sv.checks)
    if sv.has_errors:
        con.close()
        msgs = "; ".join(c.message for c in sv.errors)
        raise RuntimeError(f"Database schema errors: {msgs}")

    # Company / period (_export_info uses name/value columns)
    try:
        info = {r["name"]: r["value"]
                for r in con.execute("SELECT name, value FROM _export_info")}
    except Exception:
        info = {}
    company     = info.get("company_name", "Company").replace("(from", "").replace(")", "").strip()
    period_from = info.get("period_from", "")
    period_to   = info.get("period_to",   "")

    # Build period label e.g. "31st March 2026"
    def _ordinal(dt_str: str) -> str:
        try:
            d = date.fromisoformat(dt_str)
            day = d.day
            suffix = {1:"st",2:"nd",3:"rd"}.get(day if day < 20 else day % 10, "th")
            return f"{day}{suffix} {d.strftime('%B %Y')}"
        except Exception:
            return dt_str
    period_label = _ordinal(period_to)

    # Detect whether opening_balance column exists (older exports may omit it)
    ledger_cols = {r[1] for r in con.execute("PRAGMA table_info(mst_ledger)")}
    has_opening = "opening_balance" in ledger_cols

    # Profit & Loss A/c (special ledger, no parent group)
    # closing_balance = cumulative (prior year opening + current year profit)
    # opening_balance = prior year retained earnings not yet transferred to Reserves
    pnl_balance = 0.0
    pnl_opening = 0.0
    try:
        if has_opening:
            pnl_row = con.execute(
                "SELECT opening_balance, closing_balance FROM mst_ledger "
                "WHERE name = 'Profit & Loss A/c'"
            ).fetchone()
            if pnl_row:
                pnl_balance = safe_float(pnl_row["closing_balance"])
                pnl_opening = safe_float(pnl_row["opening_balance"])
        else:
            pnl_row = con.execute(
                "SELECT closing_balance FROM mst_ledger WHERE name = 'Profit & Loss A/c'"
            ).fetchone()
            if pnl_row:
                pnl_balance = safe_float(pnl_row["closing_balance"])
    except Exception:
        pass   # pnl_balance stays 0

    # Balance-sheet ledgers from mst_ledger (closing balances = year-end positions)
    open_col = "l.opening_balance" if has_opening else "'0'"
    bs_rows = con.execute(f"""
        SELECT
            l.name,
            l.parent,
            g.primary_group,
            {open_col} AS opening_raw,
            l.closing_balance AS closing_raw,
            COALESCE(g.is_deemedpositive, '0') AS is_dp
        FROM mst_ledger l
        LEFT JOIN mst_group g ON g.name = l.parent
        WHERE g.primary_group IS NOT NULL
        ORDER BY g.primary_group, l.parent, l.name
    """).fetchall()

    # Build parent sets per primary for auto-inference
    _pg_parents: dict[str, set[str]] = {}
    for r in bs_rows:
        pg = r["primary_group"]
        if pg:
            _pg_parents.setdefault(pg, set()).add(str(r["parent"] or ""))

    effective_remap = dict(reclassify_map or {})

    ledgers = []
    for r in bs_rows:
        pg = r["primary_group"]
        # Apply user reclassification first
        if pg in effective_remap:
            pg = effective_remap[pg]
        # Auto-infer for groups still not in any map
        elif pg not in BS_MAP and pg not in PNL_MAP:
            inferred = infer_standard_group(pg, _pg_parents.get(r["primary_group"], set()))
            if inferred:
                pg = inferred
        ledgers.append(LedgerRow(
            name=r["name"],
            parent=r["parent"],
            primary_group=pg,
            opening=safe_float(r["opening_raw"]),
            closing=safe_float(r["closing_raw"]),
            is_deemedpositive=str(r["is_dp"]) == "1",
        ))

    # P&L activity from trn_accounting (kept for future use / cross-checks).
    # The primary P&L computation uses mst_ledger closing_balance instead
    # because trn_accounting includes journal/inter-branch entries that Tally
    # excludes from its own P&L report.
    pnl_groups = (
        "'Sales Accounts','Direct Incomes','Indirect Incomes',"
        "'Purchase Accounts','Direct Expenses','Indirect Expenses'"
    )
    try:
        pnl_txn_rows = con.execute(f"""
            SELECT
                g.primary_group,
                l.parent,
                SUM(CAST(a.amount AS REAL)) AS total
            FROM trn_accounting a
            JOIN mst_ledger l ON l.name = a.ledger
            LEFT JOIN mst_group g ON g.name = l.parent
            WHERE g.primary_group IN ({pnl_groups})
            GROUP BY g.primary_group, l.parent
            ORDER BY g.primary_group, l.parent
        """).fetchall()
        pnl_rows = [
            PnlRow(
                primary_group=r["primary_group"],
                parent=r["parent"],
                total=float(r["total"] or 0),
            )
            for r in pnl_txn_rows
        ]
    except Exception:
        pnl_rows = []   # trn_accounting absent or malformed — non-fatal

    con.close()
    return FinancialData(
        company=company,
        period_from=period_from,
        period_to=period_to,
        period_label=period_label,
        ledgers=ledgers,
        pnl_rows=pnl_rows,
        pnl_balance=pnl_balance,
        pnl_opening=pnl_opening,
        opening_stock_override=opening_stock_override,
        closing_stock_override=closing_stock_override,
    )


# ─── Excel generation ─────────────────────────────────────────────────────────

# Colour palette
C_DARK_BLUE  = "1F3864"
C_MID_BLUE   = "2F5496"
C_LIGHT_BLUE = "D6E4F0"
C_HEADER_BG  = "1F3864"
C_SUBHD_BG   = "BDD7EE"
C_TOTAL_BG   = "D9E1F2"
C_WHITE      = "FFFFFF"
C_BLACK      = "000000"
C_AMBER      = "FFF2CC"

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _border(style: str = "thin") -> Border:
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def _font(bold=False, size=10, color=C_BLACK, name="Calibri") -> Font:
    return Font(bold=bold, size=size, color=color, name=name)

def _align(h="left", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

INR = '#,##0;(#,##0);"-"'   # accounting style: negatives in parens, zero as dash
INR2 = '#,##0.00;(#,##0.00);"-"'

def _fmt(ws, cell_ref: str, value: float | None, italic: bool = False) -> None:
    cell = ws[cell_ref]
    if value is not None:
        cell.value = value
    cell.number_format = INR
    if italic:
        cell.font = Font(name="Calibri", size=10, italic=True)


class ExcelWriter:
    def __init__(self, fd: FinancialData, proj: "ProjectionInputs | None" = None,
                 opts: "OutputOptions | None" = None):
        self.fd = fd
        self.proj = proj
        self.opts = opts or OutputOptions()
        self.wb = openpyxl.Workbook()
        self.wb.remove(self.wb.active)  # remove default sheet

    def _apply_page_setup(self, ws, *, landscape: bool = False, fit_height: int = 0) -> None:
        """Fit to one page width, height auto (or capped). Adds A4 print setup."""
        if not self.opts.page_setup:
            return
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = (ws.ORIENTATION_LANDSCAPE if landscape
                                     else ws.ORIENTATION_PORTRAIT)
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = fit_height  # 0 = auto
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.6, bottom=0.6,
                                       header=0.3, footer=0.3)
        ws.print_options.horizontalCentered = True
        ws.oddHeader.center.text = f"{self.fd.company}"
        ws.oddHeader.center.size = 10
        ws.oddFooter.right.text = "Page &P of &N"
        ws.oddFooter.right.size = 9
        ws.oddFooter.left.text = f"Generated {date.today().strftime('%d %b %Y')}"
        ws.oddFooter.left.size = 9

    def save(self, path: str) -> None:
        # Write note sheets FIRST so BS can reference their total cells via formulas.
        self.note_refs: dict[int, str] = {}
        self.note_refs[1]  = self._write_note_share_capital()
        self.note_refs[2]  = self._write_note_reserves()
        self.note_refs[3]  = self._write_note_lt_borrowings()
        self.note_refs[4]  = self._write_note_st_borrowings()
        self.note_refs[5]  = self._write_note_trade_payables()
        self.note_refs[8]  = self._write_note_fixed_assets()
        self.note_refs[11] = self._write_note_inventories()
        self.note_refs[12] = self._write_note_trade_receivables()
        self.note_refs[13] = self._write_note_cash_bank()
        # Now BS + PnL (can reference note total cells)
        self._write_bs()
        self._write_pnl()
        self._write_notes_index()
        pe = None
        if self.proj:
            pe = ProjectionEngine(self.fd, self.proj)
            self._write_proj_pnl(pe)
            self._write_proj_bs(pe)
            self._write_assumptions()
            if self.opts.cash_flow:
                self._write_cash_flow(pe)
        if self.opts.ratios:
            self._write_ratios(pe)
        if self.opts.common_size:
            self._write_common_size(pe)
        if self.opts.charts:
            self._write_charts(pe)
        vr = validate_financial_data(self.fd)
        self._write_validation(vr)
        # Apply page setup to every sheet
        if self.opts.page_setup:
            wide_sheets = {"Projected P&L", "Projected Balance Sheet", "Cash Flow",
                           "Ratios", "Common-Size", "N8 Fixed Assets", "Validation"}
            for ws in self.wb.worksheets:
                self._apply_page_setup(ws, landscape=(ws.title in wide_sheets))
        # Reorder + activate the Balance Sheet so the file opens to it
        self._reorder_sheets()
        self.wb.save(path)

    def _reorder_sheets(self) -> None:
        preferred = ["Balance Sheet", "P&L Statement",
                     "Projected Balance Sheet", "Projected P&L",
                     "Assumptions", "Notes Index"]
        order: list[str] = []
        for name in preferred:
            if name in self.wb.sheetnames:
                order.append(name)
        for name in self.wb.sheetnames:
            if name not in order:
                order.append(name)
        self.wb._sheets = [self.wb[n] for n in order]
        if "Balance Sheet" in self.wb.sheetnames:
            self.wb.active = self.wb.sheetnames.index("Balance Sheet")

    # ── Validation sheet ──────────────────────────────────────────────────────

    def _write_validation(self, vr: ValidationResult) -> None:
        ws = self.wb.create_sheet("Validation")
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 50
        ws.column_dimensions["D"].width = 60

        # Header
        ws.merge_cells("A1:D1")
        ws["A1"].value = f"DATA VALIDATION REPORT — {self.fd.company} — {self.fd.period_label}"
        ws["A1"].font = _font(bold=True, size=12, color=C_WHITE)
        ws["A1"].fill = _fill(C_HEADER_BG)
        ws["A1"].alignment = _align("center")

        ws["A2"].value = vr.summary()
        ws["A2"].font = _font(bold=True, size=11,
                              color="CC0000" if vr.has_errors else "2E7D32")
        ws.merge_cells("A2:D2")
        ws["A2"].alignment = _align("center")

        # Column headers
        for col, lbl in [("A", "Severity"), ("B", "Category"),
                          ("C", "Message"), ("D", "Detail")]:
            ws[f"{col}3"].value = lbl
            ws[f"{col}3"].font = _font(bold=True, size=9, color=C_WHITE)
            ws[f"{col}3"].fill = _fill(C_MID_BLUE)
            ws[f"{col}3"].alignment = _align("center")
            ws[f"{col}3"].border = _border()

        _SEVERITY_COLORS = {ERROR: "FFCCCC", WARNING: C_AMBER, INFO: "E8F5E9"}
        _SEVERITY_FG     = {ERROR: "CC0000", WARNING: "7B5800", INFO: "1B5E20"}

        for i, chk in enumerate(vr.checks, start=4):
            bg = _SEVERITY_COLORS.get(chk.severity, C_WHITE)
            fg = _SEVERITY_FG.get(chk.severity, C_BLACK)
            for col, val in [("A", chk.severity), ("B", chk.category),
                              ("C", chk.message),  ("D", chk.detail)]:
                cell = ws[f"{col}{i}"]
                cell.value = val
                cell.fill = _fill(bg)
                cell.font = _font(size=9,
                                  bold=(col == "A"),
                                  color=fg if col == "A" else C_BLACK)
                cell.alignment = _align("left", wrap=True)
                cell.border = _border()
            ws.row_dimensions[i].height = 28

        ws.freeze_panes = "A4"

    # ── Balance Sheet ─────────────────────────────────────────────────────────

    def _write_bs(self) -> None:
        ws = self.wb.create_sheet("Balance Sheet")
        fd = self.fd
        ws.column_dimensions["A"].width = 46
        ws.column_dimensions["B"].width = 8
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 18

        r = 1
        def header(text, span=4, bg=C_HEADER_BG, fg=C_WHITE, sz=12, bold=True):
            nonlocal r
            ws.merge_cells(f"A{r}:D{r}")
            cell = ws[f"A{r}"]
            cell.value = text
            cell.font = _font(bold=bold, size=sz, color=fg)
            cell.fill = _fill(bg)
            cell.alignment = _align("center")
            r += 1

        def subheader(text, bg=C_SUBHD_BG):
            nonlocal r
            ws.merge_cells(f"A{r}:D{r}")
            cell = ws[f"A{r}"]
            cell.value = text
            cell.font = _font(bold=True, size=10)
            cell.fill = _fill(bg)
            cell.alignment = _align("left")
            r += 1

        def col_header():
            nonlocal r
            ws[f"A{r}"].value = "Particulars"
            ws[f"B{r}"].value = "Note"
            ws[f"C{r}"].value = f"As at {fd.period_label}"
            ws[f"D{r}"].value = f"Previous Year"
            for col in "ABCD":
                c = ws[f"{col}{r}"]
                c.font = _font(bold=True, size=9, color=C_WHITE)
                c.fill = _fill(C_MID_BLUE)
                c.alignment = _align("center")
                c.border = _border()
            r += 1

        def row(label, amount, note=None, indent=0, bold=False, total=False, fmt=INR):
            nonlocal r
            prefix = "    " * indent
            ws[f"A{r}"].value = prefix + label
            ws[f"A{r}"].font = _font(bold=bold or total, size=10)
            ws[f"A{r}"].alignment = _align("left")
            if note is not None:
                note_num = int(note)
                note_sheet_name = self._note_sheet_name(note_num)
                ws[f"B{r}"].value = note_num
                ws[f"B{r}"].hyperlink = f"#'{note_sheet_name}'!A1"
                ws[f"B{r}"].font = Font(name="Calibri", size=9, color="1D4ED8", underline="single")
                ws[f"B{r}"].alignment = _align("center")
                # Prefer formula link to the note's total cell if available
                ref = self.note_refs.get(note_num) if hasattr(self, "note_refs") else None
                if ref:
                    ws[f"C{r}"].value = f"={ref}"
                elif amount is not None:
                    ws[f"C{r}"].value = amount
                ws[f"C{r}"].number_format = fmt
                ws[f"C{r}"].font = _font(bold=bold or total, size=10)
                ws[f"C{r}"].alignment = _align("right")
            elif amount is not None:
                ws[f"C{r}"].value = amount
                ws[f"C{r}"].number_format = fmt
                ws[f"C{r}"].font = _font(bold=bold or total, size=10)
                ws[f"C{r}"].alignment = _align("right")
            # Always ensure the amount cell has the right format/alignment/weight,
            # even on total rows whose formula is filled in by the caller after row().
            ws[f"C{r}"].number_format = fmt
            ws[f"C{r}"].alignment = _align("right")
            if bold or total:
                ws[f"A{r}"].font = _font(bold=True, size=10)
                ws[f"C{r}"].font = _font(bold=True, size=10)
            # Light borders on every data row for a clean tabular look
            for col in "ABCD":
                ws[f"{col}{r}"].border = _border()
            if total:
                for col in "ABCD":
                    ws[f"{col}{r}"].fill = _fill(C_TOTAL_BG)
            r += 1
            return r - 1  # return row number for formula references

        def spacer():
            nonlocal r
            r += 1

        header(fd.company.upper(), sz=14)
        header("BALANCE SHEET", sz=12)
        header(f"As at {fd.period_label}", sz=10, bg=C_MID_BLUE)
        header("(Amount in ₹)", sz=9, bg=C_LIGHT_BLUE, fg=C_BLACK, bold=False)
        col_header()

        # ── EQUITY & LIABILITIES ──────────────────────────────────────────────
        subheader("I.  SHAREHOLDERS' FUNDS")
        sc_r  = row("a)  Share Capital",           fd.share_capital(),       note="1",  indent=1)
        res_r = row("b)  Reserves & Surplus",      fd.reserves_surplus(),    note="2",  indent=1)
        sf_tot_r = row("Total Shareholders' Funds", None, bold=True, total=True)
        ws[f"C{sf_tot_r}"].value = f"=C{sc_r}+C{res_r}"
        ws[f"C{sf_tot_r}"].number_format = INR
        spacer()

        subheader("II.  NON-CURRENT LIABILITIES")
        ltb_r = row("a)  Long-Term Borrowings",    fd.long_term_borrowings(),  note="3", indent=1)
        dtl_r = row("b)  Deferred Tax Liability",  fd.deferred_tax_liability(), indent=1)
        ncl_tot_r = row("Total Non-Current Liabilities", None, bold=True, total=True)
        ws[f"C{ncl_tot_r}"].value = f"=C{ltb_r}+C{dtl_r}"
        ws[f"C{ncl_tot_r}"].number_format = INR
        spacer()

        subheader("III.  CURRENT LIABILITIES")
        stb_r  = row("a)  Short-Term Borrowings",      fd.short_term_borrowings(), note="4", indent=1)
        tp_r   = row("b)  Trade Payables",             fd.trade_payables(),         note="5", indent=1)
        dt_r   = row("c)  Duties & Taxes (Net)",       max(0, fd.duties_and_taxes_net()), indent=1)
        ocl_r  = row("d)  Other Current Liabilities",  fd.other_current_liabilities(), indent=1)
        prov_r = row("e)  Short-Term Provisions",      fd.short_term_provisions(), indent=1)
        cl_tot_r = row("Total Current Liabilities", None, bold=True, total=True)
        ws[f"C{cl_tot_r}"].value = f"=C{stb_r}+C{tp_r}+C{dt_r}+C{ocl_r}+C{prov_r}"
        ws[f"C{cl_tot_r}"].number_format = INR
        spacer()

        # Grand total E&L  (formula = sum of subtotals)
        tel_r = row("TOTAL EQUITY & LIABILITIES", None, bold=True)
        ws[f"C{tel_r}"].value = f"=C{sf_tot_r}+C{ncl_tot_r}+C{cl_tot_r}"
        ws[f"C{tel_r}"].number_format = INR
        for col in "ABCD":
            ws[f"{col}{tel_r}"].fill = _fill(C_HEADER_BG)
            ws[f"{col}{tel_r}"].font = Font(bold=True, size=11, color=C_WHITE, name="Calibri")
        ws[f"C{tel_r}"].alignment = _align("right")
        spacer(); spacer()

        # ── ASSETS ───────────────────────────────────────────────────────────
        subheader("I.  NON-CURRENT ASSETS")
        fa_r   = row("a)  Fixed Assets (Net Block)",     fd.net_fixed_assets(),          note="8", indent=1)
        inv_r  = row("b)  Non-Current Investments",      fd.non_current_investments(),   indent=1)
        lla_r  = row("c)  Long-Term Loans & Advances",   fd.long_term_loans_advances(),  indent=1)
        dta_r  = row("d)  Deferred Tax Asset",           fd.deferred_tax_asset(),        indent=1)
        ona_r  = row("e)  Other Non-Current Assets",     fd.other_noncurrent_assets(),   indent=1)
        nca_tot_r = row("Total Non-Current Assets", None, bold=True, total=True)
        ws[f"C{nca_tot_r}"].value = f"=C{fa_r}+C{inv_r}+C{lla_r}+C{dta_r}+C{ona_r}"
        ws[f"C{nca_tot_r}"].number_format = INR
        spacer()

        subheader("II.  CURRENT ASSETS")
        stk_r  = row("a)  Inventories (Closing Stock)",  fd.closing_stock(),             note="11", indent=1)
        tr_r   = row("b)  Trade Receivables",            fd.trade_receivables(),         note="12", indent=1)
        cb_r   = row("c)  Cash & Cash Equivalents",      fd.cash_and_bank(),             note="13", indent=1)
        oca_r  = row("d)  Other Current Assets",         fd.other_current_assets(),      indent=1)
        ca_tot_r = row("Total Current Assets", None, bold=True, total=True)
        ws[f"C{ca_tot_r}"].value = f"=C{stk_r}+C{tr_r}+C{cb_r}+C{oca_r}"
        ws[f"C{ca_tot_r}"].number_format = INR
        spacer()

        ta_r = row("TOTAL ASSETS", None, bold=True)
        ws[f"C{ta_r}"].value = f"=C{nca_tot_r}+C{ca_tot_r}"
        ws[f"C{ta_r}"].number_format = INR
        for col in "ABCD":
            ws[f"{col}{ta_r}"].fill = _fill(C_HEADER_BG)
            ws[f"{col}{ta_r}"].font = Font(bold=True, size=11, color=C_WHITE, name="Calibri")
        ws[f"C{ta_r}"].alignment = _align("right")
        spacer(); spacer()

        # ── Difference check (formula-driven) ────────────────────────────────
        diff_r = row("Balance Sheet Difference (should be 0)", None, bold=True)
        ws[f"C{diff_r}"].value = f"=C{ta_r}-C{tel_r}"
        ws[f"C{diff_r}"].number_format = INR
        # Conditional-format the diff cell via a simple Python-side check on current data
        if abs(fd.total_assets() - fd.total_equity_liabilities()) > 1:
            ws[f"C{diff_r}"].fill = _fill("FF0000")
            ws[f"C{diff_r}"].font = Font(bold=True, size=10, color=C_WHITE)
        spacer(); spacer()

        ws.freeze_panes = "A6"

    # ── Note sheet helpers ────────────────────────────────────────────────────

    def _note_sheet_name(self, note_num: int | str) -> str:
        """Return the sheet name for a given note number."""
        _NOTE_TITLES = {
            1: "N1 Share Capital",
            2: "N2 Reserves Surplus",
            3: "N3 LT Borrowings",
            4: "N4 ST Borrowings",
            5: "N5 Trade Payables",
            6: "N6 Other CL",
            7: "N7 Provisions",
            8: "N8 Fixed Assets",
            9: "N9 NC Investments",
            10: "N10 LT Loans",
            11: "N11 Inventories",
            12: "N12 Trade Receivables",
            13: "N13 Cash & Bank",
            14: "N14 Other CA",
        }
        return _NOTE_TITLES.get(int(note_num), f"Note {note_num}")

    def _start_note_sheet(self, note_num: int, title: str) -> tuple:
        """Create a note sheet, write header + back link, return (ws, next_row)."""
        sheet_name = self._note_sheet_name(note_num)
        ws = self.wb.create_sheet(sheet_name)
        ws.column_dimensions["A"].width = 46
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 20

        # Title row
        ws.merge_cells("A1:C1")
        ws["A1"].value = f"Note {note_num}:  {title}"
        ws["A1"].font = Font(name="Calibri", bold=True, size=11, color=C_WHITE)
        ws["A1"].fill = _fill(C_HEADER_BG)
        ws["A1"].alignment = _align("center")

        # Back link
        ws.merge_cells("A2:C2")
        ws["A2"].value = "<- Back to Balance Sheet"
        ws["A2"].hyperlink = "#'Balance Sheet'!A1"
        ws["A2"].font = Font(name="Calibri", size=9, color="1D4ED8", underline="single")
        ws["A2"].alignment = _align("left")

        return ws, 3   # next row to write at

    def _note_data_row(self, ws, r: int, label: str, amount: float | None = None,
                       bold: bool = False, total: bool = False, indent: int = 0) -> int:
        prefix = "    " * indent
        ws[f"A{r}"].value = prefix + label
        ws[f"A{r}"].font = _font(bold=bold or total, size=9)
        if amount is not None:
            ws[f"B{r}"].value = amount
            ws[f"B{r}"].number_format = INR
            ws[f"B{r}"].font = _font(bold=bold or total, size=9)
            ws[f"B{r}"].alignment = _align("right")
        if total:
            ws[f"A{r}"].fill = _fill(C_TOTAL_BG)
            ws[f"B{r}"].fill = _fill(C_TOTAL_BG)
        return r + 1

    def _note_total_ref(self, note_num: int, row: int) -> str:
        return f"'{self._note_sheet_name(note_num)}'!B{row}"

    def _write_note_share_capital(self) -> str:
        ws, r = self._start_note_sheet(1, "Share Capital")
        fd = self.fd
        for l in fd._ledgers_for("Capital Account"):
            if "reserve" not in l.name.lower() and "profit" not in l.name.lower():
                r = self._note_data_row(ws, r, l.name, l.closing)
        self._note_data_row(ws, r, "Total Share Capital", fd.share_capital(), total=True)
        return self._note_total_ref(1, r)

    def _write_note_reserves(self) -> str:
        ws, r = self._start_note_sheet(2, "Reserves & Surplus")
        fd = self.fd
        for l in fd._ledgers_for("Capital Account"):
            if "reserve" in l.name.lower():
                r = self._note_data_row(ws, r, l.name, l.closing)
        for l in fd._ledgers_for("Reserves & Surplus"):
            r = self._note_data_row(ws, r, l.name, l.closing)
        r = self._note_data_row(ws, r, "Profit for the year (P&L A/c)", fd.pnl_balance)
        self._note_data_row(ws, r, "Total Reserves & Surplus", fd.reserves_surplus(), total=True)
        return self._note_total_ref(2, r)

    def _write_note_lt_borrowings(self) -> str:
        ws, r = self._start_note_sheet(3, "Long-Term Borrowings")
        fd = self.fd
        for pg in ("Secured Loans", "Unsecured Loans", "Loans (Liability)"):
            group_ledgers = fd._ledgers_for(pg)
            if not group_ledgers:
                continue
            r = self._note_data_row(ws, r, pg, bold=True)
            for l in group_ledgers:
                r = self._note_data_row(ws, r, l.name, l.closing, indent=1)
        self._note_data_row(ws, r, "Total Long-Term Borrowings", fd.long_term_borrowings(), total=True)
        return self._note_total_ref(3, r)

    def _write_note_st_borrowings(self) -> str:
        ws, r = self._start_note_sheet(4, "Short-Term Borrowings")
        fd = self.fd
        for l in fd._ledgers_for("Bank OD A/c"):
            r = self._note_data_row(ws, r, l.name, l.closing)
        # Bank accounts with credit balance = OD
        od_banks = [l for l in fd._ledgers_for("Bank Accounts") if l.closing > 0]
        if od_banks:
            r = self._note_data_row(ws, r, "Bank Accounts (credit/OD balance)", bold=True)
            for l in od_banks:
                r = self._note_data_row(ws, r, l.name, l.closing, indent=1)
        self._note_data_row(ws, r, "Total Short-Term Borrowings",
                            fd.short_term_borrowings() + fd.bank_od_in_bank_accounts(), total=True)
        return self._note_total_ref(4, r)

    def _write_note_trade_payables(self) -> str:
        ws, r = self._start_note_sheet(5, "Trade Payables")
        fd = self.fd
        by_parent: dict[str, float] = {}
        for l in fd._ledgers_for("Sundry Creditors"):
            by_parent[l.parent] = by_parent.get(l.parent, 0) + l.closing
        for parent, amt in sorted(by_parent.items()):
            r = self._note_data_row(ws, r, parent, amt if amt > 0 else None)
        self._note_data_row(ws, r, "Total Trade Payables", fd.trade_payables(), total=True)
        return self._note_total_ref(5, r)

    def _write_note_fixed_assets(self) -> str:
        ws, r = self._start_note_sheet(8, "Fixed Assets")
        fd = self.fd
        # Column headers
        for col, lbl in [("A", "Asset Category"), ("B", "Gross Open"), ("C", "Additions"),
                         ("D", "Disposals"), ("E", "Gross Close"), ("F", "Accum Depr"), ("G", "Net Block")]:
            ws.column_dimensions[col].width = 18
            ws[f"{col}{r}"].value = lbl
            ws[f"{col}{r}"].font = _font(bold=True, size=9, color=C_WHITE)
            ws[f"{col}{r}"].fill = _fill(C_MID_BLUE)
            ws[f"{col}{r}"].alignment = _align("center" if col != "A" else "left")
        r += 1
        for fa in fd.fixed_assets_schedule():
            ws[f"A{r}"].value = "  " + fa["name"]
            ws[f"A{r}"].font = _font(size=9)
            for col, key in [("B","gross_open"),("C","additions"),("D","disposals"),
                              ("E","gross_close"),("F","depr_close"),("G","net_close")]:
                ws[f"{col}{r}"].value = fa[key]
                ws[f"{col}{r}"].number_format = INR
                ws[f"{col}{r}"].font = _font(size=9)
                ws[f"{col}{r}"].alignment = _align("right")
            r += 1
        # Total row
        ws[f"A{r}"].value = "  TOTAL"
        ws[f"A{r}"].font = _font(bold=True, size=9)
        ws[f"G{r}"].value = fd.net_fixed_assets()
        ws[f"G{r}"].number_format = INR
        ws[f"G{r}"].font = _font(bold=True, size=9)
        ws[f"G{r}"].fill = _fill(C_TOTAL_BG)
        ws[f"A{r}"].fill = _fill(C_TOTAL_BG)
        return f"'{self._note_sheet_name(8)}'!G{r}"

    def _write_note_inventories(self) -> str:
        ws, r = self._start_note_sheet(11, "Inventories")
        fd = self.fd
        r = self._note_data_row(ws, r, "Opening Stock (as per books / override)", fd.opening_stock())
        r = self._note_data_row(ws, r, "Closing Stock (as per books / override)", fd.closing_stock())
        self._note_data_row(ws, r, "Net Inventory on Balance Sheet", fd.closing_stock(), total=True)
        return self._note_total_ref(11, r)

    def _write_note_trade_receivables(self) -> str:
        ws, r = self._start_note_sheet(12, "Trade Receivables")
        fd = self.fd
        by_parent: dict[str, float] = {}
        for l in fd._ledgers_for("Sundry Debtors"):
            by_parent[l.parent] = by_parent.get(l.parent, 0) + (-l.closing)
        for parent, amt in sorted(by_parent.items()):
            r = self._note_data_row(ws, r, parent, amt if abs(amt) > 0 else None)
        self._note_data_row(ws, r, "Total Trade Receivables", fd.trade_receivables(), total=True)
        return self._note_total_ref(12, r)

    def _write_note_cash_bank(self) -> str:
        ws, r = self._start_note_sheet(13, "Cash & Cash Equivalents")
        fd = self.fd
        if fd._ledgers_for("Cash-in-hand"):
            r = self._note_data_row(ws, r, "Cash-in-Hand", bold=True)
            for l in fd._ledgers_for("Cash-in-hand"):
                r = self._note_data_row(ws, r, l.name, -l.closing, indent=1)
        asset_banks = [l for l in fd._ledgers_for("Bank Accounts") if l.closing < 0]
        if asset_banks:
            r = self._note_data_row(ws, r, "Bank Accounts (debit balance)", bold=True)
            for l in asset_banks:
                r = self._note_data_row(ws, r, l.name, -l.closing, indent=1)
        self._note_data_row(ws, r, "Total Cash & Cash Equivalents", fd.cash_and_bank(), total=True)
        return self._note_total_ref(13, r)

    def _write_notes_index(self):
        ws = self.wb.create_sheet("Notes Index")
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 35
        ws.column_dimensions["C"].width = 20

        ws.merge_cells("A1:C1")
        ws["A1"].value = "NOTES TO FINANCIAL STATEMENTS"
        ws["A1"].font = _font(bold=True, size=11, color=C_WHITE)
        ws["A1"].fill = _fill(C_HEADER_BG)
        ws["A1"].alignment = _align("center")

        for col, lbl in [("A","Note No."),("B","Description"),("C","Amount (Rs.)")]:
            ws[f"{col}2"].value = lbl
            ws[f"{col}2"].font = _font(bold=True, size=9, color=C_WHITE)
            ws[f"{col}2"].fill = _fill(C_MID_BLUE)
            ws[f"{col}2"].alignment = _align("center")

        fd = self.fd
        NOTE_INDEX = [
            (1,  "Share Capital",                  fd.share_capital()),
            (2,  "Reserves & Surplus",             fd.reserves_surplus()),
            (3,  "Long-Term Borrowings",           fd.long_term_borrowings()),
            (4,  "Short-Term Borrowings",          fd.short_term_borrowings() + fd.bank_od_in_bank_accounts()),
            (5,  "Trade Payables",                 fd.trade_payables()),
            (8,  "Fixed Assets (Net Block)",       fd.net_fixed_assets()),
            (11, "Inventories",                    fd.closing_stock()),
            (12, "Trade Receivables",              fd.trade_receivables()),
            (13, "Cash & Cash Equivalents",        fd.cash_and_bank()),
        ]
        for i, (num, title, amount) in enumerate(NOTE_INDEX, start=3):
            sheet_name = self._note_sheet_name(num)
            ws[f"A{i}"].value = f"Note {num}"
            ws[f"A{i}"].hyperlink = f"#'{sheet_name}'!A1"
            ws[f"A{i}"].font = Font(name="Calibri", size=9, color="1D4ED8", underline="single")
            ws[f"A{i}"].alignment = _align("center")
            ws[f"B{i}"].value = title
            ws[f"B{i}"].font = _font(size=9)
            ws[f"C{i}"].value = amount
            ws[f"C{i}"].number_format = INR
            ws[f"C{i}"].font = _font(size=9)
            ws[f"C{i}"].alignment = _align("right")
            if i % 2 == 0:
                for col in "ABC":
                    ws[f"{col}{i}"].fill = _fill(C_LIGHT_BLUE)

        ws.freeze_panes = "A3"

    # ── P&L Statement ─────────────────────────────────────────────────────────

    def _write_pnl(self) -> None:
        ws = self.wb.create_sheet("P&L Statement")
        fd = self.fd
        ws.column_dimensions["A"].width = 50
        ws.column_dimensions["B"].width = 8
        ws.column_dimensions["C"].width = 18

        r = 1

        def header(text, bg=C_HEADER_BG, fg=C_WHITE, sz=12, bold=True):
            nonlocal r
            ws.merge_cells(f"A{r}:C{r}")
            c = ws[f"A{r}"]
            c.value = text
            c.font = _font(bold=bold, size=sz, color=fg)
            c.fill = _fill(bg)
            c.alignment = _align("center")
            r += 1

        def row(label, amount=None, bold=False, total=False, indent=0, italic=False, bg=None):
            nonlocal r
            ws[f"A{r}"].value = "  " * indent + label
            ws[f"A{r}"].font = _font(bold=bold or total, size=10, name="Calibri")
            if italic:
                ws[f"A{r}"].font = Font(italic=True, size=10, name="Calibri")
            if amount is not None:
                ws[f"C{r}"].value = amount
                ws[f"C{r}"].number_format = INR
                ws[f"C{r}"].font = _font(bold=bold or total, size=10)
                ws[f"C{r}"].alignment = _align("right")
            if total or bg:
                col_bg = bg or C_TOTAL_BG
                ws[f"A{r}"].fill = _fill(col_bg)
                ws[f"C{r}"].fill = _fill(col_bg)
            r += 1
            return r - 1

        def spacer():
            nonlocal r; r += 1

        header(fd.company.upper(), sz=13)
        header("STATEMENT OF PROFIT & LOSS", sz=11)
        header(f"For the Year Ended {fd.period_label}", sz=10, bg=C_MID_BLUE)
        header("(Amount in ₹)", sz=9, bg=C_LIGHT_BLUE, fg=C_BLACK, bold=False)

        # Column labels
        ws[f"A{r}"].value = "Particulars"
        ws[f"B{r}"].value = "Note"
        ws[f"C{r}"].value = "Current Year"
        for col in "ABC":
            c = ws[f"{col}{r}"]
            c.font = _font(bold=True, size=9, color=C_WHITE)
            c.fill = _fill(C_MID_BLUE)
            c.alignment = _align("center")
            c.border = _border()
        r += 1

        # ── Revenue ──────────────────────────────────────────────────────────
        rev_r  = row("I.   Revenue from Operations",  fd.revenue_from_ops(), bold=True)
        oth_r  = row("II.  Other Income",              fd.other_income(),    bold=True)
        tot_r  = row("III. Total Revenue (I + II)",    None, bold=True, total=True)
        ws[f"C{tot_r}"].value  = f"=C{rev_r}+C{oth_r}"
        ws[f"C{tot_r}"].number_format = INR
        spacer()

        # ── Expenses ─────────────────────────────────────────────────────────
        row("IV.  Expenses", bold=True)
        pur_r  = row("     a)  Cost of Materials / Purchases", fd.purchases(), indent=1)
        chg_r  = row("     b)  Changes in Inventories",
                      fd.opening_stock() - fd.closing_stock(), indent=1,
                      italic=True)
        row("          (Opening Stock - Closing Stock)", italic=True, indent=2)
        emp_r  = row("     c)  Employee Benefits Expense",  fd.employee_costs(), indent=1)
        fin_r  = row("     d)  Finance Costs",              fd.finance_costs(), indent=1)
        dep_r  = row("     e)  Depreciation & Amortisation",fd.depreciation_from_fa(), indent=1)
        oth2_r = row("     f)  Other Expenses",             fd.other_indirect_expenses() + fd.direct_expenses(), indent=1)

        # Show breakdown of Other Expenses
        row("          Direct Expenses", fd.direct_expenses(), indent=3, italic=True)
        row("          Indirect Expenses (excl. Finance & Employee)", fd.other_indirect_expenses(), indent=3, italic=True)

        tot_exp_r = row("     Total Expenses", None, bold=True, total=True)
        ws[f"C{tot_exp_r}"].value = (
            f"=C{pur_r}+C{chg_r}+C{emp_r}+C{fin_r}+C{dep_r}+C{oth2_r}"
        )
        ws[f"C{tot_exp_r}"].number_format = INR
        spacer()

        # ── Profit lines ─────────────────────────────────────────────────────
        pbt_r = row("V.   Profit Before Tax (III − Expenses)", None, bold=True, total=True)
        ws[f"C{pbt_r}"].value = f"=C{tot_r}-C{tot_exp_r}"
        ws[f"C{pbt_r}"].number_format = INR

        tax_r = row("VI.  Tax Expense (Deferred Tax Provision)", fd.tax_expense())
        pat_r = row("VII. Profit After Tax", None, bold=True, bg=C_HEADER_BG)
        ws[f"C{pat_r}"].value = f"=C{pbt_r}-C{tax_r}"
        ws[f"C{pat_r}"].number_format = INR
        ws[f"C{pat_r}"].font = Font(bold=True, size=11, color=C_WHITE, name="Calibri")
        ws[f"A{pat_r}"].font = Font(bold=True, size=11, color=C_WHITE, name="Calibri")
        spacer(); spacer()

        # ── Key ratios ───────────────────────────────────────────────────────
        ws.merge_cells(f"A{r}:C{r}")
        ws[f"A{r}"].value = "KEY FINANCIAL RATIOS"
        ws[f"A{r}"].font = _font(bold=True, size=10, color=C_WHITE)
        ws[f"A{r}"].fill = _fill(C_MID_BLUE)
        r += 1

        rev = fd.revenue_from_ops()
        gp  = rev - fd.purchases() + (fd.closing_stock() - fd.opening_stock())
        ebitda = fd.profit_before_tax() + fd.finance_costs() + fd.depreciation_from_fa()
        ratios = [
            ("Gross Profit", gp, f"{gp/rev*100:.1f}% of Revenue" if rev else "N/A"),
            ("EBITDA", ebitda, f"{ebitda/rev*100:.1f}% of Revenue" if rev else "N/A"),
            ("PBT", fd.profit_before_tax(), f"{fd.profit_before_tax()/rev*100:.1f}% of Revenue" if rev else "N/A"),
            ("PAT", fd.profit_after_tax(), f"{fd.profit_after_tax()/rev*100:.1f}% of Revenue" if rev else "N/A"),
        ]
        for lbl, amt, ratio in ratios:
            ws[f"A{r}"].value = "  " + lbl
            ws[f"C{r}"].value = amt
            ws[f"C{r}"].number_format = INR
            ws[f"C{r}"].alignment = _align("right")
            # ratio in column B area as comment
            ws[f"B{r}"].value = ratio
            ws[f"B{r}"].font = _font(size=8, color="595959")
            ws[f"B{r}"].alignment = _align("center", wrap=True)
            ws.column_dimensions["B"].width = 20
            r += 1

        ws.freeze_panes = "A6"

    # ── Projected P&L ─────────────────────────────────────────────────────────

    def _write_proj_pnl(self, pe: "ProjectionEngine") -> None:
        ws = self.wb.create_sheet("Projected P&L")
        fd = self.fd
        ws.column_dimensions["A"].width = 42
        for col, lbl in [("B", "Base Year"), ("C", "Year 1"), ("D", "Year 2"), ("E", "Year 3")]:
            ws.column_dimensions[col].width = 16

        r = 1
        ws.merge_cells(f"A{r}:E{r}")
        ws[f"A{r}"].value = f"{fd.company.upper()} — PROJECTED PROFIT & LOSS (3 YEARS)"
        ws[f"A{r}"].font = _font(bold=True, size=12, color=C_WHITE)
        ws[f"A{r}"].fill = _fill(C_HEADER_BG)
        ws[f"A{r}"].alignment = _align("center")
        r += 1

        ws.merge_cells(f"A{r}:E{r}")
        ws[f"A{r}"].value = "Figures in ₹ | Projections are management estimates"
        ws[f"A{r}"].font = _font(size=9, color="595959")
        ws[f"A{r}"].alignment = _align("center")
        r += 1

        # Column headers
        years = [fd.period_label, "Year 1", "Year 2", "Year 3"]
        for i, col in enumerate(["A", "B", "C", "D", "E"]):
            ws[f"{col}{r}"].value = "Particulars" if i == 0 else years[i-1]
            ws[f"{col}{r}"].font = _font(bold=True, size=9, color=C_WHITE)
            ws[f"{col}{r}"].fill = _fill(C_MID_BLUE)
            ws[f"{col}{r}"].alignment = _align("center" if i > 0 else "left")
        r += 1

        pnl = pe.projected_pnl()  # list of 3 year dicts

        def proj_row(label: str, base_val: float, proj_vals: list[float],
                     bold=False, total=False, bg=None, fmt=INR):
            nonlocal r
            ws[f"A{r}"].value = label
            ws[f"A{r}"].font = _font(bold=bold or total)
            ws[f"B{r}"].value = base_val
            ws[f"B{r}"].number_format = fmt
            ws[f"B{r}"].alignment = _align("right")
            for i, (col, val) in enumerate(zip(["C", "D", "E"], proj_vals)):
                ws[f"{col}{r}"].value = val
                ws[f"{col}{r}"].number_format = fmt
                ws[f"{col}{r}"].font = _font(bold=bold or total)
                ws[f"{col}{r}"].alignment = _align("right")
            if total or bg:
                fill_c = bg or C_TOTAL_BG
                for col in ["A", "B", "C", "D", "E"]:
                    ws[f"{col}{r}"].fill = _fill(fill_c)
            r += 1

        base = fd  # actual year
        proj_row("Revenue from Operations",
                 base.revenue_from_ops(), [p["revenue"] for p in pnl], bold=True)
        proj_row("Other Income",
                 base.other_income(), [p["other_income"] for p in pnl])
        proj_row("TOTAL REVENUE",
                 base.revenue_from_ops() + base.other_income(),
                 [p["revenue"] + p["other_income"] for p in pnl],
                 total=True)
        r += 1
        proj_row("Cost of Materials / Purchases",
                 base.purchases(), [p["cogs"] for p in pnl])
        proj_row("Changes in Inventories",
                 base.opening_stock() - base.closing_stock(),
                 [p["stock_change"] for p in pnl])
        proj_row("Gross Profit",
                 base.revenue_from_ops() - base.purchases() + (base.closing_stock() - base.opening_stock()),
                 [p["gross_profit"] for p in pnl], bold=True, bg=C_LIGHT_BLUE)
        r += 1
        proj_row("Employee Benefits Expense",
                 base.employee_costs(), [p["employee"] for p in pnl])
        proj_row("Finance Costs",
                 base.finance_costs(), [p["finance"] for p in pnl])
        proj_row("Depreciation",
                 base.depreciation_from_fa(), [p["depreciation"] for p in pnl])
        proj_row("Other Expenses",
                 base.other_indirect_expenses() + base.direct_expenses(),
                 [p["other_expenses"] for p in pnl])
        proj_row("TOTAL EXPENSES",
                 base.total_expenses(),
                 [p["total_expenses"] for p in pnl], total=True)
        r += 1
        proj_row("EBITDA",
                 base.profit_before_tax() + base.finance_costs() + base.depreciation_from_fa(),
                 [p["ebitda"] for p in pnl], bold=True)
        proj_row("Profit Before Tax",
                 base.profit_before_tax(), [p["pbt"] for p in pnl], bold=True)
        proj_row("Tax Expense",
                 base.tax_expense(), [p["tax"] for p in pnl])
        proj_row("PROFIT AFTER TAX",
                 base.profit_after_tax(), [p["pat"] for p in pnl],
                 bold=True, bg=C_HEADER_BG)
        # Colour PAT row white text
        for col in ["A","B","C","D","E"]:
            ws[f"{col}{r-1}"].font = Font(bold=True, size=10, color=C_WHITE, name="Calibri")

        r += 2
        # Growth % rows
        ws.merge_cells(f"A{r}:E{r}")
        ws[f"A{r}"].value = "GROWTH & MARGIN METRICS"
        ws[f"A{r}"].font = _font(bold=True, size=9, color=C_WHITE)
        ws[f"A{r}"].fill = _fill(C_MID_BLUE)
        r += 1
        for metric, base_v, proj_vs in [
            ("Revenue Growth %", None, [p.get("revenue_growth") for p in pnl]),
            ("Gross Margin %", (base.revenue_from_ops() - base.purchases() + base.closing_stock() - base.opening_stock()) / base.revenue_from_ops() * 100 if base.revenue_from_ops() else 0,
             [p["gross_margin_pct"] for p in pnl]),
            ("EBITDA Margin %", (base.profit_before_tax() + base.finance_costs() + base.depreciation_from_fa()) / base.revenue_from_ops() * 100 if base.revenue_from_ops() else 0,
             [p["ebitda_margin"] for p in pnl]),
            ("PAT Margin %", base.profit_after_tax() / base.revenue_from_ops() * 100 if base.revenue_from_ops() else 0,
             [p["pat_margin"] for p in pnl]),
        ]:
            ws[f"A{r}"].value = "  " + metric
            ws[f"A{r}"].font = _font(size=9)
            if base_v is not None:
                ws[f"B{r}"].value = base_v; ws[f"B{r}"].number_format = "0.0%"; ws[f"B{r}"].value = base_v / 100
            for col, val in zip(["C","D","E"], proj_vs):
                if val is not None:
                    ws[f"{col}{r}"].value = val / 100
                    ws[f"{col}{r}"].number_format = "0.0%"
                    ws[f"{col}{r}"].alignment = _align("right")
            r += 1

        # Banker-view rows (DSCR, Interest Coverage) — only if requested
        if self.opts.banker_view:
            r += 1
            ws.merge_cells(f"A{r}:E{r}")
            ws[f"A{r}"].value = "BANKER METRICS"
            ws[f"A{r}"].font = _font(bold=True, size=9, color=C_WHITE)
            ws[f"A{r}"].fill = _fill(C_MID_BLUE)
            r += 1
            for metric, getter, fmt in [
                ("DSCR (Debt Service Coverage)",
                 lambda p: p.get("dscr"), '#,##0.00'),
                ("Interest Coverage Ratio",
                 lambda p: p.get("interest_coverage"), '#,##0.00'),
            ]:
                ws[f"A{r}"].value = "  " + metric
                ws[f"A{r}"].font = _font(size=9, bold=True)
                for col, p in zip(["C", "D", "E"], pnl):
                    v = getter(p)
                    if v is None:
                        ws[f"{col}{r}"].value = "—"
                    else:
                        ws[f"{col}{r}"].value = v
                        ws[f"{col}{r}"].number_format = fmt
                    ws[f"{col}{r}"].alignment = _align("right")
                r += 1

        ws.freeze_panes = "A5"

    # ── Projected Balance Sheet ───────────────────────────────────────────────

    def _write_proj_bs(self, pe: "ProjectionEngine") -> None:
        ws = self.wb.create_sheet("Projected Balance Sheet")
        fd = self.fd
        ws.column_dimensions["A"].width = 46
        for col in ["B","C","D","E"]:
            ws.column_dimensions[col].width = 16

        r = 1
        ws.merge_cells(f"A{r}:E{r}")
        ws[f"A{r}"].value = f"{fd.company.upper()} — PROJECTED BALANCE SHEET (3 YEARS)"
        ws[f"A{r}"].font = _font(bold=True, size=12, color=C_WHITE)
        ws[f"A{r}"].fill = _fill(C_HEADER_BG)
        ws[f"A{r}"].alignment = _align("center")
        r += 1

        ws.merge_cells(f"A{r}:E{r}")
        ws[f"A{r}"].value = "Provisional / Projected — For Discussion Purposes Only"
        ws[f"A{r}"].font = _font(size=9, color="FF0000")
        ws[f"A{r}"].alignment = _align("center")
        r += 1

        for i, (col, lbl) in enumerate(zip(["A","B","C","D","E"],
                                ["Particulars", "Base Year", "Year 1", "Year 2", "Year 3"])):
            ws[f"{col}{r}"].value = lbl
            ws[f"{col}{r}"].font = _font(bold=True, size=9, color=C_WHITE)
            ws[f"{col}{r}"].fill = _fill(C_MID_BLUE)
            ws[f"{col}{r}"].alignment = _align("left" if i == 0 else "center")
        r += 1

        bs = pe.projected_bs()

        def bs_row(label, base_val, proj_vals, bold=False, total=False, bg=None):
            nonlocal r
            ws[f"A{r}"].value = label
            ws[f"A{r}"].font = _font(bold=bold or total)
            ws[f"B{r}"].value = base_val; ws[f"B{r}"].number_format = INR
            ws[f"B{r}"].alignment = _align("right")
            for col, val in zip(["C","D","E"], proj_vals):
                ws[f"{col}{r}"].value = val
                ws[f"{col}{r}"].number_format = INR
                ws[f"{col}{r}"].font = _font(bold=bold or total)
                ws[f"{col}{r}"].alignment = _align("right")
            if total or bg:
                for col in ["A","B","C","D","E"]:
                    ws[f"{col}{r}"].fill = _fill(bg or C_TOTAL_BG)
            r += 1

        ws.merge_cells(f"A{r}:E{r}")
        ws[f"A{r}"].value = "EQUITY & LIABILITIES"; ws[f"A{r}"].font = _font(bold=True); ws[f"A{r}"].fill = _fill(C_SUBHD_BG); r += 1
        bs_row("  Share Capital",         fd.share_capital(),         [b["share_capital"] for b in bs])
        bs_row("  Reserves & Surplus",    fd.reserves_surplus(),      [b["reserves"] for b in bs], bold=True)
        bs_row("Total Equity",
               fd.total_equity(),         [b["total_equity"] for b in bs], total=True)
        r += 1
        ws.merge_cells(f"A{r}:E{r}")
        ws[f"A{r}"].value = "Non-Current Liabilities"; ws[f"A{r}"].font = _font(bold=True); ws[f"A{r}"].fill = _fill(C_SUBHD_BG); r += 1
        bs_row("  Long-Term Borrowings",  fd.long_term_borrowings(),  [b["lt_borrowings"] for b in bs])
        bs_row("  Deferred Tax Liability",fd.deferred_tax_liability(),[b["dtl"] for b in bs])
        r += 1
        ws.merge_cells(f"A{r}:E{r}")
        ws[f"A{r}"].value = "Current Liabilities"; ws[f"A{r}"].font = _font(bold=True); ws[f"A{r}"].fill = _fill(C_SUBHD_BG); r += 1
        bs_row("  Short-Term Borrowings", fd.short_term_borrowings(), [b["st_borrowings"] for b in bs])
        bs_row("  Trade Payables",        fd.trade_payables(),        [b["trade_payables"] for b in bs])
        bs_row("  Other Current Liab.",   fd.other_current_liabilities() + max(0, fd.duties_and_taxes_net()),
                                          [b["other_cl"] for b in bs])
        bs_row("  Short-Term Provisions", fd.short_term_provisions(), [b["provisions"] for b in bs])
        bs_row("TOTAL EQUITY & LIABILITIES",
               fd.total_equity_liabilities(), [b["total_el"] for b in bs],
               bold=True, bg=C_HEADER_BG)
        for col in ["A","B","C","D","E"]:
            ws[f"{col}{r-1}"].font = Font(bold=True, size=10, color=C_WHITE, name="Calibri")

        r += 2
        ws.merge_cells(f"A{r}:E{r}")
        ws[f"A{r}"].value = "ASSETS"; ws[f"A{r}"].font = _font(bold=True); ws[f"A{r}"].fill = _fill(C_SUBHD_BG); r += 1
        bs_row("  Fixed Assets (Net)",    fd.net_fixed_assets(),      [b["fixed_assets"] for b in bs])
        bs_row("  Long-Term Loans & Adv.",fd.long_term_loans_advances(),[b["lt_loans"] for b in bs])
        bs_row("Total Non-Current Assets",
               fd.total_noncurrent_assets(), [b["total_nca"] for b in bs], total=True)
        r += 1
        bs_row("  Inventories",           fd.closing_stock(),         [b["inventory"] for b in bs])
        bs_row("  Trade Receivables",     fd.trade_receivables(),     [b["debtors"] for b in bs])
        bs_row("  Cash & Cash Equiv.",    fd.cash_and_bank(),         [b["cash"] for b in bs])
        bs_row("  Other Current Assets",  fd.other_current_assets(),  [b["other_ca"] for b in bs])
        bs_row("Total Current Assets",
               fd.total_current_assets(), [b["total_ca"] for b in bs], total=True)
        r += 1
        bs_row("TOTAL ASSETS",
               fd.total_assets(),         [b["total_assets"] for b in bs],
               bold=True, bg=C_HEADER_BG)
        for col in ["A","B","C","D","E"]:
            ws[f"{col}{r-1}"].font = Font(bold=True, size=10, color=C_WHITE, name="Calibri")

        # Cash plug went negative in any year → flag the additional funding requirement
        shortfalls = [b.get("cash_shortfall", 0) for b in bs]
        if any(s > 0 for s in shortfalls):
            r += 2
            ws.merge_cells(f"A{r}:E{r}")
            ws[f"A{r}"].value = "⚠  FUNDING SHORTFALL — additional borrowings required to balance"
            ws[f"A{r}"].font = _font(bold=True, size=10, color="CC0000")
            ws[f"A{r}"].fill = _fill("FFE0E0")
            ws[f"A{r}"].alignment = _align("center")
            r += 1
            ws[f"A{r}"].value = "  Additional Short-Term Borrowing Needed"
            ws[f"A{r}"].font = _font(bold=True, size=10)
            for col, s in zip(["C", "D", "E"], shortfalls):
                ws[f"{col}{r}"].value = s
                ws[f"{col}{r}"].number_format = INR
                ws[f"{col}{r}"].alignment = _align("right")
                ws[f"{col}{r}"].font = _font(bold=True, color="CC0000")

        ws.freeze_panes = "A5"

    # ── Assumptions sheet ─────────────────────────────────────────────────────

    def _write_assumptions(self) -> None:
        ws = self.wb.create_sheet("Assumptions")
        p = self.proj
        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 18

        ws.merge_cells("A1:B1")
        ws["A1"].value = "PROJECTION ASSUMPTIONS"
        ws["A1"].font = _font(bold=True, size=12, color=C_WHITE)
        ws["A1"].fill = _fill(C_HEADER_BG)
        ws["A1"].alignment = _align("center")

        rows = [
            ("Revenue Growth – Year 1",         f"{p.rev_growth_y1:.1f}%"),
            ("Revenue Growth – Year 2",         f"{p.rev_growth_y2:.1f}%"),
            ("Revenue Growth – Year 3",         f"{p.rev_growth_y3:.1f}%"),
            ("Gross Margin %",                  f"{p.gross_margin_pct:.1f}%"),
            ("Operating Expenses Growth %",     f"{p.opex_growth_pct:.1f}%"),
            ("Inventory / COGS Days",           f"{p.inventory_days:.0f} days"),
            ("Debtor Days",                     f"{p.debtor_days:.0f} days"),
            ("Creditor Days",                   f"{p.creditor_days:.0f} days"),
            ("Annual Loan Repayment (₹)",       f"₹ {p.loan_repayment_pa:,.0f}"),
            ("New Borrowings per Year (₹)",     f"₹ {p.new_borrowings_pa:,.0f}"),
            ("Interest Rate %",                 f"{p.interest_rate_pct:.1f}%"),
            ("Capital Expenditure / Year (₹)",  f"₹ {p.capex_pa:,.0f}"),
            ("Depreciation Rate (WDV) %",       f"{p.depreciation_rate_pct:.1f}%"),
            ("Effective Tax Rate %",            f"{p.tax_rate_pct:.1f}%"),
            ("Other Income (fixed ₹ / year)",   f"₹ {p.other_income_pa:,.0f}"),
        ]
        for i, (lbl, val) in enumerate(rows, start=2):
            ws[f"A{i}"].value = lbl
            ws[f"A{i}"].font = _font(size=10)
            ws[f"B{i}"].value = val
            ws[f"B{i}"].font = _font(size=10, bold=True)
            ws[f"B{i}"].alignment = _align("right")
            if i % 2 == 0:
                ws[f"A{i}"].fill = _fill(C_LIGHT_BLUE)
                ws[f"B{i}"].fill = _fill(C_LIGHT_BLUE)

        ws[f"A{len(rows)+3}"].value = "Generated on"
        ws[f"B{len(rows)+3}"].value = date.today().strftime("%d %B %Y")
        ws[f"A{len(rows)+3}"].font = _font(size=9, color="595959")
        ws[f"B{len(rows)+3}"].font = _font(size=9, color="595959")

        if self.proj and self.proj.simple_mode:
            sm_r = len(rows) + 5
            ws[f"A{sm_r}"].value = "Mode: SIMPLE PROJECTION"
            ws[f"A{sm_r}"].font = _font(bold=True, size=10, color="2E7D32")
            ws[f"A{sm_r+1}"].value = ("Working capital days, OpEx growth, depreciation and interest "
                                     "rates were defaulted from base year / standard rates.")
            ws[f"A{sm_r+1}"].font = _font(size=9, color="595959")
            ws[f"A{sm_r+1}"].alignment = _align("left", wrap=True)
            ws.merge_cells(f"A{sm_r+1}:B{sm_r+1}")
            ws.row_dimensions[sm_r+1].height = 32

    # ── Ratios sheet ─────────────────────────────────────────────────────────

    def _write_ratios(self, pe: "ProjectionEngine | None") -> None:
        ws = self.wb.create_sheet("Ratios")
        fd = self.fd
        ws.column_dimensions["A"].width = 40
        for col in "BCDE":
            ws.column_dimensions[col].width = 16

        years_count = 1 + (len(pe.projected_pnl()) if pe else 0)
        last_col = chr(ord("B") + years_count - 1)

        ws.merge_cells(f"A1:{last_col}1")
        ws["A1"].value = f"{fd.company.upper()} — KEY FINANCIAL RATIOS"
        ws["A1"].font = _font(bold=True, size=12, color=C_WHITE)
        ws["A1"].fill = _fill(C_HEADER_BG)
        ws["A1"].alignment = _align("center")

        headers = ["Ratio", fd.period_label] + [f"Year {i+1}" for i in range(years_count - 1)]
        for i, h in enumerate(headers):
            col = chr(ord("A") + i)
            ws[f"{col}2"].value = h
            ws[f"{col}2"].font = _font(bold=True, size=9, color=C_WHITE)
            ws[f"{col}2"].fill = _fill(C_MID_BLUE)
            ws[f"{col}2"].alignment = _align("center" if i > 0 else "left")
            ws[f"{col}2"].border = _border()

        # Base-year metrics — protect against div-by-zero separately from the GP calc
        rev_real = fd.revenue_from_ops()
        rev = rev_real or 1
        purchases = fd.purchases() or 1
        gp_base = rev_real - fd.purchases() + (fd.closing_stock() - fd.opening_stock())
        ebitda_base = fd.profit_before_tax() + fd.finance_costs() + fd.depreciation_from_fa()
        curr_assets = fd.total_current_assets() or 0
        curr_liab = fd.total_current_liab() or 1
        debt = fd.long_term_borrowings() + fd.short_term_borrowings()
        equity = fd.total_equity() or 1
        interest = fd.finance_costs() or 0
        net_worth = equity

        def _row(label, base_val, proj_vals, fmt='#,##0.00', pct=False, group=False):
            r = ws.max_row + 1
            ws[f"A{r}"].value = label
            ws[f"A{r}"].font = _font(bold=group, size=10)
            if group:
                for col_i in range(years_count + 1):
                    col = chr(ord("A") + col_i)
                    ws[f"{col}{r}"].fill = _fill(C_SUBHD_BG)
                return
            vals = [base_val] + (proj_vals or [])
            for i, v in enumerate(vals):
                col = chr(ord("B") + i)
                if v is None:
                    ws[f"{col}{r}"].value = "—"
                else:
                    ws[f"{col}{r}"].value = v if not pct else v / 100
                    ws[f"{col}{r}"].number_format = "0.0%" if pct else fmt
                ws[f"{col}{r}"].alignment = _align("right")
                ws[f"{col}{r}"].border = _border()
            ws[f"A{r}"].border = _border()

        proj_pnl = pe.projected_pnl() if pe else []
        proj_bs  = pe.projected_bs() if pe else []

        _row("PROFITABILITY", None, None, group=True)
        _row("Gross Margin %", gp_base / rev * 100,
             [p["gross_margin_pct"] for p in proj_pnl], pct=True)
        _row("EBITDA Margin %", ebitda_base / rev * 100,
             [p["ebitda_margin"] for p in proj_pnl], pct=True)
        _row("PAT Margin %", fd.profit_after_tax() / rev * 100,
             [p["pat_margin"] for p in proj_pnl], pct=True)
        _row("Return on Equity %", fd.profit_after_tax() / equity * 100,
             [p["pat"] / max(1, b["total_equity"]) * 100 for p, b in zip(proj_pnl, proj_bs)], pct=True)
        _row("Return on Capital Employed %",
             (fd.profit_before_tax() + interest) / max(1, equity + debt) * 100,
             [(p["pbt"] + p["finance"]) / max(1, b["total_equity"] + b["lt_borrowings"] + b["st_borrowings"]) * 100
              for p, b in zip(proj_pnl, proj_bs)], pct=True)

        _row("LIQUIDITY", None, None, group=True)
        _row("Current Ratio", curr_assets / curr_liab,
             [b.get("current_ratio") for b in proj_bs])
        _row("Quick Ratio", (curr_assets - fd.closing_stock()) / curr_liab,
             [(b["debtors"] + b["cash"] + b["other_ca"]) /
              max(1, b["st_borrowings"] + b["trade_payables"] + b["other_cl"] + b["provisions"])
              for b in proj_bs])

        _row("LEVERAGE", None, None, group=True)
        _row("Debt-to-Equity", debt / equity,
             [(b["lt_borrowings"] + b["st_borrowings"]) / max(1, b["total_equity"]) for b in proj_bs])
        _row("Interest Coverage", (fd.profit_before_tax() + interest) / max(1, interest),
             [p.get("interest_coverage") for p in proj_pnl])
        _row("DSCR (Debt Service Coverage)", None,
             [p.get("dscr") for p in proj_pnl])

        _row("EFFICIENCY (DAYS)", None, None, group=True)
        _row("Inventory Days", fd.closing_stock() / purchases * 365,
             [b["inventory"] / max(1, p["cogs"]) * 365 for p, b in zip(proj_pnl, proj_bs)],
             fmt='#,##0')
        _row("Debtor Days", fd.trade_receivables() / rev * 365,
             [b["debtors"] / max(1, p["revenue"]) * 365 for p, b in zip(proj_pnl, proj_bs)],
             fmt='#,##0')
        _row("Creditor Days", fd.trade_payables() / purchases * 365,
             [b["trade_payables"] / max(1, p["cogs"]) * 365 for p, b in zip(proj_pnl, proj_bs)],
             fmt='#,##0')

        ws.freeze_panes = "B3"

    # ── Cash Flow Statement (Projected — indirect method) ────────────────────

    def _write_cash_flow(self, pe: "ProjectionEngine") -> None:
        ws = self.wb.create_sheet("Cash Flow")
        fd = self.fd
        ws.column_dimensions["A"].width = 46
        for col in "BCD":
            ws.column_dimensions[col].width = 18

        ws.merge_cells("A1:D1")
        ws["A1"].value = f"{fd.company.upper()} — PROJECTED CASH FLOW (INDIRECT METHOD)"
        ws["A1"].font = _font(bold=True, size=12, color=C_WHITE)
        ws["A1"].fill = _fill(C_HEADER_BG)
        ws["A1"].alignment = _align("center")

        for i, lbl in enumerate(["Particulars", "Year 1", "Year 2", "Year 3"]):
            col = chr(ord("A") + i)
            ws[f"{col}2"].value = lbl
            ws[f"{col}2"].font = _font(bold=True, size=9, color=C_WHITE)
            ws[f"{col}2"].fill = _fill(C_MID_BLUE)
            ws[f"{col}2"].alignment = _align("center" if i > 0 else "left")
            ws[f"{col}2"].border = _border()

        pnl = pe.projected_pnl()
        bs  = pe.projected_bs()

        prev_bs = {
            "debtors":        fd.trade_receivables(),
            "inventory":      fd.closing_stock(),
            "trade_payables": fd.trade_payables(),
            "other_ca":       fd.other_current_assets(),
            "other_cl":       fd.other_current_liabilities(),
        }

        def row(label, values, bold=False, total=False, italic=False):
            r = ws.max_row + 1
            ws[f"A{r}"].value = label
            ws[f"A{r}"].font = _font(bold=bold or total, size=10) if not italic else \
                              Font(italic=True, size=9, name="Calibri")
            for i, v in enumerate(values):
                col = chr(ord("B") + i)
                if v is None:
                    ws[f"{col}{r}"].value = "—"
                else:
                    ws[f"{col}{r}"].value = v
                    ws[f"{col}{r}"].number_format = INR
                ws[f"{col}{r}"].font = _font(bold=bold or total)
                ws[f"{col}{r}"].alignment = _align("right")
                ws[f"{col}{r}"].border = _border()
            ws[f"A{r}"].border = _border()
            if total:
                for col in "ABCD":
                    ws[f"{col}{r}"].fill = _fill(C_TOTAL_BG)
            return r

        ws[f"A3"].value = "Cash Flow from Operating Activities"
        ws[f"A3"].font = _font(bold=True, size=10)
        ws[f"A3"].fill = _fill(C_SUBHD_BG)
        for col in "BCD":
            ws[f"{col}3"].fill = _fill(C_SUBHD_BG)
            ws[f"{col}3"].border = _border()
        ws[f"A3"].border = _border()

        row("Profit Before Tax", [p["pbt"] for p in pnl])
        row("Add: Depreciation", [p["depreciation"] for p in pnl])
        row("Add: Finance Costs", [p["finance"] for p in pnl])
        # Working capital movements
        wc_rows = []
        for p, b in zip(pnl, bs):
            dr_change   = -(b["debtors"] - prev_bs["debtors"])
            inv_change  = -(b["inventory"] - prev_bs["inventory"])
            cred_change =  (b["trade_payables"] - prev_bs["trade_payables"])
            other_ca_c  = -(b["other_ca"] - prev_bs["other_ca"])
            other_cl_c  =  (b["other_cl"] - prev_bs["other_cl"])
            wc_rows.append((dr_change, inv_change, cred_change, other_ca_c, other_cl_c))
            prev_bs = {
                "debtors":        b["debtors"],
                "inventory":      b["inventory"],
                "trade_payables": b["trade_payables"],
                "other_ca":       b["other_ca"],
                "other_cl":       b["other_cl"],
            }
        row("Increase / (decrease) in Trade Payables", [w[2] for w in wc_rows])
        row("(Increase) / decrease in Trade Receivables", [w[0] for w in wc_rows])
        row("(Increase) / decrease in Inventories", [w[1] for w in wc_rows])
        row("(Increase) / decrease in Other Current Assets", [w[3] for w in wc_rows])
        row("Increase / (decrease) in Other Current Liabilities", [w[4] for w in wc_rows])
        row("Less: Taxes Paid", [-p["tax"] for p in pnl])
        row("Less: Finance Costs Paid", [-p["finance"] for p in pnl])
        cfo = []
        for i, p in enumerate(pnl):
            v = (p["pbt"] + p["depreciation"] + p["finance"]
                 + sum(wc_rows[i]) - p["tax"] - p["finance"])
            cfo.append(v)
        row("Net Cash from Operations (A)", cfo, total=True)

        # Investing
        r = ws.max_row + 1
        ws[f"A{r}"].value = "Cash Flow from Investing Activities"
        ws[f"A{r}"].font = _font(bold=True, size=10)
        for col in "ABCD":
            ws[f"{col}{r}"].fill = _fill(C_SUBHD_BG)
            ws[f"{col}{r}"].border = _border()
        capex = [-self.proj.capex_pa] * len(pnl)
        row("Capital Expenditure (CapEx)", capex)
        cfi = list(capex)
        row("Net Cash from Investing (B)", cfi, total=True)

        # Financing
        r = ws.max_row + 1
        ws[f"A{r}"].value = "Cash Flow from Financing Activities"
        ws[f"A{r}"].font = _font(bold=True, size=10)
        for col in "ABCD":
            ws[f"{col}{r}"].fill = _fill(C_SUBHD_BG)
            ws[f"{col}{r}"].border = _border()
        new_b = [self.proj.new_borrowings_pa] * len(pnl)
        rep_b = [-self.proj.loan_repayment_pa] * len(pnl)
        row("New Borrowings", new_b)
        row("Loan Repayments", rep_b)
        cff = [n + r_ for n, r_ in zip(new_b, rep_b)]
        row("Net Cash from Financing (C)", cff, total=True)

        # Net change
        net_change = [a + b_ + c for a, b_, c in zip(cfo, cfi, cff)]
        row("Net Change in Cash (A + B + C)", net_change, total=True, bold=True)
        # Opening cash and closing cash
        prev_cash = fd.cash_and_bank()
        opening = []
        closing = []
        for nc in net_change:
            opening.append(prev_cash)
            prev_cash = prev_cash + nc
            closing.append(prev_cash)
        row("Opening Cash & Equivalents", opening)
        row("Closing Cash & Equivalents", closing, total=True, bold=True)

        ws.freeze_panes = "B3"

    # ── Common-Size statements ───────────────────────────────────────────────

    def _write_common_size(self, pe: "ProjectionEngine | None") -> None:
        ws = self.wb.create_sheet("Common-Size")
        fd = self.fd
        ws.column_dimensions["A"].width = 44
        for col in "BCDEFG":
            ws.column_dimensions[col].width = 14

        ws.merge_cells("A1:G1")
        ws["A1"].value = f"{fd.company.upper()} — COMMON-SIZE STATEMENTS"
        ws["A1"].font = _font(bold=True, size=12, color=C_WHITE)
        ws["A1"].fill = _fill(C_HEADER_BG)
        ws["A1"].alignment = _align("center")

        # P&L common-size (% of revenue)
        r = 3
        ws[f"A{r}"].value = "P&L (as % of Revenue)"
        ws[f"A{r}"].font = _font(bold=True, size=10, color=C_WHITE)
        ws[f"A{r}"].fill = _fill(C_MID_BLUE)
        ws.merge_cells(f"A{r}:G{r}")
        r += 1
        proj_pnl = pe.projected_pnl() if pe else []
        years = [fd.period_label] + [f"Year {i+1}" for i in range(len(proj_pnl))]
        for i, y in enumerate(years):
            col = chr(ord("B") + i)
            ws[f"{col}{r}"].value = y
            ws[f"{col}{r}"].font = _font(bold=True, size=9, color=C_WHITE)
            ws[f"{col}{r}"].fill = _fill(C_MID_BLUE)
            ws[f"{col}{r}"].alignment = _align("center")
            ws[f"{col}{r}"].border = _border()
        ws[f"A{r}"].value = "Particulars"
        ws[f"A{r}"].font = _font(bold=True, size=9, color=C_WHITE)
        ws[f"A{r}"].fill = _fill(C_MID_BLUE)
        ws[f"A{r}"].border = _border()
        r += 1

        rev_base = fd.revenue_from_ops() or 1
        base_pnl_items = [
            ("Revenue from Operations", fd.revenue_from_ops()),
            ("Cost of Materials / Purchases", fd.purchases()),
            ("Employee Benefits", fd.employee_costs()),
            ("Finance Costs", fd.finance_costs()),
            ("Depreciation", fd.depreciation_from_fa()),
            ("Other Expenses", fd.other_indirect_expenses() + fd.direct_expenses()),
            ("Profit Before Tax", fd.profit_before_tax()),
            ("Profit After Tax", fd.profit_after_tax()),
        ]
        proj_keys = ["revenue", "cogs", "employee", "finance", "depreciation",
                     "other_expenses", "pbt", "pat"]
        for (lbl, base_v), key in zip(base_pnl_items, proj_keys):
            ws[f"A{r}"].value = lbl
            ws[f"A{r}"].border = _border()
            ws[f"B{r}"].value = base_v / rev_base
            ws[f"B{r}"].number_format = "0.0%"
            ws[f"B{r}"].border = _border()
            ws[f"B{r}"].alignment = _align("right")
            for i, p in enumerate(proj_pnl):
                col = chr(ord("C") + i)
                ws[f"{col}{r}"].value = p[key] / max(1, p["revenue"])
                ws[f"{col}{r}"].number_format = "0.0%"
                ws[f"{col}{r}"].border = _border()
                ws[f"{col}{r}"].alignment = _align("right")
            r += 1

        # BS common-size (% of total assets)
        r += 1
        ws[f"A{r}"].value = "Balance Sheet (as % of Total Assets)"
        ws[f"A{r}"].font = _font(bold=True, size=10, color=C_WHITE)
        ws[f"A{r}"].fill = _fill(C_MID_BLUE)
        ws.merge_cells(f"A{r}:G{r}")
        r += 1
        proj_bs = pe.projected_bs() if pe else []
        for i, y in enumerate(years):
            col = chr(ord("B") + i)
            ws[f"{col}{r}"].value = y
            ws[f"{col}{r}"].font = _font(bold=True, size=9, color=C_WHITE)
            ws[f"{col}{r}"].fill = _fill(C_MID_BLUE)
            ws[f"{col}{r}"].alignment = _align("center")
            ws[f"{col}{r}"].border = _border()
        ws[f"A{r}"].value = "Particulars"
        ws[f"A{r}"].font = _font(bold=True, size=9, color=C_WHITE)
        ws[f"A{r}"].fill = _fill(C_MID_BLUE)
        ws[f"A{r}"].border = _border()
        r += 1
        ta_base = fd.total_assets() or 1
        base_bs_items = [
            ("Shareholders' Funds", fd.total_equity()),
            ("Long-Term Borrowings", fd.long_term_borrowings()),
            ("Short-Term Borrowings + Trade Payables",
             fd.short_term_borrowings() + fd.trade_payables()),
            ("Fixed Assets (Net)", fd.net_fixed_assets()),
            ("Inventories", fd.closing_stock()),
            ("Trade Receivables", fd.trade_receivables()),
            ("Cash & Equivalents", fd.cash_and_bank()),
        ]
        proj_bs_keys = [
            lambda b: b["total_equity"],
            lambda b: b["lt_borrowings"],
            lambda b: b["st_borrowings"] + b["trade_payables"],
            lambda b: b["fixed_assets"],
            lambda b: b["inventory"],
            lambda b: b["debtors"],
            lambda b: b["cash"],
        ]
        for (lbl, base_v), getter in zip(base_bs_items, proj_bs_keys):
            ws[f"A{r}"].value = lbl
            ws[f"A{r}"].border = _border()
            ws[f"B{r}"].value = base_v / ta_base
            ws[f"B{r}"].number_format = "0.0%"
            ws[f"B{r}"].border = _border()
            ws[f"B{r}"].alignment = _align("right")
            for i, b in enumerate(proj_bs):
                col = chr(ord("C") + i)
                ws[f"{col}{r}"].value = getter(b) / max(1, b["total_assets"])
                ws[f"{col}{r}"].number_format = "0.0%"
                ws[f"{col}{r}"].border = _border()
                ws[f"{col}{r}"].alignment = _align("right")
            r += 1

        ws.freeze_panes = "B3"

    # ── Charts sheet ─────────────────────────────────────────────────────────

    def _write_charts(self, pe: "ProjectionEngine | None") -> None:
        ws = self.wb.create_sheet("Charts")
        fd = self.fd
        ws.column_dimensions["A"].width = 20
        for col in "BCDE":
            ws.column_dimensions[col].width = 16

        ws.merge_cells("A1:E1")
        ws["A1"].value = f"{fd.company.upper()} — KEY CHARTS"
        ws["A1"].font = _font(bold=True, size=12, color=C_WHITE)
        ws["A1"].fill = _fill(C_HEADER_BG)
        ws["A1"].alignment = _align("center")

        # Build a hidden data block for charts
        if pe:
            proj_pnl = pe.projected_pnl()
            years = [fd.period_label] + [f"Year {i+1}" for i in range(len(proj_pnl))]
            revenue = [fd.revenue_from_ops()] + [p["revenue"] for p in proj_pnl]
            pat = [fd.profit_after_tax()] + [p["pat"] for p in proj_pnl]
            ebitda = [fd.profit_before_tax() + fd.finance_costs() + fd.depreciation_from_fa()] \
                     + [p["ebitda"] for p in proj_pnl]
            pat_margin = [fd.profit_after_tax() / max(1, fd.revenue_from_ops()) * 100] \
                         + [p["pat_margin"] for p in proj_pnl]
        else:
            years = [fd.period_label]
            revenue = [fd.revenue_from_ops()]
            pat = [fd.profit_after_tax()]
            ebitda = [fd.profit_before_tax() + fd.finance_costs() + fd.depreciation_from_fa()]
            pat_margin = [fd.profit_after_tax() / max(1, fd.revenue_from_ops()) * 100]

        ws["A3"].value = "Year"
        for i, y in enumerate(years):
            ws.cell(row=3, column=2 + i, value=y)
        ws["A4"].value = "Revenue"
        for i, v in enumerate(revenue):
            ws.cell(row=4, column=2 + i, value=v).number_format = INR
        ws["A5"].value = "EBITDA"
        for i, v in enumerate(ebitda):
            ws.cell(row=5, column=2 + i, value=v).number_format = INR
        ws["A6"].value = "PAT"
        for i, v in enumerate(pat):
            ws.cell(row=6, column=2 + i, value=v).number_format = INR
        ws["A7"].value = "PAT Margin %"
        for i, v in enumerate(pat_margin):
            ws.cell(row=7, column=2 + i, value=v / 100).number_format = "0.0%"

        for row_i in range(3, 8):
            ws[f"A{row_i}"].font = _font(bold=(row_i == 3), size=9)
            for c in range(2, 2 + len(years)):
                ws.cell(row=row_i, column=c).alignment = _align("right")

        # Revenue / EBITDA / PAT bar chart
        ncols = len(years)
        bar = BarChart()
        bar.type = "col"
        bar.style = 12
        bar.title = "Revenue, EBITDA & PAT"
        bar.y_axis.title = "₹"
        bar.x_axis.title = "Year"
        data = Reference(ws, min_col=1, min_row=4, max_row=6, max_col=1 + ncols)
        cats = Reference(ws, min_col=2, min_row=3, max_col=1 + ncols, max_row=3)
        bar.add_data(data, titles_from_data=True)
        bar.set_categories(cats)
        bar.height = 9
        bar.width = 18
        ws.add_chart(bar, "A10")

        # PAT Margin trend line
        line = LineChart()
        line.title = "PAT Margin %"
        line.y_axis.title = "%"
        line.style = 12
        ldata = Reference(ws, min_col=1, min_row=7, max_row=7, max_col=1 + ncols)
        line.add_data(ldata, titles_from_data=True)
        line.set_categories(cats)
        line.height = 9
        line.width = 18
        ws.add_chart(line, "A30")


# ─── Projection Engine ────────────────────────────────────────────────────────

@dataclass
class ProjectionInputs:
    rev_growth_y1:       float = 15.0   # %
    rev_growth_y2:       float = 15.0   # %
    rev_growth_y3:       float = 10.0   # %
    gross_margin_pct:    float = 30.0   # % of revenue
    opex_growth_pct:     float = 10.0   # % growth in operating expenses
    inventory_days:      float = 45.0   # days
    debtor_days:         float = 60.0   # days
    creditor_days:       float = 45.0   # days
    loan_repayment_pa:   float = 0.0    # ₹ per year
    new_borrowings_pa:   float = 0.0    # ₹ per year
    interest_rate_pct:   float = 14.0   # % on average borrowings
    capex_pa:            float = 0.0    # ₹ per year
    depreciation_rate_pct: float = 15.0 # % WDV
    tax_rate_pct:        float = 25.0   # %
    other_income_pa:     float = 0.0    # ₹ fixed per year
    simple_mode:         bool  = False  # only use rev_growth_y1/y2/y3, gross_margin_pct, tax_rate_pct


@dataclass
class OutputOptions:
    """Optional Excel sections — keep the simple-projection output uncluttered by default."""
    cash_flow:    bool = False
    ratios:       bool = True
    charts:       bool = True
    common_size:  bool = False
    banker_view:  bool = False  # DSCR, Interest Coverage, Current Ratio in projected sheets
    page_setup:   bool = True   # fit-to-width on every sheet


class ProjectionEngine:
    def __init__(self, fd: FinancialData, p: ProjectionInputs):
        self.fd = fd
        self.p  = p
        # Apply simple-mode defaults derived from the base year so the projection
        # is sensible even when the user only set growth/margin/tax.
        if p.simple_mode:
            base_rev = fd.revenue_from_ops() or 1
            # Derive working capital days from the base year for continuity
            p.inventory_days = round(fd.closing_stock() / max(1, fd.purchases()) * 365) if fd.purchases() else 45
            p.debtor_days    = round(fd.trade_receivables() / base_rev * 365)
            p.creditor_days  = round(fd.trade_payables() / max(1, fd.purchases()) * 365) if fd.purchases() else 45
            p.opex_growth_pct = max(5.0, min(p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3))
            p.depreciation_rate_pct = 15.0
            p.interest_rate_pct = 12.0
            # Keep loan_repayment_pa, new_borrowings_pa, capex_pa, other_income_pa at 0

    def projected_pnl(self) -> list[dict]:
        fd = self.fd
        p  = self.p

        results = []
        prev_rev = fd.revenue_from_ops()
        prev_lt_borr = fd.long_term_borrowings()
        prev_net_fa  = fd.net_fixed_assets()
        prev_inventory = fd.closing_stock()
        growth_rates = [p.rev_growth_y1, p.rev_growth_y2, p.rev_growth_y3]

        for yr_idx, g in enumerate(growth_rates):
            rev = prev_rev * (1 + g / 100)
            cogs = rev * (1 - p.gross_margin_pct / 100)

            # Inventory scales every year (was previously frozen after year 0)
            inventory_close = cogs / 365 * p.inventory_days if cogs else prev_inventory
            stock_change = -(inventory_close - prev_inventory)  # +ve when stock falls
            gross_profit = rev - cogs + stock_change

            opex_mult = (1 + p.opex_growth_pct / 100) ** (yr_idx + 1)
            employee  = fd.employee_costs() * opex_mult
            other_exp = (fd.other_indirect_expenses() + fd.direct_expenses()) * opex_mult

            # Borrowings: interest is on the AVERAGE of opening and closing balance
            lt_close = max(0, prev_lt_borr + p.new_borrowings_pa - p.loan_repayment_pa)
            avg_lt   = (prev_lt_borr + lt_close) / 2
            st_borr  = fd.short_term_borrowings()  # held flat unless extended later
            finance  = (avg_lt + st_borr) * p.interest_rate_pct / 100

            # Depreciation (WDV).  CapEx contributes a half-year of depreciation
            # in the year it's added (mid-year convention) and is included in the
            # closing net block exactly once.
            fa_open  = prev_net_fa
            depr     = (fa_open + p.capex_pa * 0.5) * p.depreciation_rate_pct / 100
            net_fa   = fa_open + p.capex_pa - depr

            total_exp = cogs + employee + finance + depr + other_exp
            ebitda    = rev - cogs - employee - other_exp + stock_change
            pbt       = rev - total_exp + p.other_income_pa
            tax       = max(0, pbt * p.tax_rate_pct / 100)
            pat       = pbt - tax

            # Banker view: DSCR & Interest Coverage
            dscr = ((pat + depr + finance)
                    / (finance + p.loan_repayment_pa)) if (finance + p.loan_repayment_pa) > 0 else None
            interest_coverage = ((pat + finance + tax) / finance) if finance > 0 else None

            results.append({
                "revenue":          rev,
                "revenue_growth":   g,
                "other_income":     p.other_income_pa,
                "cogs":             cogs,
                "stock_change":     stock_change,
                "gross_profit":     gross_profit,
                "gross_margin_pct": p.gross_margin_pct,
                "employee":         employee,
                "finance":          finance,
                "depreciation":     depr,
                "other_expenses":   other_exp,
                "total_expenses":   total_exp,
                "ebitda":           ebitda,
                "ebitda_margin":    ebitda / rev * 100 if rev else 0,
                "pbt":              pbt,
                "tax":              tax,
                "pat":              pat,
                "pat_margin":       pat / rev * 100 if rev else 0,
                "net_fa":           net_fa,
                "lt_borr_close":    lt_close,
                "inventory_close":  inventory_close,
                "dscr":             dscr,
                "interest_coverage": interest_coverage,
            })
            prev_rev       = rev
            prev_lt_borr   = lt_close
            prev_net_fa    = net_fa
            prev_inventory = inventory_close

        return results

    def projected_bs(self) -> list[dict]:
        fd = self.fd
        p  = self.p
        pnl = self.projected_pnl()

        bs_list = []
        prev_reserves = fd.reserves_surplus()

        for yr_idx, yr in enumerate(pnl):
            rev  = yr["revenue"]
            cogs = yr["cogs"]
            pat  = yr["pat"]

            # Equity
            reserves    = prev_reserves + pat
            share_cap   = fd.share_capital()
            total_equity = share_cap + reserves

            # Borrowings — use the closing balance computed by the P&L pass
            lt_borr = yr["lt_borr_close"]
            st_borr = fd.short_term_borrowings()

            # Working capital
            inventory = yr["inventory_close"]
            debtors   = rev  / 365 * p.debtor_days
            creditors = cogs / 365 * p.creditor_days

            # Fixed assets
            fa = yr["net_fa"]

            # Other items (kept at base year)
            lt_loans = fd.long_term_loans_advances()
            other_ca = fd.other_current_assets()
            provisions = fd.short_term_provisions()
            dtl       = fd.deferred_tax_liability()

            # Total liabilities
            other_cl = fd.other_current_liabilities()
            total_el = total_equity + lt_borr + dtl + st_borr + creditors + other_cl + provisions

            # Cash (plug — balances the sheet)
            total_nca  = fa + lt_loans
            total_excl_cash = total_nca + inventory + debtors + other_ca
            cash = total_el - total_excl_cash

            # If cash plug goes negative the business is short of funding;
            # surface this so the UI / Validation sheet can flag it.
            cash_shortfall = max(0, -cash)
            cash_for_bs   = max(0, cash)
            if cash_shortfall > 0:
                # Treat shortfall as additional short-term borrowing needed to balance
                st_borr += cash_shortfall
                total_el += cash_shortfall

            total_ca = inventory + debtors + cash_for_bs + other_ca
            total_assets = total_nca + total_ca

            bs_list.append({
                "share_capital":  share_cap,
                "reserves":       reserves,
                "total_equity":   total_equity,
                "lt_borrowings":  lt_borr,
                "dtl":            dtl,
                "st_borrowings":  st_borr,
                "trade_payables": creditors,
                "other_cl":       other_cl,
                "provisions":     provisions,
                "total_el":       total_el,
                "fixed_assets":   fa,
                "lt_loans":       lt_loans,
                "total_nca":      total_nca,
                "inventory":      inventory,
                "debtors":        debtors,
                "cash":           cash_for_bs,
                "other_ca":       other_ca,
                "total_ca":       total_ca,
                "total_assets":   total_assets,
                "cash_shortfall": cash_shortfall,
                # Banker ratios
                "current_ratio":  ((inventory + debtors + cash_for_bs + other_ca) /
                                   (st_borr + creditors + other_cl + provisions))
                                  if (st_borr + creditors + other_cl + provisions) > 0 else None,
            })
            prev_reserves = reserves

        return bs_list


# ─── Tkinter UI ───────────────────────────────────────────────────────────────

class App:
    SETTINGS_PATH = Path.home() / ".tallyfin_settings.json"

    def __init__(self) -> None:
        if not _HAS_CTK:
            messagebox.showerror(
                "Missing dependency",
                "customtkinter is not installed.\n\n"
                "Run: pip install customtkinter\n"
                "or re-run the launcher script (run_windows.bat / run_mac.command) "
                "which installs it automatically.")
            raise SystemExit(1)

        # Theme: light by default; user can flip via toggle in title bar
        saved = self._load_settings()
        ctk.set_appearance_mode(saved.get("appearance_mode", "light"))
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        self.root.title("Tally Financial Statements Generator")
        self.root.geometry("900x820")
        self.root.minsize(820, 720)

        self.db_path     = tk.StringVar(value=saved.get("db_path", ""))
        self.out_dir     = tk.StringVar(value=saved.get("out_dir",
                                                       str(Path.home() / "Desktop")))
        self.op_stock    = tk.StringVar()
        self.cl_stock    = tk.StringVar()
        self.status_var  = tk.StringVar(value="Select a Tally SQLite file to begin.")
        self.fd: FinancialData | None = None
        self.reclassify_map: dict[str, str] = {}
        self._mapping_vars: dict[str, tk.StringVar] = {}

        # Projection mode (simple = only 3 inputs; detailed = all 15)
        self.proj_mode = tk.StringVar(value=saved.get("proj_mode", "simple"))

        # Projection inputs
        defaults = saved.get("proj_inputs", {})
        def _v(k, d): return tk.StringVar(value=str(defaults.get(k, d)))
        self.proj_vars: dict[str, tk.StringVar] = {
            "rev_growth_y1":       _v("rev_growth_y1", "15"),
            "rev_growth_y2":       _v("rev_growth_y2", "15"),
            "rev_growth_y3":       _v("rev_growth_y3", "10"),
            "gross_margin_pct":    _v("gross_margin_pct", "30"),
            "opex_growth_pct":     _v("opex_growth_pct", "10"),
            "inventory_days":      _v("inventory_days", "45"),
            "debtor_days":         _v("debtor_days", "60"),
            "creditor_days":       _v("creditor_days", "45"),
            "loan_repayment_pa":   _v("loan_repayment_pa", "0"),
            "new_borrowings_pa":   _v("new_borrowings_pa", "0"),
            "interest_rate_pct":   _v("interest_rate_pct", "14"),
            "capex_pa":            _v("capex_pa", "0"),
            "depreciation_rate_pct": _v("depreciation_rate_pct", "15"),
            "tax_rate_pct":        _v("tax_rate_pct", "25"),
            "other_income_pa":     _v("other_income_pa", "0"),
        }

        # Output options (which optional Excel sections to include)
        out_opts = saved.get("output_options", {})
        self.opt_vars: dict[str, tk.BooleanVar] = {
            "ratios":      tk.BooleanVar(value=out_opts.get("ratios", True)),
            "charts":      tk.BooleanVar(value=out_opts.get("charts", True)),
            "cash_flow":   tk.BooleanVar(value=out_opts.get("cash_flow", False)),
            "common_size": tk.BooleanVar(value=out_opts.get("common_size", False)),
            "banker_view": tk.BooleanVar(value=out_opts.get("banker_view", False)),
            "page_setup":  tk.BooleanVar(value=out_opts.get("page_setup", True)),
        }

        self._build()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _load_settings(self) -> dict:
        try:
            import json
            if self.SETTINGS_PATH.exists():
                return json.loads(self.SETTINGS_PATH.read_text())
        except Exception:
            pass
        return {}

    def _save_settings(self) -> None:
        try:
            import json
            data = {
                "db_path":  self.db_path.get(),
                "out_dir":  self.out_dir.get(),
                "proj_mode": self.proj_mode.get(),
                "proj_inputs": {k: v.get() for k, v in self.proj_vars.items()},
                "output_options": {k: v.get() for k, v in self.opt_vars.items()},
                "appearance_mode": ctk.get_appearance_mode().lower(),
            }
            self.SETTINGS_PATH.write_text(json.dumps(data, indent=2))
        except Exception:
            pass

    def _toggle_appearance(self) -> None:
        new_mode = "dark" if ctk.get_appearance_mode().lower() == "light" else "light"
        ctk.set_appearance_mode(new_mode)
        if hasattr(self, "_theme_btn"):
            self._theme_btn.configure(text="☀  Light" if new_mode == "dark" else "🌙  Dark")

    def _on_close(self) -> None:
        self._save_settings()
        self.root.destroy()

    def _build(self) -> None:
        # ── Title bar ────────────────────────────────────────────────────────
        title_frame = ctk.CTkFrame(self.root, fg_color="#1F3864", corner_radius=0, height=78)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)
        ctk.CTkLabel(title_frame, text="Tally Financial Statements Generator",
                     font=ctk.CTkFont(family="Calibri", size=18, weight="bold"),
                     text_color="white").pack(pady=(10, 0))
        ctk.CTkLabel(title_frame,
                     text="Schedule III Balance Sheet  •  P&L  •  3-Year Projections",
                     font=ctk.CTkFont(family="Calibri", size=11),
                     text_color="#BDD7EE").pack()

        # Dark/light toggle in top-right corner of the title bar
        is_dark = ctk.get_appearance_mode().lower() == "dark"
        self._theme_btn = ctk.CTkButton(title_frame, width=88, height=26,
            text="☀  Light" if is_dark else "🌙  Dark",
            command=self._toggle_appearance,
            fg_color="#2F5496", hover_color="#4472C4",
            text_color="white", corner_radius=6,
            font=ctk.CTkFont(size=10, weight="bold"))
        self._theme_btn.place(relx=1.0, x=-14, y=14, anchor="ne")

        main = ctk.CTkFrame(self.root, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=16, pady=(12, 12))

        # ── File selection ────────────────────────────────────────────────────
        file_frame = ctk.CTkFrame(main, corner_radius=10)
        file_frame.pack(fill="x", pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(file_frame, text="  Database File",
                     font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 2))

        ctk.CTkLabel(file_frame, text="Tally SQLite:").grid(
            row=1, column=0, sticky="w", padx=(14, 8), pady=4)
        ctk.CTkEntry(file_frame, textvariable=self.db_path, width=500).grid(
            row=1, column=1, sticky="ew", pady=4)
        ctk.CTkButton(file_frame, text="Browse…", width=88,
                      command=self._pick_db).grid(row=1, column=2, padx=(8, 14), pady=4)

        ctk.CTkLabel(file_frame, text="Output Folder:").grid(
            row=2, column=0, sticky="w", padx=(14, 8), pady=(4, 10))
        ctk.CTkEntry(file_frame, textvariable=self.out_dir, width=500).grid(
            row=2, column=1, sticky="ew", pady=(4, 10))
        ctk.CTkButton(file_frame, text="Browse…", width=88,
                      command=self._pick_outdir).grid(
            row=2, column=2, padx=(8, 14), pady=(4, 10))

        # ── Tabs ─────────────────────────────────────────────────────────────
        self.nb = ctk.CTkTabview(main, corner_radius=10,
                                  segmented_button_selected_color="#1F3864",
                                  segmented_button_selected_hover_color="#2F5496")
        self.nb.pack(fill="both", expand=True, pady=(0, 10))
        self.nb.add("Actual Statements")
        self.nb.add("3-Year Projections")
        self.nb.add("Validation")
        self.nb.add("Group Mapping")

        self._build_actual_tab(self.nb.tab("Actual Statements"))
        self._build_proj_tab(self.nb.tab("3-Year Projections"))
        self._build_validation_tab(self.nb.tab("Validation"))
        self._build_mapping_tab(self.nb.tab("Group Mapping"))

        # ── Status + action ───────────────────────────────────────────────────
        bottom = ctk.CTkFrame(main, fg_color="transparent")
        bottom.pack(fill="x")
        self.status_lbl = ctk.CTkLabel(bottom, textvariable=self.status_var,
                                        text_color="#2F5496",
                                        font=ctk.CTkFont(size=10),
                                        anchor="w")
        self.status_lbl.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(bottom, text="✕  Clear", width=92,
                      command=self._clear,
                      fg_color="#888", hover_color="#666").pack(side="right", padx=(8, 0))

    def _build_actual_tab(self, tab) -> None:
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        # Stock adjustments
        stock_frame = ctk.CTkFrame(tab, corner_radius=10)
        stock_frame.grid(row=0, column=0, sticky="ew", padx=4, pady=(8, 12))
        stock_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(stock_frame, text="  Stock Values",
                     font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 2))

        help_text = ("Tally's 'Closing Stock' ledger holds accounting-entry stock values. "
                     "Enter corrected values below if they differ from the audit trial balance. "
                     "Leave blank to use Tally's own figures.")
        ctk.CTkLabel(stock_frame, text=help_text, text_color="gray50",
                     font=ctk.CTkFont(size=10), wraplength=720,
                     justify="left", anchor="w").grid(
            row=1, column=0, columnspan=3, sticky="w", padx=14, pady=(0, 8))

        ctk.CTkLabel(stock_frame, text="Opening Stock (₹):").grid(
            row=2, column=0, sticky="w", padx=(14, 10), pady=4)
        self.op_entry = ctk.CTkEntry(stock_frame, textvariable=self.op_stock, width=200,
                                      placeholder_text="from Tally if blank")
        self.op_entry.grid(row=2, column=1, sticky="w", pady=4)

        ctk.CTkLabel(stock_frame, text="Closing Stock (₹):").grid(
            row=3, column=0, sticky="w", padx=(14, 10), pady=4)
        self.cl_entry = ctk.CTkEntry(stock_frame, textvariable=self.cl_stock, width=200,
                                      placeholder_text="from Tally if blank")
        self.cl_entry.grid(row=3, column=1, sticky="w", pady=4)

        ctk.CTkButton(stock_frame, text="Apply & Preview Numbers", width=200,
                      command=self._preview_actual).grid(
            row=4, column=0, columnspan=2, padx=14, pady=(10, 12), sticky="w")

        # Preview panel (monospace, scrollable)
        self.preview_text = ctk.CTkTextbox(tab,
            font=ctk.CTkFont(family="Courier", size=10),
            corner_radius=10, wrap="none")
        self.preview_text.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 8))
        self.preview_text.configure(state="disabled")

        btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        btn_frame.grid(row=2, column=0, sticky="ew", padx=4)
        ctk.CTkButton(btn_frame, text="Generate Actual Statements (Excel)",
                      command=self._gen_actual, height=36,
                      font=ctk.CTkFont(size=12, weight="bold"),
                      fg_color="#1F3864", hover_color="#2F5496").pack(side="right")

    def _build_proj_tab(self, tab) -> None:
        # CTkScrollableFrame replaces the canvas+scrollbar+inner workaround
        scroll = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=4, pady=8)
        scroll.columnconfigure(1, weight=1)
        scroll.columnconfigure(3, weight=1)

        # ── Mode selector (Simple / Detailed) ────────────────────────────────
        mode_frame = ctk.CTkFrame(scroll, corner_radius=10)
        mode_frame.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 12))
        ctk.CTkLabel(mode_frame, text="  Projection Mode",
                     font=ctk.CTkFont(size=11, weight="bold")).pack(
            anchor="w", padx=12, pady=(8, 2))
        ctk.CTkRadioButton(mode_frame,
            text="Simple — only 3 inputs (revenue growth · gross margin · tax)",
            variable=self.proj_mode, value="simple",
            command=self._refresh_proj_fields).pack(anchor="w", padx=14, pady=2)
        ctk.CTkRadioButton(mode_frame,
            text="Detailed — all 15 banker-grade inputs",
            variable=self.proj_mode, value="detailed",
            command=self._refresh_proj_fields).pack(anchor="w", padx=14, pady=(2, 10))

        simple_keys = {"rev_growth_y1", "rev_growth_y2", "rev_growth_y3",
                       "gross_margin_pct", "tax_rate_pct"}

        fields = [
            ("Revenue Growth – Year 1 (%)",            "rev_growth_y1"),
            ("Revenue Growth – Year 2 (%)",            "rev_growth_y2"),
            ("Revenue Growth – Year 3 (%)",            "rev_growth_y3"),
            ("Gross Margin % (Revenue − COGS)",        "gross_margin_pct"),
            ("Effective Tax Rate (%)",                  "tax_rate_pct"),
            ("Operating Expense Growth (% per year)",  "opex_growth_pct"),
            ("Inventory Days (Stock / COGS × 365)",    "inventory_days"),
            ("Debtor Days (Receivables / Rev × 365)",  "debtor_days"),
            ("Creditor Days (Payables / COGS × 365)",  "creditor_days"),
            ("Loan Repayment per Year (₹)",            "loan_repayment_pa"),
            ("New Borrowings per Year (₹)",            "new_borrowings_pa"),
            ("Interest Rate on Borrowings (%)",        "interest_rate_pct"),
            ("Capital Expenditure per Year (₹)",       "capex_pa"),
            ("Depreciation Rate – WDV (%)",            "depreciation_rate_pct"),
            ("Other / Fixed Income per Year (₹)",      "other_income_pa"),
        ]

        hints = {
            "gross_margin_pct":  "Gross Profit as % of revenue (Revenue − Purchases ± Stock)",
            "debtor_days":       "Days outstanding for trade receivables",
            "creditor_days":     "Days outstanding for trade payables",
            "inventory_days":    "Days of stock held (based on COGS)",
            "interest_rate_pct": "Average annualised interest on total borrowings",
            "depreciation_rate_pct": "Written-down-value rate; 15% is typical for plant & equipment",
            "tax_rate_pct":      "Effective corporate tax rate (base rate 22% + surcharge ≈ 25.17%)",
        }

        # Group fields in a CTk frame for visual cohesion
        fields_frame = ctk.CTkFrame(scroll, corner_radius=10)
        fields_frame.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(0, 12))
        fields_frame.columnconfigure(1, weight=1)
        fields_frame.columnconfigure(3, weight=1)
        ctk.CTkLabel(fields_frame, text="  Projection Inputs",
                     font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(8, 4))

        self._proj_field_widgets: list = []
        for i, (label, key) in enumerate(fields):
            row_i = 1 + i // 2
            col_base = (i % 2) * 2
            lbl = ctk.CTkLabel(fields_frame, text=label,
                                font=ctk.CTkFont(size=10), anchor="w")
            lbl.grid(row=row_i, column=col_base, sticky="w",
                     padx=(14, 8) if col_base == 0 else (8, 8), pady=3)
            e = ctk.CTkEntry(fields_frame, textvariable=self.proj_vars[key], width=120)
            e.grid(row=row_i, column=col_base + 1, sticky="w",
                   padx=(0, 14) if col_base == 2 else (0, 8), pady=3)
            if key in hints:
                e.bind("<FocusIn>", lambda ev, h=hints[key]: self._set_status(h))
            self._proj_field_widgets.append((key, lbl, e))
        # Bottom padding inside fields_frame
        ctk.CTkLabel(fields_frame, text="").grid(row=99, column=0, pady=(2, 6))
        self._simple_keys = simple_keys

        # ── Output Options ───────────────────────────────────────────────────
        opt_frame = ctk.CTkFrame(scroll, corner_radius=10)
        opt_frame.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(0, 12))
        opt_frame.columnconfigure(0, weight=1)
        opt_frame.columnconfigure(1, weight=2)
        ctk.CTkLabel(opt_frame, text="  Optional Excel Sheets",
                     font=ctk.CTkFont(size=11, weight="bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(8, 4))
        opts_layout = [
            ("ratios",      "📊  Ratios sheet",
             "Liquidity, leverage, profitability, efficiency days"),
            ("charts",      "📈  Charts (Revenue / EBITDA / PAT / Margin)",
             "Bar + line charts on a dedicated sheet"),
            ("cash_flow",   "💰  Cash Flow Statement (Projected)",
             "Indirect method, 3-year"),
            ("common_size", "📋  Common-Size statements",
             "P&L as % of revenue · BS as % of total assets"),
            ("banker_view", "🏦  Banker view (DSCR + Interest Coverage)",
             "Add bank-style coverage rows to the Projected P&L"),
            ("page_setup",  "🖨  Print-ready page setup (fit to width)",
             "A4 page setup with header/footer for every sheet"),
        ]
        for i, (key, lbl, hint) in enumerate(opts_layout):
            ctk.CTkCheckBox(opt_frame, text=lbl, variable=self.opt_vars[key],
                            font=ctk.CTkFont(size=10)).grid(
                row=i + 1, column=0, sticky="w", padx=(14, 8), pady=3)
            ctk.CTkLabel(opt_frame, text=hint, text_color="gray50",
                          font=ctk.CTkFont(size=9), anchor="w").grid(
                row=i + 1, column=1, sticky="w", pady=3, padx=(0, 14))
        ctk.CTkLabel(opt_frame, text="").grid(row=99, column=0, pady=(0, 6))

        # ── Guidance ─────────────────────────────────────────────────────────
        hint_frame = ctk.CTkFrame(scroll, fg_color="#FFF6D5", corner_radius=10)
        hint_frame.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(0, 12))
        ctk.CTkLabel(hint_frame,
            text=(
                "GUIDANCE\n"
                "• Simple mode: only revenue growth, gross margin and tax rate matter. "
                "Working-capital days, OpEx growth and interest rate are derived from base year.\n"
                "• Detailed mode: full banker model — tune every lever.\n"
                "• Tick optional sheets only when you need them — keeps the file lean.\n"
                "• If projected cash plug goes negative, the BS shows the funding shortfall in red."),
            text_color="#7B5800", font=ctk.CTkFont(size=10),
            wraplength=820, justify="left", anchor="w").pack(
            anchor="w", padx=14, pady=10, fill="x")

        # ── Action buttons ───────────────────────────────────────────────────
        btn_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_frame.grid(row=4, column=0, columnspan=4, sticky="e", pady=(4, 12))
        ctk.CTkButton(btn_frame, text="Projections Only", width=160,
                      command=self._gen_proj_only,
                      fg_color="#888", hover_color="#666").pack(side="right", padx=(0, 0))
        ctk.CTkButton(btn_frame, text="Generate Projected + Actual (All Sheets)",
                      width=320, height=36, command=self._gen_all,
                      font=ctk.CTkFont(size=12, weight="bold"),
                      fg_color="#1F3864", hover_color="#2F5496").pack(
            side="right", padx=(0, 8))

        # Apply initial enable/disable state for fields
        self._refresh_proj_fields()

    def _refresh_proj_fields(self) -> None:
        """Grey out detailed-mode fields when Simple mode is selected."""
        is_simple = self.proj_mode.get() == "simple"
        for key, lbl, entry in getattr(self, "_proj_field_widgets", []):
            if is_simple and key not in self._simple_keys:
                entry.configure(state="disabled")
                lbl.configure(text_color="gray60")
            else:
                entry.configure(state="normal")
                lbl.configure(text_color=("gray10", "gray90"))

    def _build_validation_tab(self, tab) -> None:
        tab.rowconfigure(1, weight=1)
        tab.columnconfigure(0, weight=1)

        top = ctk.CTkFrame(tab, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=4, pady=(8, 6))
        self.val_summary = ctk.CTkLabel(top,
            text="Load a database file to see validation results.",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#2F5496", anchor="w")
        self.val_summary.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(top, text="Re-run Validation", width=140,
                      command=self._run_validation).pack(side="right")

        # ttk.Treeview has no CTk equivalent — keep it but theme it to match
        tree_container = ctk.CTkFrame(tab, corner_radius=10)
        tree_container.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 8))
        tree_container.rowconfigure(0, weight=1)
        tree_container.columnconfigure(0, weight=1)

        is_dark = ctk.get_appearance_mode().lower() == "dark"
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Validation.Treeview",
                        background="#2B2B2B" if is_dark else "#FFFFFF",
                        foreground="#E0E0E0" if is_dark else "#222222",
                        fieldbackground="#2B2B2B" if is_dark else "#FFFFFF",
                        rowheight=24, borderwidth=0,
                        font=("Calibri", 10))
        style.configure("Validation.Treeview.Heading",
                        background="#1F3864", foreground="white",
                        relief="flat", font=("Calibri", 10, "bold"))
        style.map("Validation.Treeview.Heading",
                  background=[("active", "#2F5496")])

        cols = ("Severity", "Category", "Message", "Detail")
        tree = ttk.Treeview(tree_container, columns=cols, show="headings",
                             height=20, style="Validation.Treeview")
        tree.heading("Severity",  text="Severity",  anchor="center")
        tree.heading("Category",  text="Category",  anchor="w")
        tree.heading("Message",   text="Message",   anchor="w")
        tree.heading("Detail",    text="Detail",    anchor="w")
        tree.column("Severity",  width=90,  stretch=False, anchor="center")
        tree.column("Category",  width=140, stretch=False)
        tree.column("Message",   width=340)
        tree.column("Detail",    width=380)

        tree.tag_configure("ERROR",   background="#5C2222" if is_dark else "#FFCCCC",
                                       foreground="#FF8080" if is_dark else "#CC0000")
        tree.tag_configure("WARNING", background="#4A3F1A" if is_dark else "#FFF2CC",
                                       foreground="#FFD060" if is_dark else "#7B5800")
        tree.tag_configure("INFO",    background="#1F3A22" if is_dark else "#E8F5E9",
                                       foreground="#90D098" if is_dark else "#1B5E20")

        vsb = ctk.CTkScrollbar(tree_container, command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)
        vsb.grid(row=0, column=1, sticky="ns", padx=(2, 10), pady=10)

        self.val_tree = tree

    def _build_mapping_tab(self, tab) -> None:
        tab.rowconfigure(1, weight=1)
        tab.columnconfigure(0, weight=1)

        ctk.CTkLabel(tab,
            text="Assign unrecognised primary groups to standard Schedule III heads. "
                 "Groups already in the built-in map are shown but cannot be changed here.",
            font=ctk.CTkFont(size=10),
            text_color="gray50", wraplength=820, anchor="w", justify="left").grid(
            row=0, column=0, sticky="ew", padx=4, pady=(8, 6))

        self._mapping_inner = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        self._mapping_inner.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 8))

        btn_row = ctk.CTkFrame(tab, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="ew", padx=4, pady=(0, 8))
        ctk.CTkButton(btn_row, text="Apply Mapping & Reload Preview",
                      width=260, height=32,
                      command=self._apply_mapping,
                      fg_color="#1F3864", hover_color="#2F5496").pack(side="right")
        self._mapping_vars = {}

    def _populate_mapping_tab(self, fd: "FinancialData") -> None:
        """Fill the mapping tab from the loaded FinancialData."""
        from collections import defaultdict
        pg_balance: dict[str, float] = defaultdict(float)
        pg_parents: dict[str, set[str]] = defaultdict(set)
        for l in fd.ledgers:
            pg_balance[l.primary_group] += l.closing
            pg_parents[l.primary_group].add(l.parent)

        # Clear old rows
        for child in self._mapping_inner.winfo_children():
            child.destroy()
        self._mapping_vars.clear()

        ALL_STANDARD = sorted(set(BS_MAP.keys()) | set(PNL_MAP.keys()))
        DROPDOWN_VALUES = ["(Use Auto-Infer)"] + ALL_STANDARD

        self._mapping_inner.columnconfigure(0, weight=2)
        self._mapping_inner.columnconfigure(1, weight=1)
        self._mapping_inner.columnconfigure(2, weight=2)
        self._mapping_inner.columnconfigure(3, weight=2)

        # Column headers
        header_bg = "#1F3864"
        for c, lbl in enumerate(["Primary Group", "Net Balance (₹)",
                                  "Auto-Inferred Head", "Your Override"]):
            hdr = ctk.CTkLabel(self._mapping_inner, text=lbl,
                                font=ctk.CTkFont(size=10, weight="bold"),
                                fg_color=header_bg, text_color="white",
                                corner_radius=4, height=28,
                                anchor="center" if c in (0, 3) else "e" if c == 1 else "w")
            hdr.grid(row=0, column=c, sticky="ew", padx=2, pady=(0, 4))

        sorted_pgs = sorted(pg_balance.items(), key=lambda x: abs(x[1]), reverse=True)
        for row_i, (pg, bal) in enumerate(sorted_pgs, start=1):
            in_map = pg in BS_MAP or pg in PNL_MAP
            auto = pg if in_map else (infer_standard_group(pg, pg_parents[pg]) or "(no match)")
            current_override = self.reclassify_map.get(pg, "(Use Auto-Infer)")

            stripe_dark  = ("gray20" if row_i % 2 == 0 else "gray15")
            stripe_light = ("gray95" if row_i % 2 == 0 else "white")
            stripe = (stripe_light, stripe_dark)

            ctk.CTkLabel(self._mapping_inner, text=pg,
                          font=ctk.CTkFont(size=10, weight="bold" if not in_map else "normal"),
                          text_color="#CC0000" if not in_map else ("gray10", "gray90"),
                          fg_color=stripe, anchor="w",
                          corner_radius=4, height=28).grid(
                row=row_i, column=0, sticky="ew", padx=2, pady=1)

            ctk.CTkLabel(self._mapping_inner, text=f"{bal:,.0f}",
                          font=ctk.CTkFont(size=10),
                          fg_color=stripe, anchor="e",
                          corner_radius=4, height=28).grid(
                row=row_i, column=1, sticky="ew", padx=2, pady=1)

            ctk.CTkLabel(self._mapping_inner, text=auto,
                          font=ctk.CTkFont(size=10, slant="italic"),
                          text_color="gray50", fg_color=stripe, anchor="w",
                          corner_radius=4, height=28).grid(
                row=row_i, column=2, sticky="ew", padx=2, pady=1)

            if in_map:
                ctk.CTkLabel(self._mapping_inner, text="(built-in)",
                              font=ctk.CTkFont(size=10),
                              text_color="gray55", fg_color=stripe,
                              corner_radius=4, height=28).grid(
                    row=row_i, column=3, sticky="ew", padx=2, pady=1)
            else:
                var = tk.StringVar(value=current_override)
                self._mapping_vars[pg] = var
                ctk.CTkOptionMenu(self._mapping_inner, variable=var,
                                   values=DROPDOWN_VALUES, width=280,
                                   font=ctk.CTkFont(size=10)).grid(
                    row=row_i, column=3, sticky="ew", padx=2, pady=1)

    def _apply_mapping(self) -> None:
        """Save combobox selections to reclassify_map and reload preview."""
        new_map = {}
        for pg, var in self._mapping_vars.items():
            val = var.get()
            if val and val != "(Use Auto-Infer)":
                new_map[pg] = val
        self.reclassify_map = new_map
        if self.db_path.get():
            self._preview_actual()
            self._run_validation()

    # ── Actions ───────────────────────────────────────────────────────────────

    def _set_status(self, msg: str) -> None:
        self.status_var.set(msg)

    def _pick_db(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Tally SQLite File",
            filetypes=[("SQLite Database", "*.sqlite *.db"), ("All Files", "*.*")]
        )
        if path:
            self.db_path.set(path)
            self._load_db()

    def _pick_outdir(self) -> None:
        d = filedialog.askdirectory(title="Select Output Folder")
        if d:
            self.out_dir.set(d)

    def _load_db(self) -> None:
        path = self.db_path.get()
        if not path or not os.path.isfile(path):
            return
        try:
            self.fd = load_from_sqlite(
                path,
                opening_stock_override=self._parse_stock(self.op_stock.get()),
                closing_stock_override=self._parse_stock(self.cl_stock.get()),
            )
            self._set_status(
                f"Loaded: {self.fd.company}  |  {self.fd.period_from} → {self.fd.period_to}"
            )
            self._preview_actual()
            self._run_validation()
            self._populate_mapping_tab(self.fd)
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    def _run_validation(self) -> None:
        fd = self.fd
        if fd is None:
            return
        vr = validate_financial_data(fd)
        # Populate treeview
        for item in self.val_tree.get_children():
            self.val_tree.delete(item)
        for chk in vr.checks:
            self.val_tree.insert("", "end",
                values=(chk.severity, chk.category, chk.message, chk.detail),
                tags=(chk.severity,))
        # Update summary label
        summary = vr.summary()
        color = "#CC0000" if vr.has_errors else ("#7B5800" if vr.warnings else "#1B5E20")
        self.val_summary.configure(text=summary, text_color=color)

    def _parse_stock(self, val: str) -> float | None:
        val = val.strip().replace(",", "")
        try:
            return float(val) if val else None
        except ValueError:
            return None

    def _get_fd(self) -> FinancialData | None:
        """Return (possibly refreshed) FinancialData."""
        path = self.db_path.get()
        if not path or not os.path.isfile(path):
            messagebox.showwarning("No File", "Please select a Tally SQLite file first.")
            return None
        try:
            fd = load_from_sqlite(
                path,
                opening_stock_override=self._parse_stock(self.op_stock.get()),
                closing_stock_override=self._parse_stock(self.cl_stock.get()),
                reclassify_map=self.reclassify_map,
            )
            self.fd = fd
            return fd
        except Exception as e:
            messagebox.showerror("Load Error", str(e))
            return None

    def _preview_actual(self) -> None:
        fd = self._get_fd()
        if fd is None:
            return
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")

        def line(label, amount, indent=0):
            prefix = "  " * indent
            self.preview_text.insert("end", f"{prefix}{label:<45} {amount:>16,.2f}\n")

        def divider():
            self.preview_text.insert("end", "─" * 64 + "\n")

        self.preview_text.insert("end",
            f"  {fd.company}\n"
            f"  Balance Sheet Preview — as at {fd.period_label}\n"
        )
        divider()
        self.preview_text.insert("end", "  EQUITY & LIABILITIES\n")
        line("Share Capital", fd.share_capital(), 1)
        line("Reserves & Surplus", fd.reserves_surplus(), 1)
        line("Long-Term Borrowings", fd.long_term_borrowings(), 1)
        line("Deferred Tax Liability", fd.deferred_tax_liability(), 1)
        line("Short-Term Borrowings (OD/CC)", fd.short_term_borrowings(), 1)
        line("Trade Payables", fd.trade_payables(), 1)
        line("Duties & Taxes (Net payable)", max(0, fd.duties_and_taxes_net()), 1)
        line("Other Current Liabilities", fd.other_current_liabilities(), 1)
        line("Short-Term Provisions", fd.short_term_provisions(), 1)
        divider()
        line("TOTAL EQUITY & LIABILITIES", fd.total_equity_liabilities(), 0)
        divider()
        self.preview_text.insert("end", "\n  ASSETS\n")
        line("Fixed Assets (Net)", fd.net_fixed_assets(), 1)
        line("Long-Term Loans & Advances", fd.long_term_loans_advances(), 1)
        line("Closing Stock (Inventories)", fd.closing_stock(), 1)
        line("Trade Receivables", fd.trade_receivables(), 1)
        line("Cash & Cash Equivalents", fd.cash_and_bank(), 1)
        line("Duties & Taxes Refund (GST/TDS)", fd.duties_and_taxes_asset(), 1)
        line("Other Current Assets", fd.other_current_assets(), 1)
        divider()
        line("TOTAL ASSETS", fd.total_assets(), 0)
        diff = fd.total_assets() - fd.total_equity_liabilities()
        divider()
        self.preview_text.insert("end", f"\n  Difference (Assets − E&L): {diff:,.2f}")
        if abs(diff) > 10:
            self.preview_text.insert("end", "  ← REVIEW REQUIRED\n")
        else:
            self.preview_text.insert("end", "  ← OK\n")

        self.preview_text.insert("end",
            f"\n  P&L SUMMARY\n"
            f"  Revenue from Operations: {fd.revenue_from_ops():>20,.2f}\n"
            f"  Cost of Materials:       {fd.purchases():>20,.2f}\n"
            f"  Gross Profit:            {fd.revenue_from_ops()-fd.purchases()+(fd.closing_stock()-fd.opening_stock()):>20,.2f}\n"
            f"  Finance Costs:           {fd.finance_costs():>20,.2f}\n"
            f"  Profit Before Tax:       {fd.profit_before_tax():>20,.2f}\n"
            f"  Profit After Tax:        {fd.profit_after_tax():>20,.2f}\n"
        )
        self.preview_text.configure(state="disabled")

    def _get_proj_inputs(self) -> ProjectionInputs | None:
        try:
            p = ProjectionInputs(**{
                k: float(v.get().replace(",",""))
                for k, v in self.proj_vars.items()
            })
            p.simple_mode = (self.proj_mode.get() == "simple")
            return p
        except ValueError as e:
            messagebox.showerror("Invalid Input",
                                 f"Please enter valid numbers in the projection fields.\n{e}")
            return None

    def _get_output_options(self) -> "OutputOptions":
        return OutputOptions(**{k: v.get() for k, v in self.opt_vars.items()})

    def _output_path(self, suffix="") -> str:
        fd = self.fd
        company_short = (fd.company.split()[0] if fd else "Company").replace(",","").replace(".","")
        year = fd.period_to[:4] if fd else "2026"
        fname = f"{company_short}_FinStatements_{year}{suffix}.xlsx"
        return str(Path(self.out_dir.get()) / fname)

    def _gen_actual(self) -> None:
        fd = self._get_fd()
        if fd is None:
            return
        out = self._output_path("_Actual")
        try:
            ExcelWriter(fd, proj=None, opts=self._get_output_options()).save(out)
            self._save_settings()
            self._set_status(f"Saved: {out}")
            if messagebox.askyesno("Done", f"Excel saved:\n{out}\n\nOpen it now?"):
                os.startfile(out) if os.name == "nt" else os.system(f'open "{out}"')
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _gen_proj_only(self) -> None:
        fd = self._get_fd()
        if fd is None:
            return
        proj = self._get_proj_inputs()
        if proj is None:
            return
        out = self._output_path("_Projected")
        try:
            ExcelWriter(fd, proj, opts=self._get_output_options()).save(out)
            self._save_settings()
            self._set_status(f"Saved: {out}")
            if messagebox.askyesno("Done", f"Excel saved:\n{out}\n\nOpen it now?"):
                os.startfile(out) if os.name == "nt" else os.system(f'open "{out}"')
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _gen_all(self) -> None:
        fd = self._get_fd()
        if fd is None:
            return
        proj = self._get_proj_inputs()
        if proj is None:
            return
        out = self._output_path("_Full")
        try:
            ExcelWriter(fd, proj, opts=self._get_output_options()).save(out)
            self._save_settings()
            self._set_status(f"Saved: {out}")
            if messagebox.askyesno("Done", f"Excel saved:\n{out}\n\nOpen it now?"):
                os.startfile(out) if os.name == "nt" else os.system(f'open "{out}"')
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _clear(self) -> None:
        self.db_path.set("")
        self.op_stock.set("")
        self.cl_stock.set("")
        self.fd = None
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")
        self.preview_text.configure(state="disabled")
        self._set_status("Cleared. Select a Tally SQLite file to begin.")


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    App()
