#!/usr/bin/env python3
"""
Tally Financial Statements Generator — PySide6 desktop UI.

A modern Qt-based front-end for the financial_statements engine.  Launches
the GUI when run directly.  All accounting / Excel logic lives in
financial_statements.py; this file is purely the view layer.

Run:  python app.py
Requires:  PySide6, openpyxl
"""
from __future__ import annotations

import os
import sys
import json
import traceback
from pathlib import Path
from dataclasses import asdict

from PySide6.QtCore import Qt, QSize, Signal, QThread, QTimer
from PySide6.QtGui import QAction, QColor, QFont, QIcon, QPalette
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QLineEdit,
    QFileDialog, QMessageBox, QTreeWidget, QTreeWidgetItem, QHBoxLayout,
    QVBoxLayout, QGridLayout, QFrame, QGroupBox, QCheckBox, QScrollArea,
    QSpacerItem, QSizePolicy, QInputDialog, QDialog, QDialogButtonBox,
    QHeaderView, QStatusBar, QStyle, QToolButton, QStackedWidget,
    QRadioButton, QButtonGroup,
)

import financial_statements as fs

# ─────────────────────────── Theme ─────────────────────────────

C_BG          = "#0F172A"   # slate-900
C_SURFACE     = "#1E293B"   # slate-800
C_SURFACE_2   = "#334155"   # slate-700
C_BORDER      = "#475569"   # slate-600
C_TEXT        = "#F1F5F9"   # slate-100
C_TEXT_DIM    = "#94A3B8"   # slate-400
C_ACCENT      = "#3B82F6"   # blue-500
C_ACCENT_DARK = "#1D4ED8"
C_SUCCESS     = "#10B981"   # emerald-500
C_WARN        = "#F59E0B"   # amber-500
C_ERROR       = "#EF4444"   # red-500
C_TILE_BG     = "#1E293B"

QSS = f"""
* {{
    font-family: 'Segoe UI', 'Inter', 'Calibri', sans-serif;
    color: {C_TEXT};
}}
QMainWindow, QDialog {{
    background-color: {C_BG};
}}
QScrollArea, QScrollArea > QWidget > QWidget {{
    background: transparent;
    border: none;
}}
QGroupBox {{
    background: {C_SURFACE};
    border: 1px solid {C_BORDER};
    border-radius: 10px;
    margin-top: 14px;
    padding: 12px 12px 12px 12px;
    font-weight: 600;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: {C_TEXT_DIM};
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 1px;
}}
QLabel {{ background: transparent; }}
QLabel[role="title"] {{
    color: {C_TEXT};
    font-size: 22px;
    font-weight: 700;
}}
QLabel[role="subtitle"] {{
    color: {C_TEXT_DIM};
    font-size: 11px;
}}
QLabel[role="metric-label"] {{
    color: {C_TEXT_DIM};
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 0.6px;
}}
QLabel[role="metric-value"] {{
    color: {C_TEXT};
    font-size: 20px;
    font-weight: 700;
}}
QLabel[role="company"] {{
    color: {C_TEXT};
    font-size: 18px;
    font-weight: 700;
}}
QLabel[role="period"] {{
    color: {C_TEXT_DIM};
    font-size: 10px;
}}
QLabel[role="badge-ok"] {{
    background: rgba(16, 185, 129, 0.15);
    color: {C_SUCCESS};
    padding: 6px 12px;
    border-radius: 12px;
    font-weight: 600;
    font-size: 11px;
}}
QLabel[role="badge-warn"] {{
    background: rgba(245, 158, 11, 0.15);
    color: {C_WARN};
    padding: 6px 12px;
    border-radius: 12px;
    font-weight: 600;
    font-size: 11px;
}}
QLabel[role="badge-err"] {{
    background: rgba(239, 68, 68, 0.15);
    color: {C_ERROR};
    padding: 6px 12px;
    border-radius: 12px;
    font-weight: 600;
    font-size: 11px;
}}
QLineEdit {{
    background: {C_SURFACE_2};
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    padding: 6px 8px;
    color: {C_TEXT};
    selection-background-color: {C_ACCENT};
}}
QLineEdit:focus {{
    border-color: {C_ACCENT};
}}
QPushButton {{
    background: {C_SURFACE_2};
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    padding: 7px 14px;
    color: {C_TEXT};
    font-weight: 600;
}}
QPushButton:hover {{
    background: #3F526B;
}}
QPushButton:pressed {{
    background: #2A3B52;
}}
QPushButton:disabled {{
    color: {C_TEXT_DIM};
    background: {C_SURFACE};
}}
QPushButton[role="primary"] {{
    background: {C_ACCENT};
    border-color: {C_ACCENT};
    color: white;
    font-size: 13px;
    padding: 10px 18px;
}}
QPushButton[role="primary"]:hover {{
    background: {C_ACCENT_DARK};
}}
QPushButton[role="secondary"] {{
    background: transparent;
    border: 1px solid {C_ACCENT};
    color: {C_ACCENT};
    padding: 10px 18px;
}}
QPushButton[role="secondary"]:hover {{
    background: rgba(59, 130, 246, 0.12);
}}
QPushButton[role="ghost"] {{
    background: transparent;
    border: 1px solid {C_BORDER};
    padding: 4px 10px;
    font-size: 11px;
    font-weight: 500;
}}
QToolButton {{
    background: transparent;
    color: {C_TEXT_DIM};
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    padding: 6px 10px;
    font-weight: 600;
}}
QToolButton:hover {{ background: {C_SURFACE_2}; color: {C_TEXT}; }}
QTreeWidget {{
    background: {C_SURFACE_2};
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    color: {C_TEXT};
    alternate-background-color: #2E3F57;
    selection-background-color: {C_ACCENT};
    selection-color: white;
    outline: none;
}}
QHeaderView::section {{
    background: {C_SURFACE};
    color: {C_TEXT_DIM};
    padding: 6px 8px;
    border: none;
    border-right: 1px solid {C_BORDER};
    font-weight: 600;
    text-transform: uppercase;
    font-size: 10px;
    letter-spacing: 0.6px;
}}
QCheckBox {{ color: {C_TEXT}; spacing: 8px; padding: 2px 0; }}
QCheckBox::indicator {{
    width: 16px; height: 16px;
    border: 1px solid {C_BORDER};
    border-radius: 3px;
    background: {C_SURFACE_2};
}}
QCheckBox::indicator:checked {{
    background: {C_ACCENT};
    border-color: {C_ACCENT};
    image: none;
}}
QStatusBar {{
    background: {C_SURFACE};
    color: {C_TEXT_DIM};
    border-top: 1px solid {C_BORDER};
    font-size: 11px;
}}
QFrame[role="card"] {{
    background: {C_SURFACE};
    border: 1px solid {C_BORDER};
    border-radius: 10px;
}}
QFrame[role="tile"] {{
    background: {C_TILE_BG};
    border: 1px solid {C_BORDER};
    border-radius: 8px;
}}
QFrame[role="separator"] {{
    background: {C_BORDER};
    max-height: 1px;
}}
"""

# ─────────────────────── Helpers ───────────────────────

SETTINGS_FILE = Path.home() / ".tallyfin_settings_pyside.json"


def load_settings() -> dict:
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(d: dict) -> None:
    try:
        SETTINGS_FILE.write_text(json.dumps(d, indent=2), encoding="utf-8")
    except Exception:
        pass


def fmt_indian(value: float) -> str:
    """Return ₹X.XX Cr or ₹X.X L depending on magnitude."""
    if value == 0:
        return "₹0"
    crore = value / 1_00_00_000
    if abs(crore) >= 1:
        return f"₹{crore:.2f} Cr"
    lakh = value / 1_00_000
    if abs(lakh) >= 0.1:
        return f"₹{lakh:.1f} L"
    return f"₹{value:,.0f}"


# ─────────────────────── Background load worker ───────────────────────

class LoadWorker(QThread):
    """Off-thread ZIP/SQLite load so the UI stays responsive on large files."""
    succeeded = Signal(object, str)   # (FinancialData, path)
    failed    = Signal(str, str)      # (path, error message)

    def __init__(self, path: str, branch_name: str):
        super().__init__()
        self.path = path
        self.branch_name = branch_name

    def run(self):
        cleanup = False
        sqlite_path = self.path
        try:
            if self.path.lower().endswith(".zip"):
                sqlite_path = fs.extract_sqlite_from_zip(self.path)
                cleanup = True
            fd = fs.load_from_sqlite(sqlite_path, branch_name=self.branch_name)
            self.succeeded.emit(fd, self.path)
        except Exception as e:
            self.failed.emit(self.path, f"{e}\n\n{traceback.format_exc()}")
        finally:
            if cleanup:
                try: os.unlink(sqlite_path)
                except Exception: pass


# ─────────────────────── Validation popup ───────────────────────

class ValidationDialog(QDialog):
    def __init__(self, vr: "fs.ValidationResult", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Data Validation")
        self.resize(920, 520)
        self.setStyleSheet(QSS)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        summary = QLabel(vr.summary())
        summary.setStyleSheet(
            f"color: {C_ERROR if vr.has_errors else (C_WARN if vr.warnings else C_SUCCESS)};"
            "font-size: 14px; font-weight: 700;"
        )
        layout.addWidget(summary)

        tree = QTreeWidget()
        tree.setColumnCount(4)
        tree.setHeaderLabels(["Severity", "Category", "Message", "Detail"])
        tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        tree.header().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        tree.header().setSectionResizeMode(2, QHeaderView.Stretch)
        tree.header().setSectionResizeMode(3, QHeaderView.Stretch)
        tree.setAlternatingRowColors(True)
        tree.setRootIsDecorated(False)

        severity_colors = {
            fs.ERROR:   QColor(239, 68, 68, 60),
            fs.WARNING: QColor(245, 158, 11, 60),
            fs.INFO:    QColor(16, 185, 129, 60),
        }
        for chk in vr.checks:
            item = QTreeWidgetItem([chk.severity, chk.category, chk.message, chk.detail])
            bg = severity_colors.get(chk.severity)
            if bg:
                for col in range(4):
                    item.setBackground(col, bg)
            tree.addTopLevelItem(item)
        layout.addWidget(tree, stretch=1)

        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(self.reject)
        bb.accepted.connect(self.accept)
        bb.button(QDialogButtonBox.Close).clicked.connect(self.accept)
        layout.addWidget(bb)


# ─────────────────────── Metric tile ───────────────────────

class MetricTile(QFrame):
    def __init__(self, label: str, parent=None):
        super().__init__(parent)
        self.setProperty("role", "tile")
        self.setMinimumWidth(140)
        v = QVBoxLayout(self)
        v.setContentsMargins(14, 10, 14, 12)
        v.setSpacing(2)
        self._lbl = QLabel(label.upper())
        self._lbl.setProperty("role", "metric-label")
        self._val = QLabel("—")
        self._val.setProperty("role", "metric-value")
        v.addWidget(self._lbl)
        v.addWidget(self._val)

    def set_value(self, value: float, color: str | None = None) -> None:
        self._val.setText(fmt_indian(value))
        if color:
            self._val.setStyleSheet(f"color: {color};")
        else:
            self._val.setStyleSheet("")


# ─────────────────────── Main window ───────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tally Financial Statements Generator")
        self.setMinimumSize(1080, 760)
        self.resize(1180, 820)
        self.setStyleSheet(QSS)

        self._branches: list[dict] = []     # [{path, name, fd}]
        self._fd: fs.FinancialData | None = None
        self._vr: fs.ValidationResult | None = None
        self._loader: LoadWorker | None = None
        self._proj_open = False

        self._settings = load_settings()
        self._out_dir = self._settings.get("out_dir", str(Path.home() / "Desktop"))

        self._build_ui()
        self._refresh_branch_tree()

    # ── Layout ──────────────────────────────────────────────────────

    def _build_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        self.setCentralWidget(scroll)

        root = QWidget()
        scroll.setWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(22, 18, 22, 18)
        outer.setSpacing(14)

        outer.addLayout(self._build_header())
        outer.addWidget(self._build_branches_card())
        self._build_summary_card()
        outer.addWidget(self._summary_card)
        outer.addWidget(self._build_stock_card())
        outer.addWidget(self._build_options_card())
        outer.addLayout(self._build_actions())
        outer.addWidget(self._build_projection_toggle())
        outer.addWidget(self._build_projection_card())
        outer.addStretch(1)

        self._status = QStatusBar()
        self.setStatusBar(self._status)
        self._set_status("Add one or more Tally ZIP exports to begin.")

    def _build_header(self) -> QHBoxLayout:
        h = QHBoxLayout()
        h.setContentsMargins(0, 0, 0, 0)
        v = QVBoxLayout()
        v.setSpacing(2)
        t = QLabel("Tally Financial Statements")
        t.setProperty("role", "title")
        s = QLabel("Schedule III BS  •  P&L  •  Notes  •  3-Year Projections  •  Multi-branch consolidation")
        s.setProperty("role", "subtitle")
        v.addWidget(t)
        v.addWidget(s)
        h.addLayout(v)
        h.addStretch(1)

        out_lbl = QLabel("Output:")
        out_lbl.setProperty("role", "subtitle")
        self._out_edit = QLineEdit(self._out_dir)
        self._out_edit.setMinimumWidth(280)
        self._out_edit.editingFinished.connect(self._on_out_changed)
        out_btn = QPushButton("Browse")
        out_btn.setProperty("role", "ghost")
        out_btn.clicked.connect(self._pick_out_dir)
        h.addWidget(out_lbl)
        h.addWidget(self._out_edit)
        h.addWidget(out_btn)
        return h

    def _build_branches_card(self) -> QGroupBox:
        gb = QGroupBox("SOURCE  ·  Tally ZIP / SQLite (one per branch)")

        v = QVBoxLayout(gb)
        v.setContentsMargins(12, 16, 12, 12)
        v.setSpacing(8)

        self._tree = QTreeWidget()
        self._tree.setColumnCount(4)
        self._tree.setHeaderLabels(["Branch", "Company", "Period", "File"])
        self._tree.header().setSectionResizeMode(0, QHeaderView.Interactive)
        self._tree.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self._tree.header().setSectionResizeMode(2, QHeaderView.Interactive)
        self._tree.header().setSectionResizeMode(3, QHeaderView.Interactive)
        self._tree.setColumnWidth(0, 160)
        self._tree.setColumnWidth(2, 180)
        self._tree.setColumnWidth(3, 260)
        self._tree.setRootIsDecorated(False)
        self._tree.setAlternatingRowColors(True)
        self._tree.setMinimumHeight(140)
        self._tree.itemDoubleClicked.connect(lambda *_: self._rename_branch())
        v.addWidget(self._tree)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        for label, slot, role in [
            ("+ Add Branch…",  self._add_branch,    None),
            ("Rename",         self._rename_branch, "ghost"),
            ("Remove",         self._remove_branch, "ghost"),
            ("Clear All",      self._clear,         "ghost"),
        ]:
            b = QPushButton(label)
            if role:
                b.setProperty("role", role)
            b.clicked.connect(slot)
            btn_row.addWidget(b)
        btn_row.addStretch(1)
        hint = QLabel("Tip: double-click a branch to rename")
        hint.setProperty("role", "subtitle")
        btn_row.addWidget(hint)
        v.addLayout(btn_row)
        return gb

    def _build_summary_card(self):
        card = QFrame()
        card.setProperty("role", "card")
        card.setVisible(False)
        self._summary_card = card

        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(10)

        top = QHBoxLayout()
        top.setSpacing(8)
        left = QVBoxLayout()
        left.setSpacing(2)
        self._company_lbl = QLabel("—")
        self._company_lbl.setProperty("role", "company")
        self._period_lbl = QLabel("")
        self._period_lbl.setProperty("role", "period")
        left.addWidget(self._company_lbl)
        left.addWidget(self._period_lbl)
        top.addLayout(left)
        top.addStretch(1)
        self._badge_lbl = QLabel("—")
        self._badge_lbl.setProperty("role", "badge-ok")
        self._badge_lbl.setCursor(Qt.PointingHandCursor)
        self._badge_lbl.mousePressEvent = self._on_badge_clicked   # type: ignore
        top.addWidget(self._badge_lbl)
        layout.addLayout(top)

        sep = QFrame()
        sep.setProperty("role", "separator")
        sep.setFixedHeight(1)
        layout.addWidget(sep)

        tile_row = QHBoxLayout()
        tile_row.setSpacing(10)
        self._tile_rev = MetricTile("Revenue")
        self._tile_gp  = MetricTile("Gross Profit")
        self._tile_pbt = MetricTile("PBT")
        self._tile_pat = MetricTile("PAT")
        self._tile_assets = MetricTile("Total Assets")
        self._tile_diff = MetricTile("BS Diff")
        for t in (self._tile_rev, self._tile_gp, self._tile_pbt,
                  self._tile_pat, self._tile_assets, self._tile_diff):
            tile_row.addWidget(t)
        layout.addLayout(tile_row)

    def _build_stock_card(self) -> QGroupBox:
        gb = QGroupBox("STOCK OVERRIDE  ·  blank = use Tally figures")
        h = QHBoxLayout(gb)
        h.setContentsMargins(12, 16, 12, 12)
        h.setSpacing(12)
        h.addWidget(QLabel("Opening Stock (₹):"))
        self._op_stock = QLineEdit()
        self._op_stock.setPlaceholderText("e.g. 1500000")
        self._op_stock.setMaximumWidth(180)
        h.addWidget(self._op_stock)
        h.addSpacing(20)
        h.addWidget(QLabel("Closing Stock (₹):"))
        self._cl_stock = QLineEdit()
        self._cl_stock.setPlaceholderText("e.g. 1800000")
        self._cl_stock.setMaximumWidth(180)
        h.addWidget(self._cl_stock)
        apply_btn = QPushButton("Apply")
        apply_btn.clicked.connect(self._apply_stock_override)
        h.addWidget(apply_btn)
        h.addStretch(1)
        return gb

    def _build_options_card(self) -> QGroupBox:
        gb = QGroupBox("OPTIONS")
        v = QVBoxLayout(gb)
        v.setContentsMargins(12, 16, 12, 12)
        v.setSpacing(6)
        self._cb_open = QCheckBox("Open the Excel file after generation")
        self._cb_open.setChecked(self._settings.get("auto_open", True))
        v.addWidget(self._cb_open)
        return gb

    def _build_actions(self) -> QHBoxLayout:
        h = QHBoxLayout()
        h.setSpacing(10)
        self._btn_actual = QPushButton("Generate Statements (Excel)")
        self._btn_actual.setProperty("role", "primary")
        self._btn_actual.clicked.connect(self._gen_actual)
        h.addWidget(self._btn_actual)

        self._btn_proj = QPushButton("Generate + 3-Year Projections")
        self._btn_proj.setProperty("role", "secondary")
        self._btn_proj.clicked.connect(self._gen_all)
        h.addWidget(self._btn_proj)

        self._btn_cma = QPushButton("Generate + CMA (Bank format)")
        self._btn_cma.setProperty("role", "secondary")
        self._btn_cma.setToolTip("Adds the 6-sheet RBI/Tandon CMA pack used "
                                  "by banks for working-capital & term-loan "
                                  "assessment. Requires projection inputs.")
        self._btn_cma.clicked.connect(self._gen_cma)
        h.addWidget(self._btn_cma)
        h.addStretch(1)
        return h

    def _build_projection_toggle(self) -> QWidget:
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        self._proj_toggle_btn = QToolButton()
        self._proj_toggle_btn.setText("▶  3-Year Projection Inputs  (click to expand)")
        self._proj_toggle_btn.setCheckable(True)
        self._proj_toggle_btn.clicked.connect(self._toggle_proj)
        h.addWidget(self._proj_toggle_btn)
        h.addStretch(1)
        return w

    def _build_projection_card(self) -> QGroupBox:
        gb = QGroupBox("PROJECTION INPUTS")
        gb.setVisible(False)
        self._proj_card = gb

        layout = QGridLayout(gb)
        layout.setContentsMargins(12, 18, 12, 12)
        layout.setHorizontalSpacing(14)
        layout.setVerticalSpacing(8)

        fields = [
            ("Rev Growth Y1 (%)",         "rev_growth_y1",         "15"),
            ("Rev Growth Y2 (%)",         "rev_growth_y2",         "15"),
            ("Rev Growth Y3 (%)",         "rev_growth_y3",         "10"),
            ("Gross Margin (%)",          "gross_margin_pct",      "30"),
            ("OpEx Growth (%/yr)",        "opex_growth_pct",       "10"),
            ("Inventory Days",            "inventory_days",        "45"),
            ("Debtor Days",               "debtor_days",           "60"),
            ("Creditor Days",             "creditor_days",         "45"),
            ("Loan Repayment p.a. (₹)",   "loan_repayment_pa",     "0"),
            ("New Borrowings p.a. (₹)",   "new_borrowings_pa",     "0"),
            ("Interest Rate (%)",         "interest_rate_pct",     "14"),
            ("CapEx p.a. (₹)",            "capex_pa",              "0"),
            ("Depreciation WDV (%)",      "depreciation_rate_pct", "15"),
            ("Tax Rate (%)",              "tax_rate_pct",          "25"),
            ("Other Income p.a. (₹)",     "other_income_pa",       "0"),
            ("Target MPBF (₹, CMA)",      "target_mpbf",           "0"),
            ("MPBF Method (1 or 2)",      "mpbf_method",           "2"),
        ]
        self._proj_inputs: dict[str, QLineEdit] = {}
        for i, (label, key, default) in enumerate(fields):
            r, c = divmod(i, 3)
            wrap = QVBoxLayout()
            wrap.setSpacing(2)
            lbl = QLabel(label)
            lbl.setProperty("role", "metric-label")
            edit = QLineEdit(self._settings.get(f"proj.{key}", default))
            edit.setMaximumWidth(160)
            wrap.addWidget(lbl)
            wrap.addWidget(edit)
            container = QWidget()
            container.setLayout(wrap)
            layout.addWidget(container, r, c)
            self._proj_inputs[key] = edit

        # ── CMA Solver Mode (radio buttons, spans full width below grid) ──
        n_rows = (len(fields) + 2) // 3
        mode_box = QGroupBox("CMA SOLVER MODE  (when Target MPBF > 0)")
        mode_lay = QHBoxLayout(mode_box)
        mode_lay.setContentsMargins(8, 16, 8, 8)
        self._solver_mode_group = QButtonGroup(mode_box)
        saved_mode = self._settings.get("proj.solver_mode", "forward")
        modes = [
            ("forward",         "Forward  (compute MPBF from inputs)"),
            ("reverse_days",    "Reverse — Days anchor  (solve Revenue)"),
            ("reverse_revenue", "Reverse — Revenue anchor  (solve Days)"),
        ]
        self._solver_radios: dict[str, QRadioButton] = {}
        for val, label in modes:
            rb = QRadioButton(label)
            if val == saved_mode:
                rb.setChecked(True)
            rb.setProperty("solver_mode", val)
            self._solver_mode_group.addButton(rb)
            self._solver_radios[val] = rb
            mode_lay.addWidget(rb)
        mode_lay.addStretch(1)
        layout.addWidget(mode_box, n_rows, 0, 1, 3)
        return gb

    # ── Branch ops ──────────────────────────────────────────────────

    def _add_branch(self):
        last_dir = self._settings.get("last_dir", str(Path.home()))
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Tally export",
            last_dir,
            "All supported (*.zip *.sqlite *.db);;Tally ZIP (*.zip);;SQLite (*.sqlite *.db);;All files (*)",
        )
        if not path:
            return
        self._settings["last_dir"] = str(Path(path).parent)
        save_settings(self._settings)

        default_name = f"Branch {len(self._branches) + 1}"
        # Try to infer a nicer name from the filename
        stem = Path(path).stem
        if " - " in stem:
            default_name = stem.split(" - ")[1].split("_")[0]
        elif "_" in stem:
            default_name = stem.split("_")[0]

        name, ok = QInputDialog.getText(
            self, "Branch Name",
            "Give this branch a short name (used in column headers):",
            text=default_name,
        )
        if not ok or not name.strip():
            return

        self._set_status(f"Loading {Path(path).name}…  (large files may take a few seconds)")
        self._btn_actual.setEnabled(False)
        self._btn_proj.setEnabled(False)

        self._loader = LoadWorker(path, name.strip())
        self._loader.succeeded.connect(self._on_load_ok)
        self._loader.failed.connect(self._on_load_fail)
        self._loader.start()

    def _on_load_ok(self, fd: fs.FinancialData, path: str):
        name = fd.branch_name
        self._branches.append({"path": path, "name": name, "fd": fd})
        self._refresh_branch_tree()
        self._refresh_consolidated()
        self._set_status(f"Added branch ‘{name}’  •  {len(self._branches)} branch(es) loaded.")
        self._btn_actual.setEnabled(True)
        self._btn_proj.setEnabled(True)

    def _on_load_fail(self, path: str, err: str):
        QMessageBox.critical(self, "Load failed",
                             f"Could not load:\n{path}\n\n{err.splitlines()[0]}")
        self._set_status("Load failed — check the file and try again.")
        self._btn_actual.setEnabled(True)
        self._btn_proj.setEnabled(True)

    def _rename_branch(self):
        idx = self._selected_idx()
        if idx is None:
            QMessageBox.information(self, "Rename", "Select a branch first.")
            return
        b = self._branches[idx]
        name, ok = QInputDialog.getText(self, "Rename branch", "New branch name:", text=b["name"])
        if not ok or not name.strip():
            return
        b["name"] = name.strip()
        if b["fd"]:
            b["fd"].branch_name = name.strip()
            for l in b["fd"].ledgers:
                l.branch = name.strip()
            for p in b["fd"].pnl_rows:
                p.branch = name.strip()
            for s in b["fd"].stock_items:
                s.branch = name.strip()
        self._refresh_branch_tree()
        self._refresh_consolidated()

    def _remove_branch(self):
        idx = self._selected_idx()
        if idx is None:
            QMessageBox.information(self, "Remove", "Select a branch first.")
            return
        removed = self._branches.pop(idx)
        self._refresh_branch_tree()
        self._refresh_consolidated()
        self._set_status(f"Removed ‘{removed['name']}’.")

    def _clear(self):
        if not self._branches:
            return
        if QMessageBox.question(self, "Clear all", "Remove all loaded branches?") != QMessageBox.Yes:
            return
        self._branches.clear()
        self._refresh_branch_tree()
        self._refresh_consolidated()
        self._set_status("Cleared.  Add Tally ZIP exports to begin.")

    def _selected_idx(self) -> int | None:
        item = self._tree.currentItem()
        if item is None:
            return None
        idx = self._tree.indexOfTopLevelItem(item)
        return idx if idx >= 0 else None

    def _refresh_branch_tree(self):
        self._tree.clear()
        for b in self._branches:
            fd = b["fd"]
            period = f"{fd.period_from} → {fd.period_to}" if fd else "—"
            company = fd.company if fd else "(load failed)"
            item = QTreeWidgetItem([b["name"], company, period, Path(b["path"]).name])
            self._tree.addTopLevelItem(item)

    # ── Stock override ──────────────────────────────────────────────

    def _parse_amount(self, s: str) -> float | None:
        s = s.strip().replace(",", "").replace("₹", "")
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None

    def _apply_stock_override(self):
        if not self._branches:
            QMessageBox.warning(self, "No branches", "Add at least one Tally export first.")
            return
        self._refresh_consolidated()
        self._set_status("Stock overrides applied.")

    # ── Consolidate + refresh summary card ──────────────────────────

    def _refresh_consolidated(self):
        loaded = [b["fd"] for b in self._branches if b["fd"] is not None]
        if not loaded:
            self._fd = None
            self._vr = None
            self._summary_card.setVisible(False)
            return
        try:
            self._fd = fs.consolidate_branches(loaded)
        except RuntimeError as e:
            self._fd = None
            self._vr = None
            self._summary_card.setVisible(False)
            QMessageBox.critical(self, "Consolidation failed", str(e))
            return

        self._fd.opening_stock_override = self._parse_amount(self._op_stock.text())
        self._fd.closing_stock_override = self._parse_amount(self._cl_stock.text())
        self._update_summary_card()

    def _update_summary_card(self):
        fd = self._fd
        if not fd:
            return
        rev = fd.revenue_from_ops()
        gp  = rev - fd.purchases() + (fd.closing_stock() - fd.opening_stock())
        pbt = fd.profit_before_tax()
        pat = fd.profit_after_tax()
        diff = fd.total_assets() - fd.total_equity_liabilities()

        self._company_lbl.setText(fd.company)
        self._period_lbl.setText(f"For the year ended {fd.period_label}   ·   "
                                  f"{fd.period_from} → {fd.period_to}")

        self._tile_rev.set_value(rev)
        self._tile_gp.set_value(gp, C_SUCCESS if gp > 0 else C_ERROR)
        self._tile_pbt.set_value(pbt, C_SUCCESS if pbt > 0 else C_ERROR)
        self._tile_pat.set_value(pat, C_SUCCESS if pat > 0 else C_ERROR)
        self._tile_assets.set_value(fd.total_assets())
        if abs(diff) < 1:
            self._tile_diff.set_value(0, C_SUCCESS)
        elif abs(diff) < 50_000:
            self._tile_diff.set_value(diff, C_WARN)
        else:
            self._tile_diff.set_value(diff, C_ERROR)

        vr = fs.validate_financial_data(fd)
        self._vr = vr
        if vr.has_errors:
            self._badge_lbl.setProperty("role", "badge-err")
            errs = sum(1 for c in vr.checks if c.severity == fs.ERROR)
            self._badge_lbl.setText(f"  {errs} ERROR(S)  •  click to inspect  ")
        elif vr.warnings:
            self._badge_lbl.setProperty("role", "badge-warn")
            n = len(vr.warnings)
            self._badge_lbl.setText(f"  {n} warning(s)  •  click to inspect  ")
        else:
            self._badge_lbl.setProperty("role", "badge-ok")
            self._badge_lbl.setText("  Balance Sheet OK  ✓  click to inspect  ")
        # Force style refresh
        self._badge_lbl.style().unpolish(self._badge_lbl)
        self._badge_lbl.style().polish(self._badge_lbl)

        self._summary_card.setVisible(True)

    def _on_badge_clicked(self, event):
        if self._vr is None:
            return
        dlg = ValidationDialog(self._vr, self)
        dlg.exec()

    # ── Projection card ─────────────────────────────────────────────

    def _toggle_proj(self):
        self._proj_open = not self._proj_open
        self._proj_card.setVisible(self._proj_open)
        self._proj_toggle_btn.setText(
            "▼  3-Year Projection Inputs  (click to collapse)" if self._proj_open
            else "▶  3-Year Projection Inputs  (click to expand)"
        )

    def _gather_proj(self) -> fs.ProjectionInputs | None:
        try:
            args = {}
            for key, edit in self._proj_inputs.items():
                args[key] = float(edit.text().replace(",", "").strip() or "0")
                self._settings[f"proj.{key}"] = edit.text()
            if "mpbf_method" in args:
                args["mpbf_method"] = int(args["mpbf_method"])
            # Solver mode from radio group
            mode = "forward"
            for val, rb in self._solver_radios.items():
                if rb.isChecked():
                    mode = val
                    break
            args["solver_mode"] = mode
            self._settings["proj.solver_mode"] = mode
            save_settings(self._settings)
            return fs.ProjectionInputs(**args)
        except ValueError as e:
            QMessageBox.warning(self, "Invalid input",
                                f"Please enter valid numbers in the projection fields.\n{e}")
            return None

    # ── Generate Excel ──────────────────────────────────────────────

    def _on_out_changed(self):
        self._out_dir = self._out_edit.text().strip() or str(Path.home() / "Desktop")
        self._settings["out_dir"] = self._out_dir
        save_settings(self._settings)

    def _pick_out_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Select output folder", self._out_dir)
        if d:
            self._out_dir = d
            self._out_edit.setText(d)
            self._settings["out_dir"] = d
            save_settings(self._settings)

    def _output_path(self, suffix: str) -> str:
        fd = self._fd
        company_short = (fd.company.split()[0] if fd else "Company").replace(",", "").replace(".", "")
        year = fd.period_to[:4] if fd and fd.period_to else "2026"
        return str(Path(self._out_dir) / f"{company_short}_FinStatements_{year}{suffix}.xlsx")

    def _branch_fds(self) -> list[fs.FinancialData]:
        return [b["fd"] for b in self._branches if b["fd"] is not None]

    def _ensure_fd(self) -> bool:
        if self._fd is None:
            QMessageBox.warning(self, "No data",
                                "Add at least one Tally ZIP and let it load first.")
            return False
        return True

    def _gen_actual(self):
        if not self._ensure_fd():
            return
        self._refresh_consolidated()
        out = self._output_path("_Actual")
        self._save_excel(out, proj=None)

    def _gen_all(self):
        if not self._ensure_fd():
            return
        proj = self._gather_proj()
        if proj is None:
            return
        self._refresh_consolidated()
        out = self._output_path("_Full")
        self._save_excel(out, proj=proj)

    def _gen_cma(self):
        if not self._ensure_fd():
            return
        proj = self._gather_proj()
        if proj is None:
            return
        self._refresh_consolidated()
        out = self._output_path("_CMA")
        self._save_excel(out, proj=proj, cma=True)

    def _save_excel(self, out: str, proj, cma: bool = False):
        self._set_status(f"Writing {Path(out).name} …")
        QApplication.processEvents()
        try:
            fs.ExcelWriter(self._fd, proj=proj, branches=self._branch_fds(), cma=cma).save(out)
            self._set_status(f"✓ Saved: {out}")
            self._settings["auto_open"] = self._cb_open.isChecked()
            save_settings(self._settings)
            if self._cb_open.isChecked():
                self._open_file(out)
            else:
                QMessageBox.information(self, "Done", f"Excel saved:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not write Excel:\n{e}")
            self._set_status("Generation failed — see error dialog.")

    def _open_file(self, path: str):
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)   # type: ignore
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception:
            pass

    # ── Status bar ──────────────────────────────────────────────────

    def _set_status(self, msg: str):
        self._status.showMessage(msg)


# ─────────────────────── Entry point ───────────────────────

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Tally Financial Statements")
    app.setStyle("Fusion")

    pal = QPalette()
    pal.setColor(QPalette.Window, QColor(C_BG))
    pal.setColor(QPalette.WindowText, QColor(C_TEXT))
    pal.setColor(QPalette.Base, QColor(C_SURFACE))
    pal.setColor(QPalette.AlternateBase, QColor(C_SURFACE_2))
    pal.setColor(QPalette.Text, QColor(C_TEXT))
    pal.setColor(QPalette.Button, QColor(C_SURFACE_2))
    pal.setColor(QPalette.ButtonText, QColor(C_TEXT))
    pal.setColor(QPalette.Highlight, QColor(C_ACCENT))
    pal.setColor(QPalette.HighlightedText, QColor("white"))
    app.setPalette(pal)

    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
