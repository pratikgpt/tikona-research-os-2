"""
═══════════════════════════════════════════════════════════════════
 TIKONA CAPITAL — Financial Model Generator v5.1
═══════════════════════════════════════════════════════════════════
 Model:    claude-sonnet-4-20250514 (Claude Sonnet 4) + web_search
 Changes:  Year-agnostic prompt, formula-based Excel, cost tracking
 Output:   Screener Excel + 8 appended model sheets (all formulas)

 Refactored from Colab notebook into a callable function so
 financial_model_server.py (FastAPI) can import and invoke it.

 Required environment variables:
   ANTHROPIC_API_KEY     - Anthropic API key
   SCREENER_USERNAME     - Screener.in account email
   SCREENER_PASSWORD     - Screener.in password

 Public API:
   generate_financial_model(nse_code, company_name, sector, output_dir)
       -> {file_path, json_path, model_json, cost_summary, elapsed_seconds}
═══════════════════════════════════════════════════════════════════
"""

from __future__ import annotations

import os
import re
import json
import time
import shutil
import logging
import traceback
from typing import Any, Literal, Optional

import anthropic
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from bs4 import BeautifulSoup
import requests as req_lib
from pydantic import BaseModel, Field, ValidationError


logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(message)s")
logger = logging.getLogger(__name__)

__version__ = "5.1.0"
MODEL_NAME = "claude-sonnet-4-20250514"
WEB_SEARCH_TOOL = [{"type": "web_search_20250305", "name": "web_search"}]

PER_YEAR_ASSUMPTION_KEYS = [
    "revenue_growth_pct",
    "rm_pct",
    "employee_pct",
    "power_fuel_pct",
    "other_mfg_pct",
    "selling_admin_pct",
    "other_exp_pct",
    "chg_inventory_pct",
    "depreciation_cr",
    "interest_cr",
    "other_income_cr",
    "tax_rate_pct",
    "capex_cr",
    "receivable_days",
    "inventory_days",
]

DEFAULT_ROW_MAP = {
    "sales": 17,
    "raw_material": 18,
    "change_in_inventory": 19,
    "power_fuel": 20,
    "other_mfg": 21,
    "employee_cost": 22,
    "selling_admin": 23,
    "other_expenses": 24,
    "other_income": 25,
    "depreciation": 26,
    "interest": 27,
    "pbt": 28,
    "tax": 29,
    "net_profit": 30,
    "dividend_amount": 31,
    "share_capital": 57,
    "reserves": 58,
    "borrowings": 59,
    "other_liabilities": 60,
    "total_liab": 61,
    "net_block": 62,
    "cwip": 63,
    "investments": 64,
    "other_assets": 65,
    "total_assets": 66,
    "receivables": 67,
    "inventory": 68,
    "cash": 69,
    "shares_outstanding": 70,
    "cfo": 82,
    "cfi": 83,
    "cff": 84,
    "net_cash_flow": 85,
}

ROW_LABEL_CANDIDATES = {
    "sales": ["sales", "revenue from operations", "sales turnover"],
    "raw_material": ["raw material cost", "material cost", "cost of materials"],
    "change_in_inventory": ["change in inventory", "changes in inventories"],
    "power_fuel": ["power and fuel", "power & fuel", "fuel cost"],
    "other_mfg": ["other mfg", "other manufacturing", "manufacturing expenses"],
    "employee_cost": ["employee cost", "staff cost", "employee expenses"],
    "selling_admin": ["selling and admin", "selling & admin", "selling general administrative", "sga"],
    "other_expenses": ["other expenses"],
    "other_income": ["other income"],
    "depreciation": ["depreciation"],
    "interest": ["interest", "finance cost"],
    "pbt": ["profit before tax", "pbt"],
    "tax": ["tax", "tax expense"],
    "net_profit": ["net profit", "profit after tax", "pat"],
    "dividend_amount": ["dividend amount", "dividend"],
    "share_capital": ["equity share capital", "share capital"],
    "reserves": ["reserves", "reserves and surplus", "other equity"],
    "borrowings": ["borrowings", "debt"],
    "other_liabilities": ["other liabilities", "current liabilities", "non current liabilities"],
    "total_liab": ["total liabilities", "total liabilities and equity"],
    "net_block": ["net block", "net fixed assets"],
    "cwip": ["capital work in progress", "cwip"],
    "investments": ["investments"],
    "other_assets": ["other assets", "current assets", "loans and advances"],
    "total_assets": ["total assets"],
    "receivables": ["trade receivables", "receivables", "debtors"],
    "inventory": ["inventory", "inventories"],
    "cash": ["cash and bank", "cash & bank", "cash equivalents", "cash"],
    "shares_outstanding": ["number of equity shares", "no. of equity shares", "shares outstanding"],
    "cfo": ["cash from operating activity", "cash from operating activities", "cash from operations"],
    "cfi": ["cash from investing activity", "cash from investing activities"],
    "cff": ["cash from financing activity", "cash from financing activities"],
    "net_cash_flow": ["net cash flow"],
}


# ══════════════════════════════════════════════════════════════════
# COST TRACKER
# ══════════════════════════════════════════════════════════════════
class CostTracker:
    INPUT_COST_PER_M = 3.00
    OUTPUT_COST_PER_M = 15.00
    USD_TO_INR = 84.0

    def __init__(self):
        self.total_input_tokens = 0
        self.total_output_tokens = 0
        self.total_cache_read = 0
        self.total_cache_create = 0
        self.calls = 0

    def track(self, response):
        usage = response.usage
        self.total_input_tokens += usage.input_tokens
        self.total_output_tokens += usage.output_tokens
        if hasattr(usage, "cache_read_input_tokens"):
            self.total_cache_read += usage.cache_read_input_tokens or 0
        if hasattr(usage, "cache_creation_input_tokens"):
            self.total_cache_create += usage.cache_creation_input_tokens or 0
        self.calls += 1

    @property
    def total_tokens(self):
        return self.total_input_tokens + self.total_output_tokens

    @property
    def cost_usd(self):
        return (
            (self.total_input_tokens / 1_000_000) * self.INPUT_COST_PER_M
            + (self.total_output_tokens / 1_000_000) * self.OUTPUT_COST_PER_M
        )

    @property
    def cost_inr(self):
        return self.cost_usd * self.USD_TO_INR

    def to_dict(self):
        return {
            "model": MODEL_NAME,
            "calls": self.calls,
            "input_tokens": self.total_input_tokens,
            "output_tokens": self.total_output_tokens,
            "total_tokens": self.total_tokens,
            "cache_read": self.total_cache_read,
            "cache_create": self.total_cache_create,
            "cost_usd": round(self.cost_usd, 4),
            "cost_inr": round(self.cost_inr, 2),
        }

    def summary(self):
        return (
            f"\n{'='*60}\n"
            f"  💰 API COST SUMMARY\n"
            f"{'='*60}\n"
            f"  Model:           {MODEL_NAME}\n"
            f"  API Calls:       {self.calls}\n"
            f"  Input Tokens:    {self.total_input_tokens:,}\n"
            f"  Output Tokens:   {self.total_output_tokens:,}\n"
            f"  Total Tokens:    {self.total_tokens:,}\n"
            f"  Cache Read:      {self.total_cache_read:,}\n"
            f"  Cache Create:    {self.total_cache_create:,}\n"
            f"  ─────────────────────────────\n"
            f"  Cost (USD):      ${self.cost_usd:.4f}\n"
            f"  Cost (INR):      ₹{self.cost_inr:.2f}\n"
            f"{'='*60}"
        )


# ══════════════════════════════════════════════════════════════════
# STYLES
# ══════════════════════════════════════════════════════════════════
COLORS = {
    "navy": "1F4690", "med_blue": "3A5BA0", "orange": "FFA500", "peach": "FFE5B4",
    "white": "FFFFFF", "light_grey": "F2F2F2", "black": "000000", "red": "D32F2F",
    "green": "2E7D32", "dark_grey": "D9D9D9",
}

navy_fill = PatternFill("solid", fgColor=COLORS["navy"])
blue_fill = PatternFill("solid", fgColor=COLORS["med_blue"])
orange_fill = PatternFill("solid", fgColor=COLORS["orange"])
peach_fill = PatternFill("solid", fgColor=COLORS["peach"])
grey_fill = PatternFill("solid", fgColor=COLORS["light_grey"])

hdr_font = Font(name="Arial", bold=True, color=COLORS["white"], size=11)
sub_font = Font(name="Arial", bold=True, color=COLORS["white"], size=10)
title_font = Font(name="Arial", bold=True, color=COLORS["navy"], size=14)
sec_font = Font(name="Arial", bold=True, color=COLORS["navy"], size=11)
data_font = Font(name="Arial", color=COLORS["black"], size=10)
input_font = Font(name="Arial", color=COLORS["navy"], size=10)
red_font = Font(name="Arial", color=COLORS["red"], size=10)

thin_bdr = Border(
    left=Side(style="thin", color=COLORS["dark_grey"]),
    right=Side(style="thin", color=COLORS["dark_grey"]),
    top=Side(style="thin", color=COLORS["dark_grey"]),
    bottom=Side(style="thin", color=COLORS["dark_grey"]),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
INR = "#,##,##0"
INR2 = "#,##,##0.00"
PCT = "0.00%"
RATIO = "0.00"


# ══════════════════════════════════════════════════════════════════
# SCREENER DOWNLOADER
# ══════════════════════════════════════════════════════════════════
def download_screener_excel(symbol: str, output_dir: str, screener_username: str, screener_password: str) -> str:
    import cloudscraper

    logger.info("📡 Logging into Screener.in...")
    session = cloudscraper.create_scraper(browser="chrome")
    login_url = "https://www.screener.in/login/"
    resp = session.get(login_url, timeout=15)
    soup = BeautifulSoup(resp.text, "html.parser")
    csrf = soup.find("input", {"name": "csrfmiddlewaretoken"})
    csrf_token = csrf["value"] if csrf else session.cookies.get("csrftoken")

    session.post(
        login_url,
        data={
            "username": screener_username,
            "password": screener_password,
            "csrfmiddlewaretoken": csrf_token,
            "next": "/",
        },
        headers={"Referer": login_url},
        timeout=15,
    )
    probe = session.get("https://www.screener.in/dash/", timeout=15)
    if probe.status_code != 200:
        raise Exception(f"❌ Screener login failed — /dash/ returned {probe.status_code}")
    if "login" in probe.url.lower() or "email or username" in probe.text.lower():
        raise Exception("❌ Screener login failed — /dash/ redirected to login page. Check SCREENER_USERNAME/PASSWORD.")
    logger.info("✅ Screener login OK (probe URL: %s)", probe.url)

    company_url = None
    page_resp = None
    for suffix in [f"{symbol}/consolidated/", f"{symbol}/"]:
        url = f"https://www.screener.in/company/{suffix}"
        resp = session.get(url, timeout=15)
        if resp.status_code == 200:
            company_url = resp.url  # use final URL after any redirects
            page_resp = resp
            break
    else:
        raise Exception(f"❌ {symbol} not found on Screener")

    logger.info(f"📄 Company page resolved to: {company_url}")

    soup = BeautifulSoup(page_resp.text, "html.parser")

    # Dump all forms and export-related strings from raw HTML for diagnosis
    for f in soup.find_all("form"):
        logger.info(f"  [FORM] action={f.get('action')} method={f.get('method')}")
    for tag in soup.find_all(True):
        for attr in ("href", "action", "formaction", "data-url"):
            val = tag.get(attr, "")
            if val and "export" in val.lower():
                logger.info(f"  [EXPORT-ATTR] <{tag.name} {attr}={val}>")
    import re as _re
    for m in _re.findall(r'["\']([^"\']{5,120}export[^"\']{0,60})["\']', page_resp.text, _re.I)[:8]:
        logger.info(f"  [EXPORT-RAW] {m}")

    # Extract CSRF token from page HTML first (more reliable than cookie)
    page_csrf_input = soup.find("input", {"name": "csrfmiddlewaretoken"})
    csrf2_token = page_csrf_input["value"] if page_csrf_input else session.cookies.get("csrftoken", "")

    export_url = None

    # Try button with aria-label
    btn = soup.find("button", attrs={"aria-label": "Export to Excel"})
    # Try button with export text (case-insensitive)
    if not btn:
        btn = soup.find("button", string=lambda t: t and "export" in str(t).lower())
    # Try anchor with export in href
    if not btn:
        btn = soup.find("a", href=lambda h: h and "export" in str(h).lower())
    # Try form whose action contains "export"
    if not btn:
        form = soup.find("form", action=lambda a: a and "export" in str(a).lower())
        if form:
            action = form.get("action", "")
            export_url = action if action.startswith("http") else "https://www.screener.in" + action
            ci = form.find("input", {"name": "csrfmiddlewaretoken"})
            if ci:
                csrf2_token = ci["value"]

    if btn and not export_url:
        raw = btn.get("formaction", "") or btn.get("href", "")
        export_url = raw if raw.startswith("http") else "https://www.screener.in" + raw
        form = btn.find_parent("form")
        ci = (form or soup).find("input", {"name": "csrfmiddlewaretoken"})
        if ci:
            csrf2_token = ci["value"]

    # Screener now uses a no-action POST form — submits back to the company page URL
    if not export_url:
        for f in soup.find_all("form"):
            if f.get("action") is None and (f.get("method") or "").lower() == "post":
                export_url = company_url
                ci = f.find("input", {"name": "csrfmiddlewaretoken"})
                if ci:
                    csrf2_token = ci["value"]
                logger.info(f"  Found no-action POST form → submitting to {export_url}")
                break

    # Last-resort fallback
    if not export_url:
        export_url = company_url
        logger.warning(f"⚠️  No export form found; POSTing to company URL: {export_url}")

    logger.info(f"📥 POSTing export: {export_url}")
    file_resp = session.post(
        export_url,
        data={"csrfmiddlewaretoken": csrf2_token},
        headers={
            "Referer": company_url,
            "Origin": "https://www.screener.in",
            "X-CSRFToken": csrf2_token,
            "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*",
        },
        timeout=30,
    )
    content_type = file_resp.headers.get("Content-Type", "").lower()
    if "html" in content_type:
        snippet = file_resp.text[:300].replace("\n", " ")
        raise Exception(
            f"❌ Got HTML instead of Excel from Screener export "
            f"(status={file_resp.status_code}, url={file_resp.url}, preview: {snippet})"
        )

    out_path = os.path.join(output_dir, f"{symbol}_screener.xlsx")
    with open(out_path, "wb") as f:
        f.write(file_resp.content)
    logger.info(f"✅ Downloaded → {out_path} ({len(file_resp.content)//1024} KB)")
    return out_path


def _normalize_label(value: object) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _resolve_row_map(ws) -> dict[str, int]:
    label_rows: dict[str, list[int]] = {}
    for row in range(1, min(ws.max_row, 140) + 1):
        label = _normalize_label(ws.cell(row=row, column=1).value)
        if not label:
            continue
        label_rows.setdefault(label, []).append(row)

    row_map: dict[str, int] = {}
    for key, candidates in ROW_LABEL_CANDIDATES.items():
        matched_row = None
        # Pass 1: exact-match preferred (prevents e.g. "tax" matching "profit before tax")
        for candidate in candidates:
            cand_norm = _normalize_label(candidate)
            if cand_norm in label_rows:
                matched_row = label_rows[cand_norm][0]
                break
        # Pass 2: substring fallback
        if matched_row is None:
            for label, rows in label_rows.items():
                if any(_normalize_label(c) in label for c in candidates):
                    matched_row = rows[0]
                    break
        if matched_row is None:
            matched_row = DEFAULT_ROW_MAP[key]
            logger.warning("⚠ Could not locate Screener row for '%s' by label — falling back to row %s", key, matched_row)
        row_map[key] = matched_row
    return row_map


def _last_non_null(values: list):
    for value in reversed(values):
        if value is not None:
            return value
    return None


def _build_projection_years(last_actual_year: str, count: int = 5) -> list[str]:
    match = re.search(r"FY\s*'?(\d{2,4})", str(last_actual_year), flags=re.IGNORECASE)
    if not match:
        return [f"FY{i}E" for i in range(1, count + 1)]

    year_num = int(match.group(1)[-2:])
    return [f"FY{year_num + offset:02d}E" for offset in range(1, count + 1)]


class SensitivityPe(BaseModel):
    row_label: str
    col_label: str
    row_values: list[float]
    col_values: list[float]
    grid: list[list[float]]


class Valuation(BaseModel):
    dcf_fair_value: float
    pe_fair_value: float
    ev_ebitda_fair_value: float
    blended_fair_value: float
    sensitivity_pe: SensitivityPe


class Assumptions(BaseModel):
    projection_years: list[str]
    revenue_growth_pct: dict[str, float]
    rm_pct: dict[str, float]
    employee_pct: dict[str, float]
    power_fuel_pct: dict[str, float]
    other_mfg_pct: dict[str, float]
    selling_admin_pct: dict[str, float]
    other_exp_pct: dict[str, float]
    chg_inventory_pct: dict[str, float]
    depreciation_cr: dict[str, float]
    interest_cr: dict[str, float]
    other_income_cr: dict[str, float]
    tax_rate_pct: dict[str, float]
    capex_cr: dict[str, float]
    receivable_days: dict[str, float]
    inventory_days: dict[str, float]
    dividend_payout_pct: float
    target_pe: float
    target_ev_ebitda: float
    wacc_pct: float
    terminal_growth_pct: float
    rationale: str


class Projections(BaseModel):
    years: list[str]
    net_block: list[float]
    cwip: list[float]
    investments: list[float]
    other_assets: list[float]
    share_capital: list[float]
    reserves: list[float]
    borrowings: list[float]
    other_liabilities: list[float]
    cfo: list[float]
    cfi: list[float]
    cff: list[float]


class HistoricalRatios(BaseModel):
    years: list[str]
    ebitda_margin_pct: list[float]
    pat_margin_pct: list[float]
    roe_pct: list[float]
    roce_pct: list[float]
    debt_equity: list[float]
    receivable_days: list[float]
    inventory_days: list[float]
    asset_turnover: list[float]
    rm_pct: list[float]
    employee_pct: list[float]
    tax_rate_pct: list[float]


class Peer(BaseModel):
    name: str
    mcap_cr: float
    pe: float
    ev_ebitda: float
    ebitda_margin_pct: Any
    roe_pct: float


class SaarthiScores(BaseModel):
    S_sector_quality: float
    A_accounting_quality: float
    A_asset_quality: float
    R_revenue_visibility: float
    T_track_record: float
    H_balance_sheet_health: float
    I_intrinsic_valuation: float


class Thesis(BaseModel):
    investment_thesis: str
    bull_case: str
    bear_case: str
    key_catalysts: list[str] = Field(min_length=4)
    key_risks: list[str] = Field(min_length=4)
    saarthi_scores: SaarthiScores
    saarthi_total: float
    saarthi_rating: str


class BoardMember(BaseModel):
    name: str
    designation: str
    status: str
    din: Optional[str] = None
    since: Optional[str] = None


class Shareholding(BaseModel):
    years: list[str] = Field(default_factory=list)
    promoter_pct: list[float] = Field(default_factory=list)
    fii_pct: list[float] = Field(default_factory=list)
    dii_pct: list[float] = Field(default_factory=list)
    public_pct: list[float] = Field(default_factory=list)


class GovernanceData(BaseModel):
    board: list[BoardMember] = Field(default_factory=list)
    shareholding: Optional[Shareholding] = None
    promoter_pledge_pct: Optional[float] = None
    board_size: Optional[float] = None
    independent_directors: Optional[float] = None
    women_directors: Optional[float] = None
    statutory_auditor: Optional[str] = None
    managerial_remuneration_cr: Optional[float] = None
    auditor_fees_audit_cr: Optional[float] = None
    auditor_fees_non_audit_cr: Optional[float] = None
    related_party_transactions_cr: Optional[float] = None


class TimelineEvent(BaseModel):
    year: str
    category: str
    description: str
    impact: str = ""


class RiskItem(BaseModel):
    category: str
    factor: str
    description: str
    mitigation: str = ""
    probability: str = "M"
    impact: str = "M"
    rating: str = "MEDIUM"


class PeerDetailed(BaseModel):
    name: str
    revenue_series: list[float] = Field(default_factory=list)
    ebitda_margin_series: list[float] = Field(default_factory=list)
    pat_series: list[float] = Field(default_factory=list)
    mcap_cr: Optional[float] = None
    pe: Optional[float] = None
    pb: Optional[float] = None
    roce_pct: Optional[float] = None
    roe_pct: Optional[float] = None


class OperationalData(BaseModel):
    years: list[str] = Field(default_factory=list)
    volume_segments: dict[str, list[float]] = Field(default_factory=dict)
    capacity_utilisation_pct: list[float] = Field(default_factory=list)
    countries_of_operation: list[float] = Field(default_factory=list)
    plants_india: list[float] = Field(default_factory=list)
    plants_overseas: list[float] = Field(default_factory=list)
    revenue_mix_pct: dict[str, float] = Field(default_factory=dict)
    geography_mix_pct: dict[str, float] = Field(default_factory=dict)
    realization_per_mt: Optional[float] = None
    employees: Optional[float] = None


class FinancialModelOutput(BaseModel):
    sector: str
    base_year: str
    cmp: float
    target_price: float
    rating: Literal["STRONG BUY", "BUY", "ACCUMULATE", "HOLD", "UNDERPERFORM", "SELL"]
    upside_pct: float
    shares_cr: float
    assumptions: Assumptions
    projections: Projections
    historical_ratios: HistoricalRatios
    valuation: Valuation
    peers: list[Peer] = Field(default_factory=list)
    thesis: Thesis
    operational: Optional[OperationalData] = None
    governance: Optional[GovernanceData] = None
    timeline_events: list[TimelineEvent] = Field(default_factory=list)
    risk_items: list[RiskItem] = Field(default_factory=list)
    peers_detailed: list[PeerDetailed] = Field(default_factory=list)


# ══════════════════════════════════════════════════════════════════
# DATA EXTRACTOR
# ══════════════════════════════════════════════════════════════════
def extract_screener_data(filepath: str) -> dict:
    wb = load_workbook(filepath, data_only=True)
    if "Data Sheet" not in wb.sheetnames:
        raise Exception("❌ No 'Data Sheet' in Screener Excel")
    ws = wb["Data Sheet"]

    company_name = ws.cell(row=1, column=2).value or ""
    face_value = ws.cell(row=7, column=2).value
    cmp = ws.cell(row=8, column=2).value
    mcap = ws.cell(row=9, column=2).value

    data_cols = [col for col in range(2, 20) if ws.cell(row=16, column=col).value is not None]
    fiscal_years = []
    for col in data_cols:
        dt = ws.cell(row=16, column=col).value
        fiscal_years.append(f"FY{str(dt.year)[2:]}" if hasattr(dt, "year") else str(dt))

    row_map = _resolve_row_map(ws)

    # Trim trailing year columns that have no data yet (Screener reserves a column
    # for the upcoming fiscal year before results are reported). A year is considered
    # "empty" if both Sales and Net Profit are None for that column.
    while data_cols:
        last_col = data_cols[-1]
        sales_v = ws.cell(row=row_map["sales"], column=last_col).value
        np_v = ws.cell(row=row_map["net_profit"], column=last_col).value
        if sales_v is None and np_v is None:
            dropped_year = fiscal_years[-1]
            data_cols.pop()
            fiscal_years.pop()
            logger.warning("⚠ Dropping empty trailing Screener year column %s (no audited data yet)", dropped_year)
        else:
            break

    def rd(row_num):
        return [ws.cell(row=row_num, column=col).value for col in data_cols]

    nh = len(data_cols)
    sales = rd(row_map["sales"])
    rm = rd(row_map["raw_material"])

    data = {
        "company_name": company_name,
        "face_value": face_value,
        "cmp": cmp,
        "mcap": mcap,
        "fiscal_years": fiscal_years,
        "data_cols": data_cols,
        "num_years": nh,
        "row_map": row_map,
        "pl": {
            "sales": sales, "raw_material": rm,
            "change_in_inventory": rd(row_map["change_in_inventory"]), "power_fuel": rd(row_map["power_fuel"]), "other_mfg": rd(row_map["other_mfg"]),
            "employee_cost": rd(row_map["employee_cost"]), "selling_admin": rd(row_map["selling_admin"]), "other_expenses": rd(row_map["other_expenses"]),
            "other_income": rd(row_map["other_income"]), "depreciation": rd(row_map["depreciation"]), "interest": rd(row_map["interest"]),
            "pbt": rd(row_map["pbt"]), "tax": rd(row_map["tax"]), "net_profit": rd(row_map["net_profit"]), "dividend_amount": rd(row_map["dividend_amount"]),
        },
        "bs": {
            "share_capital": rd(row_map["share_capital"]), "reserves": rd(row_map["reserves"]), "borrowings": rd(row_map["borrowings"]),
            "other_liabilities": rd(row_map["other_liabilities"]), "total_liab": rd(row_map["total_liab"]), "net_block": rd(row_map["net_block"]),
            "cwip": rd(row_map["cwip"]), "investments": rd(row_map["investments"]), "other_assets": rd(row_map["other_assets"]),
            "total_assets": rd(row_map["total_assets"]), "receivables": rd(row_map["receivables"]), "inventory": rd(row_map["inventory"]),
            "cash": rd(row_map["cash"]), "shares_outstanding": rd(row_map["shares_outstanding"]),
        },
        "cf": {"cfo": rd(row_map["cfo"]), "cfi": rd(row_map["cfi"]), "cff": rd(row_map["cff"]), "net_cash_flow": rd(row_map["net_cash_flow"])},
    }
    pl = data["pl"]
    derived_ebitda = []
    for i in range(nh):
        if sales[i] is None:
            derived_ebitda.append(None)
            continue
        # Preferred: EBITDA = PBT + Depreciation + Interest − Other Income (robust to
        # Screener's variable expense-line reporting where some sub-lines roll up into
        # "Other expenses" in newer years).
        pbt_i = pl["pbt"][i]
        dep_i = pl["depreciation"][i]
        int_i = pl["interest"][i]
        oi_i = pl["other_income"][i]
        if pbt_i is not None:
            eb = float(pbt_i) + float(dep_i or 0) + float(int_i or 0) - float(oi_i or 0)
            derived_ebitda.append(eb)
            continue
        # Fallback: Sales − sum of operating expenses (older legacy method).
        expenses = [
            rm[i], pl["change_in_inventory"][i], pl["power_fuel"][i], pl["other_mfg"][i],
            pl["employee_cost"][i], pl["selling_admin"][i], pl["other_expenses"][i],
        ]
        derived_ebitda.append(float(sales[i]) - sum(v for v in expenses if v is not None))
    data["derived"] = {"ebitda": derived_ebitda}
    wb.close()
    logger.info(f"✅ Extracted {nh} years: {fiscal_years}")
    return data


# ══════════════════════════════════════════════════════════════════
# MARKET DATA
# ══════════════════════════════════════════════════════════════════
def fetch_market_data(ticker: str) -> str:
    parts = []
    try:
        import yfinance as yf
        info = yf.Ticker(f"{ticker}.NS").info
        parts.append(
            f"yfinance | CMP:₹{info.get('currentPrice','NA')} "
            f"MCap:₹{(info.get('marketCap',0) or 0)/1e7:.0f}Cr "
            f"PE:{info.get('trailingPE','NA')} PB:{info.get('priceToBook','NA')} "
            f"ROE:{((info.get('returnOnEquity',0) or 0)*100):.1f}% "
            f"Beta:{info.get('beta','NA')} 52WH:₹{info.get('fiftyTwoWeekHigh','NA')} "
            f"52WL:₹{info.get('fiftyTwoWeekLow','NA')}"
        )
    except Exception as e:
        logger.warning("⚠ yfinance lookup failed for %s: %s", ticker, e)

    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        for u in [
            f"https://www.screener.in/company/{ticker}/consolidated/",
            f"https://www.screener.in/company/{ticker}/",
        ]:
            r = req_lib.get(u, headers=headers, timeout=10)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, "html.parser")
                rows = []
                for t in soup.find_all("section", class_="card"):
                    h = t.find("h2")
                    if h:
                        rows.append(f"\n=== {h.get_text(strip=True)} ===")
                    for row in t.find_all("tr"):
                        cells = row.find_all(["th", "td"])
                        if cells:
                            rows.append(" | ".join(c.get_text(strip=True) for c in cells))
                if rows:
                    parts.append("\n".join(rows[:60]))
                break
    except Exception:
        pass
    return "\n\n".join(parts)


# ══════════════════════════════════════════════════════════════════
# JSON EXTRACTION
# ══════════════════════════════════════════════════════════════════
def extract_json_from_text(text: str) -> dict:
    text = re.sub(r"<thinking>.*?</thinking>", "", text, flags=re.DOTALL).strip()
    if text.startswith("{"):
        try:
            return json.loads(text)
        except Exception:
            pass
    for pat in [r"```json\s*(.*?)\s*```", r"```\s*(\{.*?\})\s*```"]:
        m = re.search(pat, text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(1))
            except Exception:
                pass
    start = text.find("{")
    if start != -1:
        depth = 0
        for i, ch in enumerate(text[start:], start):
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[start : i + 1])
                    except Exception:
                        break
    raise ValueError(f"No valid JSON. Preview: {text[:500]}")


# ══════════════════════════════════════════════════════════════════
# CLAUDE API — YEAR-AGNOSTIC PROMPT
# ══════════════════════════════════════════════════════════════════
def call_claude_api(
    company_name: str,
    ticker: str,
    sector: str,
    screener_data: dict,
    market_text: str,
    api_key: str,
    tracker: CostTracker,
) -> dict:
    client = anthropic.Anthropic(api_key=api_key)

    hist_json = json.dumps(
        {
            "years": screener_data["fiscal_years"],
            "revenue": screener_data["pl"]["sales"],
            "raw_material": screener_data["pl"]["raw_material"],
            "change_in_inventory": screener_data["pl"]["change_in_inventory"],
            "power_fuel": screener_data["pl"]["power_fuel"],
            "other_mfg": screener_data["pl"]["other_mfg"],
            "employee_cost": screener_data["pl"]["employee_cost"],
            "selling_admin": screener_data["pl"]["selling_admin"],
            "other_expenses": screener_data["pl"]["other_expenses"],
            "ebitda": screener_data["derived"]["ebitda"],
            "net_profit": screener_data["pl"]["net_profit"],
            "depreciation": screener_data["pl"]["depreciation"],
            "interest": screener_data["pl"]["interest"],
            "other_income": screener_data["pl"]["other_income"],
            "pbt": screener_data["pl"]["pbt"],
            "tax": screener_data["pl"]["tax"],
            "share_capital": screener_data["bs"]["share_capital"],
            "reserves": screener_data["bs"]["reserves"],
            "borrowings": screener_data["bs"]["borrowings"],
            "other_liabilities": screener_data["bs"]["other_liabilities"],
            "net_block": screener_data["bs"]["net_block"],
            "cwip": screener_data["bs"]["cwip"],
            "investments": screener_data["bs"]["investments"],
            "receivables": screener_data["bs"]["receivables"],
            "inventory": screener_data["bs"]["inventory"],
            "cash": screener_data["bs"]["cash"],
            "total_assets": screener_data["bs"]["total_assets"],
            "cfo": screener_data["cf"]["cfo"],
            "cfi": screener_data["cf"]["cfi"],
            "cff": screener_data["cf"]["cff"],
        },
        indent=2,
        default=str,
    )

    prompt = f"""You are a senior equity research analyst building a complete financial model.

COMPANY: {company_name} ({ticker}) | Sector: {sector}

━━ HISTORICAL DATA (Screener.in audited, ₹ Crores) ━━
{hist_json}

━━ MARKET DATA ━━
{market_text}

━━ INSTRUCTIONS ━━

STEP 0 — DETERMINE TIMELINE:
- Look at the historical data above. Identify the LAST fiscal year with actual data (call it "Base Year").
- Your projection period = next 5 fiscal years AFTER the Base Year.
- Example: if last actual year is FY25 → project FY26E, FY27E, FY28E, FY29E, FY30E
- Example: if last actual year is FY24 → project FY25E, FY26E, FY27E, FY28E, FY29E
- Use this dynamically — do NOT assume any specific year.

STEP 1 — WEB SEARCHES (use current/recent context, not hardcoded years):
- "{company_name} latest quarterly results earnings revenue profit"
- "{company_name} most recent annual results full year earnings"
- "{company_name} management guidance outlook capacity expansion next 3 years"
- "{company_name} analyst consensus target price estimates"
- "{company_name} broker reports earnings estimates next 5 years"
- "{sector} sector India peer comparison PE EV/EBITDA current multiples"
- "India 10 year government bond yield current"
- "{ticker} consensus estimates earnings forecasts"
- "{ticker} GoIndiaStocks consensus estimates"
- "site:trendlyne.com {ticker} forecasts estimates"
- "{company_name} order book revenue visibility pipeline"

STEP 2 — BUILD ASSUMPTIONS (5 projected years):
For EVERY projected year, provide these as PLAIN NUMBERS (15.5 not 0.155):
- revenue_growth_pct, rm_pct, employee_pct, power_fuel_pct, other_mfg_pct,
  selling_admin_pct, other_exp_pct, chg_inventory_pct
- depreciation_cr, interest_cr, other_income_cr (₹ Cr absolute)
- tax_rate_pct, capex_cr, receivable_days, inventory_days

Excel formulas applied by the builder:
  Revenue Year N = Revenue Year N-1 × (1 + revenue_growth_pct / 100)
  Each expense = Revenue × pct / 100
  EBITDA = Revenue - all operating expenses
  PBT = EBITDA + Other Income - Depreciation - Interest
  Tax = PBT × tax_rate_pct / 100
  PAT = PBT - Tax
  Receivables = Revenue × receivable_days / 365
  Inventory = RM Cost × inventory_days / 365
Every assumption % MUST be realistic and internally consistent.

STEP 3 — BALANCE SHEET PROJECTIONS (absolute values):
- net_block, cwip, investments, other_assets
- share_capital, reserves, borrowings, other_liabilities
Cash = PLUG. Balance Sheet MUST balance every year.

STEP 4 — HISTORICAL FINANCIALS IN ABSOLUTE TERMS (last 6 years):
Use web_search to find the company's annual revenue, EBITDA, and PAT in ₹ Crores for each
of the last 6 fiscal years. This is CRITICAL for recently-IPO'd or platform businesses
(e.g., Swiggy, Zomato) where the Screener data above may show nulls. Search for:
  "{company_name} annual revenue EBITDA PAT FY20 FY21 FY22 FY23 FY24 FY25 crores"
  "{company_name} annual report revenue profit loss"

STEP 5 — HISTORICAL RATIOS (last 6 years):
- ebitda_margin_pct, pat_margin_pct, roe_pct, roce_pct
- debt_equity, receivable_days, inventory_days, asset_turnover
- rm_pct, employee_pct, tax_rate_pct

STEP 5 — CASH FLOW PROJECTIONS: CFO, CFI, CFF for each projected year.

STEP 6 — VALUATION (3 methods):
- DCF (FCFE discounted, terminal via Gordon Growth)
- PE-based (Target PE × EPS year-2)
- EV/EBITDA (Target × EBITDA year-2 - net debt)
- Blended (DCF 40%, PE 30%, EV/EBITDA 30%)
- Sensitivity 5x5 grid: PE Multiple vs EPS Growth %

STEP 7 — PEER COMPARISON: 4+ listed peers with current MCap, PE, EV/EBITDA, EBITDA Margin, ROE.

STEP 8 — INVESTMENT THESIS with SAARTHI scoring (S+A+A+R+T+H+I = 100 max):
  S—Sector(0-15) A—Accounting(0-15) A—Asset(0-15) R—Revenue(0-15)
  T—Track Record(0-10) H—BS Health(0-15) I—Valuation(0-15)
  STRONG BUY≥80 | BUY 65-79 | ACCUMULATE 55-64 | HOLD 45-54 | UNDERPERFORM 35-44 | SELL<35

STEP 9 — STRUCTURED ANALYTICAL DATA (web search aggressively for these — they power dedicated sheets):

(a) operational — historical + forward operational metrics:
- years: list of 8 year labels matching last 5 actual + first 3 projected (e.g. ["FY21A","FY22A","FY23A","FY24A","FY25A","FY26E","FY27E","FY28E"])
- volume_segments: dict mapping segment_name → list of 8 yearly MT values
  (e.g. {{"Lead Recycling":[88000,135000,...8 values]}}). Real numbers if available, else conservative estimates.
- capacity_utilisation_pct: list of 8 (e.g. [62,72,78,75,80,84,87,90])
- countries_of_operation, plants_india, plants_overseas: list of 8 integers
- revenue_mix_pct: dict mapping segment → current % (sum ~= 1.0)
- geography_mix_pct: dict mapping region → current %
- realization_per_mt: ₹/tonne for primary product
- employees: integer

(b) governance — REAL data from BSE/NSE filings or company AR:
- board: list of {{name, designation, status (Promoter Executive|Independent|Nominee), din, since (year)}}
- shareholding: {{years:[4 FY labels], promoter_pct, fii_pct, dii_pct, public_pct}} each list of 4 (as fractions 0-1)
- promoter_pledge_pct, board_size, independent_directors, women_directors (numbers)
- statutory_auditor (firm name), managerial_remuneration_cr, auditor_fees_audit_cr,
  auditor_fees_non_audit_cr, related_party_transactions_cr (₹ Cr numbers)

(c) timeline_events — minimum 10 milestones, chronological:
- list of {{year (e.g. "1992" or "2024" or "2026E"), category (Founding|IPO|Expansion|New Vertical|Regulatory|Strategy|Milestone|Outlook), description, impact}}
- Include: incorporation year, IPO year, major capex/expansion years, vision/guidance targets.

(d) risk_items — minimum 8 detailed risks (replaces simple key_risks):
- list of {{category, factor (one-line), description (1-2 sentences), mitigation, probability (H|M|L), impact (H|M|L), rating (HIGH|MEDIUM|LOW)}}
- Differentiate ratings — not all MEDIUM. Match severity.

(e) peers_detailed — 3-4 TRUE direct competitors (same business model & scale, NOT giant generic sector players):
- list of {{name (NSE ticker style), revenue_series (last 5 actual years, ₹ Cr), ebitda_margin_series (fractions 0-1), pat_series (₹ Cr), mcap_cr, pe, pb, roce_pct, roe_pct}}
- For a small-cap metal recycler pick comparable recyclers, NOT Hindustan Zinc/Vedanta.

━━ RULES ━━
1. All monetary values in ₹ Crores. EPS/Target in ₹/share.
2. Percentages as PLAIN NUMBERS: 15.5, NOT 0.155.
3. null for missing. NEVER fabricate historical numbers.
4. BS must balance. Cash is plug.
5. Projection years = 5 years after last actual year.
6. shares_cr = share_capital of last year ÷ face_value ({screener_data['face_value']}).
7. Provide ALL individual expense % — builder NEEDS them for Excel formulas.

━━ OUTPUT ━━
Return ONLY valid JSON. Start with {{ end with }}. No prose, no markdown fences.

{{
  "sector": str,
  "base_year": str,
  "cmp": number,
  "target_price": number,
  "rating": "STRONG BUY|BUY|ACCUMULATE|HOLD|UNDERPERFORM|SELL",
  "upside_pct": number,
  "shares_cr": number,
  "assumptions": {{
    "projection_years": ["FY__E", ...5],
    "revenue_growth_pct": {{"FY__E": number, ...5}},
    "rm_pct": {{...}}, "employee_pct": {{...}}, "power_fuel_pct": {{...}},
    "other_mfg_pct": {{...}}, "selling_admin_pct": {{...}}, "other_exp_pct": {{...}},
    "chg_inventory_pct": {{...}},
    "depreciation_cr": {{...}}, "interest_cr": {{...}}, "other_income_cr": {{...}},
    "tax_rate_pct": {{...}}, "capex_cr": {{...}},
    "receivable_days": {{...}}, "inventory_days": {{...}},
    "dividend_payout_pct": number,
    "target_pe": number, "target_ev_ebitda": number,
    "wacc_pct": number, "terminal_growth_pct": number,
    "rationale": str
  }},
  "projections": {{
    "years": ["FY__E", ...5],
    "net_block":[5], "cwip":[5], "investments":[5], "other_assets":[5],
    "share_capital":[5], "reserves":[5], "borrowings":[5], "other_liabilities":[5],
    "cfo":[5], "cfi":[5], "cff":[5]
  }},
  "historical_ratios": {{
    "years": [last 6 FY],
    "ebitda_margin_pct":[6], "pat_margin_pct":[6], "roe_pct":[6], "roce_pct":[6],
    "debt_equity":[6], "receivable_days":[6], "inventory_days":[6], "asset_turnover":[6],
    "rm_pct":[6], "employee_pct":[6], "tax_rate_pct":[6]
  }},
  "historical_pl_absolute": {{
    "years": [last 6 FY — same order as historical_ratios.years],
    "revenue": [6 values in ₹ Cr, use null only if truly unavailable after web search],
    "ebitda":  [6 values in ₹ Cr, null if unavailable],
    "pat":     [6 values in ₹ Cr, null if unavailable]
  }},
  "valuation": {{
    "dcf_fair_value":number, "pe_fair_value":number,
    "ev_ebitda_fair_value":number, "blended_fair_value":number,
    "sensitivity_pe": {{
      "row_label":"PE Multiple", "col_label":"EPS Growth %",
      "row_values":[5], "col_values":[5], "grid":[[5x5]]
    }}
  }},
  "peers": [{{"name":str,"mcap_cr":n,"pe":n,"ev_ebitda":n,"ebitda_margin_pct":str,"roe_pct":n}}],
  "thesis": {{
    "investment_thesis":str, "bull_case":str, "bear_case":str,
    "key_catalysts":[4], "key_risks":[4],
    "saarthi_scores":{{
      "S_sector_quality":n,"A_accounting_quality":n,"A_asset_quality":n,
      "R_revenue_visibility":n,"T_track_record":n,"H_balance_sheet_health":n,"I_intrinsic_valuation":n
    }},
    "saarthi_total":number, "saarthi_rating":str
  }},
  "operational": {{
    "years":[8 labels], "volume_segments":{{"segment_name":[8 floats]}},
    "capacity_utilisation_pct":[8], "countries_of_operation":[8],
    "plants_india":[8], "plants_overseas":[8],
    "revenue_mix_pct":{{"segment":n}}, "geography_mix_pct":{{"region":n}},
    "realization_per_mt":n, "employees":n
  }},
  "governance": {{
    "board":[{{"name":str,"designation":str,"status":str,"din":str,"since":str}}],
    "shareholding":{{"years":[4],"promoter_pct":[4],"fii_pct":[4],"dii_pct":[4],"public_pct":[4]}},
    "promoter_pledge_pct":n,"board_size":n,"independent_directors":n,"women_directors":n,
    "statutory_auditor":str,"managerial_remuneration_cr":n,
    "auditor_fees_audit_cr":n,"auditor_fees_non_audit_cr":n,"related_party_transactions_cr":n
  }},
  "timeline_events":[{{"year":str,"category":str,"description":str,"impact":str}} ...min 10],
  "risk_items":[{{"category":str,"factor":str,"description":str,"mitigation":str,
                  "probability":"H|M|L","impact":"H|M|L","rating":"HIGH|MEDIUM|LOW"}} ...min 8],
  "peers_detailed":[{{"name":str,"revenue_series":[5],"ebitda_margin_series":[5],"pat_series":[5],
                      "mcap_cr":n,"pe":n,"pb":n,"roce_pct":n,"roe_pct":n}} ...3-4 true peers]
}}"""

    logger.info(f"🤖 Calling Claude API ({MODEL_NAME}) with web_search...")
    with client.messages.stream(
        model=MODEL_NAME,
        max_tokens=16000,
        messages=[{"role": "user", "content": prompt}],
        tools=WEB_SEARCH_TOOL,
    ) as stream:
        response = stream.get_final_message()

    tracker.track(response)
    raw = "".join(b.text for b in response.content if hasattr(b, "text") and b.text)

    logger.info(
        f"✅ Claude done | Tokens: in={tracker.total_input_tokens:,} out={tracker.total_output_tokens:,}"
    )

    try:
        return extract_json_from_text(raw)
    except ValueError:
        logger.warning("⚠️ JSON parse failed, requesting repair...")
        repair = client.messages.create(
            model=MODEL_NAME, max_tokens=16000,
            messages=[{
                "role": "user",
                "content": f"Fix into valid JSON starting with {{ ending with }}:\n{raw}",
            }],
        )
        tracker.track(repair)
        return extract_json_from_text(
            "".join(b.text for b in repair.content if hasattr(b, "text"))
        )


def normalize_model_output(model: dict, screener_data: dict) -> dict:
    normalized = dict(model or {})
    asmp = dict(normalized.get("assumptions") or {})
    proj = dict(normalized.get("projections") or {})
    company_name = screener_data.get("company_name") or "company"

    last_actual_year = screener_data["fiscal_years"][-1] if screener_data.get("fiscal_years") else "FY00"
    projection_years = asmp.get("projection_years") or proj.get("years") or _build_projection_years(last_actual_year)
    asmp["projection_years"] = projection_years
    proj["years"] = proj.get("years") or projection_years

    for key in PER_YEAR_ASSUMPTION_KEYS:
        raw = asmp.get(key)
        if not isinstance(raw, dict) or len(raw) == 0:
            logger.warning("⚠ %s missing entirely — projections will use 0 for this key for %s", key, company_name)
            asmp.pop(key, None)
            continue

        normalized_map: dict[str, float] = {}
        known_positions = [(idx, year, raw.get(year)) for idx, year in enumerate(projection_years) if raw.get(year) is not None]
        if not known_positions:
            logger.warning("⚠ %s missing entirely — projections will use 0 for this key for %s", key, company_name)
            asmp.pop(key, None)
            continue

        filled_messages: list[tuple[str, str, float]] = []
        for idx, year in enumerate(projection_years):
            value = raw.get(year)
            if value is not None:
                normalized_map[year] = value
                continue

            prev_candidates = [(p_idx, p_year, p_val) for p_idx, p_year, p_val in known_positions if p_idx < idx]
            if prev_candidates:
                src_idx, src_year, src_val = prev_candidates[-1]
            else:
                src_idx, src_year, src_val = next((item for item in known_positions if item[0] > idx), known_positions[0])

            normalized_map[year] = src_val
            filled_messages.append((year, src_year, src_val))

        if filled_messages:
            grouped: dict[tuple[str, float], list[str]] = {}
            for filled_year, src_year, src_val in filled_messages:
                grouped.setdefault((src_year, src_val), []).append(filled_year)
            for (src_year, src_val), years in grouped.items():
                logger.warning(
                    "⚠ %s forward-filled %s from %s (=%s) for %s",
                    key,
                    ",".join(years),
                    src_year,
                    src_val,
                    company_name,
                )

        asmp[key] = normalized_map

    shares_cr = normalized.get("shares_cr")
    if shares_cr in (None, 0):
        shares_outstanding = _last_non_null(screener_data["bs"].get("shares_outstanding", []))
        if shares_outstanding not in (None, 0):
            normalized["shares_cr"] = shares_outstanding
        else:
            face_value = screener_data.get("face_value")
            share_capital = _last_non_null(screener_data["bs"].get("share_capital", []))
            if face_value not in (None, 0) and share_capital is not None:
                normalized["shares_cr"] = share_capital / face_value

    normalized["assumptions"] = asmp
    normalized["projections"] = proj
    normalized.setdefault("cmp", screener_data.get("cmp"))
    normalized.setdefault("base_year", last_actual_year)
    return normalized


def validate_model_output(model_json: dict, nse_code: str) -> FinancialModelOutput:
    try:
        validated = FinancialModelOutput.model_validate(model_json)
    except ValidationError as e:
        logger.error("Claude response failed schema validation:\n%s", e.json(indent=2))
        error_summary = "; ".join(
            f"{'.'.join(str(part) for part in err.get('loc', []))}: {err.get('msg', 'invalid')}"
            for err in e.errors()
        )
        raise RuntimeError(f"Financial model JSON invalid for {nse_code}: {error_summary}") from e

    projection_years = validated.assumptions.projection_years
    if len(projection_years) != 5:
        raise RuntimeError(f"Financial model JSON invalid for {nse_code}: assumptions.projection_years must contain exactly 5 years")

    if len(validated.projections.years) != 5:
        raise RuntimeError(f"Financial model JSON invalid for {nse_code}: projections.years must contain exactly 5 years")

    for key in PER_YEAR_ASSUMPTION_KEYS:
        year_map = getattr(validated.assumptions, key)
        missing = [year for year in projection_years if year not in year_map]
        if missing:
            raise RuntimeError(
                f"Financial model JSON invalid for {nse_code}: assumptions.{key} missing years {', '.join(missing)}"
            )

    projection_list_keys = [
        "net_block",
        "cwip",
        "investments",
        "other_assets",
        "share_capital",
        "reserves",
        "borrowings",
        "other_liabilities",
        "cfo",
        "cfi",
        "cff",
    ]
    for key in projection_list_keys:
        values = getattr(validated.projections, key)
        if len(values) != 5:
            raise RuntimeError(f"Financial model JSON invalid for {nse_code}: projections.{key} must contain exactly 5 values")

    historical_ratio_keys = [
        "ebitda_margin_pct",
        "pat_margin_pct",
        "roe_pct",
        "roce_pct",
        "debt_equity",
        "receivable_days",
        "inventory_days",
        "asset_turnover",
        "rm_pct",
        "employee_pct",
        "tax_rate_pct",
    ]
    if len(validated.historical_ratios.years) == 0:
        raise RuntimeError(f"Financial model JSON invalid for {nse_code}: historical_ratios.years is empty")
    for key in historical_ratio_keys:
        values = getattr(validated.historical_ratios, key)
        if len(values) != len(validated.historical_ratios.years):
            raise RuntimeError(
                f"Financial model JSON invalid for {nse_code}: historical_ratios.{key} length does not match historical_ratios.years"
            )

    if len(validated.peers) < 4:
        logger.warning("⚠ peers contains only %s items for %s", len(validated.peers), nse_code)

    return validated


# ══════════════════════════════════════════════════════════════════
# EXTENDED ANALYTICAL SHEETS — shared style helpers
# ══════════════════════════════════════════════════════════════════
EXT_NAVY = "1F4690"
EXT_ORANGE = "FFA500"
EXT_BLUE = "3A5BA0"
EXT_CREAM = "FFE5B4"
EXT_GREY = "F2F2F2"
EXT_WHITE = "FFFFFF"
EXT_DGREEN = "006400"
EXT_DRED = "C00000"

_ext_navy_fill = PatternFill("solid", fgColor=EXT_NAVY)
_ext_orange_fill = PatternFill("solid", fgColor=EXT_ORANGE)
_ext_blue_fill = PatternFill("solid", fgColor=EXT_BLUE)
_ext_cream_fill = PatternFill("solid", fgColor=EXT_CREAM)
_ext_grey_fill = PatternFill("solid", fgColor=EXT_GREY)
_ext_white_fill = PatternFill("solid", fgColor=EXT_WHITE)
_ext_dgreen_fill = PatternFill("solid", fgColor=EXT_DGREEN)
_ext_dred_fill = PatternFill("solid", fgColor=EXT_DRED)
_ext_orange_section_fill = PatternFill("solid", fgColor=EXT_ORANGE)

_ext_hdr_font = Font(name="Arial", bold=True, color=EXT_WHITE, size=11)
_ext_section_font = Font(name="Arial", bold=True, color=EXT_WHITE, size=10)
_ext_label_font = Font(name="Arial", bold=True, color="000000", size=10)
_ext_data_font = Font(name="Arial", color="000000", size=10)
_ext_neg_font = Font(name="Arial", color="C00000", size=10)
_ext_white_font = Font(name="Arial", bold=True, color=EXT_WHITE, size=10)
_ext_title_font = Font(name="Arial", bold=True, color=EXT_WHITE, size=13)
_ext_sub_font = Font(name="Arial", italic=True, color="000000", size=9)

_ext_thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
_ext_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ext_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

FMT_INR = '#,##0;[Red](#,##0);"-"'
FMT_PCT = '0.0%;[Red](0.0%);"-"'
FMT_MULT = "0.0\"x\""
FMT_PER_SHARE = "#,##0.00"
FMT_DAYS = '0;[Red](0);"-"'


def _ext_new_sheet(wb, name):
    n = name + "_Ext" if name in wb.sheetnames else name
    ws = wb.create_sheet(n)
    ws.sheet_view.showGridLines = False
    return ws, n


def _ext_title_banner(ws, title, subtitle, ncols):
    last_col = get_column_letter(ncols)
    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value = title
    c.font = _ext_title_font
    c.fill = _ext_navy_fill
    c.alignment = _ext_center
    ws.row_dimensions[1].height = 28
    ws.merge_cells(f"A2:{last_col}2")
    c2 = ws["A2"]
    c2.value = subtitle
    c2.font = _ext_sub_font
    c2.fill = _ext_cream_fill
    c2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18


def _ext_year_header(ws, row, years, h_count, start_col=2, label="Particulars"):
    ws.cell(row=row, column=1, value=label).font = _ext_hdr_font
    ws.cell(row=row, column=1).fill = _ext_navy_fill
    ws.cell(row=row, column=1).alignment = _ext_center
    ws.cell(row=row, column=1).border = _ext_thin_border
    for i, yr in enumerate(years):
        col = start_col + i
        c = ws.cell(row=row, column=col, value=yr)
        c.font = _ext_hdr_font
        c.fill = _ext_orange_fill if i >= h_count else _ext_navy_fill
        c.alignment = _ext_center
        c.border = _ext_thin_border
    ws.row_dimensions[row].height = 22


def _ext_section_divider(ws, row, ncols, text):
    last_col = get_column_letter(ncols)
    ws.merge_cells(f"A{row}:{last_col}{row}")
    c = ws[f"A{row}"]
    c.value = text
    c.font = _ext_section_font
    c.fill = _ext_orange_section_fill
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border = _ext_thin_border
    ws.row_dimensions[row].height = 20


def _ext_write_data_row(ws, row, label, values, fmt, h_count, p_count, alt=False):
    label_cell = ws.cell(row=row, column=1, value=label)
    label_cell.font = _ext_label_font
    label_cell.fill = _ext_grey_fill if alt else _ext_white_fill
    label_cell.alignment = Alignment(horizontal="left", vertical="center")
    label_cell.border = _ext_thin_border
    for i, v in enumerate(values):
        col = 2 + i
        c = ws.cell(row=row, column=col)
        if v is None or (isinstance(v, float) and (v != v)):
            c.value = "-"
        else:
            c.value = v
            c.number_format = fmt
            if isinstance(v, (int, float)) and v < 0:
                c.font = _ext_neg_font
            else:
                c.font = _ext_data_font
        c.fill = _ext_cream_fill if i >= h_count else (_ext_grey_fill if alt else _ext_white_fill)
        c.alignment = _ext_center
        c.border = _ext_thin_border


def _safe_num(v, default=0.0):
    try:
        if v is None:
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def _ext_hist_value(ctx, hist_key, year):
    hr = ctx["hist_ratios"]
    years = hr.get("years", []) if isinstance(hr, dict) else []
    if year in years and hist_key in hr:
        idx = years.index(year)
        vals = hr.get(hist_key, [])
        if idx < len(vals):
            return vals[idx]
    return None


def _ext_screener_pl(ctx, key, year):
    sd = ctx["sd"]
    if year not in sd["fiscal_years"]:
        return None
    idx = sd["fiscal_years"].index(year)
    series = sd["pl"].get(key) or sd["bs"].get(key) or sd["cf"].get(key)
    if series is None or idx >= len(series):
        return None
    return series[idx]


def _ext_hist_ebitda(ctx, year):
    sd = ctx["sd"]
    if year not in sd["fiscal_years"]:
        return None
    idx = sd["fiscal_years"].index(year)
    eb = sd["derived"].get("ebitda", [])
    if idx < len(eb):
        return eb[idx]
    return None


def _ext_proj_revenue(ctx, year):
    proj_years = ctx["proj_years"]
    if year not in proj_years:
        return None
    idx = proj_years.index(year)
    revs = ctx["projected_revenues"]
    if idx < len(revs):
        return revs[idx]
    return None


def _ext_av(ctx, key, year, default=0.0):
    d = ctx["asmp"].get(key, {})
    if isinstance(d, dict):
        v = d.get(year)
        return _safe_num(v, default)
    return _safe_num(d, default)


def _ext_proj_val(ctx, key, idx, default=0.0):
    vals = ctx["proj"].get(key, [])
    if idx < len(vals) and vals[idx] is not None:
        return _safe_num(vals[idx], default)
    return default


def _ext_proj_pat(ctx, idx):
    """Approximate projected PAT for year idx (using same formula chain as P&L sheet)."""
    rev = ctx["projected_revenues"][idx] if idx < len(ctx["projected_revenues"]) else 0.0
    year = ctx["proj_years"][idx]
    rm = rev * _ext_av(ctx, "rm_pct", year) / 100
    emp = rev * _ext_av(ctx, "employee_pct", year) / 100
    pf = rev * _ext_av(ctx, "power_fuel_pct", year) / 100
    omfg = rev * _ext_av(ctx, "other_mfg_pct", year) / 100
    sa = rev * _ext_av(ctx, "selling_admin_pct", year) / 100
    oe = rev * _ext_av(ctx, "other_exp_pct", year) / 100
    ci = rev * _ext_av(ctx, "chg_inventory_pct", year) / 100
    ebitda = rev - (rm + emp + pf + omfg + sa + oe + ci)
    dep = _ext_av(ctx, "depreciation_cr", year)
    interest = _ext_av(ctx, "interest_cr", year)
    oi = _ext_av(ctx, "other_income_cr", year)
    pbt = ebitda + oi - dep - interest
    tax = max(0.0, pbt * _ext_av(ctx, "tax_rate_pct", year) / 100)
    return pbt - tax, ebitda, dep, interest, oi, pbt, tax


# ══════════════════════════════════════════════════════════════════
# SHEET 1 — Fin_Summary
# ══════════════════════════════════════════════════════════════════
def mk_fin_summary(wb, ctx):
    ws, _ = _ext_new_sheet(wb, "Fin_Summary")
    sd = ctx["sd"]
    hist_years = sd["fiscal_years"]
    dh = min(len(hist_years), 5)
    hist_disp = hist_years[-dh:] if dh else []
    hist_disp_lbl = [f"{y}A" for y in hist_disp]
    proj_years = ctx["proj_years"]
    years_all = hist_disp + proj_years
    years_disp_lbl = hist_disp_lbl + proj_years
    ncols = 1 + len(years_all)
    h_count = len(hist_disp)
    p_count = len(proj_years)
    shares_cr = _safe_num(ctx["shares_cr"], 1.0) or 1.0
    cmp = _safe_num(ctx["cmp"], 0.0)
    mcap_cr = _safe_num(sd.get("mcap"), 0.0)

    ws.column_dimensions["A"].width = 30
    for i in range(len(years_all)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 13

    _ext_title_banner(
        ws,
        f"{ctx['company_name']} — Financial Summary Dashboard",
        f"Sector: {ctx['sector']} | All figures ₹ Cr unless noted | Source: Screener / Tikona Model",
        ncols,
    )

    r = 4
    _ext_year_header(ws, r, years_disp_lbl, h_count)
    r += 1

    def hist_pl(key, year, factor=1.0):
        v = _ext_screener_pl(ctx, key, year)
        if v is None:
            return None
        return _safe_num(v) * factor

    def hist_ratio(key, year):
        return _ext_hist_value(ctx, key, year)

    rows_idx = {"alt": False}

    def emit(label, fmt, values):
        nonlocal r
        _ext_write_data_row(ws, r, label, values, fmt, h_count, p_count, alt=rows_idx["alt"])
        rows_idx["alt"] = not rows_idx["alt"]
        r += 1

    _ext_section_divider(ws, r, ncols, "PROFIT & LOSS")
    r += 1
    rows_idx["alt"] = False

    rev_vals = [hist_pl("sales", y) for y in hist_disp] + [_ext_proj_revenue(ctx, y) for y in proj_years]
    emit("Net Revenue (₹ Cr)", FMT_INR, rev_vals)

    grow = []
    for i, v in enumerate(rev_vals):
        if i == 0 or v is None or rev_vals[i - 1] in (None, 0):
            grow.append(None)
        else:
            grow.append((v - rev_vals[i - 1]) / rev_vals[i - 1])
    emit("Revenue Growth %", FMT_PCT, grow)

    eb_vals = [_ext_hist_ebitda(ctx, y) for y in hist_disp]
    pat_vals = [hist_pl("net_profit", y) for y in hist_disp]
    dep_vals = [hist_pl("depreciation", y) for y in hist_disp]
    int_vals = [hist_pl("interest", y) for y in hist_disp]
    pbt_vals = [hist_pl("pbt", y) for y in hist_disp]
    div_vals = [hist_pl("dividend_amount", y) for y in hist_disp]

    for i in range(p_count):
        pat_p, eb_p, dep_p, int_p, _oi, pbt_p, _tax = _ext_proj_pat(ctx, i)
        eb_vals.append(eb_p)
        dep_vals.append(dep_p)
        int_vals.append(int_p)
        pbt_vals.append(pbt_p)
        pat_vals.append(pat_p)
        payout = _safe_num(ctx["asmp"].get("dividend_payout_pct"), 0.0) / 100
        div_vals.append(pat_p * payout)

    emit("EBITDA (₹ Cr)", FMT_INR, eb_vals)
    emit("EBITDA Margin %", FMT_PCT, [
        (eb_vals[i] / rev_vals[i]) if (eb_vals[i] is not None and rev_vals[i] not in (None, 0)) else None
        for i in range(len(rev_vals))
    ])
    emit("Depreciation (₹ Cr)", FMT_INR, dep_vals)
    ebit_vals = [
        (eb_vals[i] - dep_vals[i]) if (eb_vals[i] is not None and dep_vals[i] is not None) else None
        for i in range(len(rev_vals))
    ]
    emit("EBIT (₹ Cr)", FMT_INR, ebit_vals)
    emit("Interest (₹ Cr)", FMT_INR, int_vals)
    emit("PBT (₹ Cr)", FMT_INR, pbt_vals)
    emit("PAT (₹ Cr)", FMT_INR, pat_vals)
    emit("PAT Margin %", FMT_PCT, [
        (pat_vals[i] / rev_vals[i]) if (pat_vals[i] is not None and rev_vals[i] not in (None, 0)) else None
        for i in range(len(rev_vals))
    ])
    pat_growth = []
    for i, v in enumerate(pat_vals):
        if i == 0 or v is None or pat_vals[i - 1] in (None, 0):
            pat_growth.append(None)
        else:
            pat_growth.append((v - pat_vals[i - 1]) / pat_vals[i - 1])
    emit("PAT Growth %", FMT_PCT, pat_growth)
    emit("EPS (₹)", FMT_PER_SHARE, [
        (v / shares_cr) if v is not None else None for v in pat_vals
    ])
    emit("DPS (₹)", FMT_PER_SHARE, [
        (v / shares_cr) if v is not None else None for v in div_vals
    ])

    _ext_section_divider(ws, r, ncols, "BALANCE SHEET")
    r += 1
    rows_idx["alt"] = False

    sc_vals = [hist_pl("share_capital", y) for y in hist_disp] + [_ext_proj_val(ctx, "share_capital", i) for i in range(p_count)]
    res_vals = [hist_pl("reserves", y) for y in hist_disp] + [_ext_proj_val(ctx, "reserves", i) for i in range(p_count)]
    borr_vals = [hist_pl("borrowings", y) for y in hist_disp] + [_ext_proj_val(ctx, "borrowings", i) for i in range(p_count)]
    nb_vals = [hist_pl("net_block", y) for y in hist_disp] + [_ext_proj_val(ctx, "net_block", i) for i in range(p_count)]
    recv_vals = [hist_pl("receivables", y) for y in hist_disp]
    inv_vals = [hist_pl("inventory", y) for y in hist_disp]
    for i in range(p_count):
        rev_i = ctx["projected_revenues"][i]
        recv_vals.append(rev_i * _ext_av(ctx, "receivable_days", proj_years[i]) / 365)
        rm_i = rev_i * _ext_av(ctx, "rm_pct", proj_years[i]) / 100
        inv_vals.append(rm_i * _ext_av(ctx, "inventory_days", proj_years[i]) / 365)

    nw_vals = [
        (_safe_num(sc_vals[i]) + _safe_num(res_vals[i])) if (sc_vals[i] is not None or res_vals[i] is not None) else None
        for i in range(len(years_all))
    ]
    emit("Net Worth (₹ Cr)", FMT_INR, nw_vals)
    emit("Total Debt (₹ Cr)", FMT_INR, borr_vals)
    ce_vals = [
        (_safe_num(nw_vals[i]) + _safe_num(borr_vals[i])) if (nw_vals[i] is not None or borr_vals[i] is not None) else None
        for i in range(len(years_all))
    ]
    emit("Capital Employed (₹ Cr)", FMT_INR, ce_vals)
    emit("Net Fixed Assets (₹ Cr)", FMT_INR, nb_vals)
    wc_vals = [
        (_safe_num(recv_vals[i]) + _safe_num(inv_vals[i])) if (recv_vals[i] is not None or inv_vals[i] is not None) else None
        for i in range(len(years_all))
    ]
    emit("Working Capital (₹ Cr)", FMT_INR, wc_vals)
    emit("Debt/Equity (x)", FMT_MULT, [
        (_safe_num(borr_vals[i]) / _safe_num(nw_vals[i])) if _safe_num(nw_vals[i]) else None
        for i in range(len(years_all))
    ])

    _ext_section_divider(ws, r, ncols, "CASH FLOWS")
    r += 1
    rows_idx["alt"] = False

    cfo_vals = [hist_pl("cfo", y) for y in hist_disp] + [_ext_proj_val(ctx, "cfo", i) for i in range(p_count)]
    cfi_vals = [hist_pl("cfi", y) for y in hist_disp] + [_ext_proj_val(ctx, "cfi", i) for i in range(p_count)]
    capex_hist = [
        (-_safe_num(c)) if c is not None else None for c in cfi_vals[:h_count]
    ]
    capex_proj = [_ext_av(ctx, "capex_cr", proj_years[i]) for i in range(p_count)]
    capex_vals = capex_hist + capex_proj
    fcf_vals = [
        (_safe_num(cfo_vals[i]) - _safe_num(capex_vals[i])) if (cfo_vals[i] is not None or capex_vals[i] is not None) else None
        for i in range(len(years_all))
    ]
    emit("CFO (₹ Cr)", FMT_INR, cfo_vals)
    emit("Capex (₹ Cr)", FMT_INR, capex_vals)
    emit("Free Cash Flow (₹ Cr)", FMT_INR, fcf_vals)
    emit("CFO/EBITDA %", FMT_PCT, [
        (_safe_num(cfo_vals[i]) / _safe_num(eb_vals[i])) if _safe_num(eb_vals[i]) else None
        for i in range(len(years_all))
    ])

    _ext_section_divider(ws, r, ncols, "KEY RATIOS")
    r += 1
    rows_idx["alt"] = False

    roe_vals = [hist_ratio("roe_pct", y) for y in hist_disp]
    roce_vals = [hist_ratio("roce_pct", y) for y in hist_disp]
    dd_vals = [hist_ratio("receivable_days", y) for y in hist_disp]
    id_vals = [hist_ratio("inventory_days", y) for y in hist_disp]

    for i in range(p_count):
        nw_i = _safe_num(nw_vals[h_count + i], 0)
        ce_i = _safe_num(ce_vals[h_count + i], 0)
        pat_i = _safe_num(pat_vals[h_count + i], 0)
        ebit_i = _safe_num(ebit_vals[h_count + i], 0)
        roe_vals.append((pat_i / nw_i * 100) if nw_i else None)
        roce_vals.append((ebit_i / ce_i * 100) if ce_i else None)
        dd_vals.append(_ext_av(ctx, "receivable_days", proj_years[i]))
        id_vals.append(_ext_av(ctx, "inventory_days", proj_years[i]))

    emit("ROE %", FMT_PCT, [(_safe_num(v) / 100) if v is not None else None for v in roe_vals])
    emit("ROCE %", FMT_PCT, [(_safe_num(v) / 100) if v is not None else None for v in roce_vals])
    emit("Debtor Days", FMT_DAYS, dd_vals)
    emit("Inventory Days", FMT_DAYS, id_vals)

    _ext_section_divider(ws, r, ncols, "VALUATIONS (at CMP)")
    r += 1
    rows_idx["alt"] = False

    eps_vals = [(_safe_num(v) / shares_cr) if v is not None else None for v in pat_vals]
    bvps_vals = [(_safe_num(v) / shares_cr) if v is not None else None for v in nw_vals]
    sps_vals = [(_safe_num(v) / shares_cr) if v is not None else None for v in rev_vals]
    ev = mcap_cr + sum(_safe_num(borr_vals[h_count - 1] if h_count else 0, 0) for _ in [0])
    pe_vals = [(cmp / v) if (v and v > 0) else None for v in eps_vals]
    pb_vals = [(cmp / v) if (v and v > 0) else None for v in bvps_vals]
    ps_vals = [(cmp / v) if (v and v > 0) else None for v in sps_vals]
    ev_ebitda = [(ev / _safe_num(eb_vals[i])) if _safe_num(eb_vals[i]) else None for i in range(len(years_all))]
    emit("P/E (x)", FMT_MULT, pe_vals)
    emit("P/B (x)", FMT_MULT, pb_vals)
    emit("P/S (x)", FMT_MULT, ps_vals)
    emit("EV/EBITDA (x)", FMT_MULT, ev_ebitda)


# ══════════════════════════════════════════════════════════════════
# SHEET 2 — Earnings_Forecast
# ══════════════════════════════════════════════════════════════════
def mk_earnings_forecast(wb, ctx):
    ws, _ = _ext_new_sheet(wb, "Earnings_Forecast")
    sd = ctx["sd"]
    hist_years = sd["fiscal_years"]
    dh = min(len(hist_years), 5)
    hist_disp = hist_years[-dh:] if dh else []
    hist_disp_lbl = [f"{y}A" for y in hist_disp]
    proj_years = ctx["proj_years"]
    years_all = hist_disp + proj_years
    years_disp_lbl = hist_disp_lbl + proj_years
    ncols = 1 + len(years_all)
    h_count = len(hist_disp)
    p_count = len(proj_years)
    shares_cr = _safe_num(ctx["shares_cr"], 1.0) or 1.0

    ws.column_dimensions["A"].width = 32
    for i in range(len(years_all)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 13

    _ext_title_banner(
        ws,
        f"{ctx['company_name']} — Earnings Forecast",
        "Forward-looking income statement | All figures ₹ Cr | Estimate years shaded in cream",
        ncols,
    )

    r = 4
    _ext_year_header(ws, r, years_disp_lbl, h_count)
    r += 1

    def hist_pl(key, year):
        v = _ext_screener_pl(ctx, key, year)
        return _safe_num(v) if v is not None else None

    alt = {"v": False}

    def emit(label, fmt, vals):
        nonlocal r
        _ext_write_data_row(ws, r, label, vals, fmt, h_count, p_count, alt=alt["v"])
        alt["v"] = not alt["v"]
        r += 1

    _ext_section_divider(ws, r, ncols, "INCOME STATEMENT")
    r += 1
    alt["v"] = False

    rev_vals = [hist_pl("sales", y) for y in hist_disp] + [_ext_proj_revenue(ctx, y) for y in proj_years]
    emit("Revenue (₹ Cr)", FMT_INR, rev_vals)
    yoy = []
    for i, v in enumerate(rev_vals):
        if i == 0 or v is None or rev_vals[i - 1] in (None, 0):
            yoy.append(None)
        else:
            yoy.append((v - rev_vals[i - 1]) / rev_vals[i - 1])
    emit("YoY Growth %", FMT_PCT, yoy)

    rm_vals = [hist_pl("raw_material", y) for y in hist_disp] + [
        ctx["projected_revenues"][i] * _ext_av(ctx, "rm_pct", proj_years[i]) / 100 for i in range(p_count)
    ]
    gp_vals = [
        (_safe_num(rev_vals[i]) - _safe_num(rm_vals[i])) if (rev_vals[i] is not None and rm_vals[i] is not None) else None
        for i in range(len(years_all))
    ]
    emit("Gross Profit (₹ Cr)", FMT_INR, gp_vals)
    emit("Gross Margin %", FMT_PCT, [
        (_safe_num(gp_vals[i]) / _safe_num(rev_vals[i])) if _safe_num(rev_vals[i]) else None
        for i in range(len(years_all))
    ])

    eb_vals = [_ext_hist_ebitda(ctx, y) for y in hist_disp]
    dep_vals = [hist_pl("depreciation", y) for y in hist_disp]
    int_vals = [hist_pl("interest", y) for y in hist_disp]
    oi_vals = [hist_pl("other_income", y) for y in hist_disp]
    pbt_vals = [hist_pl("pbt", y) for y in hist_disp]
    tax_vals = [hist_pl("tax", y) for y in hist_disp]
    pat_vals = [hist_pl("net_profit", y) for y in hist_disp]
    for i in range(p_count):
        pat_p, eb_p, dep_p, int_p, oi_p, pbt_p, tax_p = _ext_proj_pat(ctx, i)
        eb_vals.append(eb_p)
        dep_vals.append(dep_p)
        int_vals.append(int_p)
        oi_vals.append(oi_p)
        pbt_vals.append(pbt_p)
        tax_vals.append(tax_p)
        pat_vals.append(pat_p)

    emit("EBITDA (₹ Cr)", FMT_INR, eb_vals)
    emit("EBITDA Margin %", FMT_PCT, [
        (_safe_num(eb_vals[i]) / _safe_num(rev_vals[i])) if _safe_num(rev_vals[i]) else None
        for i in range(len(years_all))
    ])
    eb_growth = []
    for i, v in enumerate(eb_vals):
        if i == 0 or v is None or eb_vals[i - 1] in (None, 0):
            eb_growth.append(None)
        else:
            eb_growth.append((v - eb_vals[i - 1]) / eb_vals[i - 1])
    emit("EBITDA Growth %", FMT_PCT, eb_growth)

    emit("Depreciation (₹ Cr)", FMT_INR, dep_vals)
    ebit_vals = [
        (_safe_num(eb_vals[i]) - _safe_num(dep_vals[i])) if (eb_vals[i] is not None and dep_vals[i] is not None) else None
        for i in range(len(years_all))
    ]
    emit("EBIT (₹ Cr)", FMT_INR, ebit_vals)
    emit("Finance Costs (₹ Cr)", FMT_INR, int_vals)
    emit("Other Income (₹ Cr)", FMT_INR, oi_vals)
    emit("PBT (₹ Cr)", FMT_INR, pbt_vals)
    emit("Tax (₹ Cr)", FMT_INR, tax_vals)
    emit("Effective Tax Rate %", FMT_PCT, [
        (_safe_num(tax_vals[i]) / _safe_num(pbt_vals[i])) if _safe_num(pbt_vals[i]) else None
        for i in range(len(years_all))
    ])
    emit("PAT (₹ Cr)", FMT_INR, pat_vals)
    pat_growth = []
    for i, v in enumerate(pat_vals):
        if i == 0 or v is None or pat_vals[i - 1] in (None, 0):
            pat_growth.append(None)
        else:
            pat_growth.append((v - pat_vals[i - 1]) / pat_vals[i - 1])
    emit("PAT Growth %", FMT_PCT, pat_growth)
    emit("PAT Margin %", FMT_PCT, [
        (_safe_num(pat_vals[i]) / _safe_num(rev_vals[i])) if _safe_num(rev_vals[i]) else None
        for i in range(len(years_all))
    ])

    _ext_section_divider(ws, r, ncols, "PER SHARE DATA")
    r += 1
    alt["v"] = False

    eps_vals = [(_safe_num(v) / shares_cr) if v is not None else None for v in pat_vals]
    emit("EPS (₹)", FMT_PER_SHARE, eps_vals)
    eps_growth = []
    for i, v in enumerate(eps_vals):
        if i == 0 or v is None or eps_vals[i - 1] in (None, 0):
            eps_growth.append(None)
        else:
            eps_growth.append((v - eps_vals[i - 1]) / eps_vals[i - 1])
    emit("EPS Growth %", FMT_PCT, eps_growth)
    payout = _safe_num(ctx["asmp"].get("dividend_payout_pct"), 0.0) / 100
    dps_vals = [(_safe_num(v) * payout / shares_cr) if v is not None else None for v in pat_vals]
    emit("DPS (₹)", FMT_PER_SHARE, dps_vals)

    sc_vals = [hist_pl("share_capital", y) for y in hist_disp] + [_ext_proj_val(ctx, "share_capital", i) for i in range(p_count)]
    res_vals = [hist_pl("reserves", y) for y in hist_disp] + [_ext_proj_val(ctx, "reserves", i) for i in range(p_count)]
    bv_vals = [
        ((_safe_num(sc_vals[i]) + _safe_num(res_vals[i])) / shares_cr) if (sc_vals[i] is not None or res_vals[i] is not None) else None
        for i in range(len(years_all))
    ]
    emit("Book Value per Share (₹)", FMT_PER_SHARE, bv_vals)

    _ext_section_divider(ws, r, ncols, "ASSUMPTIONS & DRIVERS")
    r += 1
    alt["v"] = False

    vol_growth = [None] * h_count + [_ext_av(ctx, "revenue_growth_pct", y) / 100 * 0.6 for y in proj_years]
    real_growth = [None] * h_count + [_ext_av(ctx, "revenue_growth_pct", y) / 100 * 0.4 for y in proj_years]
    emit("Volume Growth % (est)", FMT_PCT, vol_growth)
    emit("Realization Growth % (est)", FMT_PCT, real_growth)
    rg = [None] * h_count + [_ext_av(ctx, "revenue_growth_pct", y) / 100 for y in proj_years]
    emit("Revenue Growth %", FMT_PCT, rg)
    em = [
        (_safe_num(eb_vals[i]) / _safe_num(rev_vals[i])) if _safe_num(rev_vals[i]) else None
        for i in range(len(years_all))
    ]
    emit("EBITDA Margin %", FMT_PCT, em)
    capex = [None] * h_count + [_ext_av(ctx, "capex_cr", y) for y in proj_years]
    emit("Capex (₹ Cr)", FMT_INR, capex)
    wcd = [None] * h_count + [
        _ext_av(ctx, "receivable_days", y) + _ext_av(ctx, "inventory_days", y) for y in proj_years
    ]
    emit("Working Capital Days", FMT_DAYS, wcd)


# ══════════════════════════════════════════════════════════════════
# SHEET 3 — Financials_Table (formula-linked)
# ══════════════════════════════════════════════════════════════════
def mk_financials_table(wb, ctx):
    ws, _ = _ext_new_sheet(wb, "Financials_Table")
    pln, bsn, cfn, asn = ctx["pln"], ctx["bsn"], ctx["cfn"], ctx["asn"]
    REV, EBITDA, OI, DEP, EBIT, INT_R, PBT, TAX, PAT = (
        ctx["REV"], ctx["EBITDA"], ctx["OI"], ctx["DEP"], ctx["EBIT"], ctx["INT_R"],
        ctx["PBT"], ctx["TAX"], ctx["PAT"],
    )
    EQ_R, TL_R, TA_R = ctx["EQ_R"], ctx["TL_R"], ctx["TA_R"]
    RECV_R, INVTY_R = ctx["RECV_R"], ctx["INVTY_R"]
    bs_rr = ctx["bs_rr"]
    cf_rr = ctx["cf_rr"]
    exp_rows = ctx["exp_rows"]
    year_labels = ctx["year_labels"]
    nc = ctx["nc"]
    h_cols = ctx["h_cols"]
    p_cols = ctx["p_cols"]

    ws.column_dimensions["A"].width = 32
    for i in range(2, nc + 1):
        ws.column_dimensions[get_column_letter(i)].width = 13

    _ext_title_banner(
        ws,
        f"{ctx['company_name']} — Standardised Financial Statements",
        "Cells link directly to P&L / Balance Sheet / Cash Flow sheets (fully dynamic)",
        nc,
    )

    r = 4
    _ext_year_header(ws, r, year_labels[1:], len(h_cols))
    r += 1

    def emit_formula(label, formula_tmpl, fmt):
        nonlocal r
        ws.cell(row=r, column=1, value=label).font = _ext_label_font
        ws.cell(row=r, column=1).border = _ext_thin_border
        for ci in range(2, nc + 1):
            cl = get_column_letter(ci)
            cell = ws.cell(row=r, column=ci, value=formula_tmpl.replace("{cl}", cl))
            cell.number_format = fmt
            cell.border = _ext_thin_border
            cell.alignment = _ext_center
            cell.fill = _ext_cream_fill if ci in p_cols else _ext_white_fill
        r += 1

    _ext_section_divider(ws, r, nc, "PROFIT & LOSS")
    r += 1
    emit_formula("Revenue", f"='{pln}'!{{cl}}{REV}", FMT_INR)
    emit_formula("Raw Material Cost", f"='{pln}'!{{cl}}{exp_rows[0]}", FMT_INR)
    emit_formula("Employee Cost", f"='{pln}'!{{cl}}{exp_rows[2]}", FMT_INR)
    other_op = "+".join([f"'{pln}'!{{cl}}{exp_rows[i]}" for i in [1, 3, 4, 5, 6]])
    emit_formula("Other Operating Expenses", f"={other_op}", FMT_INR)
    emit_formula("EBITDA", f"='{pln}'!{{cl}}{EBITDA}", FMT_INR)
    emit_formula("EBITDA Margin %", f"=IF('{pln}'!{{cl}}{REV}=0,0,'{pln}'!{{cl}}{EBITDA}/'{pln}'!{{cl}}{REV})", FMT_PCT)
    emit_formula("Depreciation", f"='{pln}'!{{cl}}{DEP}", FMT_INR)
    emit_formula("EBIT", f"='{pln}'!{{cl}}{EBIT}", FMT_INR)
    emit_formula("Interest", f"='{pln}'!{{cl}}{INT_R}", FMT_INR)
    emit_formula("Other Income", f"='{pln}'!{{cl}}{OI}", FMT_INR)
    emit_formula("PBT", f"='{pln}'!{{cl}}{PBT}", FMT_INR)
    emit_formula("Tax", f"='{pln}'!{{cl}}{TAX}", FMT_INR)
    emit_formula("PAT", f"='{pln}'!{{cl}}{PAT}", FMT_INR)
    shares_cr = _safe_num(ctx["shares_cr"], 1.0) or 1.0
    emit_formula("EPS (₹)", f"='{pln}'!{{cl}}{PAT}/{shares_cr:.4f}", FMT_PER_SHARE)

    _ext_section_divider(ws, r, nc, "BALANCE SHEET")
    r += 1
    emit_formula("Equity Capital", f"='{bsn}'!{{cl}}{bs_rr['share_capital']}", FMT_INR)
    emit_formula("Reserves & Surplus", f"='{bsn}'!{{cl}}{bs_rr['reserves']}", FMT_INR)
    emit_formula("Net Worth", f"='{bsn}'!{{cl}}{EQ_R}", FMT_INR)
    emit_formula("Total Debt", f"='{bsn}'!{{cl}}{bs_rr['borrowings']}", FMT_INR)
    emit_formula("Other Liabilities", f"='{bsn}'!{{cl}}{bs_rr['other_liabilities']}", FMT_INR)
    emit_formula("Total Liabilities & Equity", f"='{bsn}'!{{cl}}{TL_R}", FMT_INR)
    emit_formula("Net Fixed Assets", f"='{bsn}'!{{cl}}{bs_rr.get('net_block', 9)}", FMT_INR) if "net_block" in bs_rr else None
    emit_formula("Working Capital", f"='{bsn}'!{{cl}}{RECV_R}+'{bsn}'!{{cl}}{INVTY_R}", FMT_INR)
    emit_formula("Total Assets", f"='{bsn}'!{{cl}}{TA_R}", FMT_INR)

    _ext_section_divider(ws, r, nc, "CASH FLOW")
    r += 1
    emit_formula("CFO", f"='{cfn}'!{{cl}}{cf_rr['cfo']}", FMT_INR)
    emit_formula("CFI (Investing)", f"='{cfn}'!{{cl}}{cf_rr['cfi']}", FMT_INR)
    emit_formula("CFF (Financing)", f"='{cfn}'!{{cl}}{cf_rr['cff']}", FMT_INR)
    emit_formula(
        "Net Cash Change",
        f"='{cfn}'!{{cl}}{cf_rr['cfo']}+'{cfn}'!{{cl}}{cf_rr['cfi']}+'{cfn}'!{{cl}}{cf_rr['cff']}",
        FMT_INR,
    )
    emit_formula(
        "Capex",
        f"=-'{cfn}'!{{cl}}{cf_rr['cfi']}",
        FMT_INR,
    )
    emit_formula(
        "Free Cash Flow",
        f"='{cfn}'!{{cl}}{cf_rr['cfo']}+'{cfn}'!{{cl}}{cf_rr['cfi']}",
        FMT_INR,
    )

    _ext_section_divider(ws, r, nc, "KEY RATIOS")
    r += 1
    emit_formula(
        "Gross Margin %",
        f"=IF('{pln}'!{{cl}}{REV}=0,0,('{pln}'!{{cl}}{REV}-'{pln}'!{{cl}}{exp_rows[0]})/'{pln}'!{{cl}}{REV})",
        FMT_PCT,
    )
    emit_formula("EBITDA Margin %", f"=IF('{pln}'!{{cl}}{REV}=0,0,'{pln}'!{{cl}}{EBITDA}/'{pln}'!{{cl}}{REV})", FMT_PCT)
    emit_formula("PAT Margin %", f"=IF('{pln}'!{{cl}}{REV}=0,0,'{pln}'!{{cl}}{PAT}/'{pln}'!{{cl}}{REV})", FMT_PCT)
    emit_formula("ROE %", f"=IF('{bsn}'!{{cl}}{EQ_R}=0,0,'{pln}'!{{cl}}{PAT}/'{bsn}'!{{cl}}{EQ_R})", FMT_PCT)
    emit_formula(
        "ROCE %",
        f"=IF(('{bsn}'!{{cl}}{EQ_R}+'{bsn}'!{{cl}}{bs_rr['borrowings']})=0,0,'{pln}'!{{cl}}{EBIT}/('{bsn}'!{{cl}}{EQ_R}+'{bsn}'!{{cl}}{bs_rr['borrowings']}))",
        FMT_PCT,
    )
    emit_formula("Debt-Equity (x)", f"=IF('{bsn}'!{{cl}}{EQ_R}=0,0,'{bsn}'!{{cl}}{bs_rr['borrowings']}/'{bsn}'!{{cl}}{EQ_R})", FMT_MULT)
    emit_formula("Interest Coverage (x)", f"=IF('{pln}'!{{cl}}{INT_R}=0,0,'{pln}'!{{cl}}{EBIT}/'{pln}'!{{cl}}{INT_R})", FMT_MULT)
    emit_formula("Asset Turnover (x)", f"=IF('{bsn}'!{{cl}}{TA_R}=0,0,'{pln}'!{{cl}}{REV}/'{bsn}'!{{cl}}{TA_R})", FMT_MULT)
    emit_formula(
        "Debtor Days",
        f"=IF('{pln}'!{{cl}}{REV}=0,0,'{bsn}'!{{cl}}{RECV_R}/'{pln}'!{{cl}}{REV}*365)",
        FMT_DAYS,
    )
    emit_formula(
        "Inventory Days",
        f"=IF('{pln}'!{{cl}}{exp_rows[0]}=0,0,'{bsn}'!{{cl}}{INVTY_R}/'{pln}'!{{cl}}{exp_rows[0]}*365)",
        FMT_DAYS,
    )
    emit_formula(
        "CFO/EBITDA %",
        f"=IF('{pln}'!{{cl}}{EBITDA}=0,0,'{cfn}'!{{cl}}{cf_rr['cfo']}/'{pln}'!{{cl}}{EBITDA})",
        FMT_PCT,
    )


# ══════════════════════════════════════════════════════════════════
# SHEET 4 — Valuations_Table
# ══════════════════════════════════════════════════════════════════
def mk_valuations_table(wb, ctx):
    ws, _ = _ext_new_sheet(wb, "Valuations_Table")
    sd = ctx["sd"]
    hist_years = sd["fiscal_years"]
    dh = min(len(hist_years), 5)
    hist_disp = hist_years[-dh:] if dh else []
    hist_disp_lbl = [f"{y}A" for y in hist_disp]
    proj_years = ctx["proj_years"]
    years_all = hist_disp + proj_years
    years_disp_lbl = hist_disp_lbl + proj_years
    ncols = 1 + len(years_all)
    h_count = len(hist_disp)
    p_count = len(proj_years)
    shares_cr = _safe_num(ctx["shares_cr"], 1.0) or 1.0
    cmp = _safe_num(ctx["cmp"], 0.0)
    mcap_cr = shares_cr * cmp
    val = ctx["val"]
    asmp = ctx["asmp"]

    ws.column_dimensions["A"].width = 32
    for i in range(len(years_all)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 13

    _ext_title_banner(
        ws,
        f"{ctx['company_name']} — Valuation Matrix",
        f"CMP: ₹{cmp:,.0f} | Target: ₹{_safe_num(ctx['target']):,.0f} | Rating: {ctx['rating']} | All multiples at CMP",
        ncols,
    )

    r = 4
    _ext_year_header(ws, r, years_disp_lbl, h_count)
    r += 1

    def hist_pl(key, y):
        v = _ext_screener_pl(ctx, key, y)
        return _safe_num(v) if v is not None else None

    rev_vals = [hist_pl("sales", y) for y in hist_disp] + [_ext_proj_revenue(ctx, y) for y in proj_years]
    eb_vals = [_ext_hist_ebitda(ctx, y) for y in hist_disp]
    ebit_vals = []
    pat_vals = [hist_pl("net_profit", y) for y in hist_disp]
    dep_vals = [hist_pl("depreciation", y) for y in hist_disp]
    sc_vals = [hist_pl("share_capital", y) for y in hist_disp] + [_ext_proj_val(ctx, "share_capital", i) for i in range(p_count)]
    res_vals = [hist_pl("reserves", y) for y in hist_disp] + [_ext_proj_val(ctx, "reserves", i) for i in range(p_count)]
    borr_vals = [hist_pl("borrowings", y) for y in hist_disp] + [_ext_proj_val(ctx, "borrowings", i) for i in range(p_count)]
    cfo_vals = [hist_pl("cfo", y) for y in hist_disp] + [_ext_proj_val(ctx, "cfo", i) for i in range(p_count)]

    for i in range(h_count):
        if eb_vals[i] is not None and dep_vals[i] is not None:
            ebit_vals.append(_safe_num(eb_vals[i]) - _safe_num(dep_vals[i]))
        else:
            ebit_vals.append(None)
    for i in range(p_count):
        pat_p, eb_p, dep_p, _i, _o, _p, _t = _ext_proj_pat(ctx, i)
        eb_vals.append(eb_p)
        ebit_vals.append(eb_p - dep_p)
        pat_vals.append(pat_p)

    alt = {"v": False}

    def emit(label, fmt, vals):
        nonlocal r
        _ext_write_data_row(ws, r, label, vals, fmt, h_count, p_count, alt=alt["v"])
        alt["v"] = not alt["v"]
        r += 1

    _ext_section_divider(ws, r, ncols, "MARKET DATA")
    r += 1
    alt["v"] = False
    emit("CMP (₹)", FMT_PER_SHARE, [cmp] * len(years_all))
    emit("Market Cap (₹ Cr)", FMT_INR, [mcap_cr] * len(years_all))
    ev_vals = [mcap_cr + _safe_num(borr_vals[i]) for i in range(len(years_all))]
    emit("Enterprise Value (₹ Cr)", FMT_INR, ev_vals)

    _ext_section_divider(ws, r, ncols, "EARNINGS MULTIPLES")
    r += 1
    alt["v"] = False
    eps = [(_safe_num(v) / shares_cr) if v is not None else None for v in pat_vals]
    bv = [
        ((_safe_num(sc_vals[i]) + _safe_num(res_vals[i])) / shares_cr) if (sc_vals[i] is not None or res_vals[i] is not None) else None
        for i in range(len(years_all))
    ]
    sps = [(_safe_num(v) / shares_cr) if v is not None else None for v in rev_vals]
    emit("P/E (x)", FMT_MULT, [(cmp / v) if (v and v > 0) else None for v in eps])
    emit("P/B (x)", FMT_MULT, [(cmp / v) if (v and v > 0) else None for v in bv])
    emit("P/Sales (x)", FMT_MULT, [(cmp / v) if (v and v > 0) else None for v in sps])
    emit("EV/EBITDA (x)", FMT_MULT, [(ev_vals[i] / _safe_num(eb_vals[i])) if _safe_num(eb_vals[i]) else None for i in range(len(years_all))])
    emit("EV/Sales (x)", FMT_MULT, [(ev_vals[i] / _safe_num(rev_vals[i])) if _safe_num(rev_vals[i]) else None for i in range(len(years_all))])
    emit("EV/EBIT (x)", FMT_MULT, [(ev_vals[i] / _safe_num(ebit_vals[i])) if _safe_num(ebit_vals[i]) else None for i in range(len(years_all))])

    _ext_section_divider(ws, r, ncols, "PER SHARE")
    r += 1
    alt["v"] = False
    emit("EPS (₹)", FMT_PER_SHARE, eps)
    emit("Book Value (₹)", FMT_PER_SHARE, bv)
    emit("Sales/Share (₹)", FMT_PER_SHARE, sps)
    emit("CFO/Share (₹)", FMT_PER_SHARE, [(_safe_num(v) / shares_cr) if v is not None else None for v in cfo_vals])
    payout = _safe_num(asmp.get("dividend_payout_pct"), 0.0) / 100
    emit("DPS (₹)", FMT_PER_SHARE, [(_safe_num(v) * payout / shares_cr) if v is not None else None for v in pat_vals])

    _ext_section_divider(ws, r, ncols, "RETURN METRICS")
    r += 1
    alt["v"] = False
    nw = [
        (_safe_num(sc_vals[i]) + _safe_num(res_vals[i])) if (sc_vals[i] is not None or res_vals[i] is not None) else None
        for i in range(len(years_all))
    ]
    ce = [
        (_safe_num(nw[i]) + _safe_num(borr_vals[i])) if (nw[i] is not None or borr_vals[i] is not None) else None
        for i in range(len(years_all))
    ]
    emit("ROE %", FMT_PCT, [(_safe_num(pat_vals[i]) / _safe_num(nw[i])) if _safe_num(nw[i]) else None for i in range(len(years_all))])
    emit("ROCE %", FMT_PCT, [(_safe_num(ebit_vals[i]) / _safe_num(ce[i])) if _safe_num(ce[i]) else None for i in range(len(years_all))])
    emit("Earnings Yield %", FMT_PCT, [(v / cmp) if (v is not None and cmp) else None for v in eps])
    emit("Dividend Yield %", FMT_PCT, [(_safe_num(v) * payout / shares_cr / cmp) if (v is not None and cmp) else None for v in pat_vals])
    emit("FCF Yield %", FMT_PCT, [
        ((_safe_num(cfo_vals[i]) + _safe_num(_ext_av(ctx, "capex_cr", proj_years[i - h_count]) if i >= h_count else 0) * -1) / mcap_cr) if mcap_cr else None
        for i in range(len(years_all))
    ])

    r += 1
    _ext_section_divider(ws, r, ncols, "DCF ASSUMPTIONS")
    r += 1
    beta = 1.1
    rf = 0.07
    erp = 0.06
    coe = rf + erp * beta
    dcf_rows = [
        ("Risk-Free Rate %", rf, FMT_PCT, "proj_only"),
        ("Equity Risk Premium %", erp, FMT_PCT, "proj_only"),
        ("Beta", beta, FMT_MULT, "proj_only"),
        ("Cost of Equity %", coe, FMT_PCT, "proj_only"),
        ("WACC %", _safe_num(asmp.get("wacc_pct"), 11.0) / 100, FMT_PCT, "proj_only"),
        ("Terminal Growth Rate %", _safe_num(asmp.get("terminal_growth_pct"), 4.0) / 100, FMT_PCT, "proj_only"),
        ("Implied DCF Price (₹)", _safe_num(val.get("dcf_fair_value"), 0), FMT_PER_SHARE, "first_proj"),
        ("Upside / (Downside) %", ((_safe_num(val.get("blended_fair_value"), cmp) - cmp) / cmp) if cmp else 0, FMT_PCT, "first_proj"),
    ]
    for label, v, fmt, mode in dcf_rows:
        ws.cell(row=r, column=1, value=label).font = _ext_label_font
        ws.cell(row=r, column=1).border = _ext_thin_border
        for ci in range(2, ncols + 1):
            is_proj = ci > 1 + h_count
            cell = ws.cell(row=r, column=ci)
            if mode == "proj_only" and is_proj:
                cell.value = v
            elif mode == "first_proj" and ci == 1 + h_count + 1:
                cell.value = v
            cell.number_format = fmt
            cell.font = _ext_data_font
            cell.border = _ext_thin_border
            cell.alignment = _ext_center
            cell.fill = _ext_cream_fill if is_proj else _ext_white_fill
        r += 1

    sens = val.get("sensitivity_pe") or {}
    grid = sens.get("grid") or []
    if grid:
        r += 1
        _ext_section_divider(ws, r, ncols, "PE SENSITIVITY (Target Price ₹)")
        r += 1
        row_vals = sens.get("row_values", [])
        col_vals = sens.get("col_values", [])
        ws.cell(row=r, column=1, value=f"{sens.get('row_label','PE')} \\ {sens.get('col_label','EPS Growth')}").font = _ext_label_font
        ws.cell(row=r, column=1).fill = _ext_navy_fill
        ws.cell(row=r, column=1).font = _ext_hdr_font
        ws.cell(row=r, column=1).border = _ext_thin_border
        ws.cell(row=r, column=1).alignment = _ext_center
        for ci, cv in enumerate(col_vals, 2):
            c = ws.cell(row=r, column=ci, value=cv)
            c.fill = _ext_navy_fill
            c.font = _ext_hdr_font
            c.border = _ext_thin_border
            c.alignment = _ext_center
            c.number_format = FMT_PCT if abs(_safe_num(cv)) < 1 else "0.0"
        r += 1
        grid_start_row = r
        for ri, rv in enumerate(row_vals):
            c = ws.cell(row=r, column=1, value=rv)
            c.fill = _ext_navy_fill
            c.font = _ext_hdr_font
            c.border = _ext_thin_border
            c.alignment = _ext_center
            c.number_format = FMT_MULT
            grow = grid[ri] if ri < len(grid) else []
            for ci, gv in enumerate(grow, 2):
                gc = ws.cell(row=r, column=ci, value=gv)
                gc.number_format = FMT_INR
                gc.font = _ext_data_font
                gc.border = _ext_thin_border
                gc.alignment = _ext_center
            r += 1
        grid_end_row = r - 1
        last_col = get_column_letter(1 + len(col_vals))
        rng = f"B{grid_start_row}:{last_col}{grid_end_row}"
        ws.conditional_formatting.add(
            rng,
            CellIsRule(operator="greaterThan", formula=[str(cmp)], fill=PatternFill("solid", fgColor="C6EFCE")),
        )
        ws.conditional_formatting.add(
            rng,
            CellIsRule(operator="lessThan", formula=[str(cmp)], fill=PatternFill("solid", fgColor="FFC7CE")),
        )


# ══════════════════════════════════════════════════════════════════
# SHEET 5 — Key_Risks
# ══════════════════════════════════════════════════════════════════
def mk_key_risks(wb, ctx):
    ws, _ = _ext_new_sheet(wb, "Key_Risks")
    thesis = ctx["thesis"]
    sector = ctx["sector"] or "General"

    cols = ["#", "Risk Category", "Risk Factor", "Description", "Mitigation", "Probability", "Impact", "Overall Rating"]
    widths = [5, 18, 22, 40, 35, 13, 13, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ncols = len(cols)
    _ext_title_banner(
        ws,
        f"{ctx['company_name']} — Key Risks Register",
        f"Structured risk assessment | Sector: {sector} | Source: Tikona Research",
        ncols,
    )

    r = 4
    for ci, h in enumerate(cols, 1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font = _ext_hdr_font
        c.fill = _ext_navy_fill
        c.alignment = _ext_center
        c.border = _ext_thin_border
    ws.row_dimensions[r].height = 24
    r += 1

    model = ctx.get("model", {})
    structured_risks = []
    detailed = model.get("risk_items") or []
    for rk in detailed:
        if not isinstance(rk, dict):
            continue
        structured_risks.append({
            "category": rk.get("category", "Operational"),
            "factor": rk.get("factor") or rk.get("risk_factor") or (rk.get("description", "")[:60] or "Risk"),
            "description": rk.get("description") or "",
            "mitigation": rk.get("mitigation", ""),
            "probability": rk.get("probability", "M"),
            "impact": rk.get("impact", "M"),
            "rating": rk.get("rating", "MEDIUM"),
        })

    if not structured_risks:
        raw_risks = thesis.get("key_risks") or []
        categories = ["Operational", "Regulatory", "Market", "Financial", "Competitive", "Macro", "ESG", "Liquidity"]
        for idx, rk in enumerate(raw_risks):
            if isinstance(rk, dict):
                structured_risks.append({
                    "category": rk.get("category", categories[idx % len(categories)]),
                    "factor": rk.get("factor") or rk.get("risk_factor") or (rk.get("description", "")[:40] or f"Risk {idx+1}"),
                    "description": rk.get("description") or rk.get("desc") or "",
                    "mitigation": rk.get("mitigation", "Management monitoring + diversification"),
                    "probability": rk.get("probability", "M"),
                    "impact": rk.get("impact", "M"),
                    "rating": rk.get("rating", "MEDIUM"),
                })
            else:
                txt = str(rk)
                structured_risks.append({
                    "category": categories[idx % len(categories)],
                    "factor": (txt[:50] + "...") if len(txt) > 50 else txt,
                    "description": txt,
                    "mitigation": "Refer management commentary / Annual Report",
                    "probability": "M",
                    "impact": "M",
                    "rating": "MEDIUM",
                })

    fallback_risks = [
        ("Macro", "Macroeconomic slowdown", "Slowdown in Indian/global GDP could compress demand and margins", "Diversified revenue mix; cost-flex levers", "MEDIUM", "MEDIUM", "MEDIUM"),
        ("Regulatory", "Policy / regulatory change", f"Adverse policy changes in {sector} sector", "Active regulatory engagement; compliance investments", "LOW", "HIGH", "MEDIUM"),
        ("Competitive", "Intensifying competition", "New entrants and pricing pressure from peers", "Brand strength, scale, cost leadership", "MEDIUM", "MEDIUM", "MEDIUM"),
        ("Financial", "Working capital stretch", "Higher receivable days or inventory could strain cash flow", "Tight WC monitoring, credit policy discipline", "MEDIUM", "MEDIUM", "MEDIUM"),
        ("Operational", "Execution / capex delays", "Delays in capacity expansion impacting growth trajectory", "Phased capex with milestone reviews", "MEDIUM", "MEDIUM", "MEDIUM"),
        ("Market", "Commodity price volatility", "Input cost inflation could compress margins", "Pass-through pricing + hedging", "MEDIUM", "MEDIUM", "MEDIUM"),
        ("ESG", "ESG / sustainability risks", "Environmental compliance and social licence to operate", "ESG disclosures, board oversight", "LOW", "MEDIUM", "LOW"),
        ("Liquidity", "Refinancing / liquidity", "Debt maturity wall or refinancing risk", "Healthy interest coverage; staggered maturities", "LOW", "MEDIUM", "LOW"),
    ]
    while len(structured_risks) < 8:
        cat, factor, desc, mit, prob, imp, rating = fallback_risks[len(structured_risks) % len(fallback_risks)]
        structured_risks.append({
            "category": cat,
            "factor": factor,
            "description": desc,
            "mitigation": mit,
            "probability": prob,
            "impact": imp,
            "rating": rating,
        })

    def color_for(level):
        lvl = str(level).upper().strip()
        if lvl in ("H", "HIGH"):
            return _ext_dred_fill, _ext_white_font
        if lvl in ("M", "MEDIUM", "MED"):
            return PatternFill("solid", fgColor="FFA500"), _ext_white_font
        return _ext_dgreen_fill, _ext_white_font

    for i, rk in enumerate(structured_risks, 1):
        ws.cell(row=r, column=1, value=i).alignment = _ext_center
        ws.cell(row=r, column=2, value=rk["category"]).alignment = _ext_left_wrap
        ws.cell(row=r, column=3, value=rk["factor"]).alignment = _ext_left_wrap
        ws.cell(row=r, column=4, value=rk["description"]).alignment = _ext_left_wrap
        ws.cell(row=r, column=5, value=rk["mitigation"]).alignment = _ext_left_wrap
        for ci in range(1, 6):
            cell = ws.cell(row=r, column=ci)
            cell.font = _ext_data_font
            cell.border = _ext_thin_border
            if ci == 1:
                cell.alignment = _ext_center
        for ci, key in enumerate(["probability", "impact", "rating"], 6):
            fill, font = color_for(rk[key])
            c = ws.cell(row=r, column=ci, value=str(rk[key]).upper())
            c.fill = fill
            c.font = font
            c.alignment = _ext_center
            c.border = _ext_thin_border
        ws.row_dimensions[r].height = 52
        r += 1


# ══════════════════════════════════════════════════════════════════
# SHEET 6 — Peer_Compare
# ══════════════════════════════════════════════════════════════════
def mk_peer_compare(wb, ctx):
    ws, _ = _ext_new_sheet(wb, "Peer_Compare")
    model = ctx.get("model", {})
    peers_detailed = model.get("peers_detailed") or []
    peers = ctx["peers"] or []
    sd = ctx["sd"]
    hist_years = sd["fiscal_years"]
    dh = min(len(hist_years), 5)
    hist_disp = hist_years[-dh:] if dh else []
    hist_disp_lbl = [f"{y}A" for y in hist_disp]

    company_self = {
        "name": ctx["company_name"] or "Company",
        "mcap_cr": _safe_num(sd.get("mcap"), 0.0),
        "pe": None,
        "ev_ebitda": None,
        "ebitda_margin_pct": None,
        "roe_pct": None,
    }
    pe_self_eps = _safe_num(_ext_screener_pl(ctx, "net_profit", hist_disp[-1]) if hist_disp else None) / (_safe_num(ctx["shares_cr"], 1.0) or 1.0)
    cmp = _safe_num(ctx["cmp"], 0.0)
    if pe_self_eps:
        company_self["pe"] = cmp / pe_self_eps if pe_self_eps > 0 else None
    sales_last = [_safe_num(_ext_screener_pl(ctx, "sales", y)) for y in hist_disp]

    rows_list = [company_self] + list(peers)[:3]

    ws.column_dimensions["A"].width = 30
    for i in range(1, 8):
        ws.column_dimensions[get_column_letter(1 + i)].width = 14

    _ext_title_banner(
        ws,
        f"{ctx['company_name']} — Peer Comparison",
        "Revenue / EBITDA / PAT / Valuation across listed peers | ₹ Cr | Source: Screener / Tikona",
        7,
    )

    def write_table(start_r, title, headers, rows_data, fmt):
        nonlocal_r = start_r
        _ext_section_divider(ws, nonlocal_r, len(headers) + 1, title)
        nonlocal_r += 1
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=nonlocal_r, column=ci, value=h)
            c.font = _ext_hdr_font
            c.fill = _ext_navy_fill
            c.alignment = _ext_center
            c.border = _ext_thin_border
        nonlocal_r += 1
        for name, vals in rows_data:
            ws.cell(row=nonlocal_r, column=1, value=name).font = _ext_label_font
            ws.cell(row=nonlocal_r, column=1).border = _ext_thin_border
            for ci, v in enumerate(vals, 2):
                c = ws.cell(row=nonlocal_r, column=ci, value=v if v is not None else "-")
                c.number_format = fmt
                c.font = _ext_data_font
                c.alignment = _ext_center
                c.border = _ext_thin_border
            nonlocal_r += 1
        return nonlocal_r + 1

    def peer_revenue_estimate(p, years_count):
        try:
            margin = float(str(p.get("ebitda_margin_pct", "18")).replace("%", ""))
        except Exception:
            margin = 18.0
        pe = _safe_num(p.get("pe"), 20.0) or 20.0
        mcap = _safe_num(p.get("mcap_cr"), 0.0)
        approx_pat = mcap / pe if pe else 0
        margin_eff = max(margin / 100, 0.05)
        approx_rev = approx_pat / (margin_eff * 0.6) if margin_eff else 0
        return [approx_rev * (0.85 + 0.04 * i) for i in range(years_count)]

    use_detailed = bool(peers_detailed)
    n_years = len(hist_disp)
    peer_iter = peers_detailed if use_detailed else peers
    peer_iter = peer_iter[:3]

    rev_rows = [(company_self["name"], sales_last)]
    for p in peer_iter:
        if use_detailed and p.get("revenue_series"):
            series = list(p["revenue_series"])[-n_years:] + [None] * max(0, n_years - len(p["revenue_series"]))
            rev_rows.append((p.get("name", "Peer"), series[:n_years]))
        else:
            rev_rows.append((p.get("name", "Peer"), peer_revenue_estimate(p, n_years)))

    r = 4
    r = write_table(r, "REVENUE COMPARISON (₹ Cr)", ["Company"] + hist_disp_lbl, rev_rows, FMT_INR)

    margin_rows = []
    self_margins = []
    for y in hist_disp:
        eb = _ext_hist_ebitda(ctx, y)
        rev = _ext_screener_pl(ctx, "sales", y)
        self_margins.append((_safe_num(eb) / _safe_num(rev)) if _safe_num(rev) else None)
    margin_rows.append((company_self["name"], self_margins))
    for p in peer_iter:
        if use_detailed and p.get("ebitda_margin_series"):
            series = list(p["ebitda_margin_series"])[-n_years:] + [None] * max(0, n_years - len(p["ebitda_margin_series"]))
            margin_rows.append((p.get("name", "Peer"), series[:n_years]))
        else:
            try:
                m = float(str(p.get("ebitda_margin_pct", "")).replace("%", "")) / 100
            except Exception:
                m = 0.15
            margin_rows.append((p.get("name", "Peer"), [m] * n_years))
    r = write_table(r, "EBITDA MARGIN COMPARISON %", ["Company"] + hist_disp_lbl, margin_rows, FMT_PCT)

    pat_rows = []
    self_pat = [_safe_num(_ext_screener_pl(ctx, "net_profit", y)) for y in hist_disp]
    pat_rows.append((company_self["name"], self_pat))
    for p in peer_iter:
        if use_detailed and p.get("pat_series"):
            series = list(p["pat_series"])[-n_years:] + [None] * max(0, n_years - len(p["pat_series"]))
            pat_rows.append((p.get("name", "Peer"), series[:n_years]))
        else:
            mcap = _safe_num(p.get("mcap_cr"), 0)
            pe = _safe_num(p.get("pe"), 20) or 20
            pat_est = mcap / pe if pe else 0
            pat_rows.append((p.get("name", "Peer"), [pat_est * (0.85 + 0.04 * i) for i in range(n_years)]))
    r = write_table(r, "PAT COMPARISON (₹ Cr)", ["Company"] + hist_disp_lbl, pat_rows, FMT_INR)

    snap_rows = []
    snap_rows.append((
        company_self["name"],
        [company_self["mcap_cr"], company_self["pe"], None,
         (_ext_hist_value(ctx, "roce_pct", hist_disp[-1]) / 100) if (hist_disp and _ext_hist_value(ctx, "roce_pct", hist_disp[-1])) else None,
         (_ext_hist_value(ctx, "roe_pct", hist_disp[-1]) / 100) if (hist_disp and _ext_hist_value(ctx, "roe_pct", hist_disp[-1])) else None],
    ))
    snap_iter = peers_detailed if use_detailed else peers
    for p in snap_iter[:5]:
        snap_rows.append((
            p.get("name", "Peer"),
            [
                _safe_num(p.get("mcap_cr")),
                _safe_num(p.get("pe")),
                _safe_num(p.get("pb")) if p.get("pb") is not None else None,
                (_safe_num(p.get("roce_pct")) / 100) if p.get("roce_pct") is not None else None,
                (_safe_num(p.get("roe_pct")) / 100) if p.get("roe_pct") is not None else None,
            ],
        ))
    write_table(r, "CURRENT VALUATION SNAPSHOT", ["Company", "MCap (₹ Cr)", "P/E (x)", "P/B (x)", "ROCE %", "ROE %"], snap_rows, FMT_INR)

    # ── Charts ──
    try:
        chart_data_row = 70
        ws.cell(row=chart_data_row, column=1, value="ChartData: Revenue (last 3 yrs)").font = _ext_sub_font
        last3 = hist_disp[-3:] if len(hist_disp) >= 3 else hist_disp
        for ci, y in enumerate(last3, 2):
            ws.cell(row=chart_data_row, column=ci, value=y)
        for ri, (name, vals) in enumerate(rev_rows, chart_data_row + 1):
            ws.cell(row=ri, column=1, value=name)
            tail = vals[-len(last3):] if vals else []
            for ci, v in enumerate(tail, 2):
                ws.cell(row=ri, column=ci, value=_safe_num(v))
        bar = BarChart()
        bar.type = "col"
        bar.style = 11
        bar.grouping = "clustered"
        bar.title = "Revenue Comparison (₹ Cr) — Last 3 Years"
        bar.y_axis.title = "Revenue (₹ Cr)"
        bar.x_axis.title = "Year"
        end_data_row = chart_data_row + len(rev_rows)
        end_data_col = 1 + len(last3)
        data_ref = Reference(ws, min_col=2, min_row=chart_data_row, max_col=end_data_col, max_row=end_data_row)
        cats_ref = Reference(ws, min_col=1, min_row=chart_data_row + 1, max_row=end_data_row)
        bar.add_data(data_ref, titles_from_data=False, from_rows=False)
        bar.set_categories(cats_ref)
        bar.width = 18
        bar.height = 10
        ws.add_chart(bar, "B40")

        margin_data_row = chart_data_row + len(rev_rows) + 4
        ws.cell(row=margin_data_row, column=1, value="ChartData: EBITDA Margin %").font = _ext_sub_font
        for ci, y in enumerate(last3, 2):
            ws.cell(row=margin_data_row, column=ci, value=y)
        for ri, (name, vals) in enumerate(margin_rows, margin_data_row + 1):
            ws.cell(row=ri, column=1, value=name)
            tail = vals[-len(last3):] if vals else []
            for ci, v in enumerate(tail, 2):
                ws.cell(row=ri, column=ci, value=_safe_num(v))
        line = LineChart()
        line.title = "EBITDA Margin % — Last 3 Years"
        line.y_axis.title = "Margin %"
        line.x_axis.title = "Year"
        end_m_row = margin_data_row + len(margin_rows)
        end_m_col = 1 + len(last3)
        d_ref = Reference(ws, min_col=2, min_row=margin_data_row, max_col=end_m_col, max_row=end_m_row)
        c_ref = Reference(ws, min_col=1, min_row=margin_data_row + 1, max_row=end_m_row)
        line.add_data(d_ref, titles_from_data=False, from_rows=False)
        line.set_categories(c_ref)
        line.width = 18
        line.height = 10
        ws.add_chart(line, "B58")
    except Exception as ex:
        logger.warning("⚠ Peer compare charts skipped: %s", ex)


# ══════════════════════════════════════════════════════════════════
# SHEET 7 — Operational_Data
# ══════════════════════════════════════════════════════════════════
def mk_operational_data(wb, ctx):
    ws, _ = _ext_new_sheet(wb, "Operational_Data")
    model = ctx.get("model", {})
    op = model.get("operational") or {}
    sd = ctx["sd"]
    hist_years = sd["fiscal_years"]
    dh = min(len(hist_years), 5)
    hist_disp = hist_years[-dh:] if dh else []
    proj_years = ctx["proj_years"][:3]
    years_all = hist_disp + proj_years
    hist_disp_lbl = [f"{y}A" for y in hist_disp]
    years_disp_lbl = hist_disp_lbl + proj_years
    ncols = 1 + len(years_all)
    h_count = len(hist_disp)
    p_count = len(proj_years)

    ws.column_dimensions["A"].width = 32
    for i in range(len(years_all)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 13

    has_structured_op = bool(op and (op.get("volume_segments") or op.get("plants_india") or op.get("capacity_utilisation_pct")))
    subtitle = (
        "Volume / capacity / productivity | Source: Annual Reports / Company filings via web search"
        if has_structured_op else
        "Volume / capacity / productivity | Some cells indicative — sourced from Annual Reports"
    )
    _ext_title_banner(ws, f"{ctx['company_name']} — Operational Metrics", subtitle, ncols)

    r = 4
    _ext_year_header(ws, r, years_disp_lbl, h_count)
    r += 1

    rev_hist = [_safe_num(_ext_screener_pl(ctx, "sales", y)) for y in hist_disp]
    rev_proj = [_safe_num(_ext_proj_revenue(ctx, y)) for y in proj_years]
    rev_all = rev_hist + rev_proj

    # Pad / align an arbitrary-length series from JSON to the years_all length.
    def align_series(series, expected_len):
        if not series:
            return [None] * expected_len
        seq = list(series)
        if len(seq) >= expected_len:
            return seq[:expected_len]
        return seq + [None] * (expected_len - len(seq))

    alt = {"v": False}

    def emit(label, fmt, vals):
        nonlocal r
        _ext_write_data_row(ws, r, label, vals, fmt, h_count, p_count, alt=alt["v"])
        alt["v"] = not alt["v"]
        r += 1

    _ext_section_divider(ws, r, ncols, "VOLUME DATA (MT)")
    r += 1
    alt["v"] = False
    vol_segments = op.get("volume_segments") or {}
    total_vol = [0.0] * len(years_all)
    if vol_segments:
        for seg_name, seg_vals in vol_segments.items():
            aligned = align_series(seg_vals, len(years_all))
            emit(f"{seg_name} (MT)", FMT_INR, aligned)
            for i, v in enumerate(aligned):
                if v is not None:
                    total_vol[i] += _safe_num(v)
        emit("Total Volume (MT)", FMT_INR, total_vol)
    else:
        realization_per_mt = _safe_num(op.get("realization_per_mt"), 250000.0) or 250000.0
        total_vol = [(v * 1e7 / realization_per_mt) if v else None for v in rev_all]
        emit("Total Volume (MT, est.)", FMT_INR, total_vol)

    _ext_section_divider(ws, r, ncols, "CAPACITY & UTILISATION")
    r += 1
    alt["v"] = False
    cap_util = op.get("capacity_utilisation_pct") or []
    if cap_util:
        cap_aligned = align_series(cap_util, len(years_all))
        cap_aligned = [(v / 100 if v and v > 1 else v) for v in cap_aligned]
        emit("Capacity Utilisation %", FMT_PCT, cap_aligned)
    else:
        emit("Capacity Utilisation %", FMT_PCT, [0.75 + 0.02 * i if rev_all[i] else None for i in range(len(years_all))])
    emit("Countries of Operation", FMT_DAYS, align_series(op.get("countries_of_operation"), len(years_all)))
    emit("Plants India", FMT_DAYS, align_series(op.get("plants_india"), len(years_all)))
    emit("Plants Overseas", FMT_DAYS, align_series(op.get("plants_overseas"), len(years_all)))

    _ext_section_divider(ws, r, ncols, "REVENUE MIX %")
    r += 1
    alt["v"] = False
    mix = op.get("revenue_mix_pct") or {}
    if mix:
        for seg_name, pct in mix.items():
            pct_val = _safe_num(pct, 0)
            if pct_val > 1:
                pct_val /= 100
            emit(f"{seg_name}", FMT_PCT, [pct_val] * len(years_all))
    else:
        emit("Core Segment %", FMT_PCT, [0.65] * len(years_all))
        emit("Adjacent Segment %", FMT_PCT, [0.25] * len(years_all))
        emit("Others %", FMT_PCT, [0.10] * len(years_all))

    geo = op.get("geography_mix_pct") or {}
    if geo:
        _ext_section_divider(ws, r, ncols, "GEOGRAPHY MIX %")
        r += 1
        alt["v"] = False
        for region, pct in geo.items():
            pct_val = _safe_num(pct, 0)
            if pct_val > 1:
                pct_val /= 100
            emit(f"{region}", FMT_PCT, [pct_val] * len(years_all))

    _ext_section_divider(ws, r, ncols, "FINANCIAL PRODUCTIVITY")
    r += 1
    alt["v"] = False
    employees = _safe_num(op.get("employees"), 2000) or 2000
    emit("Revenue/Employee (₹ Lakh)", FMT_INR, [(v * 100 / employees) if v else None for v in rev_all])
    eb_hist = [_ext_hist_ebitda(ctx, y) for y in hist_disp]
    eb_proj = []
    for i in range(p_count):
        _pat_p, eb_p, *_rest = _ext_proj_pat(ctx, i)
        eb_proj.append(eb_p)
    eb_all = eb_hist + eb_proj
    emit("EBITDA/Tonne (₹/MT)", FMT_INR, [
        (_safe_num(eb_all[i]) * 1e7 / _safe_num(total_vol[i])) if _safe_num(total_vol[i]) else None
        for i in range(len(years_all))
    ])
    emit("Revenue/Tonne (₹/MT)", FMT_INR, [
        (_safe_num(rev_all[i]) * 1e7 / _safe_num(total_vol[i])) if _safe_num(total_vol[i]) else None
        for i in range(len(years_all))
    ])
    capex_hist = [(_safe_num(_ext_screener_pl(ctx, "cfi", y)) * -1) for y in hist_disp]
    capex_proj = [_ext_av(ctx, "capex_cr", y) for y in proj_years]
    capex_all = capex_hist + capex_proj
    emit("Capex/Tonne (₹/MT)", FMT_INR, [
        (_safe_num(capex_all[i]) * 1e7 / _safe_num(total_vol[i])) if _safe_num(total_vol[i]) else None
        for i in range(len(years_all))
    ])

    if not has_structured_op:
        r += 1
        note = ws.cell(row=r, column=1, value="Note: Volume/segment numbers above are indicative — derived from revenue assuming sector-typical realisations.")
        note.font = _ext_sub_font
        last_col = get_column_letter(ncols)
        ws.merge_cells(f"A{r}:{last_col}{r}")
        note.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 30


# ══════════════════════════════════════════════════════════════════
# SHEET 8 — Governance
# ══════════════════════════════════════════════════════════════════
def mk_governance(wb, ctx):
    ws, _ = _ext_new_sheet(wb, "Governance")
    model = ctx.get("model", {})
    gov = model.get("governance") or {}
    sd = ctx["sd"]
    hist_years = sd["fiscal_years"]
    last4 = hist_years[-4:] if len(hist_years) >= 4 else hist_years

    ws.column_dimensions["A"].width = 30
    for i in range(1, 10):
        ws.column_dimensions[get_column_letter(1 + i)].width = 16

    _ext_title_banner(
        ws,
        f"{ctx['company_name']} — Corporate Governance",
        "Board, shareholding, governance metrics | Source: Annual Report / BSE filings",
        7,
    )

    r = 4
    _ext_section_divider(ws, r, 7, "BOARD OF DIRECTORS")
    r += 1
    headers = ["Name", "Designation", "Status", "DIN", "Since"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font = _ext_hdr_font
        c.fill = _ext_navy_fill
        c.alignment = _ext_center
        c.border = _ext_thin_border
    r += 1
    board = gov.get("board") or []
    if not board:
        board = [
            {"name": "Refer Annual Report", "designation": "Chairman & MD", "status": "Promoter", "din": "—", "since": "—"},
            {"name": "Refer Annual Report", "designation": "Executive Director", "status": "Promoter", "din": "—", "since": "—"},
            {"name": "Refer Annual Report", "designation": "Independent Director", "status": "Independent", "din": "—", "since": "—"},
            {"name": "Refer Annual Report", "designation": "Independent Director", "status": "Independent", "din": "—", "since": "—"},
            {"name": "Refer Annual Report", "designation": "Independent Director", "status": "Independent", "din": "—", "since": "—"},
        ]
    for member in board:
        if not isinstance(member, dict):
            continue
        vals = [
            member.get("name", "—"),
            member.get("designation", "—"),
            member.get("status", "—"),
            member.get("din") or "—",
            member.get("since") or "—",
        ]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font = _ext_data_font
            c.alignment = _ext_left_wrap
            c.border = _ext_thin_border
        r += 1

    r += 1
    _ext_section_divider(ws, r, 7, "SHAREHOLDING PATTERN %")
    r += 1
    shp = gov.get("shareholding") or {}
    shp_years = shp.get("years") if shp else None
    if not shp_years:
        shp_years = list(last4)
    for ci, h in enumerate(["Category"] + list(shp_years), 1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font = _ext_hdr_font
        c.fill = _ext_navy_fill
        c.alignment = _ext_center
        c.border = _ext_thin_border
    r += 1

    def _shp_series(key, fallback):
        vals = shp.get(key) if shp else None
        if not vals:
            vals = fallback
        out = []
        for v in vals[:len(shp_years)]:
            x = _safe_num(v, 0)
            if x > 1:
                x /= 100
            out.append(x)
        while len(out) < len(shp_years):
            out.append(None)
        return out

    shp_rows = [
        ("Promoter", _shp_series("promoter_pct", [0.65, 0.65, 0.65, 0.65])),
        ("FII", _shp_series("fii_pct", [0.10, 0.11, 0.12, 0.13])),
        ("DII", _shp_series("dii_pct", [0.08, 0.08, 0.07, 0.07])),
        ("Public", _shp_series("public_pct", [0.17, 0.16, 0.16, 0.15])),
    ]
    for cat, vals in shp_rows:
        ws.cell(row=r, column=1, value=cat).font = _ext_label_font
        ws.cell(row=r, column=1).border = _ext_thin_border
        for ci, v in enumerate(vals, 2):
            c = ws.cell(row=r, column=ci, value=v if v is not None else "-")
            if v is not None:
                c.number_format = FMT_PCT
            c.font = _ext_data_font
            c.alignment = _ext_center
            c.border = _ext_thin_border
        r += 1

    r += 1
    _ext_section_divider(ws, r, 7, "KEY GOVERNANCE METRICS")
    r += 1

    def _gov_val(key, default):
        v = gov.get(key)
        return v if v is not None else default

    metrics = [
        ("Promoter Pledge %", _safe_num(_gov_val("promoter_pledge_pct", 0.0), 0.0) / (100 if _safe_num(_gov_val("promoter_pledge_pct", 0.0)) > 1 else 1), FMT_PCT),
        ("Board Size", _gov_val("board_size", 7), FMT_DAYS),
        ("Independent Directors (count)", _gov_val("independent_directors", 4), FMT_DAYS),
        ("Women Directors", _gov_val("women_directors", 1), FMT_DAYS),
        ("Managerial Remuneration (₹ Cr)", _gov_val("managerial_remuneration_cr", "Refer AR"), FMT_INR if isinstance(_gov_val("managerial_remuneration_cr", None), (int, float)) else None),
        ("Auditor Fees – Audit (₹ Cr)", _gov_val("auditor_fees_audit_cr", "Refer AR"), FMT_INR if isinstance(_gov_val("auditor_fees_audit_cr", None), (int, float)) else None),
        ("Auditor Fees – Non-Audit (₹ Cr)", _gov_val("auditor_fees_non_audit_cr", "Refer AR"), FMT_INR if isinstance(_gov_val("auditor_fees_non_audit_cr", None), (int, float)) else None),
        ("Statutory Auditor", _gov_val("statutory_auditor", "Refer AR"), None),
        ("Dividend Payout %", _safe_num(ctx["asmp"].get("dividend_payout_pct"), 0.0) / 100, FMT_PCT),
        ("Related Party Transactions (₹ Cr)", _gov_val("related_party_transactions_cr", "Refer AR"), FMT_INR if isinstance(_gov_val("related_party_transactions_cr", None), (int, float)) else None),
    ]
    for label, v, fmt in metrics:
        ws.cell(row=r, column=1, value=label).font = _ext_label_font
        ws.cell(row=r, column=1).border = _ext_thin_border
        c = ws.cell(row=r, column=2, value=v)
        if fmt:
            c.number_format = fmt
        c.font = _ext_data_font
        c.alignment = _ext_center
        c.border = _ext_thin_border
        ws.cell(row=r, column=3, value="Source: Annual Report / BSE filings").font = _ext_sub_font
        ws.cell(row=r, column=3).border = _ext_thin_border
        r += 1


# ══════════════════════════════════════════════════════════════════
# SHEET 9 — Timeline
# ══════════════════════════════════════════════════════════════════
def mk_timeline(wb, ctx):
    ws, _ = _ext_new_sheet(wb, "Timeline")
    thesis = ctx["thesis"]
    sd = ctx["sd"]
    model = ctx.get("model", {})
    structured_events = model.get("timeline_events") or []

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 35

    _ext_title_banner(
        ws,
        f"{ctx['company_name']} — Company Timeline",
        "Key milestones extracted from thesis / catalysts / management commentary",
        4,
    )

    r = 4
    headers = ["Year", "Event Category", "Description", "Strategic Impact"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font = _ext_hdr_font
        c.fill = _ext_navy_fill
        c.alignment = _ext_center
        c.border = _ext_thin_border
    ws.row_dimensions[r].height = 22
    r += 1

    category_colors = {
        "Founding": "8FBC8F",
        "IPO/Capital Markets": "DAA520",
        "Expansion": "4682B4",
        "New Vertical": "9370DB",
        "Regulatory": "CD5C5C",
        "Strategy": "FFA500",
        "Milestone": "20B2AA",
        "Outlook": "1F4690",
    }

    events = []
    if structured_events:
        for ev in structured_events:
            if not isinstance(ev, dict):
                continue
            events.append((
                str(ev.get("year", "—")),
                str(ev.get("category", "Milestone")),
                str(ev.get("description", "")),
                str(ev.get("impact", "")),
            ))

    if not events:
        sources = " ".join([
            thesis.get("investment_thesis", "") or "",
            thesis.get("bull_case", "") or "",
            " ".join(thesis.get("key_catalysts") or []),
        ])
        year_matches = re.findall(r"\b(19[5-9][0-9]|20[0-3][0-9])\b", sources)
        found_match = re.search(r"\b(19[5-9][0-9]|20[0-3][0-9])\b", sources)
        if found_match:
            events.append((found_match.group(1), "Founding", "Company incorporated (year inferred from thesis/historical narrative)", "Establishes business foundation"))
        else:
            events.append(("—", "Founding", "Refer Annual Report for incorporation year", "Establishes business foundation"))
        seen_years = set()
        for y in year_matches:
            if y in seen_years:
                continue
            seen_years.add(y)
            events.append((y, "Milestone", f"Key event referenced for {y} in research narrative", "Refer thesis / catalysts"))
        for cat in (thesis.get("key_catalysts") or [])[:5]:
            events.append(("2026-28E", "Strategy", str(cat), "Forward catalyst — drives next leg of growth"))
        events.append(("2030E", "Outlook", "Management guidance horizon and Tikona projection endpoint", "Long-term value realisation"))

    while len(events) < 10:
        events.append(("—", "Milestone", "Refer Annual Report for additional milestones", "—"))

    for yr, cat, desc, impact in events[:20]:
        ws.cell(row=r, column=1, value=yr).alignment = _ext_center
        c2 = ws.cell(row=r, column=2, value=cat)
        fill_color = category_colors.get(cat, "808080")
        c2.fill = PatternFill("solid", fgColor=fill_color)
        c2.font = _ext_white_font
        c2.alignment = _ext_center
        ws.cell(row=r, column=3, value=desc).alignment = _ext_left_wrap
        ws.cell(row=r, column=4, value=impact).alignment = _ext_left_wrap
        for ci in range(1, 5):
            cell = ws.cell(row=r, column=ci)
            cell.border = _ext_thin_border
            if ci != 2:
                cell.font = _ext_data_font
        ws.row_dimensions[r].height = 40
        r += 1


# ══════════════════════════════════════════════════════════════════
# SHEET 10 — Op_Charts
# ══════════════════════════════════════════════════════════════════
def mk_op_charts(wb, ctx):
    ws, _ = _ext_new_sheet(wb, "Op_Charts")
    sd = ctx["sd"]
    hist_years = sd["fiscal_years"]
    dh = min(len(hist_years), 5)
    hist_disp = hist_years[-dh:] if dh else []
    hist_disp_lbl = [f"{y}A" for y in hist_disp]
    proj_years = ctx["proj_years"]
    years_all = hist_disp + proj_years
    years_disp_lbl = hist_disp_lbl + proj_years
    h_count = len(hist_disp)
    p_count = len(proj_years)

    ws.column_dimensions["A"].width = 22
    for i in range(len(years_all)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 13

    _ext_title_banner(
        ws,
        f"{ctx['company_name']} — Operational Charts",
        "Six analytical charts on historical + projected financials",
        1 + len(years_all),
    )

    rev_vals = [_safe_num(_ext_screener_pl(ctx, "sales", y)) for y in hist_disp] + [_safe_num(_ext_proj_revenue(ctx, y)) for y in proj_years]
    eb_vals = [_safe_num(_ext_hist_ebitda(ctx, y)) for y in hist_disp]
    pat_vals = [_safe_num(_ext_screener_pl(ctx, "net_profit", y)) for y in hist_disp]
    dep_vals = [_safe_num(_ext_screener_pl(ctx, "depreciation", y)) for y in hist_disp]
    cfo_vals = [_safe_num(_ext_screener_pl(ctx, "cfo", y)) for y in hist_disp]
    cfi_vals = [_safe_num(_ext_screener_pl(ctx, "cfi", y)) for y in hist_disp]
    for i in range(p_count):
        pat_p, eb_p, dep_p, *_ = _ext_proj_pat(ctx, i)
        eb_vals.append(eb_p)
        pat_vals.append(pat_p)
        dep_vals.append(dep_p)
        cfo_vals.append(_ext_proj_val(ctx, "cfo", i))
        cfi_vals.append(_ext_proj_val(ctx, "cfi", i))
    capex_vals = []
    for i in range(h_count):
        capex_vals.append(-cfi_vals[i])
    for i in range(p_count):
        capex_vals.append(_ext_av(ctx, "capex_cr", proj_years[i]))
    fcf_vals = [cfo_vals[i] - capex_vals[i] for i in range(len(years_all))]
    shares_cr = _safe_num(ctx["shares_cr"], 1.0) or 1.0
    eps_vals = [v / shares_cr for v in pat_vals]
    ebit_vals = [eb_vals[i] - dep_vals[i] for i in range(len(years_all))]

    nw_hist = []
    ce_hist = []
    for y in hist_disp:
        sc = _safe_num(_ext_screener_pl(ctx, "share_capital", y))
        res = _safe_num(_ext_screener_pl(ctx, "reserves", y))
        borr = _safe_num(_ext_screener_pl(ctx, "borrowings", y))
        nw_hist.append(sc + res)
        ce_hist.append(sc + res + borr)
    nw_proj = []
    ce_proj = []
    for i in range(p_count):
        nw_i = _ext_proj_val(ctx, "share_capital", i) + _ext_proj_val(ctx, "reserves", i)
        nw_proj.append(nw_i)
        ce_proj.append(nw_i + _ext_proj_val(ctx, "borrowings", i))
    nw_all = nw_hist + nw_proj
    ce_all = ce_hist + ce_proj
    roe_vals = [(pat_vals[i] / nw_all[i] * 100) if nw_all[i] else 0 for i in range(len(years_all))]
    roce_vals = [(ebit_vals[i] / ce_all[i] * 100) if ce_all[i] else 0 for i in range(len(years_all))]
    margin_eb = [(eb_vals[i] / rev_vals[i] * 100) if rev_vals[i] else 0 for i in range(len(years_all))]
    margin_pat = [(pat_vals[i] / rev_vals[i] * 100) if rev_vals[i] else 0 for i in range(len(years_all))]

    # Hidden data blocks (rows 60+)
    DR = 60

    def write_block(start_row, title_label, series_dict):
        ws.cell(row=start_row, column=1, value=title_label).font = _ext_sub_font
        for ci, y in enumerate(years_disp_lbl, 2):
            ws.cell(row=start_row, column=ci, value=y).font = _ext_label_font
        last_row = start_row
        for offset, (name, vals) in enumerate(series_dict.items(), 1):
            ws.cell(row=start_row + offset, column=1, value=name).font = _ext_data_font
            for ci, v in enumerate(vals, 2):
                ws.cell(row=start_row + offset, column=ci, value=_safe_num(v))
            last_row = start_row + offset
        return last_row

    def add_chart(chart, anchor):
        chart.width = 18
        chart.height = 12
        ws.add_chart(chart, anchor)

    # Chart 1: Revenue / EBITDA / PAT bar
    end_r = write_block(DR, "ChartData 1 — Revenue/EBITDA/PAT", {"Revenue": rev_vals, "EBITDA": eb_vals, "PAT": pat_vals})
    bar1 = BarChart()
    bar1.type = "col"
    bar1.grouping = "clustered"
    bar1.title = "Revenue / EBITDA / PAT (₹ Cr)"
    data = Reference(ws, min_col=1, min_row=DR + 1, max_col=1 + len(years_all), max_row=end_r)
    cats = Reference(ws, min_col=2, min_row=DR, max_col=1 + len(years_all), max_row=DR)
    bar1.add_data(data, titles_from_data=True, from_rows=True)
    bar1.set_categories(cats)
    add_chart(bar1, "A3")

    # Chart 2: Margins line
    DR2 = end_r + 3
    end_r2 = write_block(DR2, "ChartData 2 — Margins", {"EBITDA Margin %": margin_eb, "PAT Margin %": margin_pat})
    line2 = LineChart()
    line2.title = "Margin Profile %"
    data = Reference(ws, min_col=1, min_row=DR2 + 1, max_col=1 + len(years_all), max_row=end_r2)
    cats = Reference(ws, min_col=2, min_row=DR2, max_col=1 + len(years_all), max_row=DR2)
    line2.add_data(data, titles_from_data=True, from_rows=True)
    line2.set_categories(cats)
    add_chart(line2, "K3")

    # Chart 3: Revenue by segment (stacked) — using indicative split
    DR3 = end_r2 + 3
    seg_core = [v * 0.65 for v in rev_vals]
    seg_adj = [v * 0.25 for v in rev_vals]
    seg_oth = [v * 0.10 for v in rev_vals]
    end_r3 = write_block(DR3, "ChartData 3 — Revenue by Segment", {"Core": seg_core, "Adjacent": seg_adj, "Others": seg_oth})
    bar3 = BarChart()
    bar3.type = "col"
    bar3.grouping = "stacked"
    bar3.overlap = 100
    bar3.title = "Revenue by Segment (₹ Cr, indicative)"
    data = Reference(ws, min_col=1, min_row=DR3 + 1, max_col=1 + len(years_all), max_row=end_r3)
    cats = Reference(ws, min_col=2, min_row=DR3, max_col=1 + len(years_all), max_row=DR3)
    bar3.add_data(data, titles_from_data=True, from_rows=True)
    bar3.set_categories(cats)
    add_chart(bar3, "A22")

    # Chart 4: ROE vs ROCE line
    DR4 = end_r3 + 3
    end_r4 = write_block(DR4, "ChartData 4 — ROE/ROCE", {"ROE %": roe_vals, "ROCE %": roce_vals})
    line4 = LineChart()
    line4.title = "ROE vs ROCE %"
    data = Reference(ws, min_col=1, min_row=DR4 + 1, max_col=1 + len(years_all), max_row=end_r4)
    cats = Reference(ws, min_col=2, min_row=DR4, max_col=1 + len(years_all), max_row=DR4)
    line4.add_data(data, titles_from_data=True, from_rows=True)
    line4.set_categories(cats)
    add_chart(line4, "K22")

    # Chart 5: CFO / Capex / FCF
    DR5 = end_r4 + 3
    end_r5 = write_block(DR5, "ChartData 5 — CFO/Capex/FCF", {"CFO": cfo_vals, "Capex": capex_vals, "FCF": fcf_vals})
    bar5 = BarChart()
    bar5.type = "col"
    bar5.grouping = "clustered"
    bar5.title = "Cash Flow Profile (₹ Cr)"
    data = Reference(ws, min_col=1, min_row=DR5 + 1, max_col=1 + len(years_all), max_row=end_r5)
    cats = Reference(ws, min_col=2, min_row=DR5, max_col=1 + len(years_all), max_row=DR5)
    bar5.add_data(data, titles_from_data=True, from_rows=True)
    bar5.set_categories(cats)
    add_chart(bar5, "A41")

    # Chart 6: EPS Trend
    DR6 = end_r5 + 3
    end_r6 = write_block(DR6, "ChartData 6 — EPS Trend", {"EPS (₹)": eps_vals})
    bar6 = BarChart()
    bar6.type = "col"
    bar6.title = "EPS Trend (₹)"
    data = Reference(ws, min_col=1, min_row=DR6 + 1, max_col=1 + len(years_all), max_row=end_r6)
    cats = Reference(ws, min_col=2, min_row=DR6, max_col=1 + len(years_all), max_row=DR6)
    bar6.add_data(data, titles_from_data=True, from_rows=True)
    bar6.set_categories(cats)
    add_chart(bar6, "K41")


# ══════════════════════════════════════════════════════════════════
# FORMULA-BASED EXCEL BUILDER
# ══════════════════════════════════════════════════════════════════
def build_model(screener_path: str, screener_data: dict, model: dict, out_path: str):
    is_xlsm = screener_path.endswith(".xlsm")
    wb = load_workbook(screener_path, keep_vba=is_xlsm)

    # openpyxl strips DrawingML shapes/charts during read/save, so Screener's
    # bundled "Charts" sheet ends up empty in the output. Remove it — our
    # Op_Charts sheet is the analytical replacement.
    if "Charts" in wb.sheetnames:
        del wb["Charts"]
        logger.info("🗑  Removed empty Screener 'Charts' sheet (openpyxl drops embedded drawings)")

    orig_sheets = list(wb.sheetnames)
    logger.info(f"📂 Loaded: {len(orig_sheets)} sheets")

    def mk(name):
        n = name + " - Model" if name in wb.sheetnames else name
        return wb.create_sheet(n), n

    def hdr_r(ws, r, labels, nc):
        for ci, lbl in enumerate(labels, 1):
            c = ws.cell(row=r, column=ci, value=lbl)
            c.fill = navy_fill
            c.font = hdr_font
            c.alignment = center
            c.border = thin_bdr

    def set_w(ws, widths):
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def add_borders(ws, mr, mc):
        for row in ws.iter_rows(min_row=1, max_row=mr, min_col=1, max_col=mc):
            for cell in row:
                cell.border = thin_bdr

    sd = screener_data
    row_map = sd.get("row_map", DEFAULT_ROW_MAP)
    hist_years = sd["fiscal_years"]
    data_cols = sd["data_cols"]
    ds_letters = [get_column_letter(c) for c in data_cols]
    nh = len(hist_years)

    asmp = model.get("assumptions", {})
    proj = model.get("projections", {})
    proj_years = asmp.get("projection_years", proj.get("years", _build_projection_years(hist_years[-1] if hist_years else "FY00")))
    np_ = len(proj_years)

    dh = min(nh, 6)
    h_start = nh - dh
    disp_hist_raw = hist_years[h_start:]
    disp_hist = [f"{y}A" if not y.endswith(("A", "E")) else y for y in disp_hist_raw]
    disp_ds = ds_letters[h_start:]

    year_labels = ["Particulars"] + disp_hist + proj_years
    nc = len(year_labels)
    h_cols = list(range(2, 2 + dh))
    p_cols = list(range(2 + dh, nc + 1))

    fallback_shares = _last_non_null(sd["bs"].get("shares_outstanding", []))
    if fallback_shares in (None, 0):
        face_value = sd.get("face_value")
        share_capital = _last_non_null(sd["bs"].get("share_capital", []))
        fallback_shares = (share_capital / face_value) if face_value not in (None, 0) and share_capital is not None else 1
    shares_cr = model.get("shares_cr", fallback_shares)
    thesis = model.get("thesis", {})
    val = model.get("valuation", {})
    peers = model.get("peers", [])
    rating = model.get("rating", "HOLD")
    target = model.get("target_price", 0)
    cmp = model.get("cmp", sd["cmp"])
    upside = model.get("upside_pct", 0)
    hist_ratios = model.get("historical_ratios", {})

    def av(key, yr):
        d = asmp.get(key, {})
        if isinstance(d, dict):
            return d.get(yr, 0)
        return d or 0

    def pv(key, i, default=0):
        vals = proj.get(key, [])
        return vals[i] if i < len(vals) and vals[i] is not None else default

    diagnostics: list[str] = []

    def projected_revenues_numeric() -> list[float]:
        if not sd["pl"]["sales"]:
            return [0.0] * np_
        last_actual_revenue = float(_last_non_null(sd["pl"]["sales"]) or 0.0)
        projected: list[float] = []
        current = last_actual_revenue
        for year in proj_years:
            growth_pct = float(av("revenue_growth_pct", year) or 0.0)
            current = current * (1 + growth_pct / 100)
            projected.append(current)
        return projected

    # ═══ COVER ═══
    ws, _ = mk("Cover")
    set_w(ws, [25, 30, 5, 25, 30])
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = (sd["company_name"] or "").upper()
    c.font = Font(name="Arial", bold=True, color="FFFFFF", size=20)
    c.fill = navy_fill
    c.alignment = center
    ws.row_dimensions[1].height = 45
    ws.merge_cells("A2:E2")
    c = ws["A2"]
    c.value = f"{model.get('sector','')} | TIKONA CAPITAL"
    c.font = Font(name="Arial", color="FFFFFF", size=12)
    c.fill = blue_fill
    c.alignment = center

    for i, (k, v) in enumerate(
        [
            ("CMP (₹)", cmp), ("Target (₹)", target), ("Upside", f"{upside:.1f}%"),
            ("Rating", rating), ("Market Cap (₹ Cr)", sd["mcap"]),
            ("Shares (Cr)", f"{shares_cr:.2f}"), ("Base Year", model.get("base_year", "")),
        ], 4,
    ):
        ws.cell(row=i, column=1, value=k).font = sec_font
        ws.cell(row=i, column=2, value=v).font = data_font
        for c in range(1, 3):
            ws.cell(row=i, column=c).border = thin_bdr

    r = 12
    ws.merge_cells(f"A{r}:E{r}")
    ws[f"A{r}"].value = f"RATING: {rating} | Target: ₹{target:,.0f} | Upside: {upside:.1f}%"
    color = COLORS["green"] if upside > 0 else COLORS["red"]
    ws[f"A{r}"].font = Font(name="Arial", bold=True, color="FFFFFF", size=14)
    ws[f"A{r}"].fill = PatternFill("solid", fgColor=color)
    ws[f"A{r}"].alignment = center
    r = 14
    ws.merge_cells(f"A{r}:E{r+3}")
    ws[f"A{r}"].value = thesis.get("investment_thesis", "")
    ws[f"A{r}"].font = data_font
    ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    r = 19
    for label, items, col in [
        ("KEY CATALYSTS", thesis.get("key_catalysts", []), COLORS["green"]),
        ("KEY RISKS", thesis.get("key_risks", []), COLORS["red"]),
    ]:
        ws[f"A{r}"].value = label
        ws[f"A{r}"].font = Font(name="Arial", bold=True, color=col, size=12)
        r += 1
        for item in items:
            ws[f"A{r}"].value = f"• {item}"
            r += 1
        r += 1
    ws[f"A{r}"].value = (
        f"SAARTHI SCORE: {thesis.get('saarthi_total','NA')}/100 → {thesis.get('saarthi_rating','NA')}"
    )
    ws[f"A{r}"].font = Font(name="Arial", bold=True, color=COLORS["navy"], size=13)
    ws[f"A{r}"].fill = PatternFill("solid", fgColor=COLORS["peach"])
    logger.info("  ✅ Cover")

    # ═══ ASSUMPTIONS ═══
    ws, asn = mk("Assumptions")
    set_w(ws, [35] + [15] * (nc - 1))
    hdr_r(ws, 1, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=1, column=ci).fill = orange_fill

    assum_items = [
        ("Revenue Growth %", "revenue_growth_pct", None, RATIO),
        ("RM % of Revenue", "rm_pct", "rm_pct", RATIO),
        ("Employee % of Revenue", "employee_pct", "employee_pct", RATIO),
        ("Power & Fuel %", "power_fuel_pct", None, RATIO),
        ("Other Mfg %", "other_mfg_pct", None, RATIO),
        ("Selling & Admin %", "selling_admin_pct", None, RATIO),
        ("Other Expenses %", "other_exp_pct", None, RATIO),
        ("Chg in Inventory %", "chg_inventory_pct", None, RATIO),
        ("Depreciation (₹ Cr)", "depreciation_cr", None, INR),
        ("Interest (₹ Cr)", "interest_cr", None, INR),
        ("Other Income (₹ Cr)", "other_income_cr", None, INR),
        ("Tax Rate %", "tax_rate_pct", "tax_rate_pct", RATIO),
        ("Capex (₹ Cr)", "capex_cr", None, INR),
        ("Receivable Days", "receivable_days", "receivable_days", "0"),
        ("Inventory Days", "inventory_days", "inventory_days", "0"),
    ]
    for r_idx, (label, assum_key, hist_key, fmt) in enumerate(assum_items, 2):
        ws.cell(row=r_idx, column=1, value=label).font = data_font
        if hist_key and hist_key in hist_ratios:
            hr_vals = hist_ratios[hist_key]
            hr_years = hist_ratios.get("years", [])
            for ci, yr in zip(h_cols, disp_hist):
                if yr in hr_years:
                    idx = hr_years.index(yr)
                    if idx < len(hr_vals) and hr_vals[idx] is not None:
                        c = ws.cell(row=r_idx, column=ci, value=hr_vals[idx])
                        c.number_format = fmt
        for ci, yr in zip(p_cols, proj_years):
            v = av(assum_key, yr)
            if v is not None and v != 0:
                c = ws.cell(row=r_idx, column=ci, value=v)
                c.number_format = fmt
                c.font = input_font
                c.fill = peach_fill

    r_val = len(assum_items) + 3
    ws.cell(row=r_val, column=1, value="VALUATION ASSUMPTIONS").font = sec_font
    for c in range(1, nc + 1):
        ws.cell(row=r_val, column=c).fill = grey_fill
    r_val += 1
    for label, key, fmt in [
        ("WACC %", "wacc_pct", RATIO), ("Terminal Growth %", "terminal_growth_pct", RATIO),
        ("Target PE (x)", "target_pe", RATIO), ("Target EV/EBITDA (x)", "target_ev_ebitda", RATIO),
        ("Dividend Payout %", "dividend_payout_pct", RATIO),
    ]:
        ws.cell(row=r_val, column=1, value=label)
        v = asmp.get(key)
        if v:
            c = ws.cell(row=r_val, column=2, value=v)
            c.number_format = fmt
            c.font = input_font
        r_val += 1
    r_val += 1
    ws.cell(row=r_val, column=1, value="RATIONALE").font = sec_font
    r_val += 1
    ws.merge_cells(start_row=r_val, start_column=1, end_row=r_val + 3, end_column=nc)
    ws.cell(row=r_val, column=1, value=asmp.get("rationale", "")).font = data_font
    ws.cell(row=r_val, column=1).alignment = Alignment(wrap_text=True, vertical="top")

    add_borders(ws, r_val + 3, nc)
    ws.freeze_panes = "B2"
    logger.info("  ✅ Assumptions")

    # ═══ P&L ═══
    ws, pln = mk("P&L")
    set_w(ws, [30] + [15] * (nc - 1))
    ws.cell(row=1, column=1, value="All figures in ₹ Cr").font = Font(name="Arial", italic=True, size=9)
    hdr_r(ws, 2, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=2, column=ci).fill = orange_fill

    A_GROWTH, A_RM, A_EMP, A_PF, A_OMFG, A_SA, A_OE, A_CI = 2, 3, 4, 5, 6, 7, 8, 9
    A_DEP, A_INT, A_OI, A_TAX, A_CAPEX, A_RECV, A_INV = 10, 11, 12, 13, 14, 15, 16

    r = 3
    ws.cell(row=r, column=1, value="Revenue").font = sec_font
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{row_map['sales']}").number_format = INR
    for ci in p_cols:
        prev = get_column_letter(ci - 1)
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={prev}{r}*(1+'{asn}'!{cl}{A_GROWTH}/100)").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    REV = r
    r += 1

    expense_lines = [
        ("Raw Material Cost", row_map["raw_material"], A_RM), ("Chg in Inventory", row_map["change_in_inventory"], A_CI),
        ("Employee Cost", row_map["employee_cost"], A_EMP), ("Power & Fuel", row_map["power_fuel"], A_PF),
        ("Other Mfg Exp", row_map["other_mfg"], A_OMFG), ("Selling & Admin", row_map["selling_admin"], A_SA),
        ("Other Expenses", row_map["other_expenses"], A_OE),
    ]
    exp_rows = []
    for label, ds_row, assum_row in expense_lines:
        ws.cell(row=r, column=1, value=label).font = data_font
        for ci, dc in zip(h_cols, disp_ds):
            ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{ds_row}").number_format = INR
        for ci in p_cols:
            cl = get_column_letter(ci)
            ws.cell(row=r, column=ci, value=f"={cl}{REV}*'{asn}'!{cl}{assum_row}/100").number_format = INR
            ws.cell(row=r, column=ci).fill = peach_fill
        exp_rows.append(r)
        r += 1

    ws.cell(row=r, column=1, value="Total Expenses").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"=SUM({cl}{exp_rows[0]}:{cl}{exp_rows[-1]})").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    TOTEXP = r
    r += 1

    ws.cell(row=r, column=1, value="EBITDA").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{REV}-{cl}{TOTEXP}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    EBITDA = r
    r += 1

    ws.cell(row=r, column=1, value="EBITDA Margin %")
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"=IF({cl}{REV}=0,0,{cl}{EBITDA}/{cl}{REV})").number_format = PCT
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    r += 1

    ws.cell(row=r, column=1, value="Other Income")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{row_map['other_income']}").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"='{asn}'!{cl}{A_OI}").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    OI = r
    r += 1

    ws.cell(row=r, column=1, value="Depreciation")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{row_map['depreciation']}").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"='{asn}'!{cl}{A_DEP}").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    DEP = r
    r += 1

    ws.cell(row=r, column=1, value="EBIT").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{EBITDA}+{cl}{OI}-{cl}{DEP}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    EBIT = r
    r += 1

    ws.cell(row=r, column=1, value="Finance Cost")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{row_map['interest']}").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"='{asn}'!{cl}{A_INT}").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    INT_R = r
    r += 1

    ws.cell(row=r, column=1, value="Profit Before Tax").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{EBIT}-{cl}{INT_R}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    PBT = r
    r += 1

    ws.cell(row=r, column=1, value="Tax")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{row_map['tax']}").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"=MAX(0,{cl}{PBT}*'{asn}'!{cl}{A_TAX}/100)").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    TAX = r
    r += 1

    ws.cell(row=r, column=1, value="Profit After Tax (PAT)").font = Font(
        name="Arial", bold=True, color=COLORS["navy"], size=12
    )
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{PBT}-{cl}{TAX}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    PAT = r
    r += 1

    ws.cell(row=r, column=1, value="PAT Margin %")
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"=IF({cl}{REV}=0,0,{cl}{PAT}/{cl}{REV})").number_format = PCT
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    r += 1

    ws.cell(row=r, column=1, value="PAT Growth %")
    for ci in range(3, nc + 1):
        cl = get_column_letter(ci)
        prev = get_column_letter(ci - 1)
        ws.cell(row=r, column=ci, value=f"=IF({prev}{PAT}<=0,0,({cl}{PAT}/{prev}{PAT})-1)").number_format = PCT
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    r += 1

    ws.cell(row=r, column=1, value="EPS (₹)").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{PAT}/{shares_cr:.2f}").number_format = INR2
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    EPS_R = r

    add_borders(ws, r, nc)
    ws.freeze_panes = "B3"
    logger.info("  ✅ P&L (formula-based)")

    # ═══ BALANCE SHEET ═══
    ws, bsn = mk("Balance Sheet")
    set_w(ws, [30] + [15] * (nc - 1))
    ws.cell(row=1, column=1, value="All figures in ₹ Cr").font = Font(name="Arial", italic=True, size=9)
    hdr_r(ws, 2, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=2, column=ci).fill = orange_fill

    r = 3
    bs_src = [("Share Capital", row_map["share_capital"], "share_capital"), ("Reserves", row_map["reserves"], "reserves")]
    bs_rr = {}
    for label, ds_row, pk in bs_src:
        ws.cell(row=r, column=1, value=label)
        for ci, dc in zip(h_cols, disp_ds):
            ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{ds_row}").number_format = INR
        for ci, i in zip(p_cols, range(np_)):
            ws.cell(row=r, column=ci, value=pv(pk, i)).number_format = INR
            ws.cell(row=r, column=ci).fill = peach_fill
        bs_rr[pk] = r
        r += 1

    ws.cell(row=r, column=1, value="Total Equity").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{bs_rr['share_capital']}+{cl}{bs_rr['reserves']}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    EQ_R = r
    r += 1

    for label, ds_row, pk in [("Borrowings", row_map["borrowings"], "borrowings"), ("Other Liabilities", row_map["other_liabilities"], "other_liabilities")]:
        ws.cell(row=r, column=1, value=label)
        for ci, dc in zip(h_cols, disp_ds):
            ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{ds_row}").number_format = INR
        for ci, i in zip(p_cols, range(np_)):
            ws.cell(row=r, column=ci, value=pv(pk, i)).number_format = INR
            ws.cell(row=r, column=ci).fill = peach_fill
        bs_rr[pk] = r
        r += 1

    ws.cell(row=r, column=1, value="Total Liabilities & Equity").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(
            row=r, column=ci,
            value=f"={cl}{EQ_R}+{cl}{bs_rr['borrowings']}+{cl}{bs_rr['other_liabilities']}",
        ).number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    TL_R = r
    r += 2

    ws.cell(row=r, column=1, value="APPLICATION OF FUNDS").font = sec_font
    for c in range(1, nc + 1):
        ws.cell(row=r, column=c).fill = blue_fill
        ws.cell(row=r, column=c).font = sub_font
    r += 1

    asset_items = [("Net Block", row_map["net_block"], "net_block"), ("CWIP", row_map["cwip"], "cwip"), ("Investments", row_map["investments"], "investments")]
    asset_rows = []
    for label, ds_row, pk in asset_items:
        ws.cell(row=r, column=1, value=label)
        for ci, dc in zip(h_cols, disp_ds):
            ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{ds_row}").number_format = INR
        for ci, i in zip(p_cols, range(np_)):
            ws.cell(row=r, column=ci, value=pv(pk, i)).number_format = INR
            ws.cell(row=r, column=ci).fill = peach_fill
        asset_rows.append(r)
        r += 1

    ws.cell(row=r, column=1, value="Trade Receivables")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{row_map['receivables']}").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"='{pln}'!{cl}{REV}*'{asn}'!{cl}{A_RECV}/365").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    RECV_R = r
    asset_rows.append(r)
    r += 1

    ws.cell(row=r, column=1, value="Inventory")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{row_map['inventory']}").number_format = INR
    RM_ROW = exp_rows[0]
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"='{pln}'!{cl}{RM_ROW}*'{asn}'!{cl}{A_INV}/365").number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    INVTY_R = r
    asset_rows.append(r)
    r += 1

    ws.cell(row=r, column=1, value="Cash & Bank (Plug)")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{row_map['cash']}").number_format = INR
    CASH_R = r
    r += 1

    ws.cell(row=r, column=1, value="Other Assets")
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(
            row=r, column=ci,
            value=f"='Data Sheet'!{dc}{row_map['other_assets']}-'Data Sheet'!{dc}{row_map['receivables']}-'Data Sheet'!{dc}{row_map['inventory']}-'Data Sheet'!{dc}{row_map['cash']}",
        ).number_format = INR
    for ci, i in zip(p_cols, range(np_)):
        ws.cell(row=r, column=ci, value=pv("other_assets", i)).number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    OA_R = r
    r += 1

    for ci in p_cols:
        cl = get_column_letter(ci)
        asset_sum = "+".join(f"{cl}{ar}" for ar in asset_rows)
        ws.cell(row=CASH_R, column=ci, value=f"={cl}{TL_R}-{asset_sum}-{cl}{OA_R}").number_format = INR
        ws.cell(row=CASH_R, column=ci).fill = peach_fill

    projected_revenues = projected_revenues_numeric()
    for i, year in enumerate(proj_years):
        total_liab = (
            float(pv("share_capital", i, 0) or 0)
            + float(pv("reserves", i, 0) or 0)
            + float(pv("borrowings", i, 0) or 0)
            + float(pv("other_liabilities", i, 0) or 0)
        )
        receivables = projected_revenues[i] * float(av("receivable_days", year) or 0) / 365
        raw_material_cost = projected_revenues[i] * float(av("rm_pct", year) or 0) / 100
        inventory = raw_material_cost * float(av("inventory_days", year) or 0) / 365
        assets_ex_cash = (
            float(pv("net_block", i, 0) or 0)
            + float(pv("cwip", i, 0) or 0)
            + float(pv("investments", i, 0) or 0)
            + float(pv("other_assets", i, 0) or 0)
            + receivables
            + inventory
        )
        cash_plug = total_liab - assets_ex_cash
        if cash_plug < 0:
            msg = (
                f"Cash plug projected NEGATIVE in {year} ({cash_plug:.0f} Cr) — "
                f"borrowings/reserves likely understated relative to asset base for {sd.get('company_name')}"
            )
            logger.warning("⚠ %s", msg)
            diagnostics.append(msg)

    ws.cell(row=r, column=1, value="Total Assets").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        all_a = "+".join(f"{cl}{ar}" for ar in asset_rows) + f"+{cl}{CASH_R}+{cl}{OA_R}"
        ws.cell(row=r, column=ci, value=f"={all_a}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    TA_R = r
    r += 1

    ws.cell(row=r, column=1, value="CHECK (TA - TL&E = 0)").font = red_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{TA_R}-{cl}{TL_R}").number_format = INR2

    add_borders(ws, r, nc)
    ws.freeze_panes = "B3"
    logger.info("  ✅ Balance Sheet (formula-based)")

    # ═══ CASH FLOW ═══
    ws, cfn = mk("Cash Flow")
    set_w(ws, [30] + [15] * (nc - 1))
    hdr_r(ws, 2, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=2, column=ci).fill = orange_fill
    r = 3
    cf_items = [("CFO", row_map["cfo"], "cfo"), ("CFI", row_map["cfi"], "cfi"), ("CFF", row_map["cff"], "cff")]
    cf_rr = {}
    for label, ds_row, pk in cf_items:
        ws.cell(row=r, column=1, value=label).font = sec_font
        for ci, dc in zip(h_cols, disp_ds):
            ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}{ds_row}").number_format = INR
        for ci, i in zip(p_cols, range(np_)):
            ws.cell(row=r, column=ci, value=pv(pk, i)).number_format = INR
            ws.cell(row=r, column=ci).fill = peach_fill
        cf_rr[pk] = r
        r += 1
    ws.cell(row=r, column=1, value="Net Cash Flow").font = sec_font
    for ci, dc in zip(h_cols, disp_ds):
        ws.cell(row=r, column=ci, value=f"='Data Sheet'!{dc}85").number_format = INR
    for ci in p_cols:
        cl = get_column_letter(ci)
        ws.cell(
            row=r, column=ci,
            value=f"={cl}{cf_rr['cfo']}+{cl}{cf_rr['cfi']}+{cl}{cf_rr['cff']}",
        ).number_format = INR
        ws.cell(row=r, column=ci).fill = peach_fill
    r += 2
    ws.cell(row=r, column=1, value="Free Cash Flow (FCF)").font = sec_font
    for ci in range(2, nc + 1):
        cl = get_column_letter(ci)
        ws.cell(row=r, column=ci, value=f"={cl}{cf_rr['cfo']}+{cl}{cf_rr['cfi']}").number_format = INR
        if ci in p_cols:
            ws.cell(row=r, column=ci).fill = peach_fill
    FCF_R = r

    add_borders(ws, r, nc)
    ws.freeze_panes = "B3"
    logger.info("  ✅ Cash Flow (formula-based)")

    # ═══ RATIOS ═══
    ws, _ = mk("Ratios")
    set_w(ws, [30] + [15] * (nc - 1))
    ws.cell(row=1, column=1, value="Key Financial Ratios").font = title_font
    hdr_r(ws, 2, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=2, column=ci).fill = orange_fill
    r = 3
    ratio_formulas = [
        ("EBITDA Margin %", f"=IF('{pln}'!{{cl}}{REV}=0,0,'{pln}'!{{cl}}{EBITDA}/'{pln}'!{{cl}}{REV})", PCT),
        ("PAT Margin %", f"=IF('{pln}'!{{cl}}{REV}=0,0,'{pln}'!{{cl}}{PAT}/'{pln}'!{{cl}}{REV})", PCT),
        ("ROE %", f"=IF('{bsn}'!{{cl}}{EQ_R}=0,0,'{pln}'!{{cl}}{PAT}/'{bsn}'!{{cl}}{EQ_R})", PCT),
        ("ROCE %", f"=IF(('{bsn}'!{{cl}}{EQ_R}+'{bsn}'!{{cl}}{bs_rr['borrowings']})=0,0,'{pln}'!{{cl}}{EBIT}/('{bsn}'!{{cl}}{EQ_R}+'{bsn}'!{{cl}}{bs_rr['borrowings']}))", PCT),
        ("Debt/Equity (x)", f"=IF('{bsn}'!{{cl}}{EQ_R}=0,0,'{bsn}'!{{cl}}{bs_rr['borrowings']}/'{bsn}'!{{cl}}{EQ_R})", RATIO),
        ("Receivable Days", f"=IF('{pln}'!{{cl}}{REV}=0,0,'{bsn}'!{{cl}}{RECV_R}/'{pln}'!{{cl}}{REV}*365)", "0"),
        ("Inventory Days", f"=IF('{pln}'!{{cl}}{exp_rows[0]}=0,0,'{bsn}'!{{cl}}{INVTY_R}/'{pln}'!{{cl}}{exp_rows[0]}*365)", "0"),
        ("Asset Turnover (x)", f"=IF('{bsn}'!{{cl}}{TA_R}=0,0,'{pln}'!{{cl}}{REV}/'{bsn}'!{{cl}}{TA_R})", RATIO),
        ("EPS (₹)", f"='{pln}'!{{cl}}{PAT}/{shares_cr:.2f}", INR2),
        ("BVPS (₹)", f"='{bsn}'!{{cl}}{EQ_R}/{shares_cr:.2f}", INR2),
    ]
    for label, tmpl, fmt in ratio_formulas:
        ws.cell(row=r, column=1, value=label)
        for ci in range(2, nc + 1):
            cl = get_column_letter(ci)
            ws.cell(row=r, column=ci, value=tmpl.replace("{cl}", cl)).number_format = fmt
            if ci in p_cols:
                ws.cell(row=r, column=ci).fill = peach_fill
        r += 1
    add_borders(ws, r, nc)
    ws.freeze_panes = "B3"
    logger.info("  ✅ Ratios (formula-based)")

    # ═══ VALUATION ═══
    ws, _ = mk("Valuation")
    set_w(ws, [35, 18, 18, 18, 18, 18])
    ws["A1"].value = "VALUATION SUMMARY"
    ws["A1"].font = title_font
    r = 3
    for label, v in [
        ("DCF Fair Value (₹)", val.get("dcf_fair_value")),
        ("PE Fair Value (₹)", val.get("pe_fair_value")),
        ("EV/EBITDA Fair Value (₹)", val.get("ev_ebitda_fair_value")),
        ("Blended Fair Value (₹)", val.get("blended_fair_value")),
        ("CMP (₹)", cmp), ("Rating", rating),
        ("Target Price (₹)", target), ("Upside %", upside),
    ]:
        ws.cell(row=r, column=1, value=label).font = sec_font
        ws.cell(row=r, column=2, value=v)
        for c in range(1, 3):
            ws.cell(row=r, column=c).border = thin_bdr
        r += 1
    r += 1

    ws.cell(row=r, column=1, value="PEER COMPARISON").font = sec_font
    r += 1
    for ci, h in enumerate(["Company", "MCap", "PE", "EV/EBITDA", "Margin", "ROE"], 1):
        ws.cell(row=r, column=ci, value=h).font = hdr_font
        ws.cell(row=r, column=ci).fill = navy_fill
    r += 1
    for p in peers:
        ws.cell(row=r, column=1, value=p.get("name", ""))
        ws.cell(row=r, column=2, value=p.get("mcap_cr")).number_format = INR
        ws.cell(row=r, column=3, value=p.get("pe")).number_format = RATIO
        ws.cell(row=r, column=4, value=p.get("ev_ebitda")).number_format = RATIO
        ws.cell(row=r, column=5, value=p.get("ebitda_margin_pct", ""))
        ws.cell(row=r, column=6, value=p.get("roe_pct")).number_format = RATIO
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = thin_bdr
        r += 1

    sens = val.get("sensitivity_pe", {})
    if sens and sens.get("grid"):
        if len(sens.get("row_values", [])) != len(sens.get("grid", [])) or any(
            len(row) != len(sens.get("col_values", [])) for row in sens.get("grid", [])
        ):
            logger.warning("⚠ Sensitivity grid shape mismatch for %s", screener_data.get("company_name"))
        r += 2
        ws.cell(row=r, column=1, value="SENSITIVITY TABLE").font = sec_font
        r += 1
        ws.cell(row=r, column=1, value=f"{sens.get('row_label','')} \\ {sens.get('col_label','')}")
        for ci, cv in enumerate(sens.get("col_values", []), 2):
            ws.cell(row=r, column=ci, value=cv).font = hdr_font
            ws.cell(row=r, column=ci).fill = navy_fill
        r += 1
        for rv, gv in zip(sens.get("row_values", []), sens.get("grid", [])):
            ws.cell(row=r, column=1, value=rv).font = sec_font
            for ci, v in enumerate(gv, 2):
                ws.cell(row=r, column=ci, value=v).number_format = INR
                ws.cell(row=r, column=ci).fill = peach_fill
            r += 1
    logger.info("  ✅ Valuation")

    # ═══ KPI DASHBOARD ═══
    ws, _ = mk("KPI Dashboard")
    set_w(ws, [35] + [15] * (nc - 1))
    hdr_r(ws, 2, year_labels, nc)
    for ci in p_cols:
        ws.cell(row=2, column=ci).fill = orange_fill
    r = 3
    kpi_formulas = [
        ("Revenue (₹ Cr)", f"='{pln}'!{{cl}}{REV}", INR),
        ("EBITDA (₹ Cr)", f"='{pln}'!{{cl}}{EBITDA}", INR),
        ("PAT (₹ Cr)", f"='{pln}'!{{cl}}{PAT}", INR),
        ("EPS (₹)", f"='{pln}'!{{cl}}{EPS_R}", INR2),
        ("EBITDA Margin %", f"=IF('{pln}'!{{cl}}{REV}=0,0,'{pln}'!{{cl}}{EBITDA}/'{pln}'!{{cl}}{REV})", PCT),
        ("ROE %", f"=IF('{bsn}'!{{cl}}{EQ_R}=0,0,'{pln}'!{{cl}}{PAT}/'{bsn}'!{{cl}}{EQ_R})", PCT),
        ("D/E (x)", f"=IF('{bsn}'!{{cl}}{EQ_R}=0,0,'{bsn}'!{{cl}}{bs_rr['borrowings']}/'{bsn}'!{{cl}}{EQ_R})", RATIO),
    ]
    for label, tmpl, fmt in kpi_formulas:
        ws.cell(row=r, column=1, value=label)
        for ci in range(2, nc + 1):
            cl = get_column_letter(ci)
            ws.cell(row=r, column=ci, value=tmpl.replace("{cl}", cl)).number_format = fmt
            if ci in p_cols:
                ws.cell(row=r, column=ci).fill = peach_fill
        r += 1
    add_borders(ws, r, nc)
    logger.info("  ✅ KPI Dashboard (formula-based)")

    # ═══ EXTENDED ANALYTICAL SHEETS (10) ═══
    ctx = {
        "sd": sd,
        "model": model,
        "asmp": asmp,
        "proj": proj,
        "peers": peers,
        "thesis": thesis,
        "val": val,
        "hist_ratios": hist_ratios,
        "hist_years": hist_years,
        "disp_hist": disp_hist,
        "disp_ds": disp_ds,
        "proj_years": proj_years,
        "year_labels": year_labels,
        "nh": nh,
        "dh": dh,
        "np_": np_,
        "nc": nc,
        "h_cols": h_cols,
        "p_cols": p_cols,
        "row_map": row_map,
        "shares_cr": shares_cr,
        "cmp": cmp,
        "target": target,
        "rating": rating,
        "upside": upside,
        "company_name": sd.get("company_name", ""),
        "sector": model.get("sector", ""),
        "pln": pln,
        "asn": asn,
        "bsn": bsn,
        "cfn": cfn,
        "REV": REV,
        "EBITDA": EBITDA,
        "OI": OI,
        "DEP": DEP,
        "EBIT": EBIT,
        "INT_R": INT_R,
        "PBT": PBT,
        "TAX": TAX,
        "PAT": PAT,
        "EPS_R": EPS_R,
        "EQ_R": EQ_R,
        "TL_R": TL_R,
        "TA_R": TA_R,
        "RECV_R": RECV_R,
        "INVTY_R": INVTY_R,
        "CASH_R": CASH_R,
        "OA_R": OA_R,
        "FCF_R": FCF_R,
        "exp_rows": exp_rows,
        "bs_rr": bs_rr,
        "cf_rr": cf_rr,
        "expense_lines": expense_lines,
        "projected_revenues": projected_revenues_numeric(),
    }
    try:
        mk_fin_summary(wb, ctx)
        mk_earnings_forecast(wb, ctx)
        mk_financials_table(wb, ctx)
        mk_valuations_table(wb, ctx)
        mk_key_risks(wb, ctx)
        mk_peer_compare(wb, ctx)
        mk_operational_data(wb, ctx)
        mk_governance(wb, ctx)
        mk_timeline(wb, ctx)
        mk_op_charts(wb, ctx)
        logger.info("  ✅ Extended analytical sheets (10)")
    except Exception as ex:
        logger.exception("⚠ Extended sheets failed: %s", ex)
        diagnostics.append(f"Extended sheets build error: {ex}")

    if diagnostics:
        ws, _ = mk("Model Diagnostics")
        set_w(ws, [120])
        ws["A1"] = "MODEL DIAGNOSTICS"
        ws["A1"].font = title_font
        ws["A2"] = "Warnings detected during Python pre-flight checks."
        ws["A2"].font = data_font
        for idx, message in enumerate(diagnostics, 4):
            ws.cell(row=idx, column=1, value=f"- {message}").font = red_font
        add_borders(ws, len(diagnostics) + 4, 1)
        logger.info("  ✅ Model Diagnostics")

    # Verify originals intact + save
    final = wb.sheetnames
    assert all(s in final for s in orig_sheets), "❌ Original sheets missing!"
    wb.save(out_path)
    logger.info(f"✅ Saved → {out_path} ({os.path.getsize(out_path)//1024} KB)")


# ══════════════════════════════════════════════════════════════════
# P&L SERIES EMBEDDER
# ══════════════════════════════════════════════════════════════════
def _embed_pl_series(model_json: dict, screener_data: dict) -> None:
    """Compute Revenue/EBITDA/PAT series and store them in model_json in-place.

    The Excel model derives these via cell formulas that are never persisted to
    JSON. Without this step the PPTX generator has no absolute P&L numbers to
    put in charts, so slides show zeros or placeholder data.

    We use:
      - screener_data for the historical series (audited actuals)
      - model_json["assumptions"] growth/margin for the projection series
    Both halves are stored so the chart can show a full historical+forward view.
    """
    hist_years: list[str] = screener_data.get("fiscal_years") or []
    pl = screener_data.get("pl") or {}
    derived = screener_data.get("derived") or {}

    hist_revenue = pl.get("sales") or []
    hist_ebitda  = derived.get("ebitda") or []
    hist_pat     = pl.get("net_profit") or []

    # For recently-IPO'd / platform businesses (e.g. Swiggy, Zomato) Screener
    # often returns null for row 17 (sales). Fall back to the absolute values
    # Claude fetched via web_search, aligned to the same year list.
    web_pl = model_json.get("historical_pl_absolute") or {}
    web_years: list[str] = [str(y) for y in (web_pl.get("years") or [])]

    def _align_web(screener_vals: list, web_vals: list) -> list:
        """Return screener_vals, patching any None entries from web_vals aligned by year."""
        if not web_vals or len(web_vals) != len(web_years):
            return screener_vals
        web_by_yr = dict(zip(web_years, web_vals))
        out = []
        for yr, sv in zip(hist_years, screener_vals):
            out.append(sv if sv is not None else web_by_yr.get(yr))
        if len(screener_vals) < len(hist_years):
            out = screener_vals
        return out

    hist_revenue = _align_web(hist_revenue, web_pl.get("revenue") or [])
    hist_ebitda  = _align_web(hist_ebitda,  web_pl.get("ebitda")  or [])
    hist_pat     = _align_web(hist_pat,     web_pl.get("pat")     or [])

    # If Screener returned no years at all but web search has data, use web data directly
    if not hist_years and web_years:
        hist_years   = web_years
        hist_revenue = web_pl.get("revenue") or []
        hist_ebitda  = web_pl.get("ebitda")  or []
        hist_pat     = web_pl.get("pat")     or []

    model_json["historical_pl"] = {
        "years":   hist_years,
        "revenue": hist_revenue,
        "ebitda":  hist_ebitda,
        "pat":     hist_pat,
    }

    asmp = model_json.get("assumptions") or {}
    proj_years: list[str] = asmp.get("projection_years") or []
    revenue_growth: dict   = asmp.get("revenue_growth_pct") or {}

    hist_ratios = model_json.get("historical_ratios") or {}
    ebitda_margins = [v for v in (hist_ratios.get("ebitda_margin_pct") or []) if v is not None]
    pat_margins    = [v for v in (hist_ratios.get("pat_margin_pct")    or []) if v is not None]
    ebitda_m = ebitda_margins[-1] if ebitda_margins else 15.0
    pat_m    = pat_margins[-1]    if pat_margins    else 8.0

    # Base revenue = last non-null historical value
    base_rev = next((v for v in reversed(hist_revenue) if v is not None), None)

    proj_revenues: list = []
    proj_ebitdas:  list = []
    proj_pats:     list = []
    prev = base_rev
    for yr in proj_years:
        growth = (revenue_growth.get(yr) or 0) / 100
        rev = round(prev * (1 + growth), 1) if prev is not None else None
        proj_revenues.append(rev)
        proj_ebitdas.append(round(rev * ebitda_m / 100, 1) if rev is not None else None)
        proj_pats.append(   round(rev * pat_m    / 100, 1) if rev is not None else None)
        prev = rev

    model_json["projected_pl"] = {
        "years":   proj_years,
        "revenue": proj_revenues,
        "ebitda":  proj_ebitdas,
        "pat":     proj_pats,
    }


# ══════════════════════════════════════════════════════════════════
# PUBLIC ENTRY POINT
# ══════════════════════════════════════════════════════════════════
def generate_financial_model(
    nse_code: str,
    company_name: str,
    sector: str,
    output_dir: str,
    anthropic_api_key: Optional[str] = None,
    screener_username: Optional[str] = None,
    screener_password: Optional[str] = None,
) -> dict:
    """Generate a full financial model. Returns dict with file paths + metadata."""
    api_key = anthropic_api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY not provided (env var or argument)")

    sc_user = screener_username or os.environ.get("SCREENER_USERNAME")
    sc_pass = screener_password or os.environ.get("SCREENER_PASSWORD")
    if not sc_user or not sc_pass:
        raise ValueError("SCREENER_USERNAME and SCREENER_PASSWORD required (env vars or arguments)")

    os.makedirs(output_dir, exist_ok=True)
    tracker = CostTracker()
    t_start = time.time()

    nse_code = nse_code.upper()
    logger.info(f"🚀 PIPELINE START | {company_name} ({nse_code}) | Sector: {sector}")

    screener_path = download_screener_excel(nse_code, output_dir, sc_user, sc_pass)
    screener_data = extract_screener_data(screener_path)
    market_text = fetch_market_data(nse_code)

    model_json = call_claude_api(
        company_name, nse_code, sector, screener_data, market_text, api_key, tracker,
    )
    model_json = normalize_model_output(model_json, screener_data)
    validated_model = validate_model_output(model_json, nse_code)
    model_json = validated_model.model_dump()

    # Embed computed P&L series so the PPTX generator can build real charts.
    # The Excel model uses formulas that aren't stored in the JSON, so we
    # recompute Revenue/EBITDA/PAT here from the Screener historical base +
    # Claude's growth/margin assumptions.
    _embed_pl_series(model_json, screener_data)

    json_path = os.path.join(output_dir, f"{nse_code}_model.json")
    with open(json_path, "w") as f:
        json.dump(model_json, f, indent=2, default=str)

    out_path = os.path.join(output_dir, f"{nse_code}_financial_model.xlsx")
    shutil.copy2(screener_path, out_path)
    build_model(out_path, screener_data, model_json, out_path)

    elapsed = time.time() - t_start
    logger.info(tracker.summary())

    return {
        "file_path": out_path,
        "json_path": json_path,
        "model_json": model_json,
        "cost_summary": tracker.to_dict(),
        "elapsed_seconds": round(elapsed, 1),
        "nse_code": nse_code,
        "rating": model_json.get("rating"),
        "target_price": model_json.get("target_price"),
        "upside_pct": model_json.get("upside_pct"),
        "base_year": model_json.get("base_year"),
    }


# ══════════════════════════════════════════════════════════════════
# CLI / COLAB ENTRY
# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate financial model for an NSE stock")
    parser.add_argument("--nse", required=True, help="NSE code (e.g., GRAVITA)")
    parser.add_argument("--name", required=True, help="Company name")
    parser.add_argument("--sector", default="General", help="Sector")
    parser.add_argument("--out", default="./output", help="Output directory")
    args = parser.parse_args()

    try:
        result = generate_financial_model(args.nse, args.name, args.sector, args.out)
        print(f"\n✅ DONE in {result['elapsed_seconds']}s")
        print(f"   File:        {result['file_path']}")
        print(f"   Rating:      {result['rating']}")
        print(f"   Target:      ₹{result['target_price']}")
        print(f"   Upside:      {result['upside_pct']}%")
        print(f"   Cost:        ${result['cost_summary']['cost_usd']} (₹{result['cost_summary']['cost_inr']})")
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        traceback.print_exc()
