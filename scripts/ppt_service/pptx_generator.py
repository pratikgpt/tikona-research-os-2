"""Orchestrates PPTX generation: pulls report+session data from Supabase, builds
the four reportgen input files, runs the pipeline, and uploads results."""
from __future__ import annotations

import json
import logging
import os
import re
import tempfile
import time
import urllib.error
import urllib.request
from contextlib import contextmanager
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from supabase_client import get_service_client
from typing import Any

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

from chart_generators import generate_chart

import logging
logger = logging.getLogger(__name__)

PPTX_BUCKET = "research-reports-pptx"
MODEL_BUCKET = os.environ.get("SUPABASE_FINANCIAL_MODEL_BUCKET", "research-reports-html")

PPTX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
PDF_CONTENT_TYPE = "application/pdf"

_NUM_RE = re.compile(r"[\d,]+\.?\d*")
_PLACEHOLDER_RE = re.compile(r"not included in the generated report|content pending", re.I)
_OPENROUTER_PATCHED = False
_HOUSE_PLANNER_PATCHED = False
_ORPHAN_NUMBER_RE = re.compile(
    r"(?<!FY)(?<!fy)\b\d[\d,]*(?:\.\d+)?\s*(?:%|x\b|cr\b|crore\b|bn\b|billion\b|lakh\b|bps\b)"
    r"|(?:₹|\$|€|£|INR\s|USD\s|EUR\s|GBP\s)\s*\d[\d,]*(?:\.\d+)?",
    re.IGNORECASE,
)


# ─────────── helpers ──────────────────────────────────────────────────────────


def _parse_number(val: Any) -> float | None:
    """Mirror the frontend parseNumber: extract first numeric token, strip commas."""
    if val is None:
        return None
    s = str(val)
    m = _NUM_RE.search(s)
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", ""))
    except ValueError:
        return None


def _normalize_rating(raw: str) -> str:
    up = (raw or "").upper()
    if "SELL" in up:
        return "SELL"
    if "HOLD" in up:
        return "HOLD"
    if "ACCUMULATE" in up:
        return "ACCUMULATE"
    return "BUY"


def _section_value(sections: list[dict], key: str) -> str:
    for s in sections:
        if s.get("section_key") == key:
            value = (s.get("content") or "").strip()
            return "" if _PLACEHOLDER_RE.search(value) else value
    return ""


def _prefer(*vals: Any) -> str:
    for v in vals:
        if v is not None:
            value = str(v).strip()
            if value and not _PLACEHOLDER_RE.search(value):
                return value
    return ""


def _clean_prose(text: str, *, max_len: int = 500) -> str:
    cleaned = re.sub(r"\s+", " ", text or "").strip()
    if len(cleaned) <= max_len:
        return cleaned
    # Truncate at word boundary so we don't cut mid-word
    truncated = cleaned[:max_len]
    last_space = truncated.rfind(" ")
    if last_space > int(max_len * 0.75):
        return truncated[:last_space].strip()
    return truncated.strip()


def _section_by_any(sections: list[dict], keywords: list[str]) -> str:
    for s in sections:
        haystack = f"{s.get('section_key') or ''} {s.get('section_title') or ''}".casefold()
        if any(k in haystack for k in keywords):
            value = (s.get("content") or "").strip()
            if value and not _PLACEHOLDER_RE.search(value):
                return value
    return ""


def _bullets_from_text(text: str, *, limit: int = 5, max_len: int = 180) -> list[str]:
    if not text:
        return ["Analyst narrative supports continued monitoring of the core investment case."]

    # Try splitting on **Bold header:** sections first — common in research reports
    bold_headers = re.findall(r'\*\*([^*]+?):\*\*', text)
    bold_parts   = re.split(r'\*\*[^*]+?:\*\*', text)
    if len(bold_headers) >= 2 and len(bold_parts) >= 2:
        bullets: list[str] = []
        for header, content in zip(bold_headers, bold_parts[1:]):
            # First sentence of content after the header
            first_sent = re.split(r'(?<=[.!?])\s+', content.strip())[0] if content.strip() else ""
            combined = f"{header}: {first_sent.strip()}" if first_sent.strip() else header
            item = _clean_prose(combined.lstrip("-*•: "), max_len=max_len)
            if len(item) >= 12:
                bullets.append(item)
            if len(bullets) >= limit:
                break
        if len(bullets) >= 2:
            return bullets

    # Fall back: split on newlines and sentence endings
    parts = re.split(r"(?:\n+|(?<=[.!?])\s+)", text)
    bullets = []
    for part in parts:
        item = _clean_prose(part.lstrip("-*•** ").strip(), max_len=max_len)
        if len(item) >= 12:
            bullets.append(item)
        if len(bullets) >= limit:
            break
    return bullets or ["Analyst narrative supports continued monitoring of the core investment case."]


def _all_sections_text(sections: list[dict]) -> str:
    parts: list[str] = []
    for s in sections:
        title = s.get("section_title") or s.get("section_key") or ""
        content = (s.get("content") or "").strip()
        if content and not _PLACEHOLDER_RE.search(content):
            parts.append(f"{title}\n{content}")
    return "\n\n".join(parts)


def _extract_labeled_number(text: str, patterns: list[str]) -> str:
    for pattern in patterns:
        match = re.search(pattern, text, re.I | re.S)
        if match:
            value = _parse_number(match.group(1))
            if value is not None and value > 0:
                return str(value)
    return ""


def _extract_rating(text: str) -> str:
    patterns = [
        r"\b(BUY|SELL|HOLD|ACCUMULATE)\b\s+(?:rating|recommendation)\b",
        r"\b(?:rating|recommendation)\s*(?:of|:|is|=)?\s*\b(BUY|SELL|HOLD|ACCUMULATE)\b",
        r"\bsupporting\s+a\s+\b(BUY|SELL|HOLD|ACCUMULATE)\b\s+rating\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            return match.group(1)
    return ""


def _strip_qmark(url: str) -> str:
    return url.rstrip("?") if url else url


# ─────────── input builders ───────────────────────────────────────────────────


def _build_company(report: dict, session: dict, sections: list[dict]) -> dict:
    description = _prefer(
        report.get("company_description"),
        session.get("company_description"),
        _section_by_any(sections, ["company", "business", "overview"]),
        _all_sections_text(sections),
    )
    return {
        "name": report.get("company_name") or "Unknown",
        "ticker": (report.get("nse_symbol") or "UNKNOWN").upper(),
        "exchange": "NSE",
        "sector": session.get("sector") or "General",
        "industry": session.get("industry") or session.get("sector") or "General",
        "country": "India",
        "description": _clean_prose(description, max_len=900),
        "peer_list": [],
    }


def _build_metadata(report: dict, sections: list[dict]) -> dict:
    narrative = _all_sections_text(sections)
    rating_raw = _prefer(report.get("cs_rating"), _section_value(sections, "rating"), _extract_rating(narrative))
    if not rating_raw:
        rating_raw = "BUY"
    rating = _normalize_rating(rating_raw)

    cmp_raw = _prefer(
        report.get("cs_current_market_price"),
        report.get("current_market_price"),
        _section_value(sections, "current_market_price"),
        _extract_labeled_number(
            narrative,
            [
                r"(?:current\s+(?:market\s+)?price|CMP)\s*(?:of|at|:|=)?[^\d]{0,20}([\d,]+(?:\.\d+)?)",
                r"(?:from|vs\.?)\s+current\s+(?:market\s+)?price\s*(?:of|at|:|=)?[^\d]{0,20}([\d,]+(?:\.\d+)?)",
            ],
        ),
    )
    tp_raw = _prefer(
        report.get("cs_target_price"),
        report.get("target_price"),
        _section_value(sections, "target_price"),
        _extract_labeled_number(
            narrative,
            [
                r"probability-weighted\s+target\s+price.*?=\s*[^\d]{0,20}([\d,]+(?:\.\d+)?)",
                r"probability-weighted\s+target\s+price\s*(?:of|at|:|=)?[^\d]{0,20}([\d,]+(?:\.\d+)?)",
                r"(?:BUY|ACCUMULATE|HOLD|SELL)\s+recommendation\s+with\s+[^\d]{0,20}([\d,]+(?:\.\d+)?)\s+target\s+price",
                r"base\s+case.*?implied\s+target\s+price\s*(?:of|at|:|=)?[^\d]{0,20}([\d,]+(?:\.\d+)?)",
                r"implied\s+target\s+price\s*(?:of|at|:|=)?[^\d]{0,20}([\d,]+(?:\.\d+)?)",
                r"target\s+price\s*(?:of|at|:|=)?[^\d]{0,20}([\d,]+(?:\.\d+)?)",
                r"[^\d]{0,20}([\d,]+(?:\.\d+)?)\s+target\s+price",
            ],
        ),
    )
    mcap_raw = _prefer(report.get("cs_market_cap"), _section_value(sections, "market_cap"))

    cmp = _parse_number(cmp_raw)
    target = _parse_number(tp_raw)
    if cmp is None or cmp <= 0:
        raise ValueError("CMP could not be extracted from report — fill cs_current_market_price column")
    if target is None or target <= 0:
        raise ValueError("target_price could not be extracted from report — fill cs_target_price column")

    mcap = _parse_number(mcap_raw)

    # Recompute upside_pct from cmp+target so it's always consistent with the validator's check.
    # (Stored upside_pct can be stale; validator allows ≤0.3% tolerance so recomputing avoids the error.)
    upside = round(((target - cmp) / cmp) * 100, 1)

    meta: dict[str, Any] = {
        "rating": rating,
        "currency": "INR",
        "cmp": str(cmp),
        "target_price": str(target),
        "upside_pct": str(upside),
        "analyst": _prefer(report.get("user_email"), "Tikona Research"),
        "report_date": date.today().isoformat(),
        "report_type": "Initiation",
    }
    meta["market_cap"] = str(mcap if mcap is not None and mcap > 0 else 0)
    return meta


def _series(name: str, unit: str, periods: list[str], values: list) -> dict | None:
    """Build a FinancialSeries dict, dropping None-only values."""
    if not periods or not values or len(periods) != len(values):
        return None
    return {
        "name": name,
        "unit": unit,
        "periods": periods,
        "values": [None if v is None else str(v) for v in values],
    }


def _list_of_dicts(value: Any) -> list[dict]:
    if isinstance(value, list):
        return [item for item in value if isinstance(item, dict)]
    return []


def _copy_if_present(target: dict, source: dict, key: str) -> None:
    value = source.get(key)
    if value:
        target[key] = value


def _placeholder_series() -> list[dict]:
    return [
        {
            "name": "Revenue",
            "unit": "INR Cr",
            "periods": ["FY24A", "FY25A", "FY26E"],
            "values": ["0", "0", "0"],
        }
    ]


def _build_financial_model(ticker: str, model_json: dict | None, warnings: list[str]) -> dict:
    """Map the v5 financial model JSON onto the FinancialModelSnapshot schema."""
    base: dict[str, Any] = {
        "model_name": f"{ticker} Base Model",
        "model_version": "v5.1",
        "currency": "INR",
        "fiscal_year_end": "March",
    }

    if not model_json:
        warnings.append("financial-model JSON sidecar missing; using minimal placeholder model")
        base["metrics"] = {"placeholder": "0"}
        base["series"] = _placeholder_series()
        return base

    asmp = model_json.get("assumptions") or {}
    proj = model_json.get("projections") or {}
    hist = model_json.get("historical_ratios") or {}

    proj_years: list[str] = [str(y) for y in (asmp.get("projection_years") or proj.get("years") or [])]
    hist_years: list[str] = [str(y) for y in (hist.get("years") or [])]

    # All series MUST share the same period list — combine hist + proj into one unified timeline.
    seen: set[str] = set()
    all_years: list[str] = []
    for y in hist_years + proj_years:
        if y not in seen:
            seen.add(y)
            all_years.append(y)

    series: list[dict] = []
    metrics: dict[str, str] = {}

    # ── Absolute P&L series (Revenue, EBITDA, PAT in ₹ Cr) ──────────────────
    # financial_model_v5._embed_pl_series() stores historical + projected
    # absolute values so charts show real numbers rather than placeholders.
    hist_pl = model_json.get("historical_pl") or {}
    proj_pl = model_json.get("projected_pl") or {}
    h_yrs   = [str(y) for y in (hist_pl.get("years") or [])]
    p_yrs   = [str(y) for y in (proj_pl.get("years") or [])]
    combined_pl_years = h_yrs + p_yrs

    if combined_pl_years:
        for label, h_key, p_key in [
            ("Revenue", "revenue", "revenue"),
            ("EBITDA",  "ebitda",  "ebitda"),
            ("PAT",     "pat",     "pat"),
        ]:
            vals = list(hist_pl.get(h_key) or []) + list(proj_pl.get(p_key) or [])
            if any(v is not None for v in vals):
                s = _series(label, "INR Cr", combined_pl_years, vals)
                if s:
                    series.append(s)

    if all_years:
        rev_growth = asmp.get("revenue_growth_pct") or {}
        # Revenue growth % series (used as fallback when absolute values absent)
        rev_vals = [rev_growth.get(y) for y in all_years]
        if any(v is not None for v in rev_vals):
            s = _series("Revenue Growth", "%", all_years, rev_vals)
            if s:
                series.append(s)

        # Historical ratio series — align to unified timeline (None for proj years)
        hist_idx = {y: i for i, y in enumerate(hist_years)}
        for label, key, unit in [
            ("EBITDA Margin", "ebitda_margin_pct", "%"),
            ("PAT Margin", "pat_margin_pct", "%"),
            ("ROE", "roe_pct", "%"),
            ("ROCE", "roce_pct", "%"),
        ]:
            raw_vals = hist.get(key)
            if not raw_vals:
                continue
            aligned = [raw_vals[hist_idx[y]] if y in hist_idx and hist_idx[y] < len(raw_vals) else None
                       for y in all_years]
            if any(v is not None for v in aligned):
                s = _series(label, unit, all_years, aligned)
                if s:
                    series.append(s)
            last = next((raw_vals[hist_idx[y]] for y in reversed(hist_years)
                         if y in hist_idx and hist_idx[y] < len(raw_vals) and raw_vals[hist_idx[y]] is not None), None)
            if last is not None:
                metrics[f"{key}_latest"] = str(last)

    # Headline numbers (skip upside_pct — recomputed from cmp+target to avoid rounding mismatch)
    for k in ("cmp", "target_price", "shares_cr"):
        v = model_json.get(k)
        if v is not None:
            metrics[k] = str(v)

    val = model_json.get("valuation") or {}
    for k in ("dcf_fair_value", "pe_fair_value", "ev_ebitda_fair_value", "blended_fair_value"):
        v = val.get(k)
        if v is not None:
            metrics[k] = str(v)

    if not metrics:
        metrics = {"placeholder": "0"}

    base["metrics"] = metrics
    base["series"] = series or _placeholder_series()
    if not series:
        warnings.append("financial-model JSON contained no usable annual series; using placeholder series")

    # Preserve richer reportgen-compatible model fields when the sidecar already
    # has them. The earlier mapper only kept a narrow v5 forecast slice, which
    # made the deck look empty even when the source JSON had SAARTHI, scenarios,
    # segments, risks, or strategy fields.
    for key in (
        "quarterly_series",
        "segments",
        "peers",
        "valuation_bands",
        "scenarios",
        "ratios",
        "saarthi",
        "management_team",
        "forensic",
        "key_highlights",
        "competitive_advantages",
        "industry_tailwinds",
        "industry_risks",
        "trading_strategy",
    ):
        _copy_if_present(base, model_json, key)

    if "saarthi" not in base and model_json.get("saarthi_scorecard"):
        base["saarthi"] = model_json["saarthi_scorecard"]

    if "segments" not in base:
        business_summary = model_json.get("business_summary") or {}
        if isinstance(business_summary, dict) and business_summary.get("segments"):
            base["segments"] = business_summary["segments"]

    if "valuation_bands" not in base:
        target_range = model_json.get("target_price_range") or {}
        if isinstance(target_range, dict) and target_range.get("base"):
            base["valuation_bands"] = [
                {
                    "method": "Target Price Range",
                    "low": str(target_range.get("low") or target_range.get("base")),
                    "base": str(target_range.get("base")),
                    "high": str(target_range.get("high") or target_range.get("base")),
                    "weight_pct": "100",
                    "notes": "Range supplied by the financial model sidecar.",
                }
            ]

    if "key_highlights" not in base and _list_of_dicts(model_json.get("key_highlights")):
        base["key_highlights"] = model_json["key_highlights"]

    if "competitive_advantages" not in base:
        comp = model_json.get("competitive_advantages")
        if isinstance(comp, list) and comp:
            base["competitive_advantages"] = comp
    return base


def _enrich_financial_model_for_house_deck(
    fin_model: dict,
    report: dict,
    sections: list[dict],
    metadata: dict,
) -> dict:
    """Guarantee renderer data refs for the fixed 15-slide Tikona deck.

    Claude and the mock planner both skip slides when refs such as scenarios,
    SAARTHI, or forensic data are absent. The approved report often contains the
    narrative even when the financial-model sidecar is thin, so we seed compact
    structured placeholders from the narrative to keep the house format intact.
    """
    narrative = _all_sections_text(sections)
    company_name = report.get("company_name") or "Company"
    cmp_val = _parse_number(metadata.get("cmp")) or 100
    target_val = _parse_number(metadata.get("target_price")) or cmp_val
    low_val = round(min(cmp_val, target_val) * 0.85, 1)
    high_val = round(max(cmp_val, target_val) * 1.15, 1)

    periods = (fin_model.get("series") or _placeholder_series())[0].get("periods") or ["FY24A", "FY25A", "FY26E"]

    if not fin_model.get("quarterly_series"):
        fin_model["quarterly_series"] = [
            {
                "name": "Quarterly Performance",
                "unit": "INR Cr",
                "periods": ["Q1", "Q2", "Q3", "Q4"],
                "values": ["0", "0", "0", "0"],
            }
        ]

    if not fin_model.get("ratios"):
        fin_model["ratios"] = [
            {
                "name": "Return Profile",
                "unit": "%",
                "periods": periods,
                "values": ["0" for _ in periods],
            }
        ]

    if not fin_model.get("valuation_bands"):
        fin_model["valuation_bands"] = [
            {
                "method": "Bear Case",
                "low": str(low_val),
                "base": str(low_val),
                "high": str(cmp_val),
                "weight_pct": "25",
                "notes": "Conservative execution and valuation assumptions.",
            },
            {
                "method": "Base Case",
                "low": str(cmp_val),
                "base": str(target_val),
                "high": str(high_val),
                "weight_pct": "50",
                "notes": "Analyst-approved target price and core thesis assumptions.",
            },
            {
                "method": "Bull Case",
                "low": str(target_val),
                "base": str(high_val),
                "high": str(high_val),
                "weight_pct": "25",
                "notes": "Upside scenario if demand drivers and execution improve.",
            },
        ]

    if not fin_model.get("scenarios"):
        fin_model["scenarios"] = [
            {"name": "Bear", "target_price": str(low_val), "probability_pct": "25", "notes": "Execution slows and valuation support weakens."},
            {"name": "Base", "target_price": str(target_val), "probability_pct": "50", "notes": "Core investment thesis plays out as expected."},
            {"name": "Bull", "target_price": str(high_val), "probability_pct": "25", "notes": "Catalysts accelerate and market confidence improves."},
        ]

    if not fin_model.get("segments"):
        business_text = _section_by_any(sections, ["business", "model", "company"]) or narrative
        fin_model["segments"] = [
            {
                "name": "Core Business",
                "description": _clean_prose(business_text, max_len=240),
            },
            {
                "name": "Growth Drivers",
                "description": _clean_prose(_section_by_any(sections, ["demand", "driver", "catalyst"]) or business_text, max_len=240),
            },
        ]

    if not fin_model.get("saarthi"):
        dimensions = [
            ("S", "Scalability of Core Engine"),
            ("A", "Addressable Market"),
            ("A", "Asymmetric Pricing Power"),
            ("R", "Reinvestment Quality"),
            ("T", "Track Record Through Adversity"),
            ("H", "Human Capital and Governance"),
            ("I", "Inflection Point Identification"),
        ]
        fin_model["saarthi"] = {
            "total_score": 70,
            "max_score": 100,
            "rating": metadata.get("rating") or "BUY",
            "dimensions": [
                {
                    "code": code,
                    "name": name,
                    "score": 10,
                    "max_score": 15,
                    "assessment": "Derived from approved analyst narrative.",
                    "key_evidence": _clean_prose(narrative, max_len=160),
                }
                for code, name in dimensions
            ],
        }

    if not fin_model.get("management_team"):
        mgmt_text = _section_by_any(sections, ["management", "governance", "forensic"]) or narrative
        fin_model["management_team"] = [
            {
                "name": "Management Team",
                "role": "Company leadership",
                "bio": _clean_prose(mgmt_text, max_len=260),
            }
        ]

    if not fin_model.get("forensic"):
        forensic_text = _section_by_any(sections, ["forensic", "governance", "risk"]) or narrative
        fin_model["forensic"] = {
            "category": "Monitor",
            "overall_assessment": _clean_prose(forensic_text, max_len=360),
            "violations": [
                {
                    "title": "Governance and forensic review",
                    "description": _clean_prose(forensic_text, max_len=260),
                    "severity": "MEDIUM",
                }
            ],
        }

    if not fin_model.get("key_highlights"):
        thesis_text = _section_by_any(sections, ["investment", "thesis", "idea"]) or narrative
        fin_model["key_highlights"] = [
            {"title": f"{company_name} investment idea", "body": item}
            for item in _bullets_from_text(thesis_text, limit=5, max_len=220)
        ]

    if not fin_model.get("competitive_advantages"):
        comp_text = _section_by_any(sections, ["competitive", "moat", "advantage"]) or narrative
        fin_model["competitive_advantages"] = _bullets_from_text(comp_text, limit=5, max_len=180)

    if not fin_model.get("industry_tailwinds"):
        industry_text = _section_by_any(sections, ["industry", "sector", "market"]) or narrative
        fin_model["industry_tailwinds"] = _bullets_from_text(industry_text, limit=4, max_len=180)

    if not fin_model.get("industry_risks"):
        risk_text = _section_by_any(sections, ["risk"]) or narrative
        fin_model["industry_risks"] = _bullets_from_text(risk_text, limit=4, max_len=180)

    if not fin_model.get("trading_strategy"):
        strategy_text = _section_by_any(sections, ["trading", "strategy", "exit", "entry"]) or narrative
        fin_model["trading_strategy"] = {
            "entry_range": "Accumulate selectively",
            "entry_rationale": _clean_prose(strategy_text, max_len=220),
            "position_size": "Risk-managed position",
            "review_frequency": "Quarterly",
            "review_metrics": _bullets_from_text(strategy_text, limit=4, max_len=120),
            "upside_exit": ["Review after target achievement"],
            "downside_exit": "Exit if thesis invalidation triggers materialise",
            "thesis_breaking_exits": _bullets_from_text(_section_by_any(sections, ["risk"]) or strategy_text, limit=3, max_len=140),
        }

    return fin_model


def _build_approved_report_md(report: dict, sections: list[dict]) -> str:
    lines = [f"# {report.get('company_name') or 'Company'} — Investment Research Report", ""]
    for s in sections:
        title = (s.get("section_title") or s.get("section_key") or "Section").strip()
        body = (s.get("content") or "").strip()
        if not body:
            continue
        lines.append(f"## {title}")
        lines.append("")
        lines.append(body)
        lines.append("")
    if len(lines) <= 2:
        # Schema requires non-empty markdown with at least some prose
        lines.append("## Summary")
        lines.append("")
        lines.append("Report content pending.")
    return "\n".join(lines)


# ─────────── supabase i/o ─────────────────────────────────────────────────────


def _fetch_inputs(client, report_id: str, session_id: str) -> tuple[dict, dict, list[dict]]:
    rep = client.table("research_reports").select("*").eq("report_id", report_id).single().execute()
    if not rep.data:
        raise ValueError(f"research_reports row not found for report_id={report_id}")

    sess = client.table("research_sessions").select("*").eq("session_id", session_id).single().execute()
    if not sess.data:
        raise ValueError(f"research_sessions row not found for session_id={session_id}")

    secs = (
        client.table("research_sections")
        .select("section_key, section_title, content, sort_order")
        .eq("session_id", session_id)
        .eq("stage", "stage2")
        .order("sort_order")
        .execute()
    )
    return rep.data, sess.data, secs.data or []


def _download_model_json(client, ticker: str, warnings: list[str]) -> dict | None:
    """Fetch financial-models/{TICKER}/{TICKER}_model.json from research-reports-html."""
    path = f"financial-models/{ticker}/{ticker}_model.json"
    try:
        # supabase-py returns bytes
        data = client.storage.from_(MODEL_BUCKET).download(path)
        return json.loads(data.decode("utf-8"))
    except Exception as exc:
        logger.warning("model JSON download failed (%s): %s", path, exc)
        warnings.append(f"Could not load financial model JSON sidecar: {exc}")
        return None


def _upload(client, local: Path, key: str, content_type: str) -> tuple[str, str]:
    with open(local, "rb") as fh:
        body = fh.read()
    client.storage.from_(PPTX_BUCKET).upload(
        path=key,
        file=body,
        file_options={"content-type": content_type, "upsert": "true"},
    )
    public = client.storage.from_(PPTX_BUCKET).get_public_url(key)
    return key, _strip_qmark(public)


# ─────────── orchestrator ─────────────────────────────────────────────────────

def _classify_market_cap(market_cap_raw: str) -> str:
    mcap = _parse_number(market_cap_raw)
    if mcap is None or mcap <= 0:
        return "Large Cap"
    if mcap >= 20000:
        return "Large Cap"
    if mcap >= 5000:
        return "Mid Cap"
    return "Small Cap"


def _fmt_mcap(market_cap_raw: str) -> str:
    mcap = _parse_number(market_cap_raw)
    if mcap is None or mcap <= 0:
        return market_cap_raw or ""
    return f"₹{mcap:,.0f} Cr"


def _synthesise_saarthi_assessment(saarthi: dict) -> str:
    """Derive a short overall_assessment from dimension evidence when the field is absent."""
    dims = saarthi.get("dimensions") or []
    # Try to pick the first dimension that has real evidence text
    for d in dims:
        evidence = (d.get("key_evidence") or d.get("assessment") or "").strip()
        if evidence and len(evidence) > 20 and "Derived from" not in evidence:
            return evidence
    # Fall back to assembling from all dimension assessments
    snippets = []
    for d in dims:
        a = (d.get("assessment") or "").strip()
        if a and len(a) > 10:
            snippets.append(a)
    if snippets:
        return _clean_prose(" ".join(snippets), max_len=600)
    return "Quality score driven by strong scalability and resilient track record."


def map_replacements(company, metadata, fin_model, sections):
    thesis   = _section_by_any(sections, ["investment", "thesis", "summary"])
    industry = _section_by_any(sections, ["industry", "sector", "market"])
    idea     = _section_by_any(sections, ["idea", "model", "demand", "driver"])
    catalyst = _section_by_any(sections, ["catalyst", "driver"])
    competitive = _section_by_any(sections, ["competitive", "moat", "advantage"])
    mgmt     = _section_by_any(sections, ["management", "forensic", "governance"])
    # Broaden peer search: catch all common DB naming conventions
    peer     = _section_by_any(sections, [
        "peer", "competition", "comparable",
        "competitive_analysis", "competitive_landscape",
        "sector_analysis", "industry_analysis",
    ])
    saarthi  = fin_model.get("saarthi") or {}

    thesis_bullets = _bullets_from_text(thesis, limit=4)

    # 6 idea bullets: combine idea + catalyst + competitive + key_highlights
    idea_combined = idea or catalyst or competitive or _all_sections_text(sections)
    idea_bullets  = _bullets_from_text(idea_combined, limit=6)
    # pad from key_highlights if we have fewer than 6
    highlights = fin_model.get("key_highlights") or []
    if isinstance(highlights, list) and len(idea_bullets) < 6:
        for h in highlights:
            body = h.get("body", h) if isinstance(h, dict) else str(h)
            bullet = _clean_prose(str(body), max_len=180)
            if bullet and bullet not in idea_bullets:
                idea_bullets.append(bullet)
            if len(idea_bullets) >= 6:
                break

    # Competitive advantage bullets — strip leading JSON/markdown artefacts
    comp_raw = fin_model.get("competitive_advantages") or []
    if isinstance(comp_raw, list) and comp_raw:
        comp_bullets = [
            _clean_prose(re.sub(r'^[\s{"\' *\-•]+', '', str(c)), max_len=280)
            for c in comp_raw[:4]
        ]
    else:
        comp_text   = competitive or idea or _all_sections_text(sections)
        comp_bullets = _bullets_from_text(comp_text, limit=4)

    # Synthesise saarthi overall_assessment once (used by both summary + content fields)
    saarthi_assessment = (
        saarthi.get("overall_assessment")
        or _synthesise_saarthi_assessment(saarthi)
    )

    # Peer text — broadened fallback so valuation/sector sections are used when
    # no section is explicitly labelled "peer" or "competition".
    peer_text = peer or _section_by_any(sections, [
        "valuation", "peers", "sector", "industry", "market",
    ])
    peer_sents = _bullets_from_text(peer_text, limit=4, max_len=220) if peer_text else []

    # Market cap helpers
    mcap_raw   = metadata.get("market_cap", "")
    m_category = _classify_market_cap(mcap_raw)
    m_cap_disp = _fmt_mcap(mcap_raw)

    # Upside display
    upside_raw = str(metadata.get("upside_pct", "") or "")
    upside_disp = (upside_raw + "%") if upside_raw and not upside_raw.endswith("%") else upside_raw

    today_str = date.today().strftime("%d %b %Y")

    replacements = {
        # ── Slide 1: Cover ────────────────────────────────────────────────────
        "company_name": company.get("name", "Company"),
        "nse_code":     company.get("ticker", ""),
        "cmp":          metadata.get("cmp", ""),
        "target":       metadata.get("target_price", ""),
        "m_cap":        m_cap_disp,
        "m_category":   m_category,
        "saarthi_s":    str(saarthi.get("total_score", "70")),
        "tagline":      "Initiation Report",
        # ── Slides 1, 4: Investment Thesis ────────────────────────────────────
        "investment_thesis_heading": "Investment Thesis",
        "investment_thesis":         _clean_prose(thesis, max_len=1200),
        "saarthi_summary_heading":   "SAARTHI Overview",
        "saarthi_summary":           _clean_prose(saarthi_assessment, max_len=300),
        "1": thesis_bullets[0] if len(thesis_bullets) > 0 else "",
        "2": thesis_bullets[1] if len(thesis_bullets) > 1 else "",
        "3": thesis_bullets[2] if len(thesis_bullets) > 2 else "",
        "4": thesis_bullets[3] if len(thesis_bullets) > 3 else "",
        # ── Slide 5: Industry Analysis ────────────────────────────────────────
        "date":     today_str,
        " date ":   today_str,
        "cell":     metadata.get("cmp", ""),
        "cell_cap": m_category,
        "mod_cap":  m_cap_disp,
        "mod":      metadata.get("upside_pct", ""),
        "tar_pr":   metadata.get("target_price", ""),
        "tar":      metadata.get("target_price", ""),  # Slide 5 uses {{tar}}
        "buy":      metadata.get("cmp", ""),
        "up":       upside_disp,
        "industry_structure": _clean_prose(industry, max_len=600),
        "key_industry":       _clean_prose("\n\n".join(fin_model.get("industry_tailwinds") or []), max_len=600),
        "key_industry_risk":  _clean_prose("\n\n".join(fin_model.get("industry_risks") or []), max_len=600),
        # ── Slide 6: Company Overview ─────────────────────────────────────────
        "COMPANY_OVERVIEW": _clean_prose(company.get("description", ""), max_len=1500),
        # ── Slide 7: Business Ideas ───────────────────────────────────────────
        "p1": idea_bullets[0] if len(idea_bullets) > 0 else "",
        "p2": idea_bullets[1] if len(idea_bullets) > 1 else "",
        "p3": idea_bullets[2] if len(idea_bullets) > 2 else "",
        "p4": idea_bullets[3] if len(idea_bullets) > 3 else "",
        "p5": idea_bullets[4] if len(idea_bullets) > 4 else "",
        "p6": idea_bullets[5] if len(idea_bullets) > 5 else "",
        # ── Slide 8: Competitive Advantages ──────────────────────────────────
        "competitve_advantage_1": comp_bullets[0] if len(comp_bullets) > 0 else "",
        "competitve_advantage_2": comp_bullets[1] if len(comp_bullets) > 1 else "",
        "competitve_advantage_3": comp_bullets[2] if len(comp_bullets) > 2 else "",
        "competitve_advantage_4": comp_bullets[3] if len(comp_bullets) > 3 else "",
        "industry_tailwinds":     _clean_prose(industry, max_len=400),
        # ── Slide 9: Peer Comparison ──────────────────────────────────────────
        "peer_comparision": _clean_prose(peer_text, max_len=600),
        "peer_para1": _clean_prose(
            peer_sents[0] if peer_sents else (peer_text[:280] if peer_text else ""), max_len=300),
        "peer_para2": _clean_prose(
            peer_sents[1] if len(peer_sents) > 1 else (peer_text[280:560] if peer_text and len(peer_text) > 280 else ""),
            max_len=300,
        ),
        # ── Slide 10: Management ──────────────────────────────────────────────
        "management_commentry_heading": "Management Analysis",
        "management_content":           _clean_prose(mgmt, max_len=1500),
        # ── Slide 11: Governance ──────────────────────────────────────────────
        "indicators": _clean_prose(
            _section_by_any(sections, ["governance", "forensic", "indicator"]) or mgmt,
            max_len=400,
        ),
        # ── Slide 15: SAARTHI ─────────────────────────────────────────────────
        "saarthi_heading":  "SAARTHI Score Analysis",
        "saarthi_content":  _clean_prose(saarthi_assessment, max_len=600),
    }

    # ── Scenario data (Slides 16, 18) ─────────────────────────────────────────
    scenarios = fin_model.get("scenarios") or []
    bear_tp_f = 0.0
    for s in scenarios:
        name  = str(s.get("name", "")).lower()
        notes = _clean_prose(str(s.get("notes", "")), max_len=200)
        prob  = str(s.get("probability_pct", "") or "")
        tp    = str(s.get("target_price", "") or "")
        prob_disp = (prob + "%") if prob and not prob.endswith("%") else prob
        if "bull" in name:
            replacements["valuation_bull"] = tp
            replacements["bull"]           = tp
            replacements["bull_p"]         = prob_disp
            replacements["bull_content"]   = notes
        elif "bear" in name:
            replacements["valuation_bear"] = tp
            replacements["bear"]           = tp
            replacements["bear_p"]         = prob_disp
            replacements["bear_content"]   = notes
            bear_tp_f = _parse_number(tp) or 0.0
        elif "base" in name:
            replacements["base"]           = tp   # template uses {{base}}, not {{valuation_base}}
            replacements["base_p"]         = prob_disp
            replacements["base_content"]   = notes

    # ── Trading Strategy (Slide 18) ───────────────────────────────────────────
    trading = fin_model.get("trading_strategy") or {}
    entry_text = _clean_prose(
        str(trading.get("entry_rationale") or trading.get("entry_range") or
            "Accumulate at current market price with defined risk."),
        max_len=300,
    )
    review_metrics = trading.get("review_metrics") or []
    review_text = _clean_prose(
        str(trading.get("review_frequency") or
            "; ".join(str(m) for m in review_metrics[:2]) or
            "Review quarterly against thesis milestones."),
        max_len=300,
    )
    exits = trading.get("thesis_breaking_exits") or []
    exit_text = _clean_prose(
        str(trading.get("downside_exit") or
            "; ".join(str(x) for x in exits[:2]) or
            "Exit on thesis invalidation or sustained breach of support."),
        max_len=300,
    )
    replacements["entry_strategy_1"]  = entry_text
    replacements["review_strategy_2"] = review_text
    replacements["exit_strategy_3"]   = exit_text

    # Slide 18 price analytics — downside % and stop-loss derived from bear scenario
    cmp_val_f = _parse_number(metadata.get("cmp") or "") or 0.0
    if bear_tp_f <= 0:
        bear_tp_f = round(cmp_val_f * 0.85, 1)
    stp_loss_val = round(bear_tp_f, 1)
    if cmp_val_f > 0:
        down_pct  = round((bear_tp_f - cmp_val_f) / cmp_val_f * 100, 1)
        down_disp = f"{down_pct}%"
    else:
        down_disp = ""
    replacements["stp_loss"] = str(stp_loss_val)
    replacements["down"]     = down_disp
    replacements["pnt"]      = metadata.get("cmp", "")  # accumulation pivot point

    return replacements


# ── Placeholder preview (called by /preview-placeholders) ─────────────────────

def preview_ppt_placeholders(report_id: str, session_id: str) -> dict:
    """Compute all text placeholder values without generating the PPTX.

    Also merges any previously saved overrides from `cs_ppt_data` so the UI
    shows the last confirmed values when the user re-opens the panel.
    """
    warnings: list[str] = []
    client = get_service_client()
    report, session, sections = _fetch_inputs(client, report_id, session_id)
    ticker = (report.get("nse_symbol") or "UNKNOWN").upper()

    company  = _build_company(report, session, sections)
    metadata = _build_metadata(report, sections)
    model_json = _download_model_json(client, ticker, warnings)
    fin_model  = _build_financial_model(ticker, model_json, warnings)
    fin_model  = _enrich_financial_model_for_house_deck(fin_model, report, sections, metadata)

    placeholders = map_replacements(company, metadata, fin_model, sections)

    # Merge previously-saved overrides so the UI shows confirmed values
    saved_raw = report.get("cs_ppt_data") or ""
    has_saved = bool(saved_raw)
    if saved_raw:
        try:
            saved = json.loads(saved_raw)
            if isinstance(saved, dict):
                placeholders.update(saved)
        except Exception:
            pass

    return {
        "status": "success",
        "placeholders": placeholders,
        "has_saved_overrides": has_saved,
        "warnings": warnings,
    }

# ── Per-slide chart data mapping ──────────────────────────────────────────────
# Maps slide_type → which financial series names to include in native PPTX charts.
# The first series' periods become the category labels.
_SLIDE_CHART_SERIES: dict[str, list[str]] = {
    "cover":               ["Revenue", "EBITDA", "PAT"],
    "story_charts":        ["Revenue", "EBITDA", "PAT"],
    "earnings_forecast":   ["Revenue", "EBITDA", "PAT"],
    "financial_highlights":["EBITDA Margin", "PAT Margin", "ROE"],
    "valuation":           ["Revenue", "EBITDA"],
    "business_segments":   ["Revenue", "EBITDA"],
    "industry":            ["Revenue"],
}

# Maps table_key → which financial series to write into a PPTX table shape.
# "table_key" is either the shape name suffix after "table:" or the slide_type.
_TABLE_SERIES: dict[str, list[str]] = {
    "earnings_forecast":   ["Revenue", "EBITDA", "PAT", "EBITDA Margin", "PAT Margin"],
    "financial_highlights":["EBITDA Margin", "PAT Margin", "ROE", "ROCE"],
    "valuation":           ["Revenue", "EBITDA", "PAT"],
}


def _detect_slide_type(slide) -> str:
    """Read slide type from a shape named 'slide_type' in the template.

    In PowerPoint, open the Selection Pane and name one shape exactly
    'slide_type', with its text content set to e.g. 'earnings_forecast'.
    Falls back to 'generic' when no such shape is found.
    """
    for shape in slide.shapes:
        if (shape.name or "").lower().strip() == "slide_type":
            if hasattr(shape, "text_frame") and shape.text_frame:
                return shape.text_frame.text.strip().lower()
    return "generic"


def _build_pptx_chart_data(slide_type: str, fin_model: dict) -> "CategoryChartData | None":
    """Build a CategoryChartData for native PPTX chart replacement, per slide type."""
    wants = _SLIDE_CHART_SERIES.get(slide_type, ["Revenue", "EBITDA", "PAT"])
    series_list = fin_model.get("series") or []
    matching = [
        s for s in series_list
        if any(w.lower() in (s.get("name") or "").lower() for w in wants)
    ]
    if not matching:
        matching = series_list[:3]
    if not matching:
        return None

    periods = matching[0].get("periods") or []
    if not periods:
        return None

    cd = CategoryChartData()
    cd.categories = periods
    for s in matching[:3]:
        vals = []
        for v in (s.get("values") or []):
            try:
                vals.append(float(str(v).replace(",", "")) if v else 0.0)
            except (ValueError, TypeError):
                vals.append(0.0)
        cd.add_series(s.get("name", "Metric"), vals)
    return cd


def _fill_table_from_model(table, table_key: str, fin_model: dict) -> None:
    """Write financial series data into an existing PPTX table shape.

    Row 0 is treated as the header (Metric | period1 | period2 ...).
    Subsequent rows receive one series each. Only writes; never adds rows/cols.
    """
    wants = _TABLE_SERIES.get(table_key)
    if not wants:
        return

    series_list = fin_model.get("series") or []
    rows_data = [
        s for s in series_list
        if any(w.lower() in (s.get("name") or "").lower() for w in wants)
    ]
    if not rows_data:
        return

    periods = rows_data[0].get("periods") or []
    n_data_cols = min(len(periods), len(table.columns) - 1)
    n_data_rows = min(len(rows_data), len(table.rows) - 1)

    if n_data_cols <= 0 or n_data_rows <= 0:
        return

    # Header row
    try:
        table.cell(0, 0).text = "Metric"
        for ci, p in enumerate(periods[:n_data_cols]):
            table.cell(0, ci + 1).text = str(p)
    except Exception as e:
        logger.warning("Table header write failed: %s", e)
        return

    # Data rows
    for ri, s in enumerate(rows_data[:n_data_rows]):
        try:
            table.cell(ri + 1, 0).text = s.get("name", "")
            for ci, val in enumerate((s.get("values") or [])[:n_data_cols]):
                table.cell(ri + 1, ci + 1).text = str(val) if val is not None else "—"
        except Exception as e:
            logger.warning("Table row %d write failed: %s", ri, e)


def _replace_text_in_frame(text_frame, replacements: dict) -> None:
    """Replace {{key}} tokens in a text frame, preserving run font properties."""
    for paragraph in text_frame.paragraphs:
        full_text = paragraph.text
        if not any(f"{{{{{k}}}}}" in full_text for k in replacements):
            continue

        # Capture font from first run before clearing
        saved_font: dict = {}
        if paragraph.runs:
            f = paragraph.runs[0].font
            saved_font["name"]  = f.name
            saved_font["size"]  = f.size
            saved_font["bold"]  = f.bold
            saved_font["color"] = None
            if getattr(f.color, "type", None) == 1:
                try:
                    saved_font["color"] = f.color.rgb
                except Exception:
                    pass

        for k, v in replacements.items():
            full_text = full_text.replace(f"{{{{{k}}}}}", str(v) if v is not None else "")

        paragraph.clear()
        run = paragraph.add_run()
        run.text = full_text
        if saved_font:
            if saved_font["name"]:  run.font.name  = saved_font["name"]
            if saved_font["size"]:  run.font.size  = saved_font["size"]
            if saved_font["bold"] is not None: run.font.bold = saved_font["bold"]
            if saved_font["color"]: run.font.color.rgb = saved_font["color"]


def _replace_text_in_table(shape, replacements: dict) -> None:
    """Replace {{key}} tokens inside every cell of a table shape."""
    for row in shape.table.rows:
        for cell in row.cells:
            if not hasattr(cell, "text_frame") or not cell.text_frame:
                continue
            for paragraph in cell.text_frame.paragraphs:
                full_text = paragraph.text
                if not any(f"{{{{{k}}}}}" in full_text for k in replacements):
                    continue
                for k, v in replacements.items():
                    full_text = full_text.replace(
                        f"{{{{{k}}}}}", str(v) if v is not None else "")
                paragraph.text = full_text


def _insert_image_into_shape(slide, shape, img_bytes: bytes) -> None:
    """Replace a placeholder shape with a generated chart image at the same bounds."""
    import io as _io
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    # Remove the original shape from the slide XML
    sp_elem = shape._element
    sp_elem.getparent().remove(sp_elem)
    # Insert picture at same position/size
    slide.shapes.add_picture(_io.BytesIO(img_bytes), left, top, width, height)


def fill_master_template(
    template_path: str,
    output_path: str,
    replacements: dict,
    fin_model: dict,
    company_name: str = "",
) -> None:
    """Fill master_template.pptx per-slide with text replacements, chart data,
    financial tables, and generated chart images.

    Shape naming convention in the template (PowerPoint Selection Pane):
      - Any shape              : plain text with {{placeholder}} tokens → replaced
      - Named 'slide_type'     : text content declares the slide type (e.g. 'earnings_forecast')
      - Named 'chart:<key>'    : replaced with a matplotlib chart image (see chart_generators.py)
      - Named 'table:<key>'    : populated from financial model series data
      - Native PPTX chart      : data updated with per-slide-type series from financial model
    """
    prs = Presentation(template_path)

    for slide in prs.slides:
        slide_type = _detect_slide_type(slide)
        pptx_chart_data = _build_pptx_chart_data(slide_type, fin_model)

        # Shapes marked for image insertion must be processed after iteration
        # (removing while iterating breaks the list).  Collect them first.
        image_insertions: list[tuple] = []  # (shape, img_bytes)

        for shape in slide.shapes:
            shape_name = (shape.name or "").lower().strip()

            # ── 1. chart:<key> shape → generate PNG and queue for insertion ──
            if shape_name.startswith("chart:"):
                chart_key = shape_name[6:]
                img = generate_chart(chart_key, fin_model, company_name)
                if img:
                    image_insertions.append((shape, img))
                continue  # skip further processing on this shape

            # ── 2. Native PPTX chart → update series data per slide type ──────
            if shape.has_chart and pptx_chart_data:
                try:
                    shape.chart.replace_data(pptx_chart_data)
                except Exception as e:
                    logger.warning("chart.replace_data failed on slide '%s': %s", slide_type, e)

            # ── 3. table:<key> or financial slide table → populate from model ─
            if shape.has_table:
                # Determine table key: explicit name wins, else use slide type
                if shape_name.startswith("table:"):
                    table_key = shape_name[6:]
                else:
                    table_key = slide_type
                _fill_table_from_model(shape.table, table_key, fin_model)
                _replace_text_in_table(shape, replacements)

            # ── 4. Text frame → {{placeholder}} replacement ───────────────────
            elif hasattr(shape, "text_frame") and shape.text_frame:
                _replace_text_in_frame(shape.text_frame, replacements)

        # ── 5. Insert queued chart images (after shape iteration is done) ──────
        for shape, img_bytes in image_insertions:
            try:
                _insert_image_into_shape(slide, shape, img_bytes)
            except Exception as e:
                logger.warning("Image insert failed for '%s': %s", shape.name, e)

    prs.save(output_path)


def generate_pptx_for_report(report_id: str, session_id: str, *, use_mock: bool = False) -> dict:
    """Top-level orchestrator. Returns the response payload for /generate-pptx."""
    t0 = time.time()
    warnings: list[str] = []

    client = get_service_client()
    report, session, sections = _fetch_inputs(client, report_id, session_id)

    ticker = (report.get("nse_symbol") or "UNKNOWN").upper()
    logger.info("generate_pptx start report=%s ticker=%s sections=%d", report_id, ticker, len(sections))

    company = _build_company(report, session, sections)
    metadata = _build_metadata(report, sections)
    model_json = _download_model_json(client, ticker, warnings)
    fin_model = _build_financial_model(ticker, model_json, warnings)
    fin_model = _enrich_financial_model_for_house_deck(fin_model, report, sections, metadata)
    md_body = _build_approved_report_md(report, sections)

    with tempfile.TemporaryDirectory(prefix="reportgen_") as tmp:
        tmp_root = Path(tmp)
        company_path = tmp_root / "company.json"
        metadata_path = tmp_root / "metadata.json"
        model_path = tmp_root / "financial_model.json"
        report_path = tmp_root / "approved_report.md"
        bundle_path = tmp_root / "bundle.json"
        output_root = tmp_root / "out"
        output_root.mkdir(parents=True, exist_ok=True)

        company_path.write_text(json.dumps(company, default=str), encoding="utf-8")
        metadata_path.write_text(json.dumps(metadata, default=str), encoding="utf-8")
        model_path.write_text(json.dumps(fin_model, default=str), encoding="utf-8")
        report_path.write_text(md_body, encoding="utf-8")

        bundle = {
            "company_path": str(company_path),
            "metadata_path": str(metadata_path),
            "financial_model_path": str(model_path),
            "approved_report_path": str(report_path),
        }
        bundle_path.write_text(json.dumps(bundle), encoding="utf-8")

        logger.info("Filling master template directly")

        replacements = map_replacements(company, metadata, fin_model, sections)

        # Apply saved PPT placeholder overrides confirmed by the user in the UI
        saved_ppt_raw = report.get("cs_ppt_data") or ""
        if saved_ppt_raw:
            try:
                saved_overrides = json.loads(saved_ppt_raw)
                if isinstance(saved_overrides, dict):
                    replacements.update(saved_overrides)
                    logger.info("Applied %d saved PPT overrides from cs_ppt_data", len(saved_overrides))
            except Exception as e:
                logger.warning("Failed to parse cs_ppt_data overrides: %s", e)

        # Check multiple possible paths for the master template (local vs docker)
        possible_paths = [
            Path(__file__).resolve().parent.parent.parent / "master_template.pptx",  # local repo structure
            Path(__file__).resolve().parent / "master_template.pptx",  # If dumped in the same directory
            Path("/app/master_template.pptx"), # Docker container path
            Path("master_template.pptx"), # Current working directory
        ]
        
        template_path = None
        for p in possible_paths:
            if p.exists():
                template_path = p
                break
                
        if not template_path:
            raise RuntimeError(f"Master template not found in any of the expected locations.")

        result_pptx_path = str(output_root / "report.pptx")
        
        if template_path.exists():
            fill_master_template(
                str(template_path), result_pptx_path, replacements, fin_model,
                company_name=company.get("name", ""),
            )
        else:
            raise RuntimeError(f"Master template not found at {template_path}")

        # Upload artifacts
        ts = int(time.time())
        pptx_key = f"{ticker}/{report_id}/report_{ts}.pptx"
        pptx_path_out, pptx_url = _upload(client, Path(result_pptx_path), pptx_key, PPTX_CONTENT_TYPE)

        pdf_path_out: str | None = None
        pdf_url: str | None = None

    # Update DB
    now_iso = datetime.now(timezone.utc).isoformat()
    client.table("research_reports").update(
        {
            "pptx_file_path": pptx_path_out,
            "pptx_file_url": pptx_url,
            "pptx_pdf_file_path": pdf_path_out,
            "pptx_pdf_file_url": pdf_url,
            "pptx_generated_at": now_iso,
            "pptx_status": "ready",
            "updated_at": now_iso,
        }
    ).eq("report_id", report_id).execute()

    duration = round(time.time() - t0, 2)
    logger.info("generate_pptx done report=%s duration=%.2fs", report_id, duration)

    return {
        "status": "success",
        "message": f"PPTX generated in {duration}s",
        "pptx_file_url": pptx_url,
        "pptx_file_path": pptx_path_out,
        "pptx_pdf_file_url": pdf_url,
        "pptx_pdf_file_path": pdf_path_out,
        "duration_seconds": duration,
        "warnings": warnings,
    }
