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

import reportgen
from reportgen.config import settings
from reportgen.ai.openrouter_client import OpenRouterPlanningClient
from reportgen.orchestration.pipeline import run_local_pipeline

from supabase_client import get_service_client

logger = logging.getLogger(__name__)

PPTX_BUCKET = "research-reports-pptx"
MODEL_BUCKET = os.environ.get("SUPABASE_FINANCIAL_MODEL_BUCKET", "research-reports-html")

PPTX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
PDF_CONTENT_TYPE = "application/pdf"

_NUM_RE = re.compile(r"[\d,]+\.?\d*")
_PLACEHOLDER_RE = re.compile(r"not included in the generated report|content pending", re.I)
_OPENROUTER_PATCHED = False


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


def _find_reportgen_root() -> Path:
    package_path = Path(reportgen.__file__).resolve()
    for parent in package_path.parents:
        if (parent / "prompts" / "user" / "slide_planner_input.md").exists():
            return parent
    raise RuntimeError(f"Could not find reportgen prompts root from {package_path}")


@contextmanager
def _reportgen_cwd():
    previous = Path.cwd()
    os.chdir(_find_reportgen_root())
    try:
        yield
    finally:
        os.chdir(previous)


def _patch_openrouter_client() -> None:
    """Avoid openai/httpx version skew by using OpenRouter's HTTP API directly."""
    global _OPENROUTER_PATCHED
    if _OPENROUTER_PATCHED:
        return

    def generate(self, system_prompt: str, user_prompt: str) -> str:
        if not self.api_key:
            raise RuntimeError("OpenRouter API key is not configured.")

        url = "https://openrouter.ai/api/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "HTTP-Referer": "https://github.com/pptx-research-report",
            "X-Title": "PPTX Research Report Generator",
            "Content-Type": "application/json",
        }
        data = {
            "model": self.model_name,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "max_tokens": settings.planning_max_tokens,
            "temperature": settings.planning_temperature,
        }

        req = urllib.request.Request(url, json.dumps(data).encode("utf-8"), headers)
        try:
            with urllib.request.urlopen(req, timeout=180) as response:
                result = json.loads(response.read().decode("utf-8"))
                return result["choices"][0]["message"]["content"]
        except urllib.error.HTTPError as exc:
            error_body = exc.read().decode("utf-8")
            raise RuntimeError(f"OpenRouter API error: {exc.code} - {error_body}") from exc

    OpenRouterPlanningClient.generate = generate
    _OPENROUTER_PATCHED = True


# ─────────── input builders ───────────────────────────────────────────────────


def _build_company(report: dict, session: dict) -> dict:
    return {
        "name": report.get("company_name") or "Unknown",
        "ticker": (report.get("nse_symbol") or "UNKNOWN").upper(),
        "exchange": "NSE",
        "sector": session.get("sector") or "General",
        "country": "India",
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

    if all_years:
        rev_growth = asmp.get("revenue_growth_pct") or {}
        # Build revenue growth series aligned to unified timeline (None for hist years)
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
    return base


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


def generate_pptx_for_report(report_id: str, session_id: str, *, use_mock: bool = False) -> dict:
    """Top-level orchestrator. Returns the response payload for /generate-pptx."""
    t0 = time.time()
    warnings: list[str] = []

    client = get_service_client()
    report, session, sections = _fetch_inputs(client, report_id, session_id)

    ticker = (report.get("nse_symbol") or "UNKNOWN").upper()
    logger.info("generate_pptx start report=%s ticker=%s sections=%d", report_id, ticker, len(sections))

    company = _build_company(report, session)
    metadata = _build_metadata(report, sections)
    model_json = _download_model_json(client, ticker, warnings)
    fin_model = _build_financial_model(ticker, model_json, warnings)
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

        logger.info("running reportgen pipeline use_mock=%s", use_mock)
        _patch_openrouter_client()
        with _reportgen_cwd():
            result = run_local_pipeline(bundle_path=bundle_path, output_root=output_root, use_mock=use_mock)

        if result.manifest.notes:
            warnings.extend(result.manifest.notes)

        if result.manifest.status == "render_failed" or not result.pptx_path:
            raise RuntimeError(f"reportgen render_failed: {'; '.join(result.manifest.notes) or 'unknown'}")

        # Upload artifacts
        pptx_key = f"{ticker}/{report_id}/report.pptx"
        pptx_path_out, pptx_url = _upload(client, result.pptx_path, pptx_key, PPTX_CONTENT_TYPE)

        pdf_path_out: str | None = None
        pdf_url: str | None = None
        if result.pdf_path and Path(result.pdf_path).exists():
            pdf_key = f"{ticker}/{report_id}/report.pdf"
            pdf_path_out, pdf_url = _upload(client, result.pdf_path, pdf_key, PDF_CONTENT_TYPE)
        else:
            warnings.append("PDF conversion skipped (no LibreOffice/PowerPoint detected)")

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
