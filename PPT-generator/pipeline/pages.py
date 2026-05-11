"""Phase 3: Page content generator.

For each PageBrief in the spine, make one LLM call that sees:
  - the full spine (coherence anchor)
  - the immediate neighbors' briefs (continuity)
  - the sliced narrative data for this page
  - the financial model (so tables / chart requests can use real numbers)

The model returns a PageContent JSON (blocks only — no HTML).

Parallelized via ThreadPoolExecutor since calls are I/O-bound.
"""

from __future__ import annotations

import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any

from .data_loader import slice_narrative
from .llm_client import LLMClient
from .schemas import CompanyPack, PageBrief, PageContent, PageEditorialSpec, SpineOutline


PAGE_SYSTEM = """You are a senior equity research analyst drafting ONE page of \
a LANDSCAPE 16:9 institutional research report.

CRITICAL LAYOUT CONSTRAINT:
- This is a LANDSCAPE page (wider than tall) — think PowerPoint slide proportions.
- Vertical space is VERY LIMITED (~170mm usable height).
- You MUST keep content compact so it fits within ONE page without clipping.
- A page typically has 2–4 blocks. 3 is ideal. DO NOT overload a page.
- EXCEPTION: story_charts pages MUST have EXACTLY 6 chart blocks — see STORY_CHARTS RULES below.
- Think like an editor working to a 12–13 page initiation deck, not an analyst writing a long memo.
- If forced to choose, OMIT lower-priority content rather than squeezing everything in.

Content size limits (HARD RULES — violating these will cause clipping):
- Tables: MAX 12 data rows, MAX 7 columns. Prefer 5-6 rows.
- Bullet lists: MAX 4 items. Each bullet body MAX 2 sentences.
- Paragraphs: MAX 2 sentences each. Keep body under 45 words.
- Metric strips: items count MUST be EXACTLY 2, 3, 4, or 5. NEVER use 1, 6, 7, or 8.
- Prefer fewer, high-impact blocks over many small ones.
- Never emit more than 1 table and 1 chart on the same page except:
  story_charts pages can have up to 6 charts (key financial metrics),
  valuation pages can have chart + scenario + table.
- Do NOT create placeholder disclosure, analyst-credit, or boilerplate institutional-client blocks.

STORY_CHARTS PAGE RULES (page_type = "story_charts") — MANDATORY:
- You MUST emit EXACTLY 6 chart blocks. No more, no fewer. This page renders in a 3×2 grid.
- Use ONLY metric names that appear in the financial model provided. Do NOT invent metric names.
- Suggested 6 charts: (1) Revenue bar by year, (2) EBITDA bar by year, (3) PAT bar by year,
  (4) EBITDA Margin line by year, (5) PAT Margin line by year, (6) EPS bar by year.
  Substitute with available model metrics if the above are not present (check "metrics" list in model JSON).
- Each chart title must include the unit: e.g. "Revenue (INR Cr) — FY22 to FY26E".
- Include a short caption as the title's subtitle; these render as figure footnotes.
- Do NOT include any non-chart blocks on story_charts pages.

SYMMETRY RULE (CRITICAL — your senior will reject asymmetric layouts):
- All metric card rows MUST fill their grid evenly: 2x1, 3x1, 4x1, or 5x1.
- NEVER output 4 metric cards followed by 1 card — that leaves 3 empty spaces and looks terrible.
- If you have 5 data points, use 5 cards in ONE row. If you have 6, split into TWO rows of 3.
- Same applies to highlight boxes: use 2x1 or 3x1 grids, never orphan 1 box next to 3.

NUMERIC CONSISTENCY RULES (HARD — violations will ship as visible errors):
- NEVER put a specific count-word in a page title unless the block underneath has exactly that many items. Do NOT write "Eleven Risks", "Seven Catalysts", "Six Drivers" — titles must be neutral ("Risk Register", "Key Catalysts") or use the exact rendered count. Risk blocks render at most 5 items; catalyst blocks at most 4.
- Scenario cases (bear/base/bull) MUST satisfy: bear.target_price < CMP < bull.target_price. A "bear" case whose TP is at or above CMP is not a bear case — revise it to genuine downside (typically 15–30% below CMP) or relabel.
- The `updown` field will be RECOMPUTED by the renderer from (CMP, target_price). You may omit it. If you include it, make it arithmetically correct vs CMP — do not write "-7% vs base" when the number implies -12.6%.
- Number formatting — use ONE convention consistently:
  * Large INR amounts: "₹32,259 Cr" (comma-separated, explicit "Cr")
  * Per-share prices: "₹1,408" (no suffix; understood as per-share)
  * International / capex comparisons: "US$80 bn" (never mix "$80B+" shorthand)
  * NEVER use Indian lakh-crore format like "₹1,17,102" — always decimal-thousands with explicit "Cr"

CONTENT YOU MUST NEVER INCLUDE:
- Analyst consensus, street estimates, Bloomberg consensus, or any "consensus target price" content. We are the PRIMARY research — we do not cite street consensus.
- Do NOT include CMP, Target Price, Upside, Market Cap, or Ticker as metric cards or content blocks. These ALREADY appear in every page's header bar — repeating them in the body wastes space and looks redundant.
- Do NOT include "Rating" or "BUY" as a metric card — it's already in the header badge.
- Do NOT include generic disclaimer text, analyst bylines, "for institutional clients only", or report metadata blocks in the body.

CHARTS — MANDATORY ON DATA PAGES:
- story_charts, financial_highlights, earnings_forecast, business_segments, valuation, industry, peer_comparison pages MUST emit at least one `chart` block. Text+tables only is unacceptable for these pages.
- Supported chart_types: "bar" (Revenue/EBITDA/PAT trends), "line" (margin/ROE trajectory, ARPU), "stacked_bar" (segment revenue over years), "donut" (segment mix — uses `slices: [{label, value}]`, not `metrics`).
- Suggested minimums across the report: (1) revenue/EBITDA bar by year, (2) PAT line by year, (3) segment-mix donut, (4) PE/EV band line, (5) peer EV/EBITDA bar.

VALUATION PAGE — SOTP TABLE REQUIRED (for conglomerates / multi-segment cos):
- When the company has 3+ reportable segments, the valuation page MUST include a `table` block titled "Sum-of-the-Parts (SOTP) Derivation" with headers: ["Segment", "Valuation Method", "EV (₹ Cr)", "Per Share (₹)", "Weight %"] and a final "Total / Implied Target" row. Do NOT just reference SOTP in prose without showing the maths.

SCORECARD PAGE (page_type = "scorecard") — SPECIAL RULES:
- The SAARTHI Scorecard is a key differentiator. You MUST show ALL dimensions with full detail.
- Use a "scorecard" block (NOT a table) — it renders with professional progress bars.
- Include EVERY dimension from narrative.saarthi_scorecard.dimensions:
  Each item: { name: "Dimension Name", score: "14", max_score: "15", description: "1-line assessment" }
- Also add a metrics block with: label="SAARTHI Total Score", value="79/100", delta=rating.
- Then add one callout (variant="thesis") with the overall conviction statement.
- Do NOT hand-author a framework intro or band legend — the renderer auto-injects these above the scorecard block.

Other rules:
- You write the content for exactly ONE page. Do not write other pages.
- Every claim must tie back to the report's thesis_north_star.
- Use ONLY the data provided. Do NOT invent numbers, ratings, or targets.
- Output STRUCTURED CONTENT BLOCKS (JSON), not HTML.
- Prose is tight, institutional, data-led. No hype, no filler.
- Numbers in tables come from the financial model provided. Round sensibly.
- Avoid repetition with the previous/next page. If the same point is already owned by a neighbor page, compress or omit it here.
- Cover pages are especially space-constrained because the renderer adds a financial summary and sidebar. Keep cover-page body blocks to 2-3 only.

Block types you can emit (use VARIETY, but do NOT add blocks just for variety):

  BASIC BLOCKS:
  paragraph : { kind: "paragraph", title?, body }  — body MAX 60 words
  bullets   : { kind: "bullets", title?, items: [{title?, body}] }  — MAX 4 items
  table     : { kind: "table", title?, headers, rows, footnote?, highlight_rows?: [0,1] }  — MAX 8 rows, MAX 7 cols. highlight_rows is a list of 0-indexed row indices to highlight (default [0]).
  metrics   : { kind: "metrics", title?, items: [{label, value, delta?}] }  — MAX 5 items
  callout   : { kind: "callout", variant: "thesis"|"warn"|"info"|"quote", title?, body }  — body MAX 80 words
  chart     : { kind: "chart", chart_type: "bar"|"line"|"stacked_bar", title?, metrics: [metric names in model], years?: [year labels] }

  RICH BLOCKS (use these for visual variety — they look much better than plain bullets):
  catalyst  : { kind: "catalyst", title?, items: [{icon: "🚀", title, badge?: "HIGH IMPACT", body}] }  — for catalysts, initiatives. MAX 4 items.
  risk      : { kind: "risk", title?, items: [{severity: "high"|"med"|"low", title, body}] }  — for risk pages. MAX 5 items.
  scenario  : { kind: "scenario", title?, cases: [{label, case: "bear"|"base"|"bull", target_price, updown?, description?}] }  — for valuation pages. Exactly 3 cases.
  scorecard : { kind: "scorecard", title?, items: [{name, score: "14", max_score: "15", description?}] }  — for SAARTHI scorecard. MAX 7 items.
  timeline  : { kind: "timeline", title?, items: [{year, text}] }  — for milestones, history. MAX 12 items.

PAGE-TYPE GUIDANCE (use the RIGHT block types for each page):
  cover              → metrics + callout(thesis) + catalyst  (NO table — renderer adds financial summary)
  story_charts       → EXACTLY 6 chart blocks (3×2 grid, no other block types allowed)
  thesis             → callout(thesis) + risk (summarized) + table
  industry           → chart + metrics + risk OR chart + timeline + metrics
  company_overview   → metrics + paragraph + timeline OR metrics + table + timeline
  business_segments  → chart + table + catalyst OR chart + metrics + catalyst (Business model, demand drivers, competitive landscape)
  management         → table + bullets + callout
  earnings_forecast  → table + metrics + bullets (state assumptions clearly)
  financial_highlights → EXACTLY ONE table block (Key Ratios). MANDATORY format:
    title: "Key Ratios"
    headers: ["Ratio", "FY22A", "FY23A", "FY24A", "FY25A", "FY26E", "FY27E", "FY28E"]  (use actual years from model)
    rows (in this order, values from financial model):
      ["EBITDA Margin", "49.3%", ...],   ← row 0 — highlight
      ["PAT Margin",    "30.5%", ...],   ← row 1 — highlight
      ["RoE",           "23.0%", ...],   ← green % auto-colored
      ["RoCE",          "19.0%", ...],
      ["RoIC",          "92.0%", ...],   (include only if available in model)
      ["P/E",           "37.5x", ...],
      ["P/B",           "8.7x",  ...],
      ["EV/EBITDA",     "9.0x",  ...]
    highlight_rows: [0, 1]              ← highlights EBITDA Margin and PAT Margin rows
    footnote: "E = Tikona Capital Estimates"
    Do NOT add any other blocks (no charts, no metrics) on this page.
  valuation          → scenario + table + callout + metrics
  scorecard          → scorecard (MUST use scorecard block with progress bars)
  scenario_analysis  → scenario + bullets + metrics
  entry_strategy     → callout(thesis) + metrics + bullets + risk
  catalysts          → catalyst (MUST use catalyst blocks, NOT bullets)
  risks              → risk (MUST use risk blocks with severity badges)
  peer_comparison    → table + metrics
  esg / appendix     → table + bullets + callout
"""


PAGE_USER_TEMPLATE = """Report thesis (north star): {thesis}
Report tone: {tone}

Full outline (for continuity only — do NOT rewrite other pages):
{outline}

Previous page: {prev}
THIS page:     {current}
Next page:     {next}

Editorial brief for THIS page:
{editorial}

Narrative data for THIS page:
```json
{narrative}
```

Financial model (years and metrics — use these for tables / chart requests):
```json
{financials}
```

Write THIS page. Return JSON exactly in this shape:

{{
  "page_number": {page_number},
  "page_type": "{page_type}",
  "title": "{title}",
  "subtitle": "<optional short subtitle or null>",
  "blocks": [ ...{{block objects per schema above}}... ]
}}
"""


def _outline_summary(spine: SpineOutline) -> str:
    lines = []
    for p in spine.pages:
        lines.append(f"  {p.page_number}. [{p.page_type}] {p.title} — {p.key_message}")
    return "\n".join(lines)


def _brief_str(p: PageBrief | None) -> str:
    if p is None:
        return "(none)"
    return f"{p.page_number}. [{p.page_type}] {p.title} — {p.key_message}"


def _financials_for_prompt(pack: CompanyPack) -> dict[str, Any]:
    if not pack.financials:
        return {}
    return {
        "years": pack.financials.years,
        "metrics": [r.metric for r in pack.financials.rows],
        "rows": {r.metric: r.values for r in pack.financials.rows},
        "cagr": pack.financials.cagr,
    }


def _editorial_payload(brief: PageBrief, editorial_spec: PageEditorialSpec | None) -> dict[str, Any]:
    if editorial_spec is None:
        return {
            "narrative_role": f"Advance the report through {brief.title}",
            "must_land": brief.key_message,
            "must_include": [],
            "avoid_repeating": [],
            "cut_first": [],
            "preferred_blocks": [],
            "max_blocks": 3,
        }
    return editorial_spec.model_dump()


def build_page_user_prompt(
    pack: CompanyPack,
    spine: SpineOutline,
    brief: PageBrief,
    editorial_spec: PageEditorialSpec | None = None,
    extra_feedback: str | None = None,
) -> str:
    idx = brief.page_number - 1
    prev_p = spine.pages[idx - 1] if idx > 0 else None
    next_p = spine.pages[idx + 1] if idx + 1 < len(spine.pages) else None

    narrative_slice = slice_narrative(pack, brief.data_slices) if brief.data_slices else {}
    scalars = {
        k: pack.narrative.get(k)
        for k in ("company", "ticker", "sector", "rating", "cmp",
                  "target_price", "upside_potential_pct", "tagline")
        if k in pack.narrative
    }
    user = PAGE_USER_TEMPLATE.format(
        thesis=spine.thesis_north_star,
        tone=spine.tone,
        outline=_outline_summary(spine),
        prev=_brief_str(prev_p),
        current=_brief_str(brief),
        next=_brief_str(next_p),
        editorial=json.dumps(_editorial_payload(brief, editorial_spec), indent=2, ensure_ascii=False),
        narrative=json.dumps({"_header": scalars, **narrative_slice}, indent=2, ensure_ascii=False),
        financials=json.dumps(_financials_for_prompt(pack), indent=2, ensure_ascii=False),
        page_number=brief.page_number,
        page_type=brief.page_type,
        title=brief.title,
    )
    if extra_feedback:
        user += "\n\n" + extra_feedback
    return user


def regenerate_page_with_feedback(
    pack: CompanyPack,
    spine: SpineOutline,
    brief: PageBrief,
    client: LLMClient,
    feedback: str,
    editorial_spec: PageEditorialSpec | None = None,
) -> PageContent:
    user = build_page_user_prompt(
        pack,
        spine,
        brief,
        editorial_spec=editorial_spec,
        extra_feedback=feedback,
    )
    parsed, _ = client.generate_json(PAGE_SYSTEM, user)
    parsed["page_number"] = brief.page_number
    parsed["page_type"] = brief.page_type
    parsed.setdefault("title", brief.title)
    try:
        return PageContent.model_validate(parsed)
    except Exception as e:
        fix_user = user + (
            f"\n\nYour previous output failed validation with this error:\n{e}\n\n"
            "Fix it. Every block MUST include a valid 'kind' field. "
            "Allowed block kinds are: paragraph, bullets, table, metrics, callout, chart, "
            "catalyst, risk, scenario, scorecard, timeline. "
            "For chart blocks, ONLY use: {kind, chart_type, title?, metrics, years?, slices?}. "
            "Do NOT include unsupported fields like data, footnote, chart_subtype, series, or options. "
            "For metrics items use {label, value, delta?}. "
            "For bullets items use {title?, body}. "
            "Return corrected JSON only."
        )
        parsed, _ = client.generate_json(PAGE_SYSTEM, fix_user)
        parsed["page_number"] = brief.page_number
        parsed["page_type"] = brief.page_type
        parsed.setdefault("title", brief.title)
        return PageContent.model_validate(parsed)


def generate_page(
    pack: CompanyPack,
    spine: SpineOutline,
    brief: PageBrief,
    client: LLMClient,
    editorial_spec: PageEditorialSpec | None = None,
) -> PageContent:
    user = build_page_user_prompt(pack, spine, brief, editorial_spec=editorial_spec)

    parsed, _ = client.generate_json(PAGE_SYSTEM, user)
    parsed["page_number"] = brief.page_number
    parsed["page_type"] = brief.page_type
    parsed.setdefault("title", brief.title)

    try:
        return PageContent.model_validate(parsed)
    except Exception as e:
        # One-shot retry: feed the validation error back to the model.
        fix_user = user + (
            f"\n\nYour previous output failed validation with this error:\n{e}\n\n"
            "Fix it. Every block MUST include a 'kind' field matching one of: "
            "paragraph, bullets, table, metrics, callout, chart, catalyst, risk, scenario, scorecard, timeline. "
            "For 'metrics' each item is {label, value, delta?}. "
            "For 'bullets' each item is {title?, body}. "
            "Return the corrected JSON only."
        )
        parsed, _ = client.generate_json(PAGE_SYSTEM, fix_user)
        parsed["page_number"] = brief.page_number
        parsed["page_type"] = brief.page_type
        parsed.setdefault("title", brief.title)
        return PageContent.model_validate(parsed)


def generate_all_pages(
    pack: CompanyPack,
    spine: SpineOutline,
    client: LLMClient | None = None,
    editorial_plan: Any | None = None,
    max_workers: int = 6,
) -> list[PageContent]:
    client = client or LLMClient()
    results: dict[int, PageContent] = {}
    errors: dict[int, Exception] = {}
    editorial_by_page = {
        spec.page_number: spec for spec in getattr(editorial_plan, "page_specs", [])
    }

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {
            ex.submit(
                generate_page,
                pack,
                spine,
                brief,
                client,
                editorial_by_page.get(brief.page_number),
            ): brief
            for brief in spine.pages
        }
        for fut in as_completed(futures):
            brief = futures[fut]
            try:
                results[brief.page_number] = fut.result()
            except Exception as e:  # noqa: BLE001
                errors[brief.page_number] = e

    if errors:
        msg = "; ".join(f"page {n}: {type(e).__name__}: {e}" for n, e in errors.items())
        raise RuntimeError(f"Page generation failed for {len(errors)} page(s): {msg}")

    return [results[i] for i in sorted(results)]
