"""Phase 2: Spine generator.

One LLM call that reads the CompanyPack and returns a SpineOutline:
  - thesis_north_star (the single argument the whole report defends)
  - pages: list of PageBrief (type, title, key_message, data_slices)

The output is short, so truncation is not a risk. This outline becomes the
coherence anchor — every downstream page call sees it.
"""

from __future__ import annotations

import json
from typing import Any

from .llm_client import LLMClient
from .schemas import CompanyPack, SpineOutline


ALLOWED_PAGE_TYPES = [
    "cover",
    "story_charts",
    "thesis",
    "industry",
    "company_overview",
    "business_segments",
    "management",
    "earnings_forecast",
    "financial_highlights",
    "valuation",
    "scorecard",
    "scenario_analysis",
    "risks",
    "entry_strategy",
    "catalysts",
    "peer_comparison",
    "esg",
    "appendix",
    # NOTE: "disclaimer" is auto-appended by renderer, not planned by LLM
]


SPINE_SYSTEM = """You are a senior equity research editor at an Indian institutional \
brokerage (think IIFL / Jefferies India). You plan the structure of landscape \
research reports before analysts draft them.

Your job: given a company research pack, write the title, key_message, and data_slices \
for each of the MANDATORY pages below. You CANNOT skip, reorder, or replace any of them.

MANDATORY PAGE STRUCTURE — output EXACTLY these 14 pages in this order:
   1.  cover            — Teaser page. Deal maker/breaker. Recommendation + why, in 60 seconds.
   2.  story_charts     — Story in Charts. 6 key financial charts that tell the whole story visually.
   3.  thesis           — Investment Thesis. Core bull case, valuation summary, key risks.
   4.  industry         — Industry Overview. Sector tailwinds, structural drivers, key risks.
   5.  company_overview — Company Overview. What it does, revenue split, 2-sentence explainer.
   6.  business_segments — Key Investment Idea in Detail. Business model, demand drivers, competitive landscape.
   7.  management       — Management & Corporate Governance. Board quality, track record, credibility.
   8.  earnings_forecast — Earnings Forecast. 3Y historical + 2–3Y forward P&L, key assumptions stated.
   9.  financial_highlights — Financials — Key Ratios. Margin + return + valuation multiple trajectory.
   10. valuation        — Valuations. Multi-method (P/E, EV/EBITDA, DCF/SOTP). Sensitivity / scenario.
   11. scorecard        — SAARTHI Framework. Every dimension with progress bars and detailed commentary.
   12. scenario_analysis — Scenario Analysis. Bull / Base / Bear cases with explicit assumptions.
   13. risks            — Key Risks & Thesis Invalidation Triggers. What kills the thesis.
   14. entry_strategy   — Entry, Review & Exit Strategy. When to buy, hold, and exit.

Do NOT include a "disclaimer" page — it is auto-appended by the system.
Do NOT add optional pages (peer_comparison, esg, appendix, catalysts) unless explicitly requested.

Rules:
- Output EXACTLY 14 pages in the order above. No additions, no omissions.
- Each page has one and only one "key_message" — a single sentence the page must prove.
- The thesis_north_star is the ONE argument the whole report defends.
- Use only these page types: {page_types}
- "data_slices" are dotted keys from the narrative JSON the page will consume.
- Every page must advance the thesis. No filler.
"""


SPINE_USER_TEMPLATE = """Company research pack (JSON):
```json
{pack_json}
```

Available narrative keys (top level): {narrative_keys}
Financial years in model: {years}

Design the outline. Output JSON matching this exact shape:

{{
  "thesis_north_star": "<one sentence>",
  "tone": "<short phrase>",
  "pages": [
    {{
      "page_number": 1,
      "page_type": "<one of the allowed types>",
      "title": "<page title>",
      "key_message": "<one sentence>",
      "data_slices": ["<dotted.key>", ...]
    }},
    ...
  ]
}}
"""


def _compact_pack_for_spine(pack: CompanyPack) -> dict[str, Any]:
    """Slim the pack down so the spine call stays cheap.

    Keep scalars + short descriptive fields; drop long prose sections.
    Spine doesn't need the full thesis body — just enough to plan structure.
    """
    keep_scalar = {
        "company", "ticker", "sector", "rating", "cmp", "target_price",
        "upside_potential_pct", "tagline",
    }
    out: dict[str, Any] = {k: pack.narrative.get(k) for k in keep_scalar if k in pack.narrative}

    # For each remaining section, include only a short preview so the planner
    # knows what's available without reading every word.
    for k, v in pack.narrative.items():
        if k in keep_scalar or k in out:
            continue
        if isinstance(v, str):
            out[k] = v[:400] + ("…" if len(v) > 400 else "")
        elif isinstance(v, list):
            out[k] = f"[list with {len(v)} items]"
        elif isinstance(v, dict):
            out[k] = {"_keys": list(v.keys())[:20]}
        else:
            out[k] = v
    return out


def generate_spine(pack: CompanyPack, client: LLMClient | None = None) -> tuple[SpineOutline, dict]:
    """Returns (validated SpineOutline, raw usage info)."""
    client = client or LLMClient()

    system = SPINE_SYSTEM.format(page_types=", ".join(ALLOWED_PAGE_TYPES))
    user = SPINE_USER_TEMPLATE.format(
        pack_json=json.dumps(_compact_pack_for_spine(pack), indent=2, ensure_ascii=False),
        narrative_keys=", ".join(sorted(pack.narrative.keys())),
        years=", ".join(pack.financials.years) if pack.financials else "n/a",
    )

    parsed, result = client.generate_json(system, user)

    # Ensure page numbers are sequential starting at 1.
    pages = parsed.get("pages", [])
    for i, p in enumerate(pages, start=1):
        p["page_number"] = i
    parsed["pages"] = pages

    spine = SpineOutline.model_validate(parsed)
    return spine, {"model": result.model, "usage": result.usage}
