"""Typed schemas for the report pipeline.

Three layers:
  - CompanyPack   : canonical, validated input (JSON + CSV merged)
  - SpineOutline  : model-authored outline (Phase 2 output)
  - PageContent   : model-authored per-page content (Phase 3 output)

Using pydantic so the same schemas validate LLM JSON output end-to-end.
"""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import BaseModel, Field, ConfigDict


# ─── Company pack (input) ─────────────────────────────────────────────────────

class FinancialRow(BaseModel):
    """One metric across years. Years are ISO-ish labels like 'FY24' or 'Mar-24'."""
    metric: str
    unit: str | None = None
    values: dict[str, float | None]

    model_config = ConfigDict(extra="forbid")


class FinancialModel(BaseModel):
    """Parsed CSV — deterministic numeric backbone of the report."""
    years: list[str]
    rows: list[FinancialRow]
    cagr: dict[str, dict[str, float | None]] = Field(default_factory=dict)
    raw_header: str | None = None

    def get(self, metric: str) -> FinancialRow | None:
        key = metric.strip().lower()
        # 1. Exact case-insensitive match
        for r in self.rows:
            if r.metric.strip().lower() == key:
                return r
        # 2. Partial match — key is a substring of row metric or vice versa
        # Mirrors the approach used in _render_cover_fin_table so chart rendering
        # degrades gracefully when LLM uses a slightly different metric name.
        for r in self.rows:
            rm = r.metric.strip().lower()
            if key in rm or rm in key:
                return r
        return None


class CompanyPack(BaseModel):
    """Single source of truth per company. Anything downstream reads this."""
    company: str
    ticker: str
    sector: str | None = None
    rating: str | None = None
    cmp: float | str | None = None
    target_price: float | str | None = None
    upside_potential_pct: float | int | str | None = None
    tagline: str | None = None

    # The full JSON blob (narrative fields — thesis, risks, mgmt, etc.)
    narrative: dict[str, Any]

    # Parsed CSV
    financials: FinancialModel | None = None

    model_config = ConfigDict(extra="allow")


# ─── Spine (Phase 2 output) ───────────────────────────────────────────────────

PageType = Literal[
    "cover",               # 1. Teaser / Cover page
    "story_charts",        # 2. Story in Charts
    "thesis",              # 3. Investment Thesis
    "industry",            # 4. Industry Overview (Tailwinds / Risks)
    "company_overview",    # 5. Company Overview
    "business_segments",   # 6. Key Investment Ideas / Business Model / Demand Drivers
    "management",          # 7. Management Analysis & Corporate Governance
    "earnings_forecast",   # 8. Earnings Forecast
    "financial_highlights",# 9. Financials (P&L, BS, CF)
    "valuation",           # 10. Valuations (multi-method)
    "scorecard",           # 11. SAARTHI Framework
    "scenario_analysis",   # 12. Scenario Analysis (Bull/Base/Bear)
    "risks",               # 13. Key Risks & Thesis Invalidation Triggers
    "entry_strategy",      # 14. Entry, Review, Exit Strategy
    "catalysts",           # 15. Upcoming Catalysts
    "peer_comparison",     # Peer Comparison (optional)
    "esg",                 # ESG (optional)
    "disclaimer",          # Fixed content, auto-appended
    "appendix",            # Appendix (optional)
]


class PageBrief(BaseModel):
    page_number: int
    page_type: PageType
    title: str
    key_message: str = Field(
        ...,
        description="One sentence: the single point this page must land.",
    )
    data_slices: list[str] = Field(
        default_factory=list,
        description="Dotted paths into CompanyPack.narrative this page should use.",
    )

    model_config = ConfigDict(extra="forbid")


class SpineOutline(BaseModel):
    thesis_north_star: str = Field(
        ...,
        description="The single argument the whole report defends. Every page reinforces this.",
    )
    tone: str = Field(default="institutional, data-led, measured")
    pages: list[PageBrief]

    model_config = ConfigDict(extra="forbid")


class PageEditorialSpec(BaseModel):
    page_number: int
    narrative_role: str = Field(
        ...,
        description="Why this page exists in the report arc and what job it performs.",
    )
    must_land: str = Field(
        ...,
        description="The exact analytical point the finished page must make.",
    )
    must_include: list[str] = Field(
        default_factory=list,
        description="Priority evidence or angles that should appear if space allows.",
    )
    avoid_repeating: list[str] = Field(
        default_factory=list,
        description="Points already owned by other pages and should not be repeated here.",
    )
    cut_first: list[str] = Field(
        default_factory=list,
        description="Items that should be dropped first if the page becomes crowded.",
    )
    preferred_blocks: list[str] = Field(
        default_factory=list,
        description="Preferred block mix for the page, e.g. chart+table+callout.",
    )
    max_blocks: int = 3

    model_config = ConfigDict(extra="forbid")


class EditorialPlan(BaseModel):
    document_hook: str = Field(
        ...,
        description="A concise editorial statement that sets the report's pacing and emphasis.",
    )
    compression_bias: str = Field(
        default="compress aggressively; omit lower-priority detail before expanding page count",
    )
    page_specs: list[PageEditorialSpec]

    model_config = ConfigDict(extra="forbid")


class PageRevisionRequest(BaseModel):
    page_number: int
    severity: Literal["low", "med", "high"] = "low"
    issues: list[str] = Field(default_factory=list)
    suggestions: str = ""

    model_config = ConfigDict(extra="forbid")


class DocumentReview(BaseModel):
    verdict: Literal["accept", "revise"] = "revise"
    global_issues: list[str] = Field(default_factory=list)
    page_revisions: list[PageRevisionRequest] = Field(default_factory=list)

    model_config = ConfigDict(extra="forbid")


# ─── Page content (Phase 3 output) ────────────────────────────────────────────

class Bullet(BaseModel):
    title: str | None = None
    body: str
    model_config = ConfigDict(extra="forbid")


class TableBlock(BaseModel):
    kind: Literal["table"] = "table"
    title: str | None = None
    headers: list[str]
    rows: list[list[str]]
    footnote: str | None = None
    highlight_rows: list[int] = [0]  # 0-indexed data rows to highlight; default = first row only
    model_config = ConfigDict(extra="forbid")


class MetricItem(BaseModel):
    label: str
    value: str
    delta: str | None = None
    model_config = ConfigDict(extra="ignore")


class MetricsBlock(BaseModel):
    kind: Literal["metrics"] = "metrics"
    title: str | None = None
    items: list[MetricItem]
    model_config = ConfigDict(extra="forbid")


class BulletsBlock(BaseModel):
    kind: Literal["bullets"] = "bullets"
    title: str | None = None
    items: list[Bullet]
    model_config = ConfigDict(extra="forbid")


class ParagraphBlock(BaseModel):
    kind: Literal["paragraph"] = "paragraph"
    title: str | None = None
    body: str
    model_config = ConfigDict(extra="forbid")


class CalloutBlock(BaseModel):
    kind: Literal["callout"] = "callout"
    variant: Literal["thesis", "warn", "info", "quote"] = "info"
    title: str | None = None
    body: str
    model_config = ConfigDict(extra="forbid")


class ChartBlock(BaseModel):
    """Chart is rendered deterministically from CompanyPack.financials.
    The model only *requests* a chart by referencing metric names.
    """
    kind: Literal["chart"] = "chart"
    chart_type: Literal["bar", "line", "stacked_bar", "donut"] = "bar"
    title: str | None = None
    metrics: list[str] = Field(default_factory=list)
    years: list[str] | None = None
    # For donut: inline {label, value} pairs (segment mix etc.) when not derived from financials.
    slices: list[dict] | None = None
    model_config = ConfigDict(extra="forbid")


# ─── Rich block types (matching gold-standard component library) ──────────────

class CatalystItem(BaseModel):
    icon: str = "🔹"
    title: str
    badge: str | None = None
    body: str
    model_config = ConfigDict(extra="forbid")


class CatalystBlock(BaseModel):
    kind: Literal["catalyst"] = "catalyst"
    title: str | None = None
    items: list[CatalystItem]
    model_config = ConfigDict(extra="forbid")


class RiskItem(BaseModel):
    severity: Literal["high", "med", "low"] = "med"
    title: str
    body: str
    model_config = ConfigDict(extra="forbid")


class RiskBlock(BaseModel):
    kind: Literal["risk"] = "risk"
    title: str | None = None
    items: list[RiskItem]
    model_config = ConfigDict(extra="forbid")


class ScenarioCase(BaseModel):
    label: str
    case: Literal["bear", "base", "bull"]
    target_price: str
    updown: str | None = None
    description: str | None = None
    model_config = ConfigDict(extra="forbid")


class ScenarioBlock(BaseModel):
    kind: Literal["scenario"] = "scenario"
    title: str | None = None
    cases: list[ScenarioCase]
    model_config = ConfigDict(extra="forbid")


class ScoreItem(BaseModel):
    name: str
    score: str
    max_score: str = "15"
    description: str | None = None
    model_config = ConfigDict(extra="forbid")


class ScorecardBlock(BaseModel):
    kind: Literal["scorecard"] = "scorecard"
    title: str | None = None
    items: list[ScoreItem]
    model_config = ConfigDict(extra="forbid")


class TimelineEntry(BaseModel):
    year: str
    text: str
    model_config = ConfigDict(extra="forbid")


class TimelineBlock(BaseModel):
    kind: Literal["timeline"] = "timeline"
    title: str | None = None
    items: list[TimelineEntry]
    model_config = ConfigDict(extra="forbid")


ContentBlock = Annotated[
    TableBlock | MetricsBlock | BulletsBlock | ParagraphBlock | CalloutBlock
    | ChartBlock | CatalystBlock | RiskBlock | ScenarioBlock | ScorecardBlock
    | TimelineBlock,
    Field(discriminator="kind"),
]


class PageContent(BaseModel):
    page_number: int
    page_type: PageType
    title: str
    subtitle: str | None = None
    blocks: list[ContentBlock]
    model_config = ConfigDict(extra="forbid")
