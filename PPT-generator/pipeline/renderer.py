"""Phase 5: Deterministic renderer — Tikona Capital brand design system.

Layout rules (matching tikona_reliance_landscape.html reference):
  - Cover page  : three-col-wide  (main blocks | auto fin-table | sidebar)
  - Inner pages : two-col-side    (main blocks  | persistent sidebar)
  - Sidebar always has live market data — never empty.
  - paragraph blocks → highlight-box (rotating border colour)
  - bullets blocks   → 2-col grid of highlight-box cards
  - table blocks     → fin-table-wrap with tbl-section dividers
"""

from __future__ import annotations

import html
import re
from pathlib import Path

from .schemas import (
    BulletsBlock,
    CalloutBlock,
    CatalystBlock,
    ChartBlock,
    CompanyPack,
    ContentBlock,
    MetricsBlock,
    PageContent,
    ParagraphBlock,
    RiskBlock,
    ScenarioBlock,
    ScorecardBlock,
    TableBlock,
    TimelineBlock,
)


# ─── CSS loading ──────────────────────────────────────────────────────────────

def load_brand_css() -> str:
    css_path = Path(__file__).parent / "tikona_brand.css"
    try:
        return css_path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return ""


def load_reference_css(_reference_html_path: str | Path) -> str:
    """Legacy shim — kept so orchestrate.py doesn't break."""
    return ""


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _e(s) -> str:
    return html.escape(str(s if s is not None else ""), quote=True)


def _fmt_num(v) -> str:
    if v is None:
        return "—"
    if isinstance(v, (int, float)):
        if abs(v) < 1 and v != 0:
            return f"{v * 100:.1f}%"
        return f"{v:,.0f}"
    return str(v)


def _fmt_price(v) -> str:
    if v is None or v == "":
        return "—"
    if isinstance(v, str):
        m = re.fullmatch(r"\s*₹?\s*([\d,]+(?:\.\d+)?)\s*", v)
        if not m:
            return v
        v = float(m.group(1).replace(",", ""))
    if isinstance(v, (int, float)):
        return _fmt_num(v)
    return str(v)


# ─── Block renderers ──────────────────────────────────────────────────────────

def _present(v) -> bool:
    if v is None:
        return False
    s = str(v).strip()
    return s not in {"", "—", "-", "None", "null"}


def _html_rupee(v) -> str:
    return f"&#8377;{_e(_fmt_price(v))}" if _present(v) else "—"


def _first_present(*values) -> str | None:
    for v in values:
        if _present(v):
            return str(v)
    return None


_BORDER_CYCLE = ["accent-border", "navy-border", "green-border", "accent-border",
                 "navy-border", "red-border"]


def _render_paragraph(b: ParagraphBlock, idx: int = 0) -> str:
    border = _BORDER_CYCLE[idx % len(_BORDER_CYCLE)]
    title_html = (
        f'<div class="highlight-box-title">{_e(b.title)}</div>' if b.title else ""
    )
    return (
        f'<div class="highlight-box {border}">'
        f'{title_html}<p>{_e(b.body)}</p>'
        f'</div>'
    )


def _render_bullets(b: BulletsBlock, idx: int = 0) -> str:
    head = (
        f'<div class="key-section-title">{_e(b.title)}</div>' if b.title else ""
    )
    items = b.items[:4]

    # 1 item — plain highlight box
    if len(items) == 1:
        border = _BORDER_CYCLE[idx % len(_BORDER_CYCLE)]
        lead = (
            f'<span class="analysis-bold">{_e(items[0].title)}: </span>'
            if items[0].title else ""
        )
        return (
            f'{head}<div class="highlight-box {border}">'
            f'<p>{lead}{_e(items[0].body)}</p></div>'
        )

    # 2+ items — 2-col grid of highlight boxes
    cards = []
    for i, it in enumerate(items):
        border = _BORDER_CYCLE[i % len(_BORDER_CYCLE)]
        title_html = (
            f'<div class="highlight-box-title">{_e(it.title)}</div>' if it.title else ""
        )
        cards.append(
            f'<div class="highlight-box {border}">{title_html}<p>{_e(it.body)}</p></div>'
        )

    # pack into 2-col grid
    grid_cols = 2 if len(cards) >= 2 else 1
    inner = "".join(cards)
    return (
        f'{head}'
        f'<div style="display:grid;grid-template-columns:repeat({grid_cols},1fr);gap:6px;">'
        f'{inner}</div>'
    )


# Section-divider keywords — rows whose first cell matches these become tbl-section rows
_SECTION_KEYWORDS = {
    "income statement", "revenue", "p&l", "profit & loss",
    "per share data", "per share", "valuation multiples", "valuation",
    "return ratios", "returns", "balance sheet", "cash flow",
    "key metrics", "operating metrics", "segment",
}


def _is_section_row(row: list[str]) -> bool:
    if not row:
        return False
    first = str(row[0]).strip().lower()
    return first in _SECTION_KEYWORDS or (len(row) == 1)


def _render_table(b: TableBlock) -> str:
    title = (
        f'<div class="fin-table-title">{_e(b.title)}</div>' if b.title else ""
    )
    headers = b.headers[:7]
    ths = "".join(
        f'<th{"" if i else " style=\"text-align:left;\""}'
        f'>{_e(h)}</th>'
        for i, h in enumerate(headers)
    )

    rows = b.rows[:12]
    trs = []
    last = len(rows) - 1
    for ri, row in enumerate(rows):
        cells = row[:7]

        # Section divider row
        if _is_section_row(cells):
            colspan = max(len(headers), len(cells))
            trs.append(
                f'<tr class="tbl-section">'
                f'<td colspan="{colspan}">{_e(cells[0])}</td>'
                f'</tr>'
            )
            continue

        is_highlight = ri in (b.highlight_rows or [0])
        row_cls = ' class="tbl-highlight"' if is_highlight else ""
        tds = []
        for ci, c in enumerate(cells):
            txt = str(c)
            if ci == 0:
                cls = ' class="bold"' if is_highlight else ""
            elif is_highlight:
                cls = ' class="bold"'
            elif txt.endswith("%") and not txt.startswith("-"):
                cls = ' class="green"'
            elif txt.startswith("-") and txt.endswith("%"):
                cls = ' class="red"'
            else:
                cls = ""
            tds.append(f"<td{cls}>{_e(c)}</td>")
        trs.append(f"<tr{row_cls}>{''.join(tds)}</tr>")

    foot = (
        f'<div class="figure-source">{_e(b.footnote)}</div>' if b.footnote else ""
    )
    return (
        f'<div class="fin-table-wrap">{title}'
        f'<table><thead><tr>{ths}</tr></thead>'
        f'<tbody>{"".join(trs)}</tbody></table>'
        f'{foot}</div>'
    )


_METRIC_VARIANTS = ["", "accent-bg", "mid-bg", "", "teal-bg", "light-bg", ""]


def _render_metrics(b: MetricsBlock) -> str:
    items = b.items[:8]
    n = len(items)
    # Symmetric grid counts only
    if n <= 2:
        cols = 2
    elif n == 3:
        cols = 3
    elif n == 4:
        cols = 4
    elif n == 5:
        cols = 5
    elif n == 6:
        cols = 6
    else:
        cols = 4  # wrap to 4-col grid for 7-8 items

    cards = []
    for i, it in enumerate(items):
        variant = _METRIC_VARIANTS[i % len(_METRIC_VARIANTS)]
        cls = f" {variant}" if variant else ""
        sub = (
            f'<div class="metric-card-sub">{_e(it.delta)}</div>' if it.delta else ""
        )
        cards.append(
            f'<div class="metric-card{cls}">'
            f'<div class="metric-card-label">{_e(it.label)}</div>'
            f'<div class="metric-card-value">{_e(it.value)}</div>'
            f'{sub}</div>'
        )

    head = (
        f'<div class="key-section-title">{_e(b.title)}</div>' if b.title else ""
    )
    grid = (
        f'style="display:grid;grid-template-columns:repeat({cols},1fr);'
        f'gap:5px;margin-bottom:8px;"'
    )
    return f'{head}<div class="metric-strip" {grid}>{"".join(cards)}</div>'


def _render_callout(b: CalloutBlock) -> str:
    if b.variant == "thesis":
        lead = f"<strong>{_e(b.title)}. </strong>" if b.title else ""
        return f'<div class="thesis-box"><p>{lead}{_e(b.body)}</p></div>'
    border_map = {"warn": "red-border", "info": "accent-border", "quote": "navy-border"}
    border = border_map.get(b.variant, "accent-border")
    title_html = (
        f'<div class="highlight-box-title">{_e(b.title)}</div>' if b.title else ""
    )
    return f'<div class="highlight-box {border}">{title_html}<p>{_e(b.body)}</p></div>'


def _render_donut(b: ChartBlock) -> str:
    """Inline donut from b.slices = [{label, value}, ...]."""
    slices = [s for s in (b.slices or []) if s.get("value")]
    total = sum(float(s["value"]) for s in slices) or 1.0
    colors = ["#1F4690", "#FFA500", "#3A5BA0", "#0e7490", "#1a7a4a", "#b91c1c", "#6b7c93"]
    cx, cy, r, stroke = 45, 45, 34, 14
    circ = 2 * 3.14159265 * r
    offset = 0.0
    segs = []
    legend = []
    for i, s in enumerate(slices):
        pct = float(s["value"]) / total
        dash = pct * circ
        col = colors[i % len(colors)]
        segs.append(
            f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="transparent"'
            f' stroke="{col}" stroke-width="{stroke}"'
            f' stroke-dasharray="{dash:.2f} {circ:.2f}"'
            f' stroke-dashoffset="{-offset:.2f}"'
            f' transform="rotate(-90 {cx} {cy})"/>'
        )
        legend.append(
            f'<div class="legend-item"><span class="legend-dot" style="background:{col};"></span>'
            f'<span class="legend-label">{_e(s.get("label",""))}</span>'
            f'<span class="legend-pct">{pct*100:.0f}%</span></div>'
        )
        offset += dash
    title_html = f'<div class="figure-title">{_e(b.title)}</div>' if b.title else ""
    return (
        f'<div class="figure">{title_html}'
        f'<div class="donut-wrap">'
        f'<svg viewBox="0 0 90 90" style="width:110px;height:110px;">{"".join(segs)}</svg>'
        f'<div class="donut-legend">{"".join(legend)}</div>'
        f'</div></div>'
    )


def _render_chart(b: ChartBlock, pack: CompanyPack) -> str:
    if b.chart_type == "donut":
        return _render_donut(b)
    if not pack.financials:
        return ""
    # Use block-specified years only if at least one label actually exists in the
    # model; otherwise fall back to the model's own years to avoid silent no-data.
    if b.years and any(y in pack.financials.years for y in b.years):
        years = b.years
    else:
        years = pack.financials.years
    series: list[tuple[str, list]] = []
    for m in b.metrics:
        row = pack.financials.get(m)
        if row:
            series.append((m, [row.values.get(y) for y in years]))
    if not series:
        return ""

    W, H = 560, 130
    PAD_L, PAD_R, PAD_T, PAD_B = 44, 10, 10, 26
    plot_w = W - PAD_L - PAD_R
    plot_h = H - PAD_T - PAD_B

    all_vals = [v for _, vs in series for v in vs if v is not None]
    if not all_vals:
        return ""
    vmin = min(0, min(all_vals))
    vmax = max(all_vals)
    vrange = vmax - vmin or 1

    def y_of(v: float) -> float:
        return PAD_T + plot_h * (1 - (v - vmin) / vrange)

    colors = ["#1F4690", "#FFA500", "#3A5BA0", "#FFE5B4", "#0e7490"]
    parts = [
        f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg"'
        f' style="width:100%;height:auto;">'
    ]
    for i in range(4):
        v = vmin + vrange * i / 3
        y = y_of(v)
        parts.append(
            f'<line x1="{PAD_L}" x2="{W - PAD_R}" y1="{y:.1f}" y2="{y:.1f}"'
            f' stroke="#d5dce8" stroke-width="0.5"/>'
        )
        parts.append(
            f'<text x="{PAD_L - 4}" y="{y + 3:.1f}" text-anchor="end"'
            f' font-size="7" fill="#6b7c93"'
            f' font-family="JetBrains Mono,monospace">{_fmt_num(v)}</text>'
        )

    n_years = len(years)
    slot_w = plot_w / max(n_years, 1)

    if b.chart_type == "line":
        for si, (name, vs) in enumerate(series):
            color = colors[si % len(colors)]
            pts = []
            for xi, v in enumerate(vs):
                if v is None:
                    continue
                cx = PAD_L + slot_w * (xi + 0.5)
                pts.append(f"{cx:.1f},{y_of(v):.1f}")
            if pts:
                parts.append(
                    f'<polyline fill="none" stroke="{color}" stroke-width="1.8"'
                    f' points="{" ".join(pts)}"/>'
                )
    else:
        n_s = len(series)
        bar_w = slot_w / (n_s + 1)
        for si, (name, vs) in enumerate(series):
            color = colors[si % len(colors)]
            for xi, v in enumerate(vs):
                if v is None:
                    continue
                x = PAD_L + slot_w * xi + bar_w * (0.5 + si)
                y = y_of(v)
                h = y_of(vmin) - y
                parts.append(
                    f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_w * 0.85:.1f}"'
                    f' height="{h:.1f}" fill="{color}" rx="1"/>'
                )

    for xi, yr in enumerate(years):
        cx = PAD_L + slot_w * (xi + 0.5)
        parts.append(
            f'<text x="{cx:.1f}" y="{H - PAD_B + 11}" text-anchor="middle"'
            f' font-size="7" fill="#3a4d6b"'
            f' font-family="JetBrains Mono,monospace">{_e(yr)}</text>'
        )

    legend_parts = []
    for si, (name, _) in enumerate(series):
        legend_parts.append(
            f'<span class="legend-item">'
            f'<span class="legend-dot" style="background:{colors[si % len(colors)]};"></span>'
            f'<span class="legend-label">{_e(name)}</span></span>'
        )
    parts.append("</svg>")
    legend = (
        f'<div class="donut-legend" style="justify-content:flex-start;gap:10px;margin-top:4px;">'
        f'{"".join(legend_parts)}</div>'
    )
    title_html = (
        f'<div class="figure-title">{_e(b.title)}</div>' if b.title else ""
    )
    return f'<div class="figure">{title_html}{"".join(parts)}{legend}</div>'


def _render_catalyst(b: CatalystBlock) -> str:
    head = (
        f'<div class="key-section-title">{_e(b.title)}</div>' if b.title else ""
    )
    parts = []
    for it in b.items[:5]:
        badge = (
            f'<span class="catalyst-badge">{_e(it.badge)}</span>' if it.badge else ""
        )
        parts.append(
            f'<div class="catalyst-item">'
            f'<div class="catalyst-icon">{it.icon}</div>'
            f'<div class="catalyst-body">'
            f'<div class="catalyst-title">{_e(it.title)} {badge}</div>'
            f'<div class="catalyst-desc">{_e(it.body)}</div>'
            f'</div></div>'
        )
    return head + "".join(parts)


def _render_risk(b: RiskBlock) -> str:
    head = (
        f'<div class="key-section-title">{_e(b.title)}</div>' if b.title else ""
    )
    parts = []
    for it in b.items[:5]:
        sev_cls = f"risk-{it.severity}"
        parts.append(
            f'<div class="risk-item">'
            f'<div class="risk-badge {sev_cls}">{it.severity.upper()}</div>'
            f'<div><div class="risk-title">{_e(it.title)}</div>'
            f'<div class="risk-text">{_e(it.body)}</div></div>'
            f'</div>'
        )
    return head + "".join(parts)


def _parse_price(s) -> float | None:
    if s is None:
        return None
    try:
        return float(re.sub(r"[^\d.\-]", "", str(s)) or "nan")
    except ValueError:
        return None


def _render_scenario(b: ScenarioBlock, pack: CompanyPack) -> str:
    head = (
        f'<div class="key-section-title">{_e(b.title)}</div>' if b.title else ""
    )
    cmp_val = _parse_price(pack.cmp)
    cards = []
    for c in b.cases[:3]:
        tp_val = _parse_price(c.target_price)
        # Recompute updown from (CMP, TP) so arithmetic always reconciles (Bug #2).
        updown_str = c.updown or ""
        if cmp_val and tp_val and cmp_val > 0:
            pct = (tp_val - cmp_val) / cmp_val * 100
            sign = "+" if pct >= 0 else ""
            updown_str = f"{sign}{pct:.1f}% vs CMP"
        # Flag a "bear" case that is not actually bearish (Bug #9).
        warn = ""
        if c.case == "bear" and cmp_val and tp_val and tp_val >= cmp_val:
            warn = (
                '<div class="scenario-desc" style="color:#b91c1c;font-weight:700;">'
                '⚠ Bear TP ≥ CMP — not a downside case</div>'
            )
        desc = (
            f'<div class="scenario-desc">{_e(c.description)}</div>'
            if c.description else ""
        )
        updown = (
            f'<div class="scenario-updown">{_e(updown_str)}</div>'
            if updown_str else ""
        )
        cards.append(
            f'<div class="scenario-card {c.case}">'
            f'<div class="scenario-label">{_e(c.label)}</div>'
            f'<div class="scenario-tp">{_e(c.target_price)}</div>'
            f'{updown}{warn}{desc}</div>'
        )
    return head + f'<div class="scenario-grid">{"".join(cards)}</div>'


_SCORE_COLORS = [
    "#FFA500", "#1F4690", "#3A5BA0", "#0e7490", "#1a7a4a", "#6b7c93", "#b91c1c",
]


_SAARTHI_INTRO = (
    '<div class="highlight-box accent-border" style="margin-bottom:6px;">'
    '<div class="highlight-box-title">SAARTHI Framework — Tikona Capital Proprietary Scorecard</div>'
    '<p>Seven-dimension quality assessment (Scalability, Assets, Accounting, Returns, '
    'Track Record, Hygiene, Integrity) scored out of 15 each; composite maps to rating band below.</p>'
    '<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:4px;margin-top:5px;">'
    '<div style="background:#1a7a4a;color:#fff;font-family:\'JetBrains Mono\',monospace;'
    'font-size:5.5pt;padding:3px 5px;border-radius:2px;text-align:center;">'
    '85–100 · STRONG BUY</div>'
    '<div style="background:#3A5BA0;color:#fff;font-family:\'JetBrains Mono\',monospace;'
    'font-size:5.5pt;padding:3px 5px;border-radius:2px;text-align:center;">'
    '70–84 · BUY</div>'
    '<div style="background:#FFA500;color:#1F4690;font-family:\'JetBrains Mono\',monospace;'
    'font-size:5.5pt;padding:3px 5px;border-radius:2px;text-align:center;font-weight:700;">'
    '55–69 · HOLD</div>'
    '<div style="background:#b91c1c;color:#fff;font-family:\'JetBrains Mono\',monospace;'
    'font-size:5.5pt;padding:3px 5px;border-radius:2px;text-align:center;">'
    '40–54 · REDUCE</div>'
    '<div style="background:#4a0e0e;color:#fff;font-family:\'JetBrains Mono\',monospace;'
    'font-size:5.5pt;padding:3px 5px;border-radius:2px;text-align:center;">'
    '<40 · SELL</div>'
    '</div></div>'
)


def _render_scorecard(b: ScorecardBlock) -> str:
    head = _SAARTHI_INTRO + (
        f'<div class="fin-table-title">{_e(b.title)}</div>' if b.title else ""
    )
    parts = []
    for i, it in enumerate(b.items):
        try:
            pct = min(float(it.score) / float(it.max_score) * 100, 100)
        except (ValueError, ZeroDivisionError):
            pct = 50
        color = _SCORE_COLORS[i % len(_SCORE_COLORS)]
        desc = (
            f'<div class="score-desc">{_e(it.description)}</div>'
            if it.description else ""
        )
        parts.append(
            f'<div class="score-item">'
            f'<div class="score-header">'
            f'<span class="score-name">{_e(it.name)}</span>'
            f'<span class="score-val">{_e(it.score)} / {_e(it.max_score)}</span>'
            f'</div>'
            f'<div class="score-bar-bg">'
            f'<div class="score-bar-fill" style="width:{pct:.0f}%;background:{color};"></div>'
            f'</div>'
            f'{desc}</div>'
        )
    return head + "".join(parts)


def _render_timeline(b: TimelineBlock) -> str:
    head = (
        f'<div class="key-section-title">{_e(b.title)}</div>' if b.title else ""
    )
    parts = []
    for it in b.items[:12]:
        parts.append(
            f'<div class="timeline-item">'
            f'<div class="timeline-year">{_e(it.year)}</div>'
            f'<div class="timeline-text">{_e(it.text)}</div>'
            f'</div>'
        )
    return head + f'<div class="timeline">{"".join(parts)}</div>'


def render_block(block: ContentBlock, pack: CompanyPack, idx: int = 0) -> str:
    if isinstance(block, ParagraphBlock):
        return _render_paragraph(block, idx)
    if isinstance(block, BulletsBlock):
        return _render_bullets(block, idx)
    if isinstance(block, TableBlock):
        return _render_table(block)
    if isinstance(block, MetricsBlock):
        return _render_metrics(block)
    if isinstance(block, CalloutBlock):
        return _render_callout(block)
    if isinstance(block, ChartBlock):
        return _render_chart(block, pack)
    if isinstance(block, CatalystBlock):
        return _render_catalyst(block)
    if isinstance(block, RiskBlock):
        return _render_risk(block)
    if isinstance(block, ScenarioBlock):
        return _render_scenario(block, pack)
    if isinstance(block, ScorecardBlock):
        return _render_scorecard(block)
    if isinstance(block, TimelineBlock):
        return _render_timeline(block)
    return ""


# ─── Persistent sidebar (shown on EVERY page) ─────────────────────────────────

def _render_sidebar(pack: CompanyPack) -> str:
    """Compact market-data sidebar that hides empty fields instead of showing filler dashes."""
    nar = pack.narrative
    up = pack.upside_potential_pct
    up_s = f"+{up}%" if _present(up) else None

    saarthi_raw = nar.get("saarthi_score", "") or nar.get("saarthi_scorecard", {})
    if isinstance(saarthi_raw, dict):
        saarthi = saarthi_raw.get("total_score", "") or saarthi_raw.get("score", "")
    else:
        saarthi = saarthi_raw

    ratio_rows = []
    for label, key in [
        ("P/E (TTM)", "pe_ratio"),
        ("D/E Ratio", "debt_equity"),
        ("ROE", "roe"),
        ("Net D/EBITDA", "net_debt_ebitda"),
        ("Div Yield", "dividend_yield"),
    ]:
        val = nar.get(key)
        if not _present(val):
            continue
        ratio_rows.append(
            f'<div class="side-row"><span class="side-label">{label}</span>'
            f'<span class="side-value">{_e(str(val))}</span></div>'
        )

    market_cap = _first_present(
        nar.get("market_cap"),
        nar.get("market_cap_rs_cr"),
        nar.get("mcap"),
    )
    week_52_h = _first_present(nar.get("week_52_high"))
    week_52_l = _first_present(nar.get("week_52_low"))
    credit_rating = _first_present(nar.get("credit_rating"), nar.get("sp_rating"))
    credit_outlook = _first_present(nar.get("credit_outlook"), "Stable")

    top_rows = []
    if market_cap:
        top_rows.append(
            f'<div class="side-row"><span class="side-label">Market Cap</span>'
            f'<span class="side-value">{_e(market_cap)}</span></div>'
        )
    if week_52_h and week_52_l:
        top_rows.append(
            f'<div class="side-row"><span class="side-label">52W Range</span>'
            f'<span class="side-value">{_html_rupee(week_52_l)} - {_html_rupee(week_52_h)}</span></div>'
        )

    cards = [f"""
    <div class="side-panel-card">
      <div class="side-metric-title">Target Price</div>
      <div class="side-metric-value">{_html_rupee(pack.target_price)}</div>
      <div class="side-metric-sub">{_e(up_s or 'Upside TBD')} vs CMP {_html_rupee(pack.cmp)}</div>
      {('<div class="side-divider"></div>' + ''.join(top_rows)) if top_rows else ''}
      {('<div class="side-divider"></div><div class="side-section-hdr">Key Ratios</div>' + ''.join(ratio_rows)) if ratio_rows else ''}
    </div>"""]

    if _present(saarthi):
        cards.append(f"""
    <div class="side-panel-card light compact">
      <div class="side-section-hdr" style="color:var(--primary-dark);margin-top:0;">SAARTHI Score</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:15pt;font-weight:700;
                  color:var(--primary-dark);line-height:1.1;">{_e(str(saarthi))}</div>
      <div class="side-metric-sub dark">Proprietary Framework</div>
      <div class="side-divider dark"></div>
      <div class="side-row">
        <span class="side-label dark">Rating</span>
        <span class="side-value dark" style="color:var(--green);font-weight:700;">{_e(str(pack.rating or 'â€"'))}</span>
      </div>
    </div>""")

    if credit_rating:
        cards.append(f"""
    <div class="side-panel-card compact" style="background:linear-gradient(135deg,#1F4690 0%,#3A5BA0 100%);">
      <div class="side-section-hdr">Credit Rating</div>
      <div style="color:var(--accent);font-family:'JetBrains Mono',monospace;
                  font-size:14pt;font-weight:700;line-height:1.1;">{_e(credit_rating)}</div>
      <div class="side-row">
        <span class="side-label">Outlook</span>
        <span class="side-value">{_e(credit_outlook or 'Stable')}</span>
      </div>
    </div>""")

    return "".join(cards)


_FIN_SUMMARY_METRICS = [
    # Income statement
    "Revenue", "EBITDA", "EBITDA Margin", "PAT", "PAT Margin",
    # Per-share
    "EPS",
    # Valuation multiples
    "P/E", "P/B", "EV/EBITDA",
    # Growth rates (computed as YoY % from the rows above)
    "Revenue Growth", "EBITDA Growth", "PAT Growth",
]


def _render_cover_fin_table(pack: CompanyPack) -> str:
    if not pack.financials:
        return ""
    fin = pack.financials
    years = fin.years[-6:]  # last 6 years max

    growth_metrics = {"revenue growth", "ebitda growth", "pat growth"}

    # Pass 1 — collect absolute metric rows in order
    found: list[tuple[str, dict, bool]] = []  # (display_name, year→value, is_growth)
    abs_cache: dict[str, dict] = {}           # base_name → year→value (for growth calc)
    for want in _FIN_SUMMARY_METRICS:
        want_lower = want.lower()
        if want_lower in growth_metrics:
            continue
        for row in fin.rows:
            if want_lower in row.metric.lower():
                vals = {y: row.values.get(y) for y in years}
                found.append((want, vals, False))
                abs_cache[want_lower] = vals
                break

    # Pass 2 — compute YoY growth for Revenue Growth / EBITDA Growth / PAT Growth
    # Stored as decimals (0.15) so _fmt_num renders them as "15.0%"
    for want in _FIN_SUMMARY_METRICS:
        want_lower = want.lower()
        if want_lower not in growth_metrics:
            continue
        base_name = want_lower.replace(" growth", "")
        base_vals = next(
            (v for k, v in abs_cache.items() if base_name in k),
            None,
        )
        if base_vals is None:
            continue
        growth_vals: dict = {}
        for i, y in enumerate(years):
            if i == 0:
                growth_vals[y] = None
            else:
                prev_y = years[i - 1]
                curr = base_vals.get(y)
                prev = base_vals.get(prev_y)
                if curr is not None and prev is not None and prev != 0:
                    growth_vals[y] = (curr - prev) / abs(prev)
                else:
                    growth_vals[y] = None
        found.append((want, growth_vals, True))

    if not found:
        return ""

    ths = (
        '<th style="text-align:left;">Metric</th>'
        + "".join(f"<th>{_e(y)}</th>" for y in years)
    )

    # Map exact metric names (lowercased) to the section header that should
    # appear *before* that metric.  Use exact equality so "revenue" doesn't
    # accidentally match "revenue growth".
    section_triggers = {
        "revenue":        "Income Statement",
        "eps":            "Per Share Data",
        "p/e":            "Valuation Multiples",
        "revenue growth": "Growth Rates",
    }
    trs = []
    for ri, (metric, vals, is_growth) in enumerate(found):
        metric_lower = metric.lower()
        if ri > 0 and metric_lower in section_triggers:
            colspan = len(years) + 1
            trs.append(
                f'<tr class="tbl-section">'
                f'<td colspan="{colspan}">{section_triggers[metric_lower]}</td>'
                f'</tr>'
            )

        tds = [f'<td style="text-align:left;font-weight:600;">{_e(metric)}</td>']
        for y in years:
            v = vals.get(y)
            txt = "—" if v is None else _fmt_num(v)
            if txt.endswith("%"):
                cls = ' class="red"' if txt.startswith("-") else ' class="green"'
            else:
                cls = ""
            tds.append(f"<td{cls}>{txt}</td>")
        highlight = ri == 0
        row_cls = ' class="tbl-highlight"' if highlight else ""
        trs.append(f"<tr{row_cls}>{''.join(tds)}</tr>")

    unit_note = "₹ Crore unless stated"
    return (
        f'<div class="fin-table-wrap">'
        f'<div class="fin-table-title">Financial Summary ({unit_note})</div>'
        f'<table>'
        f'<thead><tr>{ths}</tr></thead>'
        f'<tbody>{"".join(trs)}</tbody>'
        f'</table>'
        f'<div class="figure-source">e = Tikona Capital estimates</div>'
        f'</div>'
    )


# ─── Page chrome ──────────────────────────────────────────────────────────────

def _mcap_category(pack: CompanyPack) -> str:
    nar = pack.narrative
    cat = nar.get("market_cap_category", "")
    if cat:
        return cat
    mcap = nar.get("market_cap", "")
    try:
        n = float(
            re.sub(r"[^\d.]", "", str(mcap).replace("Lakh Cr", "").replace("lakh cr", ""))
        )
        if n >= 50000:
            return "Large Cap"
        if n >= 10000:
            return "Mid Cap"
        if n >= 2000:
            return "Small Cap"
        return "SME"
    except (ValueError, TypeError):
        return "Large Cap"


def _render_analyst_strip(pack: CompanyPack) -> str:
    """Render coverage analysts from pack.narrative["analysts"] ONLY.

    No defaults, no placeholders — prevents bug #10 (generic names recycled).
    """
    analysts = pack.narrative.get("analysts") or []
    if not isinstance(analysts, list) or not analysts:
        return ""
    items = []
    for a in analysts[:4]:
        name = _e(a.get("name", ""))
        title = _e(a.get("title", "") or a.get("designation", ""))
        email = _e(a.get("email", ""))
        if not name:
            continue
        items.append(
            f'<div style="display:flex;flex-direction:column;gap:1px;">'
            f'<span class="analyst-name">{name}</span>'
            f'<span class="analyst-title">{title}</span>'
            f'<span class="analyst-contact">{email}</span>'
            f'</div>'
        )
    if not items:
        return ""
    return f'<div class="analyst-strip">{"".join(items)}</div>'


def _render_tagline(pack: CompanyPack) -> str:
    tagline = pack.tagline or ""
    if not tagline:
        return ""
    return (
        f'<div class="tagline-band">'
        f'<div class="tagline-text">{_e(tagline)}</div>'
        f'</div>'
    )


def _render_full_header(pack: CompanyPack, report_type: str = "Equity Research") -> str:
    mcap_cat = _mcap_category(pack)
    return f"""
    <div class="report-header">
      <div class="header-top">
        <div class="firm-brand">
          <div class="firm-logo">T</div>
          <div>
            <div class="firm-name">Tikona Capital</div>
            <div class="firm-tagline">Driven by Research, Built with Conviction</div>
          </div>
        </div>
        <div class="report-type-badge">{_e(report_type)}</div>
      </div>
      <div class="header-company-line">
        <span class="company-name-hdr">{_e(pack.company)}</span>
        <span class="rating-badge">{_e(pack.rating or '')}</span>
        <span class="sector-tag">{_e(pack.ticker or '')} &nbsp;|&nbsp; {_e(pack.sector or '')} &nbsp;|&nbsp; {_e(mcap_cat)}</span>
      </div>
      {_render_cover_stats_bar(pack)}
      <div class="report-date">Institutional Equity Research</div>
    </div>"""


def _market_cap_text(pack: CompanyPack) -> str:
    nar = pack.narrative
    return _first_present(
        nar.get("market_cap"),
        nar.get("market_cap_rs_cr"),
        nar.get("mcap"),
        "-",
    ) or "-"


def _render_cover_stats_bar(pack: CompanyPack) -> str:
    up = pack.upside_potential_pct
    up_s = f"+{up}%" if _present(up) else "-"
    stats = [
        ("CMP", _html_rupee(pack.cmp), "As of latest pack"),
        ("Target Price", _html_rupee(pack.target_price), "SOTP-led fair value"),
        ("Upside", _e(up_s), "vs current market price"),
        ("Market Cap", _e(_market_cap_text(pack)), _e(_mcap_category(pack))),
    ]
    if _present(pack.rating):
        stats.append(("Rating", _e(pack.rating), "Tikona view"))
    parts = []
    for label, value, sub in stats:
        value_cls = "hstat-value"
        if label == "Upside":
            value_cls += " green-v"
        elif label == "Market Cap":
            value_cls += " white"
        parts.append(
            f'<div class="hstat"><div class="hstat-label">{label}</div>'
            f'<div class="{value_cls}">{value}</div>'
            f'<div class="hstat-sub">{sub}</div></div>'
        )
    return f'<div class="header-stats-bar">{"".join(parts)}</div>'


def _render_mini_header(pack: CompanyPack, page_title: str = "") -> str:
    up = pack.upside_potential_pct
    up_s = f"+{up}%" if _present(up) else "-"
    stats = [
        ("Rating", _e(pack.rating or "-"), "buy"),
        ("CMP", _html_rupee(pack.cmp), ""),
        ("TP", _html_rupee(pack.target_price), "accent"),
        ("Upside", _e(up_s), "green-v"),
        ("Cap", _e(_mcap_category(pack)), ""),
    ]
    stat_html = []
    for label, value, extra_cls in stats:
        value_cls = "mini-stat-value"
        if extra_cls:
            value_cls += f" {extra_cls}"
        stat_html.append(
            f'<div class="mini-stat"><span class="mini-stat-label">{label}</span>'
            f'<span class="{value_cls}">{value}</span></div>'
        )
    return f"""
    <div class="mini-header">
      <div class="mini-header-left">
        <div class="mini-firm">Tikona Capital</div>
        <div class="mini-company">{_e(page_title)}</div>
      </div>
      <div class="mini-stats">{''.join(stat_html)}</div>
    </div>"""


_NUM_WORDS = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11,
    "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
}


def _reconcile_title_count(title: str, actual: int) -> str:
    """Rewrite a numeric count token in the title to match actual rendered items.

    Prevents "Eleven Risks" appearing above 5 rendered risk rows (Bug #4).
    """
    if not title or actual <= 0:
        return title

    def _repl_word(m: re.Match) -> str:
        word = m.group(0)
        if _NUM_WORDS.get(word.lower()) == actual:
            return word
        # Capitalize replacement to match context roughly.
        rep = {1: "One", 2: "Two", 3: "Three", 4: "Four", 5: "Five",
               6: "Six", 7: "Seven", 8: "Eight", 9: "Nine", 10: "Ten"}.get(actual, str(actual))
        return rep

    pattern = r"\b(" + "|".join(_NUM_WORDS.keys()) + r")\b"
    new_title = re.sub(pattern, _repl_word, title, flags=re.IGNORECASE)
    new_title = re.sub(r"\b(\d+)\b",
                       lambda m: str(actual) if int(m.group(1)) != actual else m.group(0),
                       new_title)
    return new_title


def _render_section_header(page: PageContent) -> str:
    title = page.title
    # Reconcile any count-word in title with actual risk/catalyst item count.
    for block in page.blocks:
        if isinstance(block, RiskBlock):
            title = _reconcile_title_count(title, min(len(block.items), 5))
            break
        if isinstance(block, CatalystBlock):
            title = _reconcile_title_count(title, min(len(block.items), 4))
            break
    return (
        f'<div class="section-header">'
        f'<span class="section-number">{page.page_number:02d}</span>'
        f'<span class="section-title">{_e(title)}</span>'
        f'</div>'
    )


def _render_footer(page: PageContent, pack: CompanyPack, total: int) -> str:
    return f"""
    <div class="page-footer">
      <div class="footer-left">Tikona Capital Finserv Pvt. Ltd. | SEBI Reg. Research Analyst</div>
      <div class="footer-center">{_e(pack.company)} — Equity Research</div>
      <div class="footer-right">Page {page.page_number} of {total}</div>
    </div>"""


# ─── Page assembly ────────────────────────────────────────────────────────────

def render_page(
    page: PageContent,
    pack: CompanyPack,
    total_pages: int,
    is_first: bool,
) -> str:
    footer = _render_footer(page, pack, total_pages)
    sidebar = _render_sidebar(pack)

    if is_first:
        header = _render_full_header(pack, "Equity Research — Initiating Coverage")
        tagline = _render_tagline(pack)
        analysts = _render_analyst_strip(pack)
        main_blocks = "".join(render_block(b, pack, i) for i, b in enumerate(page.blocks))
        fin_col = _render_cover_fin_table(pack)
        body = f"""
        {tagline}
        {analysts}
        <div class="three-col-wide" style="padding-bottom:48px;">
          <div class="main-stack">{main_blocks}</div>
          <div>{fin_col}</div>
          <div>{sidebar}</div>
        </div>"""

    elif page.page_type == "story_charts":
        # ── STORY_CHARTS: full-width 3×2 chart grid, no sidebar ──
        header = _render_mini_header(pack, page.title)
        sec_hdr = _render_section_header(page)
        chart_blocks = [b for b in page.blocks if isinstance(b, ChartBlock)]
        other_blocks = [b for b in page.blocks if not isinstance(b, ChartBlock)]
        chart_cells = "".join(
            f'<div class="story-chart-cell">{render_block(b, pack, i)}</div>'
            for i, b in enumerate(chart_blocks)
        )
        other_html = "".join(render_block(b, pack, i) for i, b in enumerate(other_blocks))
        body = f"""
        <div class="page-content pb48" style="padding:10px 18px 0;">
          {sec_hdr}
          {other_html}
          <div class="chart-grid-3">{chart_cells}</div>
        </div>"""

    else:
        # ── INNER: two-col-side (main content + persistent sidebar) ──
        header = _render_mini_header(pack, page.title)
        sec_hdr = "" if page.page_type == "cover" else _render_section_header(page)
        main_blocks = "".join(render_block(b, pack, i) for i, b in enumerate(page.blocks))
        body = f"""
        <div class="two-col-side pb48">
          <div class="main-col">
            <div class="page-content" style="padding:0;">
              {sec_hdr}
              <div class="main-stack">{main_blocks}</div>
            </div>
          </div>
          <div class="side-col">{sidebar}</div>
        </div>"""

    return f'<div class="page">{header}{body}{footer}</div>'


# ─── Document assembly ────────────────────────────────────────────────────────


_DISCLAIMER_TEXT = (
    "About Us: Research Analyst is registered with SEBI as Research Analyst with Registration "
    "No. INH000009807. The firm got its registration on June 13, 2022, and is engaged in research services.<br><br>"
    "Disciplinary history: No penalties / directions have been issued by SEBI under the SEBI Act or "
    "Regulations made there under. There are no pending material litigations or legal proceedings, "
    "findings of inspections or investigations for which action has been taken or initiated by any "
    "regulatory authority.<br><br>"
    "Details of its associates: No associates<br><br>"
    "Disclosures with respect to Research and Recommendations Services<br>"
    "· Research Analyst may have financial interest or actual / beneficial ownership in the securities "
    "recommended in its personal portfolio.<br>"
    "· There are no actual or potential conflicts of interest arising from any connection to or "
    "association with any issuer of products / securities.<br>"
    "· Research Analyst or its employee or its associates have not received any compensation from "
    "the company in past 12 months.<br>"
    "· Research Analyst or its employee or its associates have not managed or co-managed the public "
    "offering of Subject company in past 12 months.<br>"
    "· Research Analyst or its employee or its associates have not received any compensation for "
    "investment banking or merchant banking or brokerage services from the subject company in past "
    "12 months.<br>"
    "· Research Analyst or its employee or its associates have not received any compensation for "
    "products or services other than above from the subject company in past 12 months.<br>"
    "· Research Analyst or its employee or its associates have not received any compensation or "
    "other benefits from the Subject Company or 3rd party in connection with the research report.<br>"
    "· The subject company was not a client of Research Analyst during twelve months preceding the "
    "date of distribution of the research report.<br>"
    "· Research Analysts or its employee or its associates has not served as an officer, director, "
    "or employee of the subject company.<br>"
    "· Registration granted by SEBI, membership of BASL and certification from NISM in no way "
    "guarantee performance of the Intermediary or provide any assurance of returns to investors.<br>"
    "· Investment in securities market are subject to market risks. Read all the related documents "
    "carefully before investing."
)



def _render_disclaimer_page(pack: CompanyPack, header: str, footer: str, page_number: int) -> str:
    up = pack.upside_potential_pct
    up_s = f"+{up}%" if _present(up) else "—"
    analysts = pack.narrative.get("analysts") or []
    analyst_html = "".join(
        f'<div class="side-row"><span class="side-label dark">{_e(a.get("name", ""))}</span>'
        f'<span class="side-value dark">{_e(a.get("title", "") or a.get("designation", ""))}</span></div>'
        for a in analysts[:3]
        if a.get("name")
    )
    body = f"""
    <div class="page-content pb48">
      <div class="section-header">
        <span class="section-number">{page_number:02d}</span>
        <span class="section-title">Disclaimer &amp; Disclosures</span>
      </div>
      <div class="highlight-box accent-border" style="margin-bottom:6px;">
        <div class="highlight-box-title" style="font-size:8pt;">Recommendation Recap â€" {_e(pack.company)}</div>
        <div class="metric-strip" style="display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-top:4px;">
          <div class="metric-card light-bg"><div class="metric-card-label">Rating</div><div class="metric-card-value">{_e(pack.rating or 'â€"')}</div></div>
          <div class="metric-card"><div class="metric-card-label">CMP</div><div class="metric-card-value">{_html_rupee(pack.cmp)}</div></div>
          <div class="metric-card accent-bg"><div class="metric-card-label">Target Price</div><div class="metric-card-value">{_html_rupee(pack.target_price)}</div></div>
          <div class="metric-card teal-bg"><div class="metric-card-label">Upside</div><div class="metric-card-value">{_e(up_s)}</div></div>
        </div>
      </div>
      <div class="disclosure-grid">
        <div>
          <div class="disclosure-card">
            <div class="disclosure-title">Regulatory Disclosure</div>
            <div class="disclaimer" style="font-size:5.7pt;line-height:1.55;margin-top:0;">{_DISCLAIMER_TEXT}</div>
          </div>
        </div>
        <div>
          <div class="disclosure-card">
            <div class="disclosure-title">Research Desk</div>
            <div class="side-row"><span class="side-label dark">SEBI Reg.</span><span class="side-value dark">INH000009807</span></div>
            <div class="side-row"><span class="side-label dark">Email</span><span class="side-value dark">research@tikonacapital.in</span></div>
            <div class="side-row"><span class="side-label dark">Web</span><span class="side-value dark">tikonacapital.in</span></div>
            {('<div class="side-divider dark"></div>' + analyst_html) if analyst_html else ''}
          </div>
          <div class="disclosure-card" style="margin-top:8px;">
            <div class="disclosure-title">Reader Notes</div>
            <ul class="disclosure-list">
              <li>This report is generated from a structured company pack, deterministic financial model, and editor-reviewed page plan.</li>
              <li>Valuation, scenario, and recommendation pages should be read together because target price and risk framing are linked.</li>
              <li>If company disclosures or market prices move materially after publication, the recommendation framework should be refreshed.</li>
            </ul>
          </div>
        </div>
      </div>
    </div>"""
    return f'<div class="page">{header}{body}{footer}</div>'


def render_document(
    pages: list[PageContent],
    pack: CompanyPack,
    _css: str | None = None,     # ignored — kept for backward compat
    title: str | None = None,
) -> str:
    brand_css = load_brand_css()
    total = len(pages) + 1  # +1 for disclaimer

    body_parts = [
        render_page(p, pack, total, is_first=(i == 0))
        for i, p in enumerate(pages)
    ]

    # Disclaimer page — branded, SEBI-compliant closing layout.
    dis_header = _render_mini_header(pack, "Disclaimer & Disclosures")
    from .schemas import PageContent as _PC  # local import to avoid circular
    fake = _PC(page_number=total, page_type="disclaimer", title="Disclaimer", blocks=[])
    dis_footer = _render_footer(fake, pack, total)
    body_parts.append(_render_disclaimer_page(pack, dis_header, dis_footer, total))

    doc_title = title or f"{pack.company} — {pack.rating or ''} | Tikona Capital"
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<title>{_e(doc_title)}</title>
<style>{brand_css}</style>
</head>
<body>
{"".join(body_parts)}
</body>
</html>"""
