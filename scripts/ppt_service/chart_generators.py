"""
Generate Tikona-branded chart images (PNG bytes) from financial model JSON.

Shape naming convention in master_template.pptx
------------------------------------------------
Name shapes in PowerPoint's Selection Pane using these exact prefixes:

  chart:revenue_trend       grouped bar — Revenue/EBITDA/PAT historical+projected
  chart:margin_trend        line — EBITDA Margin + PAT Margin
  chart:roe_roce            line — ROE + ROCE
  chart:pat_trend           bar  — PAT only
  chart:ebitda_trend        bar  — EBITDA only
  chart:peer_comparison     horizontal bar — peer EV/EBITDA vs subject
  chart:valuation_waterfall bar waterfall — DCF / PE / EV-EBITDA / Blended values
  chart:sensitivity_heatmap colour grid — PE sensitivity from valuation.sensitivity_pe
  chart:saarthi_radar       spider/radar — 7 SAARTHI dimensions
  chart:debt_equity         bar — D/E ratio over historical years
  chart:cash_flow           stacked bar — CFO/CFI/CFF

Usage
-----
  from chart_generators import generate_chart
  img_bytes = generate_chart("revenue_trend", fin_model, company_name="Reliance")
  # returns PNG bytes, or None if data absent / matplotlib missing
"""

from __future__ import annotations

import io
import logging
from typing import Any

logger = logging.getLogger(__name__)

# ── Brand palette ──────────────────────────────────────────────────────────────
_C = {
    "navy":      "#1F4690",
    "blue":      "#3A5BA0",
    "orange":    "#FFA500",
    "teal":      "#0e7490",
    "green":     "#1a7a4a",
    "red":       "#b91c1c",
    "grey":      "#6b7c93",
    "light_bg":  "#f8f9fc",
    "grid":      "#d5dce8",
}

_SERIES_COLORS = [_C["navy"], _C["orange"], _C["teal"], _C["green"], _C["blue"], _C["red"]]


# ── matplotlib lazy loader ─────────────────────────────────────────────────────

def _mpl():
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import numpy as np
        return plt, np
    except ImportError:
        logger.warning("matplotlib not installed — chart image generation disabled")
        return None, None


def _to_png(fig) -> bytes:
    import matplotlib.pyplot as _plt
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    _plt.close(fig)
    return buf.getvalue()


def _style(fig, ax, title: str = ""):
    fig.patch.set_facecolor(_C["light_bg"])
    ax.set_facecolor(_C["light_bg"])
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_color(_C["grid"])
    ax.spines["bottom"].set_color(_C["grid"])
    ax.tick_params(colors=_C["grey"], labelsize=7.5)
    ax.grid(axis="y", color=_C["grid"], linewidth=0.5, alpha=0.7, zorder=0)
    ax.set_axisbelow(True)
    if title:
        ax.set_title(title, fontsize=8.5, color=_C["navy"], fontweight="bold", pad=6)


def _safe_float(v) -> float | None:
    if v is None:
        return None
    try:
        return float(str(v).replace(",", ""))
    except (ValueError, TypeError):
        return None


def _fill_none(lst: list, n: int) -> list:
    return (list(lst) + [None] * n)[:n]


# ── Chart generators ───────────────────────────────────────────────────────────

def _revenue_trend(fin_model: dict, company_name: str) -> bytes | None:
    """Grouped bar: Revenue / EBITDA / PAT across historical + projected years."""
    plt, np = _mpl()
    if plt is None:
        return None

    hist = fin_model.get("historical_pl") or {}
    proj = fin_model.get("projected_pl") or {}
    h_yrs = [str(y) for y in (hist.get("years") or [])]
    p_yrs = [str(y) for y in (proj.get("years") or [])]
    years = h_yrs + p_yrs
    if not years:
        return None

    n = len(years)
    revenue = [_safe_float(v) or 0 for v in _fill_none(
        list(hist.get("revenue") or []) + list(proj.get("revenue") or []), n)]
    ebitda  = [_safe_float(v) or 0 for v in _fill_none(
        list(hist.get("ebitda") or []) + list(proj.get("ebitda") or []), n)]
    pat     = [_safe_float(v) or 0 for v in _fill_none(
        list(hist.get("pat") or []) + list(proj.get("pat") or []), n)]

    x = np.arange(n)
    w = 0.26
    fig, ax = plt.subplots(figsize=(8, 3.4))

    b1 = ax.bar(x - w, revenue, w, label="Revenue", color=_C["navy"],   zorder=3)
    b2 = ax.bar(x,     ebitda,  w, label="EBITDA",  color=_C["orange"],  zorder=3)
    b3 = ax.bar(x + w, pat,     w, label="PAT",     color=_C["teal"],    zorder=3)

    n_hist = len(h_yrs)
    for bars in (b1, b2, b3):
        for i, bar in enumerate(bars):
            if i >= n_hist:
                bar.set_hatch("///")
                bar.set_alpha(0.70)

    ax.set_xticks(x)
    ax.set_xticklabels(years, fontsize=7)
    ax.yaxis.set_major_formatter(
        __import__("matplotlib").ticker.FuncFormatter(lambda v, _: f"₹{v:,.0f}"))
    _style(fig, ax, f"{company_name} — Revenue / EBITDA / PAT  (₹ Cr)")
    ax.legend(fontsize=7, loc="upper left", framealpha=0)
    ax.annotate("/// = estimates", xy=(1, -0.14), xycoords="axes fraction",
                ha="right", fontsize=6, color=_C["grey"])
    fig.tight_layout(pad=0.6)
    return _to_png(fig)


def _margin_trend(fin_model: dict, company_name: str) -> bytes | None:
    """Line chart: EBITDA Margin + PAT Margin over historical years."""
    plt, np = _mpl()
    if plt is None:
        return None

    hist = fin_model.get("historical_ratios") or {}
    years = [str(y) for y in (hist.get("years") or [])]
    ebitda_m = [_safe_float(v) for v in (hist.get("ebitda_margin_pct") or [])]
    pat_m    = [_safe_float(v) for v in (hist.get("pat_margin_pct") or [])]

    # Also try from series in fin_model
    if not years:
        for s in (fin_model.get("series") or []):
            name = (s.get("name") or "").lower()
            if "ebitda margin" in name:
                years = s.get("periods") or []
                ebitda_m = [_safe_float(v) for v in (s.get("values") or [])]
            elif "pat margin" in name:
                pat_m = [_safe_float(v) for v in (s.get("values") or [])]

    if not years:
        return None

    n = len(years)
    ebitda_m = _fill_none(ebitda_m, n)
    pat_m    = _fill_none(pat_m, n)

    fig, ax = plt.subplots(figsize=(7, 3))
    x = range(n)
    ax.plot(x, ebitda_m, marker="o", markersize=5, linewidth=2,
            color=_C["navy"], label="EBITDA Margin %")
    ax.plot(x, pat_m, marker="s", markersize=5, linewidth=2,
            color=_C["orange"], label="PAT Margin %")

    for i, (em, pm) in enumerate(zip(ebitda_m, pat_m)):
        if em is not None:
            ax.annotate(f"{em:.1f}%", (i, em), textcoords="offset points",
                        xytext=(0, 6), ha="center", fontsize=6, color=_C["navy"])
        if pm is not None:
            ax.annotate(f"{pm:.1f}%", (i, pm), textcoords="offset points",
                        xytext=(0, -10), ha="center", fontsize=6, color=_C["orange"])

    ax.set_xticks(list(x))
    ax.set_xticklabels(years, fontsize=7)
    ax.yaxis.set_major_formatter(
        __import__("matplotlib").ticker.FuncFormatter(lambda v, _: f"{v:.0f}%"))
    _style(fig, ax, f"{company_name} — Margin Trend (%)")
    ax.legend(fontsize=7, framealpha=0)
    fig.tight_layout(pad=0.6)
    return _to_png(fig)


def _roe_roce(fin_model: dict, company_name: str) -> bytes | None:
    """Line chart: ROE + ROCE over historical years."""
    plt, np = _mpl()
    if plt is None:
        return None

    hist = fin_model.get("historical_ratios") or {}
    years = [str(y) for y in (hist.get("years") or [])]
    roe  = [_safe_float(v) for v in (hist.get("roe_pct") or [])]
    roce = [_safe_float(v) for v in (hist.get("roce_pct") or [])]

    if not years:
        return None
    n = len(years)
    roe  = _fill_none(roe, n)
    roce = _fill_none(roce, n)

    fig, ax = plt.subplots(figsize=(7, 3))
    x = range(n)
    ax.plot(x, roe,  marker="o", markersize=5, lw=2, color=_C["green"],  label="ROE %")
    ax.plot(x, roce, marker="D", markersize=5, lw=2, color=_C["blue"],   label="ROCE %")

    ax.set_xticks(list(x))
    ax.set_xticklabels(years, fontsize=7)
    ax.yaxis.set_major_formatter(
        __import__("matplotlib").ticker.FuncFormatter(lambda v, _: f"{v:.0f}%"))
    _style(fig, ax, f"{company_name} — ROE / ROCE (%)")
    ax.legend(fontsize=7, framealpha=0)
    fig.tight_layout(pad=0.6)
    return _to_png(fig)


def _pat_trend(fin_model: dict, company_name: str) -> bytes | None:
    """Bar chart: PAT historical + projected."""
    plt, np = _mpl()
    if plt is None:
        return None

    hist = fin_model.get("historical_pl") or {}
    proj = fin_model.get("projected_pl") or {}
    h_yrs = [str(y) for y in (hist.get("years") or [])]
    p_yrs = [str(y) for y in (proj.get("years") or [])]
    years = h_yrs + p_yrs
    if not years:
        return None

    n = len(years)
    pat = [_safe_float(v) or 0 for v in _fill_none(
        list(hist.get("pat") or []) + list(proj.get("pat") or []), n)]

    x = range(n)
    fig, ax = plt.subplots(figsize=(7, 3))
    bars = ax.bar(x, pat, color=_C["teal"], zorder=3)
    n_hist = len(h_yrs)
    for i, bar in enumerate(bars):
        if i >= n_hist:
            bar.set_hatch("///")
            bar.set_alpha(0.70)

    ax.set_xticks(list(x))
    ax.set_xticklabels(years, fontsize=7)
    ax.yaxis.set_major_formatter(
        __import__("matplotlib").ticker.FuncFormatter(lambda v, _: f"₹{v:,.0f}"))
    _style(fig, ax, f"{company_name} — PAT (₹ Cr)")
    fig.tight_layout(pad=0.6)
    return _to_png(fig)


def _ebitda_trend(fin_model: dict, company_name: str) -> bytes | None:
    """Bar chart: EBITDA historical + projected."""
    plt, np = _mpl()
    if plt is None:
        return None

    hist = fin_model.get("historical_pl") or {}
    proj = fin_model.get("projected_pl") or {}
    h_yrs = [str(y) for y in (hist.get("years") or [])]
    p_yrs = [str(y) for y in (proj.get("years") or [])]
    years = h_yrs + p_yrs
    if not years:
        return None

    n = len(years)
    ebitda = [_safe_float(v) or 0 for v in _fill_none(
        list(hist.get("ebitda") or []) + list(proj.get("ebitda") or []), n)]

    x = range(n)
    fig, ax = plt.subplots(figsize=(7, 3))
    bars = ax.bar(x, ebitda, color=_C["orange"], zorder=3)
    n_hist = len(h_yrs)
    for i, bar in enumerate(bars):
        if i >= n_hist:
            bar.set_hatch("///")
            bar.set_alpha(0.70)

    ax.set_xticks(list(x))
    ax.set_xticklabels(years, fontsize=7)
    ax.yaxis.set_major_formatter(
        __import__("matplotlib").ticker.FuncFormatter(lambda v, _: f"₹{v:,.0f}"))
    _style(fig, ax, f"{company_name} — EBITDA (₹ Cr)")
    fig.tight_layout(pad=0.6)
    return _to_png(fig)


def _peer_comparison(fin_model: dict, company_name: str) -> bytes | None:
    """Horizontal bar: peer EV/EBITDA and P/E vs subject company."""
    plt, np = _mpl()
    if plt is None:
        return None

    peers = fin_model.get("peers") or []
    if not peers:
        return None

    names   = [company_name] + [p.get("name", "") for p in peers[:6]]
    ev_vals = [None] + [_safe_float(p.get("ev_ebitda")) for p in peers[:6]]
    pe_vals = [None] + [_safe_float(p.get("pe")) for p in peers[:6]]

    # Subject company valuation from model metrics
    met = fin_model.get("metrics") or {}
    val = fin_model.get("valuation") or {}
    ev_vals[0] = _safe_float(met.get("ev_ebitda_fair_value") or val.get("ev_ebitda_fair_value"))
    pe_vals[0] = _safe_float(met.get("pe_fair_value") or val.get("pe_fair_value"))

    n = len(names)
    y = np.arange(n)
    w = 0.38

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(9, max(2.5, n * 0.45 + 0.5)))
    colors = [_C["orange"] if i == 0 else _C["navy"] for i in range(n)]

    def _vals(lst):
        return [v if v is not None else 0 for v in lst]

    ax1.barh(y, _vals(ev_vals), color=colors, zorder=3)
    ax1.set_yticks(y); ax1.set_yticklabels(names, fontsize=7.5)
    ax1.set_xlabel("EV/EBITDA (x)", fontsize=7.5)
    _style(fig, ax1, "EV/EBITDA Comparison")
    ax1.grid(axis="x"); ax1.grid(axis="y", alpha=0)
    ax1.invert_yaxis()

    ax2.barh(y, _vals(pe_vals), color=colors, zorder=3)
    ax2.set_yticks(y); ax2.set_yticklabels([], fontsize=7.5)
    ax2.set_xlabel("P/E (x)", fontsize=7.5)
    _style(fig, ax2, "P/E Comparison")
    ax2.grid(axis="x"); ax2.grid(axis="y", alpha=0)
    ax2.invert_yaxis()

    ax1.barh(y[0:1], _vals(ev_vals)[0:1], color=_C["orange"], label=company_name, zorder=4)
    ax1.legend(fontsize=7, framealpha=0, loc="lower right")

    fig.suptitle("Peer Valuation Comparison", fontsize=9, color=_C["navy"],
                 fontweight="bold", y=1.01)
    fig.tight_layout(pad=0.6)
    return _to_png(fig)


def _valuation_waterfall(fin_model: dict, company_name: str) -> bytes | None:
    """Bar chart: DCF / PE / EV-EBITDA / Blended fair values + CMP baseline."""
    plt, np = _mpl()
    if plt is None:
        return None

    val = fin_model.get("valuation") or {}
    met = fin_model.get("metrics") or {}
    cmp_val = _safe_float(met.get("cmp") or fin_model.get("cmp"))

    methods = []
    values  = []
    for label, keys in [
        ("DCF",       ["dcf_fair_value"]),
        ("P/E",       ["pe_fair_value"]),
        ("EV/EBITDA", ["ev_ebitda_fair_value"]),
        ("Blended",   ["blended_fair_value"]),
    ]:
        for k in keys:
            v = _safe_float(val.get(k) or met.get(k))
            if v:
                methods.append(label)
                values.append(v)
                break

    if not methods:
        return None

    # Append CMP as reference bar
    if cmp_val:
        methods.append("CMP")
        values.append(cmp_val)

    n = len(methods)
    colors = [(_C["orange"] if m == "Blended" else
               _C["grey"]   if m == "CMP"     else
               _C["navy"]) for m in methods]

    fig, ax = plt.subplots(figsize=(max(5, n * 1.2), 3.5))
    bars = ax.bar(range(n), values, color=colors, width=0.55, zorder=3)

    for i, (bar, v) in enumerate(zip(bars, values)):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.01,
                f"₹{v:,.0f}", ha="center", va="bottom", fontsize=7.5, color=_C["navy"],
                fontweight="bold")

    if cmp_val:
        ax.axhline(cmp_val, color=_C["red"], linestyle="--", linewidth=1,
                   label=f"CMP ₹{cmp_val:,.0f}")
        ax.legend(fontsize=7, framealpha=0)

    ax.set_xticks(range(n))
    ax.set_xticklabels(methods, fontsize=8)
    ax.yaxis.set_major_formatter(
        __import__("matplotlib").ticker.FuncFormatter(lambda v, _: f"₹{v:,.0f}"))
    _style(fig, ax, f"{company_name} — Valuation Summary (₹ per share)")
    fig.tight_layout(pad=0.6)
    return _to_png(fig)


def _sensitivity_heatmap(fin_model: dict, company_name: str) -> bytes | None:
    """Colour grid: PE sensitivity from valuation.sensitivity_pe."""
    plt, np = _mpl()
    if plt is None:
        return None

    val = fin_model.get("valuation") or {}
    sens = val.get("sensitivity_pe") or {}
    grid = sens.get("grid")
    row_vals = sens.get("row_values") or []
    col_vals = sens.get("col_values") or []
    row_label = sens.get("row_label", "PE Multiple")
    col_label = sens.get("col_label", "EPS Growth %")

    if not grid or not row_vals or not col_vals:
        return None

    try:
        data = np.array([[_safe_float(c) or 0 for c in row] for row in grid], dtype=float)
    except Exception:
        return None

    cmp_val = _safe_float((fin_model.get("metrics") or {}).get("cmp") or fin_model.get("cmp"))

    fig, ax = plt.subplots(figsize=(max(5, len(col_vals) * 0.9), max(3, len(row_vals) * 0.7)))
    import matplotlib.colors as mcolors
    cmap = mcolors.LinearSegmentedColormap.from_list(
        "tikona", [_C["red"], "#ffffff", _C["green"]], N=256)
    im = ax.imshow(data, cmap=cmap, aspect="auto")

    ax.set_xticks(range(len(col_vals)))
    ax.set_xticklabels([f"{v}%" for v in col_vals], fontsize=7.5)
    ax.set_yticks(range(len(row_vals)))
    ax.set_yticklabels([f"{v}x" for v in row_vals], fontsize=7.5)
    ax.set_xlabel(col_label, fontsize=8)
    ax.set_ylabel(row_label, fontsize=8)

    for i in range(len(row_vals)):
        for j in range(len(col_vals)):
            v = data[i, j]
            txt_color = "white" if abs(v - data.mean()) > data.std() else _C["navy"]
            ax.text(j, i, f"₹{v:,.0f}", ha="center", va="center",
                    fontsize=7, color=txt_color, fontweight="bold")
            if cmp_val and abs(v - cmp_val) <= cmp_val * 0.05:
                ax.add_patch(__import__("matplotlib.patches", fromlist=["Rectangle"])
                             .Rectangle((j - 0.5, i - 0.5), 1, 1,
                                        fill=False, edgecolor=_C["orange"], lw=2))

    fig.colorbar(im, ax=ax, shrink=0.8, label="Fair Value (₹)")
    ax.set_title(f"{company_name} — PE Sensitivity Matrix", fontsize=8.5,
                 color=_C["navy"], fontweight="bold", pad=6)
    fig.patch.set_facecolor(_C["light_bg"])
    fig.tight_layout(pad=0.6)
    return _to_png(fig)


def _saarthi_radar(fin_model: dict, company_name: str) -> bytes | None:
    """Spider/radar chart: 7 SAARTHI dimensions."""
    plt, np = _mpl()
    if plt is None:
        return None

    saarthi = fin_model.get("saarthi") or fin_model.get("thesis", {}).get("saarthi_scores") or {}

    # Normalised to 0-100 scale per dimension
    if isinstance(saarthi, dict) and "dimensions" in saarthi:
        dims = saarthi["dimensions"]
        labels = [d.get("name", d.get("code", "?")) for d in dims]
        scores = [min(float(d.get("score", 0)) / float(d.get("max_score", 15)) * 100, 100)
                  for d in dims]
    else:
        # Flat dict: S_sector_quality: 12, A_accounting_quality: 10 ...
        key_map = {
            "S_sector_quality":    "S — Sector",
            "A_accounting_quality":"A — Accounting",
            "A_asset_quality":     "A — Assets",
            "R_revenue_visibility":"R — Revenue",
            "T_track_record":      "T — Track Record",
            "H_balance_sheet_health":"H — Balance Sheet",
            "I_intrinsic_valuation":"I — Intrinsic Val",
        }
        labels = []
        scores = []
        for k, label in key_map.items():
            v = saarthi.get(k)
            if v is not None:
                labels.append(label)
                scores.append(min(float(v) / 15 * 100, 100))

    if len(labels) < 3:
        return None

    n = len(labels)
    angles = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist()
    angles += angles[:1]
    scores_plot = scores + scores[:1]

    fig, ax = plt.subplots(figsize=(5, 5), subplot_kw=dict(polar=True))
    fig.patch.set_facecolor(_C["light_bg"])
    ax.set_facecolor(_C["light_bg"])

    for level in [25, 50, 75, 100]:
        ax.plot(angles, [level] * (n + 1), color=_C["grid"], linewidth=0.5, linestyle="--")

    ax.fill(angles, scores_plot, alpha=0.25, color=_C["navy"])
    ax.plot(angles, scores_plot, color=_C["navy"], linewidth=2, marker="o", markersize=5)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, size=7.5, color=_C["navy"])
    ax.set_yticks([25, 50, 75, 100])
    ax.set_yticklabels(["25", "50", "75", "100"], size=6, color=_C["grey"])
    ax.set_ylim(0, 100)
    ax.spines["polar"].set_color(_C["grid"])

    total = fin_model.get("saarthi", {}).get("total_score") if isinstance(saarthi, dict) else None
    title = f"{company_name} — SAARTHI Score"
    if total:
        title += f"  ({total}/100)"
    ax.set_title(title, fontsize=8.5, color=_C["navy"], fontweight="bold", pad=14)
    fig.tight_layout()
    return _to_png(fig)


def _debt_equity(fin_model: dict, company_name: str) -> bytes | None:
    """Bar chart: Debt/Equity ratio over historical years."""
    plt, np = _mpl()
    if plt is None:
        return None

    hist = fin_model.get("historical_ratios") or {}
    years = [str(y) for y in (hist.get("years") or [])]
    de    = [_safe_float(v) for v in (hist.get("debt_equity") or [])]

    if not years or not any(v is not None for v in de):
        return None

    n = len(years)
    de_vals = [v or 0 for v in _fill_none(de, n)]

    fig, ax = plt.subplots(figsize=(6, 2.8))
    ax.bar(range(n), de_vals, color=_C["blue"], zorder=3)
    ax.set_xticks(range(n))
    ax.set_xticklabels(years, fontsize=7)
    _style(fig, ax, f"{company_name} — Debt/Equity Ratio")
    fig.tight_layout(pad=0.6)
    return _to_png(fig)


# ── Dispatcher ─────────────────────────────────────────────────────────────────

_CHART_MAP: dict[str, Any] = {
    "revenue_trend":       _revenue_trend,
    "margin_trend":        _margin_trend,
    "roe_roce":            _roe_roce,
    "pat_trend":           _pat_trend,
    "ebitda_trend":        _ebitda_trend,
    "peer_comparison":     _peer_comparison,
    "valuation_waterfall": _valuation_waterfall,
    "sensitivity_heatmap": _sensitivity_heatmap,
    "saarthi_radar":       _saarthi_radar,
    "debt_equity":         _debt_equity,
}


def generate_chart(
    chart_key: str,
    fin_model: dict,
    company_name: str = "",
) -> bytes | None:
    """Dispatcher: generate a chart PNG from financial model data.

    Args:
        chart_key:    Name after "chart:" in the shape name (e.g. "revenue_trend")
        fin_model:    The fin_model dict built by _build_financial_model()
        company_name: Used in chart titles
    Returns:
        PNG bytes, or None if data is absent or matplotlib is not installed.
    """
    fn = _CHART_MAP.get(chart_key)
    if fn is None:
        logger.warning("Unknown chart key '%s' — skipping", chart_key)
        return None
    try:
        return fn(fin_model, company_name)
    except Exception as exc:
        logger.warning("Chart generation failed for '%s': %s", chart_key, exc)
        return None
