# PPT Generator Workflow

End-to-end documentation for the equity-research PPTX pipeline that runs after
Stage 2 of `/admin/pipeline`. This supersedes the older
[`PPT_GENERATOR_HANDOFF.md`](PPT_GENERATOR_HANDOFF.md) — keep that for legacy
debugging context, but new work should follow this document.

---

## 1. High-level flow

```
┌─────────────────────────────────────────────────────────────────┐
│ Stage 2 approved (research_sections written, status=stage2_…)   │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│ runPptCopywriting()  — anthropic-pipeline.ts                    │
│   • single Sonnet call, no web search                            │
│   • input: 12 narrative + 6 atomic sections + metadata           │
│   • output: ~82 fields keyed by master_template placeholder name │
│   • sanitised + length-clipped via ppt-copy-schema.ts            │
│   • saved to research_sessions.ppt_content_json (jsonb)          │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│ POST /generate-pptx (scripts/ppt_service/main.py)                │
│   • fetch research_reports + research_sessions + sections        │
│   • download model JSON + Excel from Supabase Storage            │
│   • map_replacements() builds the deterministic chips            │
│   • merge ppt_content_json over the heuristics                   │
│   • fill_master_template() walks every slide:                    │
│       – replace {{tokens}} in text frames                        │
│       – inject matplotlib chart/table images                     │
│       – preserve aspect ratio for pies                           │
│   • slide-specific post-processors (story charts, timeline,      │
│     governance, key risks, formula tables, scenarios)            │
│   • upload .pptx (and .pdf if LibreOffice present) to Supabase   │
│   • return URL + warnings (slide-copy fields applied, injection  │
│     counts) to caller                                            │
└─────────────────────────────────────────────────────────────────┘
```

---

## 2. Tech stack

### Frontend / orchestration
- **TypeScript + React (Vite)** — UI lives in [`src/pages/ResearchPipeline.tsx`](../src/pages/ResearchPipeline.tsx) and [`src/components/pipeline/PostProductionPanel.tsx`](../src/components/pipeline/PostProductionPanel.tsx).
- **Anthropic SDK** (`@anthropic-ai/sdk`) — direct browser calls with `dangerouslyAllowBrowser: true`. Model: `claude-sonnet-4-6`.
- **Supabase JS** — DB + storage CRUD via [`src/lib/supabase.ts`](../src/lib/supabase.ts) and [`src/lib/pipeline-api.ts`](../src/lib/pipeline-api.ts).
- **Shared schema module** — [`src/lib/ppt-copy-schema.ts`](../src/lib/ppt-copy-schema.ts) is pure (no Vite deps) so the CLI imports it too.

### Python rendering service
- **FastAPI** — [`scripts/ppt_service/main.py`](../scripts/ppt_service/main.py) on `localhost:8501`.
- **python-pptx** — open / mutate / save `master_template.pptx`.
- **openpyxl** — read the financial-model `.xlsx` with formula evaluation cache (`_evaluate_excel_formula_cell`).
- **matplotlib** — render every chart and table as PNG, then injected into placeholder shapes.
- **Pillow (PIL)** — measure image dimensions for aspect-preserving pie inserts.
- **win32com (optional)** — alternative Excel rendering path; only used when running on Windows with Office installed.
- **LibreOffice (optional)** — produces companion PDF from the generated `.pptx`. If not installed, only the `.pptx` is uploaded.

### CLI tool
- **tsx** (auto-installed by `npx`) — runs `scripts/generate_ppt_copy.ts` with the shared schema, mimicking the browser flow without the UI. Loads both the root `.env` (for Vite vars) and `scripts/ppt_service/.env` (for `SUPABASE_SERVICE_KEY` which bypasses RLS on `research_sections`).

### Storage and tables
| Resource | Where | Purpose |
|---|---|---|
| `research_sessions.ppt_content_json` (jsonb) | Supabase Postgres | Output of the copywriting LLM pass. Migration: [`supabase/migrations/20260513_ppt_content_json.sql`](../supabase/migrations/20260513_ppt_content_json.sql) |
| `research_sections` (stage='stage2') | Supabase Postgres | 18 sections — 12 narrative + 6 atomic. Source for copywriting + heuristic fallback. |
| `research_reports.cs_*` | Supabase Postgres | Per-report chips (CMP, target, rating, market cap, etc.). |
| `research_reports.pptx_file_url` / `pptx_pdf_file_url` | Supabase Postgres | URLs back to the storage bucket. |
| `research-reports-pptx` bucket | Supabase Storage | Final `.pptx` and `.pdf` files. Public read. |
| `research-reports-html` bucket | Supabase Storage | `financial-models/<TICKER>/<TICKER>_model.xlsx` + `_model.json`. |

---

## 3. Master template

`master_template.pptx` lives at the repo root. It is the source of truth — never embed long-form prose in code; instead add a `{{placeholder}}` token to the template and a corresponding schema entry.

### Slide-by-slide placeholder map

#### Slide 1 — Cover + Summary Table
- **Atomic chips (deterministic):** `{{company_name}}`, `{{nse_code}}`, `{{cmp}}`, `{{target}}`, `{{m_cap}}`, `{{m_category}}`, `{{saarthi_s}}`
- **LLM copy:** `{{tagline}}`, `{{investment_thesis_s1}}`, `{{investment_ideas_1..4}}`
- **Image injection:** `{{financial_summary_image}}` → `_render_financial_summary_dashboard()` reads Excel `P&L / Balance Sheet / Cash Flow / Ratios` sheets and produces the right-side summary table.

#### Slide 2 — Story in Charts (Financial)
- **Image injection:** `{{financial_charts}}` → `_render_story_chart_collage(operational=False)` produces a 2×2 matplotlib collage with Revenue / EBITDA / PAT / P/E.
- **Y-axis:** every chart starts at 0; top is auto + ~12% headroom via `_pad_top()` so peak data doesn't kiss the top edge.
- **Labels:** xlabel "Fiscal Year", ylabel "₹ Cr" or "P/E (x)".
- **P/E source:** `Valuations_Table` sheet row "P/E (x)" — read directly to avoid CMP/EPS volatility from anomalous reported EPS.

#### Slide 3 — Story in Charts (Operational)
- **Image injection:** `{{operational_charts}}` → same renderer with `operational=True`. Series: Capacity Utilisation %, Plant Network (stacked India/Overseas), Countries of Operation, Volume by Segment (MT).
- **Y-axis pinned at 0, top padded.**

#### Slide 4 — Investment Thesis
- **LLM copy:** `{{investment_thesis_heading_s4}}`, `{{investment_thesis_s4}}` (2-3 paragraphs, `\n\n` between), `{{key_catalyst_heading_1..3}}`, `{{key_catalyst_1..3}}`, `{{saarthi_summary_s4}}`
- **Special handling:** `_replace_slide4_thesis_shapes()` handles the repeated `{{investment_thesis}}` placeholders in the template by position (left big box gets long-form, three right small boxes get short summaries, bottom strip gets medium summary).

#### Slide 5 — Industry Overview + KPI strip
- **LLM panels:** `{{industry_structure}}`, `{{key_industry_tailwinds}}`, `{{key_industry_risks}}` (~1400 chars each)
- **LLM-picked KPIs:** `{{KPI_heading_1..6}}` + `{{KPI_1..6}}`. The prompt forbids reusing the slide-1 chips (CMP / Target / Market Cap / Cap Category / SAARTHI / NSE) — pick six different operating/financial metrics with units (e.g. "Revenue FY26" → "₹4,265 Cr", "EBITDA Margin" → "11.8%").

#### Slide 6 — Company Overview
- **LLM copy:** `{{company_overview}}` (top panel), `{{competitive_moat_1..2}}`, `{{key_insights}}`
- **Image injection:** `{{percentage_revenue_pie_chart}}` + `{{percentage_EBIT_pie_chart}}` → `_render_pie_chart()`. Square (3.5×3.5) figures. Inserted with `preserve_aspect=True` so they stay circular inside the template's non-square placeholders.

#### Slide 7 — Company Timeline
- **Image injection:** `{{company_timeline}}` → `_render_timeline_table()` reads the `Timeline` sheet (Year / Event Category / Description / Strategic Impact) and renders a branded matplotlib table with category-colored cells.
- **Idempotent:** the post-cleanup re-injection pass is a no-op if the placeholder is gone (was previously stamping a duplicate small image).

#### Slide 8 — Business Model
- **LLM copy:** `{{business_model_1..6}}` — six cards, each must cover a distinct facet (revenue model / geography / value chain / cost / capital / diversification).

#### Slide 9 — Competitive Advantage
- **LLM copy:** `{{competitive_advantage}}` (left panel, 2-3 paragraphs)
- **Image injection:** `{{peer_comparison_chart_1}}` (Revenue FY26A peer bar chart) + `{{peer_comparison_chart_2}}` (EBITDA Margin FY26A peer bar chart) → `_render_peer_bar_chart()` reads the `Peer_Compare` sheet.

#### Slide 10 — Growth Catalysts & Timeline
- **LLM copy:** `{{investment_thesis_detailed}}` (left panel, 2-3 paragraphs, ~3000 chars — the longest), `{{key_catalyst}}` (right top, ~1100 chars — must be visually different from `investment_thesis_detailed`).
- **Image injection:** `{{catalyst_timeline_chart}}` → reuses `_render_timeline_table()`.

#### Slide 11 — Management Commentary
- **LLM copy:** 8 heading/content pairs `{{management_commentry_heading_1..8}}` / `{{management_content_1..8}}` covering Capital Allocation / Execution / Expansion / Leadership / Funding / Margin / Vision / Risk Controls. The prompt requires distinct facets per card.

#### Slide 12 — Management Analysis and Corporate Governance
- **Image injection:** `{{governance_table}}` → `_render_governance_table()` produces two stacked sub-tables (Board of Directors + Shareholding Pattern) from the `Governance` sheet. `chars_per_inch=10` so long names wrap inside the Name column.
- **LLM copy:** `{{indicators_1..6}}` — six governance signal cards.

#### Slide 13 — Earnings Forecast
- **Image injection:** `{{earnings_forecast_table}}` → `_render_formula_sheet_table(sheet_name="Earnings_Forecast")`.
- **LLM copy:** `{{forecast_assumptions}}` (right side panel).

#### Slide 14 — Financials
- **Image injection:** `{{financials_table}}` → `_render_formula_sheet_table(sheet_name="Financials_Table")`. Uses live formula evaluation (`use_formula_eval=True`) for that sheet.
- **LLM copy:** `{{financial_commentary}}` (right side panel).

#### Slide 15 — Valuations
- **Image injection:** `{{valuations_table}}` → `_render_formula_sheet_table(sheet_name="Valuations_Table")`.
- **LLM copy:** `{{valuation_commentary}}` (right side panel).

#### Slide 16 — SAARTHI Framework
- **LLM copy:** seven dimension cards `{{saarthi_s_content}}`, `{{saarthi_a1_content}}`, `{{saarthi_a2_content}}`, `{{saarthi_r_content}}`, `{{saarthi_t_content}}`, `{{saarthi_h_content}}`, `{{saarthi_i_content}}` (each leads with "12/15 - Dimension Name:" score header + evidence) and `{{saarthi_summary_s16}}` (bottom verdict).
- A heuristic splitter `_split_saarthi_framework()` is the fallback when `ppt_content_json` is absent — it parses the SAARTHI section text by letter prefixes (`S — `, `A — `, etc.).

#### Slide 17 — Scenario Analysis
- **Atomic chips:** `{{bull}}` / `{{base}}` / `{{bear}}` (prices), `{{bull_p}}` / `{{base_p}}` / `{{bear_p}}` (probabilities), `{{valuation_bull}}` / `{{valuation_base}}` / `{{valuation_bear}}` (mirror of prices).
- **LLM copy:** `{{bull_content}}` / `{{base_content}}` / `{{bear_content}}` (~600 chars each).
- **Image injection:** `{{probability_weight_table}}` → `_render_probability_weight_table()` produces the bottom Scenario / Target / Probability / Weighted TP table.

#### Slide 18 — Key Risks
- **Image injection:** `{{key_risks_table}}` → `_render_key_risks_table()` reads the `Key_Risks` sheet, renders the 8-row table with color-coded Probability / Impact / Overall Rating chips (`H`=red, `M`=amber, `L`=green).
- **Column widths:** `[0.04, 0.12, 0.13, 0.25, 0.225, 0.07, 0.07, 0.095]` — sums to 1.0. Description + Mitigation get the lion's share; Overall Rating is narrow.

#### Slide 19 — Entry / Review / Exit Strategy
- **Atomic chips:** `{{buy}}`, `{{tar_pr}}`, `{{stp_loss}}`, `{{up}}`, `{{down}}`, `{{pnt}}` (Risk/Reward ratio).
- **LLM copy:** `{{entry_strategy}}`, `{{review_strategy}}`, `{{exit_strategy}}` (~1400 chars each).

#### Slide 20 — Disclosure
- **Literal substitution:** `_apply_literal_text_subs()` rewrites the baked-in `(Premier Energies Ltd.)` in clause 5 of the disclosure table to `(<company_name>)`. The template's other clauses are static.

---

## 4. Override priority

When the same placeholder has multiple potential sources, the Python service applies them in this order (lowest → highest priority):

1. **Deterministic chips from `map_replacements()`** — derived from `metadata`, `fin_model`, and section text via heuristics (`_split_saarthi_framework`, `_metric_chips`, `_truncate_words`, etc.).
2. **`research_sessions.ppt_content_json`** — the LLM copywriting output. Overrides heuristics for every key it contains.
3. **`research_reports.cs_ppt_data`** — manual UI overrides from the `PPTDataPanel` (highest priority).

The Python service logs the count of fields applied at level 2 as `Slide copy fields applied: N` in the response `warnings` array. If `N == 0`, the LLM pass either hasn't run or failed — the deck falls back to heuristic copy.

---

## 5. PPT copywriting LLM pass

### Schema — [`src/lib/ppt-copy-schema.ts`](../src/lib/ppt-copy-schema.ts)

Single source of truth for both the browser path and the CLI. Every entry has:
- `kind`: `'line'` | `'card'` | `'panel'`
- `max`: hard char cap (~10-20% above the target so the LLM doesn't have to chop mid-sentence)
- `instruction`: per-field rules — slot purpose, target length, distinct-facet rules for sibling cards, paragraph-break requirement for multi-paragraph panels

Char-budget summary (target ~80-90% of max):

| Group | Fields | Target |
|---|---|---|
| Tagline | `tagline` | 10-16 words |
| Cards (slide-1 ideas) | `investment_ideas_1..4` | ~350 chars |
| Catalyst cards | `key_catalyst_1..3` | ~350 chars |
| Card content | `investment_thesis_s1` | ~1300 chars |
| Industry panels | `industry_structure`, `key_industry_tailwinds`, `key_industry_risks` | ~1400 chars |
| KPI strip | `KPI_heading_1..6`, `KPI_1..6` | 1-3 word label + number-with-units |
| Company overview | `company_overview` | ~1600 chars |
| Moats | `competitive_moat_1..2` | ~550 chars |
| Operating snapshot | `key_insights` | ~1500 chars |
| Business model | `business_model_1..6` | ~900 chars each |
| Competitive advantage | `competitive_advantage` | ~2700 chars, 2-3 paragraphs |
| Catalyst panels | `investment_thesis_detailed` / `key_catalyst` | 3000 / 1100 chars |
| Management | `management_commentry_heading_1..8` / `management_content_1..8` | 2-3 words / ~400 chars |
| Governance indicators | `indicators_1..6` | ~350 chars |
| Forecast / financial / valuation commentary | `forecast_assumptions`, `financial_commentary`, `valuation_commentary` | ~2000-2400 chars |
| SAARTHI cards | `saarthi_s/a1/a2/r/t/h/i_content` | ~450 chars |
| SAARTHI summary | `saarthi_summary_s16` | ~700 chars |
| Scenario cards | `bull/base/bear_content` | ~600 chars each |
| Strategy trio | `entry_strategy`, `review_strategy`, `exit_strategy` | ~1400 chars each |

### Prompt rules (`buildPptCopyPrompt`)
- One JSON object output, no markdown fences.
- Plain text only — no markdown bold/italics, no pipe tables, no bullet markers, no `#` headers.
- Cards must cover **distinct facets** when they sit on the same slide (eight different management facets, six different business-model angles, etc.).
- The three thesis panels (`investment_thesis_s1`, `investment_thesis_s4`, `investment_thesis_detailed`) must have **different shapes** — same facts allowed, but different opening / emphasis / length (s1 < s4 < detailed).
- KPIs must **not** reuse the slide-1 chips.
- Multi-paragraph fields: use literal `\n\n` between paragraphs.
- Numbers come straight from source material — no fabrication.
- Aim for 80-100% of each budget; draw on related sections if the primary source is thin; only fall below 70% if there's genuinely nothing more to say.

### Sanitiser (`sanitisePptContent`)
- Drops unknown keys, coerces to string.
- Normalises whitespace while **preserving `\n\n` paragraph breaks**.
- Strips markdown emphasis (`**`, `__`).
- Clips to `max` at paragraph break → word boundary → hard cut.

### Browser entry point — `runPptCopywriting()` in [`src/lib/anthropic-pipeline.ts`](../src/lib/anthropic-pipeline.ts)
- Single Sonnet call, no web search, `maxTokens: 16000`, `temperature: 0.25`.
- Streamed (`client.messages.stream`) — long calls otherwise hit Anthropic's 10-minute SSE timeout.
- Used by `PostProductionPanel.handleGenerateSlideCopy()` and auto-runs on the first PPTX generation if `ppt_content_json` is empty.

### CLI entry point — [`scripts/generate_ppt_copy.ts`](../scripts/generate_ppt_copy.ts)
Mimics the browser flow from the terminal. Reads both `.env` files, uses the **service-role** Supabase key (so it can read RLS-protected `research_sections`), then fetches → calls Anthropic → sanitises → saves.

```cmd
:: dry run — print JSON without saving
npx tsx scripts/generate_ppt_copy.ts <sessionId> --dry

:: real run — saves to research_sessions.ppt_content_json
npx tsx scripts/generate_ppt_copy.ts <sessionId>
```

Always prints a **fill report** — per-field bar chart with char count vs budget. Anything below 50% is flagged `short`. This is the fastest way to spot under-filled slots.

---

## 6. Python service internals

Entry: [`scripts/ppt_service/pptx_generator.py`](../scripts/ppt_service/pptx_generator.py). Started by [`scripts/ppt_service/main.py`](../scripts/ppt_service/main.py).

### Endpoints
- `GET /health` → `{status, reportgen, libreoffice, supabase, openrouter_key}`
- `POST /preview-placeholders` → list of every `{{token}}` found across the template
- `POST /generate-pptx` → full pipeline, returns `{status, pptx_file_url, pptx_pdf_file_url, duration_seconds, warnings[]}`

### Key functions
| Function | Purpose |
|---|---|
| `generate_pptx_for_report(report_id, session_id, use_mock)` | Top-level orchestrator. |
| `_fetch_inputs()` | Reads `research_reports`, `research_sessions`, `research_sections` rows. |
| `_download_model_json` / `_download_model_excel` | Pulls the financial model from Supabase Storage. |
| `_build_financial_model()` + `_enrich_financial_model_for_house_deck()` | Normalises the JSON model into the shape the template expects (operational, segments, scenarios, etc.). |
| `map_replacements()` | Builds the deterministic `{{token}}` → text dict from sections + metadata + fin_model. |
| `fill_master_template()` | Walks every slide, replaces text tokens, queues image inserts, then performs them. Also runs `_apply_literal_text_subs()` for non-token literals on slide 20. |
| `_replace_text_in_frame` / `_replace_text_in_table` | Run-aware token replacement that preserves font properties of the first run. |
| `_insert_image_into_shape(slide, shape, img, *, preserve_aspect=False)` | Replaces a placeholder shape with an image at the same bounds. `preserve_aspect=True` is used for pies to keep them circular. |
| `_cleanup_excel_placeholders()` | Final pass: any surviving `{{...}}` tokens get a polite fallback string ("Segment breakdown — see Excel model for details."). |
| `inject_story_chart_slides()` / `inject_company_timeline_slide()` / `inject_company_overview_slide()` / `inject_catalyst_timeline_slide()` / `inject_competitive_advantage_slide()` / `inject_peer_comparison_slide()` / `inject_governance_slide()` / `inject_formula_table_slide()` / `inject_key_risks_slide()` / `inject_probability_weight_slide()` | Slide-specific image injectors. All idempotent — second-pass calls are no-ops when placeholder is gone. |

### Standard matplotlib table conventions
All table renderers use the shared helpers:

- **`_wrap_table_cells(cell_rows, col_widths, *, fig_width_inches, chars_per_inch=11.0, skip_header=True)`** — pre-inserts `\n` line breaks so text wraps inside cells (matplotlib `Table` does not wrap natively).
- **`_apply_row_line_heights(table, line_counts)`** — redistributes cell heights proportionally to wrapped-line count so wrapped rows get more vertical room. Differential padding: `PAD=0.03` for multi-line rows (use full vertical area), `PAD=0.06` for single-line rows (give numbers breathing room).

#### chars_per_inch calibration
Empirical at standard fontsizes (proportional fonts):

| Fontsize | chars_per_inch |
|---|---|
| 7.0pt | 13-14 |
| 7.4pt | 12 |
| 7.6pt | 10-11 (formula tables) |
| 8.1pt | 10 (governance) |
| 8.3pt | 10-11 (timeline) |

#### fig_h growth caps
Slide placeholders are fixed-size. `python-pptx.add_picture(left, top, width, height)` **stretches** the image to fill — a too-tall figure gets vertically squashed when placed in a wide-short placeholder. Every table renderer caps `fig_h` so the rendered figure stays close to the placeholder's aspect ratio:

| Table | Placeholder ~aspect | fig_h cap |
|---|---|---|
| Timeline (slide 7) | wide | 7.5" |
| Peer table | tall | 9.0" |
| Key Risks (slide 18) | wide | 7.0" |
| Formula table (13/14/15) | wide | 7.5" |
| Summary dashboard (slide 1) | tall | 10.5" |

---

## 7. Local development

### One-time setup
```cmd
:: Python service deps
.\.venv\Scripts\pip.exe install -r requirements.txt

:: Frontend deps
npm install
```

### Required env
Root `.env`:
- `VITE_SUPABASE_URL`, `VITE_SUPABASE_ANON_KEY` — browser Supabase client.
- `VITE_ANTHROPIC_API_KEY` — Sonnet calls.
- `VITE_PPT_SERVICE_URL` — optional override; defaults to `/proxy/ppt`.

`scripts/ppt_service/.env`:
- `SUPABASE_URL`, `SUPABASE_SERVICE_KEY` — service role key for the Python service and the CLI (bypasses RLS on `research_sections`).
- `OPENROUTER_API_KEY` — only for the legacy mock-planner path. Not used by the LLM pass.

### Run the service
```cmd
.\.venv\Scripts\python.exe scripts\ppt_service\main.py
:: → http://localhost:8501
curl http://localhost:8501/health
```

### Apply migrations
```sql
-- supabase/migrations/20260513_ppt_content_json.sql
alter table public.research_sessions
add column if not exists ppt_content_json jsonb;
```

### Test end-to-end from the terminal
```cmd
:: 1. terminal 1 — start service (or confirm it's running)
.\.venv\Scripts\python.exe scripts\ppt_service\main.py

:: 2. terminal 2 — generate slide copy (dry run first to inspect)
npx tsx scripts/generate_ppt_copy.ts <sessionId> --dry
npx tsx scripts/generate_ppt_copy.ts <sessionId>

:: 3. terminal 3 — render
powershell -Command "Invoke-RestMethod -Method Post -Uri http://localhost:8501/generate-pptx -ContentType 'application/json' -Body '{\"reportId\":\"<reportId>\",\"sessionId\":\"<sessionId>\",\"useMock\":false}' | ConvertTo-Json -Depth 6"
```

Look for `Slide copy fields applied: ~82` in the response `warnings` array — confirms the LLM JSON flowed through.

---

## 8. Common gotchas

- **The Python service does not hot-reload.** After editing `pptx_generator.py`, kill the process and restart — old code keeps running otherwise. Symptom: warnings array missing fields you just added.
- **RLS blocks anon reads on `research_sections`.** The CLI uses `SUPABASE_SERVICE_KEY` from `scripts/ppt_service/.env`. Without it, you'll see `no stage2 sections found`.
- **`python-pptx.add_picture` stretches.** Passing both `width` and `height` ignores aspect ratio. Use `preserve_aspect=True` for anything where the rendered image is a different aspect than the placeholder.
- **`_cleanup_excel_placeholders` runs LAST.** If a slide-specific injector silently fails, the cleanup pass writes "<topic> — see Excel model for details." into the placeholder. That string in a deck always means an injector returned 0.
- **Multi-paragraph fields need `\n\n`** in the JSON. The sanitiser preserves them; the prompt explicitly requires them for `investment_thesis_s4`, `competitive_advantage`, `investment_thesis_detailed`.
- **KPI fields are LLM-picked, not metadata-derived.** The deterministic `map_replacements()` writes slide-1 chips into `KPI_*` as a fallback; the LLM pass overrides with six company-specific operating/financial metrics.
- **Date / FY labels matter.** The template assumes Indian fiscal year (`FY26A`, `FY28E`). LLM is instructed to keep those formats verbatim.

---

## 9. Recent fixes (changelog)

| Fix | Files | Why |
|---|---|---|
| SAARTHI splitter for slide 16 | `pptx_generator.py::_split_saarthi_framework` | Previously every dimension card got identical text. |
| Slide 19 `%%` formatting + `(Premier Energies)` disclosure | `pptx_generator.py::fill_master_template`, `_apply_literal_text_subs` | Double-percent rendering and stale literal company name. |
| Slide 9 mojibake (`â‚¹` → `₹`) | `pptx_generator.py::_render_peer_bar_chart` | Encoding error in title string. |
| Slide 2 P/E chart anomaly | `pptx_generator.py::_extract_financial_chart_history_from_excel` | Use `Valuations_Table` P/E directly instead of CMP/EPS. |
| Slide 10 duplicate panel | `pptx_generator.py::map_replacements` | `key_catalyst` now sources from catalyst section, not thesis box texts. |
| Slide 11 truncation | `pptx_generator.py::map_replacements` | Bumped per-card word budget 26 → 35. |
| **LLM copywriting pass** (large feature) | `ppt-copy-schema.ts`, `anthropic-pipeline.ts`, `pipeline-api.ts`, `PostProductionPanel.tsx`, `pptx_generator.py`, `scripts/generate_ppt_copy.ts`, migration | End-to-end copywriting LLM pass + CLI + Python merge. |
| Slide 7 tiny duplicate image | `pptx_generator.py::inject_company_timeline_slide` | Removed fallback path that fired on post-cleanup re-injection. |
| Y-axis starts at 0 + headroom | `pptx_generator.py::_render_story_chart_collage` | Slides 2/3 charts. |
| Axis labels on every story chart | same | "Fiscal Year" / "₹ Cr" / "Utilisation (%)" etc. |
| Pie chart oval → circle | `pptx_generator.py::_insert_image_into_shape` (added `preserve_aspect`) | Square pie image was stretched to fill non-square placeholder. |
| Table text wrapping helper | `_wrap_table_cells` + `_apply_row_line_heights` | Wraps cell text and redistributes row heights for matplotlib tables. |
| Differential cell padding | `_apply_row_line_heights` | `PAD=0.03` for wrapped rows, `0.06` for single-line numeric rows. |
| Slide 18 column rebalance | `_render_key_risks_table` | Wider Description/Mitigation, narrower Overall Rating; sums to 1.0. |
| Governance / formula chars_per_inch | `_render_governance_table`, `_render_formula_sheet_table` | Calibrated against actual fontsize to prevent cell clipping. |

---

## 10. Quick reference

### File map
- **Browser code:** [`src/lib/ppt-copy-schema.ts`](../src/lib/ppt-copy-schema.ts), [`src/lib/anthropic-pipeline.ts`](../src/lib/anthropic-pipeline.ts), [`src/lib/pipeline-api.ts`](../src/lib/pipeline-api.ts), [`src/components/pipeline/PostProductionPanel.tsx`](../src/components/pipeline/PostProductionPanel.tsx)
- **Python service:** [`scripts/ppt_service/main.py`](../scripts/ppt_service/main.py), [`scripts/ppt_service/pptx_generator.py`](../scripts/ppt_service/pptx_generator.py), [`scripts/ppt_service/excel_injector.py`](../scripts/ppt_service/excel_injector.py)
- **CLI:** [`scripts/generate_ppt_copy.ts`](../scripts/generate_ppt_copy.ts)
- **Migration:** [`supabase/migrations/20260513_ppt_content_json.sql`](../supabase/migrations/20260513_ppt_content_json.sql)
- **Template:** [`master_template.pptx`](../master_template.pptx)
- **Sample model:** [`output/GRAVITA_model.xlsx`](../output/GRAVITA_model.xlsx), [`output/GRAVITA_model.json`](../output/GRAVITA_model.json)

### Adding a new placeholder

1. Add the `{{new_token}}` to `master_template.pptx` in the appropriate slide / shape.
2. Decide: deterministic chip, LLM-generated, or image injection?
   - **Chip** → add to `map_replacements()` in `pptx_generator.py`.
   - **LLM** → add an entry to `PPT_COPY_SCHEMA` in `ppt-copy-schema.ts` with `kind`, `max`, `instruction`. That's it — the prompt and sanitiser pick it up automatically.
   - **Image** → write a new `_render_*` helper and an `inject_*_slide()` orchestrator that finds the placeholder shape and calls `_insert_image_into_shape()`.
3. If it's a multi-paragraph LLM panel, mention "2-3 paragraphs separated by a blank line (\\n\\n between paragraphs)" in the instruction.
4. Test via `npx tsx scripts/generate_ppt_copy.ts <sessionId> --dry` — the fill report will include the new field.
5. Run `/generate-pptx`, confirm the warnings array shows the expected injection counts.
