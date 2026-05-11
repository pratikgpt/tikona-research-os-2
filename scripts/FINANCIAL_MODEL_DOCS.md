# Financial Model Pipeline — Complete Reference

**Version:** 5.1.0 | **Last updated:** 2026-05-11  
**Author context:** Tikona Capital equity research platform  
**Primary files:** `scripts/financial_model_v5.py` · `scripts/financial_model_server.py`

---

## Table of Contents

1. [What it does](#1-what-it-does)
2. [Architecture overview](#2-architecture-overview)
3. [Pipeline: step-by-step](#3-pipeline-step-by-step)
4. [Pydantic schema — the output contract](#4-pydantic-schema--the-output-contract)
5. [Excel workbook structure](#5-excel-workbook-structure)
6. [JSON sidecar](#6-json-sidecar)
7. [FastAPI server](#7-fastapi-server)
8. [Deployment & environment](#8-deployment--environment)
9. [Frontend integration](#9-frontend-integration)
10. [Row-map system (Screener row resolution)](#10-row-map-system-screener-row-resolution)
11. [Known edge cases & gotchas](#11-known-edge-cases--gotchas)
12. [Cost & performance](#12-cost--performance)
13. [How to update / extend](#13-how-to-update--extend)
14. [Troubleshooting guide](#14-troubleshooting-guide)

---

## 1. What it does

Given a company's NSE ticker (`GRAVITA`, `TCS`, etc.), the pipeline:

1. Downloads 10 years of audited financial data from **Screener.in** as an Excel file.
2. Fetches live market data (CMP, PE, MCap) from **yfinance** and Screener's public web page.
3. Sends all of the above to **Claude Sonnet 4 + web_search** in a single streaming call that instructs the model to search for latest quarterly results, analyst consensus, broker estimates, and sector multiples.
4. Receives a fully structured JSON with 5-year projections, 3-method valuation, SAARTHI scores, investment thesis, peers, and assumptions.
5. **Validates** the JSON against a Pydantic schema — hard-fails if critical fields are missing or malformed.
6. **Normalises** gaps (forward-fills missing projection year assumptions, resolves `shares_cr`).
7. **Writes a JSON sidecar** (`{TICKER}_model.json`) — the machine-readable version used by Stage 1/2 prompts.
8. Appends **8 formatted Excel sheets** on top of the original Screener download using openpyxl formula strings so the Excel remains recalculation-aware.
9. Returns `{file_path, json_path, rating, target_price, upside_pct, cost_summary, elapsed_seconds}`.

---

## 2. Architecture overview

```
ResearchPipeline.tsx (frontend)
    │
    ▼
n8n webhook  POST /webhook/generate-financial-model
    │         (fires Python on VPS, returns when done)
    ▼
financial_model_server.py  (FastAPI, port 8500, /opt/financial-model/)
    │
    ▼
financial_model_v5.generate_financial_model()
    ├── download_screener_excel()  → {TICKER}_screener.xlsx
    ├── extract_screener_data()    → dict (P&L, BS, CF, row_map)
    ├── fetch_market_data()        → string (yfinance + Screener web)
    ├── call_claude_api()          → raw JSON dict
    ├── normalize_model_output()   → gap-filled dict
    ├── validate_model_output()    → FinancialModelOutput (Pydantic)
    ├── json.dump → {TICKER}_model.json   ← JSON sidecar
    └── build_model() → {TICKER}_financial_model.xlsx
    │
    ▼
financial_model_server.py  upload_model_to_supabase()
    ├── PUT xlsx → Supabase Storage  financial-models/{TICKER}/{TICKER}_model.xlsx
    └── PUT json → Supabase Storage  financial-models/{TICKER}/{TICKER}_model.json
    │
    ▼
ResearchPipeline.tsx  mirrorFinancialModelToStorage()
    → persists financial_model_file_url + financial_model_json_url on research_sessions
    │
    ▼
anthropic-pipeline.ts  getFinancialModelPromptContext()
    → fetches JSON sidecar → injects "## Financial Model Snapshot" into Stage 1 + Stage 2 prompts
```

---

## 3. Pipeline: step-by-step

### Step 1 — Screener.in login & Excel download (`download_screener_excel`)

- Uses **cloudscraper** (Chrome-UA) to bypass Screener's bot detection.
- Logs in at `https://www.screener.in/login/` with CSRF token from the page.
- Probes `https://www.screener.in/dash/` after login — if it returns non-200 or redirects to the login page, raises a hard exception (not a silent fallback).
- Tries consolidated page first (`/company/{TICKER}/consolidated/`), falls back to standalone with a warning.
- Finds the "Export to Excel" button, POSTs to its `formaction` URL.
- Saves file as `{output_dir}/{TICKER}_screener.xlsx`.

**Failure modes:** Wrong credentials, Screener rate-limiting, ticker not on Screener.

### Step 2 — Extract screener data (`extract_screener_data`)

- Reads the `Data Sheet` tab of the downloaded Excel.
- Extracts company name, face value, CMP, MCap from fixed header rows (rows 1, 7, 8, 9 column B).
- Detects fiscal year columns dynamically from row 16 (date objects → `"FY25"` etc.).
- Calls `_resolve_row_map(ws)` to find the actual row number for every data series (see [§10](#10-row-map-system-screener-row-resolution)).
- Returns a dict with keys: `company_name`, `face_value`, `cmp`, `mcap`, `fiscal_years`, `data_cols`, `num_years`, `row_map`, `pl`, `bs`, `cf`, `derived`.
- Computes `derived.ebitda` = Revenue - (all operating expenses). Uses `sum(v for v in [...] if v is not None)` — not `filter(None, ...)` which would silently drop zeros.

### Step 3 — Market data (`fetch_market_data`)

Returns a string combining two sources:
1. **yfinance** — CMP, MCap, trailing PE, PB, ROE, beta, 52-week range.
2. **Screener public page** — scrapes all `<section class="card">` rows (up to 60 lines). No auth needed.

Failures are soft-warned — the model still runs with whatever is available.

### Step 4 — Claude API call (`call_claude_api`)

**Model:** `claude-sonnet-4-20250514`  
**Tools:** `[{"type": "web_search_20250305", "name": "web_search"}]`  
**Max tokens:** 16,000  
**Mode:** Streaming — `client.messages.stream(...).get_final_message()`

The prompt sends the complete historical financial data (no truncation) and instructs Claude to:

- **Step 0** — identify the last actual year and derive 5 projection years dynamically (never hardcodes `FY26E..FY30E`).
- **Step 1** — run 10 web searches: latest quarterly results, annual results, management guidance, analyst consensus, broker estimates, sector PE/EV-EBITDA comps, India 10Y bond yield, GoIndiaStocks consensus, Trendlyne forecasts, order book.
- **Step 2** — build 5-year assumptions: 15 per-year keys (revenue growth, 7 P&L expense %, 6 absolute/working-capital drivers).
- **Step 3** — balance sheet projections (absolute values, cash as plug).
- **Step 4** — historical ratios (last 6 years, 11 metrics).
- **Step 5** — cash flow projections (CFO, CFI, CFF).
- **Step 6** — valuation: DCF + PE + EV/EBITDA → blended (40/30/30), 5×5 PE sensitivity grid.
- **Step 7** — 4+ listed peers with current MCap, PE, EV/EBITDA, EBITDA margin, ROE.
- **Step 8** — SAARTHI scoring (7 dimensions, 100-point total), investment thesis, bull/bear cases, catalysts/risks, final rating.

**Why streaming?** Anthropic's server-side `web_search_20250305` tool handles search internally — the model fires searches and accumulates results in one go. Streaming with `get_final_message()` gives Claude the time it needs (web searches add ~30–90s) without hitting timeout limits. A manual tool-use loop is wrong for server-side tools — Claude handles everything internally.

**JSON repair:** If `extract_json_from_text` fails, a second non-streaming call asks Claude to fix the JSON. Both calls are tracked by `CostTracker`.

### Step 5 — Normalise (`normalize_model_output`)

- Derives `projection_years` from `assumptions.projection_years` → `projections.years` → `_build_projection_years(last_actual_year)` (fallback cascade).
- For each of the 15 `PER_YEAR_ASSUMPTION_KEYS`: forward-fills any gap years from the nearest known value, backfills if the first year is missing. Warns with `{key}, {year}, {src_year}, {src_val}, {company}` per gap.
- Resolves `shares_cr` with a three-level cascade:
  1. Claude's value (if non-null and non-zero)
  2. `shares_outstanding` from Screener BS data
  3. `share_capital[-1] / face_value`
- Sets `cmp` and `base_year` from Screener data if Claude omitted them.

### Step 6 — Validate (`validate_model_output`)

Hard-fails with a readable `RuntimeError` listing every schema violation. Checks:
- Pydantic `FinancialModelOutput.model_validate(model_json)` — all fields, types, enum values.
- `projection_years` must have exactly 5 items.
- `projections.years` must have exactly 5 items.
- Every `PER_YEAR_ASSUMPTION_KEYS` dict must contain all 5 projection years.
- Every `projections.*` list must have exactly 5 values.
- `historical_ratios.*` lists must match `historical_ratios.years` length.
- Warns (not fails) if `peers` has fewer than 4 entries.

### Step 7 — Write JSON sidecar

```python
json_path = os.path.join(output_dir, f"{nse_code}_model.json")
with open(json_path, "w") as f:
    json.dump(validated_model.model_dump(), f, indent=2, default=str)
```

This writes the Pydantic-validated, schema-conformant JSON. The Pydantic `model_dump()` ensures all fields are present even if optional. This is what the frontend pipeline reads.

### Step 8 — Build Excel (`build_model`)

See [§5](#5-excel-workbook-structure).

### Step 9 — Cash preflight (inside `build_model`)

Before writing Excel sheets, reconstructs the balance sheet numerically to check:
- `total_assets = net_block + cwip + investments + other_assets + receivables + inventory + cash`
- `total_liabilities = share_capital + reserves + borrowings + other_liabilities`
- Flags negative projected cash → warns + records in `diagnostics` list.

If any diagnostics, a **Model Diagnostics** sheet is appended to the workbook listing all warnings.

---

## 4. Pydantic schema — the output contract

```
FinancialModelOutput
├── sector: str
├── base_year: str                         # e.g. "FY25"
├── cmp: float                             # ₹/share, from Screener/yfinance
├── target_price: float                    # ₹/share
├── rating: Literal["STRONG BUY"|"BUY"|"ACCUMULATE"|"HOLD"|"UNDERPERFORM"|"SELL"]
├── upside_pct: float
├── shares_cr: float                       # shares in crores
│
├── assumptions: Assumptions
│   ├── projection_years: list[str]        # exactly 5: ["FY26E",...,"FY30E"]
│   ├── revenue_growth_pct: dict[str,float]  # keyed by year label
│   ├── rm_pct, employee_pct, power_fuel_pct, other_mfg_pct,
│   │   selling_admin_pct, other_exp_pct, chg_inventory_pct (all dict[str,float])
│   ├── depreciation_cr, interest_cr, other_income_cr (dict[str,float], absolute ₹Cr)
│   ├── tax_rate_pct, capex_cr, receivable_days, inventory_days (dict[str,float])
│   ├── dividend_payout_pct: float
│   ├── target_pe: float
│   ├── target_ev_ebitda: float
│   ├── wacc_pct: float
│   ├── terminal_growth_pct: float
│   └── rationale: str
│
├── projections: Projections
│   ├── years: list[str]                   # exactly 5
│   ├── net_block, cwip, investments, other_assets (list[float], 5 values, absolute ₹Cr)
│   ├── share_capital, reserves, borrowings, other_liabilities (list[float], 5 values)
│   └── cfo, cfi, cff (list[float], 5 values)
│
├── historical_ratios: HistoricalRatios
│   ├── years: list[str]                   # last 6 FY years from Screener
│   └── ebitda_margin_pct, pat_margin_pct, roe_pct, roce_pct, debt_equity,
│       receivable_days, inventory_days, asset_turnover,
│       rm_pct, employee_pct, tax_rate_pct (all list[float], same length as years)
│
├── valuation: Valuation
│   ├── dcf_fair_value: float
│   ├── pe_fair_value: float
│   ├── ev_ebitda_fair_value: float
│   ├── blended_fair_value: float          # DCF 40% + PE 30% + EV/EBITDA 30%
│   └── sensitivity_pe: SensitivityPe
│       ├── row_label: str                 # "PE Multiple"
│       ├── col_label: str                 # "EPS Growth %"
│       ├── row_values: list[float]        # 5 PE multiples
│       ├── col_values: list[float]        # 5 EPS growth rates
│       └── grid: list[list[float]]        # 5×5 target price matrix
│
├── peers: list[Peer]                      # min 4 recommended
│   └── Peer: {name, mcap_cr, pe, ev_ebitda, ebitda_margin_pct, roe_pct}
│
└── thesis: Thesis
    ├── investment_thesis: str
    ├── bull_case: str
    ├── bear_case: str
    ├── key_catalysts: list[str]           # min 4
    ├── key_risks: list[str]               # min 4
    ├── saarthi_scores: SaarthiScores
    │   ├── S_sector_quality: float        # 0–15
    │   ├── A_accounting_quality: float    # 0–15
    │   ├── A_asset_quality: float         # 0–15
    │   ├── R_revenue_visibility: float    # 0–15
    │   ├── T_track_record: float          # 0–10
    │   ├── H_balance_sheet_health: float  # 0–15
    │   └── I_intrinsic_valuation: float   # 0–15
    ├── saarthi_total: float               # sum; ≥80 STRONG BUY, 65-79 BUY, etc.
    └── saarthi_rating: str
```

**Rating thresholds:**
| SAARTHI Total | Rating |
|---|---|
| ≥ 80 | STRONG BUY |
| 65–79 | BUY |
| 55–64 | ACCUMULATE |
| 45–54 | HOLD |
| 35–44 | UNDERPERFORM |
| < 35 | SELL |

---

## 5. Excel workbook structure

The output file is the original Screener Excel with **8 sheets appended**. The Screener sheets (Data Sheet, Quarterly, etc.) are kept intact.

| Sheet | Contents |
|---|---|
| **Cover** | Company name, rating badge (colour-coded), target price, CMP, upside %, base year, investment thesis, bull/bear cases, SAARTHI scores table, key catalysts, key risks |
| **Assumptions** | All 15 per-year assumption rows for 6 historical + 5 projected years. Blue = historical (pulled from Screener), orange = projected (Claude's assumptions). Rationale paragraph below. |
| **P&L** | Revenue, 7 operating expense lines, EBITDA, Depreciation, EBIT, Interest, Other Income, PBT, Tax, PAT. Historical = Screener reference formula (`='Data Sheet'!{col}{row}`). Projected = Excel formula (e.g. `=C7*(1+Assumptions!C4/100)`). |
| **Balance Sheet** | Net Block, CWIP, Investments, Other Assets, Receivables, Inventory, Cash (plug), Total Assets; Share Capital, Reserves, Borrowings, Other Liabilities, Total Liabilities. Cash cell = `=Total Assets - non-cash assets`. |
| **Cash Flow** | CFO, CFI, CFF, Net Cash Flow. Historical from Screener references; projected from `projections` dict. |
| **Ratios** | Historical and projected ratios: EBITDA Margin, PAT Margin, ROE, ROCE, D/E, Receivable Days, Inventory Days, Asset Turnover, EPS, P/E, EV/EBITDA. |
| **Valuation** | DCF, PE, EV/EBITDA fair values + blended target. PE sensitivity 5×5 grid with conditional formatting (green = above target, red = below). |
| **KPI Dashboard** | Peer comparison table. Revenue, EBITDA, PAT bar chart (openpyxl BarChart). |

**Formula convention:** Historical years reference the original Screener `Data Sheet` cells directly (`='Data Sheet'!B17`). Projected years use formulas referencing the Assumptions sheet and prior year cells. This means the Excel remains dynamic — changing an assumption cell recalculates the entire model.

**openpyxl limitation:** openpyxl writes formula strings but cannot evaluate them. If you open the file programmatically (e.g., to read projected revenue), you get `None` for formula cells. Use the JSON sidecar instead (see [§6](#6-json-sidecar)).

---

## 6. JSON sidecar

**File:** `{output_dir}/{TICKER}_model.json`  
**Supabase path:** `research-reports-html/financial-models/{TICKER}/{TICKER}_model.json`  
**DB column:** `research_sessions.financial_model_json_url`

The JSON sidecar is the canonical machine-readable output. It contains the complete `FinancialModelOutput` schema serialised via `validated_model.model_dump()`. Every assumption, projection, valuation, thesis, peer, and SAARTHI score is accessible as plain Python/JSON values.

### Why it exists

The Excel file uses openpyxl formula strings. There is no way to evaluate them without Excel recalculating. The JSON contains the pre-computed numbers that went into those formulas. It is the only file you can read programmatically after generation.

### How the frontend uses it

`getFinancialModelPromptContext(sessionId)` in `src/lib/anthropic-pipeline.ts`:

1. Queries `research_sessions.financial_model_json_url` for the session.
2. Fetches the JSON from Supabase Storage (15s timeout).
3. Builds a `## Financial Model Snapshot` markdown block containing: base year, projection years, CMP, target price, rating, upside %, SAARTHI score/rating, FM thesis/bull/bear/catalysts/risks, revenue growth assumptions, working capital assumptions, all 4 valuation anchors + target PE/EV-EBITDA, top-6 peers.
4. Injects this block into both Stage 1 (Investment Thesis) and Stage 2 (Full Report) system prompts.

The Stage 1/2 prompt instructions tell Claude: "Use this financial-model snapshot as an analyst-produced structured input. Keep outputs consistent with it unless fresher vault docs or web search clearly contradicts it, and call out contradictions explicitly."

### Old vs. new JSON schema

Models generated before v5.1.0 (pre-May-2026) may have an older schema:
```json
{ "company": {...}, "data_quality": {...}, "historical": {...}, "assumptions": {...}, "projections": {...} }
```
This schema is missing `target_price`, `rating`, `valuation`, `thesis`, etc. `getFinancialModelPromptContext` will return an empty context block for these models (all fields will be `N/A`). Re-run `/generate` on those tickers to get v5.1.0 JSON.

---

## 7. FastAPI server

**File:** `scripts/financial_model_server.py`  
**Port:** 8500 (override with `PORT` env var)  
**Base URL (from frontend):** `/proxy/fm` (Vite proxy → `http://72.61.226.16:8500`)

### Endpoints

| Method | Path | Description |
|---|---|---|
| `GET` | `/health` | Returns `{status, model_version, anthropic_key_set, screener_creds_set, output_dir}` |
| `POST` | `/generate` | Synchronous: blocks until model generated, returns full result. n8n calls this with a 15-min timeout. |
| `POST` | `/generate-async` | Returns `{job_id}` immediately, generation runs in background. |
| `GET` | `/job/{job_id}` | Poll status of an async job. Returns `{status, storage_url, json_storage_url, ...}` |
| `GET` | `/download/{ticker}` | Download the generated `.xlsx` file. |
| `POST` | `/storage/{ticker}` | Re-upload an existing model (xlsx + JSON) to Supabase Storage. |

### Request body for `/generate` and `/generate-async`

```json
{
  "nse_symbol": "GRAVITA",
  "company_name": "Gravita India Ltd",
  "sector": "Metal & Metal Products",
  "folder_id": null
}
```

### Response for `/generate`

```json
{
  "status": "success",
  "file_name": "GRAVITA_model.xlsx",
  "file_path": "/opt/financial-model/output/GRAVITA_model.xlsx",
  "storage_path": "financial-models/GRAVITA/GRAVITA_model.xlsx",
  "storage_url": "https://....supabase.co/storage/v1/object/public/research-reports-html/financial-models/GRAVITA/GRAVITA_model.xlsx",
  "json_storage_path": "financial-models/GRAVITA/GRAVITA_model.json",
  "json_storage_url": "https://....supabase.co/storage/v1/object/public/research-reports-html/financial-models/GRAVITA/GRAVITA_model.json",
  "message": "rating=BUY target=1450.0 upside=23.5% cost=$0.045",
  "duration_seconds": 87
}
```

### File layout on VPS

```
/opt/financial-model/
├── financial_model_v5.py          # main script
├── financial_model_server.py      # FastAPI server
├── financial_model_v3.py          # legacy, not used
├── venv/                          # Python virtualenv
├── backups/                       # pre-deploy backups
│   ├── financial_model_v5.py.bak
│   └── financial_model_server.py.bak
└── output/                        # MODEL_OUTPUT_DIR
    ├── GRAVITA/
    │   ├── GRAVITA_screener.xlsx
    │   ├── GRAVITA_model.json     # per-ticker subdir (created by generate_financial_model)
    │   └── ...
    ├── GRAVITA_model.xlsx         # flat copy (created by server after generation)
    ├── GRAVITA_model.json         # flat copy (used by /storage mirror endpoint)
    └── ...
```

Note: `generate_financial_model()` writes files into `{OUTPUT_DIR}/{TICKER}/`. The server then copies both xlsx and JSON to the flat `OUTPUT_DIR/` root so `upload_model_to_supabase` can find them side-by-side.

---

## 8. Deployment & environment

### VPS details

- **Host:** `72.61.226.16`
- **Service:** `systemd` unit `/etc/systemd/system/financial-model.service`
- **Python:** `3.12.3` at `/opt/financial-model/venv/bin/python3`
- **Working directory:** `/opt/financial-model/`

### Environment variables

The service reads from two systemd config files:

**`/etc/systemd/system/financial-model.service`** (main):
```ini
Environment=ANTHROPIC_API_KEY=sk-ant-api03-...
Environment=MODEL_OUTPUT_DIR=/opt/financial-model/output
```

**`/etc/systemd/system/financial-model.service.d/*.conf`** (override):
```ini
Environment="SCREENER_USERNAME=sumitpoddar@tikonacapital.com"
Environment="SCREENER_PASSWORD=Tikona@2022"
Environment="SUPABASE_URL=https://bmpvcjbfeyvkkbvclwkb.supabase.co"
Environment="SUPABASE_SERVICE_KEY=eyJhbG..."
Environment="SUPABASE_FINANCIAL_MODEL_BUCKET=research-reports-html"
```

### Deploy a new version

```bash
# On local machine
scp scripts/financial_model_v5.py root@72.61.226.16:/opt/financial-model/
scp scripts/financial_model_server.py root@72.61.226.16:/opt/financial-model/

# On VPS
python3 -m py_compile /opt/financial-model/financial_model_v5.py     # syntax check
python3 -m py_compile /opt/financial-model/financial_model_server.py
systemctl restart financial-model
curl http://localhost:8500/health  # verify model_version
```

Or use paramiko from the project root (already done for v5.1.0 deploy).

### Python dependencies (in venv)

```
fastapi
uvicorn
anthropic
yfinance
beautifulsoup4
requests
openpyxl
pandas
pydantic
cloudscraper
```

If adding a new package: `ssh root@72.61.226.16 "/opt/financial-model/venv/bin/pip install <package>"`

### Supabase Storage bucket

Bucket: `research-reports-html` (public read, service-key write)  
Path pattern: `financial-models/{TICKER}/{TICKER}_model.xlsx` and `financial-models/{TICKER}/{TICKER}_model.json`

The upload uses `x-upsert: true` — uploading for the same ticker overwrites the previous file.

### Database migration

One migration has been applied:
```sql
-- supabase/migrations/20260511_financial_model_json_url.sql
alter table public.research_sessions
add column if not exists financial_model_json_url text;
```

Status: **applied** to the live project (`bmpvcjbfeyvkkbvclwkb.supabase.co`).

---

## 9. Frontend integration

### Trigger (ResearchPipeline.tsx)

`handleGenerateFinancialModel()` (line ~350):
1. POSTs to n8n `/webhook/generate-financial-model` with `{nse_symbol, company_name, sector, folder_id}`.
2. n8n calls the VPS `/generate` endpoint and waits for the response (15-min timeout).
3. On success, calls `mirrorFinancialModelToStorage(nse_symbol)` → `POST /proxy/fm/storage/{TICKER}`.
4. Updates `research_sessions` with:
   - `financial_model_file_url` = storage URL of the xlsx
   - `financial_model_json_url` = storage URL of the JSON sidecar

### Stage 1 and Stage 2 context injection (anthropic-pipeline.ts)

Both `runStage1` and `runStage2` call `getFinancialModelPromptContext(sessionId)` before building their context block. The function:

- Returns `{ contextText: '' }` if no session / no JSON URL / fetch fails (graceful degradation).
- Returns a `## Financial Model Snapshot` markdown section if JSON is available.

The snapshot is injected **between** `financialContext` (equity universe data) and `vaultBriefing` (vault PDF summary) in the context block. This ordering is intentional — the financial model is more specific and analyst-produced than the equity universe snapshot.

### Types

- `PipelineSession.financial_model_json_url: string | null` (`src/types/pipeline.ts`)
- `mirrorFinancialModelToStorage` returns `{ fileUrl, filePath, jsonFileUrl, jsonFilePath }` (`src/lib/api.ts`)
- `updatePipelineOutput` accepts `financial_model_json_url` as a Partial Pick field (`src/lib/pipeline-api.ts`)

---

## 10. Row-map system (Screener row resolution)

Screener's Excel export has historically had inconsistent row positioning. The v5.1.0 code uses a label-based resolution system instead of hardcoded row numbers.

### How it works

`_resolve_row_map(ws)` scans column A of the `Data Sheet` (rows 1–140). For each row it normalises the label (lowercase, `&`→`and`, non-alphanumeric→space). For each of the 27 data series, it checks if any `ROW_LABEL_CANDIDATES` substring matches the normalised label.

### Fallback

If a label is not found, it falls back to `DEFAULT_ROW_MAP` hardcoded row numbers and logs a warning:
```
⚠ Could not locate Screener row for 'sales' by label — falling back to row 17
```

### DEFAULT_ROW_MAP (current hardcoded defaults)

| Series | Default row |
|---|---|
| sales | 17 |
| raw_material | 18 |
| change_in_inventory | 19 |
| power_fuel | 20 |
| other_mfg | 21 |
| employee_cost | 22 |
| selling_admin | 23 |
| other_expenses | 24 |
| other_income | 25 |
| depreciation | 26 |
| interest | 27 |
| pbt | 28 |
| tax | 29 |
| net_profit | 30 |
| dividend_amount | 31 |
| share_capital | 57 |
| reserves | 58 |
| borrowings | 59 |
| other_liabilities | 60 |
| total_liab | 61 |
| net_block | 62 |
| cwip | 63 |
| investments | 64 |
| other_assets | 65 |
| total_assets | 66 |
| receivables | 67 |
| inventory | 68 |
| cash | 69 |
| shares_outstanding | 70 |
| cfo | 82 |
| cfi | 83 |
| cff | 84 |
| net_cash_flow | 85 |

If Screener changes its export format and labels don't match, update `ROW_LABEL_CANDIDATES` to add the new label variant. Only update `DEFAULT_ROW_MAP` if the fallback row number also changed.

---

## 11. Known edge cases & gotchas

### EBITDA zero-cost lines
`sum(v for v in [...] if v is not None)` — intentional. `filter(None, ...)` would drop zero values, underreporting EBITDA for companies where a cost line is genuinely 0.

### shares_cr resolution
`shares_cr` is in **crores of shares**. `share_capital` from Screener is ₹Cr of face-value equity (not a share count). The formula `share_capital / face_value` converts it. Do not use `share_capital` directly as shares_cr or EPS will be nonsense.

### Cash-as-plug balance sheet
The balance sheet balances because cash is computed as `total_assets - all_other_assets`. Projected cash can go negative if Claude projects heavy capex + debt repayment. Negative cash is flagged in the Diagnostics sheet — it means the model implies an implicit financing need that isn't in the borrowings line.

### Consolidated vs. standalone
Screener tries consolidated first. If unavailable, falls back to standalone with a warning. Standalone numbers are smaller for conglomerates — this affects all ratios and can make the model look weaker than reality. Watch for this on holding companies.

### Projection year labelling
The prompt derives projection years from the last actual year in Screener data. If the base year is ambiguous (e.g., row 16 has partial data for the current FY), Claude may pick a different base year than expected. Check `base_year` in the JSON sidecar.

### JSON repair call
If Claude's first response produces malformed JSON (unclosed braces, escaped characters), a second non-streaming repair call is made. This adds ~$0.02–0.05 to cost and ~10–30s to time. The repair call logs `⚠️ JSON parse failed, requesting repair...`.

### Sector not matching
The `sector` field in the JSON comes from Claude (which may normalise the input sector name). The frontend passes the sector from the `ResearchPipeline` form. They should match but may differ in capitalisation or wording — the JSON's sector field is purely informational.

### Old-format JSON sidecars
10 tickers already have JSON files in Supabase from before v5.1.0: ANGELONE, GRAVITA, HBLENGINE, ICICIBANK, MOTILALOFS, PNB, PREMIERENE, SUZLON (no xlsx), SWIGGY, TCS. These have the old schema and will return empty context from `getFinancialModelPromptContext`. Re-run `/generate` to refresh.

---

## 12. Cost & performance

### Typical run

| Phase | Time | Notes |
|---|---|---|
| Screener login + download | 5–15s | Cloudscraper + POST for Excel |
| yfinance + Screener web scrape | 5–10s | Parallel in `fetch_market_data` |
| Claude API (streaming, with web search) | 60–120s | 10 web searches + 16K token output |
| JSON normalise + validate | <1s | |
| Excel build (openpyxl) | 3–8s | |
| Supabase upload (xlsx + JSON) | 2–5s | |
| **Total** | **~90–150s** | |

### Token cost (Claude Sonnet 4)

| Item | Rate |
|---|---|
| Input | $3.00 / 1M tokens |
| Output | $15.00 / 1M tokens |
| Typical input | ~8,000–12,000 tokens |
| Typical output | ~6,000–10,000 tokens |
| **Typical total cost** | **~$0.10–0.25 per model** |

Cost is logged in every server response (`message` field) and in the VPS journal.

---

## 13. How to update / extend

### Add a new assumption key

1. Add the key to `PER_YEAR_ASSUMPTION_KEYS` list.
2. Add the field to `Assumptions` Pydantic class.
3. Add it to the prompt's STEP 2 instructions with the Excel formula it maps to.
4. Add the Excel formula row in `build_model` → Assumptions sheet + P&L/BS sheet.
5. The normalise and validate functions pick it up automatically.

### Add a new projection series

1. Add the field (as `list[float]`) to `Projections` Pydantic class.
2. Add it to the prompt's STEP 3 instructions.
3. Add validation for its length in `validate_model_output`.
4. Add the sheet row in `build_model` → Balance Sheet or Cash Flow sheet.

### Change the valuation blend

Adjust the `blended_fair_value` instruction in the prompt (STEP 6) — currently `DCF 40%, PE 30%, EV/EBITDA 30%`. Also update the Cover sheet description in `build_model`.

### Change the SAARTHI scoring

SAARTHI is purely a prompt instruction — the 7 dimensions and 100-point total are in the prompt's STEP 8. The max scores per dimension (S:15, A:15, A:15, R:15, T:10, H:15, I:15) can be changed in the prompt. The Pydantic schema (`SaarthiScores`) just stores floats — no validation of individual component caps.

### Add a new Excel sheet

Add a `mk("Sheet Name")` call inside `build_model`, write cells using openpyxl, and append to the workbook. The sheet will appear after the existing 8 model sheets.

### Change the Claude model

Update `MODEL_NAME = "claude-sonnet-4-20250514"` at line 47 of `financial_model_v5.py`. The streaming call and web search tool are model-agnostic. Haiku would be too weak for 5-year financial modelling; Opus 4 would work but cost ~5x more.

### Migrate existing tickers to new JSON schema

To refresh all existing models with the new schema, POST to `/generate` for each ticker. The server returns `storage_url` and `json_storage_url` on success. You can script this:

```bash
for TICKER in ANGELONE GRAVITA TCS ICICIBANK ...; do
  curl -s -X POST http://localhost:8500/generate \
    -H "Content-Type: application/json" \
    -d "{\"nse_symbol\":\"$TICKER\",\"company_name\":\"...\",\"sector\":\"...\"}"
done
```

---

## 14. Troubleshooting guide

### `❌ Screener login failed — /dash/ returned 403`
Screener blocked the IP or the credentials are wrong. Check `SCREENER_USERNAME`/`SCREENER_PASSWORD` in the systemd override conf. Screener may require re-login from a browser first to clear a CAPTCHA block.

### `❌ Screener login failed — /dash/ redirected to login page`
Wrong credentials. Update `/etc/systemd/system/financial-model.service.d/*.conf` and `systemctl daemon-reload && systemctl restart financial-model`.

### `❌ Got HTML instead of Excel from Screener export`
Screener changed the export flow or the session expired mid-request. Usually self-heals on retry.

### `Financial model JSON invalid for {TICKER}: ...`
The Pydantic validation failed. The error message lists the specific fields. Common causes:
- Claude returned `"projection_years": 5` (int instead of list) — prompt wording issue.
- A `dict[str,float]` assumption has only 3 years instead of 5 — normalise should have caught this; check if the year labels in `assumptions.projection_years` match the dict keys.
- `historical_ratios.*` list length doesn't match `years` — Claude returned mismatched lengths.

Fix: re-run `/generate`. If it fails repeatedly, look at the raw Claude output in the logs (`journalctl -u financial-model -n 100`).

### `json_storage_url` is null in server response
The JSON sidecar file was not found at the flat output path. This happens if:
- The model was generated with old code (pre-v5.1.0) — JSON is in the per-ticker subdir, not flat.
- `generate_financial_model` failed before writing the JSON (check server logs).

Fix for old models: `cp /opt/financial-model/output/{TICKER}/{TICKER}_model.json /opt/financial-model/output/{TICKER}_model.json` then POST `/storage/{TICKER}`.

### Service not starting (`systemctl status financial-model` shows failed)
```bash
journalctl -u financial-model -n 50 --no-pager
```
Common cause: Python syntax error in newly deployed script. Run `python3 -m py_compile` on both files before restarting.

### Excel opens but cells show `#VALUE!` or `0`
openpyxl formula strings reference the `Data Sheet` by name. If you renamed or deleted the Screener sheet, formulas break. The original Screener sheets must be kept intact.

### `model_version` on `/health` still shows old version
systemd cached the old process. Run `systemctl restart financial-model` (not reload).

### Stage 1/2 gets empty financial model context
Check `research_sessions.financial_model_json_url` in Supabase — it may be null (model was generated before wiring was added, or `mirrorFinancialModelToStorage` failed silently). Fix: manually POST `/proxy/fm/storage/{TICKER}` from the browser console or re-run the financial model step in the pipeline UI.
