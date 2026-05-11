# CLAUDE.md — Tikona Research OS

Repository guide for the AI-powered equity research platform. The flagship feature is the **Research Pipeline** at `/admin/pipeline`, which produces institutional-grade Indian equity research reports end-to-end.

---

## 1. Stack

- **Frontend:** Vite + React + TypeScript, TailwindCSS, Radix UI, React Query, Sonner toasts.
- **Backend-as-a-Service:** Supabase (Postgres + Auth + Storage). All DB access goes through `src/lib/supabase.ts`.
- **LLM provider:** Anthropic SDK (`@anthropic-ai/sdk`) called **directly from the browser** with `dangerouslyAllowBrowser: true` and `VITE_ANTHROPIC_API_KEY`. Default model: `claude-sonnet-4-20250514`. Vault summarization uses Haiku (`claude-haiku-4-5-20251001`).
- **Web search:** Anthropic native `web_search_20250305` tool (max 20 uses per call). All Stage 0/1/2 calls use streaming (`client.messages.stream`) to avoid the 10-minute API timeout.
- **External services (n8n at `https://n8n.tikonacapital.com`):**
  - `POST /webhook/create-folder` — creates the Google Drive vault, returns Drive files.
  - `POST /webhook/delete-file`, `POST /webhook/upload-document` — Drive file ops.
  - `POST /webhook/generate-financial-model` — fires the Python financial model script.
  - `POST /webhook/fetch-vault-pdfs` — server-side Drive→base64 PDF fetcher (`VITE_VAULT_PDF_WEBHOOK_URL`).
  - `POST /webhook/generate-media-script`, `/synthesize-podcast`, `/generate-video` — post-production media.
- **Local services proxied via Vite (`vite.config.ts`):**
  - `/proxy/fm` → `http://72.61.226.16:8500` — financial model REST service (Python; see [scripts/financial_model_server.py](scripts/financial_model_server.py)).
  - `/proxy/ppt` → `http://localhost:8501` — PPTX/PDF generator service (see [scripts/ppt_service/](scripts/ppt_service/) and [PPT-generator/](PPT-generator/)).
  - `/proxy/n8n` → public n8n.

---

## 2. The `/admin/pipeline` flow — A to Z

Page entry: [src/pages/ResearchPipeline.tsx](src/pages/ResearchPipeline.tsx) (route registered in [src/App.tsx:64](src/App.tsx#L64)). All orchestration state lives on this single component.

### 2.1 State machine
12 states defined in [src/types/pipeline.ts](src/types/pipeline.ts):

```
company_selected
  → vault_creating → vault_ready
       ↘ financial_model_generating ↗   (optional, returns to vault_ready)
  → stage0_generating → stage0_review → stage0_approved
  → stage1_generating → stage1_review → stage1_approved
  → stage2_generating → stage2_review → stage2_approved
  → published
```

`PIPELINE_TRANSITIONS` enforces legal edges; `transitionPipelineStatus()` in [src/lib/pipeline-api.ts](src/lib/pipeline-api.ts) validates before writing `research_sessions.pipeline_status`.

### 2.2 Step-by-step

1. **Company search** — `useCompanySearch` queries `master_companies`. Selecting a row fetches `EquityUniverse` financials (`useCompanyFinancials`) and auto-fills sector by fuzzy-matching against the `SECTORS` list.
2. **Begin Pipeline (`handleCreateSession`)** — Inserts a `research_sessions` row with `pipeline_status='company_selected'` and `selected_model`, then immediately transitions to `vault_creating` and calls `createVault(nse_symbol, sector)` (n8n `/webhook/create-folder`). Vault Drive files are normalized via `processVaultResponse` and persisted to `session_documents`. State advances to `vault_ready`.
3. **Background vault summarization** — `summarizeVaultDocuments()` runs fire-and-forget after vault creation. It downloads each PDF via the `fetch-vault-pdfs` webhook (≤32 MB per PDF) and sends it directly to Haiku as a `document` content block — Haiku reads text, charts, and diagrams. Result is cached in `research_sessions.condensed_briefing` and reused as a primary source in Stage 1 + Stage 2.
4. **Optional: Generate Financial Model** — `handleGenerateFinancialModel` POSTs to `/webhook/generate-financial-model`, then `mirrorFinancialModelToStorage()` calls `/proxy/fm/storage/{TICKER}` to copy the resulting `{TICKER}_model.xlsx` into Supabase Storage. URL persisted on `research_sessions.financial_model_file_url`. The Python script is in [scripts/financial_model_v5.py](scripts/financial_model_v5.py).
5. **Stage 0 — Sector Framework** (`runStage0` in [src/lib/anthropic-pipeline.ts](src/lib/anthropic-pipeline.ts)):
   - **Path A (cached):** `getSectorPlaybook(sector)` returns existing `sector_playbooks` row → markdown extracted from `ai_writing_instructions.framework_markdown` → 0 tokens, instant.
   - **Path B (generate):** Anthropic + web search → 9-section sector brief (snapshot, KPIs, value chain, competitive landscape, regulatory, growth drivers, risks, valuation, analyst checklist). Result upserted to `sector_playbooks`; old version archived to `sector_playbook_versions`. Stored on session at `sector_framework`.
6. **Stage 1 — Investment Thesis** (`runStage1`): Single Anthropic + web-search call. Prompt is composed of (a) `buildFreshnessPreamble()` enforcing current Indian FY/quarter anchoring, (b) the SAARTHI 100-point scorecard system prompt, (c) context block: company info + sector framework slice + `formatFinancialContext(financials)` + cached vault briefing. Output saved to `research_sessions.thesis_output`.
7. **Stage 2 — Full Report** (`runStage2`): Single Anthropic + web-search call (`maxTokens: 32000`). The user prompt requests 18 deliverables separated by the literal string `===SECTION===`: 12 narrative sections (Investment Rationale, Company Background, Business Model, Management, Corporate Governance, Industry Overview, Tailwinds, Demand Drivers, Industry Risks, SAARTHI, Entry/Review/Exit Strategy, Scenario Analysis) + 6 atomic data points (Rating, Target Price, Upside %, Market Cap, Cap Category, CMP). `parseSectionsFromResponse()` splits on the separator and matches titles to `REPORT_SECTION_DEFS`. Each parsed section becomes a `research_sections` row (stage='stage2').
8. **Review/Approve loops** — Each `_review` state allows regenerate (loops back to `_generating`) or approve. `handleEditFramework`/`Thesis`/`ReportSection` persist inline edits.
9. **On stage2 approve** — A `research_reports` row is created via `createResearchReport`, and `PostProductionPanel` mounts.
10. **Post-production** ([src/components/pipeline/PostProductionPanel.tsx](src/components/pipeline/PostProductionPanel.tsx)):
    - **Step 1 — PPTX:** `generatePptx({ reportId, sessionId, useMock })` → `${PPT_SERVICE_URL}/generate-pptx`. Service returns both `pptx_file_url` and `pptx_pdf_file_url`. Health check probes `/health` on mount.
    - **Step 2 — Podcast:** n8n `/generate-media-script` then `/synthesize-podcast`. Both are async — UI polls `research_reports.podcast_script` / `audio_file_url` columns via `pollSupabaseColumn`.
    - **Step 3 — Video:** n8n `/generate-video`, polls `research_reports.video_file_url` (longer timeout).
    - **Publish:** Plan picker (`midcap_wealth` | `smallcap_alpha` | `sme_emerging`) → `publishReport(reportId, plan)` → state moves to `published`.
11. **Telegram recommendation** — On the published screen (and inside `PostProductionPanel`), `handleSendTelegramRecommendation` reads CMP / target price / rating from the report's `cs_*` custom columns (falling back to stage2 section content), builds a `RecommendationRating`, and calls `createRecommendation({ ..., send_telegram: true })`. `hasRecommendationForSession` prevents double-sends on resume.
12. **Resume / delete** — `listPipelineSessions` populates the "Recent Pipelines" list; clicking a row sets `sessionId`, and the big `useEffect` on `[sessionId]` rehydrates vault/sections/recommendation state.

### 2.3 Prompt overrides
`PromptEditor` ([src/components/pipeline/PromptEditor.tsx](src/components/pipeline/PromptEditor.tsx)) loads/saves user-scoped prompts via `getPipelinePrompt`/`savePipelinePrompt` (table `pipeline_prompts`). Overrides flow as `PromptOverrides` into `runStage0/1/2` and replace `DEFAULT_PROMPTS.stageN.system|user`. Template variables `{{COMPANY}}`, `{{NSE_SYMBOL}}`, `{{SECTOR}}` are substituted before the call.

### 2.4 Freshness enforcement
Every stage prepends `buildFreshnessPreamble()` to the system prompt. It computes today's Indian FY (`FY{YY}`), current quarter, most recent reported quarter, and injects hard rules: search with quarter-specific terms, treat data >2 quarters old as stale, never default to FY24/FY25 unless those are the latest actuals, flag missing data with `⚠`.

---

## 3. Key DB tables (see [supabase_migration.sql](supabase_migration.sql) and [src/types/database.ts](src/types/database.ts))

- `research_sessions` — pipeline state container. Important columns: `session_id`, `pipeline_status`, `current_state`, `sector_framework` (jsonb), `condensed_briefing` (vault summary), `thesis_output`, `vault_folder_id`/`url`, `financial_model_file_url`, `selected_model`, `total_tokens_used`. Legacy stage native columns (`sector_playbook_*`, `thesis_original`, `final_report_*`) remain for back-compat.
- `research_sections` — normalized stage outputs, one row per section (`stage`, `section_key`, `section_title`, `content`, `sort_order`, `tokens_used`). Stage 2 writes 18 rows.
- `research_reports` — post-production artifacts (`pptx_file_url`, `pptx_pdf_file_url`, `podcast_script`, `audio_file_url`, `video_file_url`, `is_published`, `plan`, plus `cs_*` custom-section columns extracted from the LLM output).
- `sector_playbooks` + `sector_playbook_versions` — cached/versioned sector frameworks. Markdown stored at `ai_writing_instructions.framework_markdown`.
- `sectors`, `sector_knowledge`, `sector_knowledge_embeddings` — pre-seeded sector reference data and SKB embeddings.
- `skb_suggested_updates` — pipeline-generated suggestions to enrich the SKB, reviewed by admin.
- `pipeline_prompts` — per-user stage prompt overrides.
- `session_documents` — vault Drive files mirrored per session.
- `recommendations` — Telegram recommendation log.
- `master_companies`, `equity_universe` — universe + financial snapshots powering the search/snapshot card.

⚠ Do **not** assume stage outputs live only in `research_sections` — the flat columns on `research_sessions` (`sector_framework`, `thesis_output`) and the structured columns (`final_report_*`, `thesis_*`) are also written and read on resume. Keep both in sync when touching pipeline writes.

---

## 4. Files you'll touch most often

- [src/pages/ResearchPipeline.tsx](src/pages/ResearchPipeline.tsx) — page + all orchestration handlers.
- [src/lib/anthropic-pipeline.ts](src/lib/anthropic-pipeline.ts) — Anthropic calls, `DEFAULT_PROMPTS`, freshness preamble, vault summarization, Stage 2 parser.
- [src/lib/pipeline-api.ts](src/lib/pipeline-api.ts) — Supabase CRUD + state-machine transition gate.
- [src/types/pipeline.ts](src/types/pipeline.ts) — state machine, types, model list.
- [src/lib/api.ts](src/lib/api.ts) — n8n integrations (`createVault`, `generateFinancialModel`, `mirrorFinancialModelToStorage`, `generatePptx`, `publishReport`, etc.).
- [src/components/pipeline/PostProductionPanel.tsx](src/components/pipeline/PostProductionPanel.tsx) — PPTX / podcast / video / publish / Telegram.
- [src/components/pipeline/StageReview.tsx](src/components/pipeline/StageReview.tsx), [PipelineProgressBar.tsx](src/components/pipeline/PipelineProgressBar.tsx), [PromptEditor.tsx](src/components/pipeline/PromptEditor.tsx) — UI primitives for the pipeline.
- [scripts/financial_model_v5.py](scripts/financial_model_v5.py), [scripts/financial_model_server.py](scripts/financial_model_server.py) — Python financial model generator + REST wrapper.
- [scripts/ppt_service/](scripts/ppt_service/), [PPT-generator/pipeline/](PPT-generator/pipeline/) — PPTX/PDF generation service (`orchestrate.py`, `renderer.py`, `pages.py`, `editorial.py`, `critique.py`, `pdf_export.py`).

---

## 5. Environment variables

| Var | Purpose |
|---|---|
| `VITE_SUPABASE_URL`, `VITE_SUPABASE_ANON_KEY` | Supabase client. |
| `VITE_ANTHROPIC_API_KEY` | **Required** for the pipeline. Used directly from the browser. |
| `VITE_VAULT_PDF_WEBHOOK_URL` | Optional override for the Drive→base64 PDF fetcher (defaults to `https://n8n.tikonacapital.com/webhook/fetch-vault-pdfs`). |
| `VITE_PPT_SERVICE_URL` | Optional override for the PPTX service base URL (defaults to `/proxy/ppt`). |

n8n base (`https://n8n.tikonacapital.com`) and FM service IP (`72.61.226.16:8500`) are hardcoded in `src/lib/api.ts` and `vite.config.ts`.

---

## 6. Conventions and gotchas

- **Streaming is mandatory** for Stage 1/2 LLM calls — non-streaming hits Anthropic's 10-minute SSE timeout when web search + 32K outputs are combined.
- **No markdown tables** anywhere in LLM outputs — every prompt forbids `|` pipes. Reports use bullets/numbered lists only. PPT renderer assumes this.
- **Indian fiscal year context is non-negotiable** — never bypass `buildFreshnessPreamble()`. CMP / multiples / quarters must come from web_search, not training data.
- **State transitions must validate** — always go through `transitionPipelineStatus(sessionId, next, current)` so `canTransition` enforces the FSM. Resetting on error must reverse-transition or the session sticks.
- **Vault docs are authoritative** for company-specific facts — Stage 1/2 prompts explicitly say "use the Vault Briefing as primary source, cross-check web_search against it."
- **Stage 2 output parsing** depends on the literal `===SECTION===` separator. If you change it, update both the prompt and `parseSectionsFromResponse` together.
- **Async webhook polling pattern**: n8n media steps return 200 immediately; the result lands in a Supabase column later. Use `pollSupabaseColumn(column, maxAttempts, intervalMs)` rather than trying to make the webhook synchronous.
- **`dangerouslyAllowBrowser: true`** is intentional — Anthropic key is exposed to the client. Production hardening (proxy through a backend) is out of scope for now; do not "fix" it without product sign-off.
- **Never commit `.env`** — `c:\Users\pratik\tikona-research-os-2\.env` holds the live Anthropic + Supabase keys.
