# PPT Generator Handoff

This document summarizes how the current PPT generation flow works, how it was debugged using `GRAVITA`, what code changes were made, what is currently working, and what still needs attention.

## Purpose

This file is meant to help another agent take over PPT-generator work without needing to rediscover:

- where the PPT workflow starts
- which files matter
- how the `GRAVITA` test case was set up
- what slide-specific fixes were already implemented
- what known issues still remain

## Main Workflow

### Frontend entry point

The PPT flow is triggered from:

- [src/components/pipeline/PostProductionPanel.tsx](/abs/path/c:/Users/pratik/tikona-research-os-2/src/components/pipeline/PostProductionPanel.tsx)

It calls:

- `generatePptx({ reportId, sessionId, useMock })`

from:

- [src/lib/api.ts](/abs/path/c:/Users/pratik/tikona-research-os-2/src/lib/api.ts)

### PPT service

The local PPT service is:

- [scripts/ppt_service/main.py](/abs/path/c:/Users/pratik/tikona-research-os-2/scripts/ppt_service/main.py)

Endpoints:

- `GET /health`
- `POST /preview-placeholders`
- `POST /generate-pptx`

### Core generator

The main orchestration file is:

- [scripts/ppt_service/pptx_generator.py](/abs/path/c:/Users/pratik/tikona-research-os-2/scripts/ppt_service/pptx_generator.py)

Key responsibilities:

- fetch `research_reports`, `research_sessions`, `research_sections`
- download model JSON from Supabase Storage
- download model Excel from Supabase Storage
- build replacement values
- fill `master_template.pptx`
- inject custom charts/images
- inject Excel sheet visuals where still needed
- clean up surviving placeholders
- upload final PPTX to Supabase Storage
- update `research_reports`

### Excel injection helper

The Excel fallback/image injector is:

- [scripts/ppt_service/excel_injector.py](/abs/path/c:/Users/pratik/tikona-research-os-2/scripts/ppt_service/excel_injector.py)

It does:

- `win32com` path if available
- otherwise `openpyxl + matplotlib` worksheet rendering

Important limitation:

- `openpyxl` does not preserve embedded Excel charts well
- that is why several slide visuals were switched to direct Python-generated charts instead of copied Excel chart objects

## Gravita Test Case

### IDs used for testing

- `report_id = 809150c5-c369-4570-bf83-3d541d64f520`
- `session_id = ba4c7c90-960f-4d13-a67a-4d64729b6aea`
- ticker: `GRAVITA`

### Required files in storage

The PPT service expects these in Supabase Storage bucket `research-reports-html`:

- `financial-models/GRAVITA/GRAVITA_model.xlsx`
- `financial-models/GRAVITA/GRAVITA_model.json`

### Local files used during debugging

- [output/GRAVITA_model.xlsx](/abs/path/c:/Users/pratik/tikona-research-os-2/output/GRAVITA_model.xlsx)
- [output/GRAVITA_model.json](/abs/path/c:/Users/pratik/tikona-research-os-2/output/GRAVITA_model.json)

### Typical local run commands

Start service:

```powershell
.\.venv\Scripts\python.exe scripts\ppt_service\main.py
```

Preview placeholders:

```powershell
Invoke-RestMethod -Method Post `
  -Uri http://localhost:8501/preview-placeholders `
  -ContentType "application/json" `
  -Body '{"reportId":"809150c5-c369-4570-bf83-3d541d64f520","sessionId":"ba4c7c90-960f-4d13-a67a-4d64729b6aea"}' | ConvertTo-Json -Depth 6
```

Generate PPT:

```powershell
Invoke-RestMethod -Method Post `
  -Uri http://localhost:8501/generate-pptx `
  -ContentType "application/json" `
  -Body '{"reportId":"809150c5-c369-4570-bf83-3d541d64f520","sessionId":"ba4c7c90-960f-4d13-a67a-4d64729b6aea","useMock":false}' | ConvertTo-Json -Depth 6
```

## High-Level Architecture Decisions

### What should come from JSON/model data vs Excel

The current direction is:

- story/narrative charts: generate in Python with `matplotlib`
- dense financial tables or raw model-sheet visuals: still okay to inject/render from Excel
- avoid depending on embedded Excel chart objects when possible

This is more stable and easier to style.

### Why charts moved away from Excel

`openpyxl` could render sheet cells as images, but not embedded charts reliably.

This caused:

- blank chart slides
- tiny worksheet-like snapshots
- poor fit inside PPT placeholders

So several slides were rebuilt using model-driven chart generation.

## Important Code Paths

### Placeholder computation

Key function:

- `map_replacements(...)`

This builds values like:

- `company_name`
- `cmp`
- `target`
- `COMPANY_OVERVIEW`
- `investment_thesis`
- `peer_comparision`
- `management_content`
- `financial_commentry`

### Template fill

Key function:

- `fill_master_template(...)`

This:

- walks template shapes
- replaces text placeholders
- preserves Excel placeholders
- supports custom slide-specific shape logic
- queues image insertions

### Cleanup

Key function:

- `_cleanup_excel_placeholders(...)`

This replaces surviving placeholders like:

- `{{financial_summary_image}}`
- `{{pie_chart_1}}`
- `{{pie_chart_2}}`

with fallback text.

Important implication:

- if a custom image render fails, cleanup will overwrite the placeholder with explanatory text

## Main Fixes Already Implemented

### 1. Slide 2 and Slide 3 story charts moved to Python-rendered visuals

Problem:

- Excel fallback path was poor for chart slides
- some slides appeared blank or showed sheet snapshots

Fix:

- built direct chart collages with `matplotlib`
- inserted them into the PPT using custom image injection

Related logic:

- `_render_story_chart_collage(...)`
- `inject_story_chart_slides(...)`

### 2. Slide 2 now uses historical-only actual data

Problem:

- charts were showing projected years
- some future values showed zeros or looked misleading

Fix:

- only latest 5 actual years are used
- projected `E` years are excluded for these narrative chart slides

Helpers added:

- `_last_five_actual_periods(...)`
- `_last_five_actual_operational(...)`
- `_last_five_actual_segment_series(...)`

### 3. Slide 2 revenue fallback improved

Problem:

- `GRAVITA_model.json` did not provide all needed historical P&L series in the expected shape
- revenue chart could appear blank

Fix:

- historical revenue/EBITDA/PAT fallback extraction from Excel workbook

Helpers added:

- `_parse_excel_ref(...)`
- `_evaluate_excel_formula_cell(...)`
- `_extract_financial_chart_history_from_excel(...)`

### 4. Slide 2 fourth chart changed from EBITDA Margin to P/E

Requested change:

- replace EBITDA margin chart with P/E chart

Implementation:

- P/E derived using EPS and CMP fallback logic in the story-chart generator

### 5. Slide 4 thesis boxes improved

Problem:

- left box underused
- right-side boxes were too verbose

Fix:

- repeated `{{investment_thesis}}` placeholders are now handled by position
- left panel gets long-form text
- right three boxes get short summaries
- bottom small box gets a medium summary

Helper:

- `_replace_slide4_thesis_shapes(...)`

### 6. Slide 5 bottom boxes changed to KPI chips

Problem:

- small boxes were receiving long prose

Fix:

- bottom boxes now use compact metric strings

Helper:

- `_metric_chips(...)`

### 7. Slide 1 financial summary rebuilt as a custom rendered table

Problem:

- Excel-injected summary table was tiny and unreadable

Fix:

- slide 1 right-side panel now uses a custom image renderer
- built from workbook model data, not from a pasted Excel screenshot

Helpers:

- `_extract_summary_dashboard_from_excel_safe(...)`
- `_render_financial_summary_dashboard(...)`
- `inject_financial_summary_slide(...)`

Important bug that was fixed:

- the first version failed for Gravita because the renderer hard-required a row named `CFO/EBITDA`
- the safe extractor now tolerates alternate labels and computes missing values when needed

### 8. Slide 1 title above summary table removed

Requested change:

- remove extra title text above the custom table

Fix:

- removed summary chart header so the table uses more vertical space

## Slide 6 Work Done

### Goal

Slide 6 is `Company Overview`.

The template has:

- one large `{{COMPANY_OVERVIEW}}` box
- one smaller `{{COMPANY_OVERVIEW}}` box
- `{{pie_chart_1}}`
- `{{pie_chart_2}}`

### Text improvements implemented

Instead of using the exact same overview text in both boxes:

- top box now uses fuller narrative overview text
- bottom-right box is intended to use a denser operational/business summary

Helpers added:

- `_replace_slide6_overview_shapes(...)`
- `__slide6_top_overview`
- `__slide6_bottom_overview`

### Pie-chart work implemented

New helpers:

- `_build_slide6_pie_data(...)`
- `_render_pie_chart(...)`
- `inject_company_overview_slide(...)`

Current intended logic:

- `pie_chart_1` = `Revenue Mix %`
- `pie_chart_2` = `EBIT Mix %`

Revenue mix source:

- `fin_model["operational"]["revenue_mix_pct"]`

Current EBIT mix source:

- best available proxy from operational segment data
- not yet fully validated against the exact business-segment definition user wants

### Current status of slide 6

The service log shows:

- `Injected 2 company-overview pie visuals`

but the final PPT still shows fallback text inside the circles.

This means:

- injection is occurring at some point
- but the final deck still ends with fallback text shapes visible

An additional hardening step was added:

- reinject pie visuals after cleanup
- also allow targeting shapes that already contain fallback text

Log line added:

- `Re-injected %d company-overview pie visuals after cleanup`

### Slide 6 is still not fully resolved

Known unresolved issue:

- despite injection logs, final PPT may still show:
  - `Segment breakdown — see Excel model for details.`

This is the main outstanding issue at handoff time.

Likely next debugging direction:

1. inspect the generated PPT file after each stage, not just final upload
2. verify whether slide 6 shapes are being recreated or overwritten later
3. inspect shape order and IDs before and after cleanup
4. if needed, apply the pie insertion at the absolute last step before upload and save under unique shape detection rules

## Slide 7 to 10 Work Done

Work was extended beyond slide 6 and the generator now contains custom builders for slides 7, 8, 9, and 10 as well.

### Slide 7: Company Timeline

Problem before:

- slide 7 was just showing raw `{{COMPANY_TIMELINE}}` text
- it was not using the actual timeline structure already present in the workbook

Workbook source used:

- `Timeline` sheet in [output/GRAVITA_model.xlsx](/abs/path/c:/Users/pratik/tikona-research-os-2/output/GRAVITA_model.xlsx)

That sheet contains:

- `Year`
- `Event Category`
- `Description`
- `Strategic Impact`

New helpers added:

- `_read_timeline_rows(...)`
- `_render_timeline_table(...)`
- `inject_company_timeline_slide(...)`

Implementation notes:

- timeline is now rendered as a branded table image
- event categories get colored category cells
- slide 7 placeholder handling was updated so `{{COMPANY_TIMELINE}}` does not get replaced by raw text before custom injection

Current expected behavior:

- slide 7 should use a custom generated timeline visual instead of the old overview paragraph

### Slide 8: Business Model

Problem before:

- slide 8 boxes were mostly filled with short metric chips like:
  - market cap
  - CMP
  - target price
  - staff count
- this made the slide look empty and unhelpful

Fix:

- slide 8 now derives six business-model cards from:
  - `business_model` section first
  - otherwise other nearby narrative fields like `idea`, `competitive`, `thesis`

Implementation details:

- `map_replacements(...)` now builds `business_cards`
- placeholders `p1..p6` now use those richer business-model snippets instead of KPI chips

Goal:

- each box should feel like an actual business-model insight card, not just a number chip

### Slide 9: Competitive Advantage

Problem before:

- left-side text existed but was thin
- right-side chart placeholders showed fallback text like:
  - `Competitive positioning chart — see Excel model for details.`

Workbook source used:

- `Peer_Compare` sheet in [output/GRAVITA_model.xlsx](/abs/path/c:/Users/pratik/tikona-research-os-2/output/GRAVITA_model.xlsx)

Current custom chart plan implemented:

- top chart: `Revenue FY26A (₹ Cr)` peer comparison
- bottom chart: `EBITDA Margin FY26A` peer comparison

New helpers added:

- `_read_peer_compare_sections(...)`
- `_render_peer_bar_chart(...)`
- `inject_competitive_advantage_slide(...)`

Important template observation:

- slide 9 contains duplicate/overlapping placeholders for `{{competitive_chart_1}}`
- injection logic groups overlapping chart-placeholder shapes by vertical position and replaces the grouped area with one rendered chart

Current expected behavior:

- right side should use two custom bar-chart visuals, not fallback text

### Slide 10: Peer Comparison

Problem before:

- left-side table visual was coming from the Excel snapshot renderer
- it was readable but still worksheet-like and not as controlled as the custom renders

Workbook source used:

- `Peer_Compare` sheet in [output/GRAVITA_model.xlsx](/abs/path/c:/Users/pratik/tikona-research-os-2/output/GRAVITA_model.xlsx)

Structured sections parsed from that sheet:

- `REVENUE COMPARISON (₹ Cr)`
- `EBITDA MARGIN COMPARISON %`
- `PAT COMPARISON (₹ Cr)`
- `CURRENT VALUATION SNAPSHOT`

New helper added:

- `_render_peer_table(...)`
- `inject_peer_comparison_slide(...)`

Current expected behavior:

- slide 10 left-side peer table should become a custom-rendered branded table
- right-side text boxes still use `peer_para1` and `peer_para2`

### Injection sequencing for slides 7 to 10

These new injectors were wired into `generate_pptx_for_report(...)`:

- `inject_company_timeline_slide(...)`
- `inject_competitive_advantage_slide(...)`
- `inject_peer_comparison_slide(...)`

They run:

1. before Excel sheet injection
2. and then again after cleanup, similar to the hardened slide 6 approach

This was done because several earlier slides showed a pattern where:

- custom injection seemed to happen
- but fallback text or later PPT mutations still survived into the final deck

### New warning counters added to the response

The `/generate-pptx` response now also includes:

- `Company timeline injection count: ...`
- `Competitive advantage injection count: ...`
- `Peer comparison injection count: ...`

These are useful to confirm that the new code paths actually ran during generation.

### Current state of slides 7 to 10 at handoff time

Code has been added for:

- slide 7 timeline visual
- slide 8 richer business-model cards
- slide 9 custom peer charts
- slide 10 custom peer table

However, these still need visual validation with fresh regenerated screenshots.

So the current status is:

- implemented in code
- smoke-tested locally for renderer output
- not yet fully visually verified slide-by-slide in final uploaded PPT

## Useful Observations from Gravita Workbook

### Operational data

`Operational_Data` contains:

- `Lead Recycling (MT)`
- `Aluminum Recycling (MT)`
- `Plastic Recycling (MT)`
- `Capacity Utilisation %`
- `Countries of Operation`
- `Plants India`
- `Plants Overseas`
- `REVENUE MIX %`
- `GEOGRAPHY MIX %`

### Segment sheet

`Segments` contains:

- `SEGMENT REVENUE`
- `SEGMENT PBIT`
- `SEGMENT CAPITAL EMPLOYED`

However, this workbook appears to use a different segment taxonomy in places, so care is needed before using it directly for the user-facing pie charts.

### Op_Charts

`Op_Charts` contains chart-data blocks like:

- `ChartData 1 — Revenue/EBITDA/PAT`
- `ChartData 3 — Revenue by Segment`
- `ChartData 4 — ROE/ROCE`
- `ChartData 5 — CFO/Capex/FCF`
- `ChartData 6 — EPS Trend`

This sheet may be a good source for later chart consistency work.

## Service Log Signals That Matter

When debugging, these log lines are high-signal:

- `Financial summary dashboard render failed before template fill: ...`
- `Injected 2 story-chart slide visuals`
- `Injected 2 company-overview pie visuals`
- `Financial summary dashboard injected on slide 1 (...)`
- `Re-injected 2 company-overview pie visuals after cleanup`
- `Cleaned up unreplaced Excel placeholder: ...`

If a fallback text appears in the final PPT, cleanup almost certainly saw that placeholder survive.

## Known Working Areas

These are known to be improved and generally working better than before:

- slide 1 custom financial summary table
- slide 2 custom financial charts
- slide 3 custom operational charts
- slide 4 thesis layout shaping
- slide 5 KPI-chip behavior

## Known Remaining Problems

### Slide 6

- pie charts still not sticking reliably in final PPT
- top and bottom text may still need further tuning

### Slides 7 to 10 still need verification

Recent code was added for:

- slide 7 custom timeline
- slide 8 business-model content cards
- slide 9 peer charts
- slide 10 peer table

These should be considered pending visual QA until regenerated screenshots confirm:

- the custom visuals actually persist into the final deck
- layout fit is acceptable
- the new text density is appropriate
- Excel fallback content is no longer leaking through where custom renderers should win

### Other unresolved placeholders

Recent logs still show cleanup for:

- `{{competitive_chart_1}}`
- `{{competitive_chart_2}}`
- `{{probability_weight_table}}`

These slides likely still need custom rendering or proper Excel/image injection.

### PDF generation

Local service log shows:

- LibreOffice not installed

So PDF export is skipped locally unless `soffice` becomes available.

## Suggested Next Steps for Another Agent

1. Finish slide 6 by making the pie visuals persist in the final PPT.
2. Verify whether bottom-right overview text is actually using `__slide6_bottom_overview` in final output.
3. Validate slides 7 to 10 visually after regeneration.
4. If slide 9 still shows fallback text, inspect grouped placeholder replacement behavior on overlapping competitive-chart shapes.
5. If slide 10 still shows Excel-derived visuals, inspect whether `inject_peer_comparison_slide(...)` is being overwritten by later save passes.
6. Review slides with remaining cleanup placeholders:
   - competitive charts
   - probability-weight table
7. Decide whether slide 6 EBIT pie should use:
   - segment PBIT from workbook
   - operational volume proxy
   - or a user-approved alternate definition
8. Continue slide-by-slide review with screenshots rather than broad refactors.

## Files Most Relevant for Continuing Work

- [scripts/ppt_service/pptx_generator.py](/abs/path/c:/Users/pratik/tikona-research-os-2/scripts/ppt_service/pptx_generator.py)
- [scripts/ppt_service/excel_injector.py](/abs/path/c:/Users/pratik/tikona-research-os-2/scripts/ppt_service/excel_injector.py)
- [scripts/ppt_service/main.py](/abs/path/c:/Users/pratik/tikona-research-os-2/scripts/ppt_service/main.py)
- [master_template.pptx](/abs/path/c:/Users/pratik/tikona-research-os-2/master_template.pptx)
- [output/GRAVITA_model.xlsx](/abs/path/c:/Users/pratik/tikona-research-os-2/output/GRAVITA_model.xlsx)
- [output/GRAVITA_model.json](/abs/path/c:/Users/pratik/tikona-research-os-2/output/GRAVITA_model.json)
- [CLAUDE.md](/abs/path/c:/Users/pratik/tikona-research-os-2/CLAUDE.md)
