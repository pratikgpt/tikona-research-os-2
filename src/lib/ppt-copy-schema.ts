// Pure schema + prompt builder + sanitiser for the PPT copywriting LLM pass.
// Lives in its own file (no Vite/browser deps) so both the browser path in
// anthropic-pipeline.ts AND the Node CLI (scripts/generate_ppt_copy.ts) can
// share one source of truth.

export type PptCopyKind = 'line' | 'card' | 'panel' | 'bullets';
export type PptCopySchema = Record<string, { kind: PptCopyKind; max: number; instruction: string }>;

/**
 * Per-placeholder contract for the PPT copywriting pass.
 * Every key maps 1:1 to a `{{placeholder}}` token in master_template.pptx.
 * Numeric / atomic fields (CMP, target, market cap, scenario prices) are
 * intentionally absent — those stay deterministic in the Python service.
 *
 * `max` is an absolute hard cap (clipped at word boundary in the sanitiser);
 * the per-field instruction states the desired natural length (e.g. "~1300
 * chars"). The cap is slightly above the target so the LLM is not forced to
 * end mid-sentence.
 */
export const PPT_COPY_SCHEMA: PptCopySchema = {
  // ── Cover (slide 1) ─────────────────────────────────────────────────────
  tagline: { kind: 'line', max: 120, instruction: 'One-line positioning, 10-16 words. No period at end.' },
  investment_thesis_s1: { kind: 'panel', max: 1500, instruction: 'Slide 1 left panel. ~1300 chars in flowing prose (single paragraph). Lead with the strongest reason to own; cover scale, near-term catalyst, margin trajectory, and valuation; close with the 12-month target setup. Wrap the 4-8 most important phrases (numbers, growth catalysts, competitive position) in **double-asterisks** to emphasise them. Each wrapped phrase should be 2-6 words. Do NOT bold full sentences.' },
  investment_ideas_1: { kind: 'card', max: 380, instruction: 'Slide 1 right card #1 — Market Position. ~350 chars. Concrete: scale, share, geographies, plant count, capacity.' },
  investment_ideas_2: { kind: 'card', max: 380, instruction: 'Slide 1 right card #2 — Catalyst. ~350 chars. The single biggest near-term catalyst with numbers, dates, capex.' },
  investment_ideas_3: { kind: 'card', max: 380, instruction: 'Slide 1 right card #3 — Margin / returns. ~350 chars. EBITDA / PAT margin trajectory, guidance, and quality.' },
  investment_ideas_4: { kind: 'card', max: 380, instruction: 'Slide 1 right card #4 — Valuation gap. ~350 chars. Multiple, target, upside, analyst posture.' },

  // ── Investment Thesis (slide 4) ─────────────────────────────────────────
  investment_thesis_heading_s4: { kind: 'line', max: 60, instruction: 'Slide 4 heading. 3-6 words capturing the thesis in one phrase.' },
  investment_thesis_s4: { kind: 'panel', max: 2700, instruction: 'Slide 4 left panel. ~2500 chars in 2-3 paragraphs separated by a blank line (\\n\\n between paragraphs). Long-form thesis — lead with the operating story, walk through catalysts and margin trajectory, close with valuation setup. DO NOT mirror investment_thesis_s1 wording or arc.' },
  key_catalyst_heading_1: { kind: 'line', max: 50, instruction: '2-4 word heading for catalyst card #1.' },
  key_catalyst_1: { kind: 'card', max: 380, instruction: 'Slide 4 right catalyst #1. ~350 chars. Specific, dated, quantified.' },
  key_catalyst_heading_2: { kind: 'line', max: 50, instruction: '2-4 word heading for catalyst card #2 (different angle from #1).' },
  key_catalyst_2: { kind: 'card', max: 380, instruction: 'Slide 4 right catalyst #2. ~350 chars. Different facet than #1.' },
  key_catalyst_heading_3: { kind: 'line', max: 50, instruction: '2-4 word heading for catalyst card #3 (different angle again).' },
  key_catalyst_3: { kind: 'card', max: 380, instruction: 'Slide 4 right catalyst #3. ~350 chars. Different facet than #1 and #2.' },
  saarthi_summary_s4: { kind: 'card', max: 850, instruction: 'Slide 4 bottom strip. ~800 chars. Pull together SAARTHI total score, the two-to-three highest-scoring dimensions with their specific evidence, and the resulting rating. Flowing prose.' },

  // ── Industry Overview (slide 5) ─────────────────────────────────────────
  industry_structure: { kind: 'bullets', max: 1600, instruction: 'Slide 5 left panel. 5-8 bullet points, one per line. Start each bullet with `• ` (Unicode bullet + space). No paragraphs, no headings. Each bullet ~120-200 chars. Cover market size, fragmentation, named competitors, regulatory/cost barriers, where consolidation is happening.' },
  key_industry_tailwinds: { kind: 'bullets', max: 1600, instruction: 'Slide 5 middle panel. 5-8 bullet points, one per line. Start each bullet with `• ` (Unicode bullet + space). No paragraphs, no headings. Each bullet ~120-200 chars. Sector-level tailwinds (regulation, EV / EPR / circular economy, supply-chain mandates) — not company-specific moves.' },
  key_industry_risks: { kind: 'bullets', max: 1600, instruction: 'Slide 5 right panel. 5-8 bullet points, one per line. Start each bullet with `• ` (Unicode bullet + space). No paragraphs, no headings. Each bullet ~120-200 chars. Sector-level risks — commodity, regulation, working-capital cycle, competitive intensity. Concrete, not generic.' },

  // ── KPI strip on slide 5 — 6 company-specific operating/financial tiles ─
  // CRITICAL: these are NOT the slide-1 chips (CMP / Target / Market Cap /
  // Category / SAARTHI). Pick six DIFFERENT, company-specific operating or
  // financial metrics: revenue, revenue CAGR, EBITDA margin, ROE / ROCE,
  // capacity / utilisation, plants / countries, debt-equity, FCF yield, etc.
  // Heading is a 1-3 word label. Value is the actual number with its unit.
  KPI_heading_1: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #1 heading — 1-3 words, e.g. "Revenue FY26", "EBITDA Margin". DIFFERENT from slide-1 chips.' },
  KPI_1: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #1 value — actual number with units, e.g. "₹4,265 Cr" or "11.8%".' },
  KPI_heading_2: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #2 heading — different metric from #1.' },
  KPI_2: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #2 value with units.' },
  KPI_heading_3: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #3 heading.' },
  KPI_3: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #3 value with units.' },
  KPI_heading_4: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #4 heading.' },
  KPI_4: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #4 value with units.' },
  KPI_heading_5: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #5 heading.' },
  KPI_5: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #5 value with units.' },
  KPI_heading_6: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #6 heading.' },
  KPI_6: { kind: 'line', max: 30, instruction: 'Slide 5 KPI #6 value with units.' },

  // ── Company Overview (slide 6) ──────────────────────────────────────────
  company_overview: { kind: 'panel', max: 1700, instruction: 'Slide 6 top panel. ~1600 chars. Founded year, segments, scale, geographies, plant count, employee count, market cap — anchor in facts, not adjectives.' },
  competitive_moat_1: { kind: 'card', max: 600, instruction: 'Slide 6 moat card #1. ~550 chars. One specific structural advantage explained with numbers.' },
  competitive_moat_2: { kind: 'card', max: 600, instruction: 'Slide 6 moat card #2. ~550 chars. A DIFFERENT structural advantage from #1.' },
  key_insights: { kind: 'card', max: 1600, instruction: 'Slide 6 bottom-right insights box. ~1500 chars. Operating snapshot: countries / plants / volume / utilisation / segment mix, plus the integrated-model takeaway.' },

  // ── Business Model (slide 8) ────────────────────────────────────────────
  business_model_1: { kind: 'card', max: 950, instruction: 'Slide 8 card #1. ~900 chars. Revenue model / segment mix.' },
  business_model_2: { kind: 'card', max: 950, instruction: 'Slide 8 card #2. ~900 chars. Geographic / customer footprint. Distinct from #1.' },
  business_model_3: { kind: 'card', max: 950, instruction: 'Slide 8 card #3. ~900 chars. Value-chain position / integration.' },
  business_model_4: { kind: 'card', max: 950, instruction: 'Slide 8 card #4. ~900 chars. Cost structure / unit economics.' },
  business_model_5: { kind: 'card', max: 950, instruction: 'Slide 8 card #5. ~900 chars. Capital intensity / asset base / funding model.' },
  business_model_6: { kind: 'card', max: 950, instruction: 'Slide 8 card #6. ~900 chars. Diversification / cyclicality hedge.' },

  // ── Competitive Advantage (slide 9) ─────────────────────────────────────
  competitive_advantage_heading: { kind: 'line', max: 60, instruction: 'Slide 9 heading. 3-6 words capturing the company\'s core competitive edge in one phrase.' },
  competitive_advantage: { kind: 'panel', max: 2900, instruction: 'Slide 9 left panel. ~2700 chars in 2-3 paragraphs separated by a blank line (\\n\\n between paragraphs). Why this company wins vs. the peers in the bar charts — scale, integration, mix, cost. Avoid repeating slide-1 phrasing.' },

  // ── Growth Catalysts & Timeline (slide 10) ──────────────────────────────
  investment_thesis_detailed_heading: { kind: 'line', max: 60, instruction: 'Slide 10 heading. 3-6 words framing the long-form thesis (different angle from the s4 heading).' },
  investment_thesis_detailed: { kind: 'panel', max: 3200, instruction: 'Slide 10 LEFT panel. ~3000 chars in 2-3 paragraphs separated by a blank line (\\n\\n between paragraphs). Long-form thesis narrative tying catalysts to numbers — different shape from the s1/s4 thesis panels (more arc, longer, structured by time horizon).' },
  key_catalyst: { kind: 'panel', max: 1200, instruction: 'Slide 10 RIGHT-TOP panel. ~1100 chars. The catalyst stack as continuous prose, NOT a duplicate of investment_thesis_detailed. Order catalysts by timing (immediate / 6-12m / multi-year).' },

  // ── Management Commentary (slide 11) ────────────────────────────────────
  management_commentry_heading_1: { kind: 'line', max: 40, instruction: 'Slide 11 heading #1 — 2-3 words.' },
  management_content_1: { kind: 'card', max: 440, instruction: 'Slide 11 card #1. ~400 chars. Capital allocation discipline.' },
  management_commentry_heading_2: { kind: 'line', max: 40, instruction: 'Slide 11 heading #2 — 2-3 words.' },
  management_content_2: { kind: 'card', max: 440, instruction: 'Slide 11 card #2. ~400 chars. Execution track record.' },
  management_commentry_heading_3: { kind: 'line', max: 40, instruction: 'Slide 11 heading #3 — 2-3 words.' },
  management_content_3: { kind: 'card', max: 440, instruction: 'Slide 11 card #3. ~400 chars. Expansion discipline / project execution.' },
  management_commentry_heading_4: { kind: 'line', max: 40, instruction: 'Slide 11 heading #4 — 2-3 words.' },
  management_content_4: { kind: 'card', max: 440, instruction: 'Slide 11 card #4. ~400 chars. Leadership quality / depth.' },
  management_commentry_heading_5: { kind: 'line', max: 40, instruction: 'Slide 11 heading #5 — 2-3 words.' },
  management_content_5: { kind: 'card', max: 440, instruction: 'Slide 11 card #5. ~400 chars. Funding approach / balance sheet.' },
  management_commentry_heading_6: { kind: 'line', max: 40, instruction: 'Slide 11 heading #6 — 2-3 words.' },
  management_content_6: { kind: 'card', max: 440, instruction: 'Slide 11 card #6. ~400 chars. Margin / guidance philosophy.' },
  management_commentry_heading_7: { kind: 'line', max: 40, instruction: 'Slide 11 heading #7 — 2-3 words.' },
  management_content_7: { kind: 'card', max: 440, instruction: 'Slide 11 card #7. ~400 chars. Strategic vision / long-term goals.' },
  management_commentry_heading_8: { kind: 'line', max: 40, instruction: 'Slide 11 heading #8 — 2-3 words.' },
  management_content_8: { kind: 'card', max: 440, instruction: 'Slide 11 card #8. ~400 chars. Risk controls / governance discipline.' },

  // ── Governance indicators (slide 12) ────────────────────────────────────
  indicators_1: { kind: 'card', max: 380, instruction: 'Slide 12 indicator #1. ~350 chars. One concrete governance signal with numbers.' },
  indicators_2: { kind: 'card', max: 380, instruction: 'Slide 12 indicator #2. ~350 chars. Different from #1.' },
  indicators_3: { kind: 'card', max: 380, instruction: 'Slide 12 indicator #3. ~350 chars.' },
  indicators_4: { kind: 'card', max: 380, instruction: 'Slide 12 indicator #4. ~350 chars.' },
  indicators_5: { kind: 'card', max: 380, instruction: 'Slide 12 indicator #5. ~350 chars.' },
  indicators_6: { kind: 'card', max: 380, instruction: 'Slide 12 indicator #6. ~350 chars.' },

  // ── Forecast / Financials / Valuations side panels (slides 13-15) ───────
  forecast_assumptions: { kind: 'panel', max: 2600, instruction: 'Slide 13 side panel. ~2400 chars in 2-3 paragraphs separated by a blank line (\\n\\n between paragraphs). Cover (i) volume/revenue trajectory, (ii) margin & cost assumptions, (iii) capex & working capital — one paragraph per theme. Numbers and FY labels throughout.' },
  financial_commentary: { kind: 'panel', max: 2150, instruction: 'Slide 14 side panel. ~2000 chars in 2-3 paragraphs separated by a blank line (\\n\\n between paragraphs). Para 1: P&L trajectory. Para 2: balance sheet & cash flow. Para 3: ratios / what is improving vs stretched. Use actual figures throughout.' },
  valuation_commentary: { kind: 'panel', max: 2150, instruction: 'Slide 15 side panel. ~2000 chars in 2-3 paragraphs separated by a blank line (\\n\\n between paragraphs). Para 1: P/E based fair value derivation. Para 2: EV/EBITDA or DCF cross-check. Para 3: peer benchmark + target setup. Concrete multiples and per-share numbers.' },

  // ── SAARTHI Framework (slide 16) ────────────────────────────────────────
  saarthi_s_content: { kind: 'card', max: 500, instruction: 'Slide 16 — Scalability of Core Engine. ~450 chars. Score header (e.g., "12/15") + substantive evidence on scalability with numbers.' },
  saarthi_a1_content: { kind: 'card', max: 500, instruction: 'Slide 16 — Addressable Market & Adjacency. ~450 chars. Score header + TAM / adjacency evidence.' },
  saarthi_a2_content: { kind: 'card', max: 500, instruction: 'Slide 16 — Asymmetric Pricing Power. ~450 chars. Score header + pricing / moat evidence.' },
  saarthi_r_content: { kind: 'card', max: 500, instruction: 'Slide 16 — Reinvestment Quality. ~450 chars. Score header + ROIC / capital deployment evidence.' },
  saarthi_t_content: { kind: 'card', max: 500, instruction: 'Slide 16 — Track Record Through Adversity. ~450 chars. Score header + historical resilience evidence.' },
  saarthi_h_content: { kind: 'card', max: 500, instruction: 'Slide 16 — Human Capital & Institutional DNA. ~450 chars. Score header + leadership / culture evidence.' },
  saarthi_i_content: { kind: 'card', max: 500, instruction: 'Slide 16 — Inflection Point Identification. ~450 chars. Score header + why now.' },
  saarthi_summary_s16: { kind: 'card', max: 760, instruction: 'Slide 16 bottom summary. ~700 chars. Aggregate SAARTHI verdict tying the seven dimensions to the rating, with the total score.' },

  // ── Scenario Analysis (slide 17) ────────────────────────────────────────
  bull_content: { kind: 'card', max: 650, instruction: 'Slide 17 BULL case. ~600 chars. What needs to happen for the bull target — specific operating triggers, FY28 revenue / EBITDA / margin numbers.' },
  base_content: { kind: 'card', max: 650, instruction: 'Slide 17 BASE case. ~600 chars. The most likely operating path; consistent with the published target.' },
  bear_content: { kind: 'card', max: 650, instruction: 'Slide 17 BEAR case. ~600 chars. The downside scenario — specific failure modes, not generic hedge language.' },

  // ── Entry / Review / Exit (slide 19) ────────────────────────────────────
  entry_strategy: { kind: 'bullets', max: 1600, instruction: 'Slide 19 ENTRY. 5-8 bullet points, one per line. Start each bullet with `• ` (Unicode bullet + space). No paragraphs, no headings. Each bullet ~120-200 chars. Cover price range to accumulate, position sizing principle, build cadence, what would make you wait.' },
  review_strategy: { kind: 'bullets', max: 1600, instruction: 'Slide 19 REVIEW. 5-8 bullet points, one per line. Start each bullet with `• ` (Unicode bullet + space). No paragraphs, no headings. Each bullet ~120-200 chars. Specific checkpoints (quarterly results, capex milestones, regulatory triggers, mix evolution) and what would shift the rating.' },
  exit_strategy: { kind: 'bullets', max: 1600, instruction: 'Slide 19 EXIT. 5-8 bullet points, one per line. Start each bullet with `• ` (Unicode bullet + space). No paragraphs, no headings. Each bullet ~120-200 chars. Concrete exit triggers tied to the bear case, stop loss, and partial profit booking levels.' },
};

export interface PptCopyMetadata {
  cmp?: string | number | null;
  target?: string | number | null;
  upsidePct?: string | number | null;
  marketCap?: string | number | null;
  marketCapCategory?: string | null;
  rating?: string | null;
  saarthiScore?: string | number | null;
}

/** Strip ```json fences / leading prose from a Claude response. */
export function extractJsonObject(raw: string): string {
  const fenceMatch = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenceMatch) return fenceMatch[1].trim();
  const firstBrace = raw.indexOf('{');
  const lastBrace = raw.lastIndexOf('}');
  if (firstBrace >= 0 && lastBrace > firstBrace) return raw.slice(firstBrace, lastBrace + 1);
  return raw.trim();
}

/**
 * Clip every value to its schema max, drop unknown keys, coerce to string.
 * Preserves paragraph breaks (`\n\n`) so multi-paragraph panels survive into
 * the PPTX — single line breaks and runs of whitespace within a paragraph
 * are collapsed to single spaces.
 */
export function sanitisePptContent(raw: unknown): Record<string, string> {
  const out: Record<string, string> = {};
  if (!raw || typeof raw !== 'object') return out;
  const obj = raw as Record<string, unknown>;
  for (const [key, spec] of Object.entries(PPT_COPY_SCHEMA)) {
    const v = obj[key];
    if (v == null) continue;
    let s = String(v);
    // Normalise newlines and whitespace while keeping paragraph breaks intact.
    s = s.replace(/\r\n?/g, '\n')                // CRLF → LF
      .replace(/[ \t]+/g, ' ')                // collapse runs of spaces
      .replace(/[ \t]*\n[ \t]*/g, '\n')       // trim around newlines
      .trim();
    if (spec.kind === 'bullets') {
      // Strip markdown emphasis but keep the leading bullet marker.
      s = s.replace(/\*\*|__/g, '');
      // If the LLM jammed multiple bullets onto a single line ("• A. • B. • C."
      // or "• A • B • C"), split on the bullet marker so we recover newline
      // separation. Without this, the entire string renders as one paragraph
      // on the slide.
      if (!s.includes('\n')) {
        const parts = s.split(/(?=[•●▪◦])\s*/).map((p) => p.trim()).filter(Boolean);
        if (parts.length > 1) {
          s = parts.join('\n');
        }
      }
      // Preserve single newlines (one per bullet); cap consecutive blanks.
      s = s.replace(/\n{2,}/g, '\n');
      // Normalise each line: ensure it starts with "• " and drop stray markers.
      s = s.split('\n').map((line) => {
        const trimmed = line.trim().replace(/^[\-\*•●▪◦·#>\s]+/, '').trim();
        if (!trimmed) return '';
        return `• ${trimmed}`;
      }).filter(Boolean).join('\n');
    } else {
      s = s.replace(/\n{3,}/g, '\n\n');
      // Strip markdown emphasis the renderer doesn't honour — EXCEPT for the
      // one field where inline bold is intentionally rendered (slide 1 thesis).
      if (key !== 'investment_thesis_s1') {
        s = s.replace(/\*\*|__/g, '');
      }
      s = s.replace(/^[#>\s]+/, '');
    }
    if (s.length > spec.max) {
      const trimmed = s.slice(0, spec.max);
      const lastPara = trimmed.lastIndexOf('\n\n');
      // Find the last complete sentence boundary (. or ? or !)
      const lastSentence = Math.max(
        trimmed.lastIndexOf('. '),
        trimmed.lastIndexOf('? '),
        trimmed.lastIndexOf('! ')
      );
      const lastSpace = trimmed.lastIndexOf(' ');
      const cut = lastPara > spec.max * 0.6
        ? lastPara
        : (lastSentence > spec.max * 0.7
          ? lastSentence + 1
          : (lastSpace > spec.max * 0.75 ? lastSpace : spec.max));
      s = trimmed.slice(0, cut).trim();
    }
    out[key] = s;
  }
  return out;
}

export function buildPptCopyPrompt(
  companyName: string,
  nseSymbol: string,
  sectorName: string,
  metadata: PptCopyMetadata,
  sections: Array<{ key: string; title: string; content: string }>,
): { system: string; user: string } {
  const schemaLines = Object.entries(PPT_COPY_SCHEMA)
    .map(([key, spec]) => `  "${key}": "<${spec.kind}, ≤${spec.max} chars — ${spec.instruction}>"`)
    .join(',\n');

  const sectionsBlock = sections
    .filter((s) => s.content && s.content.length > 20)
    .map((s) => `### ${s.title} (key: ${s.key})\n${s.content.slice(0, 8000)}`)
    .join('\n\n---\n\n');

  const fmt = (v: unknown) => (v == null || v === '' ? 'n/a' : String(v));

  const system = `You are a presentation copywriter for an institutional equity research deck.

Your job: take the approved long-form research report and rewrite it as box-specific PPT copy, one field per template placeholder. You are NOT generating new facts — you are reshaping existing approved content into the right length and tone for each slide.

Rules (strict):
- Output ONE JSON object. No prose before or after. No markdown fences (no \`\`\`).
- Every value is plain text. NO pipe tables. NO headers ("#").
- Inline bold: ONLY the field investment_thesis_s1 may contain **double-asterisk** emphasis (4-8 short phrases, 2-6 words each). Every OTHER field must be plain prose with no markdown bold/italics.
- Bullet markers: NO bullet markers in non-bullet fields. Bullet fields (industry_structure, key_industry_tailwinds, key_industry_risks, entry_strategy, review_strategy, exit_strategy) MUST be FORMATTED as a JSON-string with LITERAL "\\n" newline separators between every bullet — NOT a continuous paragraph. Each bullet line starts with "• " (Unicode U+2022 + space). Produce 5-8 bullets per field. CORRECT example value: "• First point about market size and structure.\\n• Second point about named competitors.\\n• Third point on regulation." WRONG (renders as one paragraph): "• First point. • Second point. • Third point." The "\\n" between bullets is REQUIRED — without it, the slide will display every bullet jammed into a single paragraph and the field will fail QC.
- Respect each field's char budget and ALL of its instruction. Cards must be single, finished sentences — no mid-clause cut-offs ("focusing on.").
- Paragraphs: these FIVE fields MUST contain 2-3 paragraphs separated by literal "\\n\\n" (JSON-escaped newline + newline): investment_thesis_s4, investment_thesis_detailed, forecast_assumptions, financial_commentary, valuation_commentary. Real paragraph breaks, not single newlines. All other non-bullet fields stay as one paragraph.
- When a slide has multiple cards (investment_ideas_1..4, business_model_1..6, management_content_1..8, indicators_1..6, key_catalyst_1..3, saarthi_*), each card MUST cover a DIFFERENT facet. Never repeat the same fact across cards on the same slide. Never start two cards with the same opening phrase.
- The thesis panels (investment_thesis_s1, investment_thesis_s4, investment_thesis_detailed) must each have a DIFFERENT shape — same facts allowed, but different opening, different emphasis, and the lengths must clearly differ (s1 < s4 < detailed).
- KPIs (KPI_heading_1..6 + KPI_1..6): pick SIX company-specific operating or financial metrics that are NOT already shown on slide 1. Slide-1 chips are CMP, Target, Market Cap, Cap Category, SAARTHI Score, NSE code — DO NOT repeat any of these as KPIs. Good KPI examples: Revenue (FY26), EBITDA Margin, PAT Growth, ROE, ROCE, Capacity, Plants, Countries, Debt/Equity, FCF Yield, etc. Heading = 1-3 words; Value = the actual number with units (e.g. "₹4,265 Cr", "11.8%", "70+", "0.3x").
- Numbers: keep ₹ symbol, keep crore/lakh units as written, keep FY labels (FY26A / FY28E). Don't fabricate figures — if a number is not in the source material, omit the claim rather than guessing.
- Length discipline: every field has a char budget. The template's text boxes are sized for those budgets — under-filling leaves visually empty boxes on the slide. Aim for 80-100% of each budget. If a field's source section looks thin, draw from RELATED sections that cover the same topic (e.g. for company_overview pull from company_background + business_model + industry_overview; for forecast_assumptions pull from investment_rationale + scenario_analysis) rather than coming in short. Only fall below 70% of budget if the source material truly has nothing more to say on that topic.
- Padding rule: extra length must come from MORE specific facts, numbers, or named details — never from filler adjectives, restatements, or generic industry observations.`;

  const user = `Company: ${companyName} (NSE: ${nseSymbol}) | Sector: ${sectorName}

## Snapshot (slide-1 chips — DO NOT reuse these as KPIs)
- CMP: ${fmt(metadata.cmp)}
- Target: ${fmt(metadata.target)}
- Upside: ${fmt(metadata.upsidePct)}
- Market Cap: ${fmt(metadata.marketCap)}
- Cap Category: ${fmt(metadata.marketCapCategory)}
- Rating: ${fmt(metadata.rating)}
- SAARTHI Score: ${fmt(metadata.saarthiScore)}

## Approved Stage 2 sections (source material)

${sectionsBlock}

---

Return a JSON object with EXACTLY these keys (no extras, no omissions):

{
${schemaLines}
}`;

  return { system, user };
}
