// Pipeline types for the research pipeline
// Maps to the ACTUAL research_sessions table in Supabase
// v2: Anthropic SDK + Web Search architecture (no vector embeddings)

// ========================
// State Machine
// ========================

export type PipelineStatus =
  | 'company_selected'           // Company chosen, user decides financial model or vault
  | 'financial_model_generating' // Financial model being generated
  | 'vault_creating'             // Creating Google Drive vault + fetching docs
  | 'vault_ready'                // Vault created, documents listed, user confirms
  | 'stage0_generating'          // Sector framework generation (Anthropic + web search)
  | 'stage0_review'
  | 'stage0_approved'
  | 'stage1_generating'          // Thesis generation (Anthropic + web search)
  | 'stage1_review'
  | 'stage1_approved'
  | 'stage2_generating'          // Report generation (Anthropic + web search)
  | 'stage2_review'
  | 'stage2_approved'
  | 'published';

// Valid state transitions (relaxed to allow jumping around for already-generated content)
export const PIPELINE_TRANSITIONS: Record<PipelineStatus, PipelineStatus[]> = {
  company_selected: ['vault_creating'],
  financial_model_generating: ['vault_ready'],       // Model done → back to vault_ready
  vault_creating: ['vault_ready'],
  vault_ready: ['financial_model_generating', 'stage0_generating', 'stage0_review', 'stage0_approved', 'stage1_generating', 'stage1_review', 'stage1_approved', 'stage2_generating', 'stage2_review', 'stage2_approved', 'published'],
  stage0_generating: ['stage0_review'],
  stage0_review: ['stage0_approved', 'stage0_generating'],
  stage0_approved: ['stage1_generating', 'stage1_review', 'stage1_approved', 'stage2_generating', 'stage2_review', 'stage2_approved', 'published'],
  stage1_generating: ['stage1_review'],
  stage1_review: ['stage1_approved', 'stage1_generating'],
  stage1_approved: ['stage2_generating', 'stage2_review', 'stage2_approved', 'published'],
  stage2_generating: ['stage2_review'],
  stage2_review: ['stage2_approved', 'stage2_generating'],
  stage2_approved: ['published', 'stage0_generating', 'stage1_generating', 'stage2_generating'],
  published: ['stage2_approved', 'stage0_generating', 'stage1_generating', 'stage2_generating'],
};

export function canTransition(from: PipelineStatus, to: PipelineStatus): boolean {
  // Relaxing transitions for better UX when resuming sessions
  if (to === 'vault_ready' || to === 'company_selected') return true;
  return PIPELINE_TRANSITIONS[from]?.includes(to) ?? true; // Default to true if not strictly mapped, to prevent UI blocks
}

// ========================
// Pipeline Session — matches actual research_sessions table
// ========================

export interface PipelineSession {
  id: string;
  session_id: string;
  company_name: string;
  company_nse_code: string;
  sector: string | null;
  sub_sector: string | null;
  current_state: string;
  // Stage 0 — native table columns
  sector_playbook_original: Record<string, unknown> | null;
  sector_playbook_approved: Record<string, unknown> | null;
  stage0_analyst_notes: string | null;
  // Stage 1 — native table columns
  condensed_briefing: string | null;
  thesis_original: Record<string, unknown> | null;
  thesis_approved: Record<string, unknown> | null;
  stage1_analyst_notes: string | null;
  coherence_changelog: Record<string, unknown> | null;
  // Stage 2 — native table columns
  final_report_raw: string | null;
  final_report_approved: string | null;
  // Vault columns
  vault_folder_id: string | null;
  vault_folder_url: string | null;
  financial_model_file_url: string | null;
  financial_model_json_url: string | null;
  // Pipeline columns
  pipeline_status: PipelineStatus | null;
  sector_framework: SectorFramework | null;
  thesis_condensed: string | null;
  thesis_output: string | null;
  report_content: string | null;
  selected_model: string | null;
  total_tokens_used: number;
  generation_time_seconds: number;
  // Metadata
  created_by: string | null;
  created_at: string;
  updated_at: string;
}

// ========================
// Sector Playbook — matches sector_playbooks table
// ========================

export interface SectorPlaybook {
  id: string;
  sector_name: string;
  sector_slug: string;
  sector_description: string;
  market_size: Record<string, unknown>;
  value_chain: Record<string, unknown>;
  industry_structure: Record<string, unknown>;
  regulatory_framework: Record<string, unknown>;
  business_model_archetypes: Record<string, unknown>[];
  cycle_position: string;
  cycle_description: string;
  sector_sentiment: string;
  consensus_view: string;
  macro_factors: Record<string, unknown>[];
  recent_developments: Record<string, unknown>[];
  contrarian_angles: Record<string, unknown>[];
  ai_writing_instructions: Record<string, unknown>;
  key_metrics_to_track: string[];
  valuation_rules: Record<string, unknown>;
  red_flags: string[];
  green_flags: string[];
  version: number;
  last_updated: string;
  created_at: string;
  updated_at: string;
  created_by: string | null;
}

// ========================
// Sector Playbook Version — historical snapshot
// ========================

export interface SectorPlaybookVersion {
  id: string;
  playbook_id: string;
  sector_name: string;
  version: number;
  framework_content: string;
  created_by: string | null;
  created_at: string;
}

// ========================
// Sector Framework — the AI-generated sector analysis markdown
// Stored in sector_playbooks and copied to research_sessions.sector_framework
// ========================

export interface SectorFramework {
  sector_name: string;
  /** The full AI-generated markdown content */
  markdown: string;
  /** Playbook version (increments on regeneration) */
  version: number;
  /** ISO date string of last update */
  last_updated: string;
}

export interface SectorKnowledge {
  id: string;
  sector_id: string;
  category: string;
  title: string;
  content: string;
  source: string | null;
  sort_order: number;
  created_at: string;
  updated_at: string;
}

// ========================
// Stage 1: Thesis Generation
// ========================

export type SaarthiRating = 'STRONG BUY' | 'BUY' | 'ACCUMULATE' | 'HOLD' | 'UNDERPERFORM' | 'SELL';

export interface ThesisOutput {
  investment_thesis: string;
  bull_case: string;
  bear_case: string;
  key_catalysts: string[];
  key_risks: string[];
  target_price_rationale: string;
  saarthi_score: number;
  saarthi_rating: SaarthiRating;
}

// ========================
// Stage 2: Report Generation
// ========================

export interface ReportOutput {
  sections: ReportSectionOutput[];
  executive_summary: string;
  recommendation: 'BUY' | 'SELL' | 'HOLD' | 'NEUTRAL';
  target_price: string | null;
}

export interface ReportSectionOutput {
  key: string;
  title: string;
  heading: string;
  content: string;
}

// ========================
// Research Sections (stored in research_sections table)
// ========================

export interface ResearchSection {
  id: string;
  session_id: string;
  section_key: string;
  section_title: string;
  stage: 'stage0' | 'stage1' | 'stage2';
  content: string;
  heading: string | null;
  sort_order: number;
  tokens_used: number;
  created_at: string;
  updated_at: string;
}

// ========================
// SKB Suggested Updates
// ========================

export interface SkbSuggestedUpdate {
  id: string;
  session_id: string;
  sector_id: string;
  category: string;
  title: string;
  suggested_content: string;
  status: 'pending' | 'approved' | 'rejected';
  created_at: string;
}

// ========================
// Pipeline UI State
// ========================

export interface PipelineProgress {
  stage: 'stage0' | 'stage1' | 'stage2';
  step: string;
  message: string;
  percent: number;
}

// Stage labels for UI
export const PIPELINE_STAGE_LABELS: Record<PipelineStatus, string> = {
  company_selected: 'Company Selected',
  financial_model_generating: 'Generating Financial Model',
  vault_creating: 'Creating Drive Vault',
  vault_ready: 'Vault Ready',
  stage0_generating: 'Generating Sector Framework',
  stage0_review: 'Review Sector Framework',
  stage0_approved: 'Sector Framework Approved',
  stage1_generating: 'Generating Investment Thesis',
  stage1_review: 'Review Investment Thesis',
  stage1_approved: 'Investment Thesis Approved',
  stage2_generating: 'Generating Report',
  stage2_review: 'Review Report',
  stage2_approved: 'Report Approved',
  published: 'Published',
};

// Stage step numbers for progress bar
export function getStageNumber(status: PipelineStatus): number {
  if (status === 'company_selected' || status === 'financial_model_generating') return 0;
  if (status === 'vault_creating' || status === 'vault_ready') return 1;
  if (status.startsWith('stage0')) return 2;
  if (status.startsWith('stage1')) return 3;
  if (status.startsWith('stage2')) return 4;
  if (status === 'published') return 5;
  return 0;
}

// AI Model options — now primarily Anthropic, with OpenRouter fallback
export const PIPELINE_MODELS = [
  { id: 'claude-sonnet', label: 'Claude Sonnet 4' },
  { id: 'claude-opus', label: 'Claude Opus 4' },
] as const;

export const DEFAULT_PIPELINE_MODEL = 'claude-sonnet';
