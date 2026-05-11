// Pipeline API layer — CRUD for pipeline sessions, sector playbooks, research sections
// Maps to the ACTUAL research_sessions table schema in Supabase

import { supabase } from '@/lib/supabase';
import type {
  PipelineSession,
  PipelineStatus,
  SectorKnowledge,
  SectorPlaybook,
  SectorPlaybookVersion,
  ResearchSection,
  SkbSuggestedUpdate,
} from '@/types/pipeline';
import { canTransition } from '@/types/pipeline';

// ========================
// Pipeline Session CRUD
// ========================

/**
 * Creates a new pipeline session using the actual table columns.
 */
export async function createPipelineSession(input: {
  company_name: string;
  company_nse_code: string;
  sector: string;
  created_by: string;
  selected_model?: string;
}): Promise<PipelineSession> {
  const sessionId = crypto.randomUUID();

  const { data, error } = await supabase
    .from('research_sessions')
    .insert({
      session_id: sessionId,
      company_name: input.company_name,
      company_nse_code: input.company_nse_code,
      sector: input.sector || null,
      current_state: 'company_selected',
      pipeline_status: 'company_selected',
      selected_model: input.selected_model || null,
      created_by: input.created_by,
      total_tokens_used: 0,
      generation_time_seconds: 0,
    })
    .select()
    .single();

  if (error) throw new Error(`Failed to create pipeline session: ${error.message}`);
  return data;
}

/**
 * Gets a pipeline session by session_id
 */
export async function getPipelineSession(sessionId: string): Promise<PipelineSession | null> {
  const { data, error } = await supabase
    .from('research_sessions')
    .select('*')
    .eq('session_id', sessionId)
    .maybeSingle();

  if (error) throw new Error(`Failed to fetch pipeline session: ${error.message}`);
  return data;
}

/**
 * Lists pipeline sessions with optional filters.
 */
export async function listPipelineSessions(options?: {
  createdBy?: string;
  pipelineStatus?: PipelineStatus;
  page?: number;
  pageSize?: number;
}): Promise<{ data: PipelineSession[]; count: number }> {
  const page = options?.page ?? 0;
  const pageSize = options?.pageSize ?? 25;
  const from = page * pageSize;
  const to = from + pageSize - 1;

  let query = supabase
    .from('research_sessions')
    .select('*', { count: 'exact' })
    .not('pipeline_status', 'is', null);

  if (options?.createdBy) {
    query = query.eq('created_by', options.createdBy);
  }
  if (options?.pipelineStatus) {
    query = query.eq('pipeline_status', options.pipelineStatus);
  }

  query = query.order('created_at', { ascending: false }).range(from, to);

  const { data, error, count } = await query;

  if (error) {
    console.warn('[Pipeline] Listing sessions fallback:', error.message);
    let fallbackQuery = supabase
      .from('research_sessions')
      .select('*', { count: 'exact' });

    if (options?.createdBy) {
      fallbackQuery = fallbackQuery.eq('created_by', options.createdBy);
    }

    fallbackQuery = fallbackQuery.order('created_at', { ascending: false }).range(from, to);

    const { data: fbData, error: fbError, count: fbCount } = await fallbackQuery;
    if (fbError) throw new Error(`Failed to list sessions: ${fbError.message}`);
    return { data: fbData ?? [], count: fbCount ?? 0 };
  }

  return { data: data ?? [], count: count ?? 0 };
}

/**
 * Transitions a pipeline session to a new status (validates state machine).
 */
export async function transitionPipelineStatus(
  sessionId: string,
  newStatus: PipelineStatus,
  currentStatus?: PipelineStatus
): Promise<PipelineSession> {
  if (currentStatus && !canTransition(currentStatus, newStatus)) {
    throw new Error(`Invalid transition: ${currentStatus} → ${newStatus}`);
  }

  const { data, error } = await supabase
    .from('research_sessions')
    .update({
      pipeline_status: newStatus,
      current_state: newStatus,
      updated_at: new Date().toISOString(),
    })
    .eq('session_id', sessionId)
    .select()
    .single();

  if (error) throw new Error(`Failed to transition status: ${error.message}`);
  return data;
}

/**
 * Updates pipeline session with stage output data.
 */
export async function updatePipelineOutput(
  sessionId: string,
  updates: Partial<Pick<
    PipelineSession,
    'sector_framework' | 'thesis_condensed' | 'thesis_output' | 'report_content' |
    'total_tokens_used' | 'generation_time_seconds' | 'selected_model' |
    'sector_playbook_original' | 'sector_playbook_approved' |
    'condensed_briefing' | 'thesis_original' | 'thesis_approved' |
    'final_report_raw' | 'final_report_approved' |
    'vault_folder_id' | 'vault_folder_url' | 'financial_model_file_url' | 'financial_model_json_url'
  >>
): Promise<PipelineSession> {
  const { data, error } = await supabase
    .from('research_sessions')
    .update({
      ...updates,
      updated_at: new Date().toISOString(),
    })
    .eq('session_id', sessionId)
    .select()
    .single();

  if (error) throw new Error(`Failed to update pipeline output: ${error.message}`);
  return data;
}

/**
 * Deletes a pipeline session and its related records
 */
export async function deletePipelineSession(sessionId: string): Promise<void> {
  await supabase.from('research_sections').delete().eq('session_id', sessionId).then(() => {});
  await supabase.from('skb_suggested_updates').delete().eq('session_id', sessionId).then(() => {});
  const { error } = await supabase
    .from('research_sessions')
    .delete()
    .eq('session_id', sessionId);

  if (error) throw new Error(`Failed to delete pipeline session: ${error.message}`);
}

// ========================
// Sector Playbook CRUD
// ========================

/**
 * Gets a sector playbook by sector name.
 * Returns null if no playbook exists for this sector.
 */
export async function getSectorPlaybook(sectorName: string): Promise<SectorPlaybook | null> {
  const { data, error } = await supabase
    .from('sector_playbooks')
    .select('*')
    .ilike('sector_name', sectorName)
    .order('version', { ascending: false })
    .limit(1)
    .maybeSingle();

  if (error) {
    console.warn('[Pipeline] Could not load sector playbook:', error.message);
    return null;
  }
  return data;
}

/**
 * Creates a new sector playbook (for a sector that doesn't have one yet).
 */
export async function createSectorPlaybook(input: {
  sector_name: string;
  sector_description: string;
  framework_content: string; // The full AI-generated markdown framework
  created_by: string;
}): Promise<SectorPlaybook> {
  const slug = input.sector_name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');

  const { data, error } = await supabase
    .from('sector_playbooks')
    .insert({
      sector_name: input.sector_name,
      sector_slug: slug,
      sector_description: input.sector_description,
      // Store the full framework as ai_writing_instructions for retrieval
      ai_writing_instructions: { framework_markdown: input.framework_content },
      version: 1,
      last_updated: new Date().toISOString().split('T')[0],
      created_by: input.created_by,
    })
    .select()
    .single();

  if (error) throw new Error(`Failed to create sector playbook: ${error.message}`);
  return data;
}

/**
 * Updates an existing sector playbook with new/updated framework content.
 * Increments the version number.
 */
export async function updateSectorPlaybook(
  playbookId: string,
  updates: {
    framework_content?: string;
    sector_description?: string;
  }
): Promise<SectorPlaybook> {
  // ── Archive the current version before overwriting ──
  try {
    const { data: current } = await supabase
      .from('sector_playbooks')
      .select('*')
      .eq('id', playbookId)
      .single();

    if (current && updates.framework_content) {
      const oldMarkdown = getFrameworkFromPlaybook(current);
      if (oldMarkdown && oldMarkdown.length > 100) {
        await supabase.from('sector_playbook_versions').insert({
          playbook_id: playbookId,
          sector_name: current.sector_name,
          version: current.version || 1,
          framework_content: oldMarkdown,
          created_by: current.created_by || null,
        });
      }
    }
  } catch (archiveErr) {
    // Non-fatal — don't block the update if archiving fails
    console.warn('[Pipeline] Could not archive playbook version:', archiveErr);
  }

  const updatePayload: Record<string, unknown> = {
    updated_at: new Date().toISOString(),
    last_updated: new Date().toISOString().split('T')[0],
  };

  if (updates.framework_content) {
    updatePayload.ai_writing_instructions = { framework_markdown: updates.framework_content };
  }
  if (updates.sector_description) {
    updatePayload.sector_description = updates.sector_description;
  }

  // Increment version
  const { data: currentForVersion } = await supabase
    .from('sector_playbooks')
    .select('version')
    .eq('id', playbookId)
    .single();

  if (currentForVersion) {
    updatePayload.version = (currentForVersion.version || 1) + 1;
  }

  const { data, error } = await supabase
    .from('sector_playbooks')
    .update(updatePayload)
    .eq('id', playbookId)
    .select()
    .single();

  if (error) throw new Error(`Failed to update sector playbook: ${error.message}`);
  return data;
}

/**
 * Fetches all archived versions for a sector playbook, newest first.
 */
export async function getSectorPlaybookVersions(
  playbookId: string
): Promise<SectorPlaybookVersion[]> {
  const { data, error } = await supabase
    .from('sector_playbook_versions')
    .select('*')
    .eq('playbook_id', playbookId)
    .order('version', { ascending: false });

  if (error) {
    console.warn('[Pipeline] Could not load playbook versions:', error.message);
    return [];
  }
  return data ?? [];
}

/**
 * Extracts the framework markdown content from a sector playbook.
 */
export function getFrameworkFromPlaybook(playbook: SectorPlaybook): string {
  const instructions = playbook.ai_writing_instructions as Record<string, unknown> | null;
  if (instructions?.framework_markdown && typeof instructions.framework_markdown === 'string') {
    return instructions.framework_markdown;
  }
  // Fallback: build a summary from the structured fields
  const parts: string[] = [];
  if (playbook.sector_description) parts.push(`## Sector Overview\n${playbook.sector_description}`);
  if (playbook.key_metrics_to_track?.length) {
    parts.push(`## Key Metrics\n${playbook.key_metrics_to_track.map(m => `- ${m}`).join('\n')}`);
  }
  if (playbook.red_flags?.length) {
    parts.push(`## Red Flags\n${playbook.red_flags.map(f => `- ${f}`).join('\n')}`);
  }
  if (playbook.green_flags?.length) {
    parts.push(`## Green Flags\n${playbook.green_flags.map(f => `- ${f}`).join('\n')}`);
  }
  return parts.join('\n\n') || 'No sector framework content available.';
}

// ========================
// Sector Knowledge CRUD (uses sector_id FK)
// ========================

/**
 * Lists all sectors from the sectors table
 */
export async function listSectors(): Promise<{ id: string; sector_name: string; description: string }[]> {
  const { data, error } = await supabase
    .from('sectors')
    .select('id, sector_name, description')
    .order('sector_name');

  if (error) {
    console.warn('[Pipeline] Could not load sectors table:', error.message);
    return [];
  }
  return data ?? [];
}

/**
 * Gets sector knowledge entries by sector_id
 */
export async function getSectorKnowledge(sectorId: string): Promise<SectorKnowledge[]> {
  const { data, error } = await supabase
    .from('sector_knowledge')
    .select('*')
    .eq('sector_id', sectorId)
    .order('category')
    .order('sort_order');

  if (error) {
    console.warn('[Pipeline] Could not load sector knowledge:', error.message);
    return [];
  }
  return data ?? [];
}

/**
 * Gets sector ID by name from sectors table
 */
export async function getSectorIdByName(sectorName: string): Promise<string | null> {
  const { data, error } = await supabase
    .from('sectors')
    .select('id')
    .ilike('sector_name', sectorName)
    .limit(1)
    .maybeSingle();

  if (error || !data) return null;
  return data.id;
}

// ========================
// Research Sections CRUD
// ========================

/**
 * Saves a research section (stage output)
 */
export async function saveResearchSection(input: {
  session_id: string;
  section_key: string;
  section_title: string;
  stage: 'stage0' | 'stage1' | 'stage2';
  content: string;
  heading?: string;
  sort_order?: number;
  tokens_used?: number;
}): Promise<ResearchSection> {
  const { data, error } = await supabase
    .from('research_sections')
    .insert({
      session_id: input.session_id,
      section_key: input.section_key,
      section_title: input.section_title,
      stage: input.stage,
      content: input.content,
      heading: input.heading || null,
      sort_order: input.sort_order || 0,
      tokens_used: input.tokens_used || 0,
    })
    .select()
    .single();

  if (error) throw new Error(`Failed to save research section: ${error.message}`);
  return data;
}

/**
 * Gets all research sections for a session, optionally filtered by stage
 */
export async function getResearchSections(
  sessionId: string,
  stage?: 'stage0' | 'stage1' | 'stage2'
): Promise<ResearchSection[]> {
  let query = supabase
    .from('research_sections')
    .select('*')
    .eq('session_id', sessionId)
    .order('sort_order')
    .order('created_at');

  if (stage) {
    query = query.eq('stage', stage);
  }

  const { data, error } = await query;
  if (error) {
    console.warn('[Pipeline] Could not load research sections:', error.message);
    return [];
  }
  return data ?? [];
}

/**
 * Updates a research section's content
 */
export async function updateResearchSection(
  sectionId: string,
  updates: { content?: string; heading?: string }
): Promise<ResearchSection> {
  const { data, error } = await supabase
    .from('research_sections')
    .update({
      ...updates,
      updated_at: new Date().toISOString(),
    })
    .eq('id', sectionId)
    .select()
    .single();

  if (error) throw new Error(`Failed to update research section: ${error.message}`);
  return data;
}

/**
 * Deletes all research sections for a session and stage (for regeneration)
 */
export async function clearResearchSections(
  sessionId: string,
  stage: 'stage0' | 'stage1' | 'stage2'
): Promise<void> {
  const { error } = await supabase
    .from('research_sections')
    .delete()
    .eq('session_id', sessionId)
    .eq('stage', stage);

  if (error) console.warn('[Pipeline] Could not clear research sections:', error.message);
}

// ========================
// SKB Suggested Updates
// ========================

/**
 * Creates a suggested update to the sector knowledge base
 */
export async function createSkbSuggestion(input: {
  session_id: string;
  sector_id: string;
  category: string;
  title: string;
  suggested_content: string;
}): Promise<SkbSuggestedUpdate> {
  const { data, error } = await supabase
    .from('skb_suggested_updates')
    .insert(input)
    .select()
    .single();

  if (error) throw new Error(`Failed to create SKB suggestion: ${error.message}`);
  return data;
}

/**
 * Lists pending SKB suggestions
 */
export async function listSkbSuggestions(options?: {
  sectorId?: string;
  status?: 'pending' | 'approved' | 'rejected';
}): Promise<SkbSuggestedUpdate[]> {
  let query = supabase
    .from('skb_suggested_updates')
    .select('*')
    .order('created_at', { ascending: false });

  if (options?.sectorId) {
    query = query.eq('sector_id', options.sectorId);
  }
  if (options?.status) {
    query = query.eq('status', options.status);
  }

  const { data, error } = await query;
  if (error) {
    console.warn('[Pipeline] Could not load SKB suggestions:', error.message);
    return [];
  }
  return data ?? [];
}

/**
 * Approves or rejects a SKB suggestion
 */
export async function reviewSkbSuggestion(
  suggestionId: string,
  status: 'approved' | 'rejected',
  reviewerEmail: string
): Promise<void> {
  const { error } = await supabase
    .from('skb_suggested_updates')
    .update({
      status,
      reviewed_by: reviewerEmail,
      reviewed_at: new Date().toISOString(),
    })
    .eq('id', suggestionId);

  if (error) throw new Error(`Failed to review SKB suggestion: ${error.message}`);
}

// ========================
// Pipeline Prompts (pipeline_prompts table)
// ========================

export type PipelineStage = 'stage0' | 'stage1' | 'stage2';

export interface PipelinePrompt {
  id: string;
  stage: PipelineStage;
  label: string;
  system_prompt: string;
  user_prompt: string;
  is_default: boolean;
  user_email: string | null;
  created_at: string;
  updated_at: string;
}

/**
 * Get the effective prompt for a stage: user's custom prompt if it exists, otherwise null.
 * The caller falls back to DEFAULT_PROMPTS from anthropic-pipeline.ts when null.
 */
export async function getPipelinePrompt(
  stage: PipelineStage,
  userEmail: string,
): Promise<PipelinePrompt | null> {
  const { data, error } = await supabase
    .from('pipeline_prompts')
    .select('*')
    .eq('stage', stage)
    .eq('user_email', userEmail)
    .maybeSingle();

  if (error) {
    console.warn('[Pipeline] Could not load pipeline prompt:', error.message);
    return null;
  }
  return data;
}

/**
 * Save (upsert) a user's custom prompt for a stage.
 */
export async function savePipelinePrompt(input: {
  stage: PipelineStage;
  label: string;
  system_prompt: string;
  user_prompt: string;
  user_email: string;
}): Promise<PipelinePrompt> {
  // Check if user already has a custom prompt for this stage
  const existing = await getPipelinePrompt(input.stage, input.user_email);

  if (existing) {
    const { data, error } = await supabase
      .from('pipeline_prompts')
      .update({
        label: input.label,
        system_prompt: input.system_prompt,
        user_prompt: input.user_prompt,
        updated_at: new Date().toISOString(),
      })
      .eq('id', existing.id)
      .select()
      .single();

    if (error) throw new Error(`Failed to update pipeline prompt: ${error.message}`);
    return data;
  }

  const { data, error } = await supabase
    .from('pipeline_prompts')
    .insert({
      stage: input.stage,
      label: input.label,
      system_prompt: input.system_prompt,
      user_prompt: input.user_prompt,
      is_default: false,
      user_email: input.user_email,
    })
    .select()
    .single();

  if (error) throw new Error(`Failed to save pipeline prompt: ${error.message}`);
  return data;
}

/**
 * Delete a user's custom prompt for a stage (reverts to defaults).
 */
export async function deletePipelinePrompt(stage: PipelineStage, userEmail: string): Promise<void> {
  const { error } = await supabase
    .from('pipeline_prompts')
    .delete()
    .eq('stage', stage)
    .eq('user_email', userEmail)
    .eq('is_default', false);

  if (error) throw new Error(`Failed to delete pipeline prompt: ${error.message}`);
}
