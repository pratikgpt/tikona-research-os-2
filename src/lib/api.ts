// API utilities for external services
import type { VaultResponse, VaultDocument, DriveFile } from '@/types/vault';
import { normalizeDriveFile as normalizeFile } from '@/types/vault';
import { supabase, getCurrentUserEmail } from '@/lib/supabase';
import type {
  ResearchSession,
  ResearchReport,
  TextSectionKey,
  CreateResearchSessionInput,
  CreateSessionDocumentInput,
  SessionDocument,
  PromptTemplate,
} from '@/types/database';

const N8N_BASE_URL = 'https://n8n.tikonacapital.com';

// Two URL strategies for the financial model service:
//
// 1. PROXY  (/proxy/fm) — safe for SHORT requests (<30 s).
//    Used for: /generate-async (enqueue, ~1-2 s) and /status/<id> polls (~1 s each).
//    NOT safe for the blocking /generate endpoint which takes ~10 minutes.
//
// 2. DIRECT (VPS IP) — no Vercel proxy timeout, but requires:
//    a) Python service sends CORS headers (Access-Control-Allow-Origin: *)
//    b) Site served over HTTP (not HTTPS) OR VPS has a TLS cert.
//    Used only as a fallback when the async endpoint is unavailable.
//
// Bottom line: prefer async mode (proxy path). The direct URL is a last-resort
// fallback that only works if the VPS is HTTP-accessible from the browser.
const FM_PROXY_URL  = '/proxy/fm';   // Vercel rewrite → VPS; fine for short calls
const FM_DIRECT_URL = import.meta.env.VITE_FINANCIAL_MODEL_URL || 'http://72.61.226.16:8500'; // direct, no proxy timeout

/**
 * Triggers financial model generation on the VPS Python service.
 *
 * WHY DIRECT URL (not /proxy/fm):
 * Vercel rewrites have a hard ~30 s response-body timeout. The Python service
 * takes ~10 minutes to generate an Excel model. Going through the Vercel proxy
 * guarantees a 502 / ROUTER_EXTERNAL_TARGET_ERROR every time and wastes credits.
 * Calling the VPS directly from the browser avoids this limit entirely.
 *
 * ASYNC POLLING:
 * Instead of holding one HTTP connection open for 10 minutes (which browsers
 * may also kill), we POST to /generate-async to enqueue the job and receive a
 * job_id immediately. We then poll /status/<job_id> every 15 s until the job
 * completes or fails (max 20 min). This is resilient to tab refreshes because
 * the job keeps running on the VPS regardless.
 *
 * Falls back to a synchronous call if the Python service does not support the
 * async endpoints (older deployments).
 */
export async function generateFinancialModel(
  ticker: string,
  companyName: string,
  sector: string,
  folderId: string
): Promise<{
  fileId: string | null;
  fileUrl: string | null;
  fileName: string;
  storageUrl: string | null;
  driveFileUrl: string | null;
}> {
  const requestBody = {
    nse_symbol: ticker.toUpperCase(),
    company_name: companyName,
    sector: sector,
    folder_id: folderId,
  };

  const fileName = `${ticker.toUpperCase()}_model.xlsx`;
  console.log('[FM] Starting generation for:', requestBody);

  // ── Try async job endpoint first ────────────────────────────────────────────
  let useAsync = false;
  let jobId: string | null = null;

  try {
    const asyncResp = await fetch(`${FM_PROXY_URL}/generate-async`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody),
      signal: AbortSignal.timeout(15_000), // 15 s to enqueue
    });

    if (asyncResp.ok) {
      const asyncData = await asyncResp.json() as Record<string, unknown>;
      if (asyncData.job_id) {
        jobId = asyncData.job_id as string;
        useAsync = true;
        console.log('[FM] Async job enqueued, job_id:', jobId);
      }
    }
  } catch {
    // Async endpoint not available — fall back to synchronous call below
    console.warn('[FM] /generate-async not available, falling back to synchronous /generate');
  }

  // ── Async polling path ───────────────────────────────────────────────────────
  if (useAsync && jobId) {
    const POLL_INTERVAL_MS = 15_000; // 15 s between polls
    const MAX_WAIT_MS = 20 * 60 * 1000; // 20 min max
    const startedAt = Date.now();
    let consecutive404s = 0;

    while (Date.now() - startedAt < MAX_WAIT_MS) {
      await new Promise(r => setTimeout(r, POLL_INTERVAL_MS));

      const statusResp = await fetch(`${FM_PROXY_URL}/job/${jobId}`, {
        signal: AbortSignal.timeout(10_000),
      });

      if (statusResp.status === 404) {
        consecutive404s++;
        console.warn(`[FM] /job/${jobId} returned 404 (${consecutive404s}x) — job endpoint may not exist on this VPS build`);
        if (consecutive404s >= 2) {
          // Job status endpoint doesn't exist on this Python service deployment.
          // Break out and fall through to the synchronous direct-VPS call below.
          console.warn('[FM] Job endpoint unavailable — falling back to synchronous /generate via direct VPS URL');
          break;
        }
        continue;
      }

      consecutive404s = 0; // reset on any non-404

      if (!statusResp.ok) {
        console.warn('[FM] Job poll returned', statusResp.status, '— retrying');
        continue;
      }

      const statusData = await statusResp.json() as Record<string, unknown>;
      const state = (statusData.status || statusData.state) as string | undefined;
      console.log(`[FM] Job ${jobId} status/state: ${state}`);

      if (state === 'SUCCESS' || state === 'success' || state === 'completed') {
        return _extractFmResult(statusData, fileName);
      }

      if (state === 'FAILURE' || state === 'error' || state === 'failed') {
        throw new Error((statusData.message as string) || 'Financial model generation failed on VPS');
      }
      // state is 'PENDING' / 'RUNNING' / 'processing' — keep polling
    }

    // Only throw timeout if we exhausted the full wait window (not a 404 break)
    if (Date.now() - startedAt >= MAX_WAIT_MS) {
      throw new Error('Financial model timed out after 20 minutes. Check the VPS service logs.');
    }
    // else: fell through from 404 break — continue to sync fallback below
  }

  // ── Synchronous fallback path (no /generate-async on this VPS build) ────────
  console.log('[FM] Using synchronous /generate endpoint (no timeout guard — ensure VPS is reachable)');

  const response = await fetch(`${FM_DIRECT_URL}/generate`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody),
    // No AbortSignal — we let the browser wait as long as it takes.
    // FM_DIRECT_URL bypasses Vercel proxy so this will NOT get a 502.
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('[FM] Error response:', errorText);
    throw new Error(`Financial model generation failed: ${response.status} ${response.statusText}`);
  }

  const responseText = await response.text();
  if (!responseText) throw new Error('Financial model service returned an empty response');

  let data: Record<string, unknown>;
  try {
    data = JSON.parse(responseText);
  } catch {
    console.error('[FM] Non-JSON response:', responseText.slice(0, 200));
    throw new Error('Financial model service returned invalid JSON');
  }

  if (data.status === 'error') {
    throw new Error((data.message as string) || 'Financial model generation failed');
  }

  return _extractFmResult(data, fileName);
}

/** Normalises a raw Python service response into the standard FM result shape. */
function _extractFmResult(
  data: Record<string, unknown>,
  fileName: string
): {
  fileId: string | null;
  fileUrl: string | null;
  fileName: string;
  storageUrl: string | null;
  driveFileUrl: string | null;
} {
  const fileId = (data.file_id as string) || (data.id as string) || null;
  const storageUrl = (data.storage_url as string) || (data.supabase_url as string) || null;
  const driveFileUrl =
    (data.file_url as string) ||
    (data.webViewLink as string) ||
    (fileId ? `https://drive.google.com/file/d/${fileId}/view` : null);
  const fileUrl = storageUrl || driveFileUrl || null;
  return { fileId, fileUrl, fileName, storageUrl, driveFileUrl };
}

export async function mirrorFinancialModelToStorage(
  ticker: string
): Promise<{ fileUrl: string; filePath: string | null; jsonFileUrl: string | null; jsonFilePath: string | null }> {
  // mirrorFinancialModelToStorage: /storage/<ticker> is a fast call (<5 s),
  // so the proxy is fine here. No timeout issues.
  const response = await fetch(`${FM_PROXY_URL}/storage/${ticker.toUpperCase()}`, {
    method: 'POST',
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Financial model storage mirror failed: ${response.status} ${errorText.slice(0, 300)}`);
  }

  const data = (await response.json()) as {
    status: string;
    message?: string | null;
    storage_url?: string | null;
    storage_path?: string | null;
    json_storage_url?: string | null;
    json_storage_path?: string | null;
  };

  if (data.status !== 'success' || !data.storage_url) {
    throw new Error(data.message || 'Financial model storage mirror did not return a URL');
  }

  return {
    fileUrl: data.storage_url,
    filePath: data.storage_path ?? null,
    jsonFileUrl: data.json_storage_url ?? null,
    jsonFilePath: data.json_storage_path ?? null,
  };
}

/**
 * Creates a research vault (Google Drive folder) for the given stock ticker
 * @param ticker - NSE stock symbol (e.g., "TATAMOTORS")
 * @param sector - Company sector (e.g., "Automobile")
 * @returns Promise with folder link, folder ID, and files array
 */
export async function createVault(ticker: string, sector: string): Promise<VaultResponse> {
  const requestBody = {
    nse_symbol: ticker.toUpperCase(),
    sector: sector,
  };

  console.log('[API] Creating vault with:', requestBody);

  const response = await fetch(`${N8N_BASE_URL}/webhook/create-folder`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(requestBody),
  });

  console.log('[API] Response status:', response.status);
  console.log('[API] Response headers:', Object.fromEntries(response.headers.entries()));

  if (!response.ok) {
    const errorText = await response.text();
    console.error('[API] Error response:', errorText);
    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  }

  // First get the text to see what we're actually receiving
  const responseText = await response.text();
  console.log('[API] Response text:', responseText);
  console.log('[API] Response text length:', responseText.length);

  // Try to parse as JSON
  let data;
  try {
    data = responseText ? JSON.parse(responseText) : null;
    console.log('[API] Parsed data:', JSON.stringify(data, null, 2));
  } catch (parseError) {
    console.error('[API] JSON parse error:', parseError);
    console.error('[API] Raw response text:', responseText);
    throw new Error(`Invalid JSON response from server: ${parseError instanceof Error ? parseError.message : 'Unknown error'}`);
  }

  if (!data) {
    throw new Error('Empty response from server. The n8n workflow may still be processing. Please check your n8n workflow configuration.');
  }

  // Handle different response formats from n8n
  let normalizedResponse: VaultResponse;

  // Case 0: n8n workflow started asynchronously (not waiting for completion)
  if (data.message === 'Workflow was started' || data.message?.includes('started')) {
    console.error('[API] n8n workflow is configured to respond immediately, not waiting for completion!');
    console.error('[API] Configure your webhook node with: Respond = "When Last Node Finishes"');
    throw new Error(
      'n8n workflow is not configured correctly. The webhook must wait for the workflow to complete before responding. ' +
      'Set your Webhook node\'s "Respond" setting to "When Last Node Finishes" and add a "Respond to Webhook" node at the end.'
    );
  }

  // Case 1: n8n returns nested array [[file1, file2]] (double-nested format)
  if (Array.isArray(data) && data.length > 0 && Array.isArray(data[0])) {
    console.log('[API] ✓ Case 1: Received nested array [[files]]');
    const files = data[0]; // Extract the inner array
    console.log('[API] Flattened to', files.length, 'files');
    const firstFile = files[0];
    console.log('[API] First file:', firstFile);
    const folderId = firstFile?.parents?.[0] || 'unknown';
    console.log('[API] Extracted folder ID:', folderId);

    normalizedResponse = {
      status: 'success',
      folder_link: `https://drive.google.com/drive/folders/${folderId}`,
      folder_id: folderId,
      files: files,
    };
    console.log('[API] Created normalized response with folder_link:', normalizedResponse.folder_link);
  }
  // Case 2: n8n returns array of files directly [file1, file2] (legacy format)
  else if (Array.isArray(data)) {
    console.log('[API] ✓ Case 2: Received flat array [files]');
    const firstFile = data[0];
    console.log('[API] First file:', firstFile);
    const folderId = firstFile?.parents?.[0] || 'unknown';
    console.log('[API] Extracted folder ID:', folderId);

    normalizedResponse = {
      status: 'success',
      folder_link: `https://drive.google.com/drive/folders/${folderId}`,
      folder_id: folderId,
      files: data,
    };
    console.log('[API] Created normalized response with folder_link:', normalizedResponse.folder_link);
  }
  // Case 3: n8n returns proper structure with status, folder_link, files
  else if (data.status === 'success' || data.folder_link) {
    console.log('[API] ✓ Case 3: Received object with status/folder_link');
    console.log('[API] data.status:', data.status);
    console.log('[API] data.folder_link:', data.folder_link);
    console.log('[API] data.folder_id:', data.folder_id);
    console.log('[API] data.files length:', data.files?.length);

    // Extract folder_id (from data or from first file's parents)
    const folderId = data.folder_id || data.main_folder_id || data.files?.[0]?.parents?.[0] || 'unknown';

    // Generate folder_link if not provided
    const folderLink = data.folder_link || `https://drive.google.com/drive/folders/${folderId}`;

    normalizedResponse = {
      status: data.status || 'success',
      folder_link: folderLink,
      folder_id: folderId,
      files: data.files || [],
    };
    console.log('[API] Created normalized response with folder_link:', normalizedResponse.folder_link);
  }
  // Case 4: Error response
  else if (data.status === 'error') {
    throw new Error(data.message || 'Failed to create vault');
  }
  // Case 5: Unknown format
  else {
    console.error('[API] Unknown response format:', data);
    throw new Error('Unexpected response format from server');
  }

  // Validate normalized response
  if (!normalizedResponse.folder_link) {
    console.error('[API] ❌ Validation failed: folder_link is missing or empty');
    console.error('[API] normalizedResponse:', normalizedResponse);
    throw new Error('Response missing folder_link. Please check the browser console for details.');
  }

  if (!normalizedResponse.folder_id) {
    console.warn('[API] Missing folder_id, using fallback from first file');
  }

  if (!normalizedResponse.files || !Array.isArray(normalizedResponse.files)) {
    console.warn('[API] No files array in response');
    normalizedResponse.files = [];
  }

  console.log('[API] Normalized response:', {
    status: normalizedResponse.status,
    folder_id: normalizedResponse.folder_id,
    filesCount: normalizedResponse.files.length,
  });

  return normalizedResponse;
}

/**
 * Processes the vault response from n8n and normalizes the files
 * @param response - Raw response from n8n webhook
 * @returns Normalized data with folder info and documents
 */
export function processVaultResponse(response: VaultResponse): {
  folderId: string;
  folderUrl: string;
  documents: VaultDocument[];
} {
  // Filter out invalid files (normalizeDriveFile returns null for invalid entries)
  const documents = (response.files || [])
    .map(normalizeFile)
    .filter((doc): doc is VaultDocument => doc !== null);

  if (documents.length < (response.files || []).length) {
    console.warn(`[API] ${(response.files || []).length - documents.length} files were invalid and filtered out`);
  }

  return {
    folderId: response.folder_id,
    folderUrl: response.folder_link,
    documents,
  };
}

/**
 * Deletes a file from Google Drive
 * @param fileId - Google Drive file ID
 * @returns Promise<void>
 */
export async function deleteDocument(fileId: string): Promise<void> {
  console.log('[API] Deleting document:', fileId);

  const response = await fetch(`${N8N_BASE_URL}/webhook/delete-file`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ file_id: fileId }),
  });

  console.log('[API] Delete response status:', response.status);

  // Read body ONCE before checking response.ok
  const responseText = await response.text();
  console.log('[API] Delete response:', responseText);

  if (!response.ok) {
    console.error('[API] Delete error:', responseText);
    throw new Error(`Failed to delete file: ${response.statusText}`);
  }

  // Parse response if any
  if (responseText) {
    try {
      const data = JSON.parse(responseText);
      if (data.status === 'error') {
        throw new Error(data.message || 'Failed to delete file');
      }
    } catch {
      // If response is not JSON or parsing fails, that's okay
      console.log('[API] Delete completed (non-JSON response)');
    }
  }
}

// ========================
// Document Upload
// ========================

/**
 * Uploads a document to Google Drive via n8n webhook
 * @param folderId - Google Drive folder ID to upload into
 * @param fileName - Desired file name
 * @param fileBase64 - Base64-encoded file content
 * @returns The uploaded file as a VaultDocument
 */
export async function uploadDocument(
  folderId: string,
  fileName: string,
  fileBase64: string
): Promise<VaultDocument> {
  console.log('[API] Uploading document:', { folderId, fileName });

  const response = await fetch(`${N8N_BASE_URL}/webhook/upload-document`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      folder_id: folderId,
      file_name: fileName,
      file_base64: fileBase64,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('[API] Upload error:', errorText);
    throw new Error(`Failed to upload document: ${response.statusText}`);
  }

  const responseText = await response.text();
  if (!responseText) {
    throw new Error('Empty response from upload webhook');
  }

  const data = JSON.parse(responseText);

  if (data.status === 'error') {
    throw new Error(data.message || 'Upload failed');
  }

  // The webhook returns the file object in data.file or data directly
  const driveFile: DriveFile = data.file || data;
  const normalizedDoc = normalizeFile(driveFile);

  if (!normalizedDoc) {
    throw new Error('Upload succeeded but returned invalid file data (missing ID)');
  }

  return normalizedDoc;
}

// ========================
// Research Session CRUD
// ========================

/**
 * Creates a new research session in the database
 */
export async function saveResearchSession(
  input: CreateResearchSessionInput
): Promise<ResearchSession> {
  const { data, error } = await supabase
    .from('research_sessions')
    .insert({
      session_id: crypto.randomUUID(),
      company_name: input.company_name,
      company_nse_code: input.nse_symbol,
      sector: input.sector || null,
      current_state: input.status || 'document_review',
      created_by: input.user_email,
      total_tokens_used: 0,
      generation_time_seconds: 0,
    })
    .select()
    .single();

  if (error) {
    throw new Error(`Failed to save research session: ${error.message}`);
  }

  return data;
}

/**
 * Lists research sessions with optional filters
 */
export async function listResearchSessions(options?: {
  userEmail?: string;
  status?: ResearchSession['status'];
  page?: number;
  pageSize?: number;
}): Promise<{ data: ResearchSession[]; count: number }> {
  const page = options?.page ?? 0;
  const pageSize = options?.pageSize ?? 25;
  const from = page * pageSize;
  const to = from + pageSize - 1;

  let query = supabase
    .from('research_sessions')
    .select('*', { count: 'exact' });

  if (options?.userEmail) {
    query = query.eq('created_by', options.userEmail);
  }
  if (options?.status) {
    query = query.eq('current_state', options.status);
  }

  query = query.order('created_at', { ascending: false }).range(from, to);

  const { data, error, count } = await query;

  if (error) {
    throw new Error(`Failed to list research sessions: ${error.message}`);
  }

  return { data: data ?? [], count: count ?? 0 };
}

/**
 * Gets a single research session by ID
 */
export async function getResearchSession(
  sessionId: string
): Promise<ResearchSession | null> {
  const { data, error } = await supabase
    .from('research_sessions')
    .select('*')
    .eq('session_id', sessionId)
    .maybeSingle();

  if (error) {
    throw new Error(`Failed to fetch research session: ${error.message}`);
  }

  return data;
}

/**
 * Updates the status of a research session
 */
export async function updateSessionStatus(
  sessionId: string,
  status: ResearchSession['status']
): Promise<ResearchSession> {
  const { data, error } = await supabase
    .from('research_sessions')
    .update({ current_state: status, updated_at: new Date().toISOString() })
    .eq('session_id', sessionId)
    .select()
    .single();

  if (error) {
    throw new Error(`Failed to update session status: ${error.message}`);
  }

  return data;
}

/**
 * Updates the selected_for_ai flag on session_documents for a research session.
 * Marks documents in selectedDocumentIds as selected, rest as unselected.
 */
export async function updateSessionDocuments(
  sessionId: string,
  selectedDocumentIds: string[]
): Promise<ResearchSession> {
  // Unselect all documents for this session first
  await supabase
    .from('session_documents')
    .update({ selected_for_ai: false })
    .eq('session_id', sessionId);

  // Select the chosen documents
  if (selectedDocumentIds.length > 0) {
    await supabase
      .from('session_documents')
      .update({ selected_for_ai: true })
      .eq('session_id', sessionId)
      .in('document_id', selectedDocumentIds);
  }

  // Return the updated session
  const { data, error } = await supabase
    .from('research_sessions')
    .select('*')
    .eq('session_id', sessionId)
    .single();

  if (error) {
    throw new Error(`Failed to update session documents: ${error.message}`);
  }

  return data;
}

/**
 * Deletes a research session and its associated documents
 */
export async function deleteResearchSession(
  sessionId: string
): Promise<void> {
  // Delete session documents first (FK constraint)
  const { error: docError } = await supabase
    .from('session_documents')
    .delete()
    .eq('session_id', sessionId);

  if (docError) {
    throw new Error(`Failed to delete session documents: ${docError.message}`);
  }

  // Delete the session
  const { error } = await supabase
    .from('research_sessions')
    .delete()
    .eq('session_id', sessionId);

  if (error) {
    throw new Error(`Failed to delete research session: ${error.message}`);
  }
}

// ========================
// Session Documents CRUD
// ========================

/**
 * Saves documents for a research session
 * Maps drive_file_id → document_id and file_name → document_name
 * for backward compatibility with existing table columns
 */
export async function saveSessionDocuments(
  documents: CreateSessionDocumentInput[]
): Promise<SessionDocument[]> {
  if (documents.length === 0) return [];

  // Map to include both old (document_id, document_name) and new (drive_file_id, file_name) columns
  const rows = documents.map((doc) => ({
    ...doc,
    document_id: doc.drive_file_id,
    document_name: doc.file_name,
  }));

  const { data, error } = await supabase
    .from('session_documents')
    .insert(rows)
    .select();

  if (error) {
    throw new Error(`Failed to save session documents: ${error.message}`);
  }

  return data ?? [];
}

/**
 * Gets documents for a research session
 */
export async function getSessionDocuments(
  sessionId: string
): Promise<SessionDocument[]> {
  const { data, error } = await supabase
    .from('session_documents')
    .select('*')
    .eq('session_id', sessionId)
    .order('created_at', { ascending: true });

  if (error) {
    throw new Error(`Failed to fetch session documents: ${error.message}`);
  }

  return data ?? [];
}

// ========================
// Research Reports CRUD
// ========================

/**
 * Creates a new research report record
 */
export async function createResearchReport(input: {
  session_id: string;
  user_email: string;
  company_name: string;
  nse_symbol: string;
}): Promise<ResearchReport> {
  const { data, error } = await supabase
    .from('research_reports')
    .insert({
      session_id: input.session_id,
      user_email: input.user_email,
      company_name: input.company_name,
      nse_symbol: input.nse_symbol,
      status: 'generating',
    })
    .select()
    .single();

  if (error) {
    throw new Error(`Failed to create report: ${error.message}`);
  }

  return data;
}

/**
 * Updates a report section after generation
 */
export async function updateReportSection(
  reportId: string,
  sectionKey: TextSectionKey,
  content: string
): Promise<void> {
  const { error } = await supabase
    .from('research_reports')
    .update({
      [sectionKey]: content,
      updated_at: new Date().toISOString(),
    })
    .eq('report_id', reportId);

  if (error) {
    throw new Error(`Failed to update report section: ${error.message}`);
  }
}

// ========================
// Dynamic Section Columns (cs_ prefixed)
// ========================

/**
 * Creates a new column in research_reports for a custom section.
 * Column will be named cs_{sectionKey} (e.g., cs_valuation_analysis).
 */
export async function addReportSectionColumn(sectionKey: string): Promise<void> {
  // Create content column (cs_<key>)
  const { error } = await supabase.rpc('add_report_section_column', {
    col_name: sectionKey,
  });

  if (error) {
    throw new Error(`Failed to create section column: ${error.message}`);
  }

  // Create heading column (cs_<key>_h)
  const { error: hError } = await supabase.rpc('add_report_section_column', {
    col_name: `${sectionKey}_h`,
  });

  if (hError) {
    throw new Error(`Failed to create heading column cs_${sectionKey}_h: ${hError.message}`);
  }
}

/**
 * Drops a custom section column from research_reports.
 * Only cs_ prefixed columns can be dropped (safety guard in RPC).
 */
export async function dropReportSectionColumn(sectionKey: string): Promise<void> {
  const { error } = await supabase.rpc('drop_report_section_column', {
    col_name: sectionKey,
  });

  if (error) {
    // Column may not exist if it was never created — that's OK
    console.warn(`[API] Failed to drop section column cs_${sectionKey}:`, error.message);
  }
}

/**
 * Updates a custom section column (cs_ prefixed) in research_reports.
 */
export async function updateCustomSection(
  reportId: string,
  sectionKey: string,
  content: string
): Promise<void> {
  const colName = `cs_${sectionKey}`;
  const { error } = await supabase
    .from('research_reports')
    .update({
      [colName]: content,
      updated_at: new Date().toISOString(),
    })
    .eq('report_id', reportId);

  if (error) {
    throw new Error(`Failed to update custom section cs_${sectionKey}: ${error.message}`);
  }
}

/**
 * Updates the generated heading for a section in research_reports.
 * Default sections use `<key>_h`, custom sections use `cs_<key>_h`.
 */
export async function updateSectionHeading(
  reportId: string,
  sectionKey: string,
  heading: string,
  isCustom: boolean
): Promise<void> {
  const colName = isCustom ? `cs_${sectionKey}_h` : `${sectionKey}_h`;
  const { error } = await supabase
    .from('research_reports')
    .update({
      [colName]: heading,
      updated_at: new Date().toISOString(),
    })
    .eq('report_id', reportId);

  if (error) {
    throw new Error(`Failed to update section heading ${colName}: ${error.message}`);
  }
}

/**
 * Finalizes a report with metadata
 */
export async function finalizeReport(
  reportId: string,
  tokensUsed: number,
  generationTimeSeconds: number
): Promise<void> {
  const { error } = await supabase
    .from('research_reports')
    .update({
      status: 'draft',
      tokens_used: tokensUsed,
      generation_time_seconds: generationTimeSeconds,
      updated_at: new Date().toISOString(),
    })
    .eq('report_id', reportId);

  if (error) {
    throw new Error(`Failed to finalize report: ${error.message}`);
  }
}

/**
 * Gets a report by session ID
 */
export async function getReportBySession(
  sessionId: string
): Promise<ResearchReport | null> {
  const { data: generatedReport, error: generatedError } = await supabase
    .from('research_reports')
    .select('*')
    .eq('session_id', sessionId)
    .or('pptx_file_url.not.is.null,pptx_file_path.not.is.null')
    .order('updated_at', { ascending: false })
    .limit(1)
    .maybeSingle();

  if (generatedError) {
    throw new Error(`Failed to fetch report: ${generatedError.message}`);
  }

  if (generatedReport) {
    return generatedReport;
  }

  const { data, error } = await supabase
    .from('research_reports')
    .select('*')
    .eq('session_id', sessionId)
    .order('updated_at', { ascending: false })
    .limit(1)
    .maybeSingle();

  if (error) {
    throw new Error(`Failed to fetch report: ${error.message}`);
  }

  return data;
}

/**
 * Gets a single report by report ID.
 */
export async function getReportById(
  reportId: string
): Promise<ResearchReport | null> {
  const { data, error } = await supabase
    .from('research_reports')
    .select('*')
    .eq('report_id', reportId)
    .maybeSingle();

  if (error) {
    throw new Error(`Failed to fetch report: ${error.message}`);
  }

  return data;
}

// ============================================================
// Prompt Template Management
// ============================================================

/**
 * List all prompt templates (default + user's custom templates)
 */
export async function listPromptTemplates(userEmail?: string): Promise<PromptTemplate[]> {
  const userEmailValue = userEmail || (await getCurrentUserEmail());

  let query = supabase
    .from('prompt_templates')
    .select('*')
    .order('sort_order', { ascending: true })
    .order('section_key', { ascending: true });

  // Get default templates OR user's custom templates
  if (userEmailValue) {
    query = query.or(`is_default.eq.true,user_email.eq.${userEmailValue}`);
  } else {
    query = query.eq('is_default', true);
  }

  const { data, error } = await query;

  if (error) {
    throw new Error(`Failed to list prompt templates: ${error.message}`);
  }

  return (data || []) as unknown as PromptTemplate[];
}

/**
 * Create a new custom prompt template
 */
export async function createPromptTemplate(input: {
  section_key: string;
  title: string;
  heading_prompt?: string;
  prompt_text: string;
  search_keywords: string[];
}) {
  const userEmail = await getCurrentUserEmail();

  const { data, error } = await supabase
    .from('prompt_templates')
    .insert({
      ...input,
      user_email: userEmail,
      is_default: false,
    })
    .select()
    .single();

  if (error) {
    throw new Error(`Failed to create prompt template: ${error.message}`);
  }

  return data;
}

/**
 * Update an existing prompt template
 */
export async function updatePromptTemplate(
  id: string,
  updates: {
    title?: string;
    heading_prompt?: string;
    prompt_text?: string;
    search_keywords?: string[];
    section_key?: string;
  }
) {
  const { data, error } = await supabase
    .from('prompt_templates')
    .update({
      ...updates,
      updated_at: new Date().toISOString(),
    })
    .eq('id', id)
    .select()
    .single();

  if (error) {
    throw new Error(`Failed to update prompt template: ${error.message}`);
  }

  return data;
}

/**
 * Delete a prompt template
 */
export async function deletePromptTemplate(id: string) {
  const { error } = await supabase
    .from('prompt_templates')
    .delete()
    .eq('id', id);

  if (error) {
    throw new Error(`Failed to delete prompt template: ${error.message}`);
  }
}

/**
 * Batch-update sort_order for multiple prompt templates
 */
export async function reorderPromptTemplates(
  updates: { id: string; sort_order: number }[]
): Promise<void> {
  // Supabase doesn't support batch update by ID natively, so update one by one
  for (const { id, sort_order } of updates) {
    const { error } = await supabase
      .from('prompt_templates')
      .update({ sort_order })
      .eq('id', id);

    if (error) {
      throw new Error(`Failed to reorder prompt template: ${error.message}`);
    }
  }
}

// ========================
// Report Publishing
// ========================

/**
 * Publishes a report so it's visible to customers
 */
export async function publishReport(reportId: string, plan?: string): Promise<void> {
  const updatePayload: { is_published: boolean; published_at: string; updated_at: string; plan?: string } = {
    is_published: true,
    published_at: new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };
  
  if (plan) {
    updatePayload.plan = plan;
  }

  const { data, error } = await supabase
    .from('research_reports')
    .update(updatePayload)
    .eq('report_id', reportId)
    .select();

  if (error) {
    throw new Error(`Failed to publish report: ${error.message}`);
  }

  if (!data || data.length === 0) {
    throw new Error('Failed to publish report. The database update was ignored. Please check Row Level Security (RLS) policies.');
  }
}

/**
 * Unpublishes a report (hides from customers)
 */
export async function unpublishReport(reportId: string): Promise<void> {
  const { data, error } = await supabase
    .from('research_reports')
    .update({
      is_published: false,
      published_at: null,
      updated_at: new Date().toISOString(),
    })
    .eq('report_id', reportId)
    .select();

  if (error) {
    throw new Error(`Failed to unpublish report: ${error.message}`);
  }

  if (!data || data.length === 0) {
    throw new Error('Failed to unpublish report. The database update was ignored. Please check Row Level Security (RLS) policies.');
  }
}

// ============================================================
// PPTX Report (reportgen pipeline)
// ============================================================

const configuredPptServiceUrl =
  (import.meta as unknown as { env?: { VITE_PPT_SERVICE_URL?: string } }).env?.VITE_PPT_SERVICE_URL?.trim();

export const PPT_SERVICE_URL =
  !configuredPptServiceUrl || configuredPptServiceUrl.includes('localhost:8501')
    ? '/proxy/ppt'
    : configuredPptServiceUrl;

export interface GeneratePptxInput {
  reportId: string;
  sessionId: string;
  useMock?: boolean;
}

export interface GeneratePptxResult {
  status: 'success' | 'error';
  message?: string | null;
  pptx_file_url?: string | null;
  pptx_file_path?: string | null;
  pptx_pdf_file_url?: string | null;
  pptx_pdf_file_path?: string | null;
  ppt_file_id?: string | null;
  ppt_file_url?: string | null;
  duration_seconds?: number | null;
  warnings?: string[] | null;
}

export interface PptPlaceholdersResult {
  status: string;
  placeholders: Record<string, string>;
  has_saved_overrides: boolean;
  warnings: string[];
}

export async function fetchPptPlaceholders(
  reportId: string,
  sessionId: string,
  ignoreOverrides?: boolean,
): Promise<PptPlaceholdersResult> {
  const response = await fetch(`${PPT_SERVICE_URL}/preview-placeholders`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ reportId, sessionId, ignoreOverrides: !!ignoreOverrides }),
  });
  if (!response.ok) {
    const err = await response.text().catch(() => '');
    throw new Error(`Failed to fetch PPT placeholders: ${response.status} ${err.slice(0, 200)}`);
  }
  return (await response.json()) as PptPlaceholdersResult;
}

export async function savePptPlaceholders(
  reportId: string,
  data: Record<string, string>,
): Promise<void> {
  // Ensure cs_ppt_data column exists (idempotent — Postgres IF NOT EXISTS)
  try {
    await supabase.rpc('add_report_section_column', { col_name: 'ppt_data' });
  } catch {
    // Column already exists — safe to ignore
  }
  const { error } = await supabase
    .from('research_reports')
    .update({ cs_ppt_data: JSON.stringify(data), updated_at: new Date().toISOString() })
    .eq('report_id', reportId);
  if (error) throw new Error(`Failed to save PPT data: ${error.message}`);
}

export async function generatePptx(
  input: GeneratePptxInput,
): Promise<GeneratePptxResult> {
  const response = await fetch(`${PPT_SERVICE_URL}/generate-pptx`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      reportId: input.reportId,
      sessionId: input.sessionId,
      useMock: input.useMock ?? false,
    }),
  });

  if (!response.ok) {
    let detail = '';
    const cloned = response.clone();
    try {
      const body = (await response.json()) as { message?: string };
      detail = body?.message ?? '';
    } catch {
      detail = (await cloned.text()).slice(0, 300);
    }
    throw new Error(`PPTX generation failed: ${response.status} ${detail}`);
  }

  return (await response.json()) as GeneratePptxResult;
}

export interface SyncSlidesInput {
  reportId: string;
  pptFileId: string;
}

export interface SyncSlidesResult {
  status: string;
  message: string;
  pptx_pdf_file_url: string;
  pptx_pdf_file_path: string;
}

export async function syncSlidesToPdf(
  input: SyncSlidesInput,
): Promise<SyncSlidesResult> {
  const response = await fetch(`${PPT_SERVICE_URL}/sync-slides-pdf`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      reportId: input.reportId,
      pptFileId: input.pptFileId,
    }),
  });

  if (!response.ok) {
    let detail = '';
    const cloned = response.clone();
    try {
      const body = (await response.json()) as { message?: string };
      detail = body?.message ?? '';
    } catch {
      detail = (await cloned.text()).slice(0, 300);
    }
    throw new Error(`Sync slides to PDF failed: ${response.status} ${detail}`);
  }

  return (await response.json()) as SyncSlidesResult;
}
