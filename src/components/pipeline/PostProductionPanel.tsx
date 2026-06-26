import { useState, useEffect, useCallback, useRef } from 'react';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select';
import { cn } from '@/lib/utils';
import { supabase } from '@/lib/supabase';
import {
  getReportBySession,
  createResearchReport,
  publishReport,
  generatePptx,
  syncSlidesToPdf,
  PPT_SERVICE_URL,
  updateCustomSection,
} from '@/lib/api';
import { runPptCopywriting } from '@/lib/anthropic-pipeline';
import { savePptContent, getPptContent } from '@/lib/pipeline-api';
import { createRecommendation } from '@/lib/recommendations-api';
import type { ResearchReport } from '@/types/database';
import type { RecommendationRating } from '@/types/recommendations';
import {
  Check,
  Loader2,
  Mic,
  Video,
  ExternalLink,
  Play,
  Download,
  ChevronDown,
  ChevronUp,
  Shield,
  Send,
  Presentation,
  Wifi,
  WifiOff,
  RefreshCw,
  CheckCircle,
  AlertCircle,
} from 'lucide-react';
import PPTDataPanel from './PPTDataPanel';

const N8N_BASE = 'https://n8n.tikonacapital.com/webhook';

interface PostProductionPanelProps {
  sessionId: string;
  companyName: string;
  nseSymbol: string;
  sector?: string | null;
  vaultId: string | null;
  financialModelFileUrl?: string | null;
  userEmail: string;
  stage2Sections: Array<{ id?: string; key: string; title: string; content: string }>;
  initialReport?: ResearchReport | null;
  onPublished: () => void;
}

export default function PostProductionPanel({
  sessionId,
  companyName,
  nseSymbol,
  sector,
  financialModelFileUrl,
  userEmail,
  stage2Sections,
  initialReport = null,
  onPublished,
}: PostProductionPanelProps) {
  // --- Report ---
  const [reportId, setReportId] = useState<string | null>(null);

  // --- PPTX generation ---
  const [pptxGenerating, setPptxGenerating] = useState(false);
  const [pptxElapsedSeconds, setPptxElapsedSeconds] = useState(0);
  const [pptxFileUrl, setPptxFileUrl] = useState<string | null>(null);
  const [pptxPdfFileUrl, setPptxPdfFileUrl] = useState<string | null>(null);
  const [pptFileId, setPptFileId] = useState<string | null>(null);
  const [pptFileUrl, setPptFileUrl] = useState<string | null>(null);
  const [slidesSyncing, setSlidesSyncing] = useState(false);
  const [useMock, setUseMock] = useState(false);
  const [previewTab, setPreviewTab] = useState<'pdf' | 'slides'>('pdf');
  const pptxTimerRef = useRef<ReturnType<typeof setInterval> | null>(null);

  // --- PPT copywriting pass (slide-specific copy) ---
  const [slideCopyReady, setSlideCopyReady] = useState(false);
  const [slideCopyGenerating, setSlideCopyGenerating] = useState(false);
  const [slideCopyProgress, setSlideCopyProgress] = useState<{ message: string; percent: number } | null>(null);

  // --- Service health ---
  const [serviceHealth, setServiceHealth] = useState<'checking' | 'ok' | 'down'>('checking');

  // --- Podcast ---
  const [scriptGenerating, setScriptGenerating] = useState(false);
  const [podcastScript, setPodcastScript] = useState<string | null>(null);
  const [lastSavedScript, setLastSavedScript] = useState<string | null>(null);
  const [scriptSaving, setScriptSaving] = useState(false);
  const [audioGenerating, setAudioGenerating] = useState(false);
  const [audioFileUrl, setAudioFileUrl] = useState<string | null>(null);

  // --- Video ---
  const [videoGenerating, setVideoGenerating] = useState(false);
  const [videoFileUrl, setVideoFileUrl] = useState<string | null>(null);
  const [videoElapsedSeconds, setVideoElapsedSeconds] = useState(0);

  // --- PPT Data ---
  const [pptDataConfirmed, setPptDataConfirmed] = useState(false);
  const [pptDataOpen, setPptDataOpen] = useState(true);

  // --- UI ---
  const [scriptExpanded, setScriptExpanded] = useState(false);
  const [selectedPlan, setSelectedPlan] = useState<string>('');

  // --- Publish & Telegram ---
  const [isPublished, setIsPublished] = useState(false);
  const [telegramSending, setTelegramSending] = useState(false);
  const [telegramSent, setTelegramSent] = useState(false);
  const [sendPush, setSendPush] = useState(true);

  const reportIdRef = useRef(reportId);
  reportIdRef.current = reportId;

  useEffect(() => {
    return () => {
      if (pptxTimerRef.current) clearInterval(pptxTimerRef.current);
    };
  }, []);

  // Health check
  useEffect(() => {
    let cancelled = false;
    setServiceHealth('checking');
    fetch(`${PPT_SERVICE_URL}/health`, { signal: AbortSignal.timeout(5000) })
      .then((r) => {
        if (!cancelled) setServiceHealth(r.ok ? 'ok' : 'down');
      })
      .catch(() => {
        if (!cancelled) setServiceHealth('down');
      });
    return () => { cancelled = true; };
  }, []);

  // Prevent browser reload/close from aborting active client-side generation processes
  useEffect(() => {
    if (!slideCopyGenerating && !pptxGenerating) return;

    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      const msg = 'Generation is in progress. Reloading or leaving the page will abort the operation. Are you sure?';
      e.preventDefault();
      e.returnValue = msg;
      return msg;
    };

    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
    };
  }, [slideCopyGenerating, pptxGenerating]);

  // Restore from existing report
  // Detect whether the PPT copywriting pass has already run for this session.
  useEffect(() => {
    let cancelled = false;
    getPptContent(sessionId)
      .then((c) => { if (!cancelled) setSlideCopyReady(!!c && Object.keys(c).length > 0); })
      .catch(() => { if (!cancelled) setSlideCopyReady(false); });
    return () => { cancelled = true; };
  }, [sessionId]);

  useEffect(() => {
    if (initialReport) {
      restoreFromReport(initialReport);
      return;
    }
    getReportBySession(sessionId).then(async (report) => {
      if (report) {
        restoreFromReport(report);
      } else {
        const created = await createResearchReport({
          session_id: sessionId,
          user_email: userEmail,
          company_name: companyName,
          nse_symbol: nseSymbol,
        });
        restoreFromReport(created);
      }
    }).catch((error) => {
      console.error('[PostProduction] Failed to load/create report', { sessionId, error });
    });
  }, [initialReport, sessionId, userEmail, companyName, nseSymbol]);

  function restoreFromReport(report: ResearchReport) {
    setReportId(report.report_id);
    if (report.pptx_file_url) setPptxFileUrl(report.pptx_file_url);
    if (report.pptx_pdf_file_url) setPptxPdfFileUrl(report.pptx_pdf_file_url);
    if (report.ppt_file_id) setPptFileId(report.ppt_file_id);
    if (report.ppt_file_url) setPptFileUrl(report.ppt_file_url);
    if (report.podcast_script) {
      setPodcastScript(report.podcast_script);
      setLastSavedScript(report.podcast_script);
    } else {
      setPodcastScript(null);
      setLastSavedScript(null);
    }
    if (report.audio_file_url) setAudioFileUrl(report.audio_file_url);
    if (report.video_file_url) setVideoFileUrl(report.video_file_url);
    if (report.is_published) {
      setIsPublished(true);
      if (report.plan) setSelectedPlan(report.plan);
    }
  }

  // ========================
  // Stage2 helpers
  // ========================

  function getSectionValue(key: string): string {
    const sec = stage2Sections.find((s) => s.key === key);
    return sec?.content?.trim() ?? '';
  }

  function parseNumber(val: string): number | null {
    if (!val) return null;
    const match = val.match(/[\d,]+\.?\d*/);
    if (!match) return null;
    const n = parseFloat(match[0].replace(/,/g, ''));
    return isNaN(n) ? null : n;
  }

  // ========================
  // Polling helper for n8n async columns
  // ========================

  const pollSupabaseColumn = useCallback(async (
    column: string,
    maxAttempts = 20,
    intervalMs = 5000,
  ): Promise<string | null> => {
    const rid = reportIdRef.current;
    if (!rid) return null;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      const { data, error } = await supabase
        .from('research_reports')
        .select(column)
        .eq('report_id', rid)
        .single();

      const record = data as Record<string, unknown> | null;
      if (!error && record?.[column]) {
        return record[column] as string;
      }
      if (attempt < maxAttempts) {
        await new Promise((r) => setTimeout(r, intervalMs));
      }
    }
    return null;
  }, []);

  // ========================
  // Step 1a: PPT copywriting pass
  // ========================

  /**
   * Runs the dedicated PPT copywriting LLM pass and persists the result on
   * the session row. The Python PPTX service reads this JSON and writes its
   * values straight into the master template, replacing the heuristic
   * truncate-and-paste path that was producing duplicate cards and mid-clause
   * cut-offs.
   *
   * Returns true on success so handleGeneratePptx can decide whether to
   * proceed with PPTX rendering after a forced regeneration.
   */
  const handleGenerateSlideCopy = useCallback(async (opts: { silent?: boolean } = {}): Promise<boolean> => {
    if (slideCopyGenerating) return false;
    if (stage2Sections.length === 0) {
      if (!opts.silent) toast.error('Stage 2 sections not loaded yet.');
      return false;
    }
    setSlideCopyGenerating(true);
    try {
      const report = await getReportBySession(sessionId);
      const reportData = report as Record<string, unknown> | null;
      const sec = (k: string) => stage2Sections.find((s) => s.key === k)?.content?.trim() ?? '';
      const meta = {
        cmp: (reportData?.cs_current_market_price as string) || sec('current_market_price'),
        target: (reportData?.cs_target_price as string) || sec('target_price'),
        upsidePct: (reportData?.cs_upside_percentage as string) || sec('upside_percentage'),
        marketCap: (reportData?.cs_market_cap as string) || sec('market_cap'),
        marketCapCategory: (reportData?.cs_market_cap_category as string) || sec('market_cap_category'),
        rating: (reportData?.cs_rating as string) || sec('rating'),
        saarthiScore: null,
      };
      const { content } = await runPptCopywriting(
        companyName,
        nseSymbol,
        sector || '',
        stage2Sections,
        meta,
        (p) => setSlideCopyProgress({ message: p.message, percent: p.percent }),
      );
      await savePptContent(sessionId, content);
      setSlideCopyReady(true);
      if (!opts.silent) toast.success(`Slide copy generated (${Object.keys(content).length} fields)`);
      return true;
    } catch (err) {
      console.error('[PostProduction] PPT copywriting failed', err);
      if (!opts.silent) toast.error(err instanceof Error ? err.message : 'PPT copywriting failed');
      return false;
    } finally {
      setSlideCopyGenerating(false);
      setSlideCopyProgress(null);
    }
  }, [slideCopyGenerating, stage2Sections, sessionId, companyName, nseSymbol, sector]);

  // ========================
  // Step 1: Generate PPTX
  // ========================

  const handleGeneratePptx = useCallback(async () => {
    if (!reportId) {
      toast.error('Report not yet created — approve stage 2 first.');
      return;
    }

    if (!financialModelFileUrl && !useMock) {
      const proceed = window.confirm(
        'Warning: No financial model Excel file was found for this session.\n\n' +
        'Generating the report now will fall back to text placeholders (charts and tables will not be injected).\n\n' +
        'We highly recommend scrolling up to the "Vault and Documents" stage, clicking "Generate Financial Model", waiting for it to complete (~10 min), and then returning here.\n\n' +
        'Do you want to proceed with the text-only fallback report anyway?'
      );
      if (!proceed) return;
    }

    setPptxGenerating(true);
    setPptxElapsedSeconds(0);
    pptxTimerRef.current = setInterval(
      () => setPptxElapsedSeconds((prev) => prev + 1),
      1000,
    );

    try {
      // Run the copywriting pass first if no cached slide copy exists. The
      // Python service still falls back to heuristic copy if this is absent,
      // but the LLM pass is what produces non-duplicate, box-budgeted content.
      if (!slideCopyReady && !useMock) {
        toast.info('Generating slide-specific copy (one-time, ~30-60s)...');
        const ok = await handleGenerateSlideCopy({ silent: true });
        if (!ok) {
          toast.warning('Slide copy step failed — falling back to heuristic copy.');
        }
      }

      const result = await generatePptx({
        reportId,
        sessionId,
        useMock,
      });

      if (result.status !== 'success' || !result.pptx_file_url) {
        throw new Error(result.message || 'PPTX generation failed');
      }

      if (result.warnings && result.warnings.length > 0) {
        for (const w of result.warnings) {
          toast.warning(w);
        }
      }

      setPptxFileUrl(result.pptx_file_url);
      setPptxPdfFileUrl(result.pptx_pdf_file_url ?? null);
      setPptFileId(result.ppt_file_id ?? null);
      setPptFileUrl(result.ppt_file_url ?? null);

      await supabase
        .from('research_reports')
        .update({ status: 'completed', updated_at: new Date().toISOString() })
        .eq('report_id', reportId);

      const dur = result.duration_seconds ? `${Math.round(result.duration_seconds)}s` : '?';
      toast.success(`PPTX generated in ${dur}${result.pptx_pdf_file_url ? ' (PDF included)' : ''}`);
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'PPTX generation failed');
    } finally {
      if (pptxTimerRef.current) {
        clearInterval(pptxTimerRef.current);
        pptxTimerRef.current = null;
      }
      setPptxGenerating(false);
    }
  }, [reportId, sessionId, useMock, slideCopyReady, handleGenerateSlideCopy, financialModelFileUrl]);

  // ========================
  // Step 1b: Sync Google Slides & Update PDF
  // ========================

  const handleSyncSlides = useCallback(async () => {
    if (!reportId) {
      toast.error('No report found to sync.');
      return;
    }
    if (!pptFileId) {
      toast.error('No Google Slides ID found for this session. Generate PPTX first.');
      return;
    }

    setSlidesSyncing(true);
    try {
      toast.info('Syncing Google Slides changes and exporting to PDF (may take ~10-15s)...');
      const result = await syncSlidesToPdf({
        reportId,
        pptFileId,
      });

      if (result.status === 'success' && result.pptx_pdf_file_url) {
        setPptxPdfFileUrl(result.pptx_pdf_file_url);
        setPreviewTab('pdf');
        toast.success('Google Slides successfully synced! PDF preview updated.');
      } else {
        throw new Error(result.message || 'Sync failed');
      }
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Failed to sync Google Slides');
    } finally {
      setSlidesSyncing(false);
    }
  }, [reportId, pptFileId]);

  // ========================
  // Step 2: Podcast
  // ========================

  const handleGenerateScript = useCallback(async () => {
    if (!reportId) return;
    setScriptGenerating(true);
    try {
      const response = await fetch(`${N8N_BASE}/generate-media-script`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ report_id: reportId }),
      });
      if (!response.ok) throw new Error('Script generation failed');

      toast.info('Script generation started — may take 1-2 minutes...');
      const script = await pollSupabaseColumn('podcast_script');

      if (script) {
        setPodcastScript(script);
        setLastSavedScript(script);
        toast.success('Podcast script generated!');
      } else {
        toast.error('Script generation timed out. Try refreshing in a minute.');
      }
    } catch (err) {
      toast.error(err instanceof Error ? err.message : 'Failed to generate script');
    } finally {
      setScriptGenerating(false);
    }
  }, [reportId, pollSupabaseColumn]);

  const handleSaveScript = useCallback(async (newScript: string, showToast = false) => {
    if (!reportId) return;
    setScriptSaving(true);
    try {
      await updateCustomSection(reportId, 'podcast_script', newScript);
      setLastSavedScript(newScript);
      if (showToast) {
        toast.success('Podcast script saved & confirmed!');
      }
    } catch (err) {
      toast.error(err instanceof Error ? err.message : 'Failed to save script');
    } finally {
      setScriptSaving(false);
    }
  }, [reportId]);

  const handleGenerateAudio = useCallback(async () => {
    if (!reportId || !podcastScript) return;
    setAudioGenerating(true);
    try {
      // Ensure the latest script version is saved to the database first
      await updateCustomSection(reportId, 'podcast_script', podcastScript);
      setLastSavedScript(podcastScript);

      const response = await fetch(`${N8N_BASE}/synthesize-podcast`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ script_text: podcastScript, report_id: reportId }),
      });
      if (!response.ok) throw new Error('Audio generation failed');

      toast.info('Audio generation started — may take 1-2 minutes...');
      const url = await pollSupabaseColumn('audio_file_url');

      if (url) {
        setAudioFileUrl(url);
        toast.success('Podcast audio generated!');
      } else {
        toast.error('Audio generation timed out. Try refreshing in a minute.');
      }
    } catch (err) {
      toast.error(err instanceof Error ? err.message : 'Failed to generate audio');
    } finally {
      setAudioGenerating(false);
    }
  }, [reportId, podcastScript, pollSupabaseColumn]);

  // ========================
  // Step 3: Video
  // ========================

  const handleGenerateVideo = useCallback(async () => {
    if (!reportId) return;
    setVideoGenerating(true);
    setVideoElapsedSeconds(0);
    const timer = setInterval(() => setVideoElapsedSeconds((p) => p + 1), 1000);

    try {
      // Fire-and-forget: use a 30 s timeout so we don't hang if the n8n
      // webhook is set to "respond when last node finishes" (which would
      // block for ~15 min and trigger a 504 from the reverse proxy / Vercel).
      // The video generation runs server-side regardless of whether we
      // receive the HTTP response — we only need to confirm the kick-off.
      let kickoffOk = false;
      try {
        const response = await fetch(`${N8N_BASE}/generate-video`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            report_id: reportId,
            company_name: companyName,
            nse_symbol: nseSymbol,
          }),
          signal: AbortSignal.timeout(30_000), // 30 s max for kickoff
        });
        kickoffOk = response.ok;
        if (!response.ok) {
          console.warn('[Video] Kickoff response not OK:', response.status);
        }
      } catch (kickoffErr) {
        // Timeout or network error — the job may still be running on n8n.
        // Log and continue to polling; if n8n truly didn't start, the poll
        // will simply time out gracefully.
        console.warn('[Video] Kickoff fetch timed out or failed (job may still be running):', kickoffErr);
      }

      toast.info(
        kickoffOk
          ? 'Video generation started — may take 10-15 minutes...'
          : 'Video generation may be starting — polling for result...',
      );

      // Poll Supabase for the video_file_url column.
      // 240 attempts × 5 s = 20 minutes — enough for a 15-min generation.
      const url = await pollSupabaseColumn('video_file_url', 240, 5000);

      if (url) {
        setVideoFileUrl(url);
        toast.success('Video generated!');
      } else {
        toast.warning('Video generation taking longer than expected. Check back later — the video may still be processing.');
      }
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Failed to generate video');
    } finally {
      clearInterval(timer);
      setVideoGenerating(false);
    }
  }, [reportId, companyName, nseSymbol, pollSupabaseColumn]);

  // ========================
  // Publish
  // ========================

  const handlePublish = async () => {
    if (!reportId || !selectedPlan) {
      toast.error('Please select a plan before publishing');
      return;
    }
    try {
      await publishReport(reportId, selectedPlan);
      setIsPublished(true);
      onPublished();
    } catch (err) {
      toast.error(`Failed to publish: ${err instanceof Error ? err.message : 'Unknown error'}`);
    }
  };

  // ========================
  // Telegram recommendation
  // ========================

  const handleSendRecommendation = async () => {
    if (!reportId || !selectedPlan) return;
    setTelegramSending(true);

    try {
      const report = await getReportBySession(sessionId);
      const reportData = report as Record<string, unknown> | null;

      const rawRating = String(reportData?.cs_rating || getSectionValue('rating') || '').toUpperCase();
      const rating: RecommendationRating = rawRating.includes('SELL') ? 'SELL' : 'BUY';

      const cmpRaw = String(reportData?.cs_current_market_price || getSectionValue('current_market_price') || '');
      const tpRaw = String(reportData?.cs_target_price || getSectionValue('target_price') || '');
      const cmp = parseNumber(cmpRaw);
      const targetPrice = parseNumber(tpRaw);
      if (!targetPrice) {
        toast.error('Target price not found in report');
        setTelegramSending(false);
        return;
      }

      // Prefer the companion PDF (better for inline preview in Telegram), fall back to PPTX.
      const reportFileUrl = pptxPdfFileUrl || pptxFileUrl || null;

      await createRecommendation({
        company_name: companyName,
        nse_symbol: nseSymbol,
        rating,
        cmp,
        target_price: targetPrice,
        validity_type: '1_year',
        validity_date: null,
        plans: [selectedPlan as never],
        trade_notes: null,
        report_file_url: reportFileUrl,
        session_id: sessionId,
        send_telegram: true,
        send_push: sendPush,
        created_by: userEmail,
        pdf_file_id: null,
      });

      setTelegramSent(true);
      toast.success('Recommendation sent to Telegram!');
    } catch (err) {
      toast.error(`Failed to send: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setTelegramSending(false);
    }
  };

  // ========================
  // Helpers
  // ========================

  const formatTime = (s: number) => `${Math.floor(s / 60)}:${(s % 60).toString().padStart(2, '0')}`;
  const isAnyGenerating = pptxGenerating || scriptGenerating || audioGenerating || videoGenerating;

  // ========================
  // RENDER
  // ========================

  return (
    <div className="rounded-xl border border-neutral-200 bg-white shadow-sm overflow-hidden">
      <div className="px-4 py-3 border-b border-neutral-100 bg-neutral-50/50 flex items-center justify-between">
        <div>
          <h2 className="text-sm font-semibold text-neutral-900">Production Workflow</h2>
          <p className="text-xs text-neutral-400 mt-1">Generate deliverables from the approved report</p>
        </div>
        <div
          className={cn(
            'flex items-center gap-1.5 text-[11px] font-medium px-2.5 py-1 rounded-full border',
            serviceHealth === 'ok'
              ? 'bg-green-50 text-green-700 border-green-200'
              : serviceHealth === 'down'
                ? 'bg-red-50 text-red-700 border-red-200'
                : 'bg-neutral-50 text-neutral-400 border-neutral-200'
          )}
          title={`PPT service: ${serviceHealth}`}
        >
          {serviceHealth === 'ok' ? (
            <><Wifi className="h-3 w-3" /> Service ✓</>
          ) : serviceHealth === 'down' ? (
            <><WifiOff className="h-3 w-3" /> Service ✗</>
          ) : (
            <><Loader2 className="h-3 w-3 animate-spin" /> Checking…</>
          )}
        </div>
      </div>

      <div className="divide-y divide-neutral-100">
        {/* === Step 0: Review & Confirm PPT Content === */}
        <div className="px-4 py-4">
          <button
            type="button"
            onClick={() => setPptDataOpen(o => !o)}
            className="w-full flex items-center justify-between group"
          >
            <div className="flex items-start gap-3">
              <div className={cn(
                'flex h-6 w-6 items-center justify-center rounded-full text-xs font-bold shrink-0 mt-0.5',
                pptDataConfirmed
                  ? 'bg-green-100 text-green-700'
                  : 'bg-accent-100 text-accent-700',
              )}>
                {pptDataConfirmed ? <Check className="h-3.5 w-3.5" strokeWidth={2.5} /> : <FileEdit className="h-3 w-3" />}
              </div>
              <div className="text-left">
                <p className={cn(
                  'text-sm font-medium',
                  pptDataConfirmed ? 'text-green-700' : 'text-neutral-900',
                )}>
                  Review PPT Content
                  {!pptDataConfirmed && (
                    <span className="ml-2 text-[10px] font-normal text-amber-600 bg-amber-50 border border-amber-200 px-1.5 py-0.5 rounded">
                      Recommended before generating
                    </span>
                  )}
                </p>
                <p className="text-xs text-neutral-400">
                  {pptDataConfirmed
                    ? 'Content confirmed — all text placeholders are set'
                    : 'Review, edit, and confirm the text that will fill each slide'}
                </p>
              </div>
            </div>
            <div className="text-neutral-400 group-hover:text-neutral-600 transition-colors ml-2">
              {pptDataOpen ? <ChevronUp className="h-4 w-4" /> : <ChevronDown className="h-4 w-4" />}
            </div>
          </button>

          {pptDataOpen && (
            <div className="mt-3 pl-9 space-y-4">
              {!slideCopyReady ? (
                <div className="rounded-lg border border-accent-100 bg-accent-50/40 p-4 space-y-3">
                  <div className="flex items-start gap-2.5">
                    <span className="text-accent-600 shrink-0 text-lg font-semibold mt-0.5">✨</span>
                    <div className="text-xs text-accent-950 leading-normal">
                      <p className="font-semibold text-[13px] text-accent-900">AI Slide Copywriting Pass Required</p>
                      <p className="mt-1 text-neutral-600">
                        This step uses the LLM to summarize and format your research sections into 
                        concise, professionally budgeted slide bullet points. Without this, the 
                        review panel and final presentation will fall back to truncated raw report text.
                      </p>
                    </div>
                  </div>
                  <div className="flex flex-col gap-3">
                    <div className="flex items-center gap-3">
                      <Button
                        type="button"
                        onClick={() => handleGenerateSlideCopy()}
                        disabled={slideCopyGenerating || stage2Sections.length === 0}
                        size="sm"
                        className="rounded-lg bg-accent-600 hover:bg-accent-700 text-white font-medium shadow-sm transition-all"
                      >
                        {slideCopyGenerating ? (
                          <><Loader2 className="h-3.5 w-3.5 mr-1.5 animate-spin" /> Generating slide copy...</>
                        ) : (
                          <><FileEdit className="h-3.5 w-3.5 mr-1.5" /> Run AI Copywriting Pass (~30s)</>
                        )}
                      </Button>
                    </div>

                    {slideCopyGenerating && slideCopyProgress && (
                      <div className="space-y-1.5 max-w-md bg-accent-50/50 p-2.5 rounded-lg border border-accent-100/40">
                        <div className="flex justify-between text-[11px]">
                          <span className="text-accent-800 font-medium animate-pulse">{slideCopyProgress.message}</span>
                          <span className="text-accent-700 font-bold">{slideCopyProgress.percent}%</span>
                        </div>
                        <div className="w-full bg-neutral-200/50 rounded-full h-1 overflow-hidden">
                          <div 
                            className="bg-accent-600 h-1 rounded-full transition-all duration-300"
                            style={{ width: `${slideCopyProgress.percent}%` }}
                          />
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              ) : (
                <>
                  <div className="flex flex-col gap-2 border-b border-neutral-100 pb-2 mb-2">
                    <div className="flex items-center justify-between text-xs">
                      <span className="text-green-700 font-semibold flex items-center gap-1.5 bg-green-50 px-2 py-0.5 rounded-full border border-green-200">
                        <Check className="h-3.5 w-3.5" /> AI Slide Copy is ready & fitted
                      </span>
                      <button
                        type="button"
                        onClick={() => handleGenerateSlideCopy()}
                        disabled={slideCopyGenerating}
                        className="text-accent-600 hover:text-accent-700 underline flex items-center gap-1 disabled:text-neutral-400 disabled:no-underline font-medium"
                      >
                        {slideCopyGenerating ? (
                          <><Loader2 className="h-3 w-3 animate-spin" /> Regenerating...</>
                        ) : (
                          <><RefreshCw className="h-3 w-3" /> Re-run AI copywriting</>
                        )}
                      </button>
                    </div>

                    {slideCopyGenerating && slideCopyProgress && (
                      <div className="space-y-1 mt-1 bg-accent-50/30 p-2 rounded-md border border-accent-100/40">
                        <div className="flex justify-between text-[10px]">
                          <span className="text-accent-800 font-medium animate-pulse">{slideCopyProgress.message}</span>
                          <span className="text-accent-700 font-bold">{slideCopyProgress.percent}%</span>
                        </div>
                        <div className="w-full bg-neutral-200/50 rounded-full h-1 overflow-hidden">
                          <div 
                            className="bg-accent-600 h-1 rounded-full transition-all duration-300"
                            style={{ width: `${slideCopyProgress.percent}%` }}
                          />
                        </div>
                      </div>
                    )}
                  </div>

                  <PPTDataPanel
                    reportId={reportId}
                    sessionId={sessionId}
                    serviceAvailable={serviceHealth === 'ok'}
                    onConfirmed={() => setPptDataConfirmed(true)}
                  />
                </>
              )}
            </div>
          )}
        </div>

        {/* === Step 1: Generate PPTX === */}
        <StepRow
          number={1}
          title="Generate PPTX Report"
          description="Build the PowerPoint deck via the reportgen pipeline"
          done={!!pptxFileUrl}
          active={!pptxFileUrl}
        >
          {!pptxFileUrl ? (
            <div className="space-y-2">
              <div className="flex items-center gap-3">
                <Button
                  onClick={handleGeneratePptx}
                  disabled={!reportId || pptxGenerating || isAnyGenerating || stage2Sections.length === 0}
                  size="sm"
                  className="rounded-lg bg-accent-600 hover:bg-accent-700"
                >
                  {pptxGenerating ? (
                    <><Loader2 className="h-3.5 w-3.5 mr-1.5 animate-spin" /> {pptxElapsedSeconds > 15 ? 'Injecting Excel Charts...' : 'Building Slides...'} ({formatTime(pptxElapsedSeconds)})</>
                  ) : (
                    <><Presentation className="h-3.5 w-3.5 mr-1.5" /> Generate PPTX</>
                  )}
                </Button>
                <label className="flex items-center gap-1.5 text-xs text-neutral-500 cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={useMock}
                    onChange={(e) => setUseMock(e.target.checked)}
                    className="h-3.5 w-3.5 rounded border-neutral-300 text-accent-600"
                  />
                  Mock planner (skip OpenRouter)
                </label>
              </div>
              {pptxGenerating && (
                <div className="text-[11px] text-accent-600 animate-pulse mt-2 flex items-center gap-1.5">
                  <div className="h-1.5 w-1.5 rounded-full bg-accent-600" />
                  Generating text, building slides, and importing high-fidelity Excel tables...
                </div>
              )}
              {!reportId && (
                <p className="text-[11px] text-neutral-400">Approve stage 2 first to create the report row.</p>
              )}

              {!financialModelFileUrl && !useMock && (
                <div className="mt-3 p-3 bg-amber-50 border border-amber-200 rounded-lg text-amber-800 text-xs">
                  <p className="font-semibold flex items-center gap-1.5 mb-1">
                    ⚠️ No Financial Model Excel Found
                  </p>
                  <p className="text-amber-700 leading-normal">
                    No financial model is associated with this session. Generating the PowerPoint now will 
                    <strong> fall back to text placeholders</strong> (charts and financial tables will not be injected). 
                    To get the complete report, please scroll up to the <strong>Vault and Documents</strong> stage and click 
                    <strong>"Generate Financial Model"</strong> first.
                  </p>
                </div>
              )}

              <div className="mt-4 p-3 bg-blue-50/50 border border-blue-100 rounded-lg">
                <p className="text-[11px] font-semibold text-blue-800 mb-2">Automated Excel Injections</p>
                <div className="grid grid-cols-2 gap-x-4 gap-y-1.5 text-[10px] text-neutral-600 font-mono">
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{financial_model_from_excel}}`}</span><span className="text-blue-700 font-semibold">Op_Charts</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{financial_model_from_excel_operational_sheet}}`}</span><span className="text-blue-700 font-semibold">Operational_Data</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{financial_summary_image}}`}</span><span className="text-blue-700 font-semibold">Fin_Summary</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{earnings_forecast_table}}`}</span><span className="text-blue-700 font-semibold">Earnings_Forecast</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{financials_table}}`}</span><span className="text-blue-700 font-semibold">Financials_Table</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{valuations_table}}`}</span><span className="text-blue-700 font-semibold">Valuations_Table</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{key_risks_table}}`}</span><span className="text-blue-700 font-semibold">Key_Risks</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{peer_comparision}}`}</span><span className="text-blue-700 font-semibold">Peer_Compare</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{governance_table}}`}</span><span className="text-blue-700 font-semibold">Governance</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{timeline}}`}</span><span className="text-blue-700 font-semibold">Timeline</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{competitive_chart_1}}`}</span><span className="text-blue-700 font-semibold">Comp_Chart_1</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{competitive_chart_2}}`}</span><span className="text-blue-700 font-semibold">Comp_Chart_2</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{pie_chart_1}}`}</span><span className="text-blue-700 font-semibold">Pie_Chart_1</span></div>
                  <div className="flex justify-between border-b border-blue-100 pb-1"><span>{`{{pie_chart_2}}`}</span><span className="text-blue-700 font-semibold">Pie_Chart_2</span></div>
                  <div className="flex justify-between pb-1"><span>{`{{probability_weight_table}}`}</span><span className="text-blue-700 font-semibold">Prob_Weight</span></div>
                </div>
              </div>
            </div>
          ) : (
            <div className="space-y-3">
              <div className="flex flex-wrap items-center gap-3">
                <a
                  href={pptxFileUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1.5 rounded-lg bg-accent-600 hover:bg-accent-700 text-white text-xs font-medium h-8 px-3"
                >
                  <Download className="h-3.5 w-3.5" /> Download PPTX
                </a>
                {pptxPdfFileUrl && (
                  <a
                    href={pptxPdfFileUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1.5 rounded-lg border border-neutral-200 hover:bg-neutral-50 text-neutral-700 text-xs font-medium h-8 px-3"
                  >
                    <ExternalLink className="h-3.5 w-3.5" /> View PDF
                  </a>
                )}
                {pptFileUrl && (
                  <a
                    href={pptFileUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1.5 rounded-lg bg-orange-600 hover:bg-orange-700 text-white text-xs font-medium h-8 px-3 transition-colors"
                  >
                    <ExternalLink className="h-3.5 w-3.5" /> Edit in Google Slides
                  </a>
                )}
                {pptFileId && (
                  <Button
                    onClick={handleSyncSlides}
                    disabled={slidesSyncing || pptxGenerating || isAnyGenerating}
                    size="sm"
                    className="h-8 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-medium px-3"
                  >
                    {slidesSyncing ? (
                      <><Loader2 className="h-3.5 w-3.5 mr-1.5 animate-spin" /> Syncing Slides...</>
                    ) : (
                      <><RefreshCw className="h-3.5 w-3.5 mr-1.5" /> Sync Slides & Update PDF</>
                    )}
                  </Button>
                )}
                <Button
                  onClick={handleGeneratePptx}
                  disabled={pptxGenerating || isAnyGenerating}
                  size="sm"
                  variant="ghost"
                  className="h-7 text-xs text-neutral-500 hover:text-neutral-800"
                >
                  {pptxGenerating ? (
                    <><Loader2 className="h-3 w-3 mr-1.5 animate-spin" /> Regenerating ({formatTime(pptxElapsedSeconds)})</>
                  ) : (
                    <><RefreshCw className="h-3 w-3 mr-1.5" /> Regenerate</>
                  )}
                </Button>
              </div>

              {pptFileId && (
                <div className="flex border-b border-neutral-200 mb-3 bg-neutral-50/50 rounded-t-lg p-1 gap-1">
                  <button
                    type="button"
                    onClick={() => setPreviewTab('pdf')}
                    className={cn(
                      'px-4 py-2 text-xs font-semibold border-b-2 transition-all flex items-center gap-1.5 rounded-md',
                      previewTab === 'pdf'
                        ? 'border-accent-600 text-accent-700 bg-white shadow-sm'
                        : 'border-transparent text-neutral-500 hover:text-neutral-700 hover:bg-neutral-100/50'
                    )}
                  >
                    <FileText className="h-3.5 w-3.5" /> PDF Preview
                  </button>
                  <button
                    type="button"
                    onClick={() => setPreviewTab('slides')}
                    className={cn(
                      'px-4 py-2 text-xs font-semibold border-b-2 transition-all flex items-center gap-1.5 rounded-md',
                      previewTab === 'slides'
                        ? 'border-accent-600 text-accent-700 bg-white shadow-sm'
                        : 'border-transparent text-neutral-500 hover:text-neutral-700 hover:bg-neutral-100/50'
                    )}
                  >
                    <Presentation className="h-3.5 w-3.5" /> Google Slides Editor
                  </button>
                </div>
              )}

              {previewTab === 'slides' && pptFileId ? (
                <div className="space-y-3">
                  <div className="p-3 bg-blue-50/50 border border-blue-100 rounded-lg flex items-start gap-2.5">
                    <span className="text-blue-700 shrink-0 text-sm font-semibold mt-0.5">💡</span>
                    <div className="text-xs text-blue-800 leading-normal">
                      <p className="font-semibold">Visual Editing Mode</p>
                      <p className="mt-0.5">
                        You can visually edit the presentation deck directly inside the workspace below. Changes are saved automatically.
                        Once you're done, click the <strong className="text-indigo-700 font-semibold">Sync Slides & Update PDF</strong> button above to re-compile your changes into the final PDF.
                      </p>
                    </div>
                  </div>
                  <iframe
                    key={pptFileId}
                    src={`https://docs.google.com/presentation/d/${pptFileId}/edit?usp=drivesdk`}
                    title={`${companyName} Google Slides Editor`}
                    className="w-full rounded-lg border border-neutral-200 bg-white"
                    style={{ height: '600px' }}
                    allowFullScreen
                  />
                </div>
              ) : (
                <>
                  {pptxPdfFileUrl ? (
                    <iframe
                      key={pptxPdfFileUrl}
                      src={pptxPdfFileUrl}
                      title={`${companyName} report preview`}
                      className="w-full rounded-lg border border-neutral-200 bg-white"
                      style={{ height: '520px' }}
                    />
                  ) : (
                    <iframe
                      key={pptxFileUrl}
                      src={`https://docs.google.com/viewer?url=${encodeURIComponent(pptxFileUrl || '')}&embedded=true`}
                      title={`${companyName} report preview`}
                      className="w-full rounded-lg border border-neutral-200 bg-white"
                      style={{ height: '520px' }}
                    />
                  )}
                </>
              )}
            </div>
          )}
        </StepRow>

        {/* === Step 2: Podcast === */}
        <StepRow
          number={2}
          title="Generate Podcast"
          description="Create script, then synthesize audio"
          done={!!audioFileUrl}
          active={!!reportId && !audioFileUrl}
          disabled={!reportId}
        >
          <div className="space-y-3">
            {!podcastScript ? (
              <Button
                onClick={handleGenerateScript}
                disabled={!reportId || scriptGenerating || isAnyGenerating}
                size="sm"
                variant={audioFileUrl ? 'outline' : 'default'}
                className={cn('rounded-lg', !audioFileUrl && 'bg-accent-600 hover:bg-accent-700')}
              >
                {scriptGenerating ? (
                  <><Loader2 className="h-3.5 w-3.5 mr-1.5 animate-spin" /> Generating Script...</>
                ) : (
                  <><Mic className="h-3.5 w-3.5 mr-1.5" /> Generate Script</>
                )}
              </Button>
            ) : (
              <>
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    {podcastScript !== lastSavedScript ? (
                      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-medium bg-amber-50 text-amber-700 border border-amber-200 animate-pulse">
                        <AlertCircle className="h-3 w-3" /> Unconfirmed Edits
                      </span>
                    ) : (
                      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-medium bg-green-50 text-green-700 border border-green-200">
                        <CheckCircle className="h-3 w-3" /> Script Confirmed
                      </span>
                    )}
                    <button
                      onClick={() => setScriptExpanded(!scriptExpanded)}
                      className="text-xs text-neutral-400 hover:text-neutral-600 flex items-center gap-1 ml-1"
                    >
                      {scriptExpanded ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />}
                      {scriptExpanded ? 'Hide Script' : 'Edit/View Script'}
                    </button>
                  </div>
                </div>

                {scriptExpanded && (
                  <div className="space-y-2">
                    <textarea
                      value={podcastScript || ''}
                      onChange={(e) => setPodcastScript(e.target.value)}
                      className="w-full rounded-lg border border-neutral-200 bg-neutral-50 p-3 text-xs text-neutral-800 font-mono leading-relaxed resize-y focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent-500/40 focus-visible:border-accent-400"
                      style={{ minHeight: '120px', maxHeight: '300px' }}
                      spellCheck={false}
                      placeholder="Podcast script content..."
                    />
                    <div className="flex justify-end">
                      <Button
                        onClick={() => handleSaveScript(podcastScript || '', true)}
                        disabled={scriptSaving || podcastScript === lastSavedScript}
                        size="xs"
                        variant={podcastScript !== lastSavedScript ? 'default' : 'outline'}
                        className={cn(
                          'rounded-md px-2.5 py-1 text-[11px] h-7',
                          podcastScript !== lastSavedScript
                            ? 'bg-blue-600 hover:bg-blue-700 text-white shadow-sm border-0'
                            : 'text-neutral-500 bg-neutral-50 border-neutral-200'
                        )}
                      >
                        {scriptSaving ? (
                          <><Loader2 className="h-3 w-3 mr-1 animate-spin" /> Saving...</>
                        ) : podcastScript !== lastSavedScript ? (
                          <><Check className="h-3 w-3 mr-1" /> Confirm & Update Script</>
                        ) : (
                          <><Check className="h-3 w-3 mr-1 text-green-600" /> Saved & Confirmed</>
                        )}
                      </Button>
                    </div>
                  </div>
                )}

                {!audioFileUrl ? (
                  <div className="space-y-2">
                    <div className="flex items-center gap-3">
                      <Button
                        onClick={handleGenerateAudio}
                        disabled={audioGenerating || isAnyGenerating || podcastScript !== lastSavedScript}
                        size="sm"
                        className={cn(
                          'rounded-lg',
                          podcastScript !== lastSavedScript
                            ? 'bg-neutral-100 text-neutral-400 cursor-not-allowed border border-neutral-200 hover:bg-neutral-100'
                            : 'bg-accent-600 hover:bg-accent-700 text-white'
                        )}
                      >
                        {audioGenerating ? (
                          <><Loader2 className="h-3.5 w-3.5 mr-1.5 animate-spin" /> Generating Audio...</>
                        ) : (
                          <><Play className="h-3.5 w-3.5 mr-1.5" /> Generate Audio</>
                        )}
                      </Button>
                    </div>
                    {podcastScript !== lastSavedScript && (
                      <p className="text-[10px] text-amber-600 leading-normal animate-pulse">
                        ⚠️ Please click <strong>"Confirm & Update Script"</strong> in the editor above to save your changes and enable audio generation.
                      </p>
                    )}
                  </div>
                ) : (
                  <div className="space-y-2">
                    <div className="flex items-center gap-3">
                      <audio controls src={audioFileUrl} className="h-8 flex-1" />
                      <a href={audioFileUrl} download className="text-xs text-neutral-500 hover:text-neutral-700 flex items-center gap-1 shrink-0">
                        <Download className="h-3 w-3" /> MP3
                      </a>
                    </div>
                    {podcastScript !== lastSavedScript && (
                      <div className="p-2.5 bg-amber-50 border border-amber-200 rounded-lg text-amber-800 text-[10px] leading-normal space-y-1">
                        <p className="font-semibold flex items-center gap-1">
                          ⚠️ Unsaved Script Edits
                        </p>
                        <p className="text-amber-700">
                          You have modified the script but the generated audio is still using the old version. Click <strong>"Confirm & Update Script"</strong> above, then you can regenerate the audio to update it.
                        </p>
                        <Button
                          onClick={handleGenerateAudio}
                          disabled={audioGenerating || isAnyGenerating}
                          size="xs"
                          variant="outline"
                          className="mt-1 h-6 text-[10px] bg-white text-amber-800 border-amber-300 hover:bg-amber-100 hover:text-amber-900 rounded"
                        >
                          {audioGenerating ? 'Regenerating Audio...' : 'Regenerate Audio with Updated Script'}
                        </Button>
                      </div>
                    )}
                  </div>
                )}
              </>
            )}
          </div>
        </StepRow>

        {/* === Step 3: Video === */}
        <StepRow
          number={3}
          title="Generate Video"
          description="Create video summary from report"
          done={!!videoFileUrl}
          active={!!reportId && !videoFileUrl}
          disabled={!reportId}
        >
          {!videoFileUrl ? (
            <Button
              onClick={handleGenerateVideo}
              disabled={!reportId || videoGenerating || isAnyGenerating}
              size="sm"
              className="rounded-lg bg-accent-600 hover:bg-accent-700"
            >
              {videoGenerating ? (
                <><Loader2 className="h-3.5 w-3.5 mr-1.5 animate-spin" /> Generating ({formatTime(videoElapsedSeconds)})...</>
              ) : (
                <><Video className="h-3.5 w-3.5 mr-1.5" /> Generate Video</>
              )}
            </Button>
          ) : (
            <div className="space-y-3">
              <video controls src={videoFileUrl} className="w-full max-h-48 rounded-lg bg-black" />
              <div className="flex flex-wrap items-center gap-3">
                <a
                  href={videoFileUrl}
                  download
                  className="inline-flex items-center gap-1.5 rounded-lg border border-neutral-200 hover:bg-neutral-50 text-neutral-700 text-xs font-medium h-8 px-3"
                >
                  <Download className="h-3.5 w-3.5" /> Download MP4
                </a>
                <Button
                  onClick={handleGenerateVideo}
                  disabled={videoGenerating || isAnyGenerating}
                  size="sm"
                  variant="ghost"
                  className="h-7 text-xs text-neutral-500 hover:text-neutral-800"
                >
                  {videoGenerating ? (
                    <><Loader2 className="h-3 w-3 mr-1.5 animate-spin" /> Regenerating ({formatTime(videoElapsedSeconds)})</>
                  ) : (
                    <><RefreshCw className="h-3 w-3 mr-1.5" /> Regenerate Video</>
                  )}
                </Button>
              </div>
            </div>
          )}
        </StepRow>
      </div>

      {/* === Publish === */}
      <div className="px-4 py-4 border-t border-neutral-100 bg-neutral-50/30 flex flex-col gap-3">
        {pptxFileUrl && (
          <div className="space-y-2 focus-within:relative z-10">
            <label className="text-xs font-semibold text-neutral-700">Select Plan to Publish For</label>
            <Select value={selectedPlan} onValueChange={setSelectedPlan}>
              <SelectTrigger className="w-full bg-white text-sm">
                <SelectValue placeholder="Choose a plan..." />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="midcap_wealth">Mid Cap Wealth Builders</SelectItem>
                <SelectItem value="smallcap_alpha">Smallcap Alpha Picks</SelectItem>
                <SelectItem value="sme_emerging">SME Emerging Business</SelectItem>
              </SelectContent>
            </Select>
          </div>
        )}

        {!isPublished ? (
          <>
            <Button
              onClick={handlePublish}
              disabled={!pptxFileUrl || !selectedPlan || isAnyGenerating}
              className={cn(
                'w-full h-10 rounded-lg font-semibold text-sm',
                pptxFileUrl && selectedPlan
                  ? 'bg-green-600 hover:bg-green-700 text-white shadow-sm'
                  : 'bg-neutral-200 text-neutral-400 cursor-not-allowed'
              )}
            >
              <Shield className="h-4 w-4 mr-2" /> Publish Report
            </Button>
            {!pptxFileUrl && (
              <p className="text-xs text-neutral-400 text-center mt-2">PPTX must be generated before publishing</p>
            )}
          </>
        ) : (
          <div className="flex items-center gap-2 text-green-600 text-sm font-medium">
            <Check className="h-4 w-4" /> Report Published
          </div>
        )}
      </div>

      {/* === Telegram === */}
      {isPublished && (
        <div className="px-4 py-4 border-t border-neutral-100 bg-gradient-to-b from-blue-50/40 to-white">
          <div className="flex items-start gap-3">
            <div className={cn(
              'flex h-6 w-6 items-center justify-center rounded-full text-xs font-bold shrink-0 mt-1',
              telegramSent ? 'bg-green-100 text-green-700' : 'bg-blue-100 text-blue-700'
            )}>
              {telegramSent ? <Check className="h-3.5 w-3.5" strokeWidth={2.5} /> : 4}
            </div>

            <div className="flex-1 min-w-0">
              <p className={cn('text-sm font-medium', telegramSent ? 'text-green-700' : 'text-neutral-900')}>
                Send Recommendation to Telegram
              </p>
              <p className="text-xs text-neutral-400 mb-3">
                Create a recommendation record and send to subscribers
              </p>

              {(() => {
                const rating = getSectionValue('rating').toUpperCase().includes('SELL') ? 'SELL' : 'BUY';
                const cmp = parseNumber(getSectionValue('current_market_price'));
                const tp = parseNumber(getSectionValue('target_price'));
                const upside = cmp && tp ? (((tp - cmp) / cmp) * 100).toFixed(1) : null;
                const planLabel = selectedPlan === 'midcap_wealth' ? 'Mid Cap Wealth Builders'
                  : selectedPlan === 'smallcap_alpha' ? 'Smallcap Alpha Picks'
                    : selectedPlan === 'sme_emerging' ? 'SME Emerging Business' : selectedPlan;

                const attachedUrl = pptxPdfFileUrl || pptxFileUrl;
                const attachedLabel = pptxPdfFileUrl ? 'PDF attached' : 'PPTX attached';

                return (
                  <div className="rounded-lg border border-neutral-200 bg-white p-3 mb-3 space-y-2">
                    <div className="grid grid-cols-2 gap-x-4 gap-y-1.5 text-xs">
                      <div>
                        <span className="text-neutral-400">Company</span>
                        <p className="font-medium text-neutral-800">{companyName} ({nseSymbol})</p>
                      </div>
                      <div>
                        <span className="text-neutral-400">Rating</span>
                        <p className={cn('font-semibold', rating === 'BUY' ? 'text-green-600' : 'text-red-600')}>{rating}</p>
                      </div>
                      <div>
                        <span className="text-neutral-400">CMP</span>
                        <p className="font-medium text-neutral-800">{cmp != null ? `₹${cmp.toLocaleString('en-IN')}` : '—'}</p>
                      </div>
                      <div>
                        <span className="text-neutral-400">Target Price</span>
                        <p className="font-medium text-neutral-800">{tp != null ? `₹${tp.toLocaleString('en-IN')}` : '—'}</p>
                      </div>
                      <div>
                        <span className="text-neutral-400">Upside</span>
                        <p className="font-medium text-neutral-800">{upside ? `${upside}%` : '—'}</p>
                      </div>
                      <div>
                        <span className="text-neutral-400">Plan</span>
                        <p className="font-medium text-neutral-800">{planLabel}</p>
                      </div>
                    </div>
                    {attachedUrl && (
                      <div className="text-xs text-neutral-400 pt-1 border-t border-neutral-100">
                        Report: <a href={attachedUrl} target="_blank" rel="noopener noreferrer" className="text-accent-600 hover:underline">{attachedLabel}</a>
                      </div>
                    )}
                  </div>
                );
              })()}

              {!telegramSent && (
                <div className="flex items-center gap-2 mb-3">
                  <input
                    type="checkbox"
                    id="send-push-checkbox-postprod"
                    checked={sendPush}
                    onChange={(e) => setSendPush(e.target.checked)}
                    className="rounded border-neutral-300 text-accent-600 focus:ring-accent-500/40 h-4 w-4 cursor-pointer"
                  />
                  <label htmlFor="send-push-checkbox-postprod" className="text-xs text-neutral-600 font-medium select-none cursor-pointer">
                    Send Push Notification to App Subscribers
                  </label>
                </div>
              )}

              {!telegramSent ? (
                <Button
                  onClick={handleSendRecommendation}
                  disabled={telegramSending}
                  size="sm"
                  className="rounded-lg bg-blue-600 hover:bg-blue-700 text-white"
                >
                  {telegramSending ? (
                    <><Loader2 className="h-3.5 w-3.5 mr-1.5 animate-spin" /> Sending...</>
                  ) : (
                    <><Send className="h-3.5 w-3.5 mr-1.5" /> Send to Telegram</>
                  )}
                </Button>
              ) : (
                <span className="text-xs text-green-600 font-medium">Recommendation sent to Telegram</span>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ========================
// Step Row sub-component
// ========================

interface StepRowProps {
  number: number;
  title: string;
  description: string;
  done: boolean;
  active: boolean;
  disabled?: boolean;
  children: React.ReactNode;
}

function StepRow({ number, title, description, done, active, disabled, children }: StepRowProps) {
  return (
    <div className={cn('px-4 py-4 flex items-start gap-3', disabled && 'opacity-50')}>
      <div className={cn(
        'flex h-6 w-6 items-center justify-center rounded-full text-xs font-bold shrink-0 mt-1',
        done ? 'bg-green-100 text-green-700' :
          active ? 'bg-accent-100 text-accent-700' :
            'bg-neutral-100 text-neutral-400'
      )}>
        {done ? <Check className="h-3.5 w-3.5" strokeWidth={2.5} /> : number}
      </div>
      <div className="flex-1 min-w-0">
        <div>
          <p className={cn('text-sm font-medium', done ? 'text-green-700' : active ? 'text-neutral-900' : 'text-neutral-500')}>
            {title}
          </p>
          <p className="text-xs text-neutral-400">{description}</p>
        </div>
        <div className="mt-2">
          {children}
        </div>
      </div>
    </div>
  );
}
