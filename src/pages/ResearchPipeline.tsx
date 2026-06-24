import { useState, useEffect, useCallback, useRef } from 'react';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Spinner } from '@/components/ui/spinner';
import PipelineProgressBar from '@/components/pipeline/PipelineProgressBar';
import StageReview from '@/components/pipeline/StageReview';
import PromptEditor from '@/components/pipeline/PromptEditor';
import PostProductionPanel from '@/components/pipeline/PostProductionPanel';
import { useCompanySearch, useCompanyFinancials } from '@/hooks/useCompanySearch';
import { useAuth } from '@/contexts/AuthContext';
import {
  createPipelineSession,
  transitionPipelineStatus,
  updatePipelineOutput,
  saveResearchSection,
  clearResearchSections,
  updateResearchSection,
  listPipelineSessions,
  deletePipelineSession,
  getPipelineSession,
  getResearchSections,
  getFrameworkFromPlaybook,
  getSectorPlaybook,
} from '@/lib/pipeline-api';
import {
  createVault,
  processVaultResponse,
  saveSessionDocuments,
  getSessionDocuments,
  generateFinancialModel,
  mirrorFinancialModelToStorage,
  unpublishReport,
  getReportBySession,
  createResearchReport,
  uploadDocument,
} from '@/lib/api';
import type { ResearchReport, SessionDocument } from '@/types/database';
import { createRecommendation, hasRecommendationForSession } from '@/lib/recommendations-api';
import type { RecommendationRating, PlanId } from '@/types/recommendations';
import { runStage0, runStage1, runStage2, DEFAULT_PROMPTS, summarizeVaultDocuments } from '@/lib/anthropic-pipeline';
import type { PromptOverrides } from '@/lib/anthropic-pipeline';
import type { PipelineSession, PipelineProgress, PipelineStatus, SectorFramework } from '@/types/pipeline';
import { PIPELINE_STAGE_LABELS, PIPELINE_MODELS, DEFAULT_PIPELINE_MODEL, getStageNumber } from '@/types/pipeline';
import type { MasterCompany } from '@/types/database';
import type { VaultDocument } from '@/types/vault';
import DocumentUploadDialog from '@/components/DocumentUploadDialog';
import { cn } from '@/lib/utils';
import { SECTORS } from '@/lib/sectors';
import {
  Search,
  Building2,
  BarChart3,
  Upload,
  Check,
  Trash2,
  ExternalLink,
  RefreshCw,
  Loader2,
  Eye,
  EyeOff,
  Sparkles,
  FileText,
  Globe,
  ChevronRight,
  Play,
  Zap,
  X,
  Clock,
  Hash,
  ChevronDown,
  Send,
} from 'lucide-react';

// ========================
// Formatting Helpers
// ========================

function formatCurrency(value: number | null): string {
  if (value == null) return '-';
  if (value >= 10000000) return `₹${(value / 10000000).toFixed(2)} Cr`;
  if (value >= 100000) return `₹${(value / 100000).toFixed(2)} L`;
  return `₹${value.toLocaleString('en-IN', { maximumFractionDigits: 2 })}`;
}

function formatPercent(value: number | null): string {
  if (value == null) return '-';
  return `${value.toFixed(1)}%`;
}

// ========================
// Main Page Component
// ========================

export default function ResearchPipeline() {
  const { user } = useAuth();

  // --- Company Search State ---
  const [searchInput, setSearchInput] = useState('');
  const [debouncedSearch, setDebouncedSearch] = useState('');
  const [selectedCompany, setSelectedCompany] = useState<MasterCompany | null>(null);
  const [isDropdownOpen, setIsDropdownOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  // --- Setup State ---
  const [selectedSector, setSelectedSector] = useState('');
  const [selectedModel, setSelectedModel] = useState(DEFAULT_PIPELINE_MODEL);

  // --- Session State ---
  const [sessionId, setSessionId] = useState<string | null>(null);
  const [session, setSession] = useState<PipelineSession | null>(null);
  const [isCreatingSession, setIsCreatingSession] = useState(false);
  const [initialReport, setInitialReport] = useState<ResearchReport | null>(null);

  // --- Vault & Documents ---
  const [vaultStatus, setVaultStatus] = useState<'idle' | 'loading' | 'success' | 'error'>('idle');
  const [vaultLink, setVaultLink] = useState<string | null>(null);
  const [vaultId, setVaultId] = useState<string | null>(null);
  const [vaultDocuments, setVaultDocuments] = useState<VaultDocument[]>([]);
  const [isUploadOpen, setIsUploadOpen] = useState(false);

  // --- Financial Model State ---
  const [financialModelStatus, setFinancialModelStatus] = useState<'idle' | 'generating' | 'success' | 'skipped'>('idle');
  const [financialModelFileUrl, setFinancialModelFileUrl] = useState<string | null>(null);

  // --- Pipeline Stage State ---
  const [pipelineStatus, setPipelineStatus] = useState<PipelineStatus>('company_selected');
  const [isRunning, setIsRunning] = useState(false);
  const [progress, setProgress] = useState<PipelineProgress | null>(null);

  // --- Stage Outputs ---
  const [sectorFramework, setSectorFramework] = useState<SectorFramework | null>(null);
  const [stage1Thesis, setStage1Thesis] = useState<string>('');
  const [stage2Sections, setStage2Sections] = useState<Array<{ id?: string; key: string; title: string; content: string }>>([]);

  // --- Prompt Overrides ---
  const [showPrompt, setShowPrompt] = useState<'stage0' | 'stage1' | 'stage2' | null>(null);
  const [stage0Prompts, setStage0Prompts] = useState<PromptOverrides>({});
  const [stage1Prompts, setStage1Prompts] = useState<PromptOverrides>({});
  const [stage2Prompts, setStage2Prompts] = useState<PromptOverrides>({});

  // --- Recent Sessions ---
  const [recentSessions, setRecentSessions] = useState<PipelineSession[]>([]);

  // --- Report Section Tabs ---
  const [activeReportTab, setActiveReportTab] = useState(0);

  // --- Telegram Recommendation ---
  const [telegramSending, setTelegramSending] = useState(false);
  const [telegramSent, setTelegramSent] = useState(false);
  const [sendPush, setSendPush] = useState(true);

  // --- Data Queries ---
  const { data: companies } = useCompanySearch(debouncedSearch);
  const { data: financials } = useCompanyFinancials(selectedCompany);

  // Auto-fill sector from financials — match to closest SECTORS entry
  useEffect(() => {
    if (financials && !selectedSector) {
      const raw = (financials.sector || financials.broad_sector || '').toLowerCase();
      if (!raw) return;
      const exact = SECTORS.find(s => s.toLowerCase() === raw);
      const partial = !exact ? SECTORS.find(s => raw.includes(s.toLowerCase()) || s.toLowerCase().includes(raw)) : undefined;
      setSelectedSector(exact || partial || financials.sector || financials.broad_sector || '');
    }
  }, [financials, selectedSector]);

  // Debounce search
  useEffect(() => {
    const timer = setTimeout(() => setDebouncedSearch(searchInput), 300);
    return () => clearTimeout(timer);
  }, [searchInput]);

  // Close dropdown on outside click
  useEffect(() => {
    if (!isDropdownOpen) return;
    const handler = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setIsDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [isDropdownOpen]);

  // Load recent sessions
  useEffect(() => {
    if (!user?.email) return;
    listPipelineSessions({ pageSize: 50 })
      .then(({ data }) => setRecentSessions(data))
      .catch(() => {});
  }, [user?.email]);

  // Load session data when sessionId changes
  useEffect(() => {
    if (!sessionId) return;
    getPipelineSession(sessionId).then((s) => {
      if (s) {
        setSession(s);
        setPipelineStatus((s.pipeline_status ?? 'company_selected') as PipelineStatus);
        setSelectedModel(s.selected_model ?? DEFAULT_PIPELINE_MODEL);

        // Restore vault data
        if (s.vault_folder_id) {
          setVaultId(s.vault_folder_id);
          setVaultLink(s.vault_folder_url || `https://drive.google.com/drive/folders/${s.vault_folder_id}`);
          setVaultStatus('success');
        }

        // Restore financial model
        if (s.financial_model_file_url) {
          setFinancialModelStatus('success');
          setFinancialModelFileUrl(s.financial_model_file_url);
        }

        if (s.sector_framework) {
          setSectorFramework(s.sector_framework);
        } else {
          const sectorName = s.sector || selectedSector;
          if (sectorName) {
            getSectorPlaybook(sectorName).then((pb) => {
              if (pb) {
                setSectorFramework({
                  sector_name: sectorName,
                  markdown: getFrameworkFromPlaybook(pb),
                  version: pb.version,
                  last_updated: pb.last_updated,
                });
              }
            });
          }
        }
        if (s.thesis_output) setStage1Thesis(s.thesis_output);
      }
    });
    // Restore stage2 sections
    getResearchSections(sessionId, 'stage2').then((sections) => {
      if (sections.length > 0) {
        setStage2Sections(sections.map(s => ({ id: s.id, key: s.section_key, title: s.section_title, content: s.content })));
      }
    });
    // Check if recommendation already sent
    hasRecommendationForSession(sessionId).then((sent) => {
      if (sent) setTelegramSent(true);
    }).catch(() => {});
    // Restore vault documents
    getSessionDocuments(sessionId).then((docs) => {
      if (docs.length > 0) {
        setVaultDocuments(docs.map((d: SessionDocument) => ({
          id: d.drive_file_id,
          name: d.file_name,
          mimeType: d.mime_type || '',
          size: d.file_size || 0,
          viewUrl: d.view_url || '',
          downloadUrl: d.download_url || '',
          type: (d.document_type || 'other') as VaultDocument['type'],
          category: (d.category || 'other') as VaultDocument['category'],
          uploadedAt: d.created_at || new Date().toISOString(),
          parentFolderId: '',
        })));
      }
    }).catch(() => {});
  }, [sessionId, selectedSector]);

  // --- Company Selection ---
  const handleSelectCompany = useCallback((company: MasterCompany) => {
    setSelectedCompany(company);
    setSearchInput(company.company_name);
    setIsDropdownOpen(false);
  }, []);

  // --- Create Session → auto-create vault ---
  const handleCreateSession = async () => {
    if (!selectedCompany || !user?.email) return;
    const sector = selectedSector || financials?.sector || financials?.broad_sector || 'General';
    if (!selectedSector) setSelectedSector(sector);
    setIsCreatingSession(true);

    try {
      const newSession = await createPipelineSession({
        company_name: selectedCompany.company_name,
        company_nse_code: selectedCompany.nse_symbol ?? '',
        sector,
        created_by: user.email,
        selected_model: selectedModel,
      });
      setSessionId(newSession.session_id);
      setSession(newSession);
      setPipelineStatus('company_selected');

      // Auto-create vault immediately after session
      setVaultStatus('loading');
      await transitionPipelineStatus(newSession.session_id, 'vault_creating', 'company_selected');
      setPipelineStatus('vault_creating');

      const vaultResponse = await createVault(selectedCompany.nse_symbol ?? '', sector);
      const { folderId, folderUrl, documents } = processVaultResponse(vaultResponse);
      setVaultLink(folderUrl);
      setVaultId(folderId);
      setVaultDocuments(documents);
      setVaultStatus('success');

      // Persist vault data to session for restore on resume
      await updatePipelineOutput(newSession.session_id, {
        vault_folder_id: folderId,
        vault_folder_url: folderUrl,
      });

      // Automatically save the incoming documents since they are physically in the vault
      if (documents.length > 0) {
        await saveSessionDocuments(
          documents.map(d => ({
            session_id: newSession.session_id,
            drive_file_id: d.id,
            file_name: d.name,
            mime_type: d.mimeType,
            file_size: d.size,
            view_url: d.viewUrl,
            download_url: d.downloadUrl,
            document_type: d.type,
            category: d.category,
          }))
        ).catch(() => console.error('Initial vault document save skipped/failed'));
      }

      await transitionPipelineStatus(newSession.session_id, 'vault_ready', 'vault_creating');
      setPipelineStatus('vault_ready');

      // Add to recent sessions list
      setRecentSessions(prev => [{ ...newSession, pipeline_status: 'vault_ready' } as PipelineSession, ...prev]);

      toast.success('Vault created — choose your next step');

      // Background: summarize vault docs with Haiku so Stage 1/2 can use them as primary source.
      // Fire-and-forget — pipeline can start before this completes; stages will pick up cached briefing.
      if (documents.length > 0) {
        summarizeVaultDocuments(newSession.session_id, selectedCompany.company_name, selectedCompany.nse_symbol ?? '', documents)
          .then(briefing => {
            if (briefing) console.log('[Pipeline] Vault briefing cached:', briefing.length, 'chars');
          })
          .catch(err => console.warn('[Pipeline] Vault summarization failed:', err));
      }
    } catch (err) {
      toast.error(`Failed to start pipeline: ${err instanceof Error ? err.message : 'Unknown error'}`);
      setVaultStatus('error');
    } finally {
      setIsCreatingSession(false);
    }
  };

  // --- Generate Financial Model (vault already exists) ---
  const handleGenerateFinancialModel = async () => {
    if (!sessionId || !selectedCompany || !vaultId) return;
    const sector = selectedSector || session?.sector || 'General';

    try {
      await transitionPipelineStatus(sessionId, 'financial_model_generating', 'vault_ready');
      setPipelineStatus('financial_model_generating');
      setFinancialModelStatus('generating');

      toast.info('Generating financial model — this takes ~10 min...');
      const modelResult = await generateFinancialModel(
        selectedCompany.nse_symbol ?? '',
        selectedCompany.company_name,
        sector,
        vaultId
      );
      const storageResult = modelResult.storageUrl
        ? { fileUrl: modelResult.storageUrl, filePath: null, jsonFileUrl: null, jsonFilePath: null }
        : await mirrorFinancialModelToStorage(selectedCompany.nse_symbol ?? '');
      setFinancialModelStatus('success');
      setFinancialModelFileUrl(storageResult.fileUrl);
      toast.success(`Financial model generated: ${modelResult.fileName}`);

      // Persist the Supabase Storage URL so downstream services can fetch it reliably.
      await updatePipelineOutput(sessionId, {
        financial_model_file_url: storageResult.fileUrl,
        financial_model_json_url: storageResult.jsonFileUrl,
      });

      // Upload the generated model to the Google Drive Vault so the user can see it
      if (storageResult.fileUrl) {
        try {
          toast.info('Uploading model to Vault...');
          const fileRes = await fetch(storageResult.fileUrl);
          const blob = await fileRes.blob();
          const reader = new FileReader();
          reader.readAsDataURL(blob);
          reader.onloadend = async () => {
            try {
              const base64data = (reader.result as string).split(',')[1];
              const uploadedDoc = await uploadDocument(
                vaultId, 
                modelResult.fileName || `${selectedCompany.nse_symbol}_Model.xlsx`, 
                base64data
              );
              setVaultDocuments(prev => [...prev, uploadedDoc]);
              toast.success('Model added to Vault');
            } catch (uploadErr) {
              console.error('Failed to upload model to Drive:', uploadErr);
            }
          };
        } catch (e) {
          console.error('Failed to fetch model for Drive upload:', e);
        }
      }

      await transitionPipelineStatus(sessionId, 'vault_ready', 'financial_model_generating');
      setPipelineStatus('vault_ready');

      const updated = await getPipelineSession(sessionId);
      if (updated) setSession(updated);
    } catch (err) {
      toast.error(`Financial model failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
      setFinancialModelStatus('idle');
      try { await transitionPipelineStatus(sessionId, 'vault_ready', 'financial_model_generating'); } catch { /* ignore */ }
      setPipelineStatus('vault_ready');
    }
  };

  // --- Skip Financial Model → go straight to sector framework ---
  const handleSkipFinancialModel = () => {
    setFinancialModelStatus('skipped');
    handleRunStage0();
  };

  // --- Save documents to session ---
  const handleConfirmDocuments = useCallback(async () => {
    if (!sessionId || vaultDocuments.length === 0) return;

    try {
      await saveSessionDocuments(
        vaultDocuments.map(d => ({
          session_id: sessionId,
          drive_file_id: d.id,
          file_name: d.name,
          mime_type: d.mimeType,
          file_size: d.size,
          view_url: d.viewUrl,
          download_url: d.downloadUrl,
          document_type: d.type,
          category: d.category,
        }))
      );
      toast.success('Documents saved to session');
    } catch {
      // Documents may already be saved
    }
  }, [sessionId, vaultDocuments]);

  // --- Stage 0: Sector Framework ---
  const handleRunStage0 = useCallback(async (forceRegenerate = false) => {
    if (!sessionId || !session) return;
    const previousStatus = pipelineStatus;
    setIsRunning(true);
    try {
      if (pipelineStatus === 'vault_ready') {
        await handleConfirmDocuments();
      }
      await transitionPipelineStatus(sessionId, 'stage0_generating', pipelineStatus);
      setPipelineStatus('stage0_generating');

      const { framework, tokensUsed, cached } = await runStage0(
        session.company_name,
        session.company_nse_code,
        selectedSector || session?.sector || '',
        setProgress,
        stage0Prompts,
        forceRegenerate
      );

      setSectorFramework(framework);

      await updatePipelineOutput(sessionId, {
        sector_framework: framework,
        total_tokens_used: (session.total_tokens_used || 0) + tokensUsed,
      });

      await transitionPipelineStatus(sessionId, 'stage0_review', 'stage0_generating');
      setPipelineStatus('stage0_review');
      toast.success(cached ? 'Existing sector framework loaded' : 'Sector framework generated');
    } catch (err) {
      toast.error(`Framework generation failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
      setPipelineStatus(previousStatus);
      try { await transitionPipelineStatus(sessionId, previousStatus); } catch { /* ignore */ }
    } finally {
      setIsRunning(false);
      setProgress(null);
    }
  }, [sessionId, session, pipelineStatus, selectedSector, stage0Prompts, handleConfirmDocuments]);

  // --- Stage 1: Investment Thesis ---
  const handleRunStage1 = useCallback(async () => {
    if (!sessionId || !session) return;
    const previousStatus = pipelineStatus;
    setIsRunning(true);
    try {
      await transitionPipelineStatus(sessionId, 'stage1_generating', pipelineStatus);
      setPipelineStatus('stage1_generating');

      const { thesis, tokensUsed } = await runStage1(
        session.company_name,
        session.company_nse_code,
        selectedSector || session?.sector || '',
        financials ?? null,
        sectorFramework?.markdown || '',
        setProgress,
        stage1Prompts,
        sessionId,
      );

      setStage1Thesis(thesis);

      await updatePipelineOutput(sessionId, {
        thesis_output: thesis,
        total_tokens_used: (session.total_tokens_used || 0) + tokensUsed,
      });

      await transitionPipelineStatus(sessionId, 'stage1_review', 'stage1_generating');
      setPipelineStatus('stage1_review');
      toast.success('Investment thesis generated');
    } catch (err) {
      toast.error(`Thesis generation failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
      setPipelineStatus(previousStatus);
      try { await transitionPipelineStatus(sessionId, previousStatus); } catch { /* ignore */ }
    } finally {
      setIsRunning(false);
      setProgress(null);
    }
  }, [sessionId, session, pipelineStatus, financials, selectedSector, sectorFramework, stage1Prompts]);

  // --- Stage 2: Full Report ---
  const handleRunStage2 = useCallback(async () => {
    if (!sessionId || !session || !stage1Thesis) return;
    const previousStatus = pipelineStatus;
    setIsRunning(true);
    try {
      await transitionPipelineStatus(sessionId, 'stage2_generating', pipelineStatus);
      setPipelineStatus('stage2_generating');

      const { sections: reportSections, tokensUsed } = await runStage2(
        session.company_name,
        session.company_nse_code,
        selectedSector || session?.sector || '',
        financials ?? null,
        stage1Thesis,
        sectorFramework?.markdown || '',
        setProgress,
        stage2Prompts,
        sessionId,
      );

      await clearResearchSections(sessionId, 'stage2');
      const savedSections: Array<{ id?: string; key: string; title: string; content: string }> = [];
      for (let i = 0; i < reportSections.length; i++) {
        const saved = await saveResearchSection({
          session_id: sessionId,
          section_key: reportSections[i].key,
          section_title: reportSections[i].title,
          stage: 'stage2',
          content: reportSections[i].content,
          sort_order: i,
          tokens_used: Math.round(tokensUsed / reportSections.length),
        });
        savedSections.push({ id: saved.id, key: reportSections[i].key, title: reportSections[i].title, content: reportSections[i].content });
      }
      setStage2Sections(savedSections);

      await updatePipelineOutput(sessionId, {
        total_tokens_used: (session.total_tokens_used || 0) + tokensUsed,
      });

      await transitionPipelineStatus(sessionId, 'stage2_review', 'stage2_generating');
      setPipelineStatus('stage2_review');
      toast.success('Report generated');
    } catch (err) {
      toast.error(`Report generation failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
      setPipelineStatus(previousStatus);
      try { await transitionPipelineStatus(sessionId, previousStatus); } catch { /* ignore */ }
    } finally {
      setIsRunning(false);
      setProgress(null);
    }
  }, [sessionId, session, pipelineStatus, financials, selectedSector, stage1Thesis, sectorFramework, stage2Prompts]);

  // --- Approve Handlers ---
  const handleApprove = async (stage: 'stage0' | 'stage1' | 'stage2') => {
    if (!sessionId) return;
    const newStatus: PipelineStatus =
      stage === 'stage0' ? 'stage0_approved' :
      stage === 'stage1' ? 'stage1_approved' : 'stage2_approved';
    try {
      await transitionPipelineStatus(sessionId, newStatus, pipelineStatus);
      if (stage === 'stage2' && session) {
        const report = await createResearchReport({
          session_id: sessionId,
          user_email: user?.email || '',
          company_name: session.company_name,
          nse_symbol: session.company_nse_code,
        });
        setInitialReport(report);
      }
      setPipelineStatus(newStatus);
      const updated = await getPipelineSession(sessionId);
      if (updated) setSession(updated);
      toast.success(`${stage === 'stage0' ? 'Sector framework' : stage === 'stage1' ? 'Thesis' : 'Report'} approved`);
    } catch (err) {
      toast.error(`Approval failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
    }
  };

  // --- Edit Handlers (persist to DB) ---
  const handleEditFramework = async (newMarkdown: string) => {
    setSectorFramework(prev => prev ? { ...prev, markdown: newMarkdown } : null);
    if (!sessionId) return;
    try {
      await updatePipelineOutput(sessionId, {
        sector_framework: sectorFramework ? { ...sectorFramework, markdown: newMarkdown } : null,
      });
    } catch {
      toast.error('Failed to save framework edit');
    }
  };

  const handleEditThesis = async (newContent: string) => {
    setStage1Thesis(newContent);
    if (!sessionId) return;
    try {
      await updatePipelineOutput(sessionId, { thesis_output: newContent });
    } catch {
      toast.error('Failed to save thesis edit');
    }
  };

  const handleEditReportSection = async (index: number, newContent: string) => {
    const section = stage2Sections[index];
    if (!section) return;
    setStage2Sections(prev => prev.map((s, i) => i === index ? { ...s, content: newContent } : s));
    if (section.id) {
      try {
        await updateResearchSection(section.id, { content: newContent });
      } catch {
        toast.error('Failed to save section edit');
      }
    }
  };

  const handlePublish = async () => {
    if (!sessionId) return;
    try {
      await transitionPipelineStatus(sessionId, 'published', pipelineStatus);
      setPipelineStatus('published');
      toast.success('Report published!');
    } catch (err) {
      toast.error(`Publish failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
    }
  };

  const handleUnpublish = async () => {
    if (!sessionId) return;
    try {
      const report = await getReportBySession(sessionId);
      if (!report || !report.report_id) {
        throw new Error('Research report record not found. Please ensure the report has been generated.');
      }
      await unpublishReport(report.report_id);
      await transitionPipelineStatus(sessionId, 'stage2_approved', pipelineStatus);
      setPipelineStatus('stage2_approved');
      toast.success('Report reverted to draft.');
    } catch (err) {
      toast.error(`Unpublish failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
    }
  };

  // --- Send Recommendation to Telegram from published report ---
  const handleSendTelegramRecommendation = async () => {
    if (!sessionId) return;
    setTelegramSending(true);
    try {
      const report = await getReportBySession(sessionId);
      if (!report) throw new Error('Report not found');

      const plan = report.plan;
      if (!plan) {
        toast.error('No plan assigned to this report. Unpublish, select a plan, and re-publish.');
        setTelegramSending(false);
        return;
      }

      // Extract data from report record (cs_ prefixed custom columns)
      // These columns may contain paragraphs — extract first number found
      const extractFirstNumber = (v: unknown): number | null => {
        if (v == null) return null;
        const s = String(v);
        // Match numbers like 1234, 1,234, 1234.56, ₹1,234.56
        const match = s.match(/[\d,]+\.?\d*/);
        if (!match) return null;
        const n = parseFloat(match[0].replace(/,/g, ''));
        return isNaN(n) ? null : n;
      };

      const rawRating = String(report.cs_rating || '').toUpperCase();
      const rating: RecommendationRating = rawRating.includes('SELL') ? 'SELL' : 'BUY';
      const cmp = extractFirstNumber(report.cs_current_market_price);
      const targetPrice = extractFirstNumber(report.cs_target_price);

      if (!targetPrice) {
        toast.error('Target price not found in report sections');
        setTelegramSending(false);
        return;
      }

      await createRecommendation({
        company_name: session?.company_name || '',
        nse_symbol: session?.company_nse_code || '',
        rating,
        cmp,
        target_price: targetPrice,
        validity_type: '1_year',
        validity_date: null,
        plans: [plan as PlanId],
        trade_notes: null,
        report_file_url: report.pptx_pdf_file_url || report.pptx_file_url || null,
        session_id: sessionId,
        send_telegram: true,
        send_push: sendPush,
        created_by: user?.email || null,
        pdf_file_id: null,
      });

      setTelegramSent(true);
      toast.success('Recommendation sent to Telegram!');
    } catch (err) {
      toast.error(`Failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
    } finally {
      setTelegramSending(false);
    }
  };

  // --- Resume a recent session ---
  const handleResumeSession = (s: PipelineSession) => {
    setSessionId(s.session_id);
    setSelectedCompany({
      company_id: 0,
      company_name: s.company_name,
      nse_symbol: s.company_nse_code,
      isin: null,
      bse_code: null,
      date_of_listing: null,
      paid_up_value: null,
      face_value: null,
      created_at: '',
      accord_code: null,
      google_code: null,
      bloomberg_ticker: null,
      yahoo_code: null,
      modified_at: null,
    });
    setSearchInput(s.company_name);
    setSelectedSector(s.sector ?? '');
    setTelegramSent(false);
    setTelegramSending(false);
  };

  // --- Delete session ---
  const handleDeleteSession = async (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (!confirm('Delete this session?')) return;
    try {
      await deletePipelineSession(id);
      setRecentSessions(prev => prev.filter(s => s.session_id !== id));
      if (sessionId === id) {
        setSessionId(null);
        setSession(null);
        setPipelineStatus('company_selected');
      }
      toast.success('Session deleted');
    } catch {
      toast.error('Failed to delete session');
    }
  };

  // --- Reset to home ---
  const handleBackToHome = () => {
    setSessionId(null);
    setSession(null);
    setSelectedCompany(null);
    setSearchInput('');
    setSelectedSector('');
    setPipelineStatus('company_selected');
    setVaultStatus('idle');
    setVaultDocuments([]);
    setSectorFramework(null);
    setStage1Thesis('');
    setStage2Sections([]);
    setFinancialModelStatus('idle');
  };

  // Computed
  const currentStage = getStageNumber(pipelineStatus);
  const hasSession = !!sessionId;

  // ========================
  // RENDER
  // ========================

  return (
    <div className="min-h-screen bg-canvas">
      {/* ==================== HEADER ==================== */}
      <header className="sticky top-0 z-30 bg-white/95 backdrop-blur-sm border-b border-neutral-200/80">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 h-14 flex items-center justify-between">
          <div className="flex items-center gap-3">
            {hasSession && (
              <button onClick={handleBackToHome} className="p-2 rounded-lg hover:bg-neutral-100 transition-colors text-neutral-400 hover:text-neutral-600">
                <X className="h-4 w-4" />
              </button>
            )}
            <div className="flex items-center gap-3">
              <div className="h-8 w-8 rounded-lg bg-accent-600 flex items-center justify-center shadow-sm shadow-accent-200">
                <Sparkles className="h-4 w-4 text-white" />
              </div>
              <div>
                <h1 className="text-sm font-semibold text-neutral-900 leading-tight">Research Pipeline</h1>
                {!hasSession && <p className="text-xs text-neutral-400 leading-tight">Powered by Claude + Web Search</p>}
              </div>
            </div>
          </div>

          {hasSession && session && (
            <div className="flex items-center gap-4">
              <div className="hidden md:flex items-center gap-2 text-xs">
                <div className="flex items-center gap-2 text-neutral-700 font-medium">
                  <Building2 className="h-3.5 w-3.5 text-neutral-400" />
                  {session.company_name}
                </div>
                <span className="text-neutral-300">|</span>
                <span className="text-neutral-500 font-mono text-xs">{session.company_nse_code}</span>
                {session.sector && (
                  <>
                    <span className="text-neutral-300">|</span>
                    <span className="text-neutral-500">{session.sector}</span>
                  </>
                )}
              </div>
              <StatusBadge status={pipelineStatus} />
            </div>
          )}
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 sm:px-6 py-6">
        {/* ==================== HOME SCREEN ==================== */}
        {!hasSession && (
          <div className="animate-fade-up">
            {/* Hero */}
            <div className="text-center mb-10 pt-4">
              <h2 className="text-2xl font-bold text-neutral-900 mb-2">Start a Research Pipeline</h2>
              <p className="text-sm text-neutral-500 max-w-md mx-auto">
                Generate institutional-grade equity research with Claude AI, real-time web data, and sector intelligence.
              </p>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-5 gap-6 max-w-5xl mx-auto">
              {/* Left Column: Search + Company Info */}
              <div className="lg:col-span-3 space-y-4">
                {/* Company Search */}
                <div className="bg-white rounded-2xl border border-neutral-200 shadow-sm p-5">
                  <label className="text-xs font-semibold text-neutral-400 uppercase tracking-wider mb-2 block">Search Company</label>
                  <div className="relative" ref={dropdownRef}>
                    <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 h-4.5 w-4.5 text-neutral-400" />
                    <Input
                      value={searchInput}
                      onChange={(e) => {
                        setSearchInput(e.target.value);
                        setIsDropdownOpen(true);
                        if (selectedCompany && e.target.value !== selectedCompany.company_name) {
                          setSelectedCompany(null);
                        }
                      }}
                      onFocus={() => searchInput.length >= 2 && setIsDropdownOpen(true)}
                      placeholder="Search by name or NSE symbol..."
                      className="pl-10 h-11 rounded-xl border-neutral-200 focus-visible:border-accent-400 focus-visible:ring-accent-500/20"
                    />

                    {isDropdownOpen && companies && companies.length > 0 && (
                      <div className="absolute z-50 w-full mt-2 bg-white border border-neutral-200 rounded-xl shadow-xl max-h-60 overflow-y-auto">
                        {companies.map((company) => (
                          <button
                            key={company.company_id}
                            className="w-full text-left px-4 py-3 hover:bg-accent-50 border-b border-neutral-50 last:border-0 transition-colors"
                            onClick={() => handleSelectCompany(company)}
                          >
                            <span className="font-medium text-sm text-neutral-900">{company.company_name}</span>
                            {company.nse_symbol && (
                              <span className="ml-2 text-xs text-neutral-400 bg-neutral-100 px-2 py-0.5 rounded font-mono">{company.nse_symbol}</span>
                            )}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>

                  {/* Model selection */}
                  {selectedCompany && (
                    <div className="mt-3 flex items-center gap-3">
                      <label className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">Model</label>
                      <div className="flex gap-2">
                        {PIPELINE_MODELS.map(m => (
                          <button
                            key={m.id}
                            onClick={() => setSelectedModel(m.id)}
                            className={cn(
                              'px-3 py-2 rounded-lg text-xs font-medium transition-all',
                              selectedModel === m.id
                                ? 'bg-accent-600 text-white shadow-sm'
                                : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'
                            )}
                          >
                            {m.label}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {/* Company Info + Financial Snapshot */}
                {selectedCompany && (
                  <div className="bg-white rounded-2xl border border-neutral-200 shadow-sm overflow-hidden animate-fade-up">
                    <div className="p-5">
                      <div className="flex items-start justify-between mb-4">
                        <div>
                          <h3 className="text-lg font-semibold text-neutral-900">{selectedCompany.company_name}</h3>
                          <div className="flex items-center gap-2 mt-1">
                            <span className="text-xs font-mono text-neutral-500 bg-neutral-100 px-2 py-0.5 rounded">NSE: {selectedCompany.nse_symbol}</span>
                            {!sessionId ? (
                              <select
                                value={selectedSector}
                                onChange={(e) => setSelectedSector(e.target.value)}
                                className="text-xs text-accent-700 bg-accent-50 border border-neutral-200 px-2 py-0.5 rounded focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent-500/40 cursor-pointer"
                              >
                                <option value="">Select Sector...</option>
                                {SECTORS.map(s => (
                                  <option key={s} value={s}>{s}</option>
                                ))}
                              </select>
                            ) : selectedSector ? (
                              <span className="text-xs text-accent-600 bg-accent-50 px-2 py-0.5 rounded">{selectedSector}</span>
                            ) : null}
                          </div>
                        </div>
                        <Button
                          onClick={handleCreateSession}
                          disabled={isCreatingSession}
                          className="h-10 px-5 rounded-xl bg-accent-600 hover:bg-accent-700 text-white font-medium shadow-sm"
                        >
                          {isCreatingSession ? (
                            <><Loader2 className="h-4 w-4 mr-2 animate-spin" /> Creating...</>
                          ) : (
                            <><Play className="h-4 w-4 mr-2" /> Begin Pipeline</>
                          )}
                        </Button>
                      </div>

                      <div className="grid grid-cols-3 sm:grid-cols-6 gap-2">
                        <MetricCard label="Market Cap" value={financials ? formatCurrency(financials.market_cap) : '-'} />
                        <MetricCard label="CMP" value={financials ? formatCurrency(financials.current_price) : '-'} />
                        <MetricCard label="P/E (TTM)" value={financials?.pe_ttm?.toFixed(1) ?? '-'} />
                        <MetricCard label="ROE" value={financials ? formatPercent(financials.roe) : '-'} />
                        <MetricCard label="ROCE" value={financials ? formatPercent(financials.roce) : '-'} />
                        <MetricCard label="EBITDA M%" value={financials ? formatPercent(financials.ebitda_margin_ttm) : '-'} />
                      </div>
                    </div>
                  </div>
                )}
              </div>

              {/* Right Column: Recent Sessions */}
              <div className="lg:col-span-2">
                <div className="bg-white rounded-2xl border border-neutral-200 shadow-sm overflow-hidden">
                  <div className="px-5 py-3 border-b border-neutral-100 flex items-center gap-2">
                    <Clock className="h-3.5 w-3.5 text-neutral-400" />
                    <h3 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">Recent Pipelines</h3>
                  </div>
                  {recentSessions.length > 0 ? (
                    <div className="divide-y divide-neutral-100">
                      {recentSessions.map((s) => {
                        const st = (s.pipeline_status ?? 'company_selected') as PipelineStatus;
                        return (
                          <div
                            key={s.session_id}
                            onClick={() => handleResumeSession(s)}
                            className="group flex items-center justify-between px-5 py-3 hover:bg-accent-50/40 transition-all cursor-pointer"
                          >
                            <div className="min-w-0 flex-1">
                              <p className="text-sm font-medium text-neutral-800 truncate">{s.company_name}</p>
                              <div className="flex items-center gap-2 mt-1">
                                <span className="text-xs font-mono text-neutral-400">{s.company_nse_code}</span>
                                <span className="text-xs text-neutral-300">{new Date(s.created_at).toLocaleDateString()}</span>
                              </div>
                            </div>
                            <div className="flex items-center gap-2 shrink-0 ml-3">
                              <StatusBadge status={st} />
                              <button
                                onClick={(e) => handleDeleteSession(s.session_id, e)}
                                className="p-1 rounded text-neutral-200 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-all"
                              >
                                <Trash2 className="h-3 w-3" />
                              </button>
                              <ChevronRight className="h-3.5 w-3.5 text-neutral-300 group-hover:text-accent-500 transition-colors" />
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  ) : (
                    <div className="p-8 text-center">
                      <FileText className="h-5 w-5 text-neutral-300 mx-auto mb-2" />
                      <p className="text-xs text-neutral-400">No recent pipelines</p>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ==================== PIPELINE WORKFLOW ==================== */}
        {hasSession && (
          <div className="max-w-4xl mx-auto animate-fade-up space-y-5">
            {/* Progress Bar */}
            <PipelineProgressBar status={pipelineStatus} />

            {/* Live Progress Indicator */}
            {progress && (
              <div className="rounded-xl border border-neutral-200 bg-accent-50/50 p-4">
                <div className="flex items-center gap-3">
                  <Globe className="h-5 w-5 text-accent-600 animate-pulse shrink-0" />
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium text-accent-800">{progress.message}</p>
                    <div className="mt-2 h-1 rounded-full bg-accent-100 overflow-hidden">
                      <div
                        className="h-full bg-accent-500 rounded-full transition-all duration-700"
                        style={{ width: `${progress.percent}%` }}
                      />
                    </div>
                  </div>
                  <span className="text-xs font-semibold text-accent-600 tabular-nums">{progress.percent}%</span>
                </div>
              </div>
            )}

            {/* ===== STEP: Vault creating (auto, shown right after session start) ===== */}
            {(pipelineStatus === 'company_selected' || pipelineStatus === 'vault_creating') && (
              <div className="bg-white rounded-xl border border-neutral-200 p-8 text-center">
                <Spinner size="sm" className="mx-auto mb-3" />
                <p className="text-sm font-medium text-neutral-700">Creating your Drive vault...</p>
                <p className="text-xs text-neutral-400 mt-1">Fetching documents and setting up folder structure</p>
              </div>
            )}

            {/* ===== STEP: Vault & Documents ===== */}
            {currentStage >= 1 && pipelineStatus !== 'company_selected' && pipelineStatus !== 'vault_creating' && (
              <CollapsibleStage
                title="Vault & Documents"
                done={currentStage > 1}
                defaultOpen={currentStage <= 1}
                trailing={vaultLink ? (
                  <a href={vaultLink} target="_blank" rel="noopener noreferrer" className="text-xs text-accent-600 hover:text-accent-700 font-medium flex items-center gap-1">
                    Open Drive <ExternalLink className="h-3 w-3" />
                  </a>
                ) : undefined}
              >
                {vaultStatus === 'error' && (
                  <div className="text-center py-6">
                    <p className="text-sm text-red-600 mb-3">Vault creation failed</p>
                    <Button onClick={handleCreateSession} variant="outline" size="sm" className="rounded-lg">
                      <RefreshCw className="h-3.5 w-3.5 mr-2" /> Retry
                    </Button>
                  </div>
                )}

                {vaultStatus === 'success' && (
                  <>
                    {/* Financial Model row — shown when generated */}
                    {financialModelStatus === 'success' && financialModelFileUrl && (
                      <div className="mb-3 flex items-center gap-3 rounded-lg bg-accent-50 border border-accent-100 px-4 py-3">
                        <BarChart3 className="h-4 w-4 text-accent-600 shrink-0" />
                        <span className="text-xs font-medium text-accent-800 flex-1">Financial Model — stored and ready</span>
                        <a
                          href={financialModelFileUrl}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="text-xs text-accent-600 hover:text-accent-700 font-medium flex items-center gap-1"
                        >
                          Open <ExternalLink className="h-3 w-3" />
                        </a>
                      </div>
                    )}

                    {/* Financial Model generating spinner */}
                    {financialModelStatus === 'generating' && (
                      <div className="mb-3 flex items-center gap-3 rounded-lg bg-neutral-50 border border-neutral-200 px-4 py-3">
                        <Spinner size="sm" className="shrink-0" />
                        <div>
                          <p className="text-xs font-medium text-neutral-700">Generating financial model...</p>
                          <p className="text-xs text-neutral-400 mt-1">This takes ~10 minutes. You can leave this tab open.</p>
                        </div>
                      </div>
                    )}

                    {/* Documents list */}
                    {vaultDocuments.length > 0 ? (
                      <div className="max-h-48 overflow-y-auto rounded-lg border border-neutral-100 divide-y divide-neutral-50">
                        {vaultDocuments.map((doc) => (
                          <div key={doc.id} className="flex items-center gap-3 px-3 py-2 hover:bg-neutral-50 transition-colors">
                            <FileText className="h-3.5 w-3.5 text-neutral-300 shrink-0" />
                            <span className="text-xs text-neutral-700 truncate flex-1">{doc.name}</span>
                            {doc.size > 0 && <span className="text-xs text-neutral-400 tabular-nums">{(doc.size / 1024).toFixed(0)} KB</span>}
                            {doc.viewUrl && (
                              <a href={doc.viewUrl} target="_blank" rel="noopener noreferrer" className="text-neutral-300 hover:text-accent-500 transition-colors">
                                <ExternalLink className="h-3 w-3" />
                              </a>
                            )}
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="text-center py-4 rounded-lg border border-dashed border-neutral-200">
                        <p className="text-xs text-neutral-400">No documents found in vault</p>
                      </div>
                    )}

                    <div className="flex items-center justify-between mt-3 pt-3 border-t border-neutral-100">
                      <Button variant="outline" size="sm" onClick={() => setIsUploadOpen(true)} className="rounded-lg text-xs">
                        <Upload className="h-3.5 w-3.5 mr-2" /> Upload
                      </Button>

                      {/* Actions for vault_ready and later stages */}
                      {['vault_ready', 'stage0_review', 'stage1_review', 'stage2_review', 'stage2_approved', 'published'].includes(pipelineStatus as string) && financialModelStatus !== 'generating' && (
                        <div className="flex items-center gap-2 mt-4">
                          <Button
                            onClick={handleGenerateFinancialModel}
                            size="sm"
                            variant="outline"
                            className="rounded-lg text-xs border-accent-200 text-accent-700 hover:bg-accent-50"
                          >
                            <BarChart3 className="h-3.5 w-3.5 mr-2" />
                            {financialModelStatus === 'success' ? 'Regenerate Financial Model' : 'Generate Financial Model'}
                          </Button>
                          
                          {pipelineStatus === 'vault_ready' && (
                            <Button
                              onClick={handleSkipFinancialModel}
                              size="sm"
                              variant="ghost"
                              className="rounded-lg text-xs text-neutral-400 hover:text-neutral-600"
                            >
                              Skip Model
                            </Button>
                          )}
                        </div>
                      )}

                      {/* Two options at vault_ready */}
                      {pipelineStatus === 'vault_ready' && financialModelStatus !== 'generating' && (
                        <div className="flex items-center gap-2">
                          <Button
                            onClick={financialModelStatus === 'success' || financialModelStatus === 'skipped'
                              ? () => handleRunStage0()
                              : handleSkipFinancialModel}
                            disabled={isRunning}
                            size="sm"
                            className="rounded-lg bg-accent-600 hover:bg-accent-700 text-xs"
                          >
                            <Zap className="h-3.5 w-3.5 mr-2" />
                            {financialModelStatus === 'success' || financialModelStatus === 'skipped'
                              ? 'Generate Sector Framework'
                              : 'Skip & Start Research'}
                          </Button>
                        </div>
                      )}

                      {/* After model done — show start research button */}
                      {pipelineStatus === 'vault_ready' && financialModelStatus === 'generating' && (
                        <p className="text-xs text-neutral-400 italic">Model generating...</p>
                      )}

                      {/* Stage already past vault_ready */}
                      {currentStage > 1 && pipelineStatus !== 'vault_ready' && (
                        <span className="text-xs text-green-600 flex items-center gap-1 font-medium">
                          <Check className="h-3.5 w-3.5" /> Vault ready
                        </span>
                      )}
                    </div>
                  </>
                )}
              </CollapsibleStage>
            )}

            {/* ===== STAGE 0: Sector Framework ===== */}
            {currentStage >= 2 && (
              <StageBlock
                title="Sector Framework"
                meta={sectorFramework ? [
                  sectorFramework.version > 1 ? `v${sectorFramework.version}` : null,
                  sectorFramework.last_updated ? `Updated ${sectorFramework.last_updated}` : null,
                ].filter(Boolean).join(' · ') || undefined : undefined}
                done={currentStage > 2 || pipelineStatus === 'stage0_approved'}
                isActive={currentStage === 2}
                showPrompt={showPrompt === 'stage0'}
                onTogglePrompt={() => setShowPrompt(showPrompt === 'stage0' ? null : 'stage0')}
                actions={
                  pipelineStatus === 'stage0_review' ? (
                    <button
                      onClick={() => handleRunStage0(true)}
                      disabled={isRunning}
                      className="text-xs text-neutral-400 hover:text-accent-600 flex items-center gap-1 transition-colors"
                    >
                      <RefreshCw className="h-3 w-3" /> Refresh
                    </button>
                  ) : undefined
                }
                promptEditor={
                  <PromptEditor
                    stage="stage0"
                    title="Sector Framework Prompt"
                    defaultSystem={DEFAULT_PROMPTS.stage0.system}
                    defaultUser={DEFAULT_PROMPTS.stage0.user}
                    userEmail={user?.email}
                    onChange={setStage0Prompts}
                  />
                }
              >
                {currentStage === 2 && (
                  <StageReview
                    title={`${selectedSector || session?.sector || ''} Sector Framework`}
                    content={sectorFramework?.markdown || ''}
                    isApproved={currentStage > 2 || pipelineStatus === 'stage0_approved'}
                    isGenerating={pipelineStatus === 'stage0_generating'}
                    onApprove={() => handleApprove('stage0')}
                    onRegenerate={() => handleRunStage0(true)}
                    onEdit={handleEditFramework}
                    maxHeight={500}
                  />
                )}
              </StageBlock>
            )}

            {/* ===== STAGE 1: Investment Thesis ===== */}
            {(currentStage >= 3 || pipelineStatus === 'stage0_approved') && (
              <StageBlock
                title="Investment Thesis"
                done={currentStage > 3 || pipelineStatus === 'stage1_approved'}
                isActive={currentStage === 3 || pipelineStatus === 'stage0_approved' || pipelineStatus === 'stage1_approved'}
                showPrompt={showPrompt === 'stage1'}
                onTogglePrompt={() => setShowPrompt(showPrompt === 'stage1' ? null : 'stage1')}
                actions={
                  pipelineStatus === 'stage0_approved' ? (
                    <Button onClick={handleRunStage1} disabled={isRunning} size="sm" className="rounded-lg bg-accent-600 hover:bg-accent-700">
                      <Zap className="h-3.5 w-3.5 mr-2" /> Generate
                    </Button>
                  ) : undefined
                }
                promptEditor={
                  <PromptEditor
                    stage="stage1"
                    title="Investment Thesis Prompt"
                    defaultSystem={DEFAULT_PROMPTS.stage1.system}
                    defaultUser={DEFAULT_PROMPTS.stage1.user}
                    userEmail={user?.email}
                    onChange={setStage1Prompts}
                  />
                }
              >
                {(currentStage === 3 || pipelineStatus === 'stage0_approved') && (
                  <StageReview
                    title="Investment Thesis"
                    content={stage1Thesis}
                    isApproved={currentStage > 3 || pipelineStatus === 'stage1_approved'}
                    isGenerating={pipelineStatus === 'stage1_generating'}
                    onApprove={() => handleApprove('stage1')}
                    onRegenerate={handleRunStage1}
                    onEdit={handleEditThesis}
                    maxHeight={500}
                  />
                )}
              </StageBlock>
            )}

            {/* ===== STAGE 2: Research Report ===== */}
            {(currentStage >= 4 || pipelineStatus === 'stage1_approved') && (
              <StageBlock
                title="Research Report"
                meta={stage2Sections.length > 0 ? `${stage2Sections.length} sections` : undefined}
                done={pipelineStatus === 'stage2_approved' || pipelineStatus === 'published'}
                isActive={true}
                showPrompt={showPrompt === 'stage2'}
                onTogglePrompt={() => setShowPrompt(showPrompt === 'stage2' ? null : 'stage2')}
                actions={
                  pipelineStatus === 'stage1_approved' ? (
                    <Button onClick={handleRunStage2} disabled={isRunning} size="sm" className="rounded-lg bg-accent-600 hover:bg-accent-700">
                      <Zap className="h-3.5 w-3.5 mr-2" /> Generate Report
                    </Button>
                  ) : undefined
                }
                promptEditor={
                  <PromptEditor
                    stage="stage2"
                    title="Full Report Prompt"
                    defaultSystem={DEFAULT_PROMPTS.stage2.system}
                    defaultUser={DEFAULT_PROMPTS.stage2.user}
                    userEmail={user?.email}
                    onChange={setStage2Prompts}
                  />
                }
              >
                {/* Loading state */}
                {pipelineStatus === 'stage2_generating' && stage2Sections.length === 0 && (
                  <StageReview
                    title="Research Report"
                    content=""
                    isApproved={false}
                    isGenerating={true}
                    onApprove={() => {}}
                    onRegenerate={() => {}}
                  />
                )}

                {/* Report sections with sidebar navigation */}
                {stage2Sections.length > 0 && (
                  <div className="flex flex-col md:flex-row">
                    {/* Section sidebar */}
                    <div className="md:w-52 shrink-0 border-b md:border-b-0 md:border-r border-neutral-100 bg-neutral-50/50">
                      <div className="md:sticky md:top-20 p-2 md:p-3">
                        <p className="text-xs font-semibold text-neutral-400 uppercase tracking-wider px-2 py-2 hidden md:block">Sections</p>
                        <div className="flex md:flex-col gap-1 overflow-x-auto md:overflow-x-visible">
                          {stage2Sections.map((section, i) => (
                            <button
                              key={`${section.key}-${i}`}
                              onClick={() => setActiveReportTab(i)}
                              className={cn(
                                'flex items-center gap-2 px-3 py-2 rounded-lg text-xs font-medium whitespace-nowrap transition-all text-left w-full',
                                activeReportTab === i
                                  ? 'bg-accent-50 text-accent-700'
                                  : 'text-neutral-500 hover:text-neutral-700 hover:bg-neutral-100'
                              )}
                            >
                              <Hash className="h-3 w-3 shrink-0 opacity-40" />
                              <span className="truncate">{section.title}</span>
                            </button>
                          ))}
                        </div>
                      </div>
                    </div>

                    {/* Active section content */}
                    <div className="flex-1 min-w-0">
                      {stage2Sections[activeReportTab] && (
                        <StageReview
                          title={stage2Sections[activeReportTab].title}
                          content={stage2Sections[activeReportTab].content}
                          isApproved={pipelineStatus === 'stage2_approved' || pipelineStatus === 'published'}
                          isGenerating={pipelineStatus === 'stage2_generating'}
                          onApprove={() => handleApprove('stage2')}
                          onRegenerate={handleRunStage2}
                          onEdit={(c) => handleEditReportSection(activeReportTab, c)}
                          maxHeight={600}
                        />
                      )}
                    </div>
                  </div>
                )}
              </StageBlock>
            )}

            {/* ===== Post-Production & Publish ===== */}
            {(pipelineStatus === 'stage2_approved' || pipelineStatus === 'published') && session && sessionId && (
              <PostProductionPanel
                sessionId={sessionId}
                companyName={session.company_name}
                nseSymbol={session.company_nse_code}
                sector={selectedSector || session.sector || null}
                vaultId={vaultId}
                financialModelFileUrl={financialModelFileUrl}
                userEmail={user?.email || ''}
                stage2Sections={stage2Sections}
                initialReport={initialReport}
                onPublished={handlePublish}
              />
            )}

            {pipelineStatus === 'published' && (
              <div className="rounded-xl border border-green-200 bg-gradient-to-b from-green-50 to-white p-8 text-center flex flex-col items-center">
                <div className="inline-flex items-center justify-center h-14 w-14 rounded-2xl bg-green-100 mb-4">
                  <Check className="h-7 w-7 text-green-600" strokeWidth={2.5} />
                </div>
                <h3 className="text-xl font-semibold text-green-800 mb-1">Report Published</h3>
                <p className="text-sm text-green-600 mb-6">This research report is now live and visible to stakeholders.</p>

                {!telegramSent && (
                  <div className="flex items-center gap-2 mb-4">
                    <input
                      type="checkbox"
                      id="send-push-checkbox-pipeline"
                      checked={sendPush}
                      onChange={(e) => setSendPush(e.target.checked)}
                      className="rounded border-neutral-300 text-accent-600 focus:ring-accent-500/40 h-4 w-4 cursor-pointer"
                    />
                    <label htmlFor="send-push-checkbox-pipeline" className="text-xs text-neutral-600 font-medium select-none cursor-pointer">
                      Send Push Notification to App Subscribers
                    </label>
                  </div>
                )}

                <div className="flex items-center gap-3">
                  {!telegramSent ? (
                    <Button
                      onClick={handleSendTelegramRecommendation}
                      disabled={telegramSending}
                      className="bg-blue-600 hover:bg-blue-700 text-white"
                    >
                      {telegramSending ? (
                        <><Loader2 className="h-4 w-4 mr-2 animate-spin" /> Sending...</>
                      ) : (
                        <><Send className="h-4 w-4 mr-2" /> Send Recommendation to Telegram</>
                      )}
                    </Button>
                  ) : (
                    <div className="flex items-center gap-2 text-green-600 font-medium text-sm px-4 py-2 bg-green-50 rounded-lg border border-green-200">
                      <Check className="h-4 w-4" /> Recommendation Sent
                    </div>
                  )}

                  <Button
                    onClick={handleUnpublish}
                    variant="outline"
                    className="border-green-200 text-green-800 hover:bg-green-50"
                  >
                    <RefreshCw className="h-4 w-4 mr-2" />
                    Revert to Draft (Unpublish)
                  </Button>
                </div>
              </div>
            )}
          </div>
        )}
      </main>

      {/* Upload Dialog */}
      {vaultId && (
        <DocumentUploadDialog
          open={isUploadOpen}
          onOpenChange={setIsUploadOpen}
          folderId={vaultId}
          nseSymbol={session?.company_nse_code || ''}
          onUploadComplete={(doc) => {
            setVaultDocuments(prev => [...prev, doc]);
            toast.success(`${doc.name} uploaded`);
          }}
        />
      )}
    </div>
  );
}

// ========================
// Sub-components
// ========================

function StatusBadge({ status }: { status: PipelineStatus }) {
  const colorMap: Record<string, string> = {
    published: 'bg-green-50 text-green-700 border-green-200',
    review: 'bg-amber-50 text-amber-700 border-amber-200',
    generating: 'bg-accent-50 text-accent-700 border-accent-200',
    approved: 'bg-green-50 text-green-600 border-green-200',
  };

  const key = status === 'published' ? 'published' :
    status.includes('review') ? 'review' :
    status.includes('generating') ? 'generating' :
    status.includes('approved') ? 'approved' : '';

  return (
    <span className={cn(
      'text-xs font-semibold px-2 py-0.5 rounded-full border',
      colorMap[key] || 'bg-neutral-50 text-neutral-500 border-neutral-200'
    )}>
      {PIPELINE_STAGE_LABELS[status]}
    </span>
  );
}

function MetricCard({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-lg bg-neutral-50 border border-neutral-100 px-3 py-2">
      <p className="text-xs text-neutral-400 uppercase tracking-wider font-semibold mb-1">{label}</p>
      <p className="text-sm font-semibold text-neutral-900 tabular-nums">{value}</p>
    </div>
  );
}

// ========================
// StageBlock — single card per stage, prompt editor + content in one container
// ========================

interface StageBlockProps {
  title: string;
  meta?: string;
  done: boolean;
  isActive: boolean;
  showPrompt: boolean;
  onTogglePrompt: () => void;
  actions?: React.ReactNode;
  promptEditor: React.ReactNode;
  children?: React.ReactNode;
}

function StageBlock({ title, meta, done, isActive, showPrompt, onTogglePrompt, actions, promptEditor, children }: StageBlockProps) {
  // Collapsed — completed stages show as a compact bar
  if (done && !isActive) {
    return (
      <div className="flex items-center gap-3 rounded-xl border border-neutral-200 bg-white px-4 py-3">
        <span className="flex h-5 w-5 items-center justify-center rounded-full bg-green-100 text-green-700 shrink-0">
          <Check className="h-3 w-3" strokeWidth={3} />
        </span>
        <span className="text-sm font-medium text-neutral-600">{title}</span>
        {meta && <span className="text-xs text-neutral-400">{meta}</span>}
      </div>
    );
  }

  return (
    <div className="rounded-xl border border-neutral-200 bg-white shadow-sm overflow-hidden">
      {/* Header */}
      <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-100">
        <div className="flex items-center gap-3">
          <h2 className="text-sm font-semibold text-neutral-900">{title}</h2>
          {meta && <span className="text-xs text-neutral-400">{meta}</span>}
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={onTogglePrompt}
            className={cn(
              'text-xs flex items-center gap-1 px-2 py-1 rounded-md transition-colors',
              showPrompt
                ? 'text-accent-600 bg-accent-50'
                : 'text-neutral-400 hover:text-neutral-600 hover:bg-neutral-50'
            )}
          >
            {showPrompt ? <EyeOff className="h-3 w-3" /> : <Eye className="h-3 w-3" />}
            Prompt
          </button>
          {actions}
        </div>
      </div>

      {/* Prompt editor — slides in, no extra border */}
      {showPrompt && (
        <div className="border-b border-neutral-100">
          {promptEditor}
        </div>
      )}

      {/* Content — no extra padding, StageReview fills edge-to-edge */}
      {children}
    </div>
  );
}

// ========================
// CollapsibleStage — for vault/documents
// ========================

interface CollapsibleStageProps {
  title: string;
  done: boolean;
  defaultOpen: boolean;
  trailing?: React.ReactNode;
  children: React.ReactNode;
}

function CollapsibleStage({ title, done, defaultOpen, trailing, children }: CollapsibleStageProps) {
  const [isOpen, setIsOpen] = useState(defaultOpen);

  if (done && !isOpen) {
    return (
      <button
        onClick={() => setIsOpen(true)}
        className="w-full flex items-center gap-3 rounded-xl border border-neutral-200 bg-white px-4 py-3 hover:bg-neutral-50 transition-colors text-left"
      >
        <span className="flex h-5 w-5 items-center justify-center rounded-full bg-green-100 text-green-700 shrink-0">
          <Check className="h-3 w-3" strokeWidth={3} />
        </span>
        <span className="text-sm font-medium text-neutral-600 flex-1">{title}</span>
        {trailing}
        <ChevronDown className="h-3.5 w-3.5 text-neutral-400" />
      </button>
    );
  }

  return (
    <div className="rounded-xl border border-neutral-200 bg-white shadow-sm overflow-hidden">
      <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-100">
        <div className="flex items-center gap-3">
          {done && (
            <span className="flex h-5 w-5 items-center justify-center rounded-full bg-green-100 text-green-700 shrink-0">
              <Check className="h-3 w-3" strokeWidth={3} />
            </span>
          )}
          <h2 className="text-sm font-semibold text-neutral-900">{title}</h2>
        </div>
        <div className="flex items-center gap-2">
          {trailing}
          {done && (
            <button onClick={() => setIsOpen(false)} className="p-1 rounded hover:bg-neutral-100 text-neutral-400">
              <X className="h-3 w-3" />
            </button>
          )}
        </div>
      </div>
      <div className="p-4">
        {children}
      </div>
    </div>
  );
}

// ========================
// Path Selection Card
// ========================
