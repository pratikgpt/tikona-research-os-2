import { useEffect, useState, useCallback } from 'react';
import {
  CheckCircle2,
  XCircle,
  AlertCircle,
  RefreshCw,
  Clock,
  Zap,
  Server,
  GitBranch,
  BarChart3,
  FileText,
  Play,
  Pause,
  FolderOpen,
  Upload,
  Trash2,
  Send,
  Video,
  Mic,
  FileOutput,
  Brain,
  ChevronDown,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { supabase } from '@/lib/supabase';

// All traffic proxied through Vite (local) / Vercel rewrites (prod) to avoid CORS
const N8N_BASE            = '/proxy/n8n'; // → https://n8n.tikonacapital.com
const N8N_API_KEY         = import.meta.env.VITE_N8N_API_KEY || '';
const PPT_SERVICE_URL     = '/proxy/ppt'; // → 72.61.226.16:8501
const FINANCIAL_MODEL_URL = '/proxy/fm';  // → 72.61.226.16:8500

// ─── Known project webhooks ────────────────────────────────────────────────────
// These are the exact n8n workflows this project calls.
// workflowName must match the name in n8n exactly (used to look up the real workflow).

const PROJECT_WEBHOOKS: {
  label: string;
  webhook: string;
  description: string;
  icon: React.ComponentType<{ className?: string }>;
  usedIn: string;
}[] = [
  {
    label: 'Generate Financial Model',
    webhook: 'generate-financial-model',
    description: 'Builds Excel financial model for a company',
    icon: BarChart3,
    usedIn: 'Report Generator',
  },
  {
    label: 'Create Vault Folder',
    webhook: 'create-folder',
    description: 'Creates Google Drive folder for research session',
    icon: FolderOpen,
    usedIn: 'Report Generator',
  },
  {
    label: 'Upload Document',
    webhook: 'upload-document',
    description: 'Uploads research docs to Google Drive vault',
    icon: Upload,
    usedIn: 'Report Generator',
  },
  {
    label: 'Delete File',
    webhook: 'delete-file',
    description: 'Removes a file from Google Drive vault',
    icon: Trash2,
    usedIn: 'Report Generator',
  },
  {
    label: 'Ingest Document',
    webhook: 'ingest-document',
    description: 'Ingests uploaded doc into RAG/search index',
    icon: Brain,
    usedIn: 'Generate Research',
  },
  {
    label: 'Send Recommendation',
    webhook: 'send-recommendation',
    description: 'Sends stock recommendation to Telegram subscribers',
    icon: Send,
    usedIn: 'Recommendations',
  },
  {
    label: 'Convert to PDF',
    webhook: 'convert-to-pdf',
    description: 'Converts research report to PDF',
    icon: FileOutput,
    usedIn: 'Post Production',
  },
  {
    label: 'Generate Media Script',
    webhook: 'generate-media-script',
    description: 'Generates video/audio script from report',
    icon: FileText,
    usedIn: 'Post Production',
  },
  {
    label: 'Synthesize Podcast',
    webhook: 'synthesize-podcast',
    description: 'Creates audio narration from script',
    icon: Mic,
    usedIn: 'Post Production',
  },
  {
    label: 'Generate Video',
    webhook: 'generate-video',
    description: 'Renders final research video',
    icon: Video,
    usedIn: 'Post Production',
  },
];

// ─── Types ────────────────────────────────────────────────────────────────────

type ServiceStatus = 'checking' | 'healthy' | 'degraded' | 'down';

interface ServiceHealth {
  status: ServiceStatus;
  latencyMs: number | null;
  details: Record<string, unknown>;
  lastChecked: Date | null;
  error?: string;
}

interface N8nWorkflow {
  id: string;
  name: string;
  active: boolean;
  updatedAt: string;
  tags?: { id: string; name: string }[];
  // Populated after fetching full detail
  webhookPaths?: string[];
}

interface N8nExecution {
  id: string;
  workflowId: string;
  workflowName?: string;
  status: 'success' | 'error' | 'waiting' | 'running' | 'canceled';
  startedAt: string;
  stoppedAt?: string;
  finished: boolean;
}

interface PipelineStats {
  total: number;
  byStatus: Record<string, number>;
  last7Days: number;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function statusColor(s: ServiceStatus) {
  switch (s) {
    case 'healthy':  return 'bg-green-500';
    case 'degraded': return 'bg-amber-500';
    case 'down':     return 'bg-red-500';
    default:         return 'bg-neutral-300 animate-pulse';
  }
}

function statusText(s: ServiceStatus) {
  switch (s) {
    case 'healthy':  return 'Healthy';
    case 'degraded': return 'Degraded';
    case 'down':     return 'Down';
    default:         return 'Checking…';
  }
}

function statusTextColor(s: ServiceStatus) {
  switch (s) {
    case 'healthy':  return 'text-green-700';
    case 'degraded': return 'text-amber-700';
    case 'down':     return 'text-red-700';
    default:         return 'text-neutral-500';
  }
}

function execStatusColor(s: N8nExecution['status']) {
  switch (s) {
    case 'success':  return 'bg-green-100 text-green-700';
    case 'error':    return 'bg-red-100 text-red-700';
    case 'running':  return 'bg-accent-100 text-accent-700';
    case 'waiting':  return 'bg-amber-100 text-amber-700';
    case 'canceled': return 'bg-neutral-100 text-neutral-500';
    default:         return 'bg-neutral-100 text-neutral-500';
  }
}

function relativeTime(iso: string) {
  const diff = Date.now() - new Date(iso).getTime();
  const s = Math.floor(diff / 1000);
  if (s < 60) return `${s}s ago`;
  const m = Math.floor(s / 60);
  if (m < 60) return `${m}m ago`;
  const h = Math.floor(m / 60);
  if (h < 24) return `${h}h ago`;
  return `${Math.floor(h / 24)}d ago`;
}

function execDuration(start: string, stop?: string) {
  const ms = new Date(stop || Date.now()).getTime() - new Date(start).getTime();
  if (ms < 1000) return `${ms}ms`;
  if (ms < 60000) return `${(ms / 1000).toFixed(1)}s`;
  return `${Math.floor(ms / 60000)}m ${Math.floor((ms % 60000) / 1000)}s`;
}

// ─── Fetch helpers ─────────────────────────────────────────────────────────────

async function checkService(url: string): Promise<ServiceHealth> {
  const start = Date.now();
  try {
    const res = await fetch(url, { signal: AbortSignal.timeout(6000) });
    const latencyMs = Date.now() - start;
    if (!res.ok) return { status: 'degraded', latencyMs, details: {}, lastChecked: new Date(), error: `HTTP ${res.status}` };
    const data = await res.json().catch(() => ({}));
    return { status: 'healthy', latencyMs, details: data, lastChecked: new Date() };
  } catch (e) {
    return { status: 'down', latencyMs: null, details: {}, lastChecked: new Date(), error: e instanceof Error ? e.message : 'Unreachable' };
  }
}

async function fetchN8nWorkflows(): Promise<N8nWorkflow[]> {
  if (!N8N_API_KEY) return [];
  const headers = { 'X-N8N-API-KEY': N8N_API_KEY };
  try {
    // 1. Get workflow list
    const listRes = await fetch(`${N8N_BASE}/api/v1/workflows?limit=100`, {
      headers,
      signal: AbortSignal.timeout(8000),
    });
    if (!listRes.ok) return [];
    const listData = await listRes.json();
    const workflows: N8nWorkflow[] = listData.data ?? [];

    // 2. For each workflow, fetch full detail to get nodes (webhook paths)
    // Do in parallel but cap at 10 concurrent to be safe
    const enriched = await Promise.all(
      workflows.map(async (wf) => {
        try {
          const detailRes = await fetch(`${N8N_BASE}/api/v1/workflows/${wf.id}`, {
            headers,
            signal: AbortSignal.timeout(6000),
          });
          if (!detailRes.ok) return wf;
          const detail = await detailRes.json();
          const nodes: Record<string, unknown>[] = detail.nodes ?? [];
          // Extract path from all Webhook-type nodes
          const webhookPaths = nodes
            .filter(n => (n.type as string)?.toLowerCase().includes('webhook'))
            .map(n => {
              const params = n.parameters as Record<string, unknown> | undefined;
              return (params?.path as string) ?? '';
            })
            .filter(Boolean);
          return { ...wf, webhookPaths };
        } catch {
          return wf;
        }
      })
    );
    return enriched;
  } catch {
    return [];
  }
}

async function fetchN8nExecutions(workflowMap?: Map<string, string>): Promise<N8nExecution[]> {
  if (!N8N_API_KEY) return [];
  try {
    const res = await fetch(`${N8N_BASE}/api/v1/executions?limit=30&includeData=false`, {
      headers: { 'X-N8N-API-KEY': N8N_API_KEY },
      signal: AbortSignal.timeout(8000),
    });
    if (!res.ok) return [];
    const data = await res.json();
    return (data.data ?? []).map((e: Record<string, unknown>) => {
      // n8n v1 API: workflow info lives at e.workflowData OR e.workflow depending on version
      const wfData = (e.workflowData ?? e.workflow ?? {}) as Record<string, unknown>;
      const workflowId = String(wfData.id ?? e.workflowId ?? '');
      // Prefer map lookup (from full workflow list) over inline name, which is often missing
      const workflowName = workflowMap?.get(workflowId) ?? (wfData.name as string) ?? workflowId ?? 'Unknown';
      return {
        id: e.id,
        workflowId,
        workflowName,
        status: e.status ?? (e.finished ? 'success' : 'running'),
        startedAt: e.startedAt,
        stoppedAt: e.stoppedAt,
        finished: e.finished,
      };
    });
  } catch {
    return [];
  }
}

async function fetchPipelineStats(): Promise<PipelineStats> {
  try {
    const { data, error } = await supabase.from('research_sessions').select('status, created_at');
    if (error || !data) return { total: 0, byStatus: {}, last7Days: 0 };
    const sevenDaysAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
    const byStatus: Record<string, number> = {};
    let last7Days = 0;
    for (const row of data) {
      byStatus[row.status] = (byStatus[row.status] ?? 0) + 1;
      if (new Date(row.created_at) > sevenDaysAgo) last7Days++;
    }
    return { total: data.length, byStatus, last7Days };
  } catch {
    return { total: 0, byStatus: {}, last7Days: 0 };
  }
}

// ─── Sub-components ───────────────────────────────────────────────────────────

function ServiceCard({ label, icon: Icon, url, health }: {
  label: string;
  icon: React.ComponentType<{ className?: string }>;
  url: string;
  health: ServiceHealth;
}) {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <div className="bg-white rounded-xl border border-neutral-200 p-4 transition-all">
      <div 
        className="flex items-center justify-between cursor-pointer select-none group"
        onClick={() => setIsOpen(!isOpen)}
      >
        <div className="flex items-center gap-3">
          <div className="h-9 w-9 rounded-lg bg-neutral-50 border border-neutral-200 flex items-center justify-center group-hover:border-neutral-300 transition-colors">
            <Icon className="h-4 w-4 text-neutral-500" />
          </div>
          <div>
            <p className="text-sm font-semibold text-neutral-900">{label}</p>
            <p className="text-xs text-neutral-400 font-mono hidden sm:block">{url}</p>
          </div>
        </div>
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2">
            <div className={`h-2 w-2 rounded-full ${statusColor(health.status)}`} />
            <span className={`text-xs font-semibold ${statusTextColor(health.status)} hidden sm:block`}>
              {statusText(health.status)}
            </span>
          </div>
          <ChevronDown className={`h-4 w-4 text-neutral-400 transition-transform ${isOpen ? 'rotate-180' : ''}`} />
        </div>
      </div>

      {isOpen && (
        <div className="mt-4 pt-4 border-t border-neutral-100 animate-in fade-in slide-in-from-top-2 duration-200">
          <div className="grid grid-cols-2 gap-3">
            <div className="rounded-lg bg-neutral-50 px-3 py-2">
              <p className="text-xs text-neutral-400 mb-1">Latency</p>
              <p className="text-sm font-semibold text-neutral-900 tabular-nums">
                {health.latencyMs !== null ? `${health.latencyMs}ms` : '—'}
              </p>
            </div>
            <div className="rounded-lg bg-neutral-50 px-3 py-2">
              <p className="text-xs text-neutral-400 mb-1">Last checked</p>
              <p className="text-sm font-semibold text-neutral-900">
                {health.lastChecked ? relativeTime(health.lastChecked.toISOString()) : '—'}
              </p>
            </div>
          </div>

          {health.error && (
            <p className="mt-3 text-xs text-red-600 bg-red-50 rounded-lg px-3 py-2 font-mono break-all">
              {health.error}
            </p>
          )}

          {health.status === 'healthy' && Object.keys(health.details).length > 0 && (
            <div className="mt-3 space-y-1 border-t border-neutral-100 pt-3">
              {Object.entries(health.details)
                .filter(([k]) => !['status', 'ok'].includes(k))
                .slice(0, 4)
                .map(([k, v]) => (
                  <div key={k} className="flex justify-between text-xs">
                    <span className="text-neutral-400">{k}</span>
                    <span className="text-neutral-700 font-mono">{String(v)}</span>
                  </div>
                ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// Matches a project webhook to a workflow from n8n and shows its status
function WebhookCard({
  webhook,
  workflows,
  recentExecutions,
}: {
  webhook: typeof PROJECT_WEBHOOKS[0];
  workflows: N8nWorkflow[];
  recentExecutions: N8nExecution[];
}) {
  const [isOpen, setIsOpen] = useState(false);
  const Icon = webhook.icon;

  // Normalise a path to bare words for fuzzy matching (handles kebab/snake/camel)
  const normalise = (s: string) => s.toLowerCase().replace(/[-_]/g, '');
  const target = normalise(webhook.webhook);

  // Match by actual webhook path inside workflow nodes (handles multiple webhooks per workflow)
  const matched = workflows.find(w =>
    w.webhookPaths?.some(p => normalise(p) === target || normalise(p).includes(target) || target.includes(normalise(p)))
  );

  // Most recent execution for this workflow
  const lastExec = matched
    ? recentExecutions.find(e => e.workflowId === matched.id || e.workflowName === matched.name)
    : undefined;

  const status: 'active' | 'inactive' | 'unknown' = matched
    ? (matched.active ? 'active' : 'inactive')
    : 'unknown';

  return (
    <div className="bg-white rounded-xl border border-neutral-200 p-4 transition-all">
      <div 
        className="flex items-center justify-between cursor-pointer select-none group"
        onClick={() => setIsOpen(!isOpen)}
      >
        <div className="flex items-center gap-3 min-w-0">
          <div className={`h-9 w-9 rounded-lg flex items-center justify-center shrink-0 group-hover:opacity-80 transition-opacity ${
            status === 'active'   ? 'bg-green-50'   :
            status === 'inactive' ? 'bg-neutral-100' :
            'bg-amber-50'
          }`}>
            <Icon className={`h-4 w-4 ${
              status === 'active'   ? 'text-green-600'   :
              status === 'inactive' ? 'text-neutral-400' :
              'text-amber-500'
            }`} />
          </div>
          <div className="min-w-0">
            <p className="text-sm font-semibold text-neutral-900 truncate">{webhook.label}</p>
            <p className="text-xs text-neutral-400 truncate hidden sm:block">{webhook.usedIn}</p>
          </div>
        </div>
        <div className="flex items-center gap-3 ml-2 shrink-0">
          <span className={`inline-flex items-center gap-1 text-xs font-semibold px-2 py-0.5 rounded-md ${
            status === 'active'   ? 'bg-green-50 text-green-700'     :
            status === 'inactive' ? 'bg-neutral-100 text-neutral-500' :
            'bg-amber-50 text-amber-700'
          }`}>
            {status === 'active'   && <><Play   className="h-3 w-3" /> Active</>}
            {status === 'inactive' && <><Pause  className="h-3 w-3" /> Inactive</>}
            {status === 'unknown'  && <><AlertCircle className="h-3 w-3" /> Not found</>}
          </span>
          <ChevronDown className={`h-4 w-4 text-neutral-400 transition-transform ${isOpen ? 'rotate-180' : ''}`} />
        </div>
      </div>

      {isOpen && (
        <div className="mt-4 pt-4 border-t border-neutral-100 animate-in fade-in slide-in-from-top-2 duration-200 flex flex-col gap-3">
          <p className="text-xs text-neutral-600">{webhook.description}</p>

          {matched && (
            <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg px-3 py-2 truncate flex flex-col">
              <span className="text-neutral-400 mb-0.5">Workflow</span>
              <span className="font-medium text-neutral-700">{matched.name}</span>
            </p>
          )}

          <div className="flex items-center justify-between text-xs pt-1">
            <span className="text-neutral-400 font-mono truncate mr-2 bg-neutral-50 px-2 py-1 rounded">
              /webhook/{webhook.webhook}
            </span>
            {lastExec ? (
              <span className={`px-2 py-1 rounded-md font-semibold whitespace-nowrap ${execStatusColor(lastExec.status)}`}>
                {lastExec.status} · {relativeTime(lastExec.startedAt)}
              </span>
            ) : (
              <span className="text-neutral-400 italic">no recent runs</span>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

function ExecutionRow({ ex }: { ex: N8nExecution }) {
  return (
    <div className="flex items-center justify-between py-3 border-b border-neutral-100 last:border-0">
      <div className="flex items-center gap-3 min-w-0">
        <span className={`text-xs font-semibold px-2 py-0.5 rounded-md shrink-0 ${execStatusColor(ex.status)}`}>
          {ex.status}
        </span>
        <span className="text-sm text-neutral-900 truncate">{ex.workflowName ?? 'Unknown'}</span>
      </div>
      <div className="flex items-center gap-4 shrink-0 ml-4 text-xs text-neutral-400">
        <span className="font-mono">{execDuration(ex.startedAt, ex.stoppedAt)}</span>
        <span>{relativeTime(ex.startedAt)}</span>
      </div>
    </div>
  );
}

// ─── Main Page ────────────────────────────────────────────────────────────────

export default function SystemHealth() {
  const hasN8nKey = !!N8N_API_KEY;

  const [pptHealth,  setPptHealth]  = useState<ServiceHealth>({ status: 'checking', latencyMs: null, details: {}, lastChecked: null });
  const [fmHealth,   setFmHealth]   = useState<ServiceHealth>({ status: 'checking', latencyMs: null, details: {}, lastChecked: null });
  const [n8nHealth,  setN8nHealth]  = useState<ServiceHealth>({ status: 'checking', latencyMs: null, details: {}, lastChecked: null });

  const [workflows,   setWorkflows]   = useState<N8nWorkflow[]>([]);
  const [executions,  setExecutions]  = useState<N8nExecution[]>([]);
  const [n8nLoading,  setN8nLoading]  = useState(false);

  const [pipelineStats, setPipelineStats] = useState<PipelineStats | null>(null);
  const [lastRefresh,   setLastRefresh]   = useState<Date | null>(null);
  const [refreshing,    setRefreshing]    = useState(false);

  const checkServices = useCallback(async () => {
    setPptHealth(h  => ({ ...h, status: 'checking' }));
    setFmHealth(h   => ({ ...h, status: 'checking' }));
    setN8nHealth(h  => ({ ...h, status: 'checking' }));

    const [ppt, fm, n8n] = await Promise.all([
      checkService(`${PPT_SERVICE_URL}/health`),
      checkService(`${FINANCIAL_MODEL_URL}/health`),
      // n8n health also through proxy to avoid CORS
      checkService(`${N8N_BASE}/healthz`),
    ]);

    setPptHealth(ppt);
    setFmHealth(fm);
    setN8nHealth(n8n);
  }, []);

  const loadN8nData = useCallback(async () => {
    if (!hasN8nKey) return;
    setN8nLoading(true);
    const wfs = await fetchN8nWorkflows();
    // Build id→name map so executions can resolve names reliably
    const workflowMap = new Map(wfs.map(w => [w.id, w.name]));
    const exs = await fetchN8nExecutions(workflowMap);
    setWorkflows(wfs);
    setExecutions(exs);
    setN8nLoading(false);
  }, [hasN8nKey]);

  const loadPipelineStats = useCallback(async () => {
    const stats = await fetchPipelineStats();
    setPipelineStats(stats);
  }, []);

  const refresh = useCallback(async () => {
    setRefreshing(true);
    await Promise.all([checkServices(), loadN8nData(), loadPipelineStats()]);
    setLastRefresh(new Date());
    setRefreshing(false);
  }, [checkServices, loadN8nData, loadPipelineStats]);

  useEffect(() => {
    const timer = setTimeout(() => {
      refresh();
    }, 0);
    return () => clearTimeout(timer);
  }, [refresh]);

  useEffect(() => {
    const t = setInterval(() => refresh(), 60000);
    return () => clearInterval(t);
  }, [refresh]);

  const allServices = [
    { label: 'PPT Generation Service', icon: FileText,  url: '72.61.226.16:8501',       health: pptHealth  },
    { label: 'Financial Model Server',  icon: BarChart3, url: '72.61.226.16:8500',        health: fmHealth   },
    { label: 'n8n Automation',          icon: GitBranch, url: 'n8n.tikonacapital.com',    health: n8nHealth  },
  ];

  const healthyCount  = allServices.filter(s => s.health.status === 'healthy').length;
  const overallStatus: ServiceStatus =
    healthyCount === allServices.length ? 'healthy' :
    healthyCount === 0 ? 'down' : 'degraded';

  const activeWorkflows = workflows.filter(w => w.active).length;
  const successRate = executions.length
    ? Math.round((executions.filter(e => e.status === 'success').length / executions.length) * 100)
    : null;

  return (
    <div className="flex-1 overflow-auto bg-canvas p-7">
      {/* Header */}
      <div className="mb-6 flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-neutral-900">System Health</h1>
          <p className="text-sm text-neutral-500 mt-1">
            VPS services · n8n workflows · pipeline stats
            {lastRefresh && (
              <span className="ml-2 text-neutral-400">· Last updated {relativeTime(lastRefresh.toISOString())}</span>
            )}
          </p>
        </div>
        <Button variant="outline" size="sm" onClick={refresh} disabled={refreshing} className="gap-2">
          <RefreshCw className={`h-3.5 w-3.5 ${refreshing ? 'animate-spin' : ''}`} />
          Refresh
        </Button>
      </div>

      {/* Overall banner */}
      <div className={`mb-6 rounded-xl border px-5 py-4 flex items-center gap-4 ${
        overallStatus === 'healthy'  ? 'bg-green-50 border-green-200'   :
        overallStatus === 'degraded' ? 'bg-amber-50 border-amber-200'   :
        overallStatus === 'down'     ? 'bg-red-50   border-red-200'     :
        'bg-neutral-50 border-neutral-200'
      }`}>
        {overallStatus === 'healthy'  && <CheckCircle2 className="h-5 w-5 text-green-600 shrink-0" />}
        {overallStatus === 'degraded' && <AlertCircle  className="h-5 w-5 text-amber-600 shrink-0" />}
        {overallStatus === 'down'     && <XCircle      className="h-5 w-5 text-red-600   shrink-0" />}
        <div>
          <p className={`text-sm font-semibold ${
            overallStatus === 'healthy'  ? 'text-green-800'  :
            overallStatus === 'degraded' ? 'text-amber-800'  :
            overallStatus === 'down'     ? 'text-red-800'    : 'text-neutral-600'
          }`}>
            {overallStatus === 'healthy'  && `All ${allServices.length} services operational`}
            {overallStatus === 'degraded' && `${healthyCount} of ${allServices.length} services healthy`}
            {overallStatus === 'down'     && 'All services unreachable'}
          </p>
          <p className="text-xs text-neutral-500 mt-1">Auto-refreshes every 60 seconds</p>
        </div>
      </div>

      {/* Summary stats */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
        {[
          { label: 'Services Up',        value: `${healthyCount}/${allServices.length}`, icon: Server,   color: 'text-green-600',   bg: 'bg-green-50'   },
          { label: 'Active Workflows',   value: hasN8nKey ? (n8nLoading ? '…' : String(activeWorkflows)) : 'No key',          icon: Play,     color: 'text-accent-600',  bg: 'bg-accent-50'  },
          { label: 'Exec Success Rate',  value: hasN8nKey ? (successRate !== null ? `${successRate}%` : '—') : 'No key',      icon: Zap,      color: 'text-amber-600',   bg: 'bg-amber-50'   },
          { label: 'Pipeline (7d)',      value: pipelineStats ? String(pipelineStats.last7Days) : '…',                         icon: BarChart3,color: 'text-neutral-600', bg: 'bg-neutral-50' },
        ].map(s => (
          <div key={s.label} className="bg-white rounded-xl border border-neutral-200 p-4 flex items-center gap-4">
            <div className={`h-10 w-10 rounded-lg ${s.bg} flex items-center justify-center shrink-0`}>
              <s.icon className={`h-5 w-5 ${s.color}`} />
            </div>
            <div>
              <p className="text-xs text-neutral-400">{s.label}</p>
              <p className="text-xl font-bold text-neutral-900 tabular-nums">{s.value}</p>
            </div>
          </div>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8">
        {/* Left Column: VPS Services */}
        <div>
          <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider mb-3">VPS Services</h2>
          <div className="flex flex-col gap-3">
            {allServices.map(s => <ServiceCard key={s.label} {...s} />)}
          </div>
        </div>

        {/* Right Column: n8n Workflows */}
        <div>
          <div className="flex items-center justify-between mb-3">
            <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">
              n8n Workflows — This Project
            </h2>
            {!hasN8nKey && (
              <span className="text-xs text-amber-600 bg-amber-50 border border-amber-200 px-2 py-0.5 rounded-md">
                Add VITE_N8N_API_KEY to see live status
              </span>
            )}
          </div>
          <div className="flex flex-col gap-3">
            {PROJECT_WEBHOOKS.map(wh => (
              <WebhookCard
                key={wh.webhook}
                webhook={wh}
                workflows={workflows}
                recentExecutions={executions}
              />
            ))}
          </div>
        </div>
      </div>

      {/* Research Pipeline stats */}
      {pipelineStats && (
        <div className="mb-8">
          <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider mb-3">Research Pipeline</h2>
          <div className="bg-white rounded-xl border border-neutral-200 p-5">
            <div className="flex items-center justify-between mb-4">
              <div>
                <p className="text-sm font-semibold text-neutral-900">Session Overview</p>
                <p className="text-xs text-neutral-400">{pipelineStats.total} total sessions in Supabase</p>
              </div>
              <span className="text-xs text-neutral-400 flex items-center gap-1">
                <Clock className="h-3 w-3" /> {pipelineStats.last7Days} in last 7 days
              </span>
            </div>
            <div className="flex flex-wrap gap-3">
              {Object.entries(pipelineStats.byStatus).map(([status, count]) => (
                <div key={status} className="flex items-center gap-2 bg-neutral-50 rounded-lg px-3 py-2 border border-neutral-100">
                  <div className={`h-2 w-2 rounded-full ${
                    status === 'completed'       ? 'bg-green-500'  :
                    status === 'document_review' ? 'bg-amber-500'  : 'bg-neutral-400'
                  }`} />
                  <span className="text-xs text-neutral-600 capitalize">{status.replace(/_/g, ' ')}</span>
                  <span className="text-xs font-bold text-neutral-900 tabular-nums">{count}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Recent executions (all workflows) */}
      {hasN8nKey && (
        <div>
          <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider mb-3">Recent Executions</h2>
          <div className="bg-white rounded-xl border border-neutral-200">
            <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-100">
              <p className="text-sm font-semibold text-neutral-900">Last 30 runs across all workflows</p>
              {successRate !== null && (
                <span className={`text-xs font-semibold px-2 py-0.5 rounded-md ${
                  successRate >= 80 ? 'bg-green-50 text-green-700' :
                  successRate >= 50 ? 'bg-amber-50 text-amber-700' : 'bg-red-50 text-red-700'
                }`}>
                  {successRate}% success
                </span>
              )}
            </div>
            <div className="px-5">
              {n8nLoading ? (
                <div className="py-8 text-center text-sm text-neutral-400">Loading…</div>
              ) : executions.length === 0 ? (
                <div className="py-8 text-center text-sm text-neutral-400">No executions found</div>
              ) : (
                executions.map(ex => <ExecutionRow key={ex.id} ex={ex} />)
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
