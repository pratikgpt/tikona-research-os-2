import { useState, useEffect, useCallback, useRef } from 'react';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Spinner } from '@/components/ui/spinner';
import { useAuth } from '@/contexts/AuthContext';
import { useCompanySearch, useCompanyFinancials } from '@/hooks/useCompanySearch';
import {
  createRecommendation,
  listRecommendations,
  closeRecommendation,
  deleteRecommendation,
  resendTelegram,
  sendPushNotification,
} from '@/lib/recommendations-api';
import {
  PLAN_OPTIONS,
  RATING_CONFIG,
  type Recommendation,
  type RecommendationRating,
  type PlanId,
  type ValidityType,
} from '@/types/recommendations';
import type { MasterCompany } from '@/types/database';
import { cn } from '@/lib/utils';
import {
  Search,
  TrendingUp,
  TrendingDown,
  Send,
  Clock,
  CheckCircle2,
  XCircle,
  MoreHorizontal,
  Trash2,
  RefreshCw,
  ExternalLink,
  FileText,
  X,
  BarChart3,
  Activity,
  Award,
  AlertCircle,
  RefreshCcw,
  Bell,
} from 'lucide-react';

// ========================
// Constants
// ========================

const RATINGS: RecommendationRating[] = ['BUY', 'SELL'];

// ========================
// Helpers
// ========================

function formatDate(iso: string): string {
  return new Date(iso).toLocaleDateString('en-IN', {
    day: '2-digit', month: 'short', year: 'numeric',
  });
}

function formatNum(n: number | null | undefined): string {
  if (n == null) return '—';
  return n.toLocaleString('en-IN', { maximumFractionDigits: 2 });
}

function validityLabel(rec: Recommendation): string {
  if (rec.validity_type === '1_year') {
    const expiry = new Date(rec.created_at);
    expiry.setFullYear(expiry.getFullYear() + 1);
    return expiry.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
  }
  return rec.validity_date ? formatDate(rec.validity_date) : '—';
}

// ========================
// Sub-components
// ========================

function RatingBadge({ rating }: { rating: RecommendationRating }) {
  const cfg = RATING_CONFIG[rating];
  return (
    <span className={cn(
      'inline-flex items-center gap-2 px-2 py-0.5 rounded-full text-xs font-semibold border',
      cfg.bg, cfg.color, cfg.border
    )}>
      <span className={cn('h-1.5 w-1.5 rounded-full', cfg.dot)} />
      {rating}
    </span>
  );
}

function PlanBadge({ planId }: { planId: PlanId }) {
  const plan = PLAN_OPTIONS.find(p => p.id === planId);
  return (
    <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-accent-50 text-accent-700 border border-accent-100">
      {plan?.label ?? planId}
    </span>
  );
}

function StatCard({
  label, value, sub, icon: Icon, color,
}: {
  label: string; value: string | number; sub?: string;
  icon: React.FC<{ className?: string }>; color: string;
}) {
  return (
    <div className="bg-white rounded-xl border border-neutral-200 p-6 flex items-center gap-4">
      <div className={cn('h-11 w-11 rounded-xl flex items-center justify-center shrink-0', color)}>
        <Icon className="h-5 w-5 text-white" />
      </div>
      <div>
        <p className="text-2xl font-bold text-neutral-900 tabular-nums">{value}</p>
        <p className="text-xs text-neutral-500 mt-1">{label}</p>
        {sub && <p className="text-xs text-neutral-400 mt-1">{sub}</p>}
      </div>
    </div>
  );
}

// ========================
// Donut Chart (SVG)
// ========================

function DonutChart({
  active, closed, total,
}: { active: number; closed: number; total: number }) {
  const r = 42;
  const circ = 2 * Math.PI * r;
  const activePct = total > 0 ? active / total : 0;
  const closedPct = total > 0 ? closed / total : 1;

  return (
    <div className="flex flex-col items-center gap-4">
      <div className="relative">
        <svg width={120} height={120} viewBox="0 0 100 100">
          <circle cx={50} cy={50} r={r} fill="none" stroke="#f1f5f9" strokeWidth={14} />
          {total > 0 && (
            <>
              <circle
                cx={50} cy={50} r={r} fill="none"
                stroke="#6366f1" strokeWidth={14}
                strokeDasharray={`${activePct * circ} ${circ}`}
                strokeLinecap="round"
                transform="rotate(-90 50 50)"
              />
              <circle
                cx={50} cy={50} r={r} fill="none"
                stroke="#10b981" strokeWidth={14}
                strokeDasharray={`${closedPct * circ} ${circ}`}
                strokeDashoffset={`${-activePct * circ}`}
                strokeLinecap="round"
                transform="rotate(-90 50 50)"
              />
            </>
          )}
        </svg>
        <div className="absolute inset-0 flex flex-col items-center justify-center">
          <span className="text-xl font-bold text-neutral-900">{total}</span>
          <span className="text-xs text-neutral-400">Total</span>
        </div>
      </div>
      <div className="flex gap-4 text-xs">
        <div className="flex items-center gap-2">
          <span className="h-2.5 w-2.5 rounded-full bg-accent-500" />
          <span className="text-neutral-600">Active <strong>{active}</strong></span>
        </div>
        <div className="flex items-center gap-2">
          <span className="h-2.5 w-2.5 rounded-full bg-green-500" />
          <span className="text-neutral-600">Closed <strong>{closed}</strong></span>
        </div>
      </div>
    </div>
  );
}

// ========================
// Horizontal Bar Chart
// ========================

function BarChart({ data }: { data: { label: string; value: number; max: number; color: string }[] }) {
  return (
    <div className="space-y-3">
      {data.map(({ label, value, max, color }) => (
        <div key={label}>
          <div className="flex justify-between text-xs mb-1">
            <span className="text-neutral-600 font-medium">{label}</span>
            <span className="text-neutral-500 tabular-nums">{value}%</span>
          </div>
          <div className="h-2 rounded-full bg-neutral-100 overflow-hidden">
            <div
              className={cn('h-full rounded-full transition-all duration-700', color)}
              style={{ width: `${max > 0 ? (value / max) * 100 : 0}%` }}
            />
          </div>
        </div>
      ))}
    </div>
  );
}

// ========================
// Action Menu
// ========================

function ActionMenu({
  rec,
  onClose,
  onDelete,
  onResend,
}: {
  rec: Recommendation;
  onClose: (id: string, success: boolean) => void;
  onDelete: (id: string) => void;
  onResend: (rec: Recommendation) => void;
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    if (open) document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [open]);

  return (
    <div className="relative" ref={ref}>
      <button
        onClick={() => setOpen(!open)}
        className="p-2 rounded-lg hover:bg-neutral-100 text-neutral-400 hover:text-neutral-600 transition-colors"
      >
        <MoreHorizontal className="h-4 w-4" />
      </button>
      {open && (
        <div className="absolute right-0 top-8 z-50 w-48 bg-white rounded-xl border border-neutral-200 shadow-lg py-1 text-sm">
          {rec.status === 'active' && (
            <>
              <button
                onClick={() => { onClose(rec.id, true); setOpen(false); }}
                className="w-full text-left px-3 py-2 hover:bg-green-50 text-green-700 flex items-center gap-2"
              >
                <CheckCircle2 className="h-3.5 w-3.5" /> Target Achieved
              </button>
              <button
                onClick={() => { onClose(rec.id, false); setOpen(false); }}
                className="w-full text-left px-3 py-2 hover:bg-red-50 text-red-700 flex items-center gap-2"
              >
                <XCircle className="h-3.5 w-3.5" /> Close (Not Hit)
              </button>
              <div className="border-t border-neutral-100 my-1" />
            </>
          )}
          {!rec.telegram_sent && (
            <button
              onClick={() => { onResend(rec); setOpen(false); }}
              className="w-full text-left px-3 py-2 hover:bg-accent-50 text-accent-700 flex items-center gap-2"
            >
              <Send className="h-3.5 w-3.5" /> Send to Telegram
            </button>
          )}
          {rec.report_file_url && (
            <a
              href={rec.report_file_url}
              target="_blank"
              rel="noopener noreferrer"
              className="w-full text-left px-3 py-2 hover:bg-neutral-50 text-neutral-700 flex items-center gap-2"
              onClick={() => setOpen(false)}
            >
              <ExternalLink className="h-3.5 w-3.5" /> Open Report
            </a>
          )}
          <div className="border-t border-neutral-100 my-1" />
          <button
            onClick={() => { onDelete(rec.id); setOpen(false); }}
            className="w-full text-left px-3 py-2 hover:bg-red-50 text-red-600 flex items-center gap-2"
          >
            <Trash2 className="h-3.5 w-3.5" /> Delete
          </button>
        </div>
      )}
    </div>
  );
}

// ========================
// Recommendations Table
// ========================

function RecTable({
  items,
  onClose,
  onDelete,
  onResend,
  emptyMsg,
}: {
  items: Recommendation[];
  onClose: (id: string, success: boolean) => void;
  onDelete: (id: string) => void;
  onResend: (rec: Recommendation) => void;
  emptyMsg: string;
}) {
  if (items.length === 0) {
    return (
      <div className="text-center py-12 text-neutral-400">
        <BarChart3 className="h-8 w-8 mx-auto mb-2 opacity-30" />
        <p className="text-sm">{emptyMsg}</p>
      </div>
    );
  }

  return (
    <div className="overflow-x-auto">
      <table className="w-full text-xs">
        <thead>
          <tr className="border-b border-neutral-100">
            {['Date', 'Company', 'Rating', 'CMP', 'Target', 'Upside', 'Validity', 'Plans', 'Telegram', 'Status', ''].map(h => (
              <th key={h} className="text-left px-3 py-3 text-xs font-semibold text-neutral-400 uppercase tracking-wider whitespace-nowrap">
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody className="divide-y divide-neutral-50">
          {items.map(rec => {
            const upside = rec.upside_pct;
            const isPositive = (upside ?? 0) >= 0;
            return (
              <tr key={rec.id} className="hover:bg-neutral-50/60 transition-colors">
                <td className="px-3 py-3 text-neutral-500 whitespace-nowrap">
                  {formatDate(rec.created_at)}
                </td>
                <td className="px-3 py-3">
                  <div className="font-semibold text-neutral-900">{rec.nse_symbol}</div>
                  <div className="text-xs text-neutral-400 truncate max-w-[120px]">{rec.company_name}</div>
                </td>
                <td className="px-3 py-3">
                  <RatingBadge rating={rec.rating} />
                </td>
                <td className="px-3 py-3 text-neutral-700 tabular-nums">
                  {rec.cmp != null ? `₹${formatNum(rec.cmp)}` : '—'}
                </td>
                <td className="px-3 py-3 font-semibold text-neutral-900 tabular-nums">
                  ₹{formatNum(rec.target_price)}
                </td>
                <td className="px-3 py-3 tabular-nums">
                  {upside != null ? (
                    <span className={cn(
                      'flex items-center gap-1 font-semibold',
                      isPositive ? 'text-green-600' : 'text-red-600'
                    )}>
                      {isPositive ? <TrendingUp className="h-3 w-3" /> : <TrendingDown className="h-3 w-3" />}
                      {upside > 0 ? '+' : ''}{upside}%
                    </span>
                  ) : '—'}
                </td>
                <td className="px-3 py-3 text-neutral-500 whitespace-nowrap">
                  {validityLabel(rec)}
                </td>
                <td className="px-3 py-3">
                  <div className="flex flex-wrap gap-1">
                    {rec.plans.map(p => <PlanBadge key={p} planId={p} />)}
                  </div>
                </td>
                <td className="px-3 py-3">
                  {rec.telegram_sent ? (
                    <span className="text-green-600 flex items-center gap-1">
                      <CheckCircle2 className="h-3.5 w-3.5" /> Sent
                    </span>
                  ) : (
                    <span className="text-neutral-400 flex items-center gap-1">
                      <Clock className="h-3.5 w-3.5" /> Pending
                    </span>
                  )}
                </td>
                <td className="px-3 py-3">
                  {rec.status === 'active' ? (
                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold bg-accent-50 text-accent-700 border border-accent-100">
                      <span className="h-1.5 w-1.5 rounded-full bg-accent-500 animate-pulse" />
                      Active
                    </span>
                  ) : (
                    <span className={cn(
                      'inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-semibold border',
                      rec.is_successful
                        ? 'bg-green-50 text-green-700 border-green-200'
                        : 'bg-neutral-100 text-neutral-500 border-neutral-200'
                    )}>
                      {rec.is_successful ? <CheckCircle2 className="h-3 w-3" /> : <XCircle className="h-3 w-3" />}
                      {rec.is_successful ? 'Hit Target' : 'Closed'}
                    </span>
                  )}
                </td>
                <td className="px-3 py-3">
                  <ActionMenu rec={rec} onClose={onClose} onDelete={onDelete} onResend={onResend} />
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ========================
// Main Page
// ========================

export default function Recommendations() {
  const { user } = useAuth();
  const [activeTab, setActiveTab] = useState<'create' | 'my' | 'performance'>('create');

  // ---- Create form state ----
  const [searchInput, setSearchInput] = useState('');
  const [debouncedSearch, setDebouncedSearch] = useState('');
  const [selectedCompany, setSelectedCompany] = useState<MasterCompany | null>(null);
  const [isDropdownOpen, setIsDropdownOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  const [rating, setRating] = useState<RecommendationRating>('BUY');
  const [cmp, setCmp] = useState('');
  const [targetPrice, setTargetPrice] = useState('');
  const [validityType, setValidityType] = useState<ValidityType>('1_year');
  const [validityDate, setValidityDate] = useState('');
  const [selectedPlans, setSelectedPlans] = useState<PlanId[]>([]);
  const [tradeNotes, setTradeNotes] = useState('');
  const [reportFileUrl, setReportFileUrl] = useState('');
  const [, setReportFile] = useState<File | null>(null);
  const [createdRec, setCreatedRec] = useState<Recommendation | null>(null);
  const [telegramSentState, setTelegramSentState] = useState(false);
  const [pushSentState, setPushSentState] = useState(false);
  const [isSendingTelegram, setIsSendingTelegram] = useState(false);
  const [isSendingPush, setIsSendingPush] = useState(false);
  const [isCreating, setIsCreating] = useState(false);

  const resetForm = () => {
    setSelectedCompany(null);
    setSearchInput('');
    setRating('BUY');
    setCmp('');
    setTargetPrice('');
    setValidityType('1_year');
    setValidityDate('');
    setSelectedPlans([]);
    setTradeNotes('');
    setReportFileUrl('');
    setReportFile(null);
    setCreatedRec(null);
    setTelegramSentState(false);
    setPushSentState(false);
  };

  // ---- My Recommendations state ----
  const [allRecs, setAllRecs] = useState<Recommendation[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [filterPlan, setFilterPlan] = useState('');
  const [filterFrom, setFilterFrom] = useState('');
  const [filterTo, setFilterTo] = useState('');

  const { data: companies } = useCompanySearch(debouncedSearch);
  const { data: financials } = useCompanyFinancials(selectedCompany);

  // debounce
  useEffect(() => {
    const t = setTimeout(() => setDebouncedSearch(searchInput), 300);
    return () => clearTimeout(t);
  }, [searchInput]);

  // close dropdown on outside click
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

  const loadRecs = useCallback(async () => {
    setIsLoading(true);
    try {
      const data = await listRecommendations({
        plan: filterPlan || undefined,
        fromDate: filterFrom || undefined,
        toDate: filterTo || undefined,
        createdBy: user?.email ?? undefined,
      });
      setAllRecs(data);
    } catch {
      toast.error('Failed to load recommendations');
    } finally {
      setIsLoading(false);
    }
  }, [filterPlan, filterFrom, filterTo, user?.email]);

  useEffect(() => {
    if (activeTab === 'my' || activeTab === 'performance') {
      loadRecs();
    }
  }, [activeTab, loadRecs]);

  // Auto-fill CMP when financials load for the selected company
  useEffect(() => {
    if (financials?.current_price) {
      setCmp(String(financials.current_price));
    }
  }, [financials?.current_price]);

  const handleFetchCmp = () => {
    const price = financials?.current_price;
    if (price) {
      setCmp(String(price));
      toast.success(`CMP updated: ₹${price}`);
    } else {
      toast.error('No price data available for this company');
    }
  };

  const handleSelectCompany = useCallback((company: MasterCompany) => {
    setSelectedCompany(company);
    setSearchInput(company.company_name);
    setIsDropdownOpen(false);
  }, []);

  const togglePlan = (planId: PlanId) => {
    setSelectedPlans(prev =>
      prev.includes(planId) ? prev.filter(p => p !== planId) : [...prev, planId]
    );
  };

  const handleCreate = async () => {
    if (!selectedCompany) return toast.error('Select a company');
    if (!targetPrice) return toast.error('Enter target price');
    if (selectedPlans.length === 0) return toast.error('Select at least one plan');
    if (validityType === 'custom' && !validityDate) return toast.error('Enter validity date');

    setIsCreating(true);
    try {
      // If file selected, convert to base64 — for now just use URL field
      // (full upload integration requires a vault folder; use URL paste)
      const newRec = await createRecommendation({
        company_name: selectedCompany.company_name,
        nse_symbol: selectedCompany.nse_symbol ?? '',
        rating,
        cmp: cmp ? parseFloat(cmp) : null,
        target_price: parseFloat(targetPrice),
        validity_type: validityType,
        validity_date: validityType === 'custom' ? validityDate : null,
        plans: selectedPlans,
        trade_notes: tradeNotes || null,
        report_file_url: reportFileUrl || null,
        session_id: null,
        send_telegram: false,
        send_push: false,
        created_by: user?.email ?? null,
      });

      toast.success('Recommendation created successfully');
      setCreatedRec(newRec);
      setTelegramSentState(false);
      setPushSentState(false);
    } catch (e) {
      toast.error(`Failed: ${e instanceof Error ? e.message : 'Unknown error'}`);
    } finally {
      setIsCreating(false);
    }
  };

  const handleClose = async (id: string, success: boolean) => {
    try {
      await closeRecommendation(id, success);
      toast.success(success ? 'Marked as Target Achieved' : 'Closed recommendation');
      loadRecs();
    } catch { toast.error('Failed to close'); }
  };

  const handleDelete = async (id: string) => {
    if (!confirm('Delete this recommendation?')) return;
    try {
      await deleteRecommendation(id);
      toast.success('Deleted');
      setAllRecs(prev => prev.filter(r => r.id !== id));
    } catch { toast.error('Failed to delete'); }
  };

  const handleResend = async (rec: Recommendation) => {
    try {
      await resendTelegram(rec);
      toast.success('Sent to Telegram');
      loadRecs();
    } catch (e) {
      toast.error(`Telegram failed: ${e instanceof Error ? e.message : 'Unknown error'}`);
    }
  };

  // ---- Derived data for performance ----
  const activeRecs = allRecs.filter(r => r.status === 'active');
  const closedRecs = allRecs.filter(r => r.status === 'closed');
  const successfulRecs = closedRecs.filter(r => r.is_successful === true);
  const successRate = closedRecs.length > 0
    ? Math.round((successfulRecs.length / closedRecs.length) * 100)
    : 0;
  const avgUpside = activeRecs.length > 0
    ? +(activeRecs.reduce((s, r) => s + (r.upside_pct ?? 0), 0) / activeRecs.length).toFixed(1)
    : 0;

  const planPerformance = PLAN_OPTIONS.map(plan => {
    const planRecs = allRecs.filter(r => r.plans.includes(plan.id));
    const planClosed = planRecs.filter(r => r.status === 'closed');
    const planSuccess = planClosed.filter(r => r.is_successful);
    const rate = planClosed.length > 0
      ? Math.round((planSuccess.length / planClosed.length) * 100)
      : 0;
    return { label: plan.label, id: plan.id, total: planRecs.length, closed: planClosed.length, rate };
  });

  const ratingBreakdown = RATINGS.map(r => ({
    rating: r,
    count: allRecs.filter(rec => rec.rating === r).length,
  })).filter(r => r.count > 0);

  const upcomingExpiry = activeRecs.filter(r => {
    const expiry = r.validity_type === '1_year'
      ? new Date(new Date(r.created_at).setFullYear(new Date(r.created_at).getFullYear() + 1))
      : r.validity_date ? new Date(r.validity_date) : null;
    if (!expiry) return false;
    const daysLeft = Math.ceil((expiry.getTime() - Date.now()) / 86400000);
    return daysLeft <= 30 && daysLeft >= 0;
  }).length;

  const tabs = [
    { id: 'create', label: 'Create Recommendation' },
    { id: 'my', label: 'My Recommendations' },
    { id: 'performance', label: 'Performance' },
  ] as const;

  const upside = cmp && targetPrice
    ? +(((parseFloat(targetPrice) - parseFloat(cmp)) / parseFloat(cmp)) * 100).toFixed(1)
    : null;

  return (
    <div className="min-h-screen bg-canvas">
      {/* Header */}
      <header className="sticky top-0 z-30 bg-white/95 backdrop-blur-sm border-b border-neutral-200/80">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 h-14 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="h-8 w-8 rounded-lg bg-accent-600 flex items-center justify-center shadow-sm">
              <Send className="h-4 w-4 text-white" />
            </div>
            <div>
              <h1 className="text-sm font-semibold text-neutral-900 leading-tight">Recommendations</h1>
              <p className="text-xs text-neutral-400 leading-tight">Send research to subscribers via Telegram</p>
            </div>
          </div>
          <div className="flex items-center gap-2 text-xs text-neutral-500">
            <span className="h-2 w-2 rounded-full bg-green-500" />
            {activeRecs.length} active
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 py-6">
        {/* Tabs */}
        <div className="flex gap-1 bg-white border border-neutral-200 rounded-xl p-1 mb-6 w-fit shadow-sm">
          {tabs.map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={cn(
                'px-4 py-2 rounded-lg text-sm font-medium transition-all',
                activeTab === tab.id
                  ? 'bg-accent-600 text-white shadow-sm'
                  : 'text-neutral-500 hover:text-neutral-800 hover:bg-neutral-50'
              )}
            >
              {tab.label}
            </button>
          ))}
        </div>

        {/* ========== CREATE TAB ========== */}
        {activeTab === 'create' && (
          <div className="max-w-3xl animate-fade-up">
            {createdRec ? (
              <div className="bg-white rounded-2xl border border-neutral-200 shadow-sm overflow-hidden p-6 text-center space-y-6">
                <div className="inline-flex items-center justify-center h-16 w-16 rounded-full bg-green-100 text-green-600 mx-auto">
                  <CheckCircle2 className="h-10 w-10" />
                </div>
                <div>
                  <h2 className="text-xl font-bold text-neutral-900">Recommendation Created!</h2>
                  <p className="text-sm text-neutral-500 mt-1">
                    The recommendation for <strong>{createdRec.company_name} ({createdRec.nse_symbol})</strong> has been saved.
                  </p>
                </div>

                {/* Recommendation Details Card */}
                <div className="max-w-md mx-auto bg-neutral-50 border border-neutral-200 rounded-xl p-4 text-left text-sm space-y-2">
                  <div className="flex justify-between border-b border-neutral-200/60 pb-2">
                    <span className="text-neutral-500 font-medium">Rating</span>
                    <RatingBadge rating={createdRec.rating} />
                  </div>
                  <div className="flex justify-between border-b border-neutral-200/60 pb-2">
                    <span className="text-neutral-500 font-medium">CMP / Target Price</span>
                    <span className="font-semibold text-neutral-800">
                      ₹{formatNum(createdRec.cmp)} / ₹{formatNum(createdRec.target_price)}
                    </span>
                  </div>
                  {createdRec.upside_pct != null && (
                    <div className="flex justify-between border-b border-neutral-200/60 pb-2">
                      <span className="text-neutral-500 font-medium">Expected Upside</span>
                      <span className="font-semibold text-green-600">+{createdRec.upside_pct}%</span>
                    </div>
                  )}
                  <div className="flex justify-between border-b border-neutral-200/60 pb-2">
                    <span className="text-neutral-500 font-medium">Plans</span>
                    <div className="flex flex-wrap gap-1 justify-end">
                      {createdRec.plans.map(p => <PlanBadge key={p} planId={p} />)}
                    </div>
                  </div>
                  <div className="flex justify-between pb-1">
                    <span className="text-neutral-500 font-medium">Validity</span>
                    <span className="text-neutral-800 font-medium">{validityLabel(createdRec)}</span>
                  </div>
                </div>

                {/* Action Buttons to Send to Telegram & Send Push */}
                <div className="flex flex-col sm:flex-row gap-3 justify-center max-w-md mx-auto">
                  <Button
                    onClick={async () => {
                      setIsSendingTelegram(true);
                      try {
                        await resendTelegram(createdRec);
                        setTelegramSentState(true);
                        toast.success('Sent to Telegram successfully!');
                      } catch (err) {
                        toast.error(`Telegram failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
                      } finally {
                        setIsSendingTelegram(false);
                      }
                    }}
                    disabled={telegramSentState || isSendingTelegram}
                    className={cn(
                      "flex-1 h-11 rounded-xl font-medium shadow-sm transition-all",
                      telegramSentState
                        ? "bg-green-50 hover:bg-green-50 text-green-700 border border-green-200 cursor-default"
                        : "bg-blue-600 hover:bg-blue-700 text-white"
                    )}
                  >
                    {isSendingTelegram ? (
                      <><Spinner size="sm" className="mr-2" /> Sending...</>
                    ) : telegramSentState ? (
                      <><CheckCircle2 className="h-4 w-4 mr-2" /> Sent to Telegram</>
                    ) : (
                      <><Send className="h-4 w-4 mr-2" /> Send to Telegram</>
                    )}
                  </Button>

                  <Button
                    onClick={async () => {
                      setIsSendingPush(true);
                      try {
                        await sendPushNotification(createdRec);
                        setPushSentState(true);
                        toast.success('Push Notification triggered successfully!');
                      } catch (err) {
                        toast.error(`Push failed: ${err instanceof Error ? err.message : 'Unknown error'}`);
                      } finally {
                        setIsSendingPush(false);
                      }
                    }}
                    disabled={pushSentState || isSendingPush}
                    className={cn(
                      "flex-1 h-11 rounded-xl font-medium shadow-sm transition-all",
                      pushSentState
                        ? "bg-green-50 hover:bg-green-50 text-green-700 border border-green-200 cursor-default"
                        : "bg-accent-600 hover:bg-accent-700 text-white"
                    )}
                  >
                    {isSendingPush ? (
                      <><Spinner size="sm" className="mr-2" /> Sending...</>
                    ) : pushSentState ? (
                      <><CheckCircle2 className="h-4 w-4 mr-2" /> Push Sent</>
                    ) : (
                      <><Bell className="h-4 w-4 mr-2" /> Send Push</>
                    )}
                  </Button>
                </div>

                {/* Navigation Buttons: My Recommendation / Next */}
                <div className="border-t border-neutral-100 pt-6 flex justify-center gap-4">
                  <Button
                    onClick={() => {
                      setActiveTab('my');
                      resetForm();
                    }}
                    variant="outline"
                    className="h-10 rounded-xl px-5 hover:bg-neutral-50 transition-colors"
                  >
                    My Recommendations
                  </Button>
                  <Button
                    onClick={() => {
                      resetForm();
                    }}
                    className="h-10 rounded-xl px-6 bg-neutral-900 hover:bg-neutral-800 text-white font-medium transition-all shadow-sm"
                  >
                    Next (Create Another)
                  </Button>
                </div>
              </div>
            ) : (
              <div className="bg-white rounded-2xl border border-neutral-200 shadow-sm overflow-hidden">
                <div className="px-6 py-4 border-b border-neutral-100">
                  <h2 className="text-sm font-semibold text-neutral-900">New Recommendation</h2>
                  <p className="text-xs text-neutral-500 mt-1">Will be sent to selected plan subscribers on Telegram</p>
                </div>

                <div className="p-6 space-y-5">
                  {/* Company search */}
                  <div>
                    <label className="block text-xs font-semibold text-neutral-500 uppercase tracking-wider mb-2">
                      Company *
                    </label>
                    <div className="relative" ref={dropdownRef}>
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-neutral-400" />
                      <Input
                        value={searchInput}
                        onChange={e => {
                          setSearchInput(e.target.value);
                          setIsDropdownOpen(true);
                          if (selectedCompany && e.target.value !== selectedCompany.company_name) {
                            setSelectedCompany(null);
                          }
                        }}
                        onFocus={() => searchInput.length >= 2 && setIsDropdownOpen(true)}
                        placeholder="Search by company name or NSE symbol..."
                        className="pl-9 h-10 rounded-xl"
                      />
                      {selectedCompany && (
                        <span className="absolute right-3 top-1/2 -translate-y-1/2 text-xs font-mono text-neutral-400 bg-neutral-100 px-2 py-0.5 rounded">
                          {selectedCompany.nse_symbol}
                        </span>
                      )}
                      {isDropdownOpen && companies && companies.length > 0 && (
                        <div className="absolute z-50 w-full mt-1 bg-white border border-neutral-200 rounded-xl shadow-xl max-h-56 overflow-y-auto">
                          {companies.map(c => (
                            <button
                              key={c.company_id}
                              className="w-full text-left px-4 py-3 hover:bg-accent-50 border-b border-neutral-50 last:border-0 transition-colors"
                              onClick={() => handleSelectCompany(c)}
                            >
                              <span className="font-medium text-sm text-neutral-900">{c.company_name}</span>
                              {c.nse_symbol && (
                                <span className="ml-2 text-xs text-neutral-400 bg-neutral-100 px-2 py-0.5 rounded font-mono">
                                  {c.nse_symbol}
                                </span>
                              )}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Rating */}
                  <div>
                    <label className="block text-xs font-semibold text-neutral-500 uppercase tracking-wider mb-2">
                      Rating *
                    </label>
                    <div className="flex flex-wrap gap-2">
                      {RATINGS.map(r => {
                        const cfg = RATING_CONFIG[r];
                        return (
                          <button
                            key={r}
                            onClick={() => setRating(r)}
                            className={cn(
                              'px-3 py-2 rounded-lg text-xs font-semibold border transition-all',
                              rating === r
                                ? `${cfg.bg} ${cfg.color} ${cfg.border} shadow-sm`
                                : 'bg-neutral-50 text-neutral-500 border-neutral-200 hover:border-neutral-300'
                            )}
                          >
                            {r}
                          </button>
                        );
                      })}
                    </div>
                  </div>

                  {/* CMP + Target */}
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-semibold text-neutral-500 uppercase tracking-wider mb-2">
                        CMP (₹)
                      </label>
                      <div className="relative">
                        <Input
                          type="number"
                          value={cmp}
                          onChange={e => setCmp(e.target.value)}
                          placeholder="e.g. 2800"
                          className="h-10 rounded-xl pr-10"
                        />
                        <button
                          type="button"
                          onClick={handleFetchCmp}
                          disabled={!selectedCompany || !financials}
                          title="Fill CMP from equity database"
                          className="absolute right-2 top-1/2 -translate-y-1/2 p-1 rounded-md text-neutral-400 hover:text-accent-600 hover:bg-accent-50 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
                        >
                          <RefreshCcw className="h-4 w-4" />
                        </button>
                      </div>
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-neutral-500 uppercase tracking-wider mb-2">
                        Target Price (₹) *
                      </label>
                      <div className="relative">
                        <Input
                          type="number"
                          value={targetPrice}
                          onChange={e => setTargetPrice(e.target.value)}
                          placeholder="e.g. 3500"
                          className="h-10 rounded-xl"
                        />
                        {upside != null && (
                          <span className={cn(
                            'absolute right-3 top-1/2 -translate-y-1/2 text-xs font-semibold',
                            upside >= 0 ? 'text-green-600' : 'text-red-600'
                          )}>
                            {upside > 0 ? '+' : ''}{upside}%
                          </span>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Validity */}
                  <div>
                    <label className="block text-xs font-semibold text-neutral-500 uppercase tracking-wider mb-2">
                      Validity *
                    </label>
                    <div className="flex gap-2 mb-2">
                      {([['1_year', '1 Year'], ['custom', 'Custom Date']] as [ValidityType, string][]).map(([v, l]) => (
                        <button
                          key={v}
                          onClick={() => setValidityType(v)}
                          className={cn(
                            'px-4 py-2 rounded-lg text-xs font-medium border transition-all',
                            validityType === v
                              ? 'bg-accent-600 text-white border-accent-600 shadow-sm'
                              : 'bg-neutral-50 text-neutral-600 border-neutral-200 hover:border-neutral-300'
                          )}
                        >
                          {l}
                        </button>
                      ))}
                    </div>
                    {validityType === '1_year' && (
                      <p className="text-xs text-neutral-400">
                        Expires on{' '}
                        <strong className="text-neutral-600">
                          {new Date(new Date().setFullYear(new Date().getFullYear() + 1))
                            .toLocaleDateString('en-IN', { day: '2-digit', month: 'long', year: 'numeric' })}
                        </strong>
                      </p>
                    )}
                    {validityType === 'custom' && (
                      <Input
                        type="date"
                        value={validityDate}
                        onChange={e => setValidityDate(e.target.value)}
                        min={new Date().toISOString().split('T')[0]}
                        className="h-10 rounded-xl w-48"
                      />
                    )}
                  </div>

                  {/* Plans */}
                  <div>
                    <label className="block text-xs font-semibold text-neutral-500 uppercase tracking-wider mb-2">
                      Share with Plans *
                    </label>
                    <div className="flex flex-wrap gap-2">
                      {PLAN_OPTIONS.map(plan => (
                        <button
                          key={plan.id}
                          onClick={() => togglePlan(plan.id)}
                          className={cn(
                            'flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-medium border transition-all',
                            selectedPlans.includes(plan.id)
                              ? 'bg-accent-600 text-white border-accent-600 shadow-sm'
                              : 'bg-neutral-50 text-neutral-600 border-neutral-200 hover:border-neutral-300'
                          )}
                        >
                          <span className={cn(
                            'h-3.5 w-3.5 rounded border-2 flex items-center justify-center transition-all',
                            selectedPlans.includes(plan.id)
                              ? 'bg-white/30 border-white'
                              : 'border-neutral-300'
                          )}>
                            {selectedPlans.includes(plan.id) && (
                              <svg className="h-2 w-2 text-white" viewBox="0 0 8 8" fill="none">
                                <path d="M1 4l2 2 4-4" stroke="currentColor" strokeWidth={1.5} strokeLinecap="round" strokeLinejoin="round" />
                              </svg>
                            )}
                          </span>
                          {plan.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Trade Notes */}
                  <div>
                    <label className="block text-xs font-semibold text-neutral-500 uppercase tracking-wider mb-2">
                      Trade Notes
                    </label>
                    <textarea
                      value={tradeNotes}
                      onChange={e => setTradeNotes(e.target.value)}
                      placeholder="Add rationale, key catalysts, risk factors..."
                      rows={3}
                      className="w-full px-3 py-3 text-sm border border-neutral-200 rounded-xl resize-none focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent-500/40 focus-visible:border-accent-400 text-neutral-800 placeholder:text-neutral-400"
                    />
                  </div>

                  {/* Report File URL */}
                  <div>
                    <label className="block text-xs font-semibold text-neutral-500 uppercase tracking-wider mb-2">
                      Research Report (URL)
                    </label>
                    <div className="relative">
                      <FileText className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-neutral-400" />
                      <Input
                        value={reportFileUrl}
                        onChange={e => setReportFileUrl(e.target.value)}
                        placeholder="Paste Google Drive or PDF link..."
                        className="pl-9 h-10 rounded-xl"
                      />
                    </div>
                    <p className="text-xs text-neutral-400 mt-1">
                      Paste the report URL from your vault or any accessible link
                    </p>
                  </div>

                  {/* Create Button only */}
                  <div className="flex items-center justify-end pt-2 border-t border-neutral-100">
                    <Button
                      onClick={handleCreate}
                      disabled={isCreating || !selectedCompany || !targetPrice || selectedPlans.length === 0}
                      className="h-10 px-6 rounded-xl bg-accent-600 hover:bg-accent-700 text-white font-medium shadow-sm transition-all"
                    >
                      {isCreating ? (
                        <><Spinner size="sm" className="mr-2 animate-spin" /> Creating...</>
                      ) : (
                        'Create Recommendation'
                      )}
                    </Button>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ========== MY RECOMMENDATIONS TAB ========== */}
        {activeTab === 'my' && (
          <div className="space-y-5 animate-fade-up">
            {/* Filters */}
            <div className="bg-white rounded-xl border border-neutral-200 p-4 flex flex-wrap items-center gap-3">
              <select
                value={filterPlan}
                onChange={e => setFilterPlan(e.target.value)}
                className="h-9 px-3 text-sm border border-neutral-200 rounded-lg text-neutral-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent-500/40"
              >
                <option value="">All Plans</option>
                {PLAN_OPTIONS.map(p => (
                  <option key={p.id} value={p.id}>{p.label}</option>
                ))}
              </select>
              <div className="flex items-center gap-2 text-sm text-neutral-500">
                <span>From</span>
                <Input
                  type="date"
                  value={filterFrom}
                  onChange={e => setFilterFrom(e.target.value)}
                  className="h-9 rounded-lg w-36"
                />
                <span>To</span>
                <Input
                  type="date"
                  value={filterTo}
                  onChange={e => setFilterTo(e.target.value)}
                  className="h-9 rounded-lg w-36"
                />
              </div>
              <Button
                onClick={loadRecs}
                size="sm"
                variant="outline"
                className="rounded-lg"
              >
                <RefreshCw className="h-3.5 w-3.5 mr-2" />
                Apply
              </Button>
              {(filterPlan || filterFrom || filterTo) && (
                <button
                  onClick={() => { setFilterPlan(''); setFilterFrom(''); setFilterTo(''); }}
                  className="text-xs text-neutral-400 hover:text-neutral-600 flex items-center gap-1"
                >
                  <X className="h-3.5 w-3.5" /> Clear
                </button>
              )}
              <span className="ml-auto text-xs text-neutral-400">{allRecs.length} total</span>
            </div>

            {isLoading ? (
              <div className="text-center py-16"><Spinner size="sm" className="mx-auto" /></div>
            ) : (
              <>
                {/* Active */}
                <div className="bg-white rounded-2xl border border-neutral-200 shadow-sm overflow-hidden">
                  <div className="px-5 py-4 border-b border-neutral-100 flex items-center gap-3">
                    <span className="h-2 w-2 rounded-full bg-accent-500 animate-pulse" />
                    <h3 className="text-sm font-semibold text-neutral-800">
                      Active Recommendations
                      <span className="ml-2 text-xs font-normal text-neutral-400">
                        ({activeRecs.length})
                      </span>
                    </h3>
                  </div>
                  <RecTable
                    items={activeRecs}
                    onClose={handleClose}
                    onDelete={handleDelete}
                    onResend={handleResend}
                    emptyMsg="No active recommendations"
                  />
                </div>

                {/* History */}
                <div className="bg-white rounded-2xl border border-neutral-200 shadow-sm overflow-hidden">
                  <div className="px-5 py-4 border-b border-neutral-100 flex items-center gap-3">
                    <Clock className="h-4 w-4 text-neutral-400" />
                    <h3 className="text-sm font-semibold text-neutral-800">
                      Trade History
                      <span className="ml-2 text-xs font-normal text-neutral-400">
                        ({closedRecs.length})
                      </span>
                    </h3>
                  </div>
                  <RecTable
                    items={closedRecs}
                    onClose={handleClose}
                    onDelete={handleDelete}
                    onResend={handleResend}
                    emptyMsg="No closed recommendations yet"
                  />
                </div>
              </>
            )}
          </div>
        )}

        {/* ========== PERFORMANCE TAB ========== */}
        {activeTab === 'performance' && (
          <div className="space-y-6 animate-fade-up">
            {isLoading ? (
              <div className="text-center py-16"><Spinner size="sm" className="mx-auto" /></div>
            ) : (
              <>
                {/* Stat cards */}
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                  <StatCard
                    label="Total Recommendations"
                    value={allRecs.length}
                    icon={BarChart3}
                    color="bg-accent-500"
                  />
                  <StatCard
                    label="Active Trades"
                    value={activeRecs.length}
                    sub={avgUpside > 0 ? `Avg upside: +${avgUpside}%` : undefined}
                    icon={Activity}
                    color="bg-sky-500"
                  />
                  <StatCard
                    label="Closed Trades"
                    value={closedRecs.length}
                    icon={CheckCircle2}
                    color="bg-green-500"
                  />
                  <StatCard
                    label="Success Rate"
                    value={`${successRate}%`}
                    sub={`${successfulRecs.length} of ${closedRecs.length} hit target`}
                    icon={Award}
                    color="bg-violet-500"
                  />
                </div>

                {upcomingExpiry > 0 && (
                  <div className="flex items-center gap-3 bg-amber-50 border border-amber-200 rounded-xl px-4 py-3">
                    <AlertCircle className="h-4 w-4 text-amber-600 shrink-0" />
                    <p className="text-sm text-amber-800">
                      <strong>{upcomingExpiry}</strong> recommendation{upcomingExpiry > 1 ? 's' : ''} expiring within 30 days
                    </p>
                  </div>
                )}

                <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                  {/* Donut: trade status */}
                  <div className="bg-white rounded-xl border border-neutral-200 p-6">
                    <h3 className="text-sm font-semibold text-neutral-800 mb-4">Trade Status</h3>
                    <DonutChart
                      active={activeRecs.length}
                      closed={closedRecs.length}
                      total={allRecs.length}
                    />
                  </div>

                  {/* Bar: success rate by plan */}
                  <div className="bg-white rounded-xl border border-neutral-200 p-6">
                    <h3 className="text-sm font-semibold text-neutral-800 mb-4">Success Rate by Plan</h3>
                    {planPerformance.every(p => p.closed === 0) ? (
                      <p className="text-xs text-neutral-400 text-center py-8">No closed trades yet</p>
                    ) : (
                      <BarChart
                        data={planPerformance.map(p => ({
                          label: p.label.split(' ').slice(0, 2).join(' '),
                          value: p.rate,
                          max: 100,
                          color: p.rate >= 60 ? 'bg-green-500' : p.rate >= 40 ? 'bg-amber-500' : 'bg-red-400',
                        }))}
                      />
                    )}
                  </div>

                  {/* Bar: rating breakdown */}
                  <div className="bg-white rounded-xl border border-neutral-200 p-6">
                    <h3 className="text-sm font-semibold text-neutral-800 mb-4">Rating Breakdown</h3>
                    {ratingBreakdown.length === 0 ? (
                      <p className="text-xs text-neutral-400 text-center py-8">No data</p>
                    ) : (
                      <BarChart
                        data={ratingBreakdown.map(r => ({
                          label: r.rating,
                          value: r.count,
                          max: Math.max(...ratingBreakdown.map(x => x.count)),
                          color: RATING_CONFIG[r.rating].dot,
                        }))}
                      />
                    )}
                  </div>
                </div>

                {/* Performance table by plan */}
                <div className="bg-white rounded-2xl border border-neutral-200 shadow-sm overflow-hidden">
                  <div className="px-5 py-4 border-b border-neutral-100">
                    <h3 className="text-sm font-semibold text-neutral-800">Performance by Plan</h3>
                  </div>
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b border-neutral-100 bg-neutral-50/50">
                        {['Plan', 'Total', 'Active', 'Closed', 'Successful', 'Success Rate'].map(h => (
                          <th key={h} className="text-left px-5 py-3 text-xs font-semibold text-neutral-400 uppercase tracking-wider">
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-neutral-50">
                      {planPerformance.map(plan => (
                        <tr key={plan.id} className="hover:bg-neutral-50/60">
                          <td className="px-5 py-4 font-medium text-neutral-800">{plan.label}</td>
                          <td className="px-5 py-4 text-neutral-600 tabular-nums">{plan.total}</td>
                          <td className="px-5 py-4 text-neutral-600 tabular-nums">{plan.total - plan.closed}</td>
                          <td className="px-5 py-4 text-neutral-600 tabular-nums">{plan.closed}</td>
                          <td className="px-5 py-4 text-green-600 tabular-nums font-medium">
                            {allRecs.filter(r => r.plans.includes(plan.id) && r.is_successful).length}
                          </td>
                          <td className="px-5 py-4">
                            <div className="flex items-center gap-3">
                              <div className="flex-1 h-1.5 rounded-full bg-neutral-100 overflow-hidden max-w-[80px]">
                                <div
                                  className={cn('h-full rounded-full', plan.rate >= 60 ? 'bg-green-500' : plan.rate >= 40 ? 'bg-amber-500' : 'bg-red-400')}
                                  style={{ width: `${plan.rate}%` }}
                                />
                              </div>
                              <span className={cn(
                                'text-sm font-semibold tabular-nums',
                                plan.rate >= 60 ? 'text-green-600' : plan.rate >= 40 ? 'text-amber-600' : plan.closed === 0 ? 'text-neutral-400' : 'text-red-600'
                              )}>
                                {plan.closed === 0 ? '—' : `${plan.rate}%`}
                              </span>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </div>
        )}
      </main>
    </div>
  );
}
