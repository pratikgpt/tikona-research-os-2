import { useState, useEffect, useCallback } from 'react';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import { cn } from '@/lib/utils';
import { fetchPptPlaceholders, savePptPlaceholders } from '@/lib/api';
import {
  Check,
  ChevronDown,
  ChevronUp,
  Loader2,
  RefreshCw,
  AlertCircle,
  FileEdit,
} from 'lucide-react';

// ── Placeholder field definitions ────────────────────────────────────────────

interface FieldDef {
  key: string;
  label: string;
  type: 'input' | 'textarea';
  hint?: string;
}

interface GroupDef {
  title: string;
  slide: string;
  fields: FieldDef[];
}

const PLACEHOLDER_GROUPS: GroupDef[] = [
  {
    title: 'Cover & Market Data',
    slide: 'Slide 1',
    fields: [
      { key: 'company_name', label: 'Company Name', type: 'input' },
      { key: 'nse_code', label: 'NSE Symbol', type: 'input' },
      { key: 'cmp', label: 'Current Market Price (CMP)', type: 'input' },
      { key: 'target', label: 'Target Price', type: 'input' },
      { key: 'm_cap', label: 'Market Cap', type: 'input' },
      { key: 'm_category', label: 'Market Category', type: 'input', hint: 'e.g. Large Cap / Mid Cap / Small Cap' },
      { key: 'saarthi_s', label: 'SAARTHI Score', type: 'input' },
      { key: 'tagline', label: 'Tagline', type: 'input' },
      { key: 'date', label: 'Report Date', type: 'input' },
    ],
  },
  {
    title: 'Investment Thesis',
    slide: 'Slides 1, 4',
    fields: [
      { key: 'investment_thesis_heading', label: 'Thesis Heading', type: 'input' },
      { key: 'investment_thesis', label: 'Thesis Summary', type: 'textarea' },
      { key: '1', label: 'Thesis Bullet 1', type: 'textarea' },
      { key: '2', label: 'Thesis Bullet 2', type: 'textarea' },
      { key: '3', label: 'Thesis Bullet 3', type: 'textarea' },
      { key: '4', label: 'Thesis Bullet 4', type: 'textarea' },
    ],
  },
  {
    title: 'Company Overview',
    slide: 'Slide 6',
    fields: [
      { key: 'COMPANY_OVERVIEW', label: 'Company Overview', type: 'textarea' },
    ],
  },
  {
    title: 'Industry Analysis',
    slide: 'Slide 5',
    fields: [
      { key: 'cell', label: 'Cell / Price Value', type: 'input' },
      { key: 'cell_cap', label: 'Small Cap Category', type: 'input' },
      { key: 'mod_cap', label: 'Market Cap Display', type: 'input' },
      { key: 'mod', label: 'Module Import Decline %', type: 'input' },
      { key: 'tar', label: 'Target Price ({{tar}})', type: 'input' },
      { key: 'tar_pr', label: 'Target Price ({{tar_pr}})', type: 'input' },
      { key: 'buy', label: 'Buy Price (CMP)', type: 'input' },
      { key: 'up', label: 'Upside %', type: 'input' },
      { key: ' date ', label: 'ALMM-II Effective Date', type: 'input' },
      { key: 'industry_structure', label: 'Industry Structure', type: 'textarea' },
      { key: 'key_industry', label: 'Key Industry Tailwinds', type: 'textarea' },
      { key: 'key_industry_risk', label: 'Key Industry Risks', type: 'textarea' },
      { key: 'industry_tailwinds', label: 'Industry Tailwinds', type: 'textarea' },
    ],
  },
  {
    title: 'Business Ideas',
    slide: 'Slide 7',
    fields: [
      { key: 'p1', label: 'Business Point 1', type: 'textarea' },
      { key: 'p2', label: 'Business Point 2', type: 'textarea' },
      { key: 'p4', label: 'Business Point 4', type: 'textarea' },
      { key: 'p5', label: 'Business Point 5', type: 'textarea' },
      { key: 'p6', label: 'Business Point 6', type: 'textarea' },
    ],
  },
  {
    title: 'Competitive Advantages',
    slide: 'Slide 8',
    fields: [
      { key: 'competitve_advantage_1', label: 'Advantage 1', type: 'textarea' },
      { key: 'competitve_advantage_2', label: 'Advantage 2', type: 'textarea' },
      { key: 'competitve_advantage_3', label: 'Advantage 3', type: 'textarea' },
      { key: 'competitve_advantage_4', label: 'Advantage 4', type: 'textarea' },
    ],
  },
  {
    title: 'Peer Comparison',
    slide: 'Slide 9',
    fields: [
      { key: 'peer_comparision', label: 'Peer Analysis (Full)', type: 'textarea' },
      { key: 'peer_para1', label: 'Peer Paragraph 1', type: 'textarea' },
      { key: 'peer_para2', label: 'Peer Paragraph 2', type: 'textarea' },
    ],
  },
  {
    title: 'Management',
    slide: 'Slide 10',
    fields: [
      { key: 'management_commentry_heading', label: 'Section Heading', type: 'input' },
      { key: 'management_content', label: 'Management Commentary', type: 'textarea' },
    ],
  },
  {
    title: 'Governance',
    slide: 'Slide 11',
    fields: [
      { key: 'indicators', label: 'Governance Indicators', type: 'textarea' },
    ],
  },
  {
    title: 'SAARTHI',
    slide: 'Slide 15',
    fields: [
      { key: 'saarthi_heading', label: 'SAARTHI Heading', type: 'input' },
      { key: 'saarthi_summary_heading', label: 'Summary Heading', type: 'input' },
      { key: 'saarthi_summary', label: 'SAARTHI Summary', type: 'textarea' },
      { key: 'saarthi_content', label: 'SAARTHI Detail', type: 'textarea' },
    ],
  },
  {
    title: 'Scenario Analysis',
    slide: 'Slide 16',
    fields: [
      { key: 'valuation_bear', label: 'Bear Case Target ({{valuation_bear}})', type: 'input' },
      { key: 'bear', label: 'Bear Target ({{bear}})', type: 'input' },
      { key: 'bear_p', label: 'Bear Probability', type: 'input', hint: 'e.g. 25%' },
      { key: 'bear_content', label: 'Bear Case Notes', type: 'textarea' },
      { key: 'base', label: 'Base Case Target ({{base}})', type: 'input' },
      { key: 'base_p', label: 'Base Probability', type: 'input', hint: 'e.g. 50%' },
      { key: 'base_content', label: 'Base Case Notes', type: 'textarea' },
      { key: 'valuation_bull', label: 'Bull Case Target ({{valuation_bull}})', type: 'input' },
      { key: 'bull', label: 'Bull Target ({{bull}})', type: 'input' },
      { key: 'bull_p', label: 'Bull Probability', type: 'input', hint: 'e.g. 25%' },
      { key: 'bull_content', label: 'Bull Case Notes', type: 'textarea' },
    ],
  },
  {
    title: 'Trading Strategy',
    slide: 'Slide 18',
    fields: [
      { key: 'entry_strategy_1', label: 'Entry Strategy', type: 'textarea' },
      { key: 'review_strategy_2', label: 'Review Strategy', type: 'textarea' },
      { key: 'exit_strategy_3', label: 'Exit Strategy', type: 'textarea' },
      { key: 'stp_loss', label: 'Stop Loss Price', type: 'input', hint: 'Bear case target / 15% below CMP' },
      { key: 'down', label: 'Downside %', type: 'input', hint: 'e.g. -12.5%' },
      { key: 'pnt', label: 'Entry Pivot Point', type: 'input', hint: 'Accumulation price level' },
    ],
  },
];

// These are filled from Excel/financial model — excluded from this panel
const EXCLUDED_KEYS = new Set([
  'financial_summary_image',
  'financial_model_from_excel',
  'financial_model_from_excel_operational_sheet',
  'timeline',
  'competitive_chart_1',
  'competitive_chart_2',
  'governance_table',
  'earnings_forecast_table',
  'financials_table',
  'valuations_table',
  'probability_weight_table',
  'key_risks_table',
]);

// ── Component ─────────────────────────────────────────────────────────────────

interface PPTDataPanelProps {
  reportId: string | null;
  sessionId: string;
  serviceAvailable: boolean;
  onConfirmed: (data: Record<string, string>) => void;
}

export default function PPTDataPanel({
  reportId,
  sessionId,
  serviceAvailable,
  onConfirmed,
}: PPTDataPanelProps) {
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [confirmed, setConfirmed] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [warnings, setWarnings] = useState<string[]>([]);
  const [values, setValues] = useState<Record<string, string>>({});
  const [openGroups, setOpenGroups] = useState<Set<string>>(new Set(['Cover & Market Data', 'Investment Thesis']));

  const loadPlaceholders = useCallback(async () => {
    if (!reportId || !serviceAvailable) return;
    setLoading(true);
    setLoadError(null);
    try {
      const result = await fetchPptPlaceholders(reportId, sessionId);
      // Filter out excluded keys
      const filtered: Record<string, string> = {};
      for (const [k, v] of Object.entries(result.placeholders)) {
        if (!EXCLUDED_KEYS.has(k)) filtered[k] = v;
      }
      setValues(filtered);
      setWarnings(result.warnings || []);
      if (result.has_saved_overrides) {
        setConfirmed(true);
        onConfirmed(filtered);
      }
    } catch (err) {
      setLoadError(err instanceof Error ? err.message : 'Failed to load PPT data');
    } finally {
      setLoading(false);
    }
  }, [reportId, sessionId, serviceAvailable, onConfirmed]);

  useEffect(() => {
    loadPlaceholders();
  }, [loadPlaceholders]);

  const handleChange = (key: string, val: string) => {
    setValues(prev => ({ ...prev, [key]: val }));
    setConfirmed(false);
  };

  const handleConfirm = async () => {
    if (!reportId) return;
    setSaving(true);
    try {
      await savePptPlaceholders(reportId, values);
      setConfirmed(true);
      onConfirmed(values);
      toast.success('PPT content confirmed and saved');
    } catch (err) {
      toast.error(err instanceof Error ? err.message : 'Failed to save PPT data');
    } finally {
      setSaving(false);
    }
  };

  const toggleGroup = (title: string) => {
    setOpenGroups(prev => {
      const next = new Set(prev);
      next.has(title) ? next.delete(title) : next.add(title);
      return next;
    });
  };

  if (!serviceAvailable) {
    return (
      <div className="rounded-lg border border-amber-200 bg-amber-50 px-4 py-3 flex items-center gap-2 text-sm text-amber-700">
        <AlertCircle className="h-4 w-4 shrink-0" />
        PPT service is offline — PPT content review unavailable. PPTX will be generated with auto-extracted values.
      </div>
    );
  }

  if (loading) {
    return (
      <div className="flex items-center gap-2 text-sm text-neutral-500 py-3">
        <Loader2 className="h-4 w-4 animate-spin" />
        Extracting PPT content from report…
      </div>
    );
  }

  if (loadError) {
    return (
      <div className="rounded-lg border border-red-200 bg-red-50 px-4 py-3 space-y-2">
        <div className="flex items-center gap-2 text-sm text-red-700">
          <AlertCircle className="h-4 w-4 shrink-0" />
          {loadError}
        </div>
        <Button size="sm" variant="outline" onClick={loadPlaceholders} className="h-7 text-xs">
          <RefreshCw className="h-3 w-3 mr-1.5" /> Retry
        </Button>
      </div>
    );
  }

  const filledCount = Object.values(values).filter(v => v.trim()).length;
  const totalCount = Object.values(values).length;

  return (
    <div className="space-y-3">
      {/* Header bar */}
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-2 text-xs text-neutral-500">
          <FileEdit className="h-3.5 w-3.5" />
          <span>{filledCount}/{totalCount} fields populated</span>
          {warnings.length > 0 && (
            <span className="text-amber-600 flex items-center gap-1">
              <AlertCircle className="h-3 w-3" /> {warnings.length} warning{warnings.length > 1 ? 's' : ''}
            </span>
          )}
        </div>
        <div className="flex items-center gap-2">
          <Button
            size="sm"
            variant="outline"
            onClick={loadPlaceholders}
            disabled={loading}
            className="h-7 text-xs"
          >
            <RefreshCw className="h-3 w-3 mr-1.5" /> Reload from Report
          </Button>
          <Button
            size="sm"
            onClick={handleConfirm}
            disabled={saving || totalCount === 0}
            className={cn(
              'h-7 text-xs font-semibold',
              confirmed
                ? 'bg-green-600 hover:bg-green-700 text-white'
                : 'bg-accent-600 hover:bg-accent-700 text-white',
            )}
          >
            {saving ? (
              <><Loader2 className="h-3 w-3 mr-1.5 animate-spin" /> Saving…</>
            ) : confirmed ? (
              <><Check className="h-3 w-3 mr-1.5" /> Confirmed</>
            ) : (
              'Confirm & Save'
            )}
          </Button>
        </div>
      </div>

      {/* Excluded placeholders notice */}
      <div className="rounded-lg border border-neutral-100 bg-neutral-50 px-3 py-2 text-xs text-neutral-400">
        <span className="font-medium text-neutral-500">Auto-filled from Excel model:</span>{' '}
        financial_summary_image, financial_model_from_excel, timeline, competitive charts, governance/earnings/financials/valuations tables — these will be inserted automatically.
      </div>

      {/* Accordion groups */}
      <div className="space-y-2">
        {PLACEHOLDER_GROUPS.map(group => {
          const isOpen = openGroups.has(group.title);
          const groupFilled = group.fields.filter(f => values[f.key]?.trim()).length;
          return (
            <div key={group.title} className="rounded-lg border border-neutral-200 overflow-hidden">
              <button
                type="button"
                onClick={() => toggleGroup(group.title)}
                className="w-full flex items-center justify-between px-4 py-2.5 bg-neutral-50 hover:bg-neutral-100 transition-colors text-left"
              >
                <div className="flex items-center gap-2">
                  <span className="text-sm font-medium text-neutral-800">{group.title}</span>
                  <span className="text-xs text-neutral-400 bg-neutral-200 px-1.5 py-0.5 rounded font-mono">{group.slide}</span>
                </div>
                <div className="flex items-center gap-2 text-xs text-neutral-400">
                  <span>{groupFilled}/{group.fields.length}</span>
                  {isOpen ? <ChevronUp className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
                </div>
              </button>

              {isOpen && (
                <div className="divide-y divide-neutral-100">
                  {group.fields.map(field => (
                    <FieldRow
                      key={field.key}
                      field={field}
                      value={values[field.key] ?? ''}
                      onChange={val => handleChange(field.key, val)}
                    />
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* Confirm button (sticky bottom) */}
      <div className="sticky bottom-0 bg-white border-t border-neutral-100 pt-3 flex items-center justify-between">
        <p className="text-xs text-neutral-400">
          {confirmed
            ? 'Content confirmed — PPTX will use these values'
            : 'Review and confirm before generating PPTX'}
        </p>
        <Button
          onClick={handleConfirm}
          disabled={saving || totalCount === 0}
          className={cn(
            'h-8 px-4 text-xs font-semibold',
            confirmed
              ? 'bg-green-600 hover:bg-green-700 text-white'
              : 'bg-accent-600 hover:bg-accent-700 text-white',
          )}
        >
          {saving ? (
            <><Loader2 className="h-3.5 w-3.5 mr-1.5 animate-spin" /> Saving…</>
          ) : confirmed ? (
            <><Check className="h-3.5 w-3.5 mr-1.5" /> Confirmed & Saved</>
          ) : (
            'Confirm & Save'
          )}
        </Button>
      </div>
    </div>
  );
}

// ── Field row ─────────────────────────────────────────────────────────────────

interface FieldRowProps {
  field: FieldDef;
  value: string;
  onChange: (val: string) => void;
}

function FieldRow({ field, value, onChange }: FieldRowProps) {
  const isEmpty = !value.trim();
  return (
    <div className="px-4 py-2.5 grid grid-cols-[200px_1fr] gap-3 items-start">
      <div className="pt-1.5">
        <p className={cn('text-xs font-medium', isEmpty ? 'text-amber-600' : 'text-neutral-700')}>
          {field.label}
        </p>
        <p className="text-[10px] text-neutral-400 font-mono mt-0.5">{`{{${field.key}}}`}</p>
        {field.hint && <p className="text-[10px] text-neutral-400 mt-0.5 italic">{field.hint}</p>}
      </div>
      {field.type === 'textarea' ? (
        <textarea
          value={value}
          onChange={e => onChange(e.target.value)}
          rows={3}
          placeholder={isEmpty ? '⚠ Empty — will be blank in PPT' : ''}
          className={cn(
            'w-full rounded-md border px-2.5 py-1.5 text-xs text-neutral-800 leading-relaxed resize-y focus:outline-none focus:ring-2 focus:ring-accent-500/40 focus:border-accent-400 transition-colors',
            isEmpty ? 'border-amber-200 bg-amber-50/30' : 'border-neutral-200 bg-white',
          )}
        />
      ) : (
        <input
          type="text"
          value={value}
          onChange={e => onChange(e.target.value)}
          placeholder={isEmpty ? '⚠ Empty' : ''}
          className={cn(
            'w-full rounded-md border px-2.5 py-1.5 text-xs text-neutral-800 focus:outline-none focus:ring-2 focus:ring-accent-500/40 focus:border-accent-400 transition-colors',
            isEmpty ? 'border-amber-200 bg-amber-50/30' : 'border-neutral-200 bg-white',
          )}
        />
      )}
    </div>
  );
}
