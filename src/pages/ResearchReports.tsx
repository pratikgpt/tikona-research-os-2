import { useState } from 'react';
import { Link } from 'react-router-dom';
import {
  FileText,
  ExternalLink,
  Trash2,
  ChevronLeft,
  ChevronRight,
  Search,
  Sparkles,
} from 'lucide-react';
import {
  useResearchSessionList,
  useDeleteResearchSession,
} from '@/hooks/useResearchSession';
import { useAuth } from '@/contexts/AuthContext';
import { Spinner, TableSkeleton } from '@/components/ui/spinner';
import { cn } from '@/lib/utils';
import StatusBadge from '@/components/StatusBadge';
import type { ResearchSession } from '@/types/database';

const PAGE_SIZE = 15;

function formatDate(iso: string): string {
  return new Date(iso).toLocaleDateString('en-IN', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  });
}

function formatTime(iso: string): string {
  return new Date(iso).toLocaleTimeString('en-IN', {
    hour: '2-digit',
    minute: '2-digit',
  });
}

type StatusFilter = 'all' | ResearchSession['status'];

export default function ResearchReports() {
  const { user } = useAuth();
  const [page, setPage] = useState(0);
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<StatusFilter>('all');
  const [deletingId, setDeletingId] = useState<string | null>(null);

  const { data: queryResult, isLoading } = useResearchSessionList(
    user?.email ?? undefined,
    page,
    PAGE_SIZE
  );

  const deleteMutation = useDeleteResearchSession();

  const sessions = queryResult?.data ?? [];
  const totalCount = queryResult?.count ?? 0;
  const totalPages = Math.ceil(totalCount / PAGE_SIZE);

  // Client-side filter
  const filtered = sessions.filter((s) => {
    const matchesSearch = !searchQuery ||
      s.company_name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      ((s.company_nse_code || '').toLowerCase().includes(searchQuery.toLowerCase())) ||
      (s.sector && s.sector.toLowerCase().includes(searchQuery.toLowerCase()));
    const matchesStatus = statusFilter === 'all' || s.current_state === statusFilter;
    return matchesSearch && matchesStatus;
  });

  const handleDelete = async (sessionId: string) => {
    if (!window.confirm('Delete this research session? This cannot be undone.')) {
      return;
    }
    setDeletingId(sessionId);
    try {
      await deleteMutation.mutateAsync(sessionId);
    } catch {
      // mutation error handled by React Query
    } finally {
      setDeletingId(null);
    }
  };

  const statusFilters: { value: StatusFilter; label: string }[] = [
    { value: 'all', label: 'All' },
    { value: 'document_review', label: 'Review' },
    { value: 'drafting', label: 'Drafting' },
    { value: 'completed', label: 'Completed' },
  ];

  return (
    <div className="flex h-full flex-col">
      {/* Page Header */}
      <header className="border-b border-neutral-200/80 bg-white px-7 py-5">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-lg font-semibold tracking-tight text-neutral-900">
              Research Reports
            </h1>
            <p className="text-sm text-neutral-500">
              {totalCount > 0
                ? `${totalCount} session${totalCount !== 1 ? 's' : ''}`
                : 'All research sessions'}
            </p>
          </div>

          <Link
            to="/admin/generate-research"
            className="inline-flex items-center gap-2 rounded-lg bg-accent-600 px-4 py-2 text-sm font-medium text-white shadow-sm transition-all duration-150 hover:bg-accent-700 active:scale-[0.97]"
          >
            <Sparkles className="h-4 w-4" />
            New Research
          </Link>
        </div>
      </header>

      {/* Main Content */}
      <div className="flex-1 overflow-auto bg-canvas p-7">
        {isLoading ? (
          <div className="overflow-hidden rounded-xl border border-neutral-200/60 bg-white">
            <TableSkeleton rows={8} cols={6} />
          </div>
        ) : sessions.length === 0 && page === 0 ? (
          /* Empty state */
          <div className="flex flex-col items-center justify-center rounded-xl border border-dashed border-neutral-200 bg-white py-16">
            <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-full bg-accent-50">
              <FileText className="h-7 w-7 text-accent-300" />
            </div>
            <h3 className="mt-4 text-sm font-medium text-neutral-900">
              No research sessions yet
            </h3>
            <p className="mt-1 text-sm text-neutral-500 max-w-sm mx-auto text-center">
              Start by generating your first research report from the Generate Research page.
            </p>
            <Link
              to="/admin/generate-research"
              className="mt-5 inline-flex items-center gap-2 rounded-lg bg-accent-600 px-4 py-2 text-sm font-medium text-white shadow-sm transition-all duration-150 hover:bg-accent-700 active:scale-[0.97]"
            >
              <Sparkles className="h-4 w-4" />
              Generate Research
            </Link>
          </div>
        ) : (
          <div className="space-y-4">
            {/* Search + Filters bar */}
            <div className="flex items-center gap-3">
              <div className="relative flex-1 max-w-sm">
                <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-neutral-400" />
                <input
                  type="text"
                  placeholder="Search by company, symbol, or sector..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full rounded-lg border border-neutral-200 bg-white py-2 pl-10 pr-4 text-sm text-neutral-900 placeholder:text-neutral-400 transition-all duration-150 focus-visible:border-accent-400 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent-500/40"
                />
              </div>
              <div className="flex items-center gap-1 rounded-lg border border-neutral-200 bg-white p-1">
                {statusFilters.map((f) => (
                  <button
                    key={f.value}
                    onClick={() => setStatusFilter(f.value)}
                    className={cn(
                      'rounded-md px-3 py-2 text-xs font-medium transition-colors',
                      statusFilter === f.value
                        ? 'bg-accent-600 text-white shadow-sm'
                        : 'text-neutral-600 hover:bg-neutral-50'
                    )}
                  >
                    {f.label}
                  </button>
                ))}
              </div>
            </div>

            {/* Table */}
            <div className="overflow-hidden rounded-xl border border-neutral-200/60 bg-white">
              <table className="min-w-full divide-y divide-neutral-100 animate-content-ready">
                <thead className="bg-neutral-50/80">
                  <tr>
                    <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wider text-neutral-500">
                      Company
                    </th>
                    <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wider text-neutral-500">
                      Symbol
                    </th>
                    <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wider text-neutral-500">
                      Sector
                    </th>
                    <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wider text-neutral-500">
                      Status
                    </th>
                    <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wider text-neutral-500">
                      Created
                    </th>
                    <th className="px-4 py-3 text-right text-xs font-semibold uppercase tracking-wider text-neutral-500">
                      Actions
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-neutral-100">
                  {filtered.length === 0 ? (
                    <tr>
                      <td
                        colSpan={6}
                        className="px-4 py-10 text-center text-sm text-neutral-400"
                      >
                        {searchQuery || statusFilter !== 'all'
                          ? 'No sessions match your filters.'
                          : 'No sessions on this page.'}
                      </td>
                    </tr>
                  ) : (
                    filtered.map((session) => (
                      <tr
                        key={session.session_id}
                        className="transition-colors hover:bg-accent-50/30"
                      >
                        <td className="whitespace-nowrap px-4 py-4 text-sm font-medium text-neutral-900">
                          {session.company_name}
                        </td>
                        <td className="whitespace-nowrap px-4 py-4 text-sm text-neutral-600">
                          <span className="rounded bg-neutral-100 px-2 py-0.5 font-mono text-xs">
                            {session.company_nse_code || '-'}
                          </span>
                        </td>
                        <td className="whitespace-nowrap px-4 py-4 text-sm text-neutral-500">
                          {session.sector || '-'}
                        </td>
                        <td className="whitespace-nowrap px-4 py-4">
                          <StatusBadge status={(session.current_state || session.status) as 'document_review' | 'drafting' | 'completed'} />
                        </td>
                        <td className="whitespace-nowrap px-4 py-4 text-sm text-neutral-500">
                          <div>{formatDate(session.created_at)}</div>
                          <div className="text-xs text-neutral-400">
                            {formatTime(session.created_at)}
                          </div>
                        </td>
                        <td className="whitespace-nowrap px-4 py-4 text-right">
                          <div className="flex items-center justify-end gap-1">
                            {Boolean(session.vault_folder_url) && (
                              <a
                                href={session.vault_folder_url as string}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="inline-flex items-center gap-1 rounded-md px-2 py-2 text-xs font-medium text-neutral-900 transition-colors hover:bg-neutral-100"
                                title="Open vault in Google Drive"
                              >
                                <ExternalLink className="h-3.5 w-3.5" />
                                Vault
                              </a>
                            )}
                            <button
                              onClick={() => handleDelete(session.session_id)}
                              disabled={deletingId === session.session_id}
                              className="inline-flex items-center gap-1 rounded-md px-2 py-2 text-xs font-medium text-red-600 transition-colors hover:bg-red-50 disabled:opacity-50"
                              title="Delete session"
                            >
                              {deletingId === session.session_id ? (
                                <Spinner size="sm" />
                              ) : (
                                <Trash2 className="h-3.5 w-3.5" />
                              )}
                              Delete
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination */}
            {totalPages > 1 && (
              <div className="flex items-center justify-between pt-1">
                <p className="text-sm text-neutral-500">
                  Showing {page * PAGE_SIZE + 1}–
                  {Math.min((page + 1) * PAGE_SIZE, totalCount)} of{' '}
                  {totalCount}
                </p>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setPage((p) => Math.max(0, p - 1))}
                    disabled={page === 0}
                    className="inline-flex items-center gap-1 rounded-lg border border-neutral-200 bg-white px-3 py-2 text-sm font-medium text-neutral-700 transition-colors hover:bg-neutral-50 disabled:cursor-not-allowed disabled:opacity-50"
                  >
                    <ChevronLeft className="h-4 w-4" />
                    Previous
                  </button>
                  <span className="text-sm text-neutral-500">
                    {page + 1} / {totalPages}
                  </span>
                  <button
                    onClick={() =>
                      setPage((p) => Math.min(totalPages - 1, p + 1))
                    }
                    disabled={page >= totalPages - 1}
                    className="inline-flex items-center gap-1 rounded-lg border border-neutral-200 bg-white px-3 py-2 text-sm font-medium text-neutral-700 transition-colors hover:bg-neutral-50 disabled:cursor-not-allowed disabled:opacity-50"
                  >
                    Next
                    <ChevronRight className="h-4 w-4" />
                  </button>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
