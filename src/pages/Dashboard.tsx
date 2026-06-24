import { Link } from 'react-router-dom';
import {
  Building2,
  Globe,
  FileText,
  Sparkles,
  ExternalLink,
  ArrowRight,
  ArrowUpRight,
  BarChart3,
} from 'lucide-react';
import { useMasterCompanyList } from '@/hooks/useMasterCompany';
import { useResearchSessionList } from '@/hooks/useResearchSession';
import { useAuth } from '@/contexts/AuthContext';
import { StatsSkeleton, TableSkeleton } from '@/components/ui/spinner';
import StatusBadge from '@/components/StatusBadge';

function formatRelativeDate(iso: string): string {
  const date = new Date(iso);
  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffMins = Math.floor(diffMs / 60000);
  const diffHours = Math.floor(diffMs / 3600000);
  const diffDays = Math.floor(diffMs / 86400000);

  if (diffMins < 1) return 'Just now';
  if (diffMins < 60) return `${diffMins}m ago`;
  if (diffHours < 24) return `${diffHours}h ago`;
  if (diffDays < 7) return `${diffDays}d ago`;

  return date.toLocaleDateString('en-IN', {
    day: '2-digit',
    month: 'short',
  });
}

export default function Dashboard() {
  const { user } = useAuth();
  const { data: queryResult, isLoading } = useMasterCompanyList();
  const { data: sessionsResult, isLoading: sessionsLoading } =
    useResearchSessionList(user?.email ?? undefined, 0, 5);

  const totalCompanies = queryResult?.count ?? 0;
  const totalSessions = sessionsResult?.count ?? 0;
  const recentSessions = sessionsResult?.data ?? [];

  return (
    <div className="flex h-full flex-col">
      {/* Page Header */}
      <header className="border-b border-neutral-200/80 bg-white px-7 py-5">
        <h1 className="text-lg font-semibold tracking-tight text-neutral-900">Dashboard</h1>
      </header>

      {/* Content */}
      <div className="flex-1 overflow-auto p-7">
        <div className="space-y-6 max-w-6xl mx-auto">
          {/* Stats Row */}
          {isLoading ? (
            <StatsSkeleton />
          ) : (
            <div className="grid grid-cols-2 gap-4">
              <div className="card-premium p-5">
                <div className="flex items-center gap-3">
                  <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-accent-50">
                    <Building2 className="h-5 w-5 text-accent-600" />
                  </div>
                  <div>
                    <p className="text-xs font-medium text-neutral-500 uppercase tracking-wider">Companies</p>
                    <p className="mt-1 text-2xl font-semibold text-neutral-900 tabular-nums">{totalCompanies.toLocaleString()}</p>
                  </div>
                </div>
              </div>
              <div className="card-premium p-5">
                <div className="flex items-center gap-3">
                  <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-accent-50">
                    <BarChart3 className="h-5 w-5 text-accent-600" />
                  </div>
                  <div>
                    <p className="text-xs font-medium text-neutral-500 uppercase tracking-wider">Research Sessions</p>
                    <p className="mt-1 text-2xl font-semibold text-neutral-900 tabular-nums">{totalSessions.toLocaleString()}</p>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Navigation Grid */}
          <div>
            <h2 className="text-xs font-medium text-neutral-400 uppercase tracking-wider mb-3">
              Navigate
            </h2>
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
              {[
                { to: '/admin/equity-database', icon: Building2, label: 'Equity Database' },
                { to: '/admin/universe', icon: Globe, label: 'Equity Universe' },
                { to: '/admin/generate-research', icon: Sparkles, label: 'Generate Research' },
                { to: '/admin/research-reports', icon: FileText, label: 'Reports' },
              ].map((action, i) => (
                <Link
                  key={action.to}
                  to={action.to}
                  className="group card-premium flex items-center gap-3 px-4 py-4 animate-fade-up-stagger"
                  style={{ animationDelay: `${i * 60}ms` }}
                >
                  <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-accent-50 transition-colors group-hover:bg-accent-100">
                    <action.icon className="h-4 w-4 text-accent-600 shrink-0" />
                  </div>
                  <span className="text-sm font-medium text-neutral-700 flex-1">{action.label}</span>
                  <ArrowUpRight className="h-3.5 w-3.5 text-neutral-300 transition-all group-hover:text-accent-500 group-hover:translate-x-0.5 group-hover:-translate-y-0.5" />
                </Link>
              ))}
            </div>
          </div>

          {/* Recent Activity */}
          <div>
            <div className="flex items-center justify-between mb-3">
              <h2 className="text-xs font-medium text-neutral-400 uppercase tracking-wider">
                Recent Sessions
              </h2>
              {totalSessions > 5 && (
                <Link
                  to="/admin/research-reports"
                  className="inline-flex items-center gap-1 text-xs font-medium text-accent-600 hover:text-accent-700"
                >
                  View all
                  <ArrowRight className="h-3 w-3" />
                </Link>
              )}
            </div>

            {sessionsLoading ? (
              <div className="rounded-xl border border-neutral-200/60 bg-white overflow-hidden">
                <TableSkeleton rows={5} cols={4} />
              </div>
            ) : recentSessions.length === 0 ? (
              <div className="rounded-xl border border-dashed border-neutral-200 bg-white py-12 text-center">
                <p className="text-sm text-neutral-500">No research sessions yet</p>
                <Link
                  to="/admin/generate-research"
                  className="mt-3 inline-flex items-center gap-2 text-sm font-medium text-accent-600 hover:text-accent-700"
                >
                  Create your first report
                  <ArrowRight className="h-3.5 w-3.5" />
                </Link>
              </div>
            ) : (
              <div className="rounded-xl border border-neutral-200/60 bg-white overflow-hidden">
                <table className="w-full">
                  <thead>
                    <tr className="border-b border-neutral-100">
                      <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wider text-neutral-500">Company</th>
                      <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wider text-neutral-500">Symbol</th>
                      <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wider text-neutral-500">Status</th>
                      <th className="px-4 py-3 text-right text-xs font-semibold uppercase tracking-wider text-neutral-500">When</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-neutral-100">
                    {recentSessions.map((session, i) => (
                      <tr
                        key={session.session_id}
                        className="transition-colors hover:bg-accent-50/30 animate-fade-up-stagger"
                        style={{ animationDelay: `${i * 50}ms` }}
                      >
                        <td className="px-4 py-3 text-sm font-medium text-neutral-900">
                          {session.company_name}
                        </td>
                        <td className="px-4 py-3">
                          <span className="font-mono text-xs text-neutral-500">
                            {session.company_nse_code || '-'}
                          </span>
                        </td>
                        <td className="px-4 py-3">
                          <StatusBadge status={(session.current_state || session.status) as 'document_review' | 'drafting' | 'completed'} />
                        </td>
                        <td className="px-4 py-3 text-right">
                          <div className="flex items-center justify-end gap-2">
                            <span className="text-xs text-neutral-400">
                              {formatRelativeDate(session.created_at)}
                            </span>
                            {Boolean(session.vault_folder_url) && (
                              <a
                                href={session.vault_folder_url as string}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-neutral-300 transition-colors hover:text-accent-600"
                                title="Open vault"
                              >
                                <ExternalLink className="h-3 w-3" />
                              </a>
                            )}
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
