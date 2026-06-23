import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { cn } from '@/lib/utils';
import { Check, RotateCcw, Edit3, Save, X, Maximize2, Minimize2, Globe } from 'lucide-react';

interface StageReviewProps {
  title: string;
  content: string;
  isApproved: boolean;
  isGenerating: boolean;
  onApprove: () => void;
  onRegenerate: () => void;
  onEdit?: (content: string) => void;
  className?: string;
  maxHeight?: number;
}

export default function StageReview({
  title,
  content,
  isApproved,
  isGenerating,
  onApprove,
  onRegenerate,
  onEdit,
  className,
  maxHeight = 500,
}: StageReviewProps) {
  const [isEditing, setIsEditing] = useState(false);
  const [editContent, setEditContent] = useState(content);
  const [isFullHeight, setIsFullHeight] = useState(false);

  useEffect(() => {
    setEditContent(content);
  }, [content]);

  const handleSave = () => {
    onEdit?.(editContent);
    setIsEditing(false);
  };

  // Loading state
  if (isGenerating && !content) {
    return (
      <div className={cn('p-8', className)}>
        <div className="flex flex-col items-center gap-4 py-6">
          <div className="relative">
            <div className="h-12 w-12 rounded-2xl bg-accent-100 flex items-center justify-center">
              <Globe className="h-6 w-6 text-accent-600 animate-pulse" />
            </div>
            <div className="absolute -top-1 -right-1 h-3 w-3 rounded-full bg-accent-500 animate-ping" />
          </div>
          <div className="text-center">
            <p className="text-sm font-semibold text-accent-800">Researching with Claude</p>
            <p className="text-xs text-accent-500 mt-1">Searching the web and analyzing data...</p>
          </div>
          <div className="w-48 h-1 rounded-full bg-accent-100 overflow-hidden">
            <div className="h-full w-1/3 rounded-full bg-accent-500 animate-[loading_1.5s_ease-in-out_infinite]" />
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className={cn('', className)}>
      {/* Approved banner */}
      {isApproved && (
        <div className="flex items-center gap-2 px-4 py-2 bg-green-50 border-b border-green-100">
          <div className="flex h-5 w-5 items-center justify-center rounded-full bg-green-500 text-white">
            <Check className="h-3 w-3" strokeWidth={2.5} />
          </div>
          <span className="text-xs font-semibold text-green-700">{title}</span>
          <span className="text-xs font-semibold text-green-600 bg-green-100 px-2 py-0.5 rounded-full ml-auto">Approved</span>
        </div>
      )}

      {/* Content */}
      {isEditing ? (
        <div className="p-4">
          <textarea
            value={editContent}
            onChange={(e) => setEditContent(e.target.value)}
            className="w-full min-h-[300px] max-h-[600px] rounded-lg border border-neutral-200 bg-neutral-50 p-4 text-sm font-mono text-neutral-800 leading-relaxed resize-y focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent-500/40 focus-visible:border-accent-400 focus-visible:bg-white transition-all"
            spellCheck={false}
          />
          <div className="flex gap-2 mt-3">
            <Button size="sm" onClick={handleSave} className="rounded-lg">
              <Save className="h-3.5 w-3.5 mr-1.5" /> Save Changes
            </Button>
            <Button size="sm" variant="outline" onClick={() => { setIsEditing(false); setEditContent(content); }} className="rounded-lg">
              <X className="h-3.5 w-3.5 mr-1.5" /> Cancel
            </Button>
          </div>
        </div>
      ) : (
        <>
          {/* Toolbar — expand toggle */}
          <div className="flex items-center justify-end px-4 py-2 border-b border-neutral-50">
            <button
              onClick={() => setIsFullHeight(!isFullHeight)}
              className="p-1 rounded hover:bg-neutral-100 text-neutral-400 hover:text-neutral-600 transition-colors"
              title={isFullHeight ? 'Collapse' : 'Expand full'}
            >
              {isFullHeight ? <Minimize2 className="h-3.5 w-3.5" /> : <Maximize2 className="h-3.5 w-3.5" />}
            </button>
          </div>

          {/* Markdown content */}
          <div
            className={cn(
              'prose prose-sm prose-neutral max-w-none text-neutral-700 leading-relaxed px-5 py-4 overflow-y-auto',
              !isFullHeight && 'scrollbar-thin'
            )}
            style={isFullHeight ? undefined : { maxHeight: `${maxHeight}px` }}
            dangerouslySetInnerHTML={{ __html: renderMarkdown(content) }}
          />

          {/* Actions */}
          {(!isApproved || onRegenerate || onEdit) && (
            <div className="flex items-center gap-2 px-4 py-3 border-t border-neutral-100 bg-neutral-50/50">
              {!isApproved && (
                <Button onClick={onApprove} size="sm" className="rounded-lg shadow-sm bg-accent-600 hover:bg-accent-700">
                  <Check className="h-3.5 w-3.5 mr-1.5" /> Approve & Continue
                </Button>
              )}
              {onRegenerate && (
                <Button onClick={onRegenerate} variant="outline" size="sm" disabled={isGenerating} className="rounded-lg">
                  <RotateCcw className="h-3.5 w-3.5 mr-1.5" /> Regenerate
                </Button>
              )}
              {onEdit && (
                <Button onClick={() => setIsEditing(true)} variant="ghost" size="sm" className="rounded-lg ml-auto">
                  <Edit3 className="h-3.5 w-3.5 mr-1.5" /> Edit
                </Button>
              )}
            </div>
          )}
        </>
      )}
    </div>
  );
}

// Simple markdown to HTML renderer
function renderMarkdown(md: string): string {
  if (!md) return '<p class="text-neutral-400 italic">No content generated yet.</p>';

  let html = md
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    // Headers
    .replace(/^#### (.+)$/gm, '<h4 class="text-sm font-semibold text-neutral-800 mt-4 mb-1">$1</h4>')
    .replace(/^### (.+)$/gm, '<h3 class="text-sm font-bold text-neutral-900 mt-5 mb-2">$1</h3>')
    .replace(/^## (.+)$/gm, '<h2 class="text-base font-bold text-neutral-900 mt-6 mb-2 pb-1 border-b border-neutral-100">$1</h2>')
    .replace(/^# (.+)$/gm, '<h1 class="text-lg font-bold text-neutral-900 mt-6 mb-3">$1</h1>')
    // Bold and italic
    .replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.+?)\*/g, '<em>$1</em>')
    // Unordered lists
    .replace(/^- (.+)$/gm, '<li class="ml-4 list-disc text-sm leading-relaxed">$1</li>')
    // Ordered lists
    .replace(/^\d+\. (.+)$/gm, '<li class="ml-4 list-decimal text-sm leading-relaxed">$1</li>')
    // Table rows
    .replace(/^\|(.+)\|$/gm, (_, row: string) => {
      const cells = row.split('|').map(c => c.trim()).filter(Boolean);
      return '<tr>' + cells.map(c => `<td class="border border-neutral-200 px-3 py-2 text-xs">${c}</td>`).join('') + '</tr>';
    })
    // Horizontal rules
    .replace(/^---+$/gm, '<hr class="my-5 border-neutral-200" />')
    // Paragraphs
    .replace(/\n\n/g, '</p><p class="text-sm leading-relaxed mb-2">')
    .replace(/\n/g, '<br />');

  if (!html.startsWith('<h') && !html.startsWith('<table') && !html.startsWith('<li')) {
    html = `<p class="text-sm leading-relaxed mb-2">${html}</p>`;
  }

  html = html.replace(/((?:<li[^>]*>.*?<\/li>\s*)+)/g, '<ul class="my-2 space-y-1">$1</ul>');
  html = html.replace(/((?:<tr>.*?<\/tr>\s*)+)/g, '<table class="w-full border-collapse my-3">$1</table>');

  return html;
}
